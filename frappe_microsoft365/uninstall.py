"""Leave a site as it was before this app: no broken attachments, no leftover fields.

Frappe's own uninstall deletes the app's DocTypes and module, but not custom fields added to
standard DocTypes (Event, File) or their columns, and it knows nothing about attachments that
were moved to SharePoint, whose links point at this app. This runs first and handles both:

1. Every attachment that lives in SharePoint is downloaded back to this server and its File row
   pointed at it again. The SharePoint copies are left in place. If any file cannot be brought
   back, the uninstall stops and says which, so no attachment is ever left pointing at nothing.
   A site that has lost access to SharePoint for good can set
   ``microsoft365_uninstall_skip_restore: 1`` in site_config.json to go ahead regardless; those
   attachments then stay readable in SharePoint only.
2. Every custom field this app added is deleted, and its column dropped.
3. Cached tokens and settings are cleared.

Frappe calls before_uninstall during ``--dry-run`` too, so a dry run only reports what would happen.
"""

import inspect

import frappe


def _dry_run():
	"""Whether this uninstall is a dry run. Frappe does not pass it to the hook, so look for it."""
	frame = inspect.currentframe()
	while frame:
		if frame.f_code.co_name == "remove_app" and "dry_run" in frame.f_locals:
			return bool(frame.f_locals["dry_run"])
		frame = frame.f_back
	return False


def before_uninstall():
	dry_run = _dry_run()
	restore_files(dry_run=dry_run)
	remove_custom_fields(dry_run=dry_run)
	if not dry_run:
		clear_caches()


def restore_files(dry_run=False):
	if not frappe.db.has_column("File", "custom_microsoft_status"):
		return
	from frappe_microsoft365 import microsoft_files as files

	names = frappe.get_all("File", filters={"custom_microsoft_status": files.STORED}, pluck="name")
	if not names:
		return
	print(f"Bringing {len(names)} attachment(s) back from SharePoint…")
	if dry_run:
		return

	failed = []
	for i, name in enumerate(names, start=1):
		try:
			files.bring_back(name)
			# Each restored file must stay restored if a later one fails.
			frappe.db.commit()  # nosemgrep
		except Exception as e:
			frappe.db.rollback()
			failed.append(f"{name}: {e}"[:300])
		if i % 25 == 0:
			print(f"  {i}/{len(names)}")

	if failed and not frappe.conf.get("microsoft365_uninstall_skip_restore"):
		frappe.throw(
			"Uninstall stopped: {0} attachment(s) could not be brought back from SharePoint, so "
			"removing the app now would leave them unopenable here. Fix the cause (usually the "
			"connection: Microsoft Settings > Troubleshoot) and run the uninstall again; files "
			"already brought back stay back. To go ahead anyway and leave these files in SharePoint "
			"only, set microsoft365_uninstall_skip_restore to 1 in site_config.json.\n{1}".format(
				len(failed), "\n".join(failed[:20])
			)
		)
	if failed:
		print(f"Skipped {len(failed)} file(s) still in SharePoint (microsoft365_uninstall_skip_restore).")


def remove_custom_fields(dry_run=False):
	from frappe_microsoft365.setup import app_custom_fields

	for doctype, fieldnames in app_custom_fields().items():
		print(f"Removing {len(fieldnames)} custom field(s) from {doctype}")
		if dry_run:
			continue
		for fieldname in fieldnames:
			name = frappe.db.get_value("Custom Field", {"dt": doctype, "fieldname": fieldname})
			if name:
				frappe.delete_doc("Custom Field", name, ignore_permissions=True, force=True)
			if frappe.db.has_column(doctype, fieldname):
				frappe.db.sql_ddl(
					f"alter table `tab{doctype}` drop column `{fieldname}`"
				)  # nosemgrep: names are this app's own constants
		frappe.clear_cache(doctype=doctype)


def clear_caches():
	from frappe_microsoft365 import microsoft_files as files
	from frappe_microsoft365 import microsoft_graph as graph

	graph.clear_app_token_cache()
	files.clear_mapping_cache()
	frappe.db.delete("Singles", {"doctype": "Microsoft Settings"})
