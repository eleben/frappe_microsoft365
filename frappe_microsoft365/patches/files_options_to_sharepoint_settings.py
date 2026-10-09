"""SharePoint storage options moved from Microsoft Settings to their own SharePoint Settings page.

Same fieldnames, so each saved value is copied across as it is and the old row removed.
"""

import frappe

FIELDS = ("files_keep_local_copy", "files_archive_links", "files_delete_remote", "files_on_uninstall")


def execute():
	for field in FIELDS:
		value = frappe.db.get_value(
			"Singles", {"doctype": "Microsoft Settings", "field": field}, "value", order_by=None
		)
		if value is not None:
			frappe.db.set_single_value("SharePoint Settings", field, value)
	frappe.db.delete("Singles", {"doctype": "Microsoft Settings", "field": ["in", FIELDS]})
	frappe.clear_document_cache("SharePoint Settings", "SharePoint Settings")
