"""SharePoint Folder became the Folders table of SharePoint Mapping.

Its rows stay where they are; each is given its mapping as parent. A mapping is named after its
DocType, so the parent is the row's own reference_doctype.

Frappe does not add the parent columns when an existing table becomes a child table (a new one
is created with them), so they are added here first.
"""

import frappe


def execute():
	if not frappe.db.table_exists("SharePoint Folder"):
		return
	for column in ("parent", "parentfield", "parenttype"):
		if not frappe.db.has_column("SharePoint Folder", column):
			frappe.db.sql_ddl(
				f"alter table `tabSharePoint Folder` add column `{column}` varchar(140)"
			)  # nosemgrep
	if not frappe.db.sql("show index from `tabSharePoint Folder` where Column_name='parent'"):
		frappe.db.sql_ddl("alter table `tabSharePoint Folder` add index `parent`(`parent`)")  # nosemgrep
	frappe.clear_cache(doctype="SharePoint Folder")
	rows = frappe.db.sql(
		"""select name, reference_doctype from `tabSharePoint Folder`
		where ifnull(parent, '') = '' order by creation asc""",
		as_dict=True,
	)
	next_idx = {}
	for row in rows:
		parent = row.reference_doctype
		if parent not in next_idx:
			next_idx[parent] = (
				frappe.db.sql(
					"select max(idx) from `tabSharePoint Folder` where parent=%s and parenttype='SharePoint Mapping'",
					parent,
				)[0][0]
				or 0
			)
		next_idx[parent] += 1
		frappe.db.sql(
			"""update `tabSharePoint Folder`
			set parent=%s, parenttype='SharePoint Mapping', parentfield='folders', idx=%s
			where name=%s""",
			(parent, next_idx[parent], row.name),
		)
