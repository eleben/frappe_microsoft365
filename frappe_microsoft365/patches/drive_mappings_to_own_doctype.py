"""Microsoft Drive Mapping was a child table of Microsoft Settings; it is now its own DocType.

The table keeps its rows across that change, but they are named with the random hashes child
rows get. Rename each to its DocType, which is how mappings are named now, and drop the
leftovers a DocType can no longer have twice.
"""

import frappe


def execute():
	if not frappe.db.table_exists("Microsoft Drive Mapping"):
		return

	rows = frappe.db.sql(
		"select name, reference_doctype from `tabMicrosoft Drive Mapping` order by creation asc",
		as_dict=True,
	)
	for row in rows:
		target = row.reference_doctype
		if not target or row.name == target:
			continue
		if frappe.db.sql("select 1 from `tabMicrosoft Drive Mapping` where name=%s", target):
			frappe.db.sql("delete from `tabMicrosoft Drive Mapping` where name=%s", row.name)
			continue
		frappe.db.sql(
			"""update `tabMicrosoft Drive Mapping`
			set name=%s, parent=NULL, parentfield=NULL, parenttype=NULL
			where name=%s""",
			(target, row.name),
		)

	frappe.cache().delete_value("frappe_microsoft365:drive_mappings")
