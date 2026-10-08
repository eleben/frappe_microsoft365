"""Microsoft Drive Mapping / Folder are now SharePoint Mapping / Folder.

"Drive" is a Microsoft Graph API term no user ever sees; Microsoft's own screens say SharePoint.
Runs before the model sync so the existing tables, and their rows, are renamed rather than new
empty ones being created beside them.
"""

import frappe

RENAMES = (
	("Microsoft Drive Mapping", "SharePoint Mapping"),
	("Microsoft Drive Folder", "SharePoint Folder"),
)


def execute():
	for old, new in RENAMES:
		if frappe.db.exists("DocType", old) and not frappe.db.exists("DocType", new):
			frappe.rename_doc("DocType", old, new, force=True)
	frappe.cache().delete_value("frappe_microsoft365:drive_mappings")
