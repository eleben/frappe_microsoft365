"""SharePoint Mapping gained "Files panel on the form", Own tab by default: give each existing
mapping's DocType its Files tab."""

import frappe


def execute():
	from frappe_microsoft365 import microsoft_files as files

	for row in frappe.get_all("SharePoint Mapping", fields=["reference_doctype", "enabled", "files_panel"]):
		mode = (row.files_panel or files.PANEL_TAB) if row.enabled else files.PANEL_HIDDEN
		files.sync_panel_fields(row.reference_doctype, mode)
	files.clear_mapping_cache()
