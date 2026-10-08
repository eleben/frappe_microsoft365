"""Microsoft Drive Mapping — where one DocType's attachments are stored in SharePoint.

One record per DocType, named after it, so a DocType can never be mapped twice. Lives beside
Microsoft Settings rather than inside it, like Microsoft Calendar: the settings hold the
connection and the global storage options, each mapping holds one destination and can be
tested on its own.
"""

import frappe
from frappe import _
from frappe.model.document import Document


class MicrosoftDriveMapping(Document):
	def validate(self):
		meta = frappe.get_meta(self.reference_doctype)
		if (
			meta.istable
			or meta.issingle
			or self.reference_doctype
			in (
				"File",
				"Microsoft Drive Folder",
				"Microsoft Drive Mapping",
			)
		):
			frappe.throw(_("{0} cannot have its own SharePoint folders.").format(self.reference_doctype))

		# The resolved ids are a cache of the URL and the library name; editing either has to
		# drop them, or files keep flowing to the old library with nothing on screen saying so.
		if not self.is_new() and (self.has_value_changed("site_url") or self.has_value_changed("library")):
			self.site_id = None
			self.drive_id = None

	def on_update(self):
		from frappe_microsoft365.microsoft_files import clear_mapping_cache

		clear_mapping_cache()

	def on_trash(self):
		from frappe_microsoft365.microsoft_files import clear_mapping_cache

		clear_mapping_cache()
