"""The SharePoint folder that holds one record's files.

Kept as its own record rather than a field on the business document so that any DocType can be
mapped without a schema change, and so the link is keyed on the folder's item id — which
survives someone renaming or moving the folder in Teams — rather than on a path.

Made by the app (an Automatic mapping, or Create folder now on a record) or by hand: someone picks
the record and either lets the folder be created where the mapping says, or links one that already
exists. A hand-made link then sends the record's attachments that are still on this server into it.
"""

import frappe
from frappe import _
from frappe.model.document import Document


class MicrosoftDriveFolder(Document):
	def validate(self):
		if self.item_id:
			return  # made by the app, already resolved

		from frappe_microsoft365 import microsoft_files as files

		row = files.mapping_for(self.reference_doctype)
		if not row:
			frappe.throw(
				_("{0} has no enabled Microsoft Drive Mapping. Create one first.").format(
					self.reference_doctype
				)
			)
		frappe.has_permission(self.reference_doctype, "write", doc=self.reference_name, throw=True)
		if files.get_folder(self.reference_doctype, self.reference_name):
			frappe.throw(
				_("{0} {1} already has a SharePoint folder.").format(
					self.reference_doctype, self.reference_name
				),
				frappe.DuplicateEntryError,
			)

		if (self.existing_folder or "").strip():
			values = files.existing_folder_values(row, self.existing_folder)
		else:
			values = files.new_folder_values(row, self.reference_doctype, self.reference_name)
		self.update(values)
		self.flags.made_by_hand = True

	def after_insert(self):
		if self.flags.made_by_hand:
			from frappe_microsoft365 import microsoft_files as files

			queued = files.queue_record(self.reference_doctype, self.reference_name)
			if queued:
				frappe.msgprint(
					_("{0} attachment(s) of this record are on their way to the folder.").format(queued),
					alert=True,
				)
