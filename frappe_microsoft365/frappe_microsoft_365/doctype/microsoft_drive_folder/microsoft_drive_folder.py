"""The SharePoint folder that holds one record's files.

Kept as its own record rather than a field on the business document so that any DocType can be
mapped without a schema change, and so the link is keyed on the folder's item id — which
survives someone renaming or moving the folder in Teams — rather than on a path.
"""

from frappe.model.document import Document


class MicrosoftDriveFolder(Document):
	pass
