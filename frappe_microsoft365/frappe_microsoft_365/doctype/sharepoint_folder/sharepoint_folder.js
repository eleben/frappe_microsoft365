// SharePoint Folder — one record's folder in SharePoint, made by the app or by hand.

frappe.ui.form.on("SharePoint Folder", {
	setup(frm) {
		// Only DocTypes that have an enabled SharePoint Mapping can have a folder.
		frm.set_query("reference_doctype", () => ({
			query: "frappe_microsoft365.microsoft_files.mapped_doctype_query",
		}));
	},
	refresh(frm) {
		if (!frm.is_new() && frm.doc.web_url) {
			frm.add_custom_button(__("Open in SharePoint"), () =>
				window.open(frm.doc.web_url, "_blank"),
			);
		}
		if (!frm.is_new() && frm.doc.reference_doctype && frm.doc.reference_name) {
			frm.add_custom_button(__("Open Record"), () =>
				frappe.set_route("Form", frm.doc.reference_doctype, frm.doc.reference_name),
			);
		}
	},
	reference_doctype(frm) {
		frm.set_value("reference_name", "");
	},
});
