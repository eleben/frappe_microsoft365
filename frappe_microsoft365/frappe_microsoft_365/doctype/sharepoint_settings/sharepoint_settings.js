// SharePoint Settings — how document storage behaves, and the tools to set it up and check it.
//
// The capability itself is switched on in Microsoft Settings, beside the others, because it
// changes what the Azure app is asked for. Everything about running it is here.

frappe.ui.form.on("SharePoint Settings", {
	refresh(frm) {
		frappe.db.get_single_value("Microsoft Settings", "use_files").then((on) => {
			const $status = frm.fields_dict.status_html.$wrapper;
			if (!on) {
				$status.html(
					`<div class="alert alert-warning small">${__(
						"SharePoint document storage is off. Tick it under Capabilities in <a href='/app/microsoft-settings'>Microsoft Settings</a>, then come back here."
					)}</div>`
				);
				return;
			}
			$status.empty();
			frappe_microsoft365.files.render_summary($status);
			frm.add_custom_button(__("Test Connection"), () => test_files(frm));
			frm.add_custom_button(__("Mappings"), () => frappe.set_route("List", "SharePoint Mapping"), __("Go to"));
			frm.add_custom_button(__("Folders"), () => frappe.set_route("List", "SharePoint Folder"), __("Go to"));
			frm.add_custom_button(__("Move Existing Attachments"), () => move_existing(frm), __("Actions"));
			frm.add_custom_button(__("Site Grant Script"), () => site_grant_script(), __("Actions"));
		});
	},
});

function test_files(frm) {
	if (frm.is_dirty()) {
		frappe.msgprint(__("Save first: the test uses the saved settings."));
		return;
	}
	frappe.call({
		method: "frappe_microsoft365.microsoft_files.test_connection",
		freeze: true,
		freeze_message: __("Signing in as the app and writing a test folder in each library…"),
		callback: (r) => {
			const result = r.message || {};
			const findings = result.findings || [];
			const counts = {};
			findings.forEach((f) => (counts[f.status] = (counts[f.status] || 0) + 1));
			frappe_microsoft365.show_findings(__("SharePoint Connection"), { findings, counts });
			frm.reload_doc();
		},
	});
}

function site_grant_script() {
	frappe.call({
		method: "frappe_microsoft365.doctor.site_grant_powershell",
		callback: (r) => {
			const out = r.message || {};
			const dialog = new frappe.ui.Dialog({
				title: __("SharePoint Site Grant Script"),
				size: "large",
				fields: [
					{
						fieldtype: "HTML",
						options: `<div class="alert alert-info small">${__(
							"Sites.Selected lets the app be granted individual sites; this script is the grant. It gives the app <b>write</b> access to the sites below and to no other site. A SharePoint or Global administrator runs it in Microsoft Graph PowerShell. Nothing runs from this screen."
						)}</div>`,
					},
					{ fieldname: "script", fieldtype: "Code", label: __("Microsoft Graph PowerShell"), read_only: 1 },
				],
			});
			dialog.set_value("script", out.script || "");
			dialog.show();
		},
	});
}

function move_existing(frm) {
	frappe.confirm(
		__(
			"Queue every attachment already on records of the mapped DocTypes for SharePoint? Files move in the background; each keeps opening from its record throughout."
		),
		() =>
			frappe.call({
				method: "frappe_microsoft365.microsoft_files.queue_existing",
				freeze: true,
				callback: (r) => {
					const out = r.message || {};
					frappe.msgprint(
						out.more
							? __("{0} attachments queued. There are more: run this again once these are done.", [out.queued])
							: __("{0} attachments queued.", [out.queued])
					);
					frappe_microsoft365.files.render_summary(frm.fields_dict.status_html.$wrapper);
				},
			})
	);
}
