// SharePoint Mapping — one DocType's SharePoint destination, testable on its own.

frappe.ui.form.on("SharePoint Mapping", {
	refresh(frm) {
		if (frm.is_new()) {
			frm.set_intro(
				__(
					"Pick the DocType, paste the SharePoint site URL (for a Teams channel: Files › Open in SharePoint) and the channel name as Base Folder, then Save and Test.",
				),
			);
			return;
		}
		frm.add_custom_button(__("Test"), () => test_mapping(frm)).addClass("btn-primary");
		frm.add_custom_button(__("Site Grant Script"), () => grant_script(frm));
		frm.add_custom_button(__("Move Existing Attachments"), () => move_existing(frm));
		frm.add_custom_button(__("Microsoft Settings"), () =>
			frappe.set_route("Form", "Microsoft Settings"),
		);
	},
});

function test_mapping(frm) {
	if (frm.is_dirty()) {
		frappe.msgprint(__("Save first: the test uses the saved mapping."));
		return;
	}
	frappe.call({
		method: "frappe_microsoft365.microsoft_files.test_connection",
		args: { mapping: frm.doc.name },
		freeze: true,
		freeze_message: __("Signing in as the app and writing a test folder…"),
		callback: (r) => {
			const findings = (r.message || {}).findings || [];
			const counts = {};
			findings.forEach((f) => (counts[f.status] = (counts[f.status] || 0) + 1));
			frappe_microsoft365.show_findings(__("SharePoint Connection"), { findings, counts });
			frm.reload_doc();
		},
	});
}

function grant_script(frm) {
	frappe.call({
		method: "frappe_microsoft365.doctor.site_grant_powershell",
		args: { mapping: frm.doc.name },
		callback: (r) => {
			const dialog = new frappe.ui.Dialog({
				title: __("SharePoint Site Grant Script"),
				size: "large",
				fields: [
					{
						fieldtype: "HTML",
						options: `<div class="alert alert-info small">${__(
							"Gives the app <b>write</b> access to this site and no other. A SharePoint or Global administrator runs it in Microsoft Graph PowerShell; nothing runs from this screen.",
						)}</div>`,
					},
					{
						fieldname: "script",
						fieldtype: "Code",
						label: __("Microsoft Graph PowerShell"),
					},
				],
			});
			dialog.set_value("script", (r.message || {}).script || "");
			dialog.show();
		},
	});
}

function move_existing(frm) {
	frappe.confirm(
		__(
			"Queue every {0} attachment that is still on this server for SharePoint? Files move in the background and keep opening from their records throughout.",
			[frm.doc.reference_doctype],
		),
		() =>
			frappe.call({
				method: "frappe_microsoft365.microsoft_files.queue_existing",
				args: { doctype: frm.doc.reference_doctype },
				freeze: true,
				callback: (r) => {
					const out = r.message || {};
					frappe.msgprint(
						out.more
							? __(
									"{0} attachments queued. There are more: run this again once these are done.",
									[out.queued],
								)
							: __("{0} attachments queued.", [out.queued]),
					);
				},
			}),
	);
}
