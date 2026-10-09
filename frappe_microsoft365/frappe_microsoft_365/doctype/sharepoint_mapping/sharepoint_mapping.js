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
		frm.add_custom_button(__("SharePoint Settings"), () =>
			frappe.set_route("Form", "SharePoint Settings"),
		);
		show_folders(frm);
	},
});

// The record folders of this DocType, listed on its mapping, so a mapping is the one place to
// see where a DocType's files live. Most recently changed first; the full list is a click away.
const FOLDER_LIMIT = 50;

function show_folders(frm) {
	const field = frm.fields_dict.folders_html;
	if (!field) return;
	const $wrapper = field.$wrapper;
	const filters = { reference_doctype: frm.doc.reference_doctype };
	Promise.all([
		frappe.db.get_list("SharePoint Folder", {
			filters,
			fields: ["name", "reference_name", "folder_name", "folder_path", "web_url"],
			order_by: "modified desc",
			limit: FOLDER_LIMIT,
		}),
		frappe.db.count("SharePoint Folder", { filters }),
	]).then(([rows, total]) => {
		const esc = (v) => frappe.utils.escape_html(v == null ? "" : String(v));
		const record_url = (name) =>
			`/app/${frappe.router.slug(frm.doc.reference_doctype)}/${encodeURIComponent(name)}`;
		const body = rows.length
			? `<table class="table table-sm" style="margin:0">
				<thead><tr class="text-muted small">
					<th>${__("Record")}</th><th>${__("Folder")}</th><th></th>
				</tr></thead>
				<tbody>${rows
					.map(
						(r) => `<tr>
							<td><a href="${esc(record_url(r.reference_name))}">${esc(r.reference_name)}</a></td>
							<td class="text-muted"><a class="text-muted" href="/app/sharepoint-folder/${encodeURIComponent(r.name)}" title="${esc(r.folder_path || "")}">${esc(r.folder_name || r.folder_path || "")}</a></td>
							<td class="text-right">${
								r.web_url
									? `<a href="${esc(r.web_url)}" target="_blank" rel="noopener">${__("Open in SharePoint")} ↗</a>`
									: ""
							}</td>
						</tr>`,
					)
					.join("")}</tbody>
			</table>`
			: `<div class="text-muted small">${__(
					"No folders yet. They appear here as {0} records get their SharePoint folder.",
					[__(frm.doc.reference_doctype)],
				)}</div>`;
		const more =
			total > rows.length
				? `<a class="ms365-all small">${__("View all {0}", [total])}</a>`
				: "";
		$wrapper.html(`
			<div style="display:flex;gap:8px;align-items:center;margin-bottom:8px">
				<div class="small text-muted" style="flex:1">${__("{0} folder(s)", [total])}</div>
				${more}
				<button class="btn btn-xs btn-default ms365-add">${__("Add or Link Folder")}</button>
			</div>
			${body}`);
		$wrapper.find(".ms365-add").on("click", () =>
			frappe.new_doc("SharePoint Folder", { reference_doctype: frm.doc.reference_doctype }),
		);
		$wrapper.find(".ms365-all").on("click", () =>
			frappe.set_route("List", "SharePoint Folder", filters),
		);
	});
}

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
