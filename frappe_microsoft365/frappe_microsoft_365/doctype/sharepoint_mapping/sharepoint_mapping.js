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

// The record folders of this DocType, shown on its mapping as a table like any child table.
//
// Not a real child table: each SharePoint Folder is its own record, made by the app the moment a
// record needs one (from any user's upload, on any worker), checked against that record's
// permissions, and there can be thousands. Rows of the mapping would serialise every folder
// creation through one document. So this draws the same grid over those records: a row opens
// its folder, Add Row creates one in place.
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
		const cell = (html, cls = "col-xs-4") =>
			`<div class="col grid-static-col ${cls}"><div class="static-area ellipsis">${html}</div></div>`;
		const check = (name) =>
			`<div class="row-check col"><input type="checkbox" class="grid-row-check ms365-check"${
				name ? ` data-name="${esc(name)}"` : ""
			} tabIndex="-1"></div>`;
		const head = `<div class="grid-row"><div class="data-row row m-0">
			${check()}
			<div class="row-index col"><span>${__("No.")}</span></div>
			${cell(esc(__("Record")))}${cell(esc(__("Folder")))}${cell(esc(__("SharePoint")), "col-xs-3")}
		</div></div>`;
		const body = rows
			.map(
				(r, n) => `<div class="grid-row ms365-row" data-name="${esc(r.name)}" style="cursor:pointer">
				<div class="data-row row m-0">
					${check(r.name)}
					<div class="row-index col"><span>${n + 1}</span></div>
					${cell(esc(r.reference_name))}
					${cell(`<span title="${esc(r.folder_path || "")}">${esc(r.folder_name || r.folder_path || "")}</span>`)}
					${cell(
						r.web_url
							? `<a class="ms365-open" href="${esc(r.web_url)}" target="_blank" rel="noopener">${__("Open")} ↗</a>`
							: "",
						"col-xs-3",
					)}
				</div>
			</div>`,
			)
			.join("");
		const more =
			total > rows.length
				? `<a class="ms365-all">${__("Showing {0} of {1}. View all", [rows.length, total])}</a>`
				: "";
		$wrapper.html(`<div class="grid-field">
			<div class="form-grid-container"><div class="form-grid">
				<div class="grid-heading-row">${head}</div>
				<div class="grid-body">
					<div class="rows">${body}</div>
					${rows.length ? "" : `<div class="grid-empty text-center text-extra-muted">${__("No folders yet")}</div>`}
				</div>
			</div></div>
			<div class="small form-clickable-section grid-footer">
				<div class="flex justify-between">
					<div class="grid-buttons">
						<button type="button" class="btn btn-xs btn-danger ms365-delete hidden">${__("Delete")}</button>
						<button type="button" class="btn btn-xs btn-secondary ms365-add">${__("Add Row")}</button>
					</div>
					<div class="text-muted">${more}</div>
				</div>
			</div>
		</div>`);
		const $checks = () => $wrapper.find(".grid-body .ms365-check");
		const selected = () => $checks().filter(":checked").map((_, el) => $(el).data("name")).get();
		const sync = () => $wrapper.find(".ms365-delete").toggleClass("hidden", !selected().length);
		$wrapper.find(".grid-heading-row .ms365-check").on("change", (e) => {
			$checks().prop("checked", e.target.checked);
			sync();
		});
		$checks().on("change", sync);
		$wrapper.find(".ms365-delete").on("click", () => remove_folders(frm, selected()));
		$wrapper.find(".ms365-row").on("click", (e) => {
			if ($(e.target).closest(".ms365-open, .row-check").length) return;
			frappe.set_route("Form", "SharePoint Folder", $(e.currentTarget).data("name"));
		});
		$wrapper.find(".ms365-add").on("click", () => add_folder(frm));
		$wrapper.find(".ms365-all").on("click", () =>
			frappe.set_route("List", "SharePoint Folder", filters),
		);
	});
}

// Delete: removes only this system's link to each folder. The folder and its files stay in
// SharePoint, and attachments already stored there keep opening; a record whose link is gone
// gets a folder again the next time it needs one (Automatic) or someone creates one (On demand).
function remove_folders(frm, names) {
	if (!names.length) return;
	frappe.confirm(
		__(
			"Remove the link to {0} folder(s)? The folders and their files stay in SharePoint, and files already stored there keep opening.",
			[names.length],
		),
		() =>
			Promise.all(
				names.map((name) =>
					frappe.call({
						method: "frappe.client.delete",
						args: { doctype: "SharePoint Folder", name },
					}),
				),
			).finally(() => show_folders(frm)),
	);
}

// Add Row: pick the record, and optionally a folder that already exists; the SharePoint Folder
// is created (or linked) on the spot, the same as saving one from its own form.
function add_folder(frm) {
	const dialog = new frappe.ui.Dialog({
		title: __("Add Folder"),
		fields: [
			{
				fieldname: "reference_name",
				fieldtype: "Link",
				options: frm.doc.reference_doctype,
				label: __(frm.doc.reference_doctype),
				reqd: 1,
			},
			{
				fieldname: "existing_folder",
				fieldtype: "Data",
				label: __("Existing Folder"),
				description: __(
					"Leave empty to create the folder where this mapping says. To use a folder that already exists, paste its SharePoint address (from the browser or Copy link) or its path in the library, e.g. Projects/PROJ-0001.",
				),
			},
		],
		primary_action_label: __("Add"),
		primary_action(values) {
			frappe.call({
				method: "frappe.client.insert",
				args: {
					doc: {
						doctype: "SharePoint Folder",
						reference_doctype: frm.doc.reference_doctype,
						reference_name: values.reference_name,
						existing_folder: values.existing_folder || null,
					},
				},
				freeze: true,
				freeze_message: __("Creating the folder in SharePoint…"),
				callback: (r) => {
					if (!r.exc) {
						dialog.hide();
						frappe.show_alert({
							message: __("Folder ready for {0}.", [values.reference_name]),
							indicator: "green",
						});
						show_folders(frm);
					}
				},
			});
		},
	});
	dialog.show();
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
