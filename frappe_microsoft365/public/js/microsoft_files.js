// "Files in Microsoft 365" — the SharePoint folder of a record, shown on its form.
//
// Loaded on every desk page (app_include_js) because which DocTypes are mapped is a setting,
// not code: the list arrives in the boot (frappe.boot.microsoft365_files_doctypes) and the
// panel draws itself only on those forms. Everything it shows comes from whitelisted methods
// that check the user can read the record; ids in the page are never trusted on their own.

frappe.provide("frappe_microsoft365.files");

(function () {
	const esc = (s) => frappe.utils.escape_html(s == null ? "" : String(s));
	const METHOD = "frappe_microsoft365.microsoft_files";

	function mapped(doctype) {
		return (frappe.boot.microsoft365_files_doctypes || []).includes(doctype);
	}

	function size(bytes) {
		if (bytes == null) return "";
		if (bytes < 1024) return `${bytes} B`;
		if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
		return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
	}

	function when(iso) {
		if (!iso) return "";
		return frappe.datetime.comment_when
			? frappe.datetime.comment_when(iso.replace("T", " ").replace("Z", ""), true)
			: iso.slice(0, 10);
	}

	function open_url(frm, item_id, download) {
		const args = new URLSearchParams({
			doctype: frm.doctype,
			name: frm.docname,
			item_id,
			download: download ? 1 : 0,
		});
		return `/api/method/${METHOD}.open_item?${args.toString()}`;
	}

	class Panel {
		constructor(frm) {
			this.frm = frm;
			this.trail = []; // [{id, name}] subfolders opened from the record's folder
			frm.dashboard.show();
			this.$body = $(frm.dashboard.add_section("", __("Files in Microsoft 365")));
			this.$body.addClass("microsoft365-files");
			this.load();
		}

		load() {
			const subfolder = this.trail.length ? this.trail[this.trail.length - 1].id : null;
			this.$body.html(`<div class="text-muted small">${__("Loading…")}</div>`);
			frappe.call({
				method: `${METHOD}.list_folder`,
				args: { doctype: this.frm.doctype, name: this.frm.docname, subfolder },
				callback: (r) => this.render(r.message || {}),
				error: () =>
					this.$body.html(
						`<div class="text-muted small">${__(
							"Could not read the SharePoint folder. Microsoft Settings > Troubleshoot > Test SharePoint Connection says why.",
						)}</div>`,
					),
			});
		}

		render(data) {
			if (!data.exists) {
				const msg = data.missing
					? __(
							"This record's SharePoint folder was deleted or moved out of the library.",
						)
					: __("No SharePoint folder yet. It is created with the first attachment.");
				const btn = data.can_create
					? `<button class="btn btn-xs btn-default ms365-create">${__("Create folder now")}</button>`
					: "";
				this.$body.html(
					`<div class="small text-muted" style="display:flex;gap:10px;align-items:center">${esc(msg)} ${btn}</div>`,
				);
				this.$body.find(".ms365-create").on("click", () =>
					frappe.call({
						method: `${METHOD}.create_folder`,
						args: { doctype: this.frm.doctype, name: this.frm.docname },
						freeze: true,
						callback: () => this.load(),
					}),
				);
				return;
			}

			const crumbs = [
				`<a class="ms365-crumb" data-depth="0">${esc(data.folder_name || this.frm.docname)}</a>`,
			]
				.concat(
					this.trail.map(
						(t, i) =>
							`<a class="ms365-crumb" data-depth="${i + 1}">${esc(t.name)}</a>`,
					),
				)
				.join(" / ");

			const rows = (data.items || []).map((item) => {
				const icon = item.is_folder ? "📁" : "📄";
				const link = item.is_folder
					? `<a class="ms365-dir" data-id="${esc(item.id)}" data-name="${esc(item.name)}">${esc(item.name)}</a>`
					: `<a href="${esc(open_url(this.frm, item.id))}" target="_blank" rel="noopener">${esc(item.name)}</a>`;
				const tag = item.attachment
					? ` <span class="text-muted small" title="${esc(__("Attached in Frappe"))}">· ${__("attachment")}</span>`
					: "";
				const extra = item.is_folder
					? esc(__("{0} items", [item.child_count || 0]))
					: esc(size(item.size));
				const download = item.is_folder
					? ""
					: `<a href="${esc(open_url(this.frm, item.id, true))}" title="${esc(__("Download"))}" class="text-muted">⤓</a>`;
				return `<tr>
					<td style="width:24px">${icon}</td>
					<td>${link}${tag}</td>
					<td class="text-muted small text-nowrap">${esc(item.modified_by || "")}</td>
					<td class="text-muted small text-nowrap">${esc(when(item.modified))}</td>
					<td class="text-muted small text-right text-nowrap">${extra}</td>
					<td style="width:24px" class="text-right">${download}</td>
				</tr>`;
			});

			const table = rows.length
				? `<table class="table table-sm" style="margin:6px 0 0"><tbody>${rows.join("")}</tbody></table>`
				: `<div class="text-muted small" style="margin-top:6px">${__(
						"Empty. Attach a file here, or drop one into the folder in Teams.",
					)}</div>`;

			this.$body.html(`
				<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
					<div class="small" style="flex:1">${crumbs}</div>
					<button class="btn btn-xs btn-default ms365-refresh">${__("Refresh")}</button>
					${
						data.web_url
							? `<a class="btn btn-xs btn-default" href="${esc(data.web_url)}" target="_blank" rel="noopener">${__(
									"Open in SharePoint",
								)} ↗</a>`
							: ""
					}
				</div>
				${table}`);

			this.$body.find(".ms365-refresh").on("click", () => this.load());
			this.$body.find(".ms365-dir").on("click", (e) => {
				const $a = $(e.currentTarget);
				this.trail.push({ id: $a.data("id"), name: $a.data("name") });
				this.load();
			});
			this.$body.find(".ms365-crumb").on("click", (e) => {
				this.trail = this.trail.slice(0, Number($(e.currentTarget).data("depth")));
				this.load();
			});
		}
	}

	frappe_microsoft365.files.Panel = Panel;

	$(document).on("form-refresh", (_e, frm) => {
		if (!frm || frm.is_new() || !mapped(frm.doctype)) return;
		// The dashboard is rebuilt on every refresh, taking any earlier panel with it.
		frm.microsoft365_files_panel = new Panel(frm);
	});
})();
