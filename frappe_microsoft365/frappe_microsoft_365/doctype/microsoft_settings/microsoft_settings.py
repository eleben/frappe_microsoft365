"""Microsoft Settings — single config page for the Azure AD app (keys + scopes)."""

import frappe
from frappe import _
from frappe.model.document import Document


class MicrosoftSettings(Document):
	def validate(self):
		# Surface the effective redirect URI so the admin can register it in Azure.
		if not self.redirect_uri:
			from frappe_microsoft365.microsoft_graph import get_redirect_uri
			# don't persist a computed value silently if integration disabled; just hint via msg
			if self.enabled:
				frappe.msgprint(
					f"Register this Redirect URI in Azure: <code>{get_redirect_uri(self)}</code>",
					indicator="blue", alert=True,
				)
		if self.enabled:
			# The secret was missing from this check, so a half-configured integration could be
			# enabled and only fail later, at sign-in, with an Azure error code.
			missing = [
				label
				for label, value in (
					(_("Tenant ID"), self.tenant_id),
					(_("Client ID"), self.client_id),
					(_("Client Secret"), self.get_password("client_secret", raise_exception=False)),
				)
				if not value
			]
			if missing:
				frappe.throw(
					_("{0} are required to enable the integration.").format(", ".join(missing))
				)

		self._reject_the_secret_id()
		self._drop_capabilities_that_lost_their_prerequisite()
		self._validate_files_mappings()

	def _validate_files_mappings(self):
		"""One row per DocType, only DocTypes that can carry attachments, and no stale ids.

		The resolved site and drive ids are a cache of the URL and library name; editing either
		has to drop them, or files keep flowing to the old library with nothing on screen saying so.
		"""
		before = {}
		previous = self.get_doc_before_save()
		if previous:
			before = {row.name: row for row in previous.get("files_mappings") or []}

		seen = set()
		for row in self.get("files_mappings") or []:
			if row.reference_doctype in seen:
				frappe.throw(_("Row {0}: {1} is mapped twice.").format(row.idx, row.reference_doctype))
			seen.add(row.reference_doctype)

			meta = frappe.get_meta(row.reference_doctype)
			if meta.istable or meta.issingle or row.reference_doctype in ("File", "Microsoft Drive Folder"):
				frappe.throw(_("Row {0}: {1} cannot have its own SharePoint folders.").format(
					row.idx, row.reference_doctype))

			old = before.get(row.name)
			if old and ((old.site_url or "") != (row.site_url or "") or (old.library or "") != (row.library or "")):
				row.site_id = None
				row.drive_id = None

	def _drop_capabilities_that_lost_their_prerequisite(self):
		"""Transcripts are reached through the meeting behind a join URL, so they cannot work
		without the standalone-meetings permission.

		The field is hidden when its prerequisite is off, and a hidden Check keeps whatever value
		it had — which would leave sign-in quietly asking Azure for transcript and recording
		consent that nothing on screen admits to wanting. Clear it instead, and say so.
		"""
		if self.use_transcripts and not self.use_teams:
			self.use_transcripts = 0
			frappe.msgprint(
				_(
					"Meeting transcripts and recordings were turned off: they need "
					"<b>Standalone Teams meetings</b>, which resolves a join link to the meeting "
					"they belong to."
				),
				indicator="orange",
				alert=True,
			)

	def _reject_the_secret_id(self):
		"""Azure shows a secret's Value and its Secret ID together, and only the Value works.

		Pasting the wrong one is the commonest setup mistake there is, and Microsoft only says
		so at sign-in, with AADSTS7000215, by which point the person has moved on. The two are
		trivially distinguishable, so say it here instead.
		"""
		from frappe_microsoft365.doctor import MASKED, looks_like_guid

		secret = (self.client_secret or "").strip()
		if not secret or MASKED.match(secret):
			return

		if looks_like_guid(secret):
			frappe.throw(
				_(
					"That looks like the Secret ID, not the secret value. In Azure, "
					"Certificates & secrets shows <b>Value</b> next to <b>Secret ID</b>: copy "
					"the Value. It is shown only once, so create a new secret if you have "
					"navigated away."
				),
				title=_("Wrong column"),
			)


@frappe.whitelist()
def get_effective_redirect_uri():
	"""Expose the default redirect URI for the Settings form help text."""
	frappe.only_for("System Manager")
	from frappe_microsoft365.microsoft_graph import get_redirect_uri
	return get_redirect_uri()
