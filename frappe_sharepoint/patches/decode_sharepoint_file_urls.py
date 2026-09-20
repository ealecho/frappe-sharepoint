import frappe
from urllib.parse import unquote


def execute():
	"""
	SharePoint links used to be stored percent-encoded in File.file_url. Frappe
	encodes file_url again when rendering, which broke the links. Store them decoded
	"""
	files = frappe.get_all(
		"File",
		filters={"uploaded_to_sharepoint": 1, "file_url": ("like", "https://%")},
		fields=["name", "file_url"]
	)

	for file in files:
		decoded = unquote(file.file_url)
		if decoded != file.file_url:
			frappe.db.set_value("File", file.name, "file_url", decoded, update_modified=False)
