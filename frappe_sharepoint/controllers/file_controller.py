import frappe
from frappe import _
import os

SETTINGS = "SharePoint Settings"


def file_upload(doc, method):
	"""
	Hook called after file insertion
	Uploads file to SharePoint if sync is enabled
	"""
	doctype = doc.attached_to_doctype
	docname = doc.attached_to_name
	is_file_uploaded = doc.uploaded_to_sharepoint
	filepath = None

	# Check if SharePoint sync is enabled and file hasn't been uploaded yet
	if (doctype and docname and method == "after_insert" and 
		frappe.db.exists("DocType", SETTINGS) and is_file_uploaded == 0):
		
		settings = frappe.get_single(SETTINGS)
		
		# Check if file sync is enabled in settings
		if settings.enable_file_sync:
			filepath = get_file_path(doc)
			
			if filepath:
				# Enqueue upload to background
				frappe.enqueue(
					"frappe_sharepoint.utils.sharepoint.trigger_sharepoint_upload",
					queue="long",
					doctype=doctype,
					docname=docname,
					filepath=filepath,
					filedoc=doc.name,
					timeout=-1
				)


def get_file_path(doc):
	"""
	Construct complete file path from File doc
	"""
	try:
		path = "private/files" if doc.is_private else "public/files"
		abspath = os.path.abspath(os.curdir)
		site_path = frappe.get_site_path(path, doc.file_name)
		filepath = f'{abspath}/{site_path}'
		return filepath
	except Exception as e:
		frappe.log_error("File path construction error", str(e))
		return None
