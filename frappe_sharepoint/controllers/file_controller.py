import frappe
from frappe import _
from frappe.utils import add_to_date, now_datetime
import os

SETTINGS = "SharePoint Settings"
RETRY_WINDOW_DAYS = 7


def file_upload(doc, method):
	"""
	Hook called after file insertion
	Uploads file to SharePoint if sync is enabled
	"""
	# Check if SharePoint sync is enabled and file hasn't been uploaded yet
	if (method == "after_insert" and is_syncable(doc) and
		frappe.db.exists("DocType", SETTINGS)):

		settings = frappe.get_single(SETTINGS)

		# Check if file sync is enabled in settings
		if settings.enable_file_sync:
			enqueue_upload(doc)


def is_syncable(doc):
	"""
	Only local files attached to a document are sent to SharePoint
	"""
	file_url = doc.file_url or ""
	return bool(
		doc.attached_to_doctype and doc.attached_to_name
		and not doc.is_folder
		and not doc.uploaded_to_sharepoint
		and file_url.startswith(("/files/", "/private/files/"))
	)


def enqueue_upload(doc):
	filepath = get_file_path(doc)

	if filepath:
		# Enqueue upload to background
		frappe.enqueue(
			"frappe_sharepoint.utils.sharepoint.trigger_sharepoint_upload",
			queue="long",
			doctype=doc.attached_to_doctype,
			docname=doc.attached_to_name,
			filepath=filepath,
			filedoc=doc.name,
			timeout=-1,
			enqueue_after_commit=True
		)


def retry_pending_uploads():
	"""
	Hourly: re-queue recent attachments whose SharePoint upload did not complete
	"""
	if not frappe.db.exists("DocType", SETTINGS):
		return

	if not frappe.db.get_single_value(SETTINGS, "enable_file_sync"):
		return

	now = now_datetime()
	files = frappe.get_all(
		"File",
		filters={
			"uploaded_to_sharepoint": 0,
			"is_folder": 0,
			"attached_to_doctype": ("is", "set"),
			"attached_to_name": ("is", "set"),
			# Leave files alone while their first upload may still be running
			"creation": ("between", [add_to_date(now, days=-RETRY_WINDOW_DAYS), add_to_date(now, minutes=-15)]),
		},
		pluck="name",
		limit=100
	)

	for name in files:
		doc = frappe.get_doc("File", name)
		if is_syncable(doc):
			enqueue_upload(doc)


def get_file_path(doc):
	"""
	Construct complete file path from File doc
	"""
	try:
		filepath = os.path.abspath(doc.get_full_path())
		if not os.path.exists(filepath):
			frappe.log_error("File path construction error", f"{doc.name}: {filepath} not found")
			return None
		return filepath
	except Exception as e:
		frappe.log_error("File path construction error", str(e))
		return None
