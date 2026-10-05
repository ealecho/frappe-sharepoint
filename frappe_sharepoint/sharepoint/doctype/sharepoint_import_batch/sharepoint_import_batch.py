# Copyright (c) 2026, Frappe Community and contributors
# For license information, please see license.txt

import csv
import os
import re
import time
from datetime import timedelta
from urllib.parse import quote, unquote

import frappe
import requests
from frappe import _
from frappe.model.document import Document
from frappe.utils import now_datetime

from frappe_sharepoint.utils import get_request_header, make_request
from frappe_sharepoint.utils import xero_matching
from frappe_sharepoint.utils.sharepoint import SETTINGS, SharePoint, sanitize_name

# Header names accepted for each mapping column, matched case-insensitively
COLUMN_ALIASES = {
	"document_type": ("document type", "doctype", "type"),
	"document_name": ("id", "name", "document", "docname", "document id", "document name"),
	"folder": ("attachmentfolder", "attachment folder", "folder"),
	"file_name": ("attachmentfilenames", "attachment filenames", "attachment filename", "filename", "file name", "file", "attachment"),
	"confidence": ("confidence",),
}

COPY_POLL_SECONDS = 1
COPY_TIMEOUT_SECONDS = 120


class SharePointImportBatch(Document):
	def validate(self):
		self.source_folder = (self.source_folder or "").strip().strip("/")

	def get_source(self):
		"""
		Archive location: the batch's own, else the one in SharePoint Settings.
		Returns dict(drive_id, folder_id, path); folder_id wins over path when set
		"""
		settings = frappe.get_single(SETTINGS)
		if self.source_folder or self.source_folder_id:
			return {"drive_id": self.source_drive_id or settings.sharepoint_drive_id,
				"folder_id": self.source_folder_id, "path": self.source_folder}
		if settings.archive_folder_path or settings.archive_folder_id:
			return {"drive_id": settings.archive_drive_id or settings.sharepoint_drive_id,
				"folder_id": settings.archive_folder_id, "path": settings.archive_folder_path}
		frappe.throw(_("No archive folder is set on this batch or in SharePoint Settings. Use Browse Folder to pick one."))

	@frappe.whitelist()
	def dry_run(self):
		self.enqueue_processing(dry_run=True)

	@frappe.whitelist()
	def run(self):
		self.enqueue_processing(dry_run=False)

	@frappe.whitelist()
	def generate_mapping(self):
		if self.status in ("Queued", "Running"):
			frappe.throw(_("This batch is already being processed"))
		if not self.xero_export_file:
			frappe.throw(_("Attach the Xero data export first"))
		if not self.match_document_types:
			frappe.throw(_("Choose at least one document type to match against"))

		self.db_set({"status": "Queued", "error": None})
		frappe.enqueue(
			"frappe_sharepoint.sharepoint.doctype.sharepoint_import_batch.sharepoint_import_batch.generate_mapping_job",
			queue="long",
			timeout=-1,
			job_id=f"sharepoint_import_generate::{self.name}",
			deduplicate=True,
			enqueue_after_commit=True,
			batch_name=self.name,
		)

	def enqueue_processing(self, dry_run):
		if self.status in ("Queued", "Running"):
			frappe.throw(_("This batch is already being processed"))

		from frappe_sharepoint.controllers.file_controller import is_sync_enabled
		if not dry_run and not is_sync_enabled():
			frappe.throw(_("SharePoint file sync is not enabled in SharePoint Settings"))
		self.get_source()

		self.db_set({"status": "Queued", "error": None})
		frappe.enqueue(
			"frappe_sharepoint.sharepoint.doctype.sharepoint_import_batch.sharepoint_import_batch.process_batch",
			queue="long",
			timeout=-1,
			job_id=f"sharepoint_import_batch::{self.name}",
			deduplicate=True,
			enqueue_after_commit=True,
			batch_name=self.name,
			dry_run=dry_run,
		)

	def read_mapping(self):
		"""
		Parse the attached CSV into rows of
		dict(document_type, document_name, folder, file_name, line)
		"""
		if not self.mapping_file:
			frappe.throw(_("Please attach a mapping file"))

		file_doc = frappe.get_doc("File", {"file_url": self.mapping_file})
		with open(file_doc.get_full_path(), newline="", encoding="utf-8-sig") as f:
			reader = csv.reader(f)
			header = next(reader, None)
			if header is None:
				frappe.throw(_("The mapping file is empty"))

			columns = self.detect_columns(header)
			rows = []
			for line_no, values in enumerate(reader, start=2):
				if not any(v.strip() for v in values):
					continue

				def get(key):
					idx = columns.get(key)
					return values[idx].strip() if idx is not None and idx < len(values) else ""

				rows.append({
					"line": line_no,
					"document_type": get("document_type") or self.default_document_type,
					"document_name": get("document_name"),
					"folder": get("folder").strip("/"),
					"file_name": get("file_name"),
					"confidence": get("confidence"),
				})

		if not rows:
			frappe.throw(_("The mapping file has no data rows"))

		return rows

	def detect_columns(self, header):
		"""
		Map our column keys to CSV column indexes using the header aliases.
		A trailing unnamed column is taken as the file name, which is how the
		Xero export CSVs come
		"""
		normalized = [h.strip().lower() for h in header]
		columns = {}
		for key, aliases in COLUMN_ALIASES.items():
			for idx, name in enumerate(normalized):
				if name in aliases and idx not in columns.values():
					columns[key] = idx
					break

		if "file_name" not in columns and normalized and normalized[-1] == "":
			columns["file_name"] = len(normalized) - 1

		missing = [k for k in ("document_name", "file_name") if k not in columns]
		if missing:
			frappe.throw(_("Could not find these columns in the mapping file: {0}. Header found: {1}").format(
				", ".join(missing), ", ".join(h for h in header if h)))

		return columns


def process_batch(batch_name, dry_run):
	"""Background job: attach every row of the batch, or only check it on a dry run"""
	batch = frappe.get_doc("SharePoint Import Batch", batch_name)
	batch.db_set({"status": "Running", "last_run": now_datetime(), "last_run_type": "Dry Run" if dry_run else "Run"})
	frappe.db.commit()

	try:
		rows = batch.read_mapping()
		importer = ArchiveImporter(batch)
		results = []
		counts = {}

		for i, row in enumerate(rows, start=1):
			result = importer.process_row(row, dry_run)
			results.append(result)
			counts[result["status"]] = counts.get(result["status"], 0) + 1
			frappe.publish_progress(i * 100 / len(rows), title=_("SharePoint Import"), description=f"{i}/{len(rows)} {row['file_name']}")
			if not dry_run and i % 20 == 0:
				frappe.db.commit()

		batch.reload()
		batch.set("results", [])
		for r in results:
			batch.append("results", r)

		errors = counts.get("Failed", 0) + counts.get("Document Not Found", 0) + counts.get("File Not Found", 0)
		batch.update({
			"total_rows": len(rows),
			"attached": counts.get("Attached", 0),
			"already_attached": counts.get("Already Attached", 0),
			"ready": counts.get("Ready", 0),
			"document_not_found": counts.get("Document Not Found", 0),
			"file_not_found": counts.get("File Not Found", 0),
			"no_file_listed": counts.get("No File Listed", 0),
			"failed": counts.get("Failed", 0),
			"skipped": counts.get("Skipped", 0),
			"status": "Dry Run Complete" if dry_run else ("Completed with Errors" if errors else "Completed"),
			"error": None,
		})
		batch.flags.ignore_permissions = True
		batch.save()
		frappe.db.commit()

	except Exception as e:
		frappe.db.rollback()
		frappe.log_error("SharePoint Import Batch Error", frappe.get_traceback())
		frappe.db.set_value("SharePoint Import Batch", batch_name, {"status": "Failed", "error": str(e)[:500]})
		frappe.db.commit()

	frappe.publish_realtime("sharepoint_import_batch_done", {"name": batch_name}, doctype="SharePoint Import Batch", docname=batch_name)


def generate_mapping_job(batch_name):
	"""Background job: match the Xero export against SmartOps documents and attach the result as the mapping file"""
	batch = frappe.get_doc("SharePoint Import Batch", batch_name)
	batch.db_set({"status": "Running", "last_run": now_datetime(), "last_run_type": "Generate Mapping"})
	frappe.db.commit()

	try:
		export = frappe.get_doc("File", {"file_url": batch.xero_export_file})
		xero_docs = xero_matching.read_xero_export(export.get_full_path())
		if not xero_docs:
			frappe.throw(_("The Xero export has no rows with attachments"))

		dates = [x["date"] for x in xero_docs.values()]
		tolerance = int(batch.match_date_tolerance or 0)
		date_from, date_to = min(dates) - timedelta(days=tolerance), max(dates) + timedelta(days=tolerance)

		smartops_docs = []
		for row in batch.match_document_types:
			smartops_docs += xero_matching.load_smartops_documents(row.document_type, batch.match_company, date_from, date_to)

		rows = xero_matching.match(xero_docs, smartops_docs, tolerance)
		counts = xero_matching.summarize(rows)

		file_name = f"{frappe.scrub(batch.name)}-mapping.csv"
		path = frappe.get_site_path("private", "files", file_name)
		xero_matching.write_csv(rows, path)
		with open(path, "rb") as f:
			mapping = frappe.get_doc({
				"doctype": "File", "file_name": file_name, "is_private": 1,
				"attached_to_doctype": "SharePoint Import Batch", "attached_to_name": batch.name,
				"attached_to_field": "mapping_file", "content": f.read(),
			})
		os.remove(path)
		mapping.insert(ignore_permissions=True)

		summary = (f"{len(xero_docs)} Xero documents, {len(rows)} files, matched against {len(smartops_docs)} SmartOps documents.\n"
			+ "\n".join(f"{k}: {v}" for k, v in sorted(counts.items())))
		batch.reload()
		batch.update({"mapping_file": mapping.file_url, "generated_summary": summary, "status": "Draft", "error": None})
		batch.flags.ignore_permissions = True
		batch.save()
		frappe.db.commit()

	except Exception as e:
		frappe.db.rollback()
		frappe.log_error("SharePoint Import Generate Mapping Error", frappe.get_traceback())
		frappe.db.set_value("SharePoint Import Batch", batch_name, {"status": "Failed", "error": str(e)[:500]})
		frappe.db.commit()

	frappe.publish_realtime("sharepoint_import_batch_done", {"name": batch_name}, doctype="SharePoint Import Batch", docname=batch_name)


class ArchiveImporter:
	"""
	Copies files from the archive folder in SharePoint into each document's
	folder (server side, nothing is downloaded) and attaches them in Frappe
	"""

	def __init__(self, batch):
		self.batch = batch
		self.settings = frappe.get_single(SETTINGS)
		self.graph = self.settings.graph_api_url
		self.dest_drive = self.settings.sharepoint_drive_id
		self.source = batch.get_source()
		self.source_drive = self.source["drive_id"]
		self._headers = None
		self._folder_cache = {}

	def headers(self):
		if not self._headers:
			self._headers = get_request_header(self.settings)
		return dict(self._headers, **{"Content-Type": "application/json"})

	def process_row(self, row, dry_run):
		result = {
			"document_type": row["document_type"], "document_name": row["document_name"],
			"folder": row["folder"], "file_name": row["file_name"], "status": "Pending", "message": "",
		}

		def done(status, message=""):
			result.update({"status": status, "message": message})
			return result

		if not row["file_name"]:
			return done("No File Listed", _("Row {0} has no file name").format(row["line"]))

		if row.get("confidence") not in xero_matching.IMPORTABLE_CONFIDENCE:
			return done("Skipped", _("Confidence is '{0}'. Change it to Confirmed once checked").format(row["confidence"]))

		if not row["document_type"] or not frappe.db.exists("DocType", row["document_type"]):
			return done("Document Not Found", _("Unknown document type {0}").format(row["document_type"]))

		if not row["document_name"] or not frappe.db.exists(row["document_type"], row["document_name"]):
			return done("Document Not Found")

		existing = self.find_existing_attachment(row)
		if existing:
			result["sharepoint_url"] = existing
			return done("Already Attached")

		try:
			source = self.get_source_item(row)
		except Exception as e:
			return done("Failed", str(e)[:300])
		if not source:
			return done("File Not Found", f"{self.source['path']}/{row['folder']}/{row['file_name']}")

		if dry_run:
			return done("Ready")

		try:
			target_folder_id = self.get_target_folder(row)
			if not target_folder_id:
				return done("Failed", _("Could not create the document folder in SharePoint"))

			item = self.copy_item(source["id"], target_folder_id, row["file_name"])
			web_url = item.get("webUrl")
			if not web_url:
				return done("Failed", _("Copy finished without a web URL"))

			self.attach(row, web_url)
			result["sharepoint_url"] = web_url
			return done("Attached")
		except Exception as e:
			frappe.log_error("SharePoint Import Row Error", f"{row}\n{frappe.get_traceback()}")
			return done("Failed", str(e)[:300])

	def find_existing_attachment(self, row):
		"""A File already attached to the document with this file name"""
		return frappe.db.get_value("File", {
			"attached_to_doctype": row["document_type"],
			"attached_to_name": row["document_name"],
			"file_name": row["file_name"],
		}, "file_url")

	def get_source_item(self, row):
		"""Drive item of the archive file, or None when it does not exist"""
		rel = "/".join(p for p in (row["folder"], row["file_name"]) if p)
		if self.source["folder_id"]:
			# By id: keeps working if the archive folder is renamed or moved
			url = f"{self.graph}/drives/{self.source_drive}/items/{self.source['folder_id']}:/{quote(rel)}"
		else:
			url = f"{self.graph}/drives/{self.source_drive}/root:/{quote(self.source['path'] + '/' + rel)}"
		response = make_request("GET", url, self.headers(), None)
		if response.status_code == 404:
			return None
		if not response.ok:
			raise Exception(f"Archive lookup failed ({response.status_code}): {response.text[:200]}")
		return response.json()

	def get_target_folder(self, row):
		"""Document folder id in the destination library, built with the app's folder rules"""
		key = (row["document_type"], row["document_name"])
		if key not in self._folder_cache:
			sharepoint = SharePoint(doctype=row["document_type"], docname=row["document_name"])
			self._folder_cache[key] = sharepoint.build_folder_structure()
		return self._folder_cache[key]

	def copy_item(self, item_id, target_folder_id, file_name):
		"""
		Server-side copy. Graph answers 202 with a monitor URL that is polled,
		without authorization headers, until the copy completes
		"""
		url = (f"{self.graph}/drives/{self.source_drive}/items/{item_id}/copy"
			"?@microsoft.graph.conflictBehavior=replace")
		body = {
			"parentReference": {"driveId": self.dest_drive, "id": target_folder_id},
			"name": sanitize_name(file_name),
		}
		response = make_request("POST", url, self.headers(), body)
		if not response.ok:
			raise Exception(f"Copy failed ({response.status_code}): {response.text[:200]}")

		monitor_url = response.headers.get("Location") if hasattr(response, "headers") else None
		if not monitor_url:
			# Synchronous copy, the item is in the body
			return response.json()

		deadline = time.time() + COPY_TIMEOUT_SECONDS
		while time.time() < deadline:
			status = requests.get(monitor_url, timeout=30, allow_redirects=False)
			if status.status_code in (301, 302, 303):
				# Finished copies redirect to the new item
				return self.get_item(status.headers.get("Location"))
			data = status.json() if status.content else {}
			if data.get("status") == "completed":
				return self.get_item(f"{self.graph}/drives/{self.dest_drive}/items/{data['resourceId']}")
			if data.get("status") == "failed":
				raise Exception(f"Copy failed: {data.get('error', {}).get('message', data)}")
			time.sleep(COPY_POLL_SECONDS)

		raise Exception("Copy timed out")

	def get_item(self, url):
		response = make_request("GET", url, self.headers(), None)
		if not response.ok:
			raise Exception(f"Could not read copied item ({response.status_code})")
		return response.json()

	def attach(self, row, web_url):
		"""File record pointing at SharePoint, flagged so the sync leaves it alone"""
		frappe.get_doc({
			"doctype": "File",
			"file_name": row["file_name"],
			"file_url": unquote(web_url),
			"attached_to_doctype": row["document_type"],
			"attached_to_name": row["document_name"],
			"is_private": 1,
			"uploaded_to_sharepoint": 1,
		}).insert(ignore_permissions=True)
