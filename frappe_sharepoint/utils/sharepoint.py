import frappe
from frappe import _
from frappe_sharepoint.utils import get_request_header, make_request

import os
import re
from urllib.parse import quote

'''
	SharePoint file synchronization using Direct Drive API
'''

SETTINGS = "SharePoint Settings"
ContentType = {"Content-Type": "application/json"}
COMPANY_STRUCTURE = "Company/Module/DocType/Document"


def sanitize_name(name):
	"""Make a string safe to use as a SharePoint file or folder name"""
	name = re.sub(r'["*:<>?/\\|]', "-", str(name or ""))
	return name.strip().rstrip(".").strip() or "_"


def trigger_sharepoint_upload(doctype=None, docname=None, filepath=None, filedoc=None):
	"""Trigger SharePoint file upload"""
	file = frappe.db.get_value("File", filedoc, ["file_url", "uploaded_to_sharepoint"], as_dict=True)
	if not file or file.uploaded_to_sharepoint or is_remote_url(file.file_url):
		# Deleted, or already handled by another job
		return

	sharepoint = SharePoint(
		doctype=doctype,
		docname=docname, 
		filepath=filepath, 
		filedoc=filedoc
	)
	sharepoint.run_sharepoint_upload()


def is_remote_url(file_url):
	return (file_url or "").startswith(("http://", "https://"))


def upload_document_bundle(doctype, docname, files):
	"""
	Upload multiple files (document PDF + attachments) to SharePoint
	
	Args:
		doctype: Document type (e.g., "Expense Claim")
		docname: Document name (e.g., "HR-EXP-2025-00033")
		files: List of file dicts with keys: filepath, filename, is_temp
		
	Returns:
		dict: Upload status with success flag and SharePoint folder URL
	"""
	try:
		frappe.logger().info(f"[SharePoint Bundle] Starting upload for {doctype}: {docname} with {len(files)} files")
		
		sharepoint = SharePoint(doctype=doctype, docname=docname, filepath=None, filedoc=None)
		frappe.logger().info(f"[SharePoint Bundle] SharePoint instance created. Drive ID: {sharepoint.drive_id}")
		frappe.logger().info(f"[SharePoint Bundle] Root folder: {sharepoint.root_folder}, Folder structure: {sharepoint.folder_structure}")
		
		# Build the folder structure first
		frappe.logger().info(f"[SharePoint Bundle] Building folder structure...")
		target_folder_id = sharepoint.build_folder_structure()
		frappe.logger().info(f"[SharePoint Bundle] Target folder ID: {target_folder_id}")
		
		if not target_folder_id:
			frappe.logger().error(f"[SharePoint Bundle] Failed to determine target folder")
			return {
				'success': False,
				'message': 'Could not determine target folder in SharePoint'
			}
		
		# Upload each file
		uploaded_count = 0
		failed_files = []
		
		for idx, file_info in enumerate(files):
			filepath = file_info.get('filepath')
			filename = file_info.get('filename')
			
			frappe.logger().info(f"[SharePoint Bundle] File {idx+1}/{len(files)}: {filename}")
			frappe.logger().info(f"[SharePoint Bundle] File path: {filepath}")
			
			if not filepath or not filename:
				frappe.logger().warning(f"[SharePoint Bundle] Skipping file {idx+1} - missing filepath or filename")
				continue
			
			# Upload file with overwrite behavior
			frappe.logger().info(f"[SharePoint Bundle] Uploading {filename} to folder {target_folder_id}")
			item = sharepoint.upload_file_to_folder(
				target_folder_id=target_folder_id,
				filepath=filepath,
				filename=filename
			)
			
			if item:
				uploaded_count += 1
				frappe.logger().info(f"[SharePoint Bundle] Successfully uploaded {filename}")
				# Update File doc if this is an attachment
				if file_info.get('file_doc'):
					sharepoint.mark_file_uploaded(file_info['file_doc'], item, filepath)
					frappe.logger().info(f"[SharePoint Bundle] Marked File {file_info['file_doc']} as uploaded")
			else:
				failed_files.append(filename)
				frappe.logger().error(f"[SharePoint Bundle] Failed to upload {filename}")
		
		# Get SharePoint folder URL
		frappe.logger().info(f"[SharePoint Bundle] Getting folder URL for {target_folder_id}")
		folder_url = sharepoint.get_folder_url(target_folder_id)
		frappe.logger().info(f"[SharePoint Bundle] Folder URL: {folder_url}")
		
		if uploaded_count > 0:
			frappe.logger().info(f"[SharePoint Bundle] Upload completed: {uploaded_count} succeeded, {len(failed_files)} failed")
			return {
				'success': True,
				'uploaded_count': uploaded_count,
				'failed_count': len(failed_files),
				'folder_url': folder_url,
				'message': f'Successfully uploaded {uploaded_count} file(s) to SharePoint'
			}
		else:
			frappe.logger().error(f"[SharePoint Bundle] All uploads failed. Failed files: {failed_files}")
			return {
				'success': False,
				'message': 'Failed to upload files to SharePoint',
				'failed_files': failed_files
			}
			
	except Exception as e:
		frappe.logger().error(f"[SharePoint Bundle] Exception: {str(e)}")
		frappe.log_error("Document Bundle Upload Error", str(e))
		return {
			'success': False,
			'message': f'Error: {str(e)}'
		}


class SharePoint(object):
	def __init__(self, **kwargs):
		self.user = frappe.session.user
		self.doctype = kwargs.get("doctype")
		self.docname = kwargs.get("docname")
		self.filepath = kwargs.get("filepath")
		self.filedoc = kwargs.get("filedoc")
		self.settings = frappe.get_single(SETTINGS)
		self.root_folder = self.settings.root_folder_path or ""
		self.folder_structure = self.settings.folder_structure or "Module/DocType/Document"
		self.company_folder = self.get_company_folder()

		# A company can live in its own document library
		self.drive_id = (self.company_folder or {}).get("drive_id") or self.settings.sharepoint_drive_id
		if not self.drive_id:
			frappe.throw(_("SharePoint Drive ID not configured in SharePoint Settings"))

		self.base_url = f'{self.settings.graph_api_url}/drives/{self.drive_id}'
		self._headers = None

	def get_headers(self, content_type="application/json"):
		'''
			Authorization headers, token fetched once per upload run
		'''
		if not self._headers:
			self._headers = get_request_header(self.settings)
		headers = dict(self._headers)
		headers["Content-Type"] = content_type
		return headers

	def get_company_folder(self):
		'''
			Resolve the top-level (country) folder from the document's Company
			Returns dict(folder_name, drive_id) or None when the level is skipped
		'''
		if self.folder_structure != COMPANY_STRUCTURE:
			return None

		company = None
		if self.doctype == "Company":
			company = self.docname
		elif self.doctype and self.docname and frappe.get_meta(self.doctype).has_field("company"):
			company = frappe.db.get_value(self.doctype, self.docname, "company")

		if not company:
			if self.settings.default_company_folder:
				return {"folder_name": self.settings.default_company_folder, "drive_id": None}
			return None

		for row in self.settings.get("company_folders") or []:
			if row.company == company:
				return {"folder_name": row.folder_name, "drive_id": row.sharepoint_drive_id}

		return {"folder_name": company, "drive_id": None}

	def get_module_folder(self):
		'''
			Second-level folder: mapped name for the DocType, else its module
		'''
		for row in self.settings.get("doctype_folders") or []:
			if row.document_type == self.doctype:
				return row.folder_name

		return frappe.db.get_value("DocType", self.doctype, "module")

	def get_folder_segments(self):
		'''
			Folder names below the root folder, based on settings
		'''
		if self.folder_structure == "Flat":
			return []

		segments = []
		if self.company_folder:
			# Mapped folder may itself be a path, e.g. "East Africa/PEAS Uganda"
			segments += [s for s in self.company_folder["folder_name"].split("/") if s.strip()]

		segments += [self.get_module_folder(), self.doctype, self.docname]
		return [sanitize_name(s) for s in segments if s]

	def get_folder_id_by_name(self, parent_folder_id, folder_name):
		'''
			Get folder ID by name within a parent folder
		'''
		url = f'{self.base_url}/items/{parent_folder_id}:/{quote(folder_name)}'
		response = make_request('GET', url, self.get_headers(), None)
		if response.ok:
			return response.json()["id"]
		if response.status_code != 404:
			frappe.log_error("SharePoint folder lookup error", response.text)
		return None

	def create_sharepoint_folder(self, parent_folder_id, folder_name):
		'''
			Create a folder in SharePoint Drive
		'''
		frappe.logger().info(f"[Create Folder] Creating '{folder_name}' in parent {parent_folder_id}")
		
		url = f'{self.base_url}/items/{parent_folder_id}/children'
		body = {
			"name": f'{folder_name}',
			"folder": {},
			"@microsoft.graph.conflictBehavior": "fail"
		}

		response = make_request('POST', url, self.get_headers(), body)
		
		if response.status_code == 409:
			# Created by a parallel upload in the meantime
			return self.get_folder_id_by_name(parent_folder_id, folder_name)

		if not response.ok:
			frappe.logger().error(f"[Create Folder] Failed to create '{folder_name}': {response.text}")
			frappe.log_error("SharePoint folder creation error", response.text)
			return None

		folder_id = response.json()["id"]
		frappe.logger().info(f"[Create Folder] Successfully created '{folder_name}' with ID: {folder_id}")
		return folder_id

	def get_or_create_folder(self, parent_folder_id, folder_name):
		'''
			Get existing folder or create new one
		'''
		folder_id = self.get_folder_id_by_name(parent_folder_id, folder_name)
		if not folder_id:
			folder_id = self.create_sharepoint_folder(parent_folder_id, folder_name)
		return folder_id

	def build_folder_structure(self):
		'''
			Build folder structure based on settings
			Returns the final folder ID where file should be uploaded
		'''
		segments = [sanitize_name(s) for s in self.root_folder.split("/") if s.strip()]
		segments += self.get_folder_segments()
		frappe.logger().info(f"[Build Folders] Target path: {'/'.join(segments) or '(drive root)'}")

		current_folder_id = "root"
		for segment in segments:
			current_folder_id = self.get_or_create_folder(current_folder_id, segment)
			if not current_folder_id:
				return None

		return current_folder_id

	def run_sharepoint_upload(self):
		'''
			Main upload function
		'''
		try:
			# Build the folder structure
			target_folder_id = self.build_folder_structure()
			
			if not target_folder_id:
				frappe.log_error("SharePoint Upload Error", "Could not determine target folder")
				return

			file_name = os.path.basename(self.filepath) if self.filepath else None
			if not file_name:
				frappe.log_error("SharePoint Upload Error", "File name is missing")
				return

			item = self.upload_file_to_folder(target_folder_id, self.filepath, file_name)
			if item:
				self.mark_file_uploaded(self.filedoc, item, self.filepath)
		
		except Exception as e:
			frappe.log_error("SharePoint Upload Error", str(e))

	def mark_file_uploaded(self, filedoc, item, filepath):
		'''
			Flag the File as uploaded and, if configured, link it to SharePoint
			and drop the local copy
		'''
		frappe.db.set_value("File", filedoc, "uploaded_to_sharepoint", 1)

		web_url = item.get('webUrl')
		if not (self.settings.replace_file_link and web_url):
			return

		file = frappe.db.get_value(
			"File", filedoc,
			["file_url", "attached_to_doctype", "attached_to_name", "attached_to_field"],
			as_dict=True
		)
		local_url = file.file_url

		if file.attached_to_field and not self.relink_attach_field(file, web_url):
			# The document still shows the local file, keep it
			return

		frappe.db.set_value("File", filedoc, "file_url", web_url)

		# Only drop the local copy once the new link is committed, a rollback
		# would otherwise leave the File pointing at a deleted path
		frappe.db.after_commit.add(lambda: self.remove_unreferenced_file(local_url, filepath))

	def relink_attach_field(self, file, web_url):
		'''
			A file uploaded through an Attach field is referenced by that field,
			point it at SharePoint too. Returns False when the file has to stay
			local: images (SharePoint links need a login, they would not render)
			and fields that cannot be resolved, e.g. inside a child table
		'''
		field = frappe.get_meta(file.attached_to_doctype).get_field(file.attached_to_field)
		if not field or field.fieldtype != "Attach":
			return False

		current = frappe.db.get_value(file.attached_to_doctype, file.attached_to_name, field.fieldname)
		if current != file.file_url:
			return False

		frappe.db.set_value(
			file.attached_to_doctype, file.attached_to_name, field.fieldname, web_url,
			update_modified=False
		)
		return True

	def remove_unreferenced_file(self, local_url, filepath):
		'''
			Frappe points identical uploads at one file on disk, keep it until
			every File using it has moved to SharePoint
		'''
		if not frappe.db.exists("File", {"file_url": local_url}):
			self.remove_file(filepath)

	def remove_file(self, filepath):
		'''
			Remove file from local filesystem after successful upload
		'''
		try:
			if filepath and os.path.exists(filepath):
				os.remove(filepath)
		except Exception as e:
			frappe.log_error("File remove error", str(e))
	
	def upload_file_to_folder(self, target_folder_id, filepath, filename):
		'''
			Upload a single file to a specific SharePoint folder
			
			Args:
				target_folder_id: SharePoint folder ID
				filepath: Local file path
				filename: Name for the file in SharePoint
				
			Returns:
				dict: SharePoint drive item if upload successful, None otherwise
		'''
		try:
			frappe.logger().info(f"[Upload File] Uploading {filename} from {filepath} to folder {target_folder_id}")
			
			with open(filepath, 'rb') as f:
				file_content = f.read()
			
			if not file_content:
				frappe.log_error("SharePoint Upload Error", f"File {filename} is empty")
				return None
			
			# Upload file with replace behavior
			url = f'{self.base_url}/items/{target_folder_id}:/{quote(sanitize_name(filename))}:/content'
			response = make_request('PUT', url, self.get_headers("application/octet-stream"), file_content)
			
			if not response.ok:
				frappe.log_error("SharePoint File Upload Error", f"File: {filename}, Status: {response.status_code}, Error: {response.text}")
				return None
			
			frappe.logger().info(f"[Upload File] Successfully uploaded {filename}")
			return response.json()
			
		except Exception as e:
			frappe.log_error("File Upload Error", f"File: {filename}, Error: {str(e)}")
			return None
	
	def get_folder_url(self, folder_id):
		'''
			Get web URL for a SharePoint folder
			
			Args:
				folder_id: SharePoint folder ID
				
			Returns:
				str: Web URL to the folder or None
		'''
		try:
			url = f'{self.base_url}/items/{folder_id}'
			response = make_request('GET', url, self.get_headers(), None)
			
			if response.ok:
				return response.json().get('webUrl')

			frappe.logger().error(f"[Get Folder URL] Failed to get URL: {response.text}")
			return None
			
		except Exception as e:
			frappe.log_error("Get Folder URL Error", str(e))
			return None
