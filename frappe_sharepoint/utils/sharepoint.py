import frappe
from frappe import _
from frappe_sharepoint.utils import get_request_header, make_request

import os

'''
	SharePoint file synchronization using Direct Drive API
'''

SETTINGS = "SharePoint Settings"
ContentType = {"Content-Type": "application/json"}


def trigger_sharepoint_upload(doctype=None, docname=None, filepath=None, filedoc=None):
	"""Trigger SharePoint file upload"""
	sharepoint = SharePoint(
		doctype=doctype,
		docname=docname, 
		filepath=filepath, 
		filedoc=filedoc
	)
	sharepoint.run_sharepoint_upload()


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
		sharepoint = SharePoint(doctype=doctype, docname=docname, filepath=None, filedoc=None)
		
		# Build the folder structure first
		target_folder_id = sharepoint.build_folder_structure()
		
		if not target_folder_id:
			return {
				'success': False,
				'message': 'Could not determine target folder in SharePoint'
			}
		
		# Upload each file
		uploaded_count = 0
		failed_files = []
		
		for file_info in files:
			filepath = file_info.get('filepath')
			filename = file_info.get('filename')
			
			if not filepath or not filename:
				continue
			
			# Upload file with overwrite behavior
			success = sharepoint.upload_file_to_folder(
				target_folder_id=target_folder_id,
				filepath=filepath,
				filename=filename
			)
			
			if success:
				uploaded_count += 1
				# Update File doc if this is an attachment
				if file_info.get('file_doc'):
					frappe.db.set_value("File", file_info['file_doc'], "uploaded_to_sharepoint", 1)
			else:
				failed_files.append(filename)
		
		# Get SharePoint folder URL
		folder_url = sharepoint.get_folder_url(target_folder_id)
		
		if uploaded_count > 0:
			return {
				'success': True,
				'uploaded_count': uploaded_count,
				'failed_count': len(failed_files),
				'folder_url': folder_url,
				'message': f'Successfully uploaded {uploaded_count} file(s) to SharePoint'
			}
		else:
			return {
				'success': False,
				'message': 'Failed to upload files to SharePoint',
				'failed_files': failed_files
			}
			
	except Exception as e:
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
		
		# Validate required settings
		if not self.settings.sharepoint_drive_id:
			frappe.throw(_("SharePoint Drive ID not configured in SharePoint Settings"))
		
		self.drive_id = self.settings.sharepoint_drive_id
		self.root_folder = self.settings.root_folder_path or ""
		self.folder_structure = self.settings.folder_structure or "Module/DocType/Document"
		self.base_url = f'{self.settings.graph_api_url}/drives/{self.drive_id}'

	def get_sharepoint_folder_items(self, folder_id):
		'''
			Fetch folder contents from SharePoint Drive
		'''
		folder_items = []
		headers = get_request_header(self.settings)
		headers.update(ContentType)
		url = f'{self.base_url}/items/{folder_id}/children'

		response = make_request('GET', url, headers, None)
		if response.status_code == 200 or response.ok:
			for item in response.json()['value']:
				folder_items.append({"name": item["name"], "id": item["id"]})
		else:
			frappe.log_error("SharePoint folder items fetch error", response.text)

		return folder_items

	def create_sharepoint_folder(self, parent_folder_id, folder_name):
		'''
			Create a folder in SharePoint Drive
		'''
		headers = get_request_header(self.settings)
		headers.update(ContentType)
		url = f'{self.base_url}/items/{parent_folder_id}/children'
		body = {
			"name": f'{folder_name}',
			"folder": {},
			"@microsoft.graph.conflictBehavior": "rename"
		}

		response = make_request('POST', url, headers, body)
		if not response.ok:
			frappe.log_error("SharePoint folder creation error", response.text)
			return None
		else:
			return response.json()["id"]

	def get_folder_id_by_name(self, parent_folder_id, folder_name):
		'''
			Get folder ID by name within a parent folder
		'''
		folder_id = None
		folder_items = self.get_sharepoint_folder_items(parent_folder_id)
		for item in folder_items:
			if folder_name == item['name']:
				folder_id = item['id']
				break
		return folder_id

	def get_or_create_folder(self, parent_folder_id, folder_name):
		'''
			Get existing folder or create new one
		'''
		folder_id = self.get_folder_id_by_name(parent_folder_id, folder_name)
		if not folder_id:
			folder_id = self.create_sharepoint_folder(parent_folder_id, folder_name)
		return folder_id

	def get_root_folder_id(self):
		'''
			Get or create the root folder for uploads
		'''
		if self.root_folder:
			# Navigate to root folder path
			headers = get_request_header(self.settings)
			url = f'{self.base_url}/root:/{self.root_folder}'
			response = make_request('GET', url, headers, None)
			
			if response.ok:
				return response.json()["id"]
			else:
				# Create root folder if it doesn't exist
				return self.create_sharepoint_folder("root", self.root_folder)
		else:
			# Use drive root
			headers = get_request_header(self.settings)
			url = f'{self.base_url}/root'
			response = make_request('GET', url, headers, None)
			if response.ok:
				return response.json()["id"]
			return "root"

	def build_folder_structure(self):
		'''
			Build folder structure based on settings
			Returns the final folder ID where file should be uploaded
		'''
		current_folder_id = self.get_root_folder_id()

		if self.folder_structure == "Flat":
			# No additional folders, upload directly to root
			return current_folder_id
		
		# Module/DocType/Document structure
		doctype_module = frappe.db.get_value("DocType", {"name": self.doctype}, "module")
		
		# Create Module folder
		if doctype_module:
			module_id = self.get_or_create_folder(current_folder_id, doctype_module)
			current_folder_id = module_id
		
		# Create DocType folder
		doctype_id = self.get_or_create_folder(current_folder_id, self.doctype)
		current_folder_id = doctype_id
		
		# Create Document folder (docname)
		if self.docname:
			document_id = self.get_or_create_folder(current_folder_id, self.docname)
			current_folder_id = document_id
		
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

			# Get file content and name
			file_content = self.get_file_content()
			file_name = self.filepath.split("/")[-1] if self.filepath else None

			if not file_content or not file_name:
				frappe.log_error("SharePoint Upload Error", "File content or name is missing")
				return

			# Upload file
			headers = get_request_header(self.settings)
			headers.update({"Content-Type": "application/octet-stream"})
			url = f'{self.base_url}/items/{target_folder_id}:/{file_name}:/content'

			response = make_request('PUT', url, headers, file_content)
			
			if not response.ok:
				frappe.log_error("SharePoint File Upload Error", response.text)
			else:
				# Mark file as uploaded
				frappe.db.set_value("File", self.filedoc, "uploaded_to_sharepoint", 1)
				
				# Replace file link if configured
				if self.settings.replace_file_link:
					web_url = response.json().get('webUrl')
					if web_url:
						frappe.db.set_value("File", self.filedoc, "file_url", web_url)
						self.remove_file()
				
				frappe.msgprint(_("File uploaded to SharePoint successfully"))
		
		except Exception as e:
			frappe.log_error("SharePoint Upload Error", str(e))

	def get_file_content(self):
		'''
			Read file content from filesystem
		'''
		try:
			if self.filepath:
				return open(self.filepath, 'rb')
			return None
		except Exception as e:
			frappe.log_error('File read error', str(e))
			return None

	def remove_file(self):
		'''
			Remove file from local filesystem after successful upload
		'''
		try:
			if self.filepath:
				os.remove(self.filepath)
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
				bool: True if upload successful, False otherwise
		'''
		try:
			# Read file content
			with open(filepath, 'rb') as f:
				file_content = f.read()
			
			if not file_content:
				frappe.log_error("SharePoint Upload Error", f"File {filename} is empty")
				return False
			
			# Upload file with replace behavior
			headers = get_request_header(self.settings)
			headers.update({"Content-Type": "application/octet-stream"})
			url = f'{self.base_url}/items/{target_folder_id}:/{filename}:/content'
			
			response = make_request('PUT', url, headers, file_content)
			
			if not response.ok:
				frappe.log_error("SharePoint File Upload Error", f"File: {filename}, Error: {response.text}")
				return False
			
			return True
			
		except Exception as e:
			frappe.log_error("File Upload Error", f"File: {filename}, Error: {str(e)}")
			return False
	
	def get_folder_url(self, folder_id):
		'''
			Get web URL for a SharePoint folder
			
			Args:
				folder_id: SharePoint folder ID
				
			Returns:
				str: Web URL to the folder or None
		'''
		try:
			headers = get_request_header(self.settings)
			url = f'{self.base_url}/items/{folder_id}'
			
			response = make_request('GET', url, headers, None)
			
			if response.ok:
				return response.json().get('webUrl')
			
			return None
			
		except Exception as e:
			frappe.log_error("Get Folder URL Error", str(e))
			return None
