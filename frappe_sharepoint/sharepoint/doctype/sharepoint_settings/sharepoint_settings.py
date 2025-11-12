# Copyright (c) 2023, Frappe Community and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document

class SharePointSettings(Document):
	@frappe.whitelist()
	def test_connection(self):
		"""Test connection to Microsoft Graph API with provided credentials"""
		try:
			from frappe_sharepoint.utils import get_access_token, make_request
			
			# Get access token using credentials
			access_token = get_access_token(self.tenant_id, self.client_id, self.get_password("client_secret"))
			
			if not access_token:
				frappe.throw(_("Failed to authenticate. Please check your credentials."))
			
			# Test API connection
			headers = {'Authorization': f'Bearer {access_token}'}
			test_url = f"{self.graph_api_url}/sites/root"
			response = make_request('GET', test_url, headers, None)
			
			if response and response.ok:
				frappe.msgprint(_("Connection successful! Credentials are valid."), indicator='green')
				return True
			else:
				frappe.throw(_("Connection failed. Please verify your credentials and permissions."))
		except Exception as e:
			frappe.log_error("SharePoint Connection Test Error", str(e))
			frappe.throw(_("Connection test failed: {0}").format(str(e)))
	

	@frappe.whitelist()
	def get_sharepoint_sites(self):
		"""Get all SharePoint sites in the tenant"""
		try:
			from frappe_sharepoint.utils import get_request_header, make_request
			
			headers = get_request_header(self)
			
			# Get all sites in the tenant
			sites_url = f"{self.graph_api_url}/sites?search=*"
			response = make_request('GET', sites_url, headers, None)
			
			if response and response.ok:
				data = response.json()
				sites = []
				
				for site in data.get('value', []):
					sites.append({
						'id': site.get('id'),
						'name': site.get('name'),
						'displayName': site.get('displayName'),
						'webUrl': site.get('webUrl'),
						'description': site.get('description', '')
					})
				
				return sites
			else:
				frappe.throw(_("Failed to fetch SharePoint sites"))
		except Exception as e:
			frappe.log_error("SharePoint Sites Fetch Error", str(e))
			frappe.throw(_("Error fetching SharePoint sites: {0}").format(str(e)))
	
	@frappe.whitelist()
	def get_site_drives(self, site_id):
		"""Get all document libraries (drives) for a specific site"""
		try:
			from frappe_sharepoint.utils import get_request_header, make_request
			
			headers = get_request_header(self)
			
			# Get all drives for the site
			drives_url = f"{self.graph_api_url}/sites/{site_id}/drives"
			response = make_request('GET', drives_url, headers, None)
			
			if response and response.ok:
				data = response.json()
				drives = []
				
				for drive in data.get('value', []):
					drives.append({
						'id': drive.get('id'),
						'name': drive.get('name'),
						'description': drive.get('description', ''),
						'driveType': drive.get('driveType'),
						'webUrl': drive.get('webUrl')
					})
				
				return drives
			else:
				frappe.throw(_("Failed to fetch drives for the site"))
		except Exception as e:
			frappe.log_error("SharePoint Drives Fetch Error", str(e))
			frappe.throw(_("Error fetching drives: {0}").format(str(e)))
	
	@frappe.whitelist()
	def get_drive_folders(self, drive_id, folder_path=None):
		"""Get folders in a drive at the specified path"""
		try:
			from frappe_sharepoint.utils import get_request_header, make_request
			
			headers = get_request_header(self)
			
			# Build URL based on whether we're at root or in a subfolder
			if folder_path and folder_path != '/':
				# Get children of specific folder
				folders_url = f"{self.graph_api_url}/drives/{drive_id}/root:{folder_path}:/children"
			else:
				# Get root level folders
				folders_url = f"{self.graph_api_url}/drives/{drive_id}/root/children"
			
			response = make_request('GET', folders_url, headers, None)
			
			if response and response.ok:
				data = response.json()
				folders = []
				
				for item in data.get('value', []):
					# Only return folders, not files
					if 'folder' in item:
						folders.append({
							'id': item.get('id'),
							'name': item.get('name'),
							'path': item.get('parentReference', {}).get('path', '') + '/' + item.get('name'),
							'webUrl': item.get('webUrl'),
							'childCount': item.get('folder', {}).get('childCount', 0)
						})
				
				return folders
			else:
				frappe.throw(_("Failed to fetch folders"))
		except Exception as e:
			frappe.log_error("SharePoint Folders Fetch Error", str(e))
			frappe.throw(_("Error fetching folders: {0}").format(str(e)))
