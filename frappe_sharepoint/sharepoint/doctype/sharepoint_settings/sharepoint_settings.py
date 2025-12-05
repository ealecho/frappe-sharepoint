# Copyright (c) 2023, Frappe Community and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document
from urllib.parse import urlparse

class SharePointSettings(Document):
	def validate(self):
		"""Validate settings before saving"""
		self.validate_root_folder_path()
	
	def validate_root_folder_path(self):
		"""Validate and sanitize root folder path"""
		if self.root_folder_path:
			# Strip leading and trailing whitespace
			self.root_folder_path = self.root_folder_path.strip()
			
			# Remove leading and trailing slashes
			self.root_folder_path = self.root_folder_path.strip('/')
			
			# If the path is now empty (was just "/"), clear it
			if not self.root_folder_path:
				self.root_folder_path = ""
				return
			
			# Check for invalid characters in folder names
			invalid_chars = ['\\', ':', '*', '?', '"', '<', '>', '|']
			for char in invalid_chars:
				if char in self.root_folder_path:
					frappe.throw(
						_("Root Folder Path contains invalid character: {0}").format(char),
						title=_("Invalid Path")
					)
			
			# Check for double slashes (invalid path)
			if '//' in self.root_folder_path:
				frappe.throw(
					_("Root Folder Path cannot contain consecutive slashes (//)"),
					title=_("Invalid Path")
				)
			
			# Validate path segments (no empty segments between slashes)
			segments = self.root_folder_path.split('/')
			for segment in segments:
				if not segment.strip():
					frappe.throw(
						_("Root Folder Path cannot have empty folder names"),
						title=_("Invalid Path")
					)
	
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
	
	def parse_sharepoint_url(self, url):
		"""
		Parse SharePoint URL into hostname and site path
		
		Args:
			url: SharePoint site URL (e.g., https://peasuk.sharepoint.com/sites/SmartOps)
			
		Returns:
			tuple: (hostname, site_path) e.g., ("peasuk.sharepoint.com", "/sites/SmartOps")
			
		Raises:
			frappe.ValidationError if URL format is invalid
		"""
		if not url:
			frappe.throw(
				_("SharePoint Site URL is required"),
				title=_("Missing URL")
			)
		
		# Parse the URL
		parsed = urlparse(url.strip())
		
		# Validate hostname
		hostname = parsed.netloc
		if not hostname:
			frappe.throw(
				_("Invalid SharePoint URL format. Expected format: https://yourtenant.sharepoint.com/sites/YourSite"),
				title=_("Invalid URL")
			)
		
		if not hostname.endswith('.sharepoint.com'):
			frappe.throw(
				_("Invalid SharePoint URL. Hostname must end with .sharepoint.com. Expected format: https://yourtenant.sharepoint.com/sites/YourSite"),
				title=_("Invalid URL")
			)
		
		# Get site path
		site_path = parsed.path.rstrip('/')
		
		# Validate site path - must have /sites/ or /teams/ prefix
		if not site_path:
			frappe.throw(
				_("SharePoint Site URL must include a site path. Expected format: https://yourtenant.sharepoint.com/sites/YourSite"),
				title=_("Invalid URL")
			)
		
		if not (site_path.startswith('/sites/') or site_path.startswith('/teams/')):
			frappe.throw(
				_("SharePoint Site URL must include /sites/ or /teams/ in the path. Expected format: https://yourtenant.sharepoint.com/sites/YourSite"),
				title=_("Invalid URL")
			)
		
		# Validate that there's a site name after /sites/ or /teams/
		path_parts = site_path.split('/')
		if len(path_parts) < 3 or not path_parts[2]:
			frappe.throw(
				_("SharePoint Site URL must include a site name after /sites/ or /teams/. Expected format: https://yourtenant.sharepoint.com/sites/YourSite"),
				title=_("Invalid URL")
			)
		
		return hostname, site_path
	
	@frappe.whitelist()
	def fetch_sharepoint_details(self):
		"""
		Fetch Site ID and available Drives from SharePoint Site URL.
		Uses direct site lookup which only requires access to the specific site,
		not the Sites.Read.All permission needed for site search.
		
		Returns:
			dict: {
				'site_id': str,
				'site_name': str,
				'drives': list of drive dicts
			}
		"""
		try:
			from frappe_sharepoint.utils import get_request_header, make_request
			
			# Parse and validate the SharePoint URL
			hostname, site_path = self.parse_sharepoint_url(self.sharepoint_site_url)
			
			frappe.logger().info(f"[Fetch Details] Parsed URL - hostname: {hostname}, site_path: {site_path}")
			
			headers = get_request_header(self)
			
			# Step 1: Get site details using direct path lookup
			# API: GET /sites/{hostname}:{site_path}
			site_url = f"{self.graph_api_url}/sites/{hostname}:{site_path}"
			frappe.logger().info(f"[Fetch Details] Fetching site from: {site_url}")
			
			response = make_request('GET', site_url, headers, None)
			
			if not response or not response.ok:
				error_msg = response.text if response else "No response from server"
				frappe.logger().error(f"[Fetch Details] Failed to fetch site: {error_msg}")
				frappe.throw(
					_("Failed to fetch SharePoint site details. Please verify the URL and your permissions. Error: {0}").format(error_msg),
					title=_("Site Fetch Failed")
				)
			
			site_data = response.json()
			site_id = site_data.get('id')
			site_name = site_data.get('displayName') or site_data.get('name')
			
			frappe.logger().info(f"[Fetch Details] Found site: {site_name} (ID: {site_id})")
			
			if not site_id:
				frappe.throw(
					_("Could not retrieve Site ID from SharePoint response"),
					title=_("Invalid Response")
				)
			
			# Step 2: Get available drives (document libraries) for the site
			drives_url = f"{self.graph_api_url}/sites/{site_id}/drives"
			frappe.logger().info(f"[Fetch Details] Fetching drives from: {drives_url}")
			
			drives_response = make_request('GET', drives_url, headers, None)
			
			drives = []
			if drives_response and drives_response.ok:
				drives_data = drives_response.json()
				for drive in drives_data.get('value', []):
					drives.append({
						'id': drive.get('id'),
						'name': drive.get('name'),
						'description': drive.get('description', ''),
						'driveType': drive.get('driveType'),
						'webUrl': drive.get('webUrl')
					})
				frappe.logger().info(f"[Fetch Details] Found {len(drives)} drives")
			else:
				frappe.logger().warning(f"[Fetch Details] Could not fetch drives: {drives_response.text if drives_response else 'No response'}")
			
			return {
				'site_id': site_id,
				'site_name': site_name,
				'drives': drives
			}
			
		except frappe.ValidationError:
			# Re-raise validation errors as-is
			raise
		except Exception as e:
			frappe.log_error("SharePoint Fetch Details Error", str(e))
			frappe.throw(_("Error fetching SharePoint details: {0}").format(str(e)))
