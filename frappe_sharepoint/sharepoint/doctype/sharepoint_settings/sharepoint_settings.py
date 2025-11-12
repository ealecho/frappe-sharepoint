# Copyright (c) 2023, aptitudetech and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document

class SharePointSettings(Document):
	@frappe.whitelist()
	def fetch_sharepoint_details(self):
		"""Fetch SharePoint site and drive details from Graph API"""
		if not self.sharepoint_site_url:
			frappe.throw(_("Please provide SharePoint Site URL"))
		
		try:
			from frappe_sharepoint.utils import get_request_header, make_request
			
			headers = get_request_header(self)
			
			# Extract site path from URL
			# e.g., https://tenant.sharepoint.com/sites/SiteName -> /sites/SiteName
			parts = self.sharepoint_site_url.split('.com')
			if len(parts) > 1:
				site_path = parts[1]
				hostname = parts[0].replace('https://', '')
				
				# Get site details
				site_url = f"{self.graph_api_url}/sites/{hostname}.sharepoint.com:{site_path}"
				site_response = make_request('GET', site_url, headers, None)
				
				if site_response and site_response.ok:
					site_data = site_response.json()
					self.sharepoint_site_id = site_data.get('id')
					
					# Get default drive
					drive_url = f"{self.graph_api_url}/sites/{self.sharepoint_site_id}/drive"
					drive_response = make_request('GET', drive_url, headers, None)
					
					if drive_response and drive_response.ok:
						drive_data = drive_response.json()
						self.sharepoint_drive_id = drive_data.get('id')
						frappe.msgprint(_("SharePoint details fetched successfully"))
					else:
						frappe.throw(_("Failed to fetch drive details"))
				else:
					frappe.throw(_("Failed to fetch site details. Please check the URL."))
		except Exception as e:
			frappe.log_error("SharePoint Details Fetch Error", str(e))
			frappe.throw(_("Error fetching SharePoint details: {0}").format(str(e)))
