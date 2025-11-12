// Copyright (c) 2023, Frappe Community and contributors
// For license information, please see license.txt

frappe.ui.form.on('SharePoint Settings', {
	refresh: function(frm) {
		// Add Test Connection button
		if (frm.doc.tenant_id && frm.doc.client_id && frm.doc.client_secret) {
			frm.add_custom_button(__('Test Connection'), function() {
				frappe.call({
					method: 'test_connection',
					doc: frm.doc,
					callback: function(r) {
						if (r.message) {
							frappe.show_alert({
								message: __('Connection successful!'),
								indicator: 'green'
							});
						}
					}
				});
			});
		}
		
		// Add Fetch SharePoint Details button
		if (frm.doc.enable_file_sync && frm.doc.sharepoint_site_url) {
			frm.add_custom_button(__('Fetch SharePoint Details'), function() {
				frm.call({
					method: 'fetch_sharepoint_details',
					doc: frm.doc,
					callback: function(r) {
						frm.refresh_field('sharepoint_site_id');
						frm.refresh_field('sharepoint_drive_id');
					}
				});
			});
		}
	}
});
