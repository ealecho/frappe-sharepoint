// Copyright (c) 2023, aptitudetech and contributors
// For license information, please see license.txt

frappe.ui.form.on('SharePoint Settings', {
	refresh: function(frm) {
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
