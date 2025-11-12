frappe.ui.form.on('Expense Claim', {
	refresh(frm) {
		// Add "Upload to SharePoint" button
		if (!frm.is_new()) {
			frm.add_custom_button(__('Upload to SharePoint'), function() {
				upload_to_sharepoint(frm);
			}, __('SharePoint'));
		}
	}
});

function upload_to_sharepoint(frm) {
	frappe.confirm(
		__('Upload this Expense Claim document (PDF) and all attachments to SharePoint?'),
		function() {
			// User confirmed - proceed with upload
			frappe.call({
				method: 'frappe_sharepoint.utils.document_upload.upload_document_to_sharepoint',
				args: {
					doctype: frm.doctype,
					docname: frm.docname
				},
				freeze: true,
				freeze_message: __('Uploading to SharePoint...'),
				callback: function(r) {
					if (r.message && r.message.success) {
						frappe.show_alert({
							message: __('Successfully uploaded to SharePoint'),
							indicator: 'green'
						}, 5);
						
						// Show folder URL if available
						if (r.message.folder_url) {
							frappe.msgprint({
								title: __('Upload Successful'),
								message: __('Files uploaded: {0}<br>SharePoint Folder: <a href="{1}" target="_blank">Open Folder</a>', 
									[r.message.uploaded_count || 0, r.message.folder_url]),
								indicator: 'green'
							});
						}
					}
				},
				error: function() {
					frappe.msgprint({
						title: __('Upload Failed'),
						message: __('Failed to upload to SharePoint. Please check the error log.'),
						indicator: 'red'
					});
				}
			});
		}
	);
}
