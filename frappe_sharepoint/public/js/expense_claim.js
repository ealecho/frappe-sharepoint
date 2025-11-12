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
			// Show loading dialog with progress indicator
			let dialog = frappe.msgprint({
				title: __('SharePoint Upload'),
				message: __('<div style="text-align: center; padding: 20px;">\
					<div class="progress">\
						<div class="progress-bar progress-bar-striped active" role="progressbar" \
							aria-valuenow="100" aria-valuemin="0" aria-valuemax="100" style="width: 100%">\
						</div>\
					</div>\
					<p style="margin-top: 15px; color: #8D99A6;">\
						<i class="fa fa-cloud-upload" style="margin-right: 8px;"></i>\
						Preparing files and uploading to SharePoint...\
					</p>\
					<p style="margin-top: 10px; font-size: 12px; color: #A8B4C0;">\
						This may take a few moments depending on file sizes\
					</p>\
				</div>'),
				indicator: 'blue',
				primary_action: null
			});
			
			frappe.call({
				method: 'frappe_sharepoint.utils.document_upload.upload_document_to_sharepoint',
				args: {
					doctype: frm.doctype,
					docname: frm.docname
				},
				callback: function(r) {
					// Close the progress dialog
					if (dialog) {
						dialog.hide();
					}
					
					if (r.message && r.message.success) {
						// Show success message with details
						let uploaded_files = r.message.uploaded_count || 0;
						let folder_url = r.message.folder_url || '';
						
						frappe.msgprint({
							title: __('Upload Successful'),
							message: __('<div style="padding: 10px;">\
								<p style="margin-bottom: 15px;">\
									<i class="fa fa-check-circle" style="color: #98D85B; margin-right: 8px;"></i>\
									<strong>{0} file(s)</strong> successfully uploaded to SharePoint\
								</p>\
								{1}\
							</div>', [
								uploaded_files,
								folder_url ? '<a href="' + folder_url + '" target="_blank" class="btn btn-primary btn-sm">\
									<i class="fa fa-external-link" style="margin-right: 5px;"></i>Open SharePoint Folder\
								</a>' : ''
							]),
							indicator: 'green',
							primary_action: {
								label: __('Close'),
								action: function() {
									frappe.hide_msgprint();
								}
							}
						});
						
						// Also show brief alert
						frappe.show_alert({
							message: __('Successfully uploaded to SharePoint'),
							indicator: 'green'
						}, 5);
					} else {
						// Show error message
						let error_msg = r.message && r.message.error 
							? r.message.error 
							: __('An unknown error occurred during upload');
						
						frappe.msgprint({
							title: __('Upload Failed'),
							message: __('<div style="padding: 10px;">\
								<p style="margin-bottom: 15px;">\
									<i class="fa fa-exclamation-triangle" style="color: #FF6B6B; margin-right: 8px;"></i>\
									Failed to upload to SharePoint\
								</p>\
								<p style="color: #8D99A6; font-size: 13px;">{0}</p>\
								<p style="margin-top: 15px; font-size: 12px; color: #A8B4C0;">\
									Check the error log for more details or contact your administrator\
								</p>\
							</div>', [error_msg]),
							indicator: 'red'
						});
					}
				},
				error: function(r) {
					// Close the progress dialog
					if (dialog) {
						dialog.hide();
					}
					
					// Show error message
					let error_msg = r.message || __('Network error or server unavailable');
					
					frappe.msgprint({
						title: __('Upload Failed'),
						message: __('<div style="padding: 10px;">\
							<p style="margin-bottom: 15px;">\
								<i class="fa fa-exclamation-triangle" style="color: #FF6B6B; margin-right: 8px;"></i>\
								Failed to upload to SharePoint\
							</p>\
							<p style="color: #8D99A6; font-size: 13px;">{0}</p>\
							<p style="margin-top: 15px; font-size: 12px; color: #A8B4C0;">\
								Check the error log for more details or contact your administrator\
							</p>\
						</div>', [error_msg]),
						indicator: 'red'
					});
				}
			});
		}
	);
}
