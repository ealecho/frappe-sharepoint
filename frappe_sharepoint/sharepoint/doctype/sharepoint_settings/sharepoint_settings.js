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
		
		// Add Browse SharePoint Sites button
		if (frm.doc.enable_file_sync && frm.doc.tenant_id && frm.doc.client_id && frm.doc.client_secret) {
			frm.add_custom_button(__('Browse SharePoint Sites'), function() {
				frappe_sharepoint.browse_sites(frm);
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

// SharePoint Browser functionality
frappe_sharepoint = {
	browse_sites: function(frm) {
		let selected_site = null;
		let selected_drive = null;
		let selected_folder = null;
		
		// Step 1: Show sites selector
		frappe.call({
			method: 'get_sharepoint_sites',
			doc: frm.doc,
			callback: function(r) {
				if (r.message && r.message.length > 0) {
					let sites = r.message;
					
					let d = new frappe.ui.Dialog({
						title: __('Select SharePoint Site'),
						fields: [
							{
								fieldtype: 'HTML',
								fieldname: 'sites_list',
								options: frappe_sharepoint.render_sites_list(sites)
							}
						],
						primary_action_label: __('Next'),
						primary_action: function() {
							if (!selected_site) {
								frappe.msgprint(__('Please select a site'));
								return;
							}
							d.hide();
							frappe_sharepoint.browse_drives(frm, selected_site);
						}
					});
					
					d.show();
					
					// Attach click handlers
					setTimeout(function() {
						d.$wrapper.find('.site-item').on('click', function() {
							d.$wrapper.find('.site-item').removeClass('selected');
							$(this).addClass('selected');
							selected_site = sites[$(this).data('index')];
						});
					}, 100);
				} else {
					frappe.msgprint(__('No SharePoint sites found in your tenant'));
				}
			}
		});
	},
	
	browse_drives: function(frm, site) {
		let selected_drive = null;
		
		frappe.call({
			method: 'get_site_drives',
			doc: frm.doc,
			args: {
				site_id: site.id
			},
			callback: function(r) {
				if (r.message && r.message.length > 0) {
					let drives = r.message;
					
					let d = new frappe.ui.Dialog({
						title: __('Select Document Library'),
						fields: [
							{
								fieldtype: 'HTML',
								fieldname: 'drives_list',
								options: frappe_sharepoint.render_drives_list(drives)
							}
						],
						primary_action_label: __('Next'),
						primary_action: function() {
							if (!selected_drive) {
								frappe.msgprint(__('Please select a document library'));
								return;
							}
							d.hide();
							frappe_sharepoint.browse_folders(frm, site, selected_drive);
						},
						secondary_action_label: __('Back'),
						secondary_action: function() {
							d.hide();
							frappe_sharepoint.browse_sites(frm);
						}
					});
					
					d.show();
					
					// Attach click handlers
					setTimeout(function() {
						d.$wrapper.find('.drive-item').on('click', function() {
							d.$wrapper.find('.drive-item').removeClass('selected');
							$(this).addClass('selected');
							selected_drive = drives[$(this).data('index')];
						});
					}, 100);
				} else {
					frappe.msgprint(__('No document libraries found for this site'));
				}
			}
		});
	},
	
	browse_folders: function(frm, site, drive, current_path = null) {
		let selected_folder = null;
		
		frappe.call({
			method: 'get_drive_folders',
			doc: frm.doc,
			args: {
				drive_id: drive.id,
				folder_path: current_path
			},
			callback: function(r) {
				let folders = r.message || [];
				
				let d = new frappe.ui.Dialog({
					title: __('Select Root Folder'),
					fields: [
						{
							fieldtype: 'HTML',
							fieldname: 'path_info',
							options: `<div style="padding: 10px; background: #f8f9fa; border-radius: 4px; margin-bottom: 10px;">
								<strong>Site:</strong> ${site.displayName}<br>
								<strong>Library:</strong> ${drive.name}<br>
								<strong>Current Path:</strong> ${current_path || '/'}<br>
								<small class="text-muted">Click a folder to browse inside, or click "Select This Folder" to use current location</small>
							</div>`
						},
						{
							fieldtype: 'HTML',
							fieldname: 'folders_list',
							options: frappe_sharepoint.render_folders_list(folders)
						}
					],
					primary_action_label: __('Select This Folder'),
					primary_action: function() {
						d.hide();
						frappe_sharepoint.confirm_selection(frm, site, drive, current_path || '/');
					},
					secondary_action_label: __('Back'),
					secondary_action: function() {
						d.hide();
						if (current_path && current_path !== '/') {
							// Go up one level
							let parent_path = current_path.substring(0, current_path.lastIndexOf('/')) || '/';
							frappe_sharepoint.browse_folders(frm, site, drive, parent_path);
						} else {
							// Go back to drives selection
							frappe_sharepoint.browse_drives(frm, site);
						}
					}
				});
				
				d.show();
				
				// Attach click handlers for folders
				setTimeout(function() {
					d.$wrapper.find('.folder-item').on('click', function() {
						let folder = folders[$(this).data('index')];
						d.hide();
						frappe_sharepoint.browse_folders(frm, site, drive, folder.path);
					});
				}, 100);
			}
		});
	},
	
	confirm_selection: function(frm, site, drive, folder_path) {
		frappe.confirm(
			__('Confirm your selection:<br><br>' +
			   '<strong>Site:</strong> {0}<br>' +
			   '<strong>Library:</strong> {1}<br>' +
			   '<strong>Root Folder:</strong> {2}<br><br>' +
			   'This will update the SharePoint Settings. Continue?',
			   [site.displayName, drive.name, folder_path]),
			function() {
				// Set values on form
				frm.set_value('sharepoint_site_url', site.webUrl);
				frm.set_value('sharepoint_site_id', site.id);
				frm.set_value('sharepoint_drive_id', drive.id);
				frm.set_value('root_folder_path', folder_path);
				
				frappe.show_alert({
					message: __('SharePoint settings updated successfully'),
					indicator: 'green'
				});
			}
		);
	},
	
	render_sites_list: function(sites) {
		let html = `<div style="max-height: 400px; overflow-y: auto;">`;
		
		sites.forEach(function(site, index) {
			html += `
				<div class="site-item" data-index="${index}" style="
					padding: 12px;
					border: 1px solid #d1d8dd;
					border-radius: 4px;
					margin-bottom: 8px;
					cursor: pointer;
					transition: all 0.2s;
				">
					<div style="font-weight: 600; margin-bottom: 4px;">${site.displayName || site.name}</div>
					<div style="font-size: 12px; color: #6c757d;">${site.webUrl}</div>
					${site.description ? `<div style="font-size: 11px; color: #8d99a6; margin-top: 4px;">${site.description}</div>` : ''}
				</div>
			`;
		});
		
		html += `</div>
		<style>
			.site-item:hover {
				background-color: #f8f9fa;
				border-color: #2490ef;
			}
			.site-item.selected {
				background-color: #e7f2ff;
				border-color: #2490ef;
			}
		</style>`;
		
		return html;
	},
	
	render_drives_list: function(drives) {
		let html = `<div style="max-height: 400px; overflow-y: auto;">`;
		
		drives.forEach(function(drive, index) {
			html += `
				<div class="drive-item" data-index="${index}" style="
					padding: 12px;
					border: 1px solid #d1d8dd;
					border-radius: 4px;
					margin-bottom: 8px;
					cursor: pointer;
					transition: all 0.2s;
				">
					<div style="font-weight: 600; margin-bottom: 4px;">${drive.name}</div>
					<div style="font-size: 12px; color: #6c757d;">Type: ${drive.driveType}</div>
					${drive.description ? `<div style="font-size: 11px; color: #8d99a6; margin-top: 4px;">${drive.description}</div>` : ''}
				</div>
			`;
		});
		
		html += `</div>
		<style>
			.drive-item:hover {
				background-color: #f8f9fa;
				border-color: #2490ef;
			}
			.drive-item.selected {
				background-color: #e7f2ff;
				border-color: #2490ef;
			}
		</style>`;
		
		return html;
	},
	
	render_folders_list: function(folders) {
		if (!folders || folders.length === 0) {
			return `<div style="padding: 20px; text-align: center; color: #6c757d;">
				<p>No folders found. You can select this location as your root folder.</p>
			</div>`;
		}
		
		let html = `<div style="max-height: 300px; overflow-y: auto;">`;
		
		folders.forEach(function(folder, index) {
			html += `
				<div class="folder-item" data-index="${index}" style="
					padding: 10px;
					border: 1px solid #d1d8dd;
					border-radius: 4px;
					margin-bottom: 8px;
					cursor: pointer;
					transition: all 0.2s;
					display: flex;
					align-items: center;
				">
					<svg style="width: 20px; height: 20px; margin-right: 8px; color: #2490ef;" viewBox="0 0 24 24" fill="currentColor">
						<path d="M10 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/>
					</svg>
					<div>
						<div style="font-weight: 500;">${folder.name}</div>
						<div style="font-size: 11px; color: #8d99a6;">${folder.childCount} items</div>
					</div>
				</div>
			`;
		});
		
		html += `</div>
		<style>
			.folder-item:hover {
				background-color: #f8f9fa;
				border-color: #2490ef;
			}
		</style>`;
		
		return html;
	}
};
