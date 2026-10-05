// Copyright (c) 2026, Frappe Community and contributors
// For license information, please see license.txt

frappe.ui.form.on('SharePoint Import Batch', {
	onload(frm) {
		// New batches start from the archive configured in SharePoint Settings
		if (frm.is_new() && !frm.doc.source_folder) {
			frappe.db.get_doc('SharePoint Settings').then((s) => {
				if (s.archive_folder_path) {
					frm.set_value({
						source_folder: s.archive_folder_path,
						source_folder_id: s.archive_folder_id,
						source_drive_id: s.archive_drive_id || s.sharepoint_drive_id,
					});
				}
			});
		}
	},

	refresh(frm) {
		frm.set_intro('');
		if (!['Queued', 'Running'].includes(frm.doc.status)) {
			frm.add_custom_button(__('Browse Folder'), () => {
				frappe_sharepoint_picker.pick_folder({
					title: __('Select Archive Folder'),
					drive_id: frm.doc.source_drive_id || null,
					folder_id: frm.doc.source_folder_id || null,
					on_select(folder) {
						frm.set_value({source_folder: folder.path, source_folder_id: folder.id, source_drive_id: folder.drive_id});
					},
				});
			});
		}
		if (frm.doc.generated_summary && frm.doc.mapping_file) {
			frm.set_intro(__('Mapping generated. Download it from the Mapping File field, fix any Review rows (set Confidence to Confirmed), re-attach if changed, then Dry Run.'), 'blue');
		}
		if (!frm.doc.source_folder) {
			frm.set_intro(__('No archive folder set here or in SharePoint Settings. Use Browse Folder.'), 'orange');
		}
		if (['Queued', 'Running'].includes(frm.doc.status)) {
			frm.set_intro(__('Processing in the background. This page refreshes when it finishes.'), 'blue');
		} else if (!frm.is_new()) {
			frm.add_custom_button(__('Generate Mapping'), () => {
				if (frm.is_dirty()) {
					frappe.msgprint(__('Please save the batch first'));
					return;
				}
				frappe.confirm(
					__('Match the Xero export against the selected document types and replace this batch\'s Mapping File with the result?'),
					() => frappe.call({method: 'generate_mapping', doc: frm.doc, freeze: true, callback: () => frm.reload_doc()})
				);
			}, __('SharePoint'));
			frm.add_custom_button(__('Dry Run'), () => run_batch(frm, 'dry_run'));
			frm.add_custom_button(__('Run'), () => {
				frappe.confirm(
					__('Copy the listed files into each document\'s SharePoint folder and attach them?<br>Rows already attached are skipped.'),
					() => run_batch(frm, 'run')
				);
			}).addClass('btn-primary');
		}

		frappe.realtime.off('sharepoint_import_batch_done');
		frappe.realtime.on('sharepoint_import_batch_done', (data) => {
			if (data.name === frm.doc.name) {
				frm.reload_doc();
			}
		});
	},
});

function run_batch(frm, method) {
	if (frm.is_dirty()) {
		frappe.msgprint(__('Please save the batch first'));
		return;
	}
	frappe.call({
		method,
		doc: frm.doc,
		freeze: true,
		callback: () => {
			frappe.show_alert({message: __('Queued. Progress is shown at the top of the page.'), indicator: 'blue'});
			frm.reload_doc();
		},
	});
}
