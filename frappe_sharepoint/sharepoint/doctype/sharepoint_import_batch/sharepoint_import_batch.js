// Copyright (c) 2026, Frappe Community and contributors
// For license information, please see license.txt

frappe.ui.form.on('SharePoint Import Batch', {
	refresh(frm) {
		frm.set_intro('');
		if (['Queued', 'Running'].includes(frm.doc.status)) {
			frm.set_intro(__('Processing in the background. This page refreshes when it finishes.'), 'blue');
		} else if (!frm.is_new()) {
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
