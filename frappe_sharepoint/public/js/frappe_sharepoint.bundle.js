frappe.provide("frappe");
frappe.provide("frappe_sharepoint_picker");

frappe.realtime.on("sharepoint_sync", function (output) {
    frappe.show_alert(output, 15);
});

/**
 * Folder picker for SharePoint. Browses by folder id so renamed folders keep
 * working. on_select receives {drive_id, id, name, path}.
 */
frappe_sharepoint_picker.pick_folder = function (opts) {
	const state = {drive_id: opts.drive_id || null, folder_id: opts.folder_id || null};
	let selected = null;

	const d = new frappe.ui.Dialog({
		title: opts.title || __("Select SharePoint Folder"),
		size: "large",
		fields: [
			{fieldtype: "HTML", fieldname: "path"},
			{fieldtype: "HTML", fieldname: "list"},
		],
		primary_action_label: __("Select This Folder"),
		primary_action() {
			if (!selected) {
				frappe.msgprint(__("Open the folder you want, or click one in the list, then select"));
				return;
			}
			d.hide();
			opts.on_select(selected);
		},
		secondary_action_label: __("Up One Level"),
		secondary_action() {
			if (state.parent_id !== undefined) {
				load(state.parent_id);
			}
		},
	});

	function load(folder_id) {
		frappe.call({
			method: "frappe_sharepoint.utils.sharepoint.list_folders",
			args: {drive_id: state.drive_id, folder_id: folder_id},
			freeze: true,
			callback(r) {
				const data = r.message;
				state.drive_id = data.drive_id;
				state.folder_id = data.current.id;
				state.parent_id = data.current.id ? data.parent_id : undefined;
				selected = data.current.id ? {drive_id: data.drive_id, id: data.current.id, name: data.current.name, path: data.current.path} : null;

				d.get_field("path").$wrapper.html(`
					<div style="padding:8px 10px;background:var(--bg-light-gray);border-radius:4px;margin-bottom:8px;">
						<strong>${__("Current folder")}:</strong> ${frappe.utils.escape_html(data.current.path || __("(library root)"))}
					</div>`);

				const rows = data.folders.map((f, i) => `
					<div class="sp-folder" data-i="${i}" style="padding:8px 10px;border:1px solid var(--border-color);border-radius:4px;margin-bottom:6px;cursor:pointer;display:flex;justify-content:space-between;">
						<span>📁 ${frappe.utils.escape_html(f.name)}</span>
						<span class="text-muted small">${f.childCount} ${__("items")}</span>
					</div>`).join("");
				d.get_field("list").$wrapper.html(`<div style="max-height:320px;overflow-y:auto;">${rows || `<p class="text-muted">${__("No sub-folders")}</p>`}</div>`);
				d.get_field("list").$wrapper.find(".sp-folder").on("click", function () {
					load(data.folders[$(this).data("i")].id);
				});
				d.get_secondary_btn().toggle(!!data.current.id);
			},
		});
	}

	d.show();
	load(state.folder_id);
};
