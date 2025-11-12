frappe.provide("frappe");

frappe.realtime.on("sharepoint_sync", function (output) {
    frappe.show_alert(output, 15);
});