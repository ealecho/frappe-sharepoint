from frappe import _

def get_data():
	return [
		{
			"module_name": "SharePoint",
			"category": "Modules",
			"type": "module",
			"label": _("SharePoint"),
			"color": "#0078D4",
			"icon": "octicon octicon-cloud-upload",
			"description": _("SharePoint file synchronization")
		}
	]
