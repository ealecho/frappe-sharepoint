import frappe
from frappe import _
from frappe.utils.pdf import get_pdf
import os
import tempfile

SETTINGS = "SharePoint Settings"


@frappe.whitelist()
def upload_document_to_sharepoint(doctype, docname):
	"""
	Upload document PDF along with all attachments to SharePoint
	
	Args:
		doctype: Document type (e.g., "Expense Claim")
		docname: Document name (e.g., "HR-EXP-2025-00033")
	
	Returns:
		dict: Upload status and SharePoint folder URL
	"""
	try:
		# Check if SharePoint sync is enabled
		settings = frappe.get_single(SETTINGS)
		if not settings.enable_file_sync:
			frappe.throw(_("SharePoint file sync is not enabled in SharePoint Settings"))
		
		# Generate document PDF
		pdf_file_path = generate_document_pdf(doctype, docname)
		
		# Get all attachments for the document
		attachments = get_document_attachments(doctype, docname)
		
		# Prepare files list for upload
		files_to_upload = []
		
		# Add PDF to upload list
		if pdf_file_path:
			files_to_upload.append({
				'filepath': pdf_file_path,
				'filename': f"{docname}.pdf",
				'is_temp': True  # Mark for cleanup after upload
			})
		
		# Add attachments to upload list
		for attachment in attachments:
			files_to_upload.append({
				'filepath': attachment['file_path'],
				'filename': attachment['file_name'],
				'is_temp': False,
				'file_doc': attachment['name']
			})
		
		if not files_to_upload:
			frappe.msgprint(_("No files to upload. Document has no attachments."))
			return {'success': False, 'message': 'No files to upload'}
		
		# Upload to SharePoint
		from frappe_sharepoint.utils.sharepoint import upload_document_bundle
		result = upload_document_bundle(
			doctype=doctype,
			docname=docname,
			files=files_to_upload
		)
		
		# Cleanup temporary PDF file
		if pdf_file_path and os.path.exists(pdf_file_path):
			os.remove(pdf_file_path)
		
		if result.get('success'):
			frappe.msgprint(
				_("Document uploaded to SharePoint successfully!<br>Files uploaded: {0}").format(
					len(files_to_upload)
				),
				indicator='green',
				title=_('Upload Successful')
			)
		
		return result
		
	except Exception as e:
		frappe.log_error("Document SharePoint Upload Error", str(e))
		frappe.throw(_("Failed to upload document to SharePoint: {0}").format(str(e)))


def generate_document_pdf(doctype, docname):
	"""
	Generate PDF for a document using its print format
	
	Args:
		doctype: Document type
		docname: Document name
	
	Returns:
		str: Path to temporary PDF file
	"""
	try:
		# Get the document
		doc = frappe.get_doc(doctype, docname)
		
		# Generate PDF using standard print format
		# frappe.get_print() returns the HTML content for the print format
		html_content = frappe.get_print(doctype, docname, print_format="Standard")
		pdf_content = get_pdf(html_content)
		
		# Create temporary file
		temp_dir = tempfile.gettempdir()
		pdf_filename = f"{docname}.pdf"
		pdf_path = os.path.join(temp_dir, pdf_filename)
		
		# Write PDF to temporary file
		with open(pdf_path, 'wb') as f:
			f.write(pdf_content)
		
		return pdf_path
		
	except Exception as e:
		frappe.log_error("PDF Generation Error", str(e))
		frappe.msgprint(_("Failed to generate PDF: {0}").format(str(e)), indicator='red')
		return None


def get_document_attachments(doctype, docname):
	"""
	Get all file attachments for a document
	
	Args:
		doctype: Document type
		docname: Document name
	
	Returns:
		list: List of attachment details
	"""
	try:
		# Query all files attached to the document
		files = frappe.get_all(
			"File",
			filters={
				"attached_to_doctype": doctype,
				"attached_to_name": docname
			},
			fields=["name", "file_name", "file_url", "is_private"]
		)
		
		attachments = []
		for file_doc in files:
			# Get full file path
			file_path = get_file_path(file_doc)
			if file_path and os.path.exists(file_path):
				attachments.append({
					'name': file_doc.name,
					'file_name': file_doc.file_name,
					'file_path': file_path
				})
		
		return attachments
		
	except Exception as e:
		frappe.log_error("Get Attachments Error", str(e))
		return []


def get_file_path(file_doc):
	"""
	Get absolute file path for a File document
	
	Args:
		file_doc: File document dict with file_url and is_private
	
	Returns:
		str: Absolute file path
	"""
	try:
		# Check if file_url exists
		if not file_doc.get('file_url'):
			return None
		
		# Extract file path from URL
		file_url = file_doc.get('file_url')
		
		# Handle private and public files
		if file_doc.get('is_private'):
			# Private file: /private/files/filename.ext
			if '/private/files/' in file_url:
				filename = file_url.split('/private/files/')[-1]
				site_path = frappe.get_site_path('private', 'files', filename)
			else:
				return None
		else:
			# Public file: /files/filename.ext
			if '/files/' in file_url:
				filename = file_url.split('/files/')[-1]
				site_path = frappe.get_site_path('public', 'files', filename)
			else:
				return None
		
		# Return absolute path if file exists
		if os.path.exists(site_path):
			return site_path
		
		return None
		
	except Exception as e:
		frappe.log_error("File Path Error", str(e))
		return None
