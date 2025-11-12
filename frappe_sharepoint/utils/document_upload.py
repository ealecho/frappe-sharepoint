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
		frappe.logger().info(f"[SharePoint Upload] Starting upload for {doctype}: {docname}")
		
		# Check if SharePoint sync is enabled
		settings = frappe.get_single(SETTINGS)
		frappe.logger().info(f"[SharePoint Upload] SharePoint sync enabled: {settings.enable_file_sync}")
		
		if not settings.enable_file_sync:
			frappe.throw(_("SharePoint file sync is not enabled in SharePoint Settings"))
		
		# Generate document PDF
		frappe.logger().info(f"[SharePoint Upload] Generating PDF for {docname}")
		pdf_file_path = generate_document_pdf(doctype, docname)
		frappe.logger().info(f"[SharePoint Upload] PDF generated at: {pdf_file_path}")
		
		# Get all attachments for the document
		frappe.logger().info(f"[SharePoint Upload] Fetching attachments for {docname}")
		attachments = get_document_attachments(doctype, docname)
		frappe.logger().info(f"[SharePoint Upload] Found {len(attachments)} attachments: {[a['file_name'] for a in attachments]}")
		
		# Prepare files list for upload
		files_to_upload = []
		
		# Add PDF to upload list
		if pdf_file_path:
			files_to_upload.append({
				'filepath': pdf_file_path,
				'filename': f"{docname}.pdf",
				'is_temp': True  # Mark for cleanup after upload
			})
			frappe.logger().info(f"[SharePoint Upload] Added PDF to upload list: {docname}.pdf")
		else:
			frappe.logger().warning(f"[SharePoint Upload] PDF generation failed, no PDF to upload")
		
		# Add attachments to upload list
		for attachment in attachments:
			files_to_upload.append({
				'filepath': attachment['file_path'],
				'filename': attachment['file_name'],
				'is_temp': False,
				'file_doc': attachment['name']
			})
			frappe.logger().info(f"[SharePoint Upload] Added attachment: {attachment['file_name']}")
		
		if not files_to_upload:
			frappe.logger().warning(f"[SharePoint Upload] No files to upload for {docname}")
			frappe.msgprint(_("No files to upload. Document has no attachments."))
			return {'success': False, 'message': 'No files to upload'}
		
		frappe.logger().info(f"[SharePoint Upload] Total files to upload: {len(files_to_upload)}")
		
		# Upload to SharePoint
		from frappe_sharepoint.utils.sharepoint import upload_document_bundle
		frappe.logger().info(f"[SharePoint Upload] Calling upload_document_bundle with {len(files_to_upload)} files")
		result = upload_document_bundle(
			doctype=doctype,
			docname=docname,
			files=files_to_upload
		)
		frappe.logger().info(f"[SharePoint Upload] Upload result: {result}")
		
		# Cleanup temporary PDF file
		if pdf_file_path and os.path.exists(pdf_file_path):
			os.remove(pdf_file_path)
			frappe.logger().info(f"[SharePoint Upload] Cleaned up temp PDF: {pdf_file_path}")
		
		if result.get('success'):
			frappe.logger().info(f"[SharePoint Upload] Upload completed successfully for {docname}")
			frappe.msgprint(
				_("Document uploaded to SharePoint successfully!<br>Files uploaded: {0}").format(
					len(files_to_upload)
				),
				indicator='green',
				title=_('Upload Successful')
			)
		else:
			frappe.logger().error(f"[SharePoint Upload] Upload failed: {result.get('message')}")
		
		return result
		
	except Exception as e:
		frappe.logger().error(f"[SharePoint Upload] Exception occurred: {str(e)}")
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
		frappe.logger().info(f"[PDF Generation] Starting for {doctype}: {docname}")
		
		# Get the document
		doc = frappe.get_doc(doctype, docname)
		frappe.logger().info(f"[PDF Generation] Document fetched successfully")
		
		# Generate PDF using standard print format
		# frappe.get_print() returns the HTML content for the print format
		frappe.logger().info(f"[PDF Generation] Generating print HTML with Standard format")
		html_content = frappe.get_print(doctype, docname, print_format="Standard")
		frappe.logger().info(f"[PDF Generation] HTML content length: {len(html_content)} chars")
		
		frappe.logger().info(f"[PDF Generation] Converting HTML to PDF")
		pdf_content = get_pdf(html_content)
		frappe.logger().info(f"[PDF Generation] PDF content size: {len(pdf_content)} bytes")
		
		# Create temporary file
		temp_dir = tempfile.gettempdir()
		pdf_filename = f"{docname}.pdf"
		pdf_path = os.path.join(temp_dir, pdf_filename)
		frappe.logger().info(f"[PDF Generation] Saving to: {pdf_path}")
		
		# Write PDF to temporary file
		with open(pdf_path, 'wb') as f:
			f.write(pdf_content)
		
		frappe.logger().info(f"[PDF Generation] PDF saved successfully at {pdf_path}")
		return pdf_path
		
	except Exception as e:
		frappe.logger().error(f"[PDF Generation] Exception: {str(e)}")
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
		frappe.logger().info(f"[Get Attachments] Querying files for {doctype}: {docname}")
		
		# Query all files attached to the document
		files = frappe.get_all(
			"File",
			filters={
				"attached_to_doctype": doctype,
				"attached_to_name": docname
			},
			fields=["name", "file_name", "file_url", "is_private"]
		)
		
		frappe.logger().info(f"[Get Attachments] Found {len(files)} files in database")
		
		attachments = []
		for file_doc in files:
			frappe.logger().info(f"[Get Attachments] Processing file: {file_doc.file_name} (private={file_doc.is_private}, url={file_doc.file_url})")
			
			# Get full file path
			file_path = get_file_path(file_doc)
			
			if file_path and os.path.exists(file_path):
				attachments.append({
					'name': file_doc.name,
					'file_name': file_doc.file_name,
					'file_path': file_path
				})
				frappe.logger().info(f"[Get Attachments] Added file: {file_doc.file_name} at {file_path}")
			else:
				frappe.logger().warning(f"[Get Attachments] Skipped file (not found): {file_doc.file_name}, path={file_path}")
		
		frappe.logger().info(f"[Get Attachments] Total valid attachments: {len(attachments)}")
		return attachments
		
	except Exception as e:
		frappe.logger().error(f"[Get Attachments] Exception: {str(e)}")
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
			frappe.logger().warning(f"[Get File Path] No file_url for file: {file_doc.get('file_name')}")
			return None
		
		# Extract file path from URL
		file_url = file_doc.get('file_url')
		frappe.logger().info(f"[Get File Path] Processing URL: {file_url}")
		
		# Handle private and public files
		if file_doc.get('is_private'):
			# Private file: /private/files/filename.ext
			if '/private/files/' in file_url:
				filename = file_url.split('/private/files/')[-1]
				site_path = frappe.get_site_path('private', 'files', filename)
				frappe.logger().info(f"[Get File Path] Private file path: {site_path}")
			else:
				frappe.logger().warning(f"[Get File Path] Private file URL format not recognized: {file_url}")
				return None
		else:
			# Public file: /files/filename.ext
			if '/files/' in file_url:
				filename = file_url.split('/files/')[-1]
				site_path = frappe.get_site_path('public', 'files', filename)
				frappe.logger().info(f"[Get File Path] Public file path: {site_path}")
			else:
				frappe.logger().warning(f"[Get File Path] Public file URL format not recognized: {file_url}")
				return None
		
		# Return absolute path if file exists
		if os.path.exists(site_path):
			frappe.logger().info(f"[Get File Path] File exists at: {site_path}")
			return site_path
		else:
			frappe.logger().warning(f"[Get File Path] File does not exist at: {site_path}")
		
		return None
		
	except Exception as e:
		frappe.logger().error(f"[Get File Path] Exception: {str(e)}")
		frappe.log_error("File Path Error", str(e))
		return None
