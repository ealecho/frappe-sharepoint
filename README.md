<div align="center">
    <h1>Frappe SharePoint Integration</h1>
</div>

A universal SharePoint file synchronization solution for Frappe/ERPNext. This app automatically uploads files from your Frappe system to your SharePoint site, with flexible folder structure options and OAuth2 authentication.

## Features

- **Universal SharePoint Integration**: Connect to any SharePoint site using your own Azure AD tenant
- **Automatic File Sync**: Automatically upload files to SharePoint when they're attached to documents
- **Flexible Folder Structure**: Choose between hierarchical (Module/DocType/Document) or flat folder organization
- **Simple Authentication**: Direct credential configuration with Azure AD App Registration
- **Optional File Replacement**: Keep files on SharePoint only or maintain local copies
- **Supports Frappe v13 and v14**

## Why Use SharePoint Integration?

**Centralized File Management:** Keep all your files in SharePoint for better organization and easier sharing with external stakeholders.

**Enhanced Security:** Leverage SharePoint's enterprise-grade security features and compliance tools.

**Better Collaboration:** Share files easily with team members using SharePoint's built-in sharing capabilities.

**Reduced Storage Costs:** Optionally remove local file copies after uploading to SharePoint to save storage space.

**Backup and Recovery:** Benefit from SharePoint's built-in versioning and backup capabilities.

---

## Installation

### Self Hosting:

```bash
# Get the app
bench get-app https://github.com/yourusername/frappe-sharepoint.git

# Install on your site
bench --site [your.site.name] install-app frappe_sharepoint

# Run migrations
bench --site [your.site.name] migrate

# Restart
bench restart
```

---

## Setup Instructions

### 1. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Configure your app:
   - **Name**: Frappe SharePoint Sync (or any name you prefer)
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Not required (leave blank)

<img src="./app_registration.png" height="480">

4. After creation, note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

5. Go to "Certificates & secrets" → Create a new client secret
   - Note down the **Value** (you won't be able to see it again)

6. Go to "API permissions" → Add the following Microsoft Graph **Application** permissions:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`

7. Click "Grant admin consent" for your organization

### 2. Configure SharePoint Settings

1. Go to **SharePoint Settings** in ERPNext
2. Fill in the following fields:

   **Azure AD Credentials:**
   - **Tenant ID**: Your Azure AD tenant ID
   - **Client ID**: Your app registration client ID
   - **Client Secret**: Your app registration client secret
   - Click **Test Connection** to verify your credentials

   **SharePoint Configuration:**
   - **Graph API URL**: `https://graph.microsoft.com/v1.0` (default)
   - **Enable File Sync**: Check to enable automatic file upload
   - **SharePoint Site URL**: Full URL of your SharePoint site (e.g., `https://yourtenant.sharepoint.com/sites/YourSite`)
   - Click **Fetch SharePoint Details** button to automatically retrieve Site ID and Drive ID
   - **Root Folder Path**: (Optional) Specify a root folder within the drive (e.g., `/Frappe Files`)

   **File Handling:**
   - **Store Files on SharePoint Only**: Check to link attachments to their SharePoint URL and remove the copy on the Frappe server once the upload succeeds. Users open these files directly in SharePoint, so they need access to the SharePoint site.
   - **Folder Structure**: Choose between:
     - `Company/Module/DocType/Document`: One top-level folder per company (e.g. per country), then hierarchical folders
     - `Module/DocType/Document`: Creates hierarchical folders
     - `Flat`: Uploads all files to root folder

   - **Excluded Document Types**: Attachments of these document types stay on the Frappe server and are never sent to SharePoint. Data Import, Bank Statement Import, Prepared Report, Letter Head, Package Import, Repost Item Valuation, Import Supplier Invoice, User Font and Communication (emails) are always excluded, because Frappe reads their files back from disk. Add your own import tools and template doctypes here.

   Files uploaded through an **Attach** field, on the document or in one of its child tables, are moved too and the field is pointed at the SharePoint URL. If the document was not saved yet when the upload ran, this happens on the next hourly retry. Files in **Attach Image** fields are copied to SharePoint but kept on the server, since a SharePoint link needs a login and would not render as an image.

   **Folder Mapping:** (optional, use **SharePoint > Load Default Mappings** for a starting point; folders are created automatically)
   - **Company Folders**: Map each Company to its top-level folder (e.g. `PEAS Uganda Ltd` → `PEAS Uganda`), optionally with its own Drive ID if that country uses a separate document library. Unmapped companies use the company name.
   - **Folder for Documents without a Company**: Used for doctypes that have no Company field. Leave blank to skip the company level for those.
   - **Document Type Folders**: Map a DocType to a friendlier second-level folder (e.g. `Expense Claim` → `Expenses`, `Purchase Order` → `Procurement`). Unmapped doctypes use their module name.

<img src="./m365_settings.png" height="580">

3. Save the settings

---

## Usage

Once configured, the app will automatically:

1. Upload any new files attached to Frappe documents to SharePoint
2. Create the folder structure based on your settings
3. Mark files as "Uploaded to SharePoint"
4. Optionally replace the local file with a SharePoint link
5. Retry uploads that did not complete (hourly, for files attached in the last 7 days)

Characters SharePoint does not allow in names (`" * : < > ? / \ |`) are replaced with `-`, so a document named `ACC-SINV/2025/0001` gets the folder `ACC-SINV-2025-0001`.

### Folder Structure Examples

**Company/Module/DocType/Document:**
```
SharePoint Drive
└── [Root Folder Path]
    └── PEAS Uganda                  (Company Folders mapping)
        └── Expenses                 (Document Type Folders mapping, else module name)
            └── Expense Claim
                └── HR-EXP-2025-00033
                    └── [File]
```

**Module/DocType/Document:**
```
SharePoint Drive
└── [Root Folder Path]
    └── [Module Name]
        └── [DocType Name]
            └── [Document Name]
                └── [File]
```

**Flat:**
```
SharePoint Drive
└── [Root Folder Path]
    └── [File]
```

---

## Troubleshooting

### Files not uploading?

1. Check that "Enable File Sync" is enabled in SharePoint Settings
2. Click "Test Connection" to verify your Azure AD credentials
3. Check Error Log in Frappe for specific error messages
4. Ensure the SharePoint Site ID and Drive ID are correctly fetched

### Development and staging sites

A database restored from production brings the production SharePoint Settings with it. To make sure such a site never syncs, add this to its `site_config.json` (it is not part of database backups, so it survives restores):

```json
"disable_sharepoint_sync": 1
```

This overrides **Enable File Sync**: no uploads, no retries, no local files removed. Remove the key, or point the site at a separate test document library, to test the integration.

### Permission errors?

1. Verify all required Microsoft Graph **Application** permissions are granted
2. Ensure admin consent was granted in Azure AD
3. Check that your Azure AD app has access to the SharePoint site

### Can't fetch SharePoint details?

1. Verify the SharePoint Site URL is correct
2. Click "Test Connection" to verify your credentials
3. Ensure your Azure AD app has proper permissions

---

## Dependencies

- [Frappe Framework](https://github.com/frappe/frappe) v13 or v14
- Microsoft 365 subscription with SharePoint Online
- Azure AD tenant with app registration permissions

---

## Bug Reports

Please create an issue on [GitHub Issues](https://github.com/yourusername/frappe-sharepoint/issues/new)

---

## License

MIT
