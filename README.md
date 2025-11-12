<div align="center">
    <h1>Frappe SharePoint Integration</h1>
</div>

A universal SharePoint file synchronization solution for Frappe/ERPNext. This app automatically uploads files from your Frappe system to your SharePoint site, with flexible folder structure options and OAuth2 authentication.

## Features

- **Universal SharePoint Integration**: Connect to any SharePoint site using your own Azure AD tenant
- **Automatic File Sync**: Automatically upload files to SharePoint when they're attached to documents
- **Flexible Folder Structure**: Choose between hierarchical (Module/DocType/Document) or flat folder organization
- **OAuth2 Authentication**: Secure authentication using Frappe Connected Apps
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
   - **Redirect URI**: Web → `https://[your-frappe-site]/api/method/frappe.integrations.oauth2.authorize_redirect`

<img src="./app_registration.png" height="480">

4. After creation, note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

5. Go to "Certificates & secrets" → Create a new client secret
   - Note down the **Value** (you won't be able to see it again)

6. Go to "API permissions" → Add the following Microsoft Graph **Delegated** permissions:
   - `Files.ReadWrite.All`
   - `Sites.Read.All` (or `Sites.ReadWrite.All`)
   - `offline_access`
   - `User.Read`

7. Click "Grant admin consent" for your organization

### 2. Configure Connected App in Frappe

1. Go to **Connected App** doctype in Frappe
2. Create a new Connected App with these settings:
   - **App Name**: Microsoft SharePoint
   - **Provider**: Custom
   - **Client ID**: [Application (client) ID from Azure]
   - **Client Secret**: [Secret value from Azure]
   - **Authorization URL**: `https://login.microsoftonline.com/[tenant-id]/oauth2/v2.0/authorize`
   - **Token URL**: `https://login.microsoftonline.com/[tenant-id]/oauth2/v2.0/token`
   - **Scopes**: `Files.ReadWrite.All Sites.Read.All offline_access User.Read`
   - **Redirect URI**: `https://[your-frappe-site]/api/method/frappe.integrations.oauth2.authorize_redirect`

3. Save the Connected App

### 3. Authorize a User

1. Click "Connect User" in the Connected App
2. Select the user who will authenticate with Microsoft
3. You'll be redirected to Microsoft login
4. Grant the requested permissions
5. You'll be redirected back to Frappe

### 4. Configure SharePoint Settings

1. Go to **SharePoint Settings** in Frappe
2. Fill in the following fields:

   **OAuth Settings:**
   - **Connected App**: Select the Connected App you created
   - **Connected User**: Select the user you authorized
   - **Graph API URL**: `https://graph.microsoft.com/v1.0` (default)

   **SharePoint Configuration:**
   - **SharePoint Site URL**: Full URL of your SharePoint site (e.g., `https://yourtenant.sharepoint.com/sites/YourSite`)
   - Click **Fetch SharePoint Details** button to automatically retrieve Site ID and Drive ID
   - **Root Folder Path**: (Optional) Specify a root folder within the drive (e.g., `FrappeFiles`)

   **File Sync Settings:**
   - **Enable File Sync**: Check to enable automatic file upload
   - **Replace File Link**: Check to replace local files with SharePoint links (saves local storage)
   - **Folder Structure**: Choose between:
     - `Module/DocType/Document`: Creates hierarchical folders
     - `Flat`: Uploads all files to root folder

<img src="./m365_settings.png" height="580">

3. Save the settings

---

## Usage

Once configured, the app will automatically:

1. Upload any new files attached to Frappe documents to SharePoint
2. Create the folder structure based on your settings
3. Mark files as "Uploaded to SharePoint"
4. Optionally replace the local file with a SharePoint link

### Folder Structure Examples

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
2. Verify the Connected User has a valid OAuth token
3. Check Error Log in Frappe for specific error messages
4. Ensure the SharePoint Site ID and Drive ID are correctly fetched

### Permission errors?

1. Verify all required Microsoft Graph permissions are granted
2. Ensure admin consent was granted in Azure AD
3. Check that the Connected User has access to the SharePoint site

### Can't fetch SharePoint details?

1. Verify the SharePoint Site URL is correct
2. Check that the OAuth token is valid
3. Ensure the user has access to the specified SharePoint site

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
