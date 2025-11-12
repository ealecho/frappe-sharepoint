from __future__ import absolute_import

import frappe
from frappe import _
import requests

# Get access token using client credentials flow
def get_access_token(tenant_id, client_id, client_secret):
    """
    Authenticate with Azure AD using client credentials flow
    Returns access token for Microsoft Graph API
    """
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    
    try:
        response = requests.post(token_url, data=data)
        if response.ok:
            return response.json().get('access_token')
        else:
            frappe.log_error("Azure AD Token Error", response.text)
            return None
    except Exception as e:
        frappe.log_error("Azure AD Authentication Error", str(e))
        return None

# Make request headers with bearer token
def get_request_header(settings):
    """
    Generate authorization headers using Azure AD credentials
    """
    access_token = get_access_token(
        settings.tenant_id,
        settings.client_id,
        settings.get_password("client_secret")
    )
    
    if not access_token:
        frappe.throw(_("Failed to authenticate with Azure AD. Please check your credentials."))
    
    headers = {'Authorization': f'Bearer {access_token}'}
    return headers
    
# General API request handler
def make_request(request, url, headers, body=None):
    """
    Make HTTP requests to Microsoft Graph API
    """
    if request == 'POST':
        return requests.post(url, headers=headers, json=body)
    elif request == 'PATCH':
        return requests.patch(url, headers=headers, json=body)
    elif request == 'GET':
        return requests.get(url, headers=headers)
    elif request == 'DELETE':
        return requests.delete(url, headers=headers)
    elif request == "PUT":
        return requests.put(url, headers=headers, data=body)