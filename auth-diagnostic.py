#!/usr/bin/env python
"""SharePoint authentication diagnostic script"""

import os
import sys
import json
import requests
from dotenv import load_dotenv

def run_auth_diagnostic():
    """Execute SharePoint authentication diagnostic"""
    print("=== SharePoint Authentication Diagnostic ===")
    
    # Check for .env file
    if not os.path.exists(".env"):
        print("❌ Error: .env file not found")
        print("   Please copy .env.example and configure it")
        return False
    
    # Load environment variables
    load_dotenv()
    
    # Check required variables
    required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "SITE_URL"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"❌ Error: The following environment variables are not set: {', '.join(missing_vars)}")
        return False
    
    print("✅ All required environment variables are set")
    
    # Check SharePoint site URL format
    site_url = os.getenv("SITE_URL")
    if not site_url.startswith("https://") or ".sharepoint.com/" not in site_url.lower():
        print(f"❌ Error: Invalid SharePoint site URL: {site_url}")
        print("   URL must be in the format: https://your-tenant.sharepoint.com/sites/your-site")
        return False
    
    print(f"✅ SharePoint site URL format is valid: {site_url}")
    
    # Test authentication
    print("\n--- Testing Authentication and Requests ---")
    
    try:
        import msal
        
        # Set up token cache
        cache = msal.SerializableTokenCache()
        
        # Create MSAL client application
        tenant_id = os.getenv("TENANT_ID")
        client_id = os.getenv("CLIENT_ID")
        client_secret = os.getenv("CLIENT_SECRET")
        
        print(f"Tenant ID: {tenant_id[:5]}...{tenant_id[-5:]}")
        print(f"Client ID: {client_id[:5]}...{client_id[-5:]}")
        
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret,
            token_cache=cache
        )
        
        # Try client credential flow
        print("Attempting client credential flow...")
        scope = ["https://graph.microsoft.com/.default"]
        result = app.acquire_token_for_client(scopes=scope)
        
        if "access_token" not in result:
            error_code = result.get("error", "unknown")
            error_description = result.get("error_description", "Unknown error")
            print(f"❌ Error: Authentication failed: {error_code} - {error_description}")
            
            # Detailed explanations for common errors
            if "AADSTS700016" in str(error_description):
                print("   Application not found or is disabled for the tenant")
                print("   Please check your application registration in Azure AD")
            elif "AADSTS7000215" in str(error_description):
                print("   Invalid client secret")
                print("   Check if your client secret is correct and not expired")
            elif "AADSTS650057" in str(error_description):
                print("   Invalid client credentials")
                print("   Check your client ID and client secret")
            elif "AADSTS70011" in str(error_description):
                print("   Application was not found in the tenant")
                print("   Make sure the application is registered in this tenant")
            
            print(f"\nFull error response: \n{json.dumps(result, indent=2)}")
            return False
            
        print("✅ Successfully obtained access token")
        
        # Try accessing SharePoint site
        print("Attempting to access SharePoint site...")
        
        headers = {
            "Authorization": f"Bearer {result['access_token']}",
            "Content-Type": "application/json",
        }
        
        # Extract domain and site name from site URL
        site_parts = site_url.replace("https://", "").split("/")
        domain = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else "root"
        
        print(f"Domain: {domain}, Site: {site_name}")
        
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{domain}:/sites/{site_name}"
        print(f"Request: {graph_url}")
        
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code != 200:
            print(f"❌ Error: Failed to access SharePoint site: HTTP {response.status_code}")
            print(f"Response: {response.text}")
            
            if response.status_code == 404:
                print("   Specified site not found")
                print("   Please check that your site URL is correct")
            elif response.status_code == 401:
                print("   No access permission")
                print("   Please check that your application has been granted appropriate permissions")
                print("   Permissions such as Sites.Read.All, Files.Read.All are required")
            
            return False
            
        site_info = response.json()
        print("✅ Successfully accessed SharePoint site")
        print(f"Site name: {site_info.get('displayName', 'Unknown')}")
        print(f"Site ID: {site_info.get('id', 'Unknown')}")
        
        # Try listing document libraries
        print("\nAttempting to list document libraries...")
        
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['id']}/drives"
        response = requests.get(drives_url, headers=headers)
        
        if response.status_code != 200:
            print(f"❌ Error: Failed to list document libraries: HTTP {response.status_code}")
            print(f"Response: {response.text}")
        else:
            drives = response.json().get("value", [])
            print(f"✅ Successfully listed {len(drives)} document libraries")
            for drive in drives:
                print(f"  - {drive.get('name', 'Unknown')}")
        
    except Exception as e:
        print(f"❌ Error: Exception during diagnostic: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    print("\n✅ Authentication diagnostic completed successfully")
    return True

if __name__ == "__main__":
    try:
        result = run_auth_diagnostic()
        if not result:
            print("\n❌ Authentication diagnostic failed")
            sys.exit(1)
    except Exception as e:
        print(f"❌ Error: An error occurred during diagnostic execution: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)