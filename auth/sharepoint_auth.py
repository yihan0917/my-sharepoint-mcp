"""SharePoint authentication handler module."""

from dataclasses import dataclass
from datetime import datetime, timedelta
import json
import os

import msal
from config.settings import SHAREPOINT_CONFIG, TOKEN_CACHE_FILE

@dataclass
class SharePointContext:
    """Context object for SharePoint connection."""
    access_token: str
    token_expiry: datetime
    graph_url: str = "https://graph.microsoft.com/v1.0"

    @property
    def headers(self) -> dict[str, str]:
        """Get authorization headers for API calls."""
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def is_token_valid(self) -> bool:
        """Check if the access token is still valid."""
        return datetime.now() < self.token_expiry


async def get_auth_context() -> SharePointContext:
    """Get SharePoint authentication context."""
    # Set up token cache
    cache = msal.SerializableTokenCache()
    
    # Load cached token if available
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, 'r') as cache_file:
            cache.deserialize(cache_file.read())
    
    # Create MSAL client application
    app = msal.ConfidentialClientApplication(
        SHAREPOINT_CONFIG["client_id"],
        authority=f"https://login.microsoftonline.com/{SHAREPOINT_CONFIG['tenant_id']}",
        client_credential=SHAREPOINT_CONFIG["client_secret"],
        token_cache=cache
    )
    
    # Get valid token from cache
    accounts = app.get_accounts()
    result = None
    if accounts:
        # Try silent authentication with existing account
        result = app.acquire_token_silent(SHAREPOINT_CONFIG["scope"], account=accounts[0])
    
    # If silent authentication fails, get a new token
    if not result:
        # Try client credentials flow
        result = app.acquire_token_for_client(scopes=SHAREPOINT_CONFIG["scope"])
        
        # If that fails, try username/password flow
        if "access_token" not in result:
            result = app.acquire_token_by_username_password(
                SHAREPOINT_CONFIG["username"], 
                SHAREPOINT_CONFIG["password"],
                scopes=SHAREPOINT_CONFIG["scope"]
            )
    
    # Raise error if token acquisition fails
    if "access_token" not in result:
        error_description = result.get("error_description", "Unknown error")
        raise Exception(f"Authentication failed: {error_description}")
    
    # Save token cache
    with open(TOKEN_CACHE_FILE, 'w') as cache_file:
        cache_file.write(cache.serialize())
    
    # Calculate token expiry (default is 1 hour)
    expiry = datetime.now() + timedelta(seconds=result.get("expires_in", 3600))
    
    # Return auth context
    return SharePointContext(
        access_token=result["access_token"],
        token_expiry=expiry
    )


async def refresh_token_if_needed(context: SharePointContext) -> None:
    """Refresh token if needed."""
    if not context.is_token_valid():
        print("Refreshing expired token...")
        # Re-authenticate to get a new token
        new_context = await get_auth_context()
        
        # Update the context
        context.access_token = new_context.access_token
        context.token_expiry = new_context.token_expiry