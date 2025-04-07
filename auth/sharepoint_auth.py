"""SharePoint authentication handler module."""

from dataclasses import dataclass
from datetime import datetime, timedelta
import json
import os
import logging

import msal
import requests
from config.settings import SHAREPOINT_CONFIG, TOKEN_CACHE_FILE

# Set up logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("sharepoint_auth")

@dataclass
class SharePointContext:
    """Context object for SharePoint connection."""
    access_token: str
    token_expiry: datetime
    graph_url: str = "https://graph.microsoft.com/v1.0"

    @property
    def headers(self) -> dict[str, str]:
        """Get authorization headers for API calls."""
        # ヘッダーの内容をログに出力（トークンは一部のみ表示）
        token_preview = f"{self.access_token[:10]}...{self.access_token[-10:]}" if self.access_token else "None"
        logger.debug(f"Using token (preview): {token_preview}")
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def is_token_valid(self) -> bool:
        """Check if the access token is still valid."""
        # Add safety check to handle None expiry
        if not self.token_expiry:
            return False
        is_valid = datetime.now() < self.token_expiry
        logger.debug(f"Token valid: {is_valid}, Expires: {self.token_expiry}")
        return is_valid

    def test_connection(self) -> bool:
        """Test the connection to SharePoint."""
        try:
            # Extract site domain and name from site URL
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            # Get site information via Microsoft Graph API
            site_url = f"{self.graph_url}/sites/{domain}:/sites/{site_name}"
            logger.debug(f"Testing connection to: {site_url}")
            
            response = requests.get(site_url, headers=self.headers)
            
            if response.status_code != 200:
                logger.error(f"Connection test failed: HTTP {response.status_code} - {response.text}")
                return False
                
            logger.info(f"Connection test successful: {response.status_code}")
            return True
        except Exception as e:
            logger.error(f"Error during connection test: {e}")
            return False


def validate_config() -> None:
    """Validate SharePoint configuration."""
    missing_vars = []
    for key in ["tenant_id", "client_id", "client_secret", "site_url"]:
        if not SHAREPOINT_CONFIG.get(key):
            missing_vars.append(key)
    
    if missing_vars:
        error_msg = f"Missing required configuration: {', '.join(missing_vars)}"
        logger.error(error_msg)
        raise ValueError(error_msg)
    
    # Validate site URL format
    site_url = SHAREPOINT_CONFIG["site_url"]
    if not site_url.startswith("https://") or ".sharepoint.com/" not in site_url.lower():
        logger.error(f"Invalid SharePoint site URL: {site_url}")
        raise ValueError(f"Invalid SharePoint site URL: {site_url}")


async def get_auth_context() -> SharePointContext:
    """Get SharePoint authentication context."""
    # Validate configuration first
    validate_config()
    
    # Set up token cache
    cache = msal.SerializableTokenCache()
    
    # 既存のキャッシュファイルは削除して強制的に新しいトークンを取得
    if os.path.exists(TOKEN_CACHE_FILE):
        try:
            os.remove(TOKEN_CACHE_FILE)
            logger.info("Removed existing token cache to force new token acquisition")
        except Exception as e:
            logger.warning(f"Error removing token cache: {e}")
    
    # Create MSAL client application
    app = msal.ConfidentialClientApplication(
        SHAREPOINT_CONFIG["client_id"],
        authority=f"https://login.microsoftonline.com/{SHAREPOINT_CONFIG['tenant_id']}",
        client_credential=SHAREPOINT_CONFIG["client_secret"],
        token_cache=cache
    )
    
    # Get new token
    logger.info("Acquiring new token with client credentials flow")
    result = app.acquire_token_for_client(scopes=SHAREPOINT_CONFIG["scope"])
    
    # Raise error if token acquisition fails
    if "access_token" not in result:
        error_code = result.get("error", "unknown")
        error_description = result.get("error_description", "Unknown error")
        logger.error(f"Authentication failed: {error_code} - {error_description}")
        
        # Log more detailed error information for troubleshooting
        if "AADSTS" in error_description:
            if "AADSTS50034" in error_description:
                logger.error("User account doesn't exist or is invalid")
            elif "AADSTS50126" in error_description:
                logger.error("Invalid username or password")
            elif "AADSTS65001" in error_description:
                logger.error("Application does not have the required permissions")
            elif "AADSTS70011" in error_description:
                logger.error("Application specified in the request is not found in the tenant")
        
        raise Exception(f"Authentication failed: {error_code} - {error_description}")
    
    # ログに表示するトークンの一部（セキュリティのため）
    token_preview = f"{result['access_token'][:10]}...{result['access_token'][-10:]}"
    logger.info(f"Token acquired successfully: {token_preview}")
    
    # テスト目的で、トークンに必要なクレームが含まれているか確認
    try:
        import base64
        import json
        
        # トークンのペイロード部分（2番目の部分）をデコード
        token_parts = result['access_token'].split('.')
        if len(token_parts) >= 2:
            # パディングを追加
            payload = token_parts[1]
            payload += '=' * ((4 - len(payload) % 4) % 4)
            
            # Base64でデコード
            decoded = base64.b64decode(payload)
            claims = json.loads(decoded)
            
            # ロールを確認
            roles = claims.get('roles', [])
            if roles:
                logger.info(f"Token contains roles: {roles}")
            else:
                logger.warning("Token does not contain roles claim")
                
            # スコープを確認
            scp = claims.get('scp', '')
            if scp:
                logger.info(f"Token contains scp claim: {scp}")
            else:
                logger.info("Token does not contain scp claim (expected for app-only tokens)")
                
            if not roles and not scp:
                logger.error("Token does not contain either roles or scp claim - this will cause errors!")
    except Exception as e:
        logger.warning(f"Error analyzing token claims: {e}")
    
    # Save token cache
    try:
        with open(TOKEN_CACHE_FILE, 'w') as cache_file:
            cache_file.write(cache.serialize())
        logger.info("Token cache saved to file")
    except Exception as e:
        logger.warning(f"Error saving token cache: {e}")
    
    # Calculate token expiry (default is 1 hour)
    expiry = datetime.now() + timedelta(seconds=result.get("expires_in", 3600))
    logger.info(f"Authentication successful, token expires at {expiry}")
    
    # Return auth context
    context = SharePointContext(
        access_token=result["access_token"],
        token_expiry=expiry
    )
    
    # Test connection immediately
    logger.info("Testing connection with acquired token...")
    if not context.test_connection():
        logger.warning("Connection test failed, but continuing anyway...")
    
    return context


async def refresh_token_if_needed(context: SharePointContext) -> None:
    """Refresh token if needed."""
    if not context.is_token_valid():
        logger.info("Token expired, refreshing...")
        try:
            # Re-authenticate to get a new token
            new_context = await get_auth_context()
            
            # Update the context
            context.access_token = new_context.access_token
            context.token_expiry = new_context.token_expiry
            logger.info("Token refreshed successfully")
        except Exception as e:
            logger.error(f"Error refreshing token: {e}")
            raise