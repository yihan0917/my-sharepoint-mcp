"""Microsoft Graph API client for SharePoint MCP server."""

import requests
import logging
from typing import Dict, Any, Optional

from auth.sharepoint_auth import SharePointContext

# Set up logging
logger = logging.getLogger("graph_client")

class GraphClient:
    """Client for interacting with Microsoft Graph API."""
    
    def __init__(self, context: SharePointContext):
        """Initialize Graph client with SharePoint context.
        
        Args:
            context: SharePoint authentication context
        """
        self.context = context
        self.base_url = context.graph_url
        logger.debug(f"GraphClient initialized with base URL: {self.base_url}")
    
    async def get(self, endpoint: str) -> Dict[str, Any]:
        """Send GET request to Graph API.
        
        Args:
            endpoint: API endpoint path (without base URL)
            
        Returns:
            Response from the API as dictionary
            
        Raises:
            Exception: If the request fails
        """
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        logger.debug(f"Making GET request to: {url}")
        
        # ヘッダーはコンテキストから取得（認証トークンを含む）
        headers = self.context.headers
        
        # リクエストを送信
        response = requests.get(url, headers=headers)
        
        # レスポンスをログに記録
        logger.debug(f"Response status code: {response.status_code}")
        
        if response.status_code != 200:
            error_text = response.text
            logger.error(f"Graph API error: {response.status_code} - {error_text}")
            
            # 認証関連のエラーを詳細に記録
            if response.status_code in (401, 403):
                logger.error("Authentication or authorization error detected")
                if "scp or roles claim" in error_text:
                    logger.error("Token does not have required claims (scp or roles)")
                    logger.error("Please check application permissions in Azure AD")
            
            raise Exception(f"Graph API error: {response.status_code} - {error_text}")
        
        # 正常なレスポンスをJSON形式で返す
        return response.json()
    
    async def post(self, endpoint: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """Send POST request to Graph API.
        
        Args:
            endpoint: API endpoint path (without base URL)
            data: JSON data to send
            
        Returns:
            Response from the API as dictionary
            
        Raises:
            Exception: If the request fails
        """
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        logger.debug(f"Making POST request to: {url}")
        logger.debug(f"With data: {data}")
        
        # ヘッダーはコンテキストから取得（認証トークンを含む）
        headers = self.context.headers
        
        # リクエストを送信
        response = requests.post(url, headers=headers, json=data)
        
        # レスポンスをログに記録
        logger.debug(f"Response status code: {response.status_code}")
        
        if response.status_code not in (200, 201):
            error_text = response.text
            logger.error(f"Graph API error: {response.status_code} - {error_text}")
            
            # 認証関連のエラーを詳細に記録
            if response.status_code in (401, 403):
                logger.error("Authentication or authorization error detected")
                if "scp or roles claim" in error_text:
                    logger.error("Token does not have required claims (scp or roles)")
                    logger.error("Please check application permissions in Azure AD")
            
            raise Exception(f"Graph API error: {response.status_code} - {error_text}")
        
        # 正常なレスポンスをJSON形式で返す
        return response.json()
        
    async def get_site_info(self, domain: str, site_name: str) -> Dict[str, Any]:
        """Get SharePoint site information.
        
        Args:
            domain: SharePoint domain
            site_name: Name of the site
            
        Returns:
            Site information
        """
        endpoint = f"sites/{domain}:/sites/{site_name}"
        logger.info(f"Getting site info for domain: {domain}, site: {site_name}")
        return await self.get(endpoint)
    
    async def list_document_libraries(self, domain: str, site_name: str) -> Dict[str, Any]:
        """List all document libraries in the site.
        
        Args:
            domain: SharePoint domain
            site_name: Name of the site
            
        Returns:
            List of document libraries
        """
        endpoint = f"sites/{domain}:/sites/{site_name}:/drives"
        logger.info(f"Listing document libraries for domain: {domain}, site: {site_name}")
        return await self.get(endpoint)