"""SharePoint site information resources."""

import json
import requests
from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from utils.graph_client import GraphClient
from config.settings import SHAREPOINT_CONFIG

def register_site_resources(mcp: FastMCP):
    """Register SharePoint site resources with the MCP server."""
    
    @mcp.resource("sharepoint://site-info")
    async def get_site_info(ctx: Context) -> str:
        """Get basic information about the SharePoint site."""
        await refresh_token_if_needed(ctx.request_context.lifespan_context)
        sp_ctx = ctx.request_context.lifespan_context
        
        # Create Graph client
        graph_client = GraphClient(sp_ctx)
        
        try:
            # Extract site domain and name from site URL
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            # Get site information via Microsoft Graph API
            site_url = f"{sp_ctx.graph_url}/sites/{domain}:/sites/{site_name}"
            response = requests.get(site_url, headers=sp_ctx.headers)
            
            if response.status_code != 200:
                return f"Error retrieving site info: {response.status_code} - {response.text}"
            
            site_info = response.json()
            
            # Format the output
            result = {
                "name": site_info.get("displayName", "Unknown"),
                "description": site_info.get("description", "No description"),
                "created": site_info.get("createdDateTime", "Unknown"),
                "last_modified": site_info.get("lastModifiedDateTime", "Unknown"),
                "web_url": site_info.get("webUrl", SHAREPOINT_CONFIG["site_url"])
            }
            
            return json.dumps(result, indent=2)
        except Exception as e:
            return f"Error accessing SharePoint: {str(e)}"