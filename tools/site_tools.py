"""SharePoint site information tools."""

import json
import requests
from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from config.settings import SHAREPOINT_CONFIG

def register_site_tools(mcp: FastMCP):
    """Register SharePoint site tools with the MCP server."""
    
    @mcp.tool()
    async def get_site_info(ctx: Context) -> str:
        """Get basic information about the SharePoint site."""
        # コンテキストオブジェクトから認証情報を取得
        await refresh_token_if_needed(ctx.request_context.lifespan_context)
        sp_ctx = ctx.request_context.lifespan_context
        
        try:
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            site_url = f"{sp_ctx.graph_url}/sites/{domain}:/sites/{site_name}"
            response = requests.get(site_url, headers=sp_ctx.headers)
            
            if response.status_code != 200:
                return f"Error retrieving site info: {response.status_code} - {response.text}"
            
            site_info = response.json()
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
            
    @mcp.tool()
    async def list_document_libraries(ctx: Context) -> str:
        """List all document libraries in the SharePoint site."""
        await refresh_token_if_needed(ctx.request_context.lifespan_context)
        sp_ctx = ctx.request_context.lifespan_context
        
        try:
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            drives_url = f"{sp_ctx.graph_url}/sites/{domain}:/sites/{site_name}:/drives"
            response = requests.get(drives_url, headers=sp_ctx.headers)
            
            if response.status_code != 200:
                return f"Error retrieving document libraries: {response.status_code} - {response.text}"
            
            drives = response.json().get("value", [])
            formatted_drives = [{
                    "name": drive.get("name", "Unknown"),
                    "description": drive.get("description", "No description"),
                    "web_url": drive.get("webUrl", "Unknown"),
                    "drive_type": drive.get("driveType", "Unknown"),
                    "id": drive.get("id", "Unknown")
                } for drive in drives]
            
            return json.dumps(formatted_drives, indent=2)
        except Exception as e:
            return f"Error accessing SharePoint document libraries: {str(e)}"