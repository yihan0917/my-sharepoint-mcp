"""SharePoint site information tools."""

import json
import logging
from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from utils.graph_client import GraphClient
from config.settings import SHAREPOINT_CONFIG

# Set up logging
logger = logging.getLogger("sharepoint_tools")

def register_site_tools(mcp: FastMCP):
    """Register SharePoint site tools with the MCP server."""
    
    @mcp.tool()
    async def get_site_info(ctx: Context) -> str:
        """Get basic information about the SharePoint site."""
        logger.info("Tool called: get_site_info")
        
        try:
            # コンテキストオブジェクトから認証情報を取得
            sp_ctx = ctx.request_context.lifespan_context
            
            # 必要に応じてトークンを更新
            await refresh_token_if_needed(sp_ctx)
            
            # GraphClientを作成
            graph_client = GraphClient(sp_ctx)
            
            # サイトURLからドメインとサイト名を抽出
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            logger.info(f"Getting info for site: {site_name} in domain: {domain}")
            
            # GraphClientを使用してサイト情報を取得
            site_info = await graph_client.get_site_info(domain, site_name)
            
            # レスポンスを整形
            result = {
                "name": site_info.get("displayName", "Unknown"),
                "description": site_info.get("description", "No description"),
                "created": site_info.get("createdDateTime", "Unknown"),
                "last_modified": site_info.get("lastModifiedDateTime", "Unknown"),
                "web_url": site_info.get("webUrl", SHAREPOINT_CONFIG["site_url"])
            }
            
            logger.info(f"Successfully retrieved site info for: {result['name']}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in get_site_info: {str(e)}")
            return f"Error accessing SharePoint: {str(e)}"
            
    @mcp.tool()
    async def list_document_libraries(ctx: Context) -> str:
        """List all document libraries in the SharePoint site."""
        logger.info("Tool called: list_document_libraries")
        
        try:
            # コンテキストオブジェクトから認証情報を取得
            sp_ctx = ctx.request_context.lifespan_context
            
            # 必要に応じてトークンを更新
            await refresh_token_if_needed(sp_ctx)
            
            # GraphClientを作成
            graph_client = GraphClient(sp_ctx)
            
            # サイトURLからドメインとサイト名を抽出
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            logger.info(f"Listing document libraries for site: {site_name} in domain: {domain}")
            
            # GraphClientを使用してドキュメントライブラリを一覧表示
            result = await graph_client.list_document_libraries(domain, site_name)
            
            # レスポンスからドライブ情報を抽出
            drives = result.get("value", [])
            formatted_drives = [{
                    "name": drive.get("name", "Unknown"),
                    "description": drive.get("description", "No description"),
                    "web_url": drive.get("webUrl", "Unknown"),
                    "drive_type": drive.get("driveType", "Unknown"),
                    "id": drive.get("id", "Unknown")
                } for drive in drives]
            
            logger.info(f"Successfully retrieved {len(formatted_drives)} document libraries")
            return json.dumps(formatted_drives, indent=2)
            
        except Exception as e:
            logger.error(f"Error in list_document_libraries: {str(e)}")
            return f"Error accessing SharePoint document libraries: {str(e)}"
            
    @mcp.tool()
    async def search_sharepoint(ctx: Context, query: str) -> str:
        """Search content in the SharePoint site.
        
        Args:
            query: Search query string
        """
        logger.info(f"Tool called: search_sharepoint with query: {query}")
        
        try:
            # コンテキストオブジェクトから認証情報を取得
            sp_ctx = ctx.request_context.lifespan_context
            
            # 必要に応じてトークンを更新
            await refresh_token_if_needed(sp_ctx)
            
            # GraphClientを作成
            graph_client = GraphClient(sp_ctx)
            
            # サイトURLからドメインとサイト名を抽出
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            logger.info(f"Searching for '{query}' in site: {site_name}")
            
            # まずサイト情報を取得してサイトIDを取得
            site_info = await graph_client.get_site_info(domain, site_name)
            site_id = site_info.get("id")
            
            if not site_id:
                logger.error("Failed to get site ID")
                return "Error: Could not retrieve site ID"
            
            # 検索リクエストを実行
            search_url = f"sites/{site_id}/search"
            search_data = {
                "requests": [
                    {
                        "entityTypes": [
                            "driveItem",
                            "listItem",
                            "list"
                        ],
                        "query": {
                            "queryString": query
                        }
                    }
                ]
            }
            
            logger.debug(f"Search request: {search_data}")
            search_results = await graph_client.post(search_url, search_data)
            
            # 検索結果を整形
            formatted_results = []
            for result in search_results.get("value", [])[0].get("hitsContainers", []):
                for hit in result.get("hits", []):
                    formatted_results.append({
                        "title": hit.get("resource", {}).get("name", "Unknown"),
                        "url": hit.get("resource", {}).get("webUrl", "Unknown"),
                        "type": hit.get("resource", {}).get("@odata.type", "Unknown"),
                        "summary": hit.get("summary", "No summary available")
                    })
            
            logger.info(f"Search returned {len(formatted_results)} results")
            return json.dumps(formatted_results, indent=2)
            
        except Exception as e:
            logger.error(f"Error in search_sharepoint: {str(e)}")
            return f"Error searching SharePoint: {str(e)}"