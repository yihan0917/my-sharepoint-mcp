"""Main implementation of the SharePoint MCP Server."""

from contextlib import asynccontextmanager
from collections.abc import AsyncIterator

from mcp.server.fastmcp import FastMCP

from auth.sharepoint_auth import SharePointContext, get_auth_context
from config.settings import APP_NAME

# Import tool registrations
from tools.site_tools import register_site_tools

@asynccontextmanager
async def sharepoint_lifespan(server: FastMCP) -> AsyncIterator[SharePointContext]:
    """Manage the lifecycle of the SharePoint connection."""
    print("Initializing SharePoint connection...")
    
    try:
        # Get SharePoint authentication context
        context = await get_auth_context()
        yield context
    except Exception as e:
        print(f"Error during SharePoint authentication: {e}")
        # Return a dummy context to avoid errors
        dummy_context = SharePointContext(
            access_token="error",
            token_expiry=None,  # ここでNoneを使用
            graph_url="https://graph.microsoft.com/v1.0"
        )
        yield dummy_context
    finally:
        print("Closing SharePoint connection...")

# Create the MCP server
mcp = FastMCP(APP_NAME, lifespan=sharepoint_lifespan)

# Register tools
register_site_tools(mcp)

# Main execution
if __name__ == "__main__":
    mcp.run()