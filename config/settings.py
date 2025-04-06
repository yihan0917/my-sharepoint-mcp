"""Configuration settings for the SharePoint MCP Server."""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Basic settings
APP_NAME = "SharePoint MCP"
DEBUG = os.getenv("DEBUG", "False").lower() in ("true", "1", "t")

# SharePoint connection settings
SHAREPOINT_CONFIG = {
    "tenant_id": os.getenv("TENANT_ID", ""),
    "client_id": os.getenv("CLIENT_ID", ""),
    "client_secret": os.getenv("CLIENT_SECRET", ""),
    "site_url": os.getenv("SITE_URL", ""),
    "username": os.getenv("USERNAME", ""),
    "password": os.getenv("PASSWORD", ""),
    "scope": ["https://graph.microsoft.com/.default"],
}

# Microsoft Graph API settings
GRAPH_API_VERSION = "v1.0"
GRAPH_BASE_URL = f"https://graph.microsoft.com/{GRAPH_API_VERSION}"

# Token settings
TOKEN_CACHE_FILE = ".token_cache"