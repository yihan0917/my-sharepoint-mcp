"""Main implementation of the SharePoint MCP Server."""

import os
import sys
import logging
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
from datetime import datetime, timedelta

from mcp.server.fastmcp import FastMCP

from auth.sharepoint_auth import SharePointContext, get_auth_context
from config.settings import APP_NAME

# デバッグモードを強制的に有効化
DEBUG = True

# ログレベルの設定
logging_level = logging.DEBUG if DEBUG else logging.INFO
logging.basicConfig(level=logging_level, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("sharepoint_mcp")

# ツール登録をインポート
from tools.site_tools import register_site_tools

@asynccontextmanager
async def sharepoint_lifespan(server: FastMCP) -> AsyncIterator[SharePointContext]:
    """SharePoint接続のライフサイクルを管理する。"""
    logger.info("SharePoint接続を初期化中...")
    
    try:
        # SharePoint認証コンテキストを取得
        logger.debug("認証コンテキストの取得を試行中...")
        context = await get_auth_context()
        logger.info(f"認証成功。トークンの有効期限: {context.token_expiry}")
        
        # コンテキストをアプリケーションで使用するためにyield
        yield context
        
    except Exception as e:
        logger.error(f"SharePoint認証中にエラーが発生しました: {e}")
        
        # エラーコンテキストを作成
        error_context = SharePointContext(
            access_token="error",
            token_expiry=datetime.now() + timedelta(seconds=10),  # 短い有効期限
            graph_url="https://graph.microsoft.com/v1.0"
        )
        
        logger.warning("認証失敗のためエラーコンテキストを使用します")
        yield error_context
        
    finally:
        logger.info("SharePoint接続を終了中...")

# MCPサーバーを作成
mcp = FastMCP(APP_NAME, lifespan=sharepoint_lifespan)

# ツールを登録
register_site_tools(mcp)

# メイン実行
if __name__ == "__main__":
    try:
        logger.info(f"{APP_NAME}サーバーを起動中...")
        mcp.run()
    except Exception as e:
        logger.error(f"MCPサーバーの起動中にエラーが発生しました: {e}")
        raise