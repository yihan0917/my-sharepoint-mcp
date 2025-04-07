#!/usr/bin/env python
"""SharePoint認証の診断スクリプト"""

import os
import sys
import json
import requests
from dotenv import load_dotenv

def run_auth_diagnostic():
    """SharePoint認証の診断を実行する"""
    print("=== SharePoint認証診断 ===")
    
    # .envファイルの確認
    if not os.path.exists(".env"):
        print("❌ エラー: .envファイルが見つかりません")
        print("   .env.exampleをコピーして設定してください")
        return False
    
    # 環境変数の読み込み
    load_dotenv()
    
    # 必須変数のチェック
    required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "SITE_URL"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"❌ エラー: 以下の環境変数が設定されていません: {', '.join(missing_vars)}")
        return False
    
    print("✅ すべての必須環境変数が設定されています")
    
    # SharePointのサイトURL形式をチェック
    site_url = os.getenv("SITE_URL")
    if not site_url.startswith("https://") or ".sharepoint.com/" not in site_url.lower():
        print(f"❌ エラー: 無効なSharePointサイトURL: {site_url}")
        print("   URLは以下の形式である必要があります: https://your-tenant.sharepoint.com/sites/your-site")
        return False
    
    print(f"✅ SharePointサイトURLの形式は有効です: {site_url}")
    
    # 認証テスト
    print("\n--- 認証とリクエストのテスト ---")
    
    try:
        import msal
        
        # トークンキャッシュの設定
        cache = msal.SerializableTokenCache()
        
        # MSALクライアントアプリケーションの作成
        tenant_id = os.getenv("TENANT_ID")
        client_id = os.getenv("CLIENT_ID")
        client_secret = os.getenv("CLIENT_SECRET")
        
        print(f"テナントID: {tenant_id[:5]}...{tenant_id[-5:]}")
        print(f"クライアントID: {client_id[:5]}...{client_id[-5:]}")
        
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret,
            token_cache=cache
        )
        
        # クライアント認証フローを試行
        print("クライアント認証フローを試行中...")
        scope = ["https://graph.microsoft.com/.default"]
        result = app.acquire_token_for_client(scopes=scope)
        
        if "access_token" not in result:
            error_code = result.get("error", "unknown")
            error_description = result.get("error_description", "Unknown error")
            print(f"❌ エラー: 認証に失敗しました: {error_code} - {error_description}")
            
            # 共通のエラーに対する詳細説明
            if "AADSTS700016" in str(error_description):
                print("   アプリケーションが見つからないか、テナントに対して無効になっています")
                print("   Azure ADでアプリケーション登録を確認してください")
            elif "AADSTS7000215" in str(error_description):
                print("   無効なクライアントシークレットです")
                print("   クライアントシークレットが正しいか、有効期限が切れていないか確認してください")
            elif "AADSTS650057" in str(error_description):
                print("   クライアント資格情報が無効です")
                print("   クライアントIDとクライアントシークレットを確認してください")
            elif "AADSTS70011" in str(error_description):
                print("   このアプリケーションはテナントに見つかりませんでした")
                print("   アプリケーションがこのテナントに登録されているか確認してください")
            
            print(f"\n完全なエラーレスポンス: \n{json.dumps(result, indent=2)}")
            return False
            
        print("✅ アクセストークンの取得に成功しました")
        
        # SharePointサイトへのアクセスを試行
        print("SharePointサイトへのアクセスを試行中...")
        
        headers = {
            "Authorization": f"Bearer {result['access_token']}",
            "Content-Type": "application/json",
        }
        
        # サイトURLからドメインとサイト名を抽出
        site_parts = site_url.replace("https://", "").split("/")
        domain = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else "root"
        
        print(f"ドメイン: {domain}, サイト: {site_name}")
        
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{domain}:/sites/{site_name}"
        print(f"リクエスト: {graph_url}")
        
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code != 200:
            print(f"❌ エラー: SharePointサイトへのアクセスに失敗しました: HTTP {response.status_code}")
            print(f"レスポンス: {response.text}")
            
            if response.status_code == 404:
                print("   指定されたサイトが見つかりません")
                print("   サイトURLが正しいか確認してください")
            elif response.status_code == 401:
                print("   アクセス権限がありません")
                print("   アプリケーションに適切な権限が付与されているか確認してください")
                print("   Sites.Read.All, Files.Read.All などの権限が必要です")
            
            return False
            
        site_info = response.json()
        print("✅ SharePointサイトへのアクセスに成功しました")
        print(f"サイト名: {site_info.get('displayName', '不明')}")
        print(f"サイトID: {site_info.get('id', '不明')}")
        
        # ドキュメントライブラリの一覧を試行
        print("\nドキュメントライブラリの一覧取得を試行中...")
        
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['id']}/drives"
        response = requests.get(drives_url, headers=headers)
        
        if response.status_code != 200:
            print(f"❌ エラー: ドキュメントライブラリの一覧取得に失敗しました: HTTP {response.status_code}")
            print(f"レスポンス: {response.text}")
        else:
            drives = response.json().get("value", [])
            print(f"✅ {len(drives)}個のドキュメントライブラリの一覧取得に成功しました")
            for drive in drives:
                print(f"  - {drive.get('name', '不明')}")
        
    except Exception as e:
        print(f"❌ エラー: 診断中に例外が発生しました: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    print("\n✅ 認証診断が正常に完了しました")
    return True

if __name__ == "__main__":
    try:
        result = run_auth_diagnostic()
        if not result:
            print("\n❌ 認証診断に失敗しました")
            sys.exit(1)
    except Exception as e:
        print(f"❌ エラー: 診断の実行中にエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)