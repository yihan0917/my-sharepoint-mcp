#!/usr/bin/env python
"""アクセストークンの内容を確認するスクリプト"""

import os
import sys
import json
import base64
import msal
from dotenv import load_dotenv

def decode_jwt(token):
    """JWTトークンをデコードして内容を表示する"""
    try:
        # トークンの部分を取得（ヘッダー.ペイロード.署名）
        parts = token.split('.')
        if len(parts) != 3:
            print("❌ 無効なJWTトークン形式です")
            return None
            
        # ペイロード部分をデコード（2番目の部分）
        # パディングを追加
        payload = parts[1]
        payload += '=' * ((4 - len(payload) % 4) % 4)
        
        # Base64でデコード
        decoded = base64.b64decode(payload)
        claims = json.loads(decoded)
        
        return claims
    except Exception as e:
        print(f"❌ トークンのデコード中にエラーが発生しました: {e}")
        return None

def get_and_analyze_token():
    """トークンを取得して分析する"""
    print("=== アクセストークン分析 ===")
    
    # 環境変数を読み込む
    load_dotenv()
    
    try:
        # MSALクライアントアプリケーションの作成
        tenant_id = os.getenv("TENANT_ID")
        client_id = os.getenv("CLIENT_ID")
        client_secret = os.getenv("CLIENT_SECRET")
        
        if not all([tenant_id, client_id, client_secret]):
            print("❌ 認証に必要な環境変数が設定されていません")
            return False
        
        # トークンキャッシュの設定
        cache = msal.SerializableTokenCache()
        
        # MSALクライアントアプリケーションの作成
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret,
            token_cache=cache
        )
        
        # トークンの取得
        print("トークンの取得中...")
        scope = ["https://graph.microsoft.com/.default"]
        result = app.acquire_token_for_client(scopes=scope)
        
        if "access_token" not in result:
            print(f"❌ トークンの取得に失敗しました: {result.get('error', '不明')}")
            return False
        
        token = result["access_token"]
        print("✅ アクセストークンの取得に成功しました")
        
        # トークンの分析
        print("\n--- トークンの詳細分析 ---")
        claims = decode_jwt(token)
        
        if not claims:
            return False
        
        # 重要な情報を表示
        print("\n重要なクレーム情報:")
        print(f"発行者 (iss): {claims.get('iss', '不明')}")
        print(f"対象者 (aud): {claims.get('aud', '不明')}")
        print(f"アプリID (appid): {claims.get('appid', '不明')}")
        
        # roles と scp の確認（これが問題の原因と関連）
        roles = claims.get('roles', [])
        scp = claims.get('scp', '')
        
        print("\n権限情報:")
        if roles:
            print("✅ rolesクレームが存在します:")
            for role in roles:
                print(f"  - {role}")
        else:
            print("❌ rolesクレームが存在しません")
        
        if scp:
            print("✅ scpクレームが存在します:")
            print(f"  {scp}")
        else:
            print("❌ scpクレームが存在しません")
            
        # エラーメッセージの原因に関連する検証
        if not roles and not scp:
            print("\n⚠️ 警告: トークンにrolesもscpも含まれていません")
            print("   このことが「Either scp or roles claim need to be present in the token」エラーの原因です")
            print("   Azure ADでアプリケーション権限を正しく設定し、管理者の同意を得てください")
        
        # 全てのクレームを表示（オプション）
        print("\n全てのクレーム:")
        print(json.dumps(claims, indent=2))
        
        return True
        
    except Exception as e:
        print(f"❌ エラー: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    get_and_analyze_token()
