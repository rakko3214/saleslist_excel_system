#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
更新データロジックのテスト
fm_import_resultフィールドによる分類が正しく動作するかテスト
"""

import os
import sys
from datetime import datetime, timedelta
import pymysql

# データベース接続設定
DB_CONFIG = {
    'host': '192.168.0.133',
    'port': 3307,
    'user': 'saleslist_user',
    'password': 'saleslist_pass',
    'database': 'sales_list',
    'charset': 'utf8mb4'
}

def test_updated_logic():
    """修正されたロジックをテストする"""
    try:
        # データベース接続
        connection = pymysql.connect(**DB_CONFIG)
        cursor = connection.cursor()
        
        # テスト条件
        area_id = 1  # 札幌支店
        account_id = 1  # 管理部
        test_date = '2025-10-10'  # テスト日付
        
        print(f"=== 更新データロジックテスト ===")
        print(f"対象: 支店ID={area_id}, アカウントID={account_id}, 日付={test_date}")
        print()
        
        # 新規データ: fm_import_result = 2
        new_query = """
            SELECT COUNT(*) 
            FROM companies 
            WHERE fm_area_id = %s 
            AND imported_fm_account_id = %s 
            AND fm_import_result = 2 
            AND DATE(created_at) = %s
        """
        cursor.execute(new_query, (area_id, account_id, test_date))
        new_count = cursor.fetchone()[0]
        print(f"新規データ (fm_import_result=2): {new_count}件")
        
        # 更新データ: fm_import_result = 1
        update_query = """
            SELECT COUNT(*) 
            FROM companies 
            WHERE fm_area_id = %s 
            AND imported_fm_account_id = %s 
            AND fm_import_result = 1 
            AND DATE(updated_at) = %s
        """
        cursor.execute(update_query, (area_id, account_id, test_date))
        update_count = cursor.fetchone()[0]
        print(f"更新データ (fm_import_result=1): {update_count}件")
        print()
        
        # 参考: 旧ロジックでの結果も表示
        print("=== 参考: 旧ロジックでの結果 ===")
        old_new_query = """
            SELECT COUNT(*) 
            FROM companies 
            WHERE fm_area_id = %s 
            AND imported_fm_account_id = %s 
            AND DATE(created_at) = %s
        """
        cursor.execute(old_new_query, (area_id, account_id, test_date))
        old_new_count = cursor.fetchone()[0]
        print(f"旧ロジック新規データ (created_at基準): {old_new_count}件")
        
        old_update_query = """
            SELECT COUNT(*) 
            FROM companies 
            WHERE fm_area_id = %s 
            AND imported_fm_account_id = %s 
            AND DATE(updated_at) = %s 
            AND created_at != updated_at
        """
        cursor.execute(old_update_query, (area_id, account_id, test_date))
        old_update_count = cursor.fetchone()[0]
        print(f"旧ロジック更新データ (created_at≠updated_at): {old_update_count}件")
        print()
        
        # 結果比較
        print("=== 結果比較 ===")
        print(f"新規データ: 旧ロジック={old_new_count}件 → 新ロジック={new_count}件")
        print(f"更新データ: 旧ロジック={old_update_count}件 → 新ロジック={update_count}件")
        
        if update_count > 0:
            print("✅ 修正成功: 更新データが正しく取得できています")
        else:
            print("⚠️  更新データが0件です。日付や条件を確認してください")
        
        connection.close()
        
    except Exception as e:
        print(f"エラー: {e}")

if __name__ == "__main__":
    test_updated_logic()