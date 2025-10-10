#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from dotenv import load_dotenv
import pymysql

load_dotenv()

# .envファイルからデータベース接続設定を取得
host = os.getenv('DB_HOST', '192.168.0.133')
port = int(os.getenv('DB_PORT', '3307'))
database = os.getenv('DB_DATABASE', 'sales_list')
user = os.getenv('DB_USERNAME', 'root')
password = os.getenv('DB_PASSWORD', 'root')

try:
    connection = pymysql.connect(
        host=host,
        port=port,
        user=user,
        password=password,
        database=database,
        charset='utf8mb4'
    )
    
    with connection.cursor() as cursor:
        # テーブル一覧を取得
        cursor.execute('SHOW TABLES')
        tables = cursor.fetchall()
        
        print('=== データベースのテーブル一覧 ===')
        for table in tables:
            print(f'- {table[0]}')
        
        # 売上やデータに関連しそうなテーブルを探す
        sales_related_tables = []
        for table in tables:
            table_name = table[0]
            if any(keyword in table_name.lower() for keyword in ['sales', 'data', 'count', 'log', 'record', 'transaction', 'job', 'work']):
                sales_related_tables.append(table_name)
        
        if sales_related_tables:
            print('\n=== 売上・データ関連テーブル ===')
            for table_name in sales_related_tables:
                print(f'- {table_name}')
                
        # 既知のテーブル構造を確認
        for table_name in ['companies', 'fm_areas', 'fm_accounts', 'fm_area_accounts']:
            try:
                cursor.execute(f'DESCRIBE {table_name}')
                columns = cursor.fetchall()
                print(f'\n=== {table_name} テーブル構造 ===')
                for col in columns:
                    print(f'  {col[0]}: {col[1]}')
                    
                # サンプルデータも確認
                cursor.execute(f'SELECT COUNT(*) FROM {table_name}')
                count = cursor.fetchone()[0]
                print(f'  データ件数: {count}件')
                
            except Exception as e:
                print(f'{table_name} テーブルが見つかりません: {e}')
        
        # companiesテーブルにcreated_atやupdated_atカラムがあるか確認
        try:
            cursor.execute('DESCRIBE companies')
            columns = cursor.fetchall()
            date_columns = [col[0] for col in columns if 'date' in col[0].lower() or 'time' in col[0].lower() or 'created' in col[0].lower() or 'updated' in col[0].lower()]
            if date_columns:
                print(f'\n=== companiesテーブルの日付関連カラム ===')
                for col in date_columns:
                    print(f'  - {col}')
        except Exception as e:
            print(f'companiesテーブルの日付カラム確認でエラー: {e}')
                
except Exception as e:
    print(f'データベース接続エラー: {e}')
finally:
    try:
        if 'connection' in locals():
            connection.close()
    except:
        pass