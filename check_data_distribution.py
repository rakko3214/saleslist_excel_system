#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from dotenv import load_dotenv
import pymysql
from datetime import datetime, timedelta

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
        # kanri_regist_historiesテーブルの構造確認
        try:
            cursor.execute('DESCRIBE kanri_regist_histories')
            columns = cursor.fetchall()
            print('=== kanri_regist_histories テーブル構造 ===')
            for col in columns:
                print(f'  {col[0]}: {col[1]}')
            
            cursor.execute('SELECT COUNT(*) FROM kanri_regist_histories')
            count = cursor.fetchone()[0]
            print(f'  データ件数: {count}件')
            
            # サンプルデータを確認
            cursor.execute('SELECT * FROM kanri_regist_histories LIMIT 5')
            samples = cursor.fetchall()
            print('\n=== サンプルデータ ===')
            for i, sample in enumerate(samples, 1):
                print(f'  データ{i}: {sample}')
                
        except Exception as e:
            print(f'kanri_regist_histories テーブル確認エラー: {e}')
        
        # companiesテーブルの日付分布を確認
        print('\n=== companies テーブルの日付分布 ===')
        try:
            # 最新と最古のデータ
            cursor.execute('SELECT MIN(created_at), MAX(created_at) FROM companies WHERE created_at IS NOT NULL')
            date_range = cursor.fetchone()
            print(f'  最古のデータ: {date_range[0]}')
            print(f'  最新のデータ: {date_range[1]}')
            
            # 今日、今週、今月のデータ件数
            today = datetime.now().date()
            week_ago = today - timedelta(days=7)
            month_ago = today - timedelta(days=30)
            
            cursor.execute('SELECT COUNT(*) FROM companies WHERE DATE(created_at) = %s', (today,))
            today_count = cursor.fetchone()[0]
            print(f'  今日のデータ: {today_count}件')
            
            cursor.execute('SELECT COUNT(*) FROM companies WHERE DATE(created_at) >= %s', (week_ago,))
            week_count = cursor.fetchone()[0]
            print(f'  1週間のデータ: {week_count}件')
            
            cursor.execute('SELECT COUNT(*) FROM companies WHERE DATE(created_at) >= %s', (month_ago,))
            month_count = cursor.fetchone()[0]
            print(f'  1ヶ月のデータ: {month_count}件')
            
        except Exception as e:
            print(f'日付分布確認エラー: {e}')
        
        # 支店・アカウント別のデータ分布
        print('\n=== 支店・アカウント別データ分布 ===')
        try:
            cursor.execute('''
                SELECT 
                    fa.area_name_ja,
                    fac.department_name,
                    COUNT(c.id) as company_count,
                    COUNT(CASE WHEN DATE(c.created_at) = %s THEN 1 END) as today_count,
                    COUNT(CASE WHEN DATE(c.created_at) >= %s THEN 1 END) as week_count
                FROM companies c
                LEFT JOIN fm_areas fa ON c.fm_area_id = fa.id
                LEFT JOIN fm_accounts fac ON c.imported_fm_account_id = fac.id
                WHERE fa.id IS NOT NULL AND fac.needs_hellowork = 1
                GROUP BY fa.id, fac.id
                ORDER BY fa.id, fac.sort_order
                LIMIT 10
            ''', (today, week_ago))
            
            results = cursor.fetchall()
            print('  支店名 | アカウント名 | 総件数 | 今日 | 1週間')
            print('  ' + '-' * 60)
            for row in results:
                print(f'  {row[0][:10]:10} | {row[1][:12]:12} | {row[2]:6} | {row[3]:4} | {row[4]:6}')
                
        except Exception as e:
            print(f'支店別データ分布確認エラー: {e}')
            
except Exception as e:
    print(f'データベース接続エラー: {e}')
finally:
    try:
        if 'connection' in locals():
            connection.close()
    except:
        pass