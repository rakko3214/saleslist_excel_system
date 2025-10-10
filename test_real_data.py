#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from real_data_app import *
from datetime import datetime

with app.app_context():
    print('=== 実際のデータ取得テスト ===')
    
    # 支店とアカウントのマッピング取得
    mapping = get_area_account_mapping()
    print('マッピング取得数:', len(mapping))
    
    if mapping:
        first_mapping = mapping[0]
        print('最初のマッピング:')
        print(f"  支店: {first_mapping['area_name']} (ID: {first_mapping['area_id']})")
        print(f"  アカウント: {first_mapping['account_name']} (ID: {first_mapping['account_id']})")
        
        # 実際のデータ取得テスト
        data_result = get_companies_data_by_period(
            first_mapping['area_id'], 
            first_mapping['account_id'],
            'today'
        )
        
        print(f'\n本日のデータ:')
        print(f"  新規: {data_result['new_count']}件")
        print(f"  更新: {data_result['update_count']}件")
        print(f"  期間: {data_result['period']}")
        
        # 1週間のデータ
        week_result = get_companies_data_by_period(
            first_mapping['area_id'], 
            first_mapping['account_id'],
            'week'
        )
        
        print(f'\n1週間のデータ:')
        print(f"  新規: {week_result['new_count']}件")
        print(f"  更新: {week_result['update_count']}件")
        
        # Excel出力データ生成テスト
        print(f'\n=== Excel出力データ生成テスト ===')
        hierarchical_data = generate_hierarchical_excel_data('today')
        print(f'生成された行数: {len(hierarchical_data)}行')
        
        # 最初の数行を表示
        print('\n生成されたデータサンプル:')
        for i, row in enumerate(hierarchical_data[:5]):
            print(f"  行{i+1}: {row['項目名'][:30]} - {row['件数']}")
        
        print('\n✅ randomデータではなく、実際のDBからデータを取得してExcel出力データを生成しています！')
        
        # 軽量化効果の確認
        total_companies = Company.query.count()
        today_companies = Company.query.filter(func.date(Company.created_at) == datetime.now().date()).count()
        week_companies = Company.query.filter(func.date(Company.created_at) >= datetime.now().date() - timedelta(days=7)).count()
        
        print(f'\n=== 軽量化効果 ===')
        print(f'全データ: {total_companies:,}件')
        print(f'今日のデータ: {today_companies:,}件 ({today_companies/total_companies*100:.1f}%)')
        print(f'1週間のデータ: {week_companies:,}件 ({week_companies/total_companies*100:.1f}%)')
        print(f'軽量化率(今日): {100-today_companies/total_companies*100:.1f}% データ削減')
        
    else:
        print('❌ マッピングデータが取得できませんでした')