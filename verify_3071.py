from real_data_app import *
from datetime import datetime

with app.app_context():
    today = datetime.now().date()
    print('=== アプリ画面の3071件の検証 ===')
    
    # 支店・アカウントマッピングを取得
    mapping = get_area_account_mapping()
    print(f'支店・アカウント組み合わせ数: {len(mapping)}件')
    
    # 全支店・アカウントの合計を計算
    all_total_new = 0
    all_total_update = 0
    
    for item in mapping:
        data = get_companies_data_by_period(
            item['area_id'], 
            item['account_id'],
            'today'
        )
        all_total_new += data['new_count']
        all_total_update += data['update_count']
    
    print(f'\n全支店・アカウントの合計:')
    print(f'新規: {all_total_new}件')
    print(f'更新: {all_total_update}件')
    print(f'合計: {all_total_new + all_total_update}件')
    
    print(f'\n=== 比較結果 ===')
    print(f'DB全体 (updated_at = 今日): 5101件')
    print(f'アプリ集計 (支店・アカウント別): {all_total_new + all_total_update}件')
    print(f'アプリ画面表示: 3071件')
    
    # 差異の原因調査
    print(f'\n=== 差異の原因 ===')
    if (all_total_new + all_total_update) == 3071:
        print('✅ アプリ集計とアプリ画面表示が一致')
        print('→ 支店・アカウントでフィルタリングされたデータが正しく表示されている')
    else:
        print('❌ アプリ集計とアプリ画面表示が不一致')
        print(f'差異: {abs((all_total_new + all_total_update) - 3071)}件')
    
    print(f'\nDB全体とアプリ集計の差異: {5101 - (all_total_new + all_total_update)}件')
    print('→ この差異は、支店・アカウントに割り当てられていないデータ')