# 実データ対応のためのクエリ例

def get_real_hellowork_data(area_id, account_id, date_range=None):
    """実際のハローワークデータを取得する関数（例）"""
    
    if date_range is None:
        # デフォルトでは当月のデータを取得
        today = datetime.now().date()
        start_date = today.replace(day=1)  # 月初
        end_date = today  # 今日まで
    else:
        start_date, end_date = date_range
    
    # 実際のcompaniesテーブルからデータを集計
    # （この例は実際のテーブル構造に合わせて調整が必要）
    query = db.session.query(
        func.count(Company.id).label('count'),
        Company.status  # 'new' or 'updated' など
    ).filter(
        Company.fm_area_id == area_id,
        Company.fm_account_id == account_id,
        Company.created_at.between(start_date, end_date)
    ).group_by(Company.status)
    
    results = query.all()
    
    data = {'new': 0, 'updated': 0}
    for result in results:
        if result.status in data:
            data[result.status] = result.count
    
    return data

def generate_real_hierarchical_excel_data(date_range=None):
    """実データに基づく階層構造のExcel出力用データを生成"""
    
    mapping = get_area_account_mapping()
    hierarchical_data = []
    
    # 日付範囲の設定
    if date_range is None:
        today = datetime.now().date()
        start_date = today.replace(day=1)
        end_date = today
        date_str = f"{start_date.strftime('%Y年%m月%d日')} - {end_date.strftime('%Y年%m月%d日')}"
    else:
        start_date, end_date = date_range
        if start_date == end_date:
            date_str = start_date.strftime("%Y年%m月%d日")
        else:
            date_str = f"{start_date.strftime('%Y年%m月%d日')} - {end_date.strftime('%Y年%m月%d日')}"
    
    # 支店ごとにグループ化
    areas = {}
    for item in mapping:
        area_name = item['area_name']
        if area_name not in areas:
            areas[area_name] = {
                'area_id': item['area_id'],
                'accounts': []
            }
        areas[area_name]['accounts'].append({
            'account_id': item['account_id'],
            'account_name': item['account_name']
        })
    
    # 各支店・アカウントに対して実データを取得
    for area_name, area_data in areas.items():
        hierarchical_data.append({
            'レベル': 1,
            '項目名': f"📍 {area_name}",
            '種別': '',
            '件数': '',
            '備考': f'支店ID: {area_data["area_id"]}'
        })
        
        area_total = 0
        
        for account in area_data['accounts']:
            # 実データを取得
            real_data = get_real_hellowork_data(
                area_data['area_id'], 
                account['account_id'], 
                (start_date, end_date) if date_range else None
            )
            
            new_count = real_data['new']
            update_count = real_data['updated']
            
            hierarchical_data.append({
                'レベル': 2,
                '項目名': f"  📂 {account['account_name']}",
                '種別': '',
                '件数': '',
                '備考': f'アカウントID: {account["account_id"]}'
            })
            
            hierarchical_data.append({
                'レベル': 3,
                '項目名': f"    📝 新規",
                '種別': '新規',
                '件数': new_count,
                '備考': f'{date_str}の実データ'
            })
            
            hierarchical_data.append({
                'レベル': 3,
                '項目名': f"    🔄 更新",
                '種別': '更新',
                '件数': update_count,
                '備考': f'{date_str}の実データ'
            })
            
            account_total = new_count + update_count
            area_total += account_total
            
            hierarchical_data.append({
                'レベル': 2,
                '項目名': f"  └─ 小計",
                '種別': '小計',
                '件数': account_total,
                '備考': f'{account["account_name"]}の{date_str}合計'
            })
        
        hierarchical_data.append({
            'レベル': 1,
            '項目名': f"🔢 {area_name} 合計",
            '種別': '支店合計',
            '件数': area_total,
            '備考': f'{area_name}の{date_str}総計'
        })
        
        # 区切り行
        hierarchical_data.append({
            'レベル': 0,
            '項目名': '',
            '種別': '',
            '件数': '',
            '備考': ''
        })
    
    return hierarchical_data