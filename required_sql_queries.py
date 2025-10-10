# 実際のハローワークデータを取得するためのSQL文例

# 1. 日別の新規・更新件数を取得するクエリ
"""
SELECT 
    c.fm_area_id,
    c.fm_account_id,
    DATE(c.created_at) as date,
    COUNT(CASE WHEN c.status = 'new' THEN 1 END) as new_count,
    COUNT(CASE WHEN c.status = 'updated' THEN 1 END) as update_count,
    COUNT(*) as total_count
FROM companies c
WHERE c.created_at >= '2025-10-01'  -- 対象期間の開始日
  AND c.created_at <= '2025-10-08'  -- 対象期間の終了日
  AND c.is_hellowork = 1            -- ハローワーク対象のみ
GROUP BY c.fm_area_id, c.fm_account_id, DATE(c.created_at)
ORDER BY c.fm_area_id, c.fm_account_id, DATE(c.created_at);
"""

# 2. 期間合計の件数を取得するクエリ
"""
SELECT 
    c.fm_area_id,
    c.fm_account_id,
    COUNT(CASE WHEN c.status = 'new' THEN 1 END) as new_count,
    COUNT(CASE WHEN c.status = 'updated' THEN 1 END) as update_count,
    COUNT(*) as total_count
FROM companies c
WHERE c.created_at >= '2025-10-01'
  AND c.created_at <= '2025-10-08'
  AND c.is_hellowork = 1
GROUP BY c.fm_area_id, c.fm_account_id
ORDER BY c.fm_area_id, c.fm_account_id;
"""

# 3. 支店・アカウント別の月次集計クエリ
"""
SELECT 
    fa.area_name_ja,
    fac.department_name,
    c.fm_area_id,
    c.fm_account_id,
    YEAR(c.created_at) as year,
    MONTH(c.created_at) as month,
    COUNT(CASE WHEN c.status = 'new' THEN 1 END) as new_count,
    COUNT(CASE WHEN c.status = 'updated' THEN 1 END) as update_count,
    COUNT(*) as total_count
FROM companies c
INNER JOIN fm_areas fa ON c.fm_area_id = fa.id
INNER JOIN fm_accounts fac ON c.fm_account_id = fac.id
WHERE c.created_at >= '2025-10-01'
  AND c.created_at <= '2025-10-31'
  AND c.is_hellowork = 1
  AND fac.needs_hellowork = 1
GROUP BY c.fm_area_id, c.fm_account_id, YEAR(c.created_at), MONTH(c.created_at)
ORDER BY fa.id, fac.sort_order, year, month;
"""

# 4. SQLAlchemy版の実装例
def get_hellowork_data_by_period(start_date, end_date):
    """指定期間のハローワークデータを取得"""
    
    query = db.session.query(
        Company.fm_area_id,
        Company.fm_account_id,
        FmArea.area_name_ja,
        FmAccount.department_name,
        func.count(
            case([(Company.status == 'new', 1)], else_=None)
        ).label('new_count'),
        func.count(
            case([(Company.status == 'updated', 1)], else_=None)
        ).label('update_count'),
        func.count(Company.id).label('total_count')
    ).join(
        FmArea, Company.fm_area_id == FmArea.id
    ).join(
        FmAccount, Company.fm_account_id == FmAccount.id
    ).filter(
        Company.created_at >= start_date,
        Company.created_at <= end_date,
        Company.is_hellowork == 1,  # ハローワーク対象
        FmAccount.needs_hellowork == 1
    ).group_by(
        Company.fm_area_id,
        Company.fm_account_id,
        FmArea.area_name_ja,
        FmAccount.department_name
    ).order_by(
        FmArea.id,
        FmAccount.sort_order
    ).all()
    
    return query

# 5. 日別詳細データ取得のSQLAlchemy版
def get_daily_hellowork_data(start_date, end_date):
    """日別のハローワークデータを取得"""
    
    query = db.session.query(
        Company.fm_area_id,
        Company.fm_account_id,
        FmArea.area_name_ja,
        FmAccount.department_name,
        func.date(Company.created_at).label('date'),
        func.count(
            case([(Company.status == 'new', 1)], else_=None)
        ).label('new_count'),
        func.count(
            case([(Company.status == 'updated', 1)], else_=None)
        ).label('update_count')
    ).join(
        FmArea, Company.fm_area_id == FmArea.id
    ).join(
        FmAccount, Company.fm_account_id == FmAccount.id
    ).filter(
        Company.created_at >= start_date,
        Company.created_at <= end_date,
        Company.is_hellowork == 1,
        FmAccount.needs_hellowork == 1
    ).group_by(
        Company.fm_area_id,
        Company.fm_account_id,
        func.date(Company.created_at),
        FmArea.area_name_ja,
        FmAccount.department_name
    ).order_by(
        FmArea.id,
        FmAccount.sort_order,
        func.date(Company.created_at)
    ).all()
    
    return query