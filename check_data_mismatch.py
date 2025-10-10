from real_data_app import *
from datetime import datetime

with app.app_context():
    today = datetime.now().date()
    print(f'=== {today} のデータ検証 ===')
    
    # データベース直接クエリ（画像と同じクエリ）
    with db.engine.connect() as connection:
        result = connection.execute(text("SELECT COUNT(*) FROM companies WHERE DATE(updated_at) = '2025-10-10'"))
        db_count = result.fetchone()[0]
        print(f'DB直接クエリ (updated_at): {db_count}件')
        
        result2 = connection.execute(text("SELECT COUNT(*) FROM companies WHERE DATE(created_at) = '2025-10-10'"))
        created_count = result2.fetchone()[0]
        print(f'DB直接クエリ (created_at): {created_count}件')
    
    # アプリの集計ロジック確認
    print(f'\n=== アプリの集計ロジック確認 ===')
    
    # 新規データ（created_at基準）
    new_total = db.session.query(Company).filter(
        func.date(Company.created_at) == today
    ).count()
    print(f'新規データ (created_at = 今日): {new_total}件')
    
    # 更新データ（updated_at基準で作成日と更新日が異なる）
    update_total = db.session.query(Company).filter(
        func.date(Company.updated_at) == today,
        func.date(Company.created_at) != func.date(Company.updated_at)
    ).count()
    print(f'更新データ (updated_at = 今日 AND created_at != updated_at): {update_total}件')
    
    # 合計
    app_total = new_total + update_total
    print(f'アプリ表示合計: {app_total}件')
    
    print(f'\n=== 差異の原因調査 ===')
    print(f'DB直接クエリ: {db_count}件')
    print(f'アプリ集計: {app_total}件') 
    print(f'差異: {db_count - app_total}件')
    
    # updated_atが今日で、created_atも今日のデータ（新規作成かつ同日更新）
    both_today = db.session.query(Company).filter(
        func.date(Company.updated_at) == today,
        func.date(Company.created_at) == today
    ).count()
    print(f'\n作成も更新も今日: {both_today}件')
    print(f'これがnew_totalと重複している可能性があります')
    
    print(f'\n=== 正しい集計方法の提案 ===')
    print('DB側: updated_atが今日のデータ（すべて）')
    print('アプリ側: created_atが今日 + (updated_atが今日 かつ created_at≠updated_at)')
    print('→ アプリ側では同日作成&更新のデータが二重カウントを避けているため数が少ない')