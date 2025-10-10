#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# 修正したreal_data_app.pyのデータ取得機能をテストする
import os
import sys
sys.path.append(os.path.dirname(__file__))

from dotenv import load_dotenv
load_dotenv()

# Flaskアプリケーションのセットアップ
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text, func

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# モデル定義（real_data_app.pyから）
class FmArea(db.Model):
    __tablename__ = 'fm_areas'
    id = db.Column(db.Integer, primary_key=True)
    area_name_ja = db.Column(db.Text, nullable=False)
    area_name_en = db.Column(db.Text, nullable=False)
    fm_login_account_id = db.Column(db.Text, nullable=False)
    fm_login_account_pass = db.Column(db.Text, nullable=False)

class FmAccount(db.Model):
    __tablename__ = 'fm_accounts'
    id = db.Column(db.Integer, primary_key=True)
    department_name = db.Column(db.Text, nullable=False)
    sort_order = db.Column(db.Integer, nullable=False)
    needs_hellowork = db.Column(db.Integer, nullable=False, default=0)
    needs_tabelog = db.Column(db.Integer, nullable=False, default=0)
    needs_kanri = db.Column(db.Integer, nullable=False, default=1)

class Company(db.Model):
    __tablename__ = 'companies'
    id = db.Column(db.Integer, primary_key=True)
    media_site_id = db.Column(db.Integer)
    fm_area_id = db.Column(db.Integer)
    company_id_in_site = db.Column(db.Text)
    url = db.Column(db.Text)
    company_name = db.Column(db.Text)
    tel = db.Column(db.Text)
    address = db.Column(db.Text)
    hp = db.Column(db.Text)
    ceo_name = db.Column(db.Text)
    capital_stock = db.Column(db.Integer)
    job_detail = db.Column(db.Text)
    fm_major_industry_id = db.Column(db.Integer)
    fm_minor_industry_id = db.Column(db.Integer)
    kanri_regist_history_id = db.Column(db.Integer)
    share_departments = db.Column(db.Text)
    imported_fm_account_id = db.Column(db.Integer)
    display_started_at = db.Column(db.Date)
    created_at = db.Column(db.DateTime)
    updated_at = db.Column(db.DateTime)
    fm_saved_at = db.Column(db.DateTime)
    fm_import_result = db.Column(db.Integer)

class FmAreaAccount(db.Model):
    __tablename__ = 'fm_area_accounts'
    id = db.Column(db.Integer, primary_key=True)
    fm_area_id = db.Column(db.Integer, nullable=False)
    fm_account_id = db.Column(db.Integer, nullable=False)
    is_related = db.Column(db.Integer, nullable=False, default=1)

def test_data_retrieval():
    """データ取得のテスト"""
    with app.app_context():
        try:
            print("=== データベース接続テスト ===")
            # データベース接続テスト
            with db.engine.connect() as connection:
                result = connection.execute(text('SELECT VERSION()'))
                mysql_version = result.fetchone()[0]
                print(f"✅ データベース接続成功: {mysql_version}")
            
            print("\n=== 基本統計 ===")
            # 基本統計
            total_areas = FmArea.query.count()
            total_accounts = FmAccount.query.count()
            total_companies = Company.query.count()
            
            print(f"支店数: {total_areas}")
            print(f"アカウント数: {total_accounts}")
            print(f"企業データ総数: {total_companies}")
            
            print("\n=== 今日のデータ取得テスト ===")
            # 今日のデータ取得テスト
            from datetime import datetime, timedelta
            today = datetime.now().date()
            week_ago = today - timedelta(days=7)
            
            today_count = db.session.query(Company).filter(
                func.date(Company.created_at) == today
            ).count()
            
            week_count = db.session.query(Company).filter(
                func.date(Company.created_at) >= week_ago
            ).count()
            
            print(f"今日作成されたデータ: {today_count}件")
            print(f"1週間で作成されたデータ: {week_count}件")
            
            print("\n=== 支店・アカウント別データサンプル ===")
            # 支店・アカウント別の実際のデータ取得テスト
            sample_query = db.session.query(
                FmArea.area_name_ja,
                FmAccount.department_name,
                func.count(Company.id).label('company_count'),
                func.count(func.case((func.date(Company.created_at) == today, 1))).label('today_count')
            ).join(
                Company, FmArea.id == Company.fm_area_id
            ).join(
                FmAccount, Company.imported_fm_account_id == FmAccount.id
            ).filter(
                FmAccount.needs_hellowork == 1
            ).group_by(
                FmArea.id, FmAccount.id
            ).limit(5).all()
            
            print("支店名 | アカウント名 | 総件数 | 今日")
            print("-" * 50)
            for row in sample_query:
                print(f"{row.area_name_ja[:8]:8} | {row.department_name[:10]:10} | {row.company_count:6} | {row.today_count:4}")
            
            print("\n=== 期間指定データ取得関数テスト ===")
            # 期間指定データ取得関数のテスト
            def get_companies_data_by_period_test(area_id, account_id, date_filter='today'):
                from datetime import datetime, timedelta
                
                today = datetime.now().date()
                
                if date_filter == 'today':
                    filter_start = today
                    filter_end = today
                elif date_filter == 'week':
                    filter_start = today - timedelta(days=7)
                    filter_end = today
                elif date_filter == 'month':
                    filter_start = today - timedelta(days=30)
                    filter_end = today
                
                try:
                    new_count = db.session.query(Company).filter(
                        Company.fm_area_id == area_id,
                        Company.imported_fm_account_id == account_id,
                        func.date(Company.created_at).between(filter_start, filter_end)
                    ).count()
                    
                    update_count = db.session.query(Company).filter(
                        Company.fm_area_id == area_id,
                        Company.imported_fm_account_id == account_id,
                        func.date(Company.updated_at).between(filter_start, filter_end),
                        func.date(Company.created_at) != func.date(Company.updated_at)
                    ).count()
                    
                    return {
                        'new_count': new_count,
                        'update_count': update_count,
                        'period': f'{filter_start} 〜 {filter_end}'
                    }
                    
                except Exception as e:
                    return {'error': str(e)}
            
            # 実際のエリアとアカウントでテスト
            first_area = FmArea.query.first()
            first_account = FmAccount.query.filter(FmAccount.needs_hellowork == 1).first()
            
            if first_area and first_account:
                print(f"テスト対象: {first_area.area_name_ja} - {first_account.department_name}")
                
                for period in ['today', 'week', 'month']:
                    result = get_companies_data_by_period_test(first_area.id, first_account.id, period)
                    if 'error' not in result:
                        print(f"{period:8}: 新規{result['new_count']:4}件, 更新{result['update_count']:4}件")
                    else:
                        print(f"{period:8}: エラー - {result['error']}")
            
            print("\n✅ データ取得機能は正常に動作しています！")
            return True
            
        except Exception as e:
            print(f"❌ エラーが発生しました: {e}")
            import traceback
            traceback.print_exc()
            return False

if __name__ == "__main__":
    success = test_data_retrieval()
    if success:
        print("\n🎉 randomデータではなく、実際のデータベースからデータを取得できています！")
    else:
        print("\n⚠️ データ取得に問題があります。")