#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
実際のデータ存在確認とテスト
"""

from dotenv import load_dotenv
load_dotenv()
import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from datetime import date

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

class Company(db.Model):
    __tablename__ = 'companies'
    id = db.Column(db.Integer, primary_key=True)
    fm_area_id = db.Column(db.Integer)
    imported_fm_account_id = db.Column(db.Integer)
    fm_import_result = db.Column(db.Integer)
    created_at = db.Column(db.DateTime)
    updated_at = db.Column(db.DateTime)

def find_actual_data():
    """実際にデータが存在する条件を確認"""
    with app.app_context():
        print('=== 更新データ(fm_import_result=1)が存在する支店・アカウント ===')
        results = db.session.query(
            Company.fm_area_id, 
            Company.imported_fm_account_id, 
            func.count(Company.id).label('count')
        ).filter(
            Company.fm_import_result == 1,
            func.date(Company.updated_at) == date(2025, 10, 10)
        ).group_by(
            Company.fm_area_id, 
            Company.imported_fm_account_id
        ).order_by(
            func.count(Company.id).desc()
        ).limit(5).all()
        
        if results:
            for result in results:
                print(f'支店ID={result.fm_area_id}, アカウントID={result.imported_fm_account_id}: {result.count}件')
            
            # 最も件数が多い条件でテスト
            test_area_id = results[0].fm_area_id
            test_account_id = results[0].imported_fm_account_id
            
            print(f'\n=== テスト実行 (支店ID={test_area_id}, アカウントID={test_account_id}) ===')
            
            # 新ロジックでの件数取得
            new_count = db.session.query(Company).filter(
                Company.fm_area_id == test_area_id,
                Company.imported_fm_account_id == test_account_id,
                Company.fm_import_result == 2,
                func.date(Company.created_at) == date(2025, 10, 10)
            ).count()
            
            update_count = db.session.query(Company).filter(
                Company.fm_area_id == test_area_id,
                Company.imported_fm_account_id == test_account_id,
                Company.fm_import_result == 1,
                func.date(Company.updated_at) == date(2025, 10, 10)
            ).count()
            
            print(f'新規データ (fm_import_result=2): {new_count}件')
            print(f'更新データ (fm_import_result=1): {update_count}件')
            
            if update_count > 0:
                print('✅ 修正成功: 更新データが正しく取得できています')
            else:
                print('⚠️ 更新データが0件でした')
        else:
            print('該当データが見つかりませんでした')

if __name__ == "__main__":
    find_actual_data()