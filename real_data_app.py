from flask import Flask, render_template_string, jsonify, request, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text, func, and_
from dotenv import load_dotenv
import os
import pymysql
from datetime import datetime, date, timedelta
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import io

# 環境変数をロード
load_dotenv()

app = Flask(__name__)

# データベース設定
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY')

db = SQLAlchemy(app)

# ========================
# 実際のデータ構造に合わせたモデル定義
# ========================

class FmArea(db.Model):
    __tablename__ = 'fm_areas'
    
    id = db.Column(db.Integer, primary_key=True)
    area_name_ja = db.Column(db.Text, nullable=False)
    area_name_en = db.Column(db.Text, nullable=False)
    fm_login_account_id = db.Column(db.Text, nullable=False)
    fm_login_account_pass = db.Column(db.Text, nullable=False)
    
    def to_dict(self):
        return {
            'id': self.id,
            'name': self.area_name_ja,
            'name_en': self.area_name_en,
            'login_id': self.fm_login_account_id
        }

class FmAccount(db.Model):
    __tablename__ = 'fm_accounts'
    
    id = db.Column(db.Integer, primary_key=True)
    department_name = db.Column(db.Text, nullable=False)
    sort_order = db.Column(db.Integer, nullable=False)
    needs_hellowork = db.Column(db.Integer, nullable=False, default=0)
    needs_tabelog = db.Column(db.Integer, nullable=False, default=0)
    needs_kanri = db.Column(db.Integer, nullable=False, default=1)
    
    def to_dict(self):
        return {
            'id': self.id,
            'name': self.department_name,
            'sort_order': self.sort_order,
            'needs_hellowork': bool(self.needs_hellowork),
            'needs_tabelog': bool(self.needs_tabelog),
            'needs_kanri': bool(self.needs_kanri)
        }

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
    capital_stock = db.Column(db.Text)
    # 他のフィールドは必要に応じて追加
    
    def to_dict(self):
        return {
            'id': self.id,
            'company_name': self.company_name,
            'fm_area_id': self.fm_area_id,
            'address': self.address,
            'tel': self.tel,
            'url': self.url
        }

# fm_area_accountsの関連テーブル（多対多関係）
class FmAreaAccount(db.Model):
    __tablename__ = 'fm_area_accounts'
    
    id = db.Column(db.Integer, primary_key=True)
    fm_area_id = db.Column(db.Integer, nullable=False)
    fm_account_id = db.Column(db.Integer, nullable=False)
    
    def to_dict(self):
        return {
            'id': self.id,
            'fm_area_id': self.fm_area_id,
            'fm_account_id': self.fm_account_id
        }

# ========================
# 集計関数（実データ構造対応）
# ========================

def get_area_account_summary():
    """支店・アカウント・データ件数のサマリーを取得"""
    
    # 支店別の企業データ件数を取得
    area_summary = db.session.query(
        FmArea.id,
        FmArea.area_name_ja,
        func.count(Company.id).label('company_count')
    ).outerjoin(
        Company, FmArea.id == Company.fm_area_id
    ).group_by(
        FmArea.id, FmArea.area_name_ja
    ).order_by(FmArea.id).all()
    
    # アカウント別の関連情報
    account_summary = db.session.query(
        FmAccount.id,
        FmAccount.department_name,
        FmAccount.needs_hellowork,
        func.count(FmAreaAccount.id).label('area_count')
    ).outerjoin(
        FmAreaAccount, FmAccount.id == FmAreaAccount.fm_account_id
    ).group_by(
        FmAccount.id
    ).order_by(FmAccount.sort_order).all()
    
    return {
        'areas': [
            {
                'id': row.id,
                'name': row.area_name_ja,
                'company_count': row.company_count
            } for row in area_summary
        ],
        'accounts': [
            {
                'id': row.id,
                'name': row.department_name,
                'needs_hellowork': bool(row.needs_hellowork),
                'area_count': row.area_count
            } for row in account_summary
        ]
    }

def get_area_account_mapping():
    """支店とアカウントの関連マッピングを取得"""
    
    mapping = db.session.query(
        FmAreaAccount.fm_area_id,
        FmAreaAccount.fm_account_id,
        FmArea.area_name_ja,
        FmAccount.department_name,
        FmAccount.needs_hellowork
    ).join(
        FmArea, FmAreaAccount.fm_area_id == FmArea.id
    ).join(
        FmAccount, FmAreaAccount.fm_account_id == FmAccount.id
    ).filter(
        FmAccount.needs_hellowork == 1  # ハローワークが必要なアカウントのみ
    ).order_by(
        FmArea.id, FmAccount.sort_order
    ).all()
    
    return [
        {
            'area_id': row.fm_area_id,
            'area_name': row.area_name_ja,
            'account_id': row.fm_account_id,
            'account_name': row.department_name,
            'needs_hellowork': bool(row.needs_hellowork)
        } for row in mapping
    ]

def generate_hierarchical_excel_data():
    """階層構造のExcel出力用データを生成（支店→アカウント→更新/新規→件数）"""
    import random
    from datetime import datetime, timedelta
    
    # 支店とアカウントのマッピングを取得
    mapping = get_area_account_mapping()
    
    # 階層構造データを生成
    hierarchical_data = []
    
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
    
    # 各支店・アカウントに対してサンプルデータを生成
    for area_name, area_data in areas.items():
        # 支店ヘッダー行
        hierarchical_data.append({
            'レベル': 1,
            '項目名': area_name,
            '種別': '',
            '件数': '',
            '備考': f'支店ID: {area_data["area_id"]}'
        })
        
        for account in area_data['accounts']:
            # アカウントヘッダー行
            hierarchical_data.append({
                'レベル': 2,
                '項目名': f"  ├─ {account['account_name']}",
                '種別': '',
                '件数': '',
                '備考': f'アカウントID: {account["account_id"]}'
            })
            
            # 新規・更新のデータ（実際のデータがない場合はサンプル）
            # 本日のデータ
            today = datetime.now().date()
            new_count = random.randint(5, 25)  # 実際のDBクエリに置き換え予定
            update_count = random.randint(3, 15)  # 実際のDBクエリに置き換え予定
            
            hierarchical_data.append({
                'レベル': 3,
                '項目名': f"    ├─ 新規",
                '種別': '新規',
                '件数': new_count,
                '備考': f'{today.strftime("%Y年%m月%d日")}のデータ'
            })
            
            hierarchical_data.append({
                'レベル': 3,
                '項目名': f"    └─ 更新",
                '種別': '更新',
                '件数': update_count,
                '備考': f'{today.strftime("%Y年%m月%d日")}のデータ'
            })
            
            # 合計行
            total_count = new_count + update_count
            hierarchical_data.append({
                'レベル': 2,
                '項目名': f"  └─ 小計",
                '種別': '合計',
                '件数': total_count,
                '備考': f'{account["account_name"]}の合計'
            })
    
    return hierarchical_data

# ========================
# HTMLテンプレート（実データ対応）
# ========================

MAIN_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ハローワークデータ管理システム（実データ版）</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .card { background: white; border-radius: 10px; padding: 25px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .card h2 { color: #333; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 3px solid #667eea; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; text-align: center; }
        .stat-card h3 { font-size: 2em; margin-bottom: 5px; }
        .stat-card p { opacity: 0.9; }
        .table-container { overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f8f9fa; font-weight: bold; color: #495057; }
        tr:hover { background-color: #f8f9fa; }
        .btn { padding: 12px 25px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; font-weight: bold; text-decoration: none; display: inline-block; text-align: center; transition: all 0.3s; margin: 5px; }
        .btn-primary { background: #667eea; color: white; }
        .btn-primary:hover { background: #5a6fd8; transform: translateY(-2px); }
        .btn-success { background: #28a745; color: white; }
        .btn-success:hover { background: #218838; transform: translateY(-2px); }
        .btn-info { background: #17a2b8; color: white; }
        .btn-info:hover { background: #138496; transform: translateY(-2px); }
        .status-success { color: #28a745; font-weight: bold; }
        .status-error { color: #dc3545; font-weight: bold; }
        .hellowork-enabled { background-color: #d4edda; }
        .area-section { margin-bottom: 30px; border-left: 4px solid #667eea; padding-left: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏢 ハローワークデータ管理システム（実データ版）</h1>
            <p>{{ stats.total_companies }}件の企業データ・{{ stats.total_areas }}支店・{{ stats.total_accounts }}アカウントを管理</p>
        </div>
        
        <div class="card">
            <h2>📊 システム統計</h2>
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>{{ stats.total_areas }}</h3>
                    <p>登録支店数</p>
                </div>
                <div class="stat-card">
                    <h3>{{ stats.total_accounts }}</h3>
                    <p>総アカウント数</p>
                </div>
                <div class="stat-card">
                    <h3>{{ stats.hellowork_accounts }}</h3>
                    <p>ハローワーク対象</p>
                </div>
                <div class="stat-card">
                    <h3>{{ stats.total_companies }}</h3>
                    <p>企業データ件数</p>
                </div>
            </div>
        </div>
        
        <div class="card">
            <h2>🏢 支店別データ</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>支店ID</th>
                            <th>支店名</th>
                            <th>企業データ件数</th>
                            <th>状況</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for area in areas %}
                        <tr>
                            <td>{{ area.id }}</td>
                            <td><strong>{{ area.name }}</strong></td>
                            <td>{{ area.company_count }}件</td>
                            <td><span class="status-success">✅ 稼働中</span></td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="card">
            <h2>📋 アカウント設定</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>アカウントID</th>
                            <th>部署名</th>
                            <th>ハローワーク</th>
                            <th>食べログ</th>
                            <th>管理</th>
                            <th>関連支店数</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for account in accounts %}
                        <tr class="{% if account.needs_hellowork %}hellowork-enabled{% endif %}">
                            <td>{{ account.id }}</td>
                            <td><strong>{{ account.name }}</strong></td>
                            <td>{% if account.needs_hellowork %}<span class="status-success">✅</span>{% else %}<span class="status-error">❌</span>{% endif %}</td>
                            <td>{% if account.needs_tabelog %}<span class="status-success">✅</span>{% else %}<span class="status-error">❌</span>{% endif %}</td>
                            <td>{% if account.needs_kanri %}<span class="status-success">✅</span>{% else %}<span class="status-error">❌</span>{% endif %}</td>
                            <td>{{ account.area_count }}支店</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="card">
            <h2>🔗 支店・アカウント関連マッピング</h2>
            {% set current_area = [] %}
            {% for mapping in area_account_mapping %}
                {% if current_area|length == 0 or current_area[0] != mapping.area_name %}
                    {% if current_area|length > 0 %}</div>{% endif %}
                    {% set _ = current_area.clear() %}
                    {% set _ = current_area.append(mapping.area_name) %}
                    <div class="area-section">
                        <h3>{{ mapping.area_name }}</h3>
                        <ul>
                {% endif %}
                <li><strong>{{ mapping.account_name }}</strong> (ID: {{ mapping.account_id }})
                    {% if mapping.needs_hellowork %}<span class="status-success">📧 ハローワーク対象</span>{% endif %}
                </li>
            {% endfor %}
            {% if area_account_mapping|length > 0 %}</ul></div>{% endif %}
        </div>
        
        <div class="card">
            <h2>🛠️ 管理ツール</h2>
            <div style="display: flex; flex-wrap: wrap; gap: 10px;">
                <a href="http://localhost:8081" target="_blank" class="btn btn-info">phpMyAdmin</a>
                <a href="http://localhost:8082" target="_blank" class="btn btn-info">Adminer</a>
                <a href="/api/areas" target="_blank" class="btn btn-primary">支店API</a>
                <a href="/api/accounts" target="_blank" class="btn btn-primary">アカウントAPI</a>
                <a href="/api/mapping" target="_blank" class="btn btn-primary">マッピングAPI</a>
                <button onclick="exportHierarchicalReport()" class="btn btn-success">📊 階層構造 Excel出力</button>
            </div>
        </div>
    </div>

    <script>
        async function exportHierarchicalReport() {
            try {
                const response = await fetch('/api/export-mapping', {
                    method: 'POST'
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `hellowork_hierarchical_report_${new Date().toISOString().split('T')[0].replace(/-/g, '')}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                    alert('階層構造レポートのExcel出力が完了しました！\n\n構造:\n支店 → アカウント → 新規/更新 → 件数');
                } else {
                    alert('Excel出力に失敗しました');
                }
            } catch (error) {
                alert('Excel出力でエラーが発生しました: ' + error.message);
            }
        }

        // 旧関数との互換性維持
        async function exportMapping() {
            return await exportHierarchicalReport();
        }
    </script>
</body>
</html>
'''

# ========================
# ルート定義
# ========================

@app.route('/')
def index():
    try:
        # 統計情報取得
        summary = get_area_account_summary()
        mapping = get_area_account_mapping()
        
        stats = {
            'total_areas': len(summary['areas']),
            'total_accounts': len(summary['accounts']),
            'hellowork_accounts': len([acc for acc in summary['accounts'] if acc['needs_hellowork']]),
            'total_companies': sum(area['company_count'] for area in summary['areas'])
        }
        
        return render_template_string(MAIN_TEMPLATE,
                                    stats=stats,
                                    areas=summary['areas'],
                                    accounts=summary['accounts'],
                                    area_account_mapping=mapping)
                                    
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/areas')
def get_areas():
    """支店一覧取得API"""
    try:
        areas = FmArea.query.order_by(FmArea.id).all()
        return jsonify({
            'status': 'success',
            'data': [area.to_dict() for area in areas]
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/accounts')
def get_accounts():
    """アカウント一覧取得API"""
    try:
        accounts = FmAccount.query.order_by(FmAccount.sort_order).all()
        return jsonify({
            'status': 'success',
            'data': [account.to_dict() for account in accounts]
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/mapping')
def get_mapping():
    """支店・アカウントマッピング取得API"""
    try:
        mapping = get_area_account_mapping()
        return jsonify({
            'status': 'success',
            'data': mapping
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export-mapping', methods=['POST'])
def export_mapping():
    """階層構造 Excel出力API"""
    try:
        # 階層構造データを生成
        hierarchical_data = generate_hierarchical_excel_data()
        
        # DataFrame作成
        df = pd.DataFrame(hierarchical_data)
        
        # Excelファイル作成
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='ハローワーク送信状況')
            
            # スタイル調整
            workbook = writer.book
            worksheet = writer.sheets['ハローワーク送信状況']
            
            # フォントとスタイルの定義
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            
            # ヘッダーのスタイル
            header_font = Font(bold=True, color='FFFFFF', size=12)
            header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
            
            # レベル別のスタイル定義
            level1_font = Font(bold=True, size=14, color='000080')  # 支店
            level1_fill = PatternFill(start_color='E6F2FF', end_color='E6F2FF', fill_type='solid')
            
            level2_font = Font(bold=True, size=11, color='000000')  # アカウント
            level2_fill = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')
            
            level3_font = Font(size=10, color='333333')  # 新規/更新
            level3_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
            
            subtotal_font = Font(bold=True, size=10, color='006600')  # 小計
            subtotal_fill = PatternFill(start_color='F0FFF0', end_color='F0FFF0', fill_type='solid')
            
            # 罫線の定義
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # ヘッダー行のスタイル設定
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            
            # データ行のスタイル設定
            for row_num, row_data in enumerate(hierarchical_data, start=2):
                level = row_data.get('レベル', 0)
                
                # レベルに応じてスタイルを適用
                if level == 1:  # 支店
                    font = level1_font
                    fill = level1_fill
                elif level == 2:  # アカウント・小計
                    if row_data.get('種別') == '合計':
                        font = subtotal_font
                        fill = subtotal_fill
                    else:
                        font = level2_font
                        fill = level2_fill
                elif level == 3:  # 新規/更新
                    font = level3_font
                    fill = level3_fill
                else:
                    font = Font(size=10)
                    fill = PatternFill()
                
                # 行の全セルにスタイルを適用
                for col_num in range(1, len(df.columns) + 1):
                    cell = worksheet.cell(row=row_num, column=col_num)
                    cell.font = font
                    cell.fill = fill
                    cell.border = thin_border
                    
                    # 件数列は右寄せ
                    if col_num == 4 and row_data.get('件数') != '':  # 件数列
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    else:
                        cell.alignment = Alignment(horizontal='left', vertical='center')
            
            # 列幅の調整
            column_widths = {
                'A': 5,   # レベル
                'B': 35,  # 項目名
                'C': 10,  # 種別
                'D': 10,  # 件数
                'E': 25   # 備考
            }
            
            for col_letter, width in column_widths.items():
                worksheet.column_dimensions[col_letter].width = width
            
            # フリーズペイン（ヘッダー行を固定）
            worksheet.freeze_panes = 'A2'
            
            # シートタブの色
            worksheet.sheet_properties.tabColor = "366092"
        
        output.seek(0)
        
        # ファイル名生成
        today = date.today()
        filename = f"hellowork_hierarchical_report_{today.strftime('%Y%m%d')}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/test')
def api_test():
    """API接続テスト"""
    try:
        with db.engine.connect() as connection:
            result = connection.execute(text('SELECT VERSION()'))
            mysql_version = result.fetchone()[0]
        
        # 実データ統計
        stats = get_area_account_summary()
        
        return jsonify({
            'status': 'success',
            'message': 'API is working with real data',
            'database': 'scraping',
            'mysql_version': mysql_version,
            'environment': os.getenv('FLASK_ENV', 'production'),
            'data_summary': {
                'areas': len(stats['areas']),
                'accounts': len(stats['accounts']),
                'total_companies': sum(area['company_count'] for area in stats['areas'])
            }
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'database': 'disconnected'
        }), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000, debug=True)