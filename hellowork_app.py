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
# データモデル定義
# ========================

class FmArea(db.Model):
    __tablename__ = 'fm_areas'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, comment='支店名')
    code = db.Column(db.String(20), unique=True, nullable=False, comment='支店コード')
    created_at = db.Column(db.TIMESTAMP, default=func.current_timestamp())
    updated_at = db.Column(db.TIMESTAMP, default=func.current_timestamp(), onupdate=func.current_timestamp())
    
    # リレーション
    accounts = db.relationship('FmAccount', backref='area', lazy=True)
    
    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'code': self.code
        }

class FmAccount(db.Model):
    __tablename__ = 'fm_accounts'
    
    id = db.Column(db.Integer, primary_key=True, comment='アカウントID')
    area_id = db.Column(db.Integer, db.ForeignKey('fm_areas.id'), nullable=False, comment='支店ID')
    name = db.Column(db.String(100), nullable=False, comment='アカウント名')
    email = db.Column(db.String(255), comment='メールアドレス')
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.TIMESTAMP, default=func.current_timestamp())
    updated_at = db.Column(db.TIMESTAMP, default=func.current_timestamp(), onupdate=func.current_timestamp())
    
    # リレーション
    hellowork_data = db.relationship('HelloworkData', backref='account', lazy=True)
    
    def to_dict(self):
        return {
            'id': self.id,
            'area_id': self.area_id,
            'area_name': self.area.name if self.area else None,
            'name': self.name,
            'email': self.email,
            'is_active': self.is_active
        }

class HelloworkData(db.Model):
    __tablename__ = 'hellowork_data'
    
    id = db.Column(db.Integer, primary_key=True)
    fm_account_id = db.Column(db.Integer, db.ForeignKey('fm_accounts.id'), nullable=False, comment='送信先アカウントID')
    data_type = db.Column(db.Enum('新規', '更新'), nullable=False, comment='データ種別')
    job_title = db.Column(db.String(255), comment='求人タイトル')
    company_name = db.Column(db.String(255), comment='会社名')
    sent_date = db.Column(db.Date, nullable=False, comment='送信日')
    created_at = db.Column(db.TIMESTAMP, default=func.current_timestamp())
    
    def to_dict(self):
        return {
            'id': self.id,
            'fm_account_id': self.fm_account_id,
            'data_type': self.data_type,
            'job_title': self.job_title,
            'company_name': self.company_name,
            'sent_date': self.sent_date.isoformat() if self.sent_date else None,
            'account_name': self.account.name if self.account else None,
            'area_name': self.account.area.name if self.account and self.account.area else None
        }

# 既存のモデル（互換性維持）
class User(db.Model):
    __tablename__ = 'user'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    created_at = db.Column(db.TIMESTAMP, default=func.current_timestamp())

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'email': self.email
        }

# ========================
# 集計関数
# ========================

def get_daily_report_data(date_from=None, date_to=None, area_ids=None):
    """日別レポートデータを取得"""
    
    # デフォルト日付設定（過去30日）
    if not date_to:
        date_to = date.today()
    if not date_from:
        date_from = date_to - timedelta(days=30)
    
    # ベースクエリ
    query = db.session.query(
        HelloworkData.sent_date,
        FmArea.id.label('area_id'),
        FmArea.name.label('area_name'),
        FmAccount.id.label('account_id'),
        FmAccount.name.label('account_name'),
        HelloworkData.data_type,
        func.count(HelloworkData.id).label('count')
    ).join(
        FmAccount, HelloworkData.fm_account_id == FmAccount.id
    ).join(
        FmArea, FmAccount.area_id == FmArea.id
    ).filter(
        and_(
            HelloworkData.sent_date >= date_from,
            HelloworkData.sent_date <= date_to
        )
    )
    
    # 支店フィルター
    if area_ids:
        query = query.filter(FmArea.id.in_(area_ids))
    
    # グループ化とソート
    query = query.group_by(
        HelloworkData.sent_date,
        FmArea.id,
        FmAccount.id,
        HelloworkData.data_type
    ).order_by(
        HelloworkData.sent_date.desc(),
        FmArea.id,
        FmAccount.id,
        HelloworkData.data_type
    )
    
    return query.all()

# ========================
# HTMLテンプレート
# ========================

MAIN_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ハローワークデータ管理システム</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .card { background: white; border-radius: 10px; padding: 25px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .card h2 { color: #333; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 3px solid #667eea; }
        .filter-section { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 20px; }
        .form-group { display: flex; flex-direction: column; }
        .form-group label { font-weight: bold; margin-bottom: 5px; color: #555; }
        .form-group input, .form-group select { padding: 10px; border: 1px solid #ddd; border-radius: 5px; font-size: 14px; }
        .btn { padding: 12px 25px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; font-weight: bold; text-decoration: none; display: inline-block; text-align: center; transition: all 0.3s; }
        .btn-primary { background: #667eea; color: white; }
        .btn-primary:hover { background: #5a6fd8; transform: translateY(-2px); }
        .btn-success { background: #28a745; color: white; }
        .btn-success:hover { background: #218838; transform: translateY(-2px); }
        .btn-info { background: #17a2b8; color: white; }
        .btn-info:hover { background: #138496; transform: translateY(-2px); }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; text-align: center; }
        .stat-card h3 { font-size: 2em; margin-bottom: 5px; }
        .stat-card p { opacity: 0.9; }
        .table-container { overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f8f9fa; font-weight: bold; color: #495057; }
        tr:hover { background-color: #f8f9fa; }
        .status-success { color: #28a745; font-weight: bold; }
        .status-error { color: #dc3545; font-weight: bold; }
        .actions { margin-top: 20px; display: flex; gap: 10px; flex-wrap: wrap; }
        .loading { text-align: center; padding: 20px; color: #666; }
        @media (max-width: 768px) {
            .header h1 { font-size: 2em; }
            .filter-section { grid-template-columns: 1fr; }
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏢 ハローワークデータ管理システム</h1>
            <p>支店・アカウント別のデータ送信状況を管理・分析</p>
        </div>
        
        <div class="card">
            <h2>📊 システム情報</h2>
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>{{ stats.total_areas }}</h3>
                    <p>登録支店数</p>
                </div>
                <div class="stat-card">
                    <h3>{{ stats.total_accounts }}</h3>
                    <p>アクティブアカウント</p>
                </div>
                <div class="stat-card">
                    <h3>{{ stats.today_data }}</h3>
                    <p>本日の送信件数</p>
                </div>
                <div class="stat-card">
                    <h3>{{ stats.total_data }}</h3>
                    <p>総データ件数</p>
                </div>
            </div>
        </div>
        
        <div class="card">
            <h2>🔍 レポート生成</h2>
            <form id="reportForm">
                <div class="filter-section">
                    <div class="form-group">
                        <label for="dateFrom">開始日</label>
                        <input type="date" id="dateFrom" name="date_from" value="{{ default_date_from }}">
                    </div>
                    <div class="form-group">
                        <label for="dateTo">終了日</label>
                        <input type="date" id="dateTo" name="date_to" value="{{ default_date_to }}">
                    </div>
                    <div class="form-group">
                        <label for="areaFilter">支店フィルター</label>
                        <select id="areaFilter" name="area_ids" multiple>
                            {% for area in areas %}
                            <option value="{{ area.id }}">{{ area.name }}</option>
                            {% endfor %}
                        </select>
                    </div>
                </div>
                <div class="actions">
                    <button type="button" onclick="loadReport()" class="btn btn-primary">📋 レポート表示</button>
                    <button type="button" onclick="exportExcel()" class="btn btn-success">📊 Excel出力</button>
                    <a href="/api/areas" target="_blank" class="btn btn-info">📝 API テスト</a>
                </div>
            </form>
        </div>
        
        <div class="card">
            <h2>📈 レポート結果</h2>
            <div id="reportContent">
                <div class="loading">フィルターを設定してレポートを表示してください</div>
            </div>
        </div>
        
        <div class="card">
            <h2>🛠️ 管理ツール</h2>
            <div class="actions">
                <a href="http://localhost:8081" target="_blank" class="btn btn-info">phpMyAdmin</a>
                <a href="http://localhost:8082" target="_blank" class="btn btn-info">Adminer</a>
                <a href="/api/daily-report" target="_blank" class="btn btn-primary">API確認</a>
            </div>
        </div>
    </div>

    <script>
        // レポート読み込み
        async function loadReport() {
            const form = document.getElementById('reportForm');
            const formData = new FormData(form);
            const params = new URLSearchParams();
            
            for (let [key, value] of formData.entries()) {
                params.append(key, value);
            }
            
            document.getElementById('reportContent').innerHTML = '<div class="loading">データを読み込み中...</div>';
            
            try {
                const response = await fetch(`/api/daily-report?${params}`);
                const data = await response.json();
                
                if (data.status === 'success') {
                    displayReport(data.data);
                } else {
                    document.getElementById('reportContent').innerHTML = 
                        `<div class="status-error">エラー: ${data.message}</div>`;
                }
            } catch (error) {
                document.getElementById('reportContent').innerHTML = 
                    '<div class="status-error">データの読み込みに失敗しました</div>';
            }
        }
        
        // レポート表示
        function displayReport(data) {
            if (data.length === 0) {
                document.getElementById('reportContent').innerHTML = 
                    '<div class="loading">指定された期間にデータがありません</div>';
                return;
            }
            
            let html = '<div class="table-container"><table><thead><tr>';
            html += '<th>送信日</th><th>支店</th><th>アカウント</th><th>種別</th><th>件数</th>';
            html += '</tr></thead><tbody>';
            
            data.forEach(row => {
                html += `<tr>
                    <td>${row.sent_date}</td>
                    <td>${row.area_name}</td>
                    <td>${row.account_name}</td>
                    <td><span class="status-${row.data_type === '新規' ? 'success' : 'info'}">${row.data_type}</span></td>
                    <td><strong>${row.count}</strong></td>
                </tr>`;
            });
            
            html += '</tbody></table></div>';
            document.getElementById('reportContent').innerHTML = html;
        }
        
        // Excel出力
        async function exportExcel() {
            const form = document.getElementById('reportForm');
            const formData = new FormData(form);
            
            try {
                const response = await fetch('/api/export-excel', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `hellowork_report_${new Date().toISOString().split('T')[0]}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                } else {
                    alert('Excel出力に失敗しました');
                }
            } catch (error) {
                alert('Excel出力でエラーが発生しました');
            }
        }
        
        // 初期読み込み
        window.onload = function() {
            loadReport();
        };
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
        stats = {
            'total_areas': FmArea.query.count(),
            'total_accounts': FmAccount.query.filter_by(is_active=True).count(),
            'today_data': HelloworkData.query.filter_by(sent_date=date.today()).count(),
            'total_data': HelloworkData.query.count()
        }
        
        # 支店一覧取得
        areas = FmArea.query.order_by(FmArea.id).all()
        
        # デフォルト日付
        today = date.today()
        week_ago = today - timedelta(days=7)
        
        return render_template_string(MAIN_TEMPLATE,
                                    stats=stats,
                                    areas=areas,
                                    default_date_from=week_ago.isoformat(),
                                    default_date_to=today.isoformat())
                                    
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
        area_id = request.args.get('area_id', type=int)
        
        query = FmAccount.query.filter_by(is_active=True)
        if area_id:
            query = query.filter_by(area_id=area_id)
        
        accounts = query.order_by(FmAccount.id).all()
        return jsonify({
            'status': 'success',
            'data': [account.to_dict() for account in accounts]
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/daily-report')
def api_daily_report():
    """日別レポートAPI"""
    try:
        # パラメータ取得
        date_from_str = request.args.get('date_from')
        date_to_str = request.args.get('date_to')
        area_ids_str = request.args.getlist('area_ids')
        
        # 日付変換
        date_from = datetime.strptime(date_from_str, '%Y-%m-%d').date() if date_from_str else None
        date_to = datetime.strptime(date_to_str, '%Y-%m-%d').date() if date_to_str else None
        area_ids = [int(aid) for aid in area_ids_str if aid] if area_ids_str else None
        
        # データ取得
        results = get_daily_report_data(date_from, date_to, area_ids)
        
        # レスポンス形成
        data = []
        for row in results:
            data.append({
                'sent_date': row.sent_date.isoformat(),
                'area_id': row.area_id,
                'area_name': row.area_name,
                'account_id': row.account_id,
                'account_name': row.account_name,
                'data_type': row.data_type,
                'count': row.count
            })
        
        return jsonify({
            'status': 'success',
            'data': data,
            'total_records': len(data)
        })
        
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    """Excel出力API"""
    try:
        # パラメータ取得
        date_from_str = request.form.get('date_from')
        date_to_str = request.form.get('date_to')
        area_ids_str = request.form.getlist('area_ids')
        
        # 日付変換
        date_from = datetime.strptime(date_from_str, '%Y-%m-%d').date() if date_from_str else None
        date_to = datetime.strptime(date_to_str, '%Y-%m-%d').date() if date_to_str else None
        area_ids = [int(aid) for aid in area_ids_str if aid] if area_ids_str else None
        
        # データ取得
        results = get_daily_report_data(date_from, date_to, area_ids)
        
        # DataFrame作成
        data = []
        for row in results:
            data.append({
                '送信日': row.sent_date.isoformat(),
                '支店名': row.area_name,
                'アカウント名': row.account_name,
                'データ種別': row.data_type,
                '件数': row.count
            })
        
        df = pd.DataFrame(data)
        
        # Excelファイル作成
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='ハローワークレポート')
            
            # スタイル調整
            workbook = writer.book
            worksheet = writer.sheets['ハローワークレポート']
            
            # ヘッダーのスタイル
            header_font = Font(bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
            
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 列幅調整
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        
        # ファイル名生成
        filename = f"hellowork_report_{date.today().isoformat()}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# 既存API（互換性維持）
@app.route('/api/test')
def api_test():
    """API接続テスト"""
    try:
        with db.engine.connect() as connection:
            result = connection.execute(text('SELECT VERSION()'))
            mysql_version = result.fetchone()[0]
        
        return jsonify({
            'status': 'success',
            'message': 'API is working',
            'database': 'connected',
            'mysql_version': mysql_version,
            'environment': os.getenv('FLASK_ENV', 'production'),
            'new_features': ['ハローワークデータ管理', 'Excel出力', '支店別レポート']
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'database': 'disconnected'
        }), 500

@app.route('/api/init-db')
def init_database():
    """データベース初期化API"""
    try:
        db.create_all()
        
        # 既存のUser テーブルチェック
        user_count = User.query.count()
        if user_count == 0:
            sample_users = [
                User(name='田中太郎', email='tanaka@example.com'),
                User(name='佐藤花子', email='sato@example.com'),
                User(name='鈴木一郎', email='suzuki@example.com')
            ]
            for user in sample_users:
                db.session.add(user)
        
        db.session.commit()
        
        # 新テーブルの状況確認
        stats = {
            'fm_areas': FmArea.query.count(),
            'fm_accounts': FmAccount.query.count(),
            'hellowork_data': HelloworkData.query.count(),
            'users': User.query.count()
        }
        
        return jsonify({
            'status': 'success',
            'message': 'データベースが正常に初期化されました',
            'table_counts': stats
        })
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'データベース初期化エラー: {str(e)}'
        }), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000, debug=True)