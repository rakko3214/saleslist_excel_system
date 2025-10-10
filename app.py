from flask import Flask, render_template_string, jsonify
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from dotenv import load_dotenv
import os
import pymysql

# 環境変数をロード
load_dotenv()

app = Flask(__name__)

# データベース設定
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY')

db = SQLAlchemy(app)

# サンプルモデル
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'email': self.email
        }

# HTMLテンプレート
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Python Docker Development Environment</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        .status { padding: 10px; margin: 10px 0; border-radius: 5px; }
        .success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .error { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info { background-color: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .button { background-color: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; margin: 5px; }
    </style>
</head>
<body>
    <h1>Python Docker 開発環境</h1>
    
    <div class="info">
        <h3>環境情報:</h3>
        <p><strong>Python:</strong> {{ python_version }}</p>
        <p><strong>Flask:</strong> {{ flask_version }}</p>
        <p><strong>データベース:</strong> MySQL</p>
    </div>
    
    <div class="status {{ db_status_class }}">
        <h3>データベース接続状態:</h3>
        <p>{{ db_status_message }}</p>
    </div>
    
    {% if users %}
    <h3>ユーザーデータ:</h3>
    <table>
        <tr>
            <th>ID</th>
            <th>名前</th>
            <th>メールアドレス</th>
        </tr>
        {% for user in users %}
        <tr>
            <td>{{ user.id }}</td>
            <td>{{ user.name }}</td>
            <td>{{ user.email }}</td>
        </tr>
        {% endfor %}
    </table>
    {% endif %}
    
    <div style="margin-top: 30px;">
        <h3>管理ツール:</h3>
        <a href="http://localhost:8080" target="_blank" class="button">phpMyAdmin を開く</a>
        <a href="/api/test" target="_blank" class="button">API テスト</a>
        <a href="/api/init-db" class="button">データベース初期化</a>
    </div>
</body>
</html>
'''

@app.route('/')
def index():
    try:
        # データベース接続テスト
        with db.engine.connect() as connection:
            connection.execute(text('SELECT 1'))
        db_status_class = 'success'
        db_status_message = '✅ データベースに正常に接続されています'
        
        # ユーザーデータを取得
        users = User.query.all()
        
    except Exception as e:
        db_status_class = 'error'
        db_status_message = f'❌ データベース接続エラー: {str(e)}'
        users = []
    
    import sys
    import flask
    
    return render_template_string(HTML_TEMPLATE,
                                python_version=sys.version.split()[0],
                                flask_version=flask.__version__,
                                db_status_class=db_status_class,
                                db_status_message=db_status_message,
                                users=users)

@app.route('/api/test')
def api_test():
    """API接続テスト"""
    try:
        # データベース接続テスト
        with db.engine.connect() as connection:
            result = connection.execute(text('SELECT VERSION()'))
            mysql_version = result.fetchone()[0]
        
        return jsonify({
            'status': 'success',
            'message': 'API is working',
            'database': 'connected',
            'mysql_version': mysql_version,
            'environment': os.getenv('FLASK_ENV', 'production')
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'database': 'disconnected'
        }), 500

@app.route('/api/init-db')
def init_database():
    """データベースとサンプルデータの初期化"""
    try:
        # テーブルを作成
        db.create_all()
        
        # サンプルデータが既に存在するかチェック
        if User.query.count() == 0:
            # サンプルユーザーを追加
            sample_users = [
                User(name='田中太郎', email='tanaka@example.com'),
                User(name='佐藤花子', email='sato@example.com'),
                User(name='鈴木一郎', email='suzuki@example.com')
            ]
            
            for user in sample_users:
                db.session.add(user)
            
            db.session.commit()
            message = 'データベースとサンプルデータを初期化しました'
        else:
            message = 'データベースは既に初期化されています'
        
        return jsonify({
            'status': 'success',
            'message': message,
            'user_count': User.query.count()
        })
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'データベース初期化エラー: {str(e)}'
        }), 500

@app.route('/api/users')
def get_users():
    """ユーザー一覧取得API"""
    try:
        users = User.query.all()
        return jsonify({
            'status': 'success',
            'users': [user.to_dict() for user in users]
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000, debug=True)