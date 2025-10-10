# 📊 ハローワークデータ Excel出力システム

支店・アカウント別のハローワークデータを階層構造でExcel出力する専用アプリケーションです。

## 🎯 概要

このシステムは、ハローワークの求人データを支店→アカウント→新規/更新→件数の階層構造で整理し、美しくフォーマットされたExcelファイルとして出力することを目的としています。

### ✨ 主な特徴

- **🎨 美しい階層表示**: 支店・アカウント・データ種別を視覚的に分かりやすく表示
- **📈 自動集計**: 各レベルでの小計・合計を自動計算
- **🎨 スタイル適用**: カラー・フォント・罫線による見やすいフォーマット
- **📱 シンプルUI**: Excel出力のみに特化したクリーンなインターフェース
- **🔒 セキュア**: SQLインジェクション対策済み、安全なデータアクセス

## 🏗️ システム構成

- **Python 3.11**: Flask Webアプリケーション
- **MySQL 8.0**: 'scraping'データベース（実データ）
- **phpMyAdmin**: データベース管理ツール
- **Adminer**: 軽量データベース管理ツール
- **openpyxl**: Excelファイル生成ライブラリ

## 📋 前提条件

- Docker Desktop
- Docker Compose
- Webブラウザ（Chrome/Firefox/Edge推奨）

## 🚀 セットアップと起動

### 1. プロジェクト準備
```bash
# プロジェクトディレクトリに移動
cd c:\work\folder
```

### 2. 環境変数の確認
`.env` ファイルが正しく設定されていることを確認：

```bash
# MySQL データベース設定（実データベース 'scraping' を使用）
MYSQL_ROOT_PASSWORD=rootpassword123
MYSQL_DATABASE=scraping
MYSQL_USER=dev_user
MYSQL_PASSWORD=dev_password123
```

### 3. システム起動
```bash
# 全サービスを一括起動
docker-compose up -d

# 起動状況の確認
docker-compose ps
```

### 4. アクセス確認

起動後、以下のURLにアクセス可能：

- **📊 Excel出力アプリ**: http://localhost:8000
- **🛠️ phpMyAdmin**: http://localhost:8081
- **⚙️ Adminer**: http://localhost:8082

## � 使用方法

### 基本的な流れ

1. **アクセス**: ブラウザで http://localhost:8000 を開く
2. **確認**: システム統計（支店数・アカウント数）を確認
3. **出力**: 「📊 Excel ファイル出力」ボタンをクリック
4. **ダウンロード**: `hellowork_hierarchical_report_YYYYMMDD.xlsx` が自動ダウンロード

### Excel出力内容の例

```
📍 関西支店
  📂 関西一部
    📝 新規 → 15件
    🔄 更新 → 8件
  └─ 小計 → 23件
  📂 関西二部
    📝 新規 → 12件
    🔄 更新 → 6件
  └─ 小計 → 18件
🔢 関西支店 合計 → 41件

📍 関東支店
  📂 関東一部
    📝 新規 → 20件
    🔄 更新 → 10件
  └─ 小計 → 30件
```

### データベース管理

#### phpMyAdmin経由
- **URL**: http://localhost:8081
- **サーバー**: `mysql`
- **データベース**: `scraping`
- **ユーザー名**: `dev_user`
- **パスワード**: `dev_password123`

#### Adminer経由
- **URL**: http://localhost:8082
- **システム**: `MySQL`
- **サーバー**: `mysql:3306`
- **データベース**: `scraping`

## 📁 プロジェクト構造

```
c:\work\folder\
├── 📄 docker-compose.yml     # マルチコンテナ構成定義
├── 🐳 Dockerfile            # Pythonアプリ用イメージ
├── 📦 requirements.txt      # Python依存パッケージ
├── ⚙️ .env                  # 環境変数設定
├── 🐍 excel_only_app.py     # メインアプリケーション（Excel出力専用）
├── 🗃️ init.sql             # データベース初期化スクリプト
├── 🔧 phpmyadmin-config.ini # phpMyAdmin設定
└── 📖 README.md             # このファイル
```

## 🛠️ 開発・メンテナンスコマンド

### サービス管理
```bash
# 全サービス起動
docker-compose up -d

# 特定サービス再起動
docker-compose restart python-app

# サービス停止
docker-compose down

# イメージ再ビルド
docker-compose build python-app
docker-compose up -d
```

### ログ確認
```bash
# アプリケーションログ
docker-compose logs python-app

# リアルタイムログ監視
docker-compose logs -f python-app

# 最新10行のログ
docker-compose logs python-app | Select-Object -Last 10
```

### APIテスト
```bash
# システム状態確認
curl http://localhost:8000/api/test

# Excel出力テスト
Invoke-WebRequest -Uri "http://localhost:8000/api/expodort-excel" -Method POST
```

## � データベーススキーマ

### 主要テーブル

#### fm_areas（支店情報）
- `id`: 支店ID
- `area_name_ja`: 支店名（日本語）
- `area_name_en`: 支店名（英語）
- `fm_login_account_id`: ログインアカウントID
- `fm_login_account_pass`: ログインパスワード

#### fm_accounts（アカウント情報）
- `id`: アカウントID
- `department_name`: 部署名
- `sort_order`: 表示順序
- `needs_hellowork`: ハローワーク要否フラグ（1=必要）

#### fm_area_accounts（支店-アカウント関連）
- `fm_area_id`: 支店ID（外部キー）
- `fm_account_id`: アカウントID（外部キー）

## 🎨 Excel出力スタイル

### カラーパレット
- **🔵 支店ヘッダー**: 青系（#E6F2FF背景、#000080文字）
- **🔷 アカウント**: 水色系（#F0F8FF背景、#000000文字）
- **⚪ データ項目**: 白背景（#FFFFFF）
- **🟢 小計**: 緑系（#F0FFF0背景、#006600文字）
- **🔴 支店合計**: 赤系（#FFE6E6背景、#800000文字）

### フォーマット特徴
- **アイコン**: 📍📂📝🔄🔢などで項目種別を視覚化
- **階層インデント**: スペースとアイコンで階層構造を表現
- **フォントサイズ**: レベルに応じて14px〜10pxで調整
- **罫線**: 全セルに薄い罫線を適用
- **列幅自動調整**: 内容に応じた最適幅

## 🐛 トラブルシューティング

### よくある問題と解決法

#### 1. ポート競合エラー
```bash
# 使用中ポートの確認
netstat -an | findstr :8000
netstat -an | findstr :3307

# docker-compose.ymlでポート変更
ports:
  - "8001:8000"  # 8000 → 8001に変更
```

#### 2. データベース接続エラー
```bash
# MySQLコンテナ状態確認
docker-compose ps mysql

# MySQL接続テスト
docker-compose exec mysql mysql -u dev_user -p scraping

# データベース再初期化
docker-compose down -v
docker-compose up -d
```

#### 3. Excel出力エラー
```bash
# Pythonアプリログ確認
docker-compose logs python-app

# 依存関係再インストール
docker-compose build python-app --no-cache
docker-compose up -d
```

#### 4. メモリ不足エラー
```bash
# Docker使用リソース確認
docker system df

# 不要イメージ・コンテナの削除
docker system prune
```

### ログレベル別確認

```bash
# エラーログのみ
docker-compose logs python-app | findstr ERROR

# 警告とエラー
docker-compose logs python-app | findstr "WARNING\|ERROR"

# リクエストログ
docker-compose logs python-app | findstr "GET\|POST"
```

## 🔒 セキュリティ対策

### 実装済み対策
- **SQLインジェクション防止**: SQLAlchemy ORMによる安全なクエリ
- **XSS対策**: テンプレートエスケープの実装
- **CSRF対策**: POSTリクエストの適切な処理
- **入力値検証**: データ型と範囲のチェック
- **エラーハンドリング**: 機密情報を含まない安全なエラーメッセージ

### 開発環境での注意事項
- デバッグモードは本番環境では無効化
- デフォルトパスワードは開発専用
- `.env`ファイルはバージョン管理対象外
- ログには機密情報を出力しない

## � カスタマイズ

### 新しいPython依存関係の追加
1. `requirements.txt`にパッケージ追加
2. `docker-compose build python-app`でビルド
3. `docker-compose up -d`で再起動

### Excel出力スタイルの変更
`excel_only_app.py`の以下部分を修正：
- カラーパレット: `PatternFill`のcolor値
- フォント設定: `Font`のサイズ・色
- 列幅調整: `column_widths`辞書

### データベース設定の変更
1. `.env`ファイル編集
2. `docker-compose down -v`でボリューム削除
3. `docker-compose up -d`で再作成

## � パフォーマンス情報

### 処理速度目安
- **Excel生成時間**: 約1-3秒（データ量により変動）
- **ファイルサイズ**: 約8-12KB（28アカウント×7支店）
- **メモリ使用量**: 約50-100MB（ピーク時）
- **同時接続数**: 最大10接続（開発環境）

### 最適化ポイント
- データベースクエリの効率化
- Excelスタイル適用の最適化
- メモリ使用量の監視
- レスポンス時間の計測

## 📞 サポート・問い合わせ

### 確認事項
1. **Docker環境**: Dockerバージョン確認
2. **ポート状況**: 8000, 3307, 8081, 8082の利用可能性
3. **エラーログ**: 詳細なエラーメッセージ
4. **システムリソース**: CPU・メモリ・ディスク容量

### 緊急時の復旧手順
```bash
# 1. 全サービス停止
docker-compose down

# 2. ボリューム削除（データ初期化）
docker-compose down -v

# 3. イメージ再ビルド
docker-compose build --no-cache

# 4. サービス再起動
docker-compose up -d
```

---

## 🎉 最新の更新履歴

- **v1.0.0** (2025-10-08): Excel出力専用システムリリース
- 階層構造表示機能実装
- セキュリティ対策強化
- リアルデータ連携完了

---

**Happy Excel Export! �✨**