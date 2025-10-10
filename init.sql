-- 開発用データベースの初期化スクリプト
-- このファイルはMySQLコンテナ起動時に自動実行されます

-- ユーザー権限の設定
GRANT ALL PRIVILEGES ON *.* TO 'dev_user'@'%' WITH GRANT OPTION;
FLUSH PRIVILEGES;

USE development_db;

-- 支店マスタテーブル
CREATE TABLE IF NOT EXISTS fm_areas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL COMMENT '支店名',
    code VARCHAR(20) UNIQUE NOT NULL COMMENT '支店コード',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- アカウントマスタテーブル
CREATE TABLE IF NOT EXISTS fm_accounts (
    id INT NOT NULL PRIMARY KEY COMMENT 'アカウントID',
    area_id INT NOT NULL COMMENT '支店ID',
    name VARCHAR(100) NOT NULL COMMENT 'アカウント名',
    email VARCHAR(255) COMMENT 'メールアドレス',
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (area_id) REFERENCES fm_areas(id)
);

-- ハローワークデータテーブル
CREATE TABLE IF NOT EXISTS hellowork_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    fm_account_id INT NOT NULL COMMENT '送信先アカウントID',
    data_type ENUM('新規', '更新') NOT NULL COMMENT 'データ種別',
    job_title VARCHAR(255) COMMENT '求人タイトル',
    company_name VARCHAR(255) COMMENT '会社名',
    sent_date DATE NOT NULL COMMENT '送信日',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (fm_account_id) REFERENCES fm_accounts(id),
    INDEX idx_sent_date (sent_date),
    INDEX idx_account_date (fm_account_id, sent_date)
);

-- 支店データの初期投入
INSERT IGNORE INTO fm_areas (id, name, code) VALUES 
(1, '関西支店', 'KANSAI'),
(2, '関東支店', 'KANTO'),
(3, '中部支店', 'CHUBU'),
(4, '九州支店', 'KYUSHU');

-- アカウントデータの初期投入
INSERT IGNORE INTO fm_accounts (id, area_id, name, email) VALUES 
(90000007, 1, '関西一部', 'kansai1@example.com'),
(90000008, 1, '関西二部', 'kansai2@example.com'),
(90000009, 2, '関東一部', 'kanto1@example.com'),
(90000010, 2, '関東二部', 'kanto2@example.com'),
(90000011, 3, '中部一部', 'chubu1@example.com'),
(90000012, 4, '九州一部', 'kyushu1@example.com');

-- サンプルハローワークデータの投入
INSERT IGNORE INTO hellowork_data (fm_account_id, data_type, job_title, company_name, sent_date) VALUES 
-- 関西一部のデータ（今日）
(90000007, '新規', 'システムエンジニア', '株式会社テクノロジー', CURDATE()),
(90000007, '新規', 'Webデザイナー', '株式会社クリエイト', CURDATE()),
(90000007, '更新', '営業職', '商事株式会社', CURDATE()),

-- 関西二部のデータ（今日）
(90000008, '新規', 'プログラマー', 'IT企業株式会社', CURDATE()),
(90000008, '更新', '事務職', '総合商社', CURDATE()),
(90000008, '更新', 'マーケティング', 'マーケ会社', CURDATE()),

-- 関東一部のデータ（今日）
(90000009, '新規', 'データアナリスト', 'データ分析会社', CURDATE()),
(90000009, '新規', 'エンジニア', 'スタートアップ', CURDATE()),
(90000009, '新規', 'デザイナー', 'デザイン事務所', CURDATE()),
(90000009, '更新', 'コンサルタント', 'コンサル会社', CURDATE()),

-- 昨日のデータ
(90000007, '新規', '経理職', '会計事務所', DATE_SUB(CURDATE(), INTERVAL 1 DAY)),
(90000008, '更新', '人事職', '人材会社', DATE_SUB(CURDATE(), INTERVAL 1 DAY)),
(90000009, '新規', '企画職', '企画会社', DATE_SUB(CURDATE(), INTERVAL 1 DAY)),

-- 一週間前のデータ
(90000007, '新規', '技術職', '製造業', DATE_SUB(CURDATE(), INTERVAL 7 DAY)),
(90000008, '更新', 'サービス職', 'サービス業', DATE_SUB(CURDATE(), INTERVAL 7 DAY));

-- 既存のuserテーブル（既存機能との互換性維持）
CREATE TABLE IF NOT EXISTS user (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    email VARCHAR(120) UNIQUE NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 既存のsample_dataテーブル（既存機能との互換性維持）
CREATE TABLE IF NOT EXISTS sample_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    title VARCHAR(255) NOT NULL,
    description TEXT,
    status ENUM('active', 'inactive') DEFAULT 'active',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- 既存データの投入（重複を避けるため）
INSERT IGNORE INTO user (name, email) VALUES 
('田中太郎', 'tanaka@example.com'),
('佐藤花子', 'sato@example.com'),
('鈴木一郎', 'suzuki@example.com');

INSERT IGNORE INTO sample_data (title, description, status) VALUES 
('サンプル1', 'これは最初のサンプルデータです', 'active'),
('サンプル2', 'これは2番目のサンプルデータです', 'active'),
('サンプル3', 'これは3番目のサンプルデータです', 'inactive');