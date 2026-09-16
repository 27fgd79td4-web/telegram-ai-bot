<?php
/**
 * install.php — Ma'lumotlar bazasi jadvallarini yaratish.
 * Bir marta ishlatiladi:  php install.php
 */

require __DIR__ . '/vendor/autoload.php';

use Dotenv\Dotenv;

$dotenv = Dotenv::createImmutable(__DIR__);
if (file_exists(__DIR__ . '/.env')) $dotenv->load();

function getEnvVal(string $key, mixed $default = null): mixed {
    $val = $_ENV[$key] ?? $_SERVER[$key] ?? getenv($key);
    return ($val !== false && $val !== null && $val !== '') ? $val : $default;
}

$dbPath = (string)getEnvVal('DATABASE_PATH', './data/bot.db');
$dataDir = (string)getEnvVal('DATA_DIR', './data');
foreach ([$dataDir, (string)getEnvVal('TEMP_DIR', './temp'), (string)getEnvVal('LOGS_DIR', './logs'),
          (string)getEnvVal('UPLOADS_DIR', './uploads'), (string)getEnvVal('DOWNLOADS_DIR', './downloads')] as $d) {
    if (!is_dir($d)) mkdir($d, 0777, true);
}

$pdo = new PDO('sqlite:' . $dbPath);
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('PRAGMA foreign_keys = ON');

$pdo->exec("
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telegram_id INTEGER UNIQUE NOT NULL,
    username TEXT,
    first_name TEXT,
    language TEXT DEFAULT 'uz',
    is_admin INTEGER DEFAULT 0,
    is_blocked INTEGER DEFAULT 0,
    daily_limit_override INTEGER,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    last_active TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_users_tg ON users(telegram_id);

CREATE TABLE IF NOT EXISTS conversations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telegram_user_id INTEGER NOT NULL,
    title TEXT DEFAULT 'Yangi suhbat',
    is_active INTEGER DEFAULT 1,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (telegram_user_id) REFERENCES users(telegram_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_conv_user ON conversations(telegram_user_id);

CREATE TABLE IF NOT EXISTS messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    conversation_id INTEGER NOT NULL,
    role TEXT NOT NULL,
    content TEXT NOT NULL,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (conversation_id) REFERENCES conversations(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_msg_conv ON messages(conversation_id);

CREATE TABLE IF NOT EXISTS usages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telegram_user_id INTEGER NOT NULL,
    request_type TEXT NOT NULL,
    tokens INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (telegram_user_id) REFERENCES users(telegram_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_usage_user ON usages(telegram_user_id);
CREATE INDEX IF NOT EXISTS idx_usage_date ON usages(created_at);

CREATE TABLE IF NOT EXISTS files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telegram_user_id INTEGER NOT NULL,
    filename TEXT NOT NULL,
    file_type TEXT NOT NULL,
    file_size INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (telegram_user_id) REFERENCES users(telegram_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS activity_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telegram_user_id INTEGER NOT NULL,
    username TEXT,
    full_name TEXT,
    action TEXT NOT NULL,
    user_query TEXT,
    ai_response TEXT,
    tokens INTEGER DEFAULT 0,
    status TEXT DEFAULT 'OK',
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (telegram_user_id) REFERENCES users(telegram_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_activity_user ON activity_logs(telegram_user_id);
CREATE INDEX IF NOT EXISTS idx_activity_date ON activity_logs(created_at);
");

echo "✅ Database tayyor: {$dbPath}\n";