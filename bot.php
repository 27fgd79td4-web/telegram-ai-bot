<?php
/**
 * ============================================================================
 * UNIVERSAL TELEGRAM AI ASSISTANT BOT — ALL-IN-ONE (PHP 8.1+)
 * ============================================================================
 * Texnologiyalar:
 *  - PHP 8.1+
 *  - GuzzleHTTP (Telegram & Google Gemini)
 *  - SQLite (PDO)
 *  - PDF/DOCX/XLSX/PPTX parsers
 * ============================================================================
 */

declare(strict_types=1);

require __DIR__ . '/vendor/autoload.php';

use Dotenv\Dotenv;
use GuzzleHttp\Client;
use GuzzleHttp\Exception\GuzzleException;
use Smalot\PdfParser\Parser as PdfParser;
use PhpOffice\PhpWord\IOFactory as WordIOFactory;
use PhpOffice\PhpSpreadsheet\IOFactory as SpreadsheetIOFactory;
use PhpOffice\PhpPresentation\IOFactory as PresentationIOFactory;
use Dompdf\Dompdf;
use Dompdf\Options;

// ============================================================================
// SYSTEM PROMPTS
// ============================================================================

class Prompts {
    public const SYSTEM = <<<EOT
Siz Universal AI Assistant — Telegram bot orqali ishlaydigan aqlli, samimiy va professional yordamchisiz.
Asosiy vazifangiz: foydalanuvchilarning savollariga aniq, tushunarli va chiroyli formatlangan javob berish.
Javoblaringizda Telegram HTML formatidan (<b>, <i>, <code>, <pre>) unumli foydalaning.
O'zbek tilida ravon gapiring. Zarur bo'lsa boshqa tillarda ham javob bera olasiz.
EOT;

    public const DOCUMENT = <<<EOT
Siz hujjatlarni tahlil qilish bo'yicha professional ekspertsiz.
Berilgan hujjat matnini sinchkovlik bilan o'rganib, aniq faktlarga asoslangan xulosalar bering.
EOT;

    public const CODING = <<<EOT
Siz tajribali Senior Software Engineer siz.
Toza, xavfsiz va samarali kod yozasiz. Har bir qismni tushuntirib berasiz.
Format: Kodni ```til ko'rinishida bering.
EOT;

    public const TOOLS = <<<EOT
You are an advanced text and language processing specialist.
Provide high quality, well structured and concise output.
Adapt to user's language.
EOT;
}


// ============================================================================
// 1. SOZLAMALAR
// ============================================================================

$dotenv = Dotenv::createImmutable(__DIR__);
if (file_exists(__DIR__ . '/.env')) $dotenv->load();

function getEnvVal(string $key, mixed $default = null): mixed {
    $val = $_ENV[$key] ?? $_SERVER[$key] ?? getenv($key);
    return ($val !== false && $val !== null && $val !== '') ? $val : $default;
}

class Settings {
    public string $BOT_TOKEN;
    public int $ADMIN_ID;
    public string $GEMINI_API_KEY;
    public string $GEMINI_MODEL;
    public float $GEMINI_TEMPERATURE;
    public string $DATABASE_PATH;
    public int $FREE_DAILY_LIMIT;
    public int $MAX_FILE_SIZE_MB;
    public string $GOOGLE_SHEETS_WEBHOOK_URL;
    public string $GOOGLE_SHEET_URL;
    public string $TEMP_DIR, $UPLOADS_DIR, $DOWNLOADS_DIR, $DATA_DIR, $LOGS_DIR, $LOG_LEVEL;

    public function __construct() {
        $this->BOT_TOKEN            = (string)getEnvVal('BOT_TOKEN', '');
        $this->ADMIN_ID             = (int)getEnvVal('ADMIN_ID', 0);
        $this->GEMINI_API_KEY       = (string)getEnvVal('GEMINI_API_KEY', (string)getEnvVal('OPENAI_API_KEY', ''));
        $this->GEMINI_MODEL         = (string)getEnvVal('GEMINI_MODEL', 'gemini-3.6-flash');
        $this->GEMINI_TEMPERATURE   = (float)getEnvVal('GEMINI_TEMPERATURE', (float)getEnvVal('OPENAI_TEMPERATURE', 0.7));
        $this->GOOGLE_SHEETS_WEBHOOK_URL = (string)getEnvVal('GOOGLE_SHEETS_WEBHOOK_URL', '');
        $this->GOOGLE_SHEET_URL           = (string)getEnvVal('GOOGLE_SHEET_URL', '');
        $this->DATABASE_PATH        = (string)getEnvVal('DATABASE_PATH', './data/bot.db');
        $this->FREE_DAILY_LIMIT     = (int)getEnvVal('FREE_DAILY_LIMIT', 30);
        $this->MAX_FILE_SIZE_MB     = (int)getEnvVal('MAX_FILE_SIZE_MB', 20);
        $this->TEMP_DIR             = (string)getEnvVal('TEMP_DIR', './temp');
        $this->UPLOADS_DIR          = (string)getEnvVal('UPLOADS_DIR', './uploads');
        $this->DOWNLOADS_DIR        = (string)getEnvVal('DOWNLOADS_DIR', './downloads');
        $this->DATA_DIR             = (string)getEnvVal('DATA_DIR', './data');
        $this->LOGS_DIR             = (string)getEnvVal('LOGS_DIR', './logs');
        $this->LOG_LEVEL            = (string)getEnvVal('LOG_LEVEL', 'INFO');
        foreach ([$this->TEMP_DIR, $this->UPLOADS_DIR, $this->DOWNLOADS_DIR, $this->DATA_DIR, $this->LOGS_DIR] as $d)
            if (!is_dir($d)) mkdir($d, 0777, true);
    }

    public function maxFileSizeBytes(): int { return $this->MAX_FILE_SIZE_MB * 1024 * 1024; }
    public function isAdmin(int $uid): bool { return $this->ADMIN_ID !== 0 && $uid === $this->ADMIN_ID; }
}

$settings = new Settings();

// ============================================================================
// 2. LOGGER
// ============================================================================

class Logger {
    private string $file;
    private string $level;
    private static array $order = ['DEBUG'=>0,'INFO'=>1,'WARNING'=>2,'ERROR'=>3];

    public function __construct(string $file, string $level = 'INFO') {
        $this->file = $file;
        $this->level = strtoupper($level);
    }
    public function log(string $level, string $msg): void {
        if ((self::$order[$level] ?? 1) < (self::$order[$this->level] ?? 1)) return;
        $line = sprintf("[%s] %s | %s\n", date('Y-m-d H:i:s'), str_pad($level, 8), $msg);
        file_put_contents($this->file, $line, FILE_APPEND | LOCK_EX);
        echo $line;
    }
    public function info(string $m): void    { $this->log('INFO', $m); }
    public function debug(string $m): void   { $this->log('DEBUG', $m); }
    public function warning(string $m): void { $this->log('WARNING', $m); }
    public function error(string $m): void   { $this->log('ERROR', $m); }
}

$logger = new Logger($settings->LOGS_DIR . '/bot.log', $settings->LOG_LEVEL);

// ============================================================================
// 3. HELPERLAR
// ============================================================================

class Helpers {
    public const ALLOWED_DOC_EXT = ['.pdf', '.docx', '.txt', '.csv', '.xlsx', '.pptx'];
    public const ALLOWED_IMG_EXT = ['.jpg', '.jpeg', '.png', '.webp'];

    public static function sanitizeFilename(string $f): string {
        $b = basename($f);
        $c = preg_replace('/[^\w\s\.-]/u', '', $b);
        $c = trim($c);
        return $c !== '' ? $c : 'uploaded_file';
    }
    public static function validateExt(string $f, array $allowed): bool {
        return in_array('.' . strtolower(pathinfo($f, PATHINFO_EXTENSION)), $allowed, true);
    }
    public static function formatSize(int $b): string {
        if ($b < 1024) return "$b B";
        if ($b < 1048576) return round($b / 1024, 1) . " KB";
        if ($b < 1073741824) return round($b / 1048576, 2) . " MB";
        return round($b / 1073741824, 2) . " GB";
    }
    public static function splitText(string $text, int $max = 4000): array {
        if ($text === '') return [''];
        if (mb_strlen($text) <= $max) return [$text];
        $chunks = []; $cur = []; $len = 0; $inCode = false; $lang = '';
        foreach (explode("\n", $text) as $line) {
            $ll = mb_strlen($line) + 1;
            $st = trim($line);
            if (str_starts_with($st, '```')) {
                if (!$inCode) { $inCode = true; $lang = trim(substr($st, 3)); }
                else { $inCode = false; $lang = ''; }
            }
            if ($len + $ll > $max) {
                if ($inCode) $cur[] = '```';
                $chunks[] = implode("\n", $cur);
                $cur = []; $len = 0;
                if ($inCode) { $cur[] = '```' . $lang; $len += mb_strlen($lang) + 4; }
            }
            $cur[] = $line; $len += $ll;
        }
        if ($cur) $chunks[] = implode("\n", $cur);
        return $chunks;
    }
    public static function now(): string { return date('Y-m-d H:i:s'); }
    public static function uuid(): string {
        return bin2hex(random_bytes(8));
    }
}

// ============================================================================
// 4. MA'LUMOTLAR BAZASI (REPOSITORIES)
// ============================================================================

class Database {
    private static ?PDO $pdo = null;
    public static function pdo(string $path): PDO {
        if (self::$pdo === null) {
            $dir = dirname($path);
            if (!is_dir($dir)) mkdir($dir, 0777, true);
            self::$pdo = new PDO('sqlite:' . $path);
            self::$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
            self::$pdo->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_ASSOC);
            self::$pdo->exec('PRAGMA foreign_keys = ON');
        }
        return self::$pdo;
    }
}

class UserRepository {
    public function __construct(private PDO $pdo, private Settings $settings) {}

    public function getOrCreate(int $tid, ?string $username, ?string $firstName, string $lang = 'uz'): array {
        $stmt = $this->pdo->prepare('SELECT * FROM users WHERE telegram_id = ?');
        $stmt->execute([$tid]);
        $user = $stmt->fetch();

        $isAdmin = $this->settings->isAdmin($tid) ? 1 : 0;

        if (!$user) {
            $ins = $this->pdo->prepare('INSERT INTO users (telegram_id, username, first_name, language, is_admin) VALUES (?,?,?,?,?)');
            $ins->execute([$tid, $username, $firstName, $lang, $isAdmin]);
            $stmt->execute([$tid]);
            return $stmt->fetch();
        }

        $upd = $this->pdo->prepare('UPDATE users SET last_active = ?, username = COALESCE(?, username), first_name = COALESCE(?, first_name), is_admin = ? WHERE telegram_id = ?');
        $upd->execute([Helpers::now(), $username, $firstName, $isAdmin, $tid]);
        $stmt->execute([$tid]);
        return $stmt->fetch();
    }

    public function updateLanguage(int $tid, string $lang): void {
        $this->pdo->prepare('UPDATE users SET language = ? WHERE telegram_id = ?')->execute([$lang, $tid]);
    }
    public function getAllActiveIds(): array {
        return $this->pdo->query('SELECT telegram_id FROM users WHERE is_blocked = 0')->fetchAll(PDO::FETCH_COLUMN);
    }
    public function totalCount(): int {
        return (int)$this->pdo->query('SELECT COUNT(*) FROM users')->fetchColumn();
    }
    public function activeCount(int $days = 7): int {
        $since = date('Y-m-d H:i:s', time() - $days * 86400);
        $s = $this->pdo->prepare('SELECT COUNT(*) FROM users WHERE last_active >= ?');
        $s->execute([$since]);
        return (int)$s->fetchColumn();
    }
    public function blockedCount(): int {
        return (int)$this->pdo->query('SELECT COUNT(*) FROM users WHERE is_blocked = 1')->fetchColumn();
    }
    public function getRecentUsers(int $limit = 15): array {
        $s = $this->pdo->prepare('SELECT * FROM users ORDER BY last_active DESC LIMIT ?');
        $s->bindValue(1, $limit, PDO::PARAM_INT);
        $s->execute();
        return $s->fetchAll();
    }
    public function getByTelegramId(int $tid): ?array {
        $s = $this->pdo->prepare('SELECT * FROM users WHERE telegram_id = ?');
        $s->execute([$tid]);
        return $s->fetch() ?: null;
    }
    public function setBlocked(int $tid, bool $blocked): bool {
        $s = $this->pdo->prepare('UPDATE users SET is_blocked = ? WHERE telegram_id = ?');
        return $s->execute([$blocked ? 1 : 0, $tid]);
    }
    public function setLimitOverride(int $tid, ?int $limit): bool {
        $s = $this->pdo->prepare('UPDATE users SET daily_limit_override = ? WHERE telegram_id = ?');
        return $s->execute([$limit, $tid]);
    }
}

class ConversationRepository {
    public function __construct(private PDO $pdo) {}

    public function getOrCreateActive(int $tid): array {
        $s = $this->pdo->prepare('SELECT * FROM conversations WHERE telegram_user_id = ? AND is_active = 1 ORDER BY updated_at DESC LIMIT 1');
        $s->execute([$tid]);
        $c = $s->fetch();
        if ($c) return $c;

        $ins = $this->pdo->prepare('INSERT INTO conversations (telegram_user_id, title, is_active) VALUES (?,?,1)');
        $ins->execute([$tid, 'Yangi suhbat']);
        $id = (int)$this->pdo->lastInsertId();
        $s->execute([$tid]);
        return $s->fetch();
    }

    public function createNew(int $tid, string $title = 'Yangi suhbat'): array {
        $this->pdo->prepare('UPDATE conversations SET is_active = 0 WHERE telegram_user_id = ?')->execute([$tid]);
        $ins = $this->pdo->prepare('INSERT INTO conversations (telegram_user_id, title, is_active) VALUES (?,?,1)');
        $ins->execute([$tid, $title]);
        $s = $this->pdo->prepare('SELECT * FROM conversations WHERE id = ?');
        $s->execute([(int)$this->pdo->lastInsertId()]);
        return $s->fetch();
    }

    public function getUserConversations(int $tid, int $limit = 10): array {
        $s = $this->pdo->prepare('SELECT * FROM conversations WHERE telegram_user_id = ? ORDER BY updated_at DESC LIMIT ?');
        $s->bindValue(1, $tid, PDO::PARAM_INT);
        $s->bindValue(2, $limit, PDO::PARAM_INT);
        $s->execute();
        return $s->fetchAll();
    }

    public function setActive(int $tid, int $cid): ?array {
        $this->pdo->prepare('UPDATE conversations SET is_active = 0 WHERE telegram_user_id = ?')->execute([$tid]);
        $u = $this->pdo->prepare('UPDATE conversations SET is_active = 1, updated_at = ? WHERE id = ? AND telegram_user_id = ?');
        $u->execute([Helpers::now(), $cid, $tid]);
        $s = $this->pdo->prepare('SELECT * FROM conversations WHERE id = ?');
        $s->execute([$cid]);
        return $s->fetch() ?: null;
    }

    public function updateTitleIfDefault(int $cid, string $title): void {
        $s = $this->pdo->prepare('SELECT title FROM conversations WHERE id = ?');
        $s->execute([$cid]);
        $t = $s->fetchColumn();
        if ($t === 'Yangi suhbat') {
            $this->pdo->prepare('UPDATE conversations SET title = ?, updated_at = ? WHERE id = ?')
                ->execute([mb_substr($title, 0, 40), Helpers::now(), $cid]);
        }
    }

    public function touch(int $cid): void {
        $this->pdo->prepare('UPDATE conversations SET updated_at = ? WHERE id = ?')->execute([Helpers::now(), $cid]);
    }
}

class MessageRepository {
    public function __construct(private PDO $pdo, private ConversationRepository $convRepo) {}

    public function add(int $cid, string $role, string $content): void {
        $this->pdo->prepare('INSERT INTO messages (conversation_id, role, content) VALUES (?,?,?)')
            ->execute([$cid, $role, $content]);
        $this->convRepo->touch($cid);
    }

    public function recent(int $cid, int $limit = 10): array {
        $s = $this->pdo->prepare('SELECT role, content FROM messages WHERE conversation_id = ? ORDER BY id DESC LIMIT ?');
        $s->bindValue(1, $cid, PDO::PARAM_INT);
        $s->bindValue(2, $limit, PDO::PARAM_INT);
        $s->execute();
        return array_reverse($s->fetchAll());
    }

    public function getAllForConversation(int $cid): array {
        $s = $this->pdo->prepare('SELECT role, content, created_at FROM messages WHERE conversation_id = ? ORDER BY id ASC');
        $s->execute([$cid]);
        return $s->fetchAll();
    }

    public function getAllForUser(int $uid, int $limit = 200): array {
        $s = $this->pdo->prepare('
            SELECT m.role, m.content, m.created_at, c.title as conversation_title
            FROM messages m
            JOIN conversations c ON m.conversation_id = c.id
            WHERE c.telegram_user_id = ?
            ORDER BY m.id ASC
            LIMIT ?
        ');
        $s->bindValue(1, $uid, PDO::PARAM_INT);
        $s->bindValue(2, $limit, PDO::PARAM_INT);
        $s->execute();
        return $s->fetchAll();
    }
}

class UsageRepository {
    public function __construct(private PDO $pdo, private Settings $settings) {}

    public function record(int $tid, string $type, int $tokens = 0): void {
        $this->pdo->prepare('INSERT INTO usages (telegram_user_id, request_type, tokens) VALUES (?,?,?)')
            ->execute([$tid, $type, $tokens]);
    }

    public function todayCount(int $tid): int {
        $start = date('Y-m-d 00:00:00');
        $s = $this->pdo->prepare('SELECT COUNT(*) FROM usages WHERE telegram_user_id = ? AND created_at >= ?');
        $s->execute([$tid, $start]);
        return (int)$s->fetchColumn();
    }

    public function canRequest(int $tid): array {
        if ($this->settings->isAdmin($tid)) return [true, 0, PHP_INT_MAX];

        $s = $this->pdo->prepare('SELECT daily_limit_override FROM users WHERE telegram_id = ?');
        $s->execute([$tid]);
        $ovr = $s->fetchColumn();
        $limit = ($ovr !== false && $ovr !== null) ? (int)$ovr : $this->settings->FREE_DAILY_LIMIT;

        $used = $this->todayCount($tid);
        return [$used < $limit, $used, $limit];
    }

    public function statsSummary(): array {
        $todayStart = date('Y-m-d 00:00:00');
        $monthStart = date('Y-m-01 00:00:00');

        $todayQ = $this->pdo->prepare('SELECT COUNT(*) FROM usages WHERE created_at >= ?');
        $todayQ->execute([$todayStart]);
        $monthQ = $this->pdo->prepare('SELECT COUNT(*) FROM usages WHERE created_at >= ?');
        $monthQ->execute([$monthStart]);

        $br = $this->pdo->query('SELECT request_type, COUNT(*) as c FROM usages GROUP BY request_type')->fetchAll();
        $breakdown = [];
        foreach ($br as $row) $breakdown[$row['request_type']] = (int)$row['c'];

        $tokens = (int)$this->pdo->query('SELECT COALESCE(SUM(tokens),0) FROM usages')->fetchColumn();

        return [
            'today_requests' => (int)$todayQ->fetchColumn(),
            'month_requests' => (int)$monthQ->fetchColumn(),
            'breakdown'      => $breakdown,
            'total_tokens'   => $tokens,
        ];
    }
}

class FileRepository {
    public function __construct(private PDO $pdo) {}
    public function record(int $tid, string $name, string $type, int $size): void {
        $this->pdo->prepare('INSERT INTO files (telegram_user_id, filename, file_type, file_size) VALUES (?,?,?,?)')
            ->execute([$tid, $name, $type, $size]);
    }
}

class ActivityLogRepository {
    public function __construct(private PDO $pdo) {}

    public function record(
        int $userId,
        ?string $username,
        ?string $fullName,
        string $action,
        string $query = '',
        string $response = '',
        int $tokens = 0,
        string $status = 'OK'
    ): int {
        try {
            $stmt = $this->pdo->prepare('
                INSERT INTO activity_logs (telegram_user_id, username, full_name, action, user_query, ai_response, tokens, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ');
            $stmt->execute([
                $userId,
                $username,
                $fullName,
                $action,
                mb_substr($query, 0, 2000),
                mb_substr($response, 0, 2000),
                $tokens,
                $status
            ]);
            return (int)$this->pdo->lastInsertId();
        } catch (Throwable $e) {
            return 0;
        }
    }

    public function getRecent(int $limit = 50): array {
        try {
            $stmt = $this->pdo->prepare('
                SELECT * FROM activity_logs ORDER BY id DESC LIMIT ?
            ');
            $stmt->bindValue(1, $limit, PDO::PARAM_INT);
            $stmt->execute();
            return array_reverse($stmt->fetchAll());
        } catch (Throwable $e) {
            return [];
        }
    }

    public function countTotal(): int {
        try {
            return (int)$this->pdo->query('SELECT COUNT(*) FROM activity_logs')->fetchColumn();
        } catch (Throwable $e) {
            return 0;
        }
    }

    public function getUserActivities(int $uid, int $limit = 100): array {
        try {
            $stmt = $this->pdo->prepare('
                SELECT * FROM activity_logs WHERE telegram_user_id = ? ORDER BY id DESC LIMIT ?
            ');
            $stmt->bindValue(1, $uid, PDO::PARAM_INT);
            $stmt->bindValue(2, $limit, PDO::PARAM_INT);
            $stmt->execute();
            return array_reverse($stmt->fetchAll());
        } catch (Throwable) {
            return [];
        }
    }

    public function getAllForAudit(int $limit = 1000, bool $todayOnly = false): array {
        try {
            if ($todayOnly) {
                $todayStart = date('Y-m-d 00:00:00');
                $stmt = $this->pdo->prepare('
                    SELECT * FROM activity_logs WHERE created_at >= ? ORDER BY id DESC LIMIT ?
                ');
                $stmt->bindValue(1, $todayStart);
                $stmt->bindValue(2, $limit, PDO::PARAM_INT);
            } else {
                $stmt = $this->pdo->prepare('
                    SELECT * FROM activity_logs ORDER BY id DESC LIMIT ?
                ');
                $stmt->bindValue(1, $limit, PDO::PARAM_INT);
            }
            $stmt->execute();
            return array_reverse($stmt->fetchAll());
        } catch (Throwable) {
            return [];
        }
    }
}

class SettingsRepository {
    public function __construct(private PDO $pdo) {
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS bot_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
        ");
    }

    public function get(string $key, ?string $default = null): ?string {
        try {
            $stmt = $this->pdo->prepare("SELECT value FROM bot_settings WHERE key = ?");
            $stmt->execute([$key]);
            $val = $stmt->fetchColumn();
            return $val !== false ? (string)$val : $default;
        } catch (Throwable) {
            return $default;
        }
    }

    public function set(string $key, string $value): void {
        try {
            $stmt = $this->pdo->prepare("
                INSERT INTO bot_settings (key, value, updated_at) 
                VALUES (?, ?, datetime('now'))
                ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')
            ");
            $stmt->execute([$key, $value]);
        } catch (Throwable) {}
    }
}

// ============================================================================
// 5. SERVISLAR
// ============================================================================

class TelegramAPI {
    private Client $http;
    public function __construct(private string $token) {
        $this->http = new Client(['base_uri' => "https://api.telegram.org/bot{$token}/", 'timeout' => 60]);
    }
    public function call(string $method, array $params = []): array {
        try {
            $r = $this->http->post($method, ['json' => $params]);
            $data = json_decode((string)$r->getBody(), true);
            return $data ?? ['ok' => false];
        } catch (GuzzleException $e) {
            return ['ok' => false, 'description' => $e->getMessage()];
        }
    }
    public function sendMessage(int $chatId, string $text, ?array $replyMarkup = null, string $parseMode = 'HTML'): array {
        return $this->call('sendMessage', array_filter([
            'chat_id' => $chatId, 'text' => $text,
            'parse_mode' => $parseMode,
            'reply_markup' => $replyMarkup ? json_encode($replyMarkup, JSON_UNESCAPED_UNICODE) : null,
        ], fn($v) => $v !== null));
    }
    public function editMessageText(int $chatId, int $msgId, string $text, ?array $replyMarkup = null): array {
        return $this->call('editMessageText', array_filter([
            'chat_id' => $chatId, 'message_id' => $msgId, 'text' => $text,
            'parse_mode' => 'HTML',
            'reply_markup' => $replyMarkup ? json_encode($replyMarkup, JSON_UNESCAPED_UNICODE) : null,
        ], fn($v) => $v !== null));
    }
    public function answerCallback(string $id, string $text = '', bool $alert = false): void {
        $this->call('answerCallbackQuery', ['callback_query_id' => $id, 'text' => $text, 'show_alert' => $alert]);
    }
    public function sendChatAction(int $chatId, string $action): void {
        $this->call('sendChatAction', ['chat_id' => $chatId, 'action' => $action]);
    }
    public function deleteMessage(int $chatId, int $msgId): void {
        $this->call('deleteMessage', ['chat_id' => $chatId, 'message_id' => $msgId]);
    }
    public function getFile(string $fileId): ?string {
        $r = $this->call('getFile', ['file_id' => $fileId]);
        return $r['ok'] ? $r['result']['file_path'] : null;
    }
    public function downloadFile(string $filePath, string $savePath): bool {
        try {
            $url = "https://api.telegram.org/file/bot{$this->token}/{$filePath}";
            $this->http->get($url, ['sink' => $savePath]);
            return file_exists($savePath);
        } catch (GuzzleException $e) {
            return false;
        }
    }
    public function sendDocument(int $chatId, string $filePath, ?string $caption = null, ?array $replyMarkup = null): array {
        try {
            if (!file_exists($filePath)) {
                return ['ok' => false, 'description' => 'Fayl topilmadi: ' . $filePath];
            }
            $multipart = [
                ['name' => 'chat_id', 'contents' => (string)$chatId],
                [
                    'name'     => 'document',
                    'contents' => fopen($filePath, 'r'),
                    'filename' => basename($filePath),
                ],
            ];
            if ($caption !== null) {
                $multipart[] = ['name' => 'caption', 'contents' => $caption];
                $multipart[] = ['name' => 'parse_mode', 'contents' => 'HTML'];
            }
            if ($replyMarkup !== null) {
                $multipart[] = ['name' => 'reply_markup', 'contents' => json_encode($replyMarkup, JSON_UNESCAPED_UNICODE)];
            }
            $r = $this->http->post('sendDocument', ['multipart' => $multipart]);
            $data = json_decode((string)$r->getBody(), true);
            return $data ?? ['ok' => false];
        } catch (GuzzleException $e) {
            return ['ok' => false, 'description' => $e->getMessage()];
        }
    }
}

class GeminiService {
    private Client $http;

    public function __construct(private Settings $settings) {
        $this->http = new Client([
            'base_uri' => 'https://generativelanguage.googleapis.com/v1beta/',
            'timeout'  => 120,
            'headers'  => [
                'x-goog-api-key' => $settings->GEMINI_API_KEY,
                'Content-Type'   => 'application/json',
            ],
        ]);
    }

    private function post(string $path, array $body): array {
        try {
            $r = $this->http->post($path, ['json' => $body]);
            return json_decode((string)$r->getBody(), true) ?? [];
        } catch (GuzzleException $e) {
            $errMsg = $e->getMessage();
            if (method_exists($e, 'getResponse') && $e->getResponse()) {
                $errBody = (string)$e->getResponse()->getBody();
                $errData = json_decode($errBody, true);
                if (isset($errData['error']['message'])) {
                    $errMsg = $errData['error']['message'];
                }
            }
            throw new RuntimeException('Gemini xatosi: ' . $errMsg);
        }
    }

    private function extractTextAndTokens(array $data): array {
        $content = '';
        if (isset($data['candidates'][0]['content']['parts'])) {
            foreach ($data['candidates'][0]['content']['parts'] as $part) {
                if (isset($part['text'])) {
                    $content .= $part['text'];
                }
            }
        }
        $tokens = $data['usageMetadata']['totalTokenCount'] ?? 0;
        return [
            'content' => $content,
            'tokens'  => $tokens,
        ];
    }

    public function chat(array $messages, ?string $systemPrompt = null, ?float $temp = null): array {
        $contents = [];
        foreach ($messages as $m) {
            $role = ($m['role'] === 'assistant' || $m['role'] === 'model') ? 'model' : 'user';
            $text = trim((string)($m['content'] ?? ''));
            if ($text === '') continue;

            $count = count($contents);
            if ($count > 0 && $contents[$count - 1]['role'] === $role) {
                $contents[$count - 1]['parts'][0]['text'] .= "\n" . $text;
            } else {
                $contents[] = [
                    'role'  => $role,
                    'parts' => [['text' => $text]],
                ];
            }
        }

        if (!empty($contents) && $contents[0]['role'] === 'model') {
            array_unshift($contents, [
                'role'  => 'user',
                'parts' => [['text' => 'Salom']],
            ]);
        }

        $body = [
            'contents' => $contents,
            'generationConfig' => [
                'temperature' => $temp ?? $this->settings->GEMINI_TEMPERATURE,
            ],
        ];

        $sys = $systemPrompt ?? Prompts::SYSTEM;
        if (!empty($sys)) {
            $body['systemInstruction'] = [
                'parts' => [['text' => $sys]],
            ];
        }

        $model = !empty($this->settings->GEMINI_MODEL) ? $this->settings->GEMINI_MODEL : 'gemini-3.6-flash';
        $endpoint = 'models/' . urlencode($model) . ':generateContent';
        $data = $this->post($endpoint, $body);
        return $this->extractTextAndTokens($data);
    }

    public function transcribe(string $audioPath): string {
        if (!file_exists($audioPath) || filesize($audioPath) === 0) {
            throw new RuntimeException("Ovozli fayl topilmadi yoki bo'sh.");
        }

        $ext = strtolower(pathinfo($audioPath, PATHINFO_EXTENSION));
        $mimeMap = [
            'ogg'  => 'audio/ogg',
            'oga'  => 'audio/ogg',
            'mp3'  => 'audio/mp3',
            'wav'  => 'audio/wav',
            'm4a'  => 'audio/m4a',
            'aac'  => 'audio/aac',
            'flac' => 'audio/flac',
        ];
        $mimeType = $mimeMap[$ext] ?? 'audio/ogg';

        $b64 = base64_encode(file_get_contents($audioPath));

        $body = [
            'contents' => [
                [
                    'role'  => 'user',
                    'parts' => [
                        [
                            'inlineData' => [
                                'mimeType' => $mimeType,
                                'data'     => $b64,
                            ],
                        ],
                        [
                            'text' => "Iltimos, ushbu audio yozuvdagi barcha so'zlarni eshitganingizdek aniq matnga aylantiring (transkripsiya qiling). Faqat audio ichidagi aytilgan gaplarni yozing, hech qanday kirish so'zlari, izoh yoki tarjima qo'shmang.",
                        ],
                    ],
                ],
            ],
            'generationConfig' => [
                'temperature' => 0.1,
            ],
        ];

        $model = !empty($this->settings->GEMINI_MODEL) ? $this->settings->GEMINI_MODEL : 'gemini-3.6-flash';
        $endpoint = 'models/' . urlencode($model) . ':generateContent';
        $data = $this->post($endpoint, $body);
        $res = $this->extractTextAndTokens($data);
        return trim($res['content']);
    }

    public function analyzeImage(string $base64, string $prompt, string $mimeType = 'image/jpeg'): array {
        $body = [
            'contents' => [
                [
                    'role'  => 'user',
                    'parts' => [
                        [
                            'inlineData' => [
                                'mimeType' => $mimeType,
                                'data'     => $base64,
                            ],
                        ],
                        [
                            'text' => $prompt,
                        ],
                    ],
                ],
            ],
            'systemInstruction' => [
                'parts' => [['text' => Prompts::SYSTEM]],
            ],
            'generationConfig' => [
                'temperature' => $this->settings->GEMINI_TEMPERATURE,
            ],
        ];

        $model = !empty($this->settings->GEMINI_MODEL) ? $this->settings->GEMINI_MODEL : 'gemini-3.6-flash';
        $endpoint = 'models/' . urlencode($model) . ':generateContent';
        $data = $this->post($endpoint, $body);
        return $this->extractTextAndTokens($data);
    }

    public function analyzeDocument(string $docText, string $action, ?string $question = null): array {
        $prompts = [
            'summary'         => "Quyidagi hujjatning asosiy mazmunini qisqa, aniq va lo'nda qilib tushuntiring (Executive Summary).",
            'deep_analysis'   => "Quyidagi hujjatni chuqur tahlil qiling: asosiy g'oyalar, bo'limlar, maqsad va natijalarni tizimli yoriting.",
            'issues'          => "Ushbu hujjatdagi mavjud xatolar, risklar, qarama-qarshiliklar va kamchiliklarni aniqlang.",
            'recommendations' => "Ushbu hujjat yuzasidan amaliy tavsiyalar, optimallashtirish yo'llari va keyingi qadamlarni bering.",
            'data'            => "Hujjatdagi barcha muhim raqamlar, statistikalar, sanalar va asosiy ko'rsatkichlarni jadval shaklida ajratib bering.",
            'question'        => "Foydalanuvchi savoli: {$question}\n\nUshbu savolga faqat quyidagi hujjat ma'lumotlariga tayangan holda aniq javob bering.",
        ];
        $instr = $prompts[$action] ?? $prompts['summary'];
        $prompt = "{$instr}\n\n--- HUJJAT MATNI ---\n{$docText}\n--- HUJJAT MATNI TUGADI ---";
        return $this->chat([['role' => 'user', 'content' => $prompt]], Prompts::DOCUMENT);
    }

    public function coding(string $query, string $mode = 'general'): array {
        $modes = [
            'write'    => "Vazifa: Talab qilingan kodni to'liq, ishlab turgan, toza va arxitekturaviy to'g'ri yozib bering.",
            'bug'      => "Vazifa: Ushbu kodni tekshiring, xatolar (bugs) va risklarni aniqlang.",
            'fix'      => "Vazifa: Kodning xatolarini tuzatib, to'liq ishlaydigan variantini bering.",
            'optimize' => "Vazifa: Ushbu kodni tezlik va xotira bo'yicha optimallashtiring.",
            'explain'  => "Vazifa: Ushbu kodning ishlashini bosqichma-bosqich tushuntiring.",
            'refactor' => "Vazifa: Ushbu kodni Clean Code tamoyillariga muvofiq qayta yozing.",
            'test'     => "Vazifa: Ushbu kod uchun pytest unit testlar yozib bering.",
            'general'  => "Vazifa: Dasturlash so'rovi bo'yicha toza yechim bering.",
        ];
        $instr = $modes[$mode] ?? $modes['general'];
        return $this->chat([['role' => 'user', 'content' => "[{$instr}]\n\nFoydalanuvchi so'rovi:\n{$query}"]], Prompts::CODING);
    }

    public function runTool(string $tool, string $input, array $params = []): array {
        $directives = [
            'text_gen'      => "Text Generator: '{$input}' mavzusida professional matn yozing.",
            'translator'    => "Translator: Quyidagi matnni " . ($params['target_lang'] ?? "o'zbek") . " tiliga professional tarjima qiling:\n\n{$input}",
            'summarizer'    => "Summarizer: Quyidagi matnni eng muhim nuqtalarini saqlab xulosa qiling:\n\n{$input}",
            'idea_gen'      => "Idea Generator: '{$input}' bo'yicha 5 ta eng innovatsion va amaliy g'oya bering.",
            'rewrite'       => "Rewrite: Quyidagi matnni " . ($params['style'] ?? 'professional') . " uslubda qayta yozing:\n\n{$input}",
            'marketing'     => "Marketing: '{$input}' uchun reklama matni, Telegram post va CTA yozing.",
            'study'         => "Study Assistant: '{$input}' mavzusini o'rganish bo'yicha tushuntirish va test savollari tuzing.",
            'data_analysis' => "Data Analysis: Quyidagi ma'lumotlarni tahlil qiling va qonuniyatlarni bering:\n\n{$input}",
        ];
        $directive = $directives[$tool] ?? "Quyidagi vazifani bajaring:\n{$input}";
        return $this->chat([['role' => 'user', 'content' => $directive]], Prompts::TOOLS);
    }
}

class DocumentService {
    private const MAX_LEN = 30000;

    public function extractText(string $path): string {
        $ext = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        switch ($ext) {
            case 'pdf':
                $parser = new PdfParser();
                $pdf = $parser->parseFile($path);
                $pages = [];
                $i = 1;
                foreach ($pdf->getPages() as $page) {
                    $txt = trim($page->getText());
                    if ($txt !== '') $pages[] = "[Sahifa {$i}]\n{$txt}";
                    $i++;
                }
                return $this->truncate(implode("\n\n", $pages));

            case 'docx':
                $phpWord = WordIOFactory::load($path);
                $parts = [];
                foreach ($phpWord->getSections() as $section) {
                    foreach ($section->getElements() as $el) {
                        if (method_exists($el, 'getText')) {
                            $t = trim($el->getText());
                            if ($t !== '') $parts[] = $t;
                        } elseif (method_exists($el, 'getRows')) {
                            foreach ($el->getRows() as $row) {
                                $cells = [];
                                foreach ($row->getCells() as $cell) {
                                    $cells[] = trim($cell->getText());
                                }
                                $parts[] = implode(' | ', $cells);
                            }
                        }
                    }
                }
                return $this->truncate(implode("\n\n", $parts));

            case 'txt':
            case 'log':
            case 'md':
                $content = @file_get_contents($path);
                if ($content === false) throw new RuntimeException('Faylni o\'qib bo\'lmadi.');
                $content = mb_convert_encoding($content, 'UTF-8', 'UTF-8, Windows-1251, ISO-8859-1');
                return $this->truncate($content);

            case 'csv':
                $lines = [];
                if (($h = fopen($path, 'r')) !== false) {
                    $i = 0;
                    while (($row = fgetcsv($h)) !== false) {
                        if ($i > 200) { $lines[] = '... [qolgan qatorlar qisqartirildi]'; break; }
                        $lines[] = implode(' | ', $row);
                        $i++;
                    }
                    fclose($h);
                }
                return $this->truncate(implode("\n", $lines));

            case 'xlsx':
                $spreadsheet = SpreadsheetIOFactory::load($path);
                $parts = [];
                $sheetCount = 0;
                foreach ($spreadsheet->getWorksheetIterator() as $sheet) {
                    if ($sheetCount >= 5) break;
                    $parts[] = "--- {$sheet->getTitle()} ---";
                    $r = 0;
                    foreach ($sheet->getRowIterator() as $row) {
                        if ($r > 100) break;
                        $cells = [];
                        foreach ($row->getCellIterator() as $cell) {
                            $cells[] = (string)$cell->getValue();
                        }
                        $parts[] = implode(' | ', $cells);
                        $r++;
                    }
                    $sheetCount++;
                }
                return $this->truncate(implode("\n", $parts));

            case 'pptx':
                $presentation = PresentationIOFactory::load($path);
                $slides = [];
                $i = 1;
                foreach ($presentation->getAllSlides() as $slide) {
                    $txts = [];
                    foreach ($slide->getShapeCollection() as $shape) {
                        if ($shape instanceof \PhpOffice\PhpPresentation\Shape\RichText) {
                            $t = trim($shape->getPlainText());
                            if ($t !== '') $txts[] = $t;
                        }
                    }
                    if ($txts) $slides[] = "[Slayd {$i}]\n" . implode("\n", $txts);
                    $i++;
                }
                return $this->truncate(implode("\n\n", $slides));
        }
        throw new InvalidArgumentException("Qo'llab-quvvatlanmaydigan format: {$ext}");
    }

    private function truncate(string $text): string {
        if (mb_strlen($text) <= self::MAX_LEN) return $text;
        $notice = "\n\n[...Hujjat hajmi katta bo'lgani sababli o'rta qismlar qisqartirildi...]\n\n";
        return mb_substr($text, 0, 20000) . $notice . mb_substr($text, -8000);
    }
}

class SpeechService {
    public function __construct(
        private TelegramAPI $tg,
        private GeminiService $gemini
    ) {}

    public function processVoice(string $fileId, string $ext = 'ogg'): array {
        $tempPath = rtrim((string)getEnvVal('TEMP_DIR', './temp'), '/') . '/voice_' . Helpers::uuid() . '.' . $ext;
        try {
            $filePath = $this->tg->getFile($fileId);
            if (!$filePath) return [null, "Ovozli faylni yuklab bo'lmadi."];
            if (!$this->tg->downloadFile($filePath, $tempPath))
                return [null, "Ovozli faylni saqlab bo'lmadi."];

            $text = $this->gemini->transcribe($tempPath);
            if ($text === '')
                return [null, "❌ Ovozli xabarni tushunib bo'lmadi. Iltimos, yana bir marta yuboring."];
            return [$text, null];
        } catch (Throwable $e) {
            return [null, "❌ Ovozni tahlil qilishda xatolik: " . $e->getMessage()];
        } finally {
            @unlink($tempPath);
        }
    }
}

class ImageService {
    public function __construct(
        private TelegramAPI $tg,
        private GeminiService $gemini
    ) {}

    public function processPhoto(string $fileId, string $prompt): array {
        $tempPath = rtrim((string)getEnvVal('TEMP_DIR', './temp'), '/') . '/photo_' . Helpers::uuid() . '.jpg';
        try {
            $filePath = $this->tg->getFile($fileId);
            if (!$filePath) throw new RuntimeException('Rasmni yuklab bo\'lmadi.');
            if (!$this->tg->downloadFile($filePath, $tempPath))
                throw new RuntimeException('Rasmni saqlab bo\'lmadi.');
            $b64 = base64_encode(file_get_contents($tempPath));
            return $this->gemini->analyzeImage($b64, $prompt, 'image/jpeg');
        } finally {
            @unlink($tempPath);
        }
    }
}

class GoogleSheetsService {
    private Client $http;

    public function __construct(
        private Settings $settings,
        private ?Logger $logger = null
    ) {
        $this->http = new Client([
            'timeout'         => 4.0,
            'connect_timeout' => 2.5,
            'http_errors'     => false,
        ]);
    }

    public function isConfigured(): bool {
        return !empty($this->settings->GOOGLE_SHEETS_WEBHOOK_URL) &&
               $this->settings->GOOGLE_SHEETS_WEBHOOK_URL !== 'YOUR_GOOGLE_SHEETS_WEBHOOK_URL_HERE';
    }

    public function log(
        int $userId,
        ?string $username,
        ?string $fullName,
        string $action,
        string $query = '',
        string $response = '',
        int $tokens = 0,
        string $status = 'OK'
    ): bool {
        if (!$this->isConfigured()) return false;

        $payload = [
            'timestamp' => Helpers::now(),
            'user_id'   => $userId,
            'username'  => $username ?? '',
            'full_name' => $fullName ?? '',
            'action'    => $action,
            'query'     => mb_substr($query, 0, 800),
            'response'  => mb_substr($response, 0, 800),
            'tokens'    => $tokens,
            'status'    => $status,
        ];

        try {
            $resp = $this->http->post($this->settings->GOOGLE_SHEETS_WEBHOOK_URL, [
                'json' => $payload,
            ]);
            return $resp->getStatusCode() === 200;
        } catch (Throwable $e) {
            if ($this->logger) {
                $this->logger->error("Google Sheets log xatosi: " . $e->getMessage());
            }
            return false;
        }
    }

    public function syncBatch(array $logs): int {
        if (!$this->isConfigured() || empty($logs)) return 0;

        $items = [];
        foreach ($logs as $l) {
            $items[] = [
                'timestamp' => $l['created_at'] ?? Helpers::now(),
                'user_id'   => $l['telegram_user_id'],
                'username'  => $l['username'] ?? '',
                'full_name' => $l['full_name'] ?? '',
                'action'    => $l['action'] ?? '',
                'query'     => mb_substr($l['user_query'] ?? '', 0, 800),
                'response'  => mb_substr($l['ai_response'] ?? '', 0, 800),
                'tokens'    => (int)($l['tokens'] ?? 0),
                'status'    => $l['status'] ?? 'OK',
            ];
        }

        try {
            $resp = $this->http->post($this->settings->GOOGLE_SHEETS_WEBHOOK_URL, [
                'json' => ['batch' => $items],
            ]);
            return $resp->getStatusCode() === 200 ? count($items) : 0;
        } catch (Throwable $e) {
            if ($this->logger) {
                $this->logger->error("Google Sheets batch sync xatosi: " . $e->getMessage());
            }
            return 0;
        }
    }
}

class ExportService {
    public function __construct(private string $tempDir) {
        if (!is_dir($this->tempDir)) @mkdir($this->tempDir, 0777, true);
    }

    private function renderPdf(string $html, string $outputPath, string $orientation = 'portrait'): bool {
        try {
            $options = new Options();
            $options->set('isHtml5ParserEnabled', true);
            $options->set('isRemoteEnabled', true);
            $options->set('defaultFont', 'DejaVu Sans');

            $dompdf = new Dompdf($options);
            $dompdf->loadHtml($html, 'UTF-8');
            $dompdf->setPaper('A4', $orientation);
            $dompdf->render();

            file_put_contents($outputPath, $dompdf->output());
            return file_exists($outputPath) && filesize($outputPath) > 0;
        } catch (Throwable $e) {
            return false;
        }
    }

    public function exportUserMessagesPdf(int $uid, string $userName, array $messages, string $title): ?string {
        $now = date('d.m.Y H:i:s');
        $total = count($messages);
        $safeTitle = htmlspecialchars($title, ENT_QUOTES, 'UTF-8');
        $safeName = htmlspecialchars($userName, ENT_QUOTES, 'UTF-8');

        $itemsHtml = '';
        foreach ($messages as $m) {
            $isUser = $m['role'] === 'user';
            $cls = $isUser ? 'msg-user' : 'msg-ai';
            $badge = $isUser ? '<span class="badge badge-user">👤 Foydalanuvchi</span>' : '<span class="badge badge-ai">🤖 AI Assistant</span>';
            $time = htmlspecialchars($m['created_at'] ?? $now, ENT_QUOTES, 'UTF-8');
            $body = nl2br(htmlspecialchars($m['content'], ENT_QUOTES, 'UTF-8'));

            $itemsHtml .= "
            <div class='msg-card {$cls}'>
                <div class='msg-meta'>{$badge} &nbsp; <span style='color:#64748b; font-weight:normal;'>{$time}</span></div>
                <div class='msg-body'>{$body}</div>
            </div>";
        }

        if ($itemsHtml === '') {
            $itemsHtml = "<p style='color:#64748b; text-align:center;'>Ushbu suhbatda hali xabarlar mavjud emas.</p>";
        }

        $html = "
        <!DOCTYPE html>
        <html lang='uz'>
        <head>
        <meta charset='UTF-8'>
        <style>
            @page { margin: 25px 25px 35px 25px; }
            body { font-family: 'DejaVu Sans', sans-serif; font-size: 11px; color: #1e293b; line-height: 1.5; margin: 0; padding: 0; }
            .header { background: #0f172a; color: #ffffff; padding: 18px 20px; border-radius: 8px; margin-bottom: 20px; }
            .header h1 { margin: 0 0 6px 0; font-size: 18px; color: #38bdf8; }
            .header .meta { font-size: 10px; color: #94a3b8; }
            .badge { display: inline-block; padding: 3px 8px; font-size: 9px; font-weight: bold; border-radius: 4px; }
            .badge-user { background: #e0f2fe; color: #0369a1; }
            .badge-ai { background: #dcfce7; color: #15803d; }
            .msg-card { border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px 14px; margin-bottom: 12px; }
            .msg-user { background: #f8fafc; border-left: 4px solid #0284c7; }
            .msg-ai { background: #ffffff; border-left: 4px solid #10b981; }
            .msg-meta { font-size: 10px; color: #64748b; margin-bottom: 6px; font-weight: bold; }
            .msg-body { font-size: 10.5px; word-wrap: break-word; }
            .footer { position: fixed; bottom: -20px; left: 0; right: 0; text-align: center; font-size: 9px; color: #94a3b8; }
        </style>
        </head>
        <body>
            <div class='header'>
                <h1>🤖 Universal AI Assistant — Suhbat Hisoboti</h1>
                <div class='meta'>
                    Suhbat mavzusi: <b>{$safeTitle}</b> &nbsp;|&nbsp; 
                    Foydalanuvchi: <b>{$safeName}</b> (ID: {$uid}) &nbsp;|&nbsp; 
                    Jami xabarlar: <b>{$total}</b> ta &nbsp;|&nbsp; 
                    Sana: {$now}
                </div>
            </div>
            {$itemsHtml}
            <div class='footer'>Universal AI Assistant Bot (@assistanduzz_bot) — Shaxsiy suhbat hisoboti</div>
        </body>
        </html>";

        $path = rtrim($this->tempDir, '/') . '/chat_' . $uid . '_' . time() . '.pdf';
        return $this->renderPdf($html, $path) ? $path : null;
    }

    public function exportUserMessagesMd(int $uid, string $userName, array $messages, string $title): ?string {
        $now = date('d.m.Y H:i:s');
        $md = "# 🤖 Universal AI Assistant — Suhbat Tarixi\n\n";
        $md .= "- **Suhbat mavzusi:** {$title}\n";
        $md .= "- **Foydalanuvchi:** {$userName} (ID: `{$uid}`)\n";
        $md .= "- **Eksport sanasi:** {$now}\n";
        $md .= "- **Xabarlar soni:** " . count($messages) . " ta\n\n";
        $md .= "---\n\n";

        foreach ($messages as $m) {
            $role = $m['role'] === 'user' ? '👤 Foydalanuvchi' : '🤖 AI Assistant';
            $time = $m['created_at'] ?? $now;
            $md .= "### {$role} ({$time})\n\n";
            $md .= trim($m['content']) . "\n\n";
            $md .= "---\n\n";
        }

        $path = rtrim($this->tempDir, '/') . '/chat_' . $uid . '_' . time() . '.md';
        file_put_contents($path, $md);
        return file_exists($path) ? $path : null;
    }

    public function exportUserActivitiesPdf(int $uid, string $userName, array $activities): ?string {
        $now = date('d.m.Y H:i:s');
        $safeName = htmlspecialchars($userName, ENT_QUOTES, 'UTF-8');
        $total = count($activities);

        $rows = '';
        $i = 1;
        foreach ($activities as $a) {
            $time = htmlspecialchars($a['created_at'] ?? '', ENT_QUOTES, 'UTF-8');
            $action = htmlspecialchars($a['action'] ?? '', ENT_QUOTES, 'UTF-8');
            $query = nl2br(htmlspecialchars(mb_substr($a['user_query'] ?? '', 0, 300), ENT_QUOTES, 'UTF-8'));
            $resp = nl2br(htmlspecialchars(mb_substr($a['ai_response'] ?? '', 0, 400), ENT_QUOTES, 'UTF-8'));
            $tok = number_format((int)($a['tokens'] ?? 0));
            $st = $a['status'] === 'ERROR' ? '<span class="badge badge-err">ERROR</span>' : '<span class="badge badge-ok">OK</span>';

            $rows .= "<tr>
                <td style='text-align:center;'>{$i}</td>
                <td>{$time}</td>
                <td><b>{$action}</b></td>
                <td>{$query}</td>
                <td>{$resp}</td>
                <td style='text-align:center;'>{$tok}</td>
                <td style='text-align:center;'>{$st}</td>
            </tr>";
            $i++;
        }

        $html = "
        <!DOCTYPE html>
        <html lang='uz'>
        <head>
        <meta charset='UTF-8'>
        <style>
            @page { margin: 20px; }
            body { font-family: 'DejaVu Sans', sans-serif; font-size: 9.5px; color: #1e293b; line-height: 1.4; }
            .header { background: #0f172a; color: #ffffff; padding: 14px 18px; border-radius: 6px; margin-bottom: 15px; }
            .header h1 { margin: 0 0 4px 0; font-size: 16px; color: #38bdf8; }
            .header .meta { font-size: 9.5px; color: #94a3b8; }
            .badge { display: inline-block; padding: 2px 6px; font-size: 8.5px; font-weight: bold; border-radius: 3px; }
            .badge-ok { background: #dcfce7; color: #166534; }
            .badge-err { background: #fee2e2; color: #991b1b; }
            table { width: 100%; border-collapse: collapse; }
            th { background: #1e293b; color: #ffffff; padding: 7px 5px; font-size: 9px; text-align: left; }
            td { padding: 6px 5px; border-bottom: 1px solid #e2e8f0; font-size: 8.5px; vertical-align: top; }
            tr:nth-child(even) td { background: #f8fafc; }
            .footer { position: fixed; bottom: -15px; left: 0; right: 0; text-align: center; font-size: 8px; color: #94a3b8; }
        </style>
        </head>
        <body>
            <div class='header'>
                <h1>📊 Foydalanuvchi Amallari Tarixi</h1>
                <div class='meta'>
                    Foydalanuvchi: <b>{$safeName}</b> (ID: {$uid}) &nbsp;|&nbsp; 
                    Jami amallar: <b>{$total}</b> ta &nbsp;|&nbsp; 
                    Sana: {$now}
                </div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style='width:25px; text-align:center;'>#</th>
                        <th style='width:90px;'>Vaqt</th>
                        <th style='width:90px;'>Bo'lim</th>
                        <th style='width:180px;'>So'rov</th>
                        <th>Javob xulosasi</th>
                        <th style='width:45px; text-align:center;'>Token</th>
                        <th style='width:40px; text-align:center;'>Status</th>
                    </tr>
                </thead>
                <tbody>{$rows}</tbody>
            </table>
            <div class='footer'>Universal AI Assistant Bot (@assistanduzz_bot) — Barcha huquqlar himoyalangan</div>
        </body>
        </html>";

        $path = rtrim($this->tempDir, '/') . '/activity_' . $uid . '_' . time() . '.pdf';
        return $this->renderPdf($html, $path, 'landscape') ? $path : null;
    }

    public function exportAdminAuditPdf(array $activities, string $title, bool $todayOnly = false): ?string {
        $now = date('d.m.Y H:i:s');
        $safeTitle = htmlspecialchars($title, ENT_QUOTES, 'UTF-8');
        $total = count($activities);

        $totalTokens = 0;
        $successCount = 0;
        $uniqueUsers = [];

        foreach ($activities as $a) {
            $totalTokens += (int)($a['tokens'] ?? 0);
            if (($a['status'] ?? '') === 'OK') $successCount++;
            $uniqueUsers[$a['telegram_user_id']] = true;
        }

        $userCount = count($uniqueUsers);
        $successRate = $total > 0 ? round(($successCount / $total) * 100, 1) : 100;

        $rows = '';
        $i = 1;
        foreach ($activities as $a) {
            $time = htmlspecialchars($a['created_at'] ?? '', ENT_QUOTES, 'UTF-8');
            $user = htmlspecialchars(($a['full_name'] ?: ($a['username'] ? '@' . $a['username'] : 'User')) . " ({$a['telegram_user_id']})", ENT_QUOTES, 'UTF-8');
            $action = htmlspecialchars($a['action'] ?? '', ENT_QUOTES, 'UTF-8');
            $query = nl2br(htmlspecialchars(mb_substr($a['user_query'] ?? '', 0, 220), ENT_QUOTES, 'UTF-8'));
            $resp = nl2br(htmlspecialchars(mb_substr($a['ai_response'] ?? '', 0, 300), ENT_QUOTES, 'UTF-8'));
            $tok = number_format((int)($a['tokens'] ?? 0));
            $st = $a['status'] === 'ERROR' ? '<span class="badge badge-err">ERR</span>' : '<span class="badge badge-ok">OK</span>';

            $rows .= "<tr>
                <td style='text-align:center;'>{$i}</td>
                <td>{$time}</td>
                <td><b>{$user}</b></td>
                <td>{$action}</td>
                <td>{$query}</td>
                <td>{$resp}</td>
                <td style='text-align:center;'>{$tok}</td>
                <td style='text-align:center;'>{$st}</td>
            </tr>";
            $i++;
        }

        $html = "
        <!DOCTYPE html>
        <html lang='uz'>
        <head>
        <meta charset='UTF-8'>
        <style>
            @page { margin: 18px; }
            body { font-family: 'DejaVu Sans', sans-serif; font-size: 9px; color: #1e293b; line-height: 1.35; }
            .header { background: #0f172a; color: #ffffff; padding: 14px 18px; border-radius: 6px; margin-bottom: 12px; }
            .header h1 { margin: 0 0 5px 0; font-size: 16px; color: #38bdf8; }
            .stats-bar { margin-top: 8px; font-size: 9.5px; }
            .stat-pill { display: inline-block; background: #1e293b; padding: 4px 10px; border-radius: 4px; margin-right: 8px; color: #f8fafc; }
            .stat-pill b { color: #38bdf8; }
            .badge { display: inline-block; padding: 2px 5px; font-size: 8px; font-weight: bold; border-radius: 3px; }
            .badge-ok { background: #dcfce7; color: #166534; }
            .badge-err { background: #fee2e2; color: #991b1b; }
            table { width: 100%; border-collapse: collapse; }
            th { background: #1e293b; color: #ffffff; padding: 6px 4px; font-size: 8.5px; text-align: left; }
            td { padding: 5px 4px; border-bottom: 1px solid #e2e8f0; font-size: 8px; vertical-align: top; }
            tr:nth-child(even) td { background: #f8fafc; }
            .footer { position: fixed; bottom: -12px; left: 0; right: 0; text-align: center; font-size: 8px; color: #94a3b8; }
        </style>
        </head>
        <body>
            <div class='header'>
                <h1>👑 {$safeTitle}</h1>
                <div class='stats-bar'>
                    <div class='stat-pill'>Jami amallar: <b>{$total}</b></div>
                    <div class='stat-pill'>Foydalanuvchilar: <b>{$userCount}</b></div>
                    <div class='stat-pill'>Jami tokenlar: <b>" . number_format($totalTokens) . "</b></div>
                    <div class='stat-pill'>Muvaffaqiyat: <b>{$successRate}%</b></div>
                    <div class='stat-pill' style='float:right;'>Sana: {$now}</div>
                </div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style='width:22px; text-align:center;'>#</th>
                        <th style='width:80px;'>Vaqt</th>
                        <th style='width:120px;'>Foydalanuvchi</th>
                        <th style='width:80px;'>Amal turi</th>
                        <th style='width:160px;'>So'rov</th>
                        <th>AI Javobi / Natija</th>
                        <th style='width:40px; text-align:center;'>Token</th>
                        <th style='width:35px; text-align:center;'>Status</th>
                    </tr>
                </thead>
                <tbody>{$rows}</tbody>
            </table>
            <div class='footer'>Universal AI Assistant Bot (@assistanduzz_bot) — Administrator Audit Hisoboti</div>
        </body>
        </html>";

        $prefix = $todayOnly ? 'today_audit_' : 'full_audit_';
        $path = rtrim($this->tempDir, '/') . '/' . $prefix . time() . '.pdf';
        return $this->renderPdf($html, $path, 'landscape') ? $path : null;
    }

    public function exportAdminAuditCsv(array $activities): ?string {
        $path = rtrim($this->tempDir, '/') . '/audit_' . time() . '.csv';
        $fp = fopen($path, 'w');
        if (!$fp) return null;

        // UTF-8 BOM for Microsoft Excel compatibility
        fwrite($fp, "\xEF\xBB\xBF");

        fputcsv($fp, ['ID', 'Sana & Vaqt', 'Telegram ID', 'Username', 'FIO', 'Amal Turi', 'Foydalanuvchi So\'rovi', 'AI Javobi', 'Tokenlar', 'Status']);

        foreach ($activities as $a) {
            fputcsv($fp, [
                $a['id'] ?? '',
                $a['created_at'] ?? '',
                $a['telegram_user_id'] ?? '',
                $a['username'] ? '@' . $a['username'] : '',
                $a['full_name'] ?? '',
                $a['action'] ?? '',
                $a['user_query'] ?? '',
                $a['ai_response'] ?? '',
                $a['tokens'] ?? 0,
                $a['status'] ?? 'OK',
            ]);
        }

        fclose($fp);
        return file_exists($path) ? $path : null;
    }

    public function exportAdminAuditMd(array $activities, string $title): ?string {
        $now = date('d.m.Y H:i:s');
        $md = "# 👑 {$title}\n\n";
        $md .= "- **Eksport sanasi:** {$now}\n";
        $md .= "- **Jami yozuvlar:** " . count($activities) . " ta\n\n";
        $md .= "| ID | Sana | User ID | Username | FIO | Amal | So'rov | Token | Status |\n";
        $md .= "|:---|:---|:---|:---|:---|:---|:---|:---|:---|\n";

        foreach ($activities as $a) {
            $uname = $a['username'] ? '@' . $a['username'] : '—';
            $fname = str_replace('|', '/', $a['full_name'] ?? '—');
            $action = str_replace('|', '/', $a['action'] ?? '');
            $query = str_replace(["\r", "\n", '|'], [' ', ' ', '/'], mb_substr($a['user_query'] ?? '', 0, 80));
            $md .= "| {$a['id']} | {$a['created_at']} | `{$a['telegram_user_id']}` | {$uname} | {$fname} | {$action} | {$query} | {$a['tokens']} | {$a['status']} |\n";
        }

        $path = rtrim($this->tempDir, '/') . '/audit_' . time() . '.md';
        file_put_contents($path, $md);
        return file_exists($path) ? $path : null;
    }
}

// ============================================================================
// 6. KLAVIATURALAR
// ============================================================================

class Keyboards {
    public static function main(): array {
        return [
            'keyboard' => [
                [['text' => '🤖 AI Assistant'], ['text' => '🎙 Ovozli AI']],
                [['text' => '📄 Hujjat tahlili'], ['text' => '🖼 Rasm tahlili']],
                [['text' => '💻 Kod yordamchisi'], ['text' => '🛠 AI Tools']],
                [['text' => '🗑 Yangi suhbat'], ['text' => '📚 Tarix'], ['text' => '⚙️ Sozlamalar']],
            ],
            'resize_keyboard' => true,
        ];
    }

    public static function cancel(): array {
        return ['keyboard' => [[['text' => '⬅️ Asosiy menyu']]], 'resize_keyboard' => true];
    }

    public static function documentActions(): array {
        return ['inline_keyboard' => [
            [['text' => '📌 Qisqa xulosa', 'callback_data' => 'doc_action:summary'],
             ['text' => '🔍 Chuqur tahlil', 'callback_data' => 'doc_action:deep_analysis']],
            [['text' => '❓ Savol berish', 'callback_data' => 'doc_action:question'],
             ['text' => '⚠️ Muammolarni topish', 'callback_data' => 'doc_action:issues']],
            [['text' => '💡 Tavsiyalar', 'callback_data' => 'doc_action:recommendations'],
             ['text' => '📊 Asosiy ma\'lumotlar', 'callback_data' => 'doc_action:data']],
        ]];
    }

    public static function codingModes(): array {
        return ['inline_keyboard' => [
            [['text' => '🧑‍💻 Kod yozish', 'callback_data' => 'code_mode:write'],
             ['text' => '🐛 Bug topish', 'callback_data' => 'code_mode:bug']],
            [['text' => '🔧 Kodni tuzatish', 'callback_data' => 'code_mode:fix'],
             ['text' => '⚡ Optimallashtirish', 'callback_data' => 'code_mode:optimize']],
            [['text' => '📚 Tushuntirish', 'callback_data' => 'code_mode:explain'],
             ['text' => '🔄 Refactor', 'callback_data' => 'code_mode:refactor']],
            [['text' => '🧪 Test yozish', 'callback_data' => 'code_mode:test']],
        ]];
    }

    public static function aiTools(): array {
        return ['inline_keyboard' => [
            [['text' => '📝 Text Generator', 'callback_data' => 'tool:text_gen'],
             ['text' => '🌐 Translator', 'callback_data' => 'tool:translator']],
            [['text' => '📋 Summarizer', 'callback_data' => 'tool:summarizer'],
             ['text' => '💡 Idea Generator', 'callback_data' => 'tool:idea_gen']],
            [['text' => '✍️ Rewrite', 'callback_data' => 'tool:rewrite'],
             ['text' => '📢 Marketing', 'callback_data' => 'tool:marketing']],
            [['text' => '🎓 Study Assistant', 'callback_data' => 'tool:study'],
             ['text' => '📊 Data Analysis', 'callback_data' => 'tool:data_analysis']],
        ]];
    }

    public static function rewriteStyles(): array {
        return ['inline_keyboard' => [
            [['text' => '👔 Professional', 'callback_data' => 'style:professional'],
             ['text' => '😊 Friendly', 'callback_data' => 'style:friendly']],
            [['text' => '📜 Formal', 'callback_data' => 'style:formal'],
             ['text' => '⚡ Short', 'callback_data' => 'style:short']],
            [['text' => '💡 Simple', 'callback_data' => 'style:simple']],
        ]];
    }

    public static function translatorLanguages(): array {
        return ['inline_keyboard' => [
            [['text' => '🇺🇿 O\'zbekcha', 'callback_data' => 'tr_lang:o\'zbek'],
             ['text' => '🇷🇺 Русский', 'callback_data' => 'tr_lang:rus']],
            [['text' => '🇬🇧 English', 'callback_data' => 'tr_lang:ingliz'],
             ['text' => '🇩🇪 Deutsch', 'callback_data' => 'tr_lang:nemis']],
            [['text' => '🇹🇷 Türkçe', 'callback_data' => 'tr_lang:turk'],
             ['text' => '🇸🇦 العربية', 'callback_data' => 'tr_lang:arab']],
        ]];
    }

    public static function settings(): array {
        return ['inline_keyboard' => [
            [['text' => '🌐 Tilni tanlash', 'callback_data' => 'settings:language'],
             ['text' => '🧠 AI Model', 'callback_data' => 'settings:model']],
            [['text' => '📥 Tarixni yuklab olish (PDF/MD)', 'callback_data' => 'settings:export']],
            [['text' => '🗑 Tarixni tozalash', 'callback_data' => 'settings:clear_history'],
             ['text' => 'ℹ️ Bot haqida', 'callback_data' => 'settings:about']],
        ]];
    }

    public static function userExport(): array {
        return ['inline_keyboard' => [
            [['text' => '📑 Faol suhbatni PDF yuklash', 'callback_data' => 'export:chat_pdf'],
             ['text' => '📝 Faol suhbatni MD yuklash', 'callback_data' => 'export:chat_md']],
            [['text' => '📊 Barcha amallarim (PDF)', 'callback_data' => 'export:user_all_pdf'],
             ['text' => '📝 Barcha amallarim (MD)', 'callback_data' => 'export:user_all_md']],
            [['text' => '⬅️ Orqaga', 'callback_data' => 'export:back_to_settings']],
        ]];
    }

    public static function languageSelect(): array {
        return ['inline_keyboard' => [[
            ['text' => '🇺🇿 O\'zbekcha', 'callback_data' => 'lang:uz'],
            ['text' => '🇷🇺 Русский', 'callback_data' => 'lang:ru'],
            ['text' => '🇬🇧 English', 'callback_data' => 'lang:en'],
        ]]];
    }

    public static function history(array $convs): array {
        $rows = [];
        foreach ($convs as $c) {
            $icon = $c['is_active'] ? '🟢 ' : '⚪ ';
            $rows[] = [['text' => $icon . mb_substr($c['title'], 0, 25), 'callback_data' => 'switch_conv:' . $c['id']]];
        }
        $rows[] = [
            ['text' => '📑 PDF yuklab olish', 'callback_data' => 'export:chat_pdf'],
            ['text' => '📝 MD yuklab olish', 'callback_data' => 'export:chat_md'],
        ];
        return ['inline_keyboard' => $rows];
    }

    public static function admin(): array {
        return ['inline_keyboard' => [
            [['text' => '📊 Kengaytirilgan Statistika', 'callback_data' => 'admin:stats'],
             ['text' => '📨 Xabar tarqatish', 'callback_data' => 'admin:broadcast']],
            [['text' => '👥 Foydalanuvchilar', 'callback_data' => 'admin:users'],
             ['text' => '⚙️ Tizim holati', 'callback_data' => 'admin:health']],
            [['text' => '📥 Hisobotlarni Eksport qilish (PDF/CSV)', 'callback_data' => 'admin:exports']],
            [['text' => '📈 Google Sheets', 'callback_data' => 'admin:sheets'],
             ['text' => '🧹 Keshni tozalash', 'callback_data' => 'admin:cleanup']],
        ]];
    }

    public static function adminExports(): array {
        return ['inline_keyboard' => [
            [['text' => '📑 Barcha amallar (PDF hisobot)', 'callback_data' => 'admin_exp:all_pdf']],
            [['text' => '📊 Excel / CSV yuklab olish', 'callback_data' => 'admin_exp:all_csv'],
             ['text' => '📝 Markdown (.md) hisobot', 'callback_data' => 'admin_exp:all_md']],
            [['text' => '📅 Bugungi amallar (PDF)', 'callback_data' => 'admin_exp:today_pdf']],
            [['text' => '⬅️ Orqaga', 'callback_data' => 'admin:back']],
        ]];
    }

    public static function sheets(?string $sheetUrl = null): array {
        $buttons = [];
        if (!empty($sheetUrl)) {
            $buttons[] = [['text' => '🔗 Jadvalni ochish (Google Sheets)', 'url' => $sheetUrl]];
        }
        $buttons[] = [
            ['text' => '🧪 Sinov yuborish (Ping)', 'callback_data' => 'sheets:ping'],
            ['text' => '📥 Oxirgi 50 ta harakatni yuklash', 'callback_data' => 'sheets:sync'],
        ];
        $buttons[] = [
            ['text' => '⚙️ Webhook URL kiritish', 'callback_data' => 'sheets:set_webhook'],
            ['text' => 'ℹ️ Sozlash qo\'llanmasi', 'callback_data' => 'sheets:guide'],
        ];
        $buttons[] = [
            ['text' => '⬅️ Orqaga', 'callback_data' => 'admin:back'],
        ];
        return ['inline_keyboard' => $buttons];
    }
}

// ============================================================================
// 7. FSM (STATE) — faylga yozib boramiz
// ============================================================================

class StateManager {
    private string $dir;
    public function __construct(string $dir) {
        $this->dir = rtrim($dir, '/');
        if (!is_dir($this->dir)) mkdir($this->dir, 0777, true);
    }
    private function file(int $uid): string { return "{$this->dir}/state_{$uid}.json"; }

    public function get(int $uid): array {
        $f = $this->file($uid);
        if (!file_exists($f)) return ['state' => null, 'data' => []];
        $j = json_decode(file_get_contents($f), true);
        return $j ?: ['state' => null, 'data' => []];
    }

    public function setState(int $uid, ?string $state): void {
        $d = $this->get($uid); $d['state'] = $state;
        file_put_contents($this->file($uid), json_encode($d, JSON_UNESCAPED_UNICODE));
    }

    public function updateData(int $uid, array $data): void {
        $d = $this->get($uid);
        $d['data'] = array_merge($d['data'] ?? [], $data);
        file_put_contents($this->file($uid), json_encode($d, JSON_UNESCAPED_UNICODE));
    }

    public function clear(int $uid): void {
        @unlink($this->file($uid));
    }

    public function getState(int $uid): ?string { return $this->get($uid)['state'] ?? null; }
    public function getData(int $uid): array { return $this->get($uid)['data'] ?? []; }
}

// ============================================================================
// 8. RATE LIMIT MIDDLEWARE (in-memory)
// ============================================================================

class RateLimiter {
    private array $lastTime = [];
    public function __construct(private float $delay = 0.8) {}

    /** true qaytarsa — so'rov rad etiladi */
    public function check(int $uid): bool {
        $now = microtime(true);
        if ($now - ($this->lastTime[$uid] ?? 0.0) < $this->delay) return true;
        $this->lastTime[$uid] = $now;
        return false;
    }
}

// ============================================================================
// 9. ASOSIY BOT HANDLER SINFI
// ============================================================================

class TelegramBot {
    private TelegramAPI $tg;
    private UserRepository $users;
    private ConversationRepository $convs;
    private MessageRepository $msgs;
    private UsageRepository $usages;
    private FileRepository $files;
    private ActivityLogRepository $activityLogs;
    private SettingsRepository $botSettings;
    private ExportService $exporter;
    private GoogleSheetsService $sheets;
    private GeminiService $gemini;
    private DocumentService $docService;
    private SpeechService $speechService;
    private ImageService $imageService;
    private StateManager $states;
    private RateLimiter $rateLimit;
    private Logger $logger;
    private Settings $settings;
    private PDO $pdo;
    private int $offset = 0;

    public function __construct(
        Settings $settings,
        PDO $pdo,
        TelegramAPI $tg,
        Logger $logger
    ) {
        $this->settings = $settings;
        $this->pdo = $pdo;
        $this->tg = $tg;
        $this->logger = $logger;

        $this->users   = new UserRepository($pdo, $settings);
        $this->convs   = new ConversationRepository($pdo);
        $this->msgs    = new MessageRepository($pdo, $this->convs);
        $this->usages  = new UsageRepository($pdo, $settings);
        $this->files   = new FileRepository($pdo);

        $this->activityLogs  = new ActivityLogRepository($pdo);
        $this->botSettings   = new SettingsRepository($pdo);
        $this->exporter      = new ExportService($settings->TEMP_DIR);

        $savedWebhook = $this->botSettings->get('GOOGLE_SHEETS_WEBHOOK_URL');
        if (!empty($savedWebhook)) {
            $this->settings->GOOGLE_SHEETS_WEBHOOK_URL = $savedWebhook;
        }
        $savedSheetUrl = $this->botSettings->get('GOOGLE_SHEET_URL');
        if (!empty($savedSheetUrl)) {
            $this->settings->GOOGLE_SHEET_URL = $savedSheetUrl;
        }

        $this->sheets        = new GoogleSheetsService($settings, $logger);

        $this->gemini        = new GeminiService($settings);
        $this->docService    = new DocumentService();
        $this->speechService = new SpeechService($tg, $this->gemini);
        $this->imageService  = new ImageService($tg, $this->gemini);

        $this->states    = new StateManager($settings->DATA_DIR . '/states');
        $this->rateLimit = new RateLimiter(0.8);
    }

    private function logAction(
        int $uid,
        ?string $username,
        ?string $fullName,
        string $action,
        string $query = '',
        string $response = '',
        int $tokens = 0,
        string $status = 'OK'
    ): void {
        try {
            $this->activityLogs->record($uid, $username, $fullName, $action, $query, $response, $tokens, $status);
            $this->sheets->log($uid, $username, $fullName, $action, $query, $response, $tokens, $status);
        } catch (Throwable $e) {
            $this->logger->error("Activity log error: " . $e->getMessage());
        }
    }

    // ------------------------------------------------------------------
    // ASOSIY POLLING LOOP
    // ------------------------------------------------------------------
    public function run(): void {
        $this->tg->call('deleteWebhook', ['drop_pending_updates' => true]);
        $this->registerCommands();

        $me = $this->tg->call('getMe');
        $username = $me['result']['username'] ?? 'unknown';
        $this->logger->info("Bot started: @{$username} | Model: {$this->settings->GEMINI_MODEL}");

        while (true) {
            try {
                $resp = $this->tg->call('getUpdates', [
                    'offset' => $this->offset,
                    'timeout' => 25,
                    'allowed_updates' => ['message', 'callback_query'],
                ]);
                if (!($resp['ok'] ?? false)) {
                    sleep(3);
                    continue;
                }
                foreach ($resp['result'] as $update) {
                    $this->offset = $update['update_id'] + 1;
                    try {
                        $this->handleUpdate($update);
                    } catch (Throwable $e) {
                        $this->logger->error("Update handler error: " . $e->getMessage());
                    }
                }
            } catch (Throwable $e) {
                $this->logger->error("Polling error: " . $e->getMessage());
                sleep(3);
            }
        }
    }

    private function registerCommands(): void {
        $this->tg->call('setMyCommands', ['commands' => [
            ['command' => 'start',    'description' => 'Botni ishga tushirish'],
            ['command' => 'newchat',  'description' => 'Yangi suhbat boshlash'],
            ['command' => 'history',  'description' => 'Suhbatlar tarixi'],
            ['command' => 'settings', 'description' => 'Sozlamalar va til'],
            ['command' => 'help',     'description' => 'Yordam'],
        ]]);
    }

    private function handleUpdate(array $update): void {
        if (isset($update['message'])) $this->handleMessage($update['message']);
        elseif (isset($update['callback_query'])) $this->handleCallback($update['callback_query']);
    }

    // ------------------------------------------------------------------
    // XABARLARNI QAYTA ISHLASH
    // ------------------------------------------------------------------
    private function handleMessage(array $msg): void {
        $user = $msg['from'] ?? null;
        if (!$user) return;
        $uid = (int)$user['id'];
        $chatId = (int)$msg['chat']['id'];

        // Auth + rate limit
        $dbUser = $this->users->getOrCreate($uid, $user['username'] ?? null, $user['first_name'] ?? null,
                                  $user['language_code'] ?? 'uz');

        if (!empty($dbUser['is_blocked']) && !$this->settings->isAdmin($uid)) {
            $this->tg->sendMessage($chatId, "⛔ <b>Hisobingiz bloklangan!</b>\nSavollar bo'yicha administratorga murojaat qiling.");
            return;
        }

        if ($this->rateLimit->check($uid) && !$this->settings->isAdmin($uid)) {
            $this->tg->sendMessage($chatId, "⚠️ Iltimos, biroz kuting.");
            return;
        }

        // Buyruqlar
        $text = $msg['text'] ?? '';
        if ($text !== '') {
            if ($text === '/start' || $text === '/help')    { $this->cmdStart($chatId, $text === '/help'); return; }
            if ($text === '/newchat')                       { $this->cmdNewChat($uid, $chatId, $user); return; }
            if ($text === '/history')                       { $this->cmdHistory($uid, $chatId); return; }
            if ($text === '/export' || $text === '/pdf')    { $this->tg->sendMessage($chatId, "📥 <b>Suhbat yoki amallar tarixingizni yuklab olish:</b>\n\nKerakli formatni tanlang 👇", Keyboards::userExport()); return; }
            if ($text === '/settings' || $text === '⚙️ Sozlamalar') { $this->cmdSettings($chatId); return; }
            if ($text === '/admin' && $this->settings->isAdmin($uid)) { $this->cmdAdmin($chatId); return; }
            if (str_starts_with($text, '/set_sheets') && $this->settings->isAdmin($uid)) {
                $parts = explode(' ', trim($text), 2);
                $url = trim($parts[1] ?? '');
                if (empty($url)) {
                    $this->tg->sendMessage($chatId, "ℹ️ <b>Foydalanish:</b>\n<code>/set_sheets https://script.google.com/macros/s/.../exec</code>");
                    return;
                }
                $this->handleSetWebhook($uid, $chatId, $url);
                return;
            }
            if ($this->settings->isAdmin($uid)) {
                if (str_starts_with($text, '/block ')) {
                    $targetId = (int)trim(substr($text, 7));
                    $this->users->setBlocked($targetId, true);
                    $this->tg->sendMessage($chatId, "⛔ Foydalanuvchi <code>{$targetId}</code> muvaffaqiyatli bloklandi.");
                    return;
                }
                if (str_starts_with($text, '/unblock ')) {
                    $targetId = (int)trim(substr($text, 9));
                    $this->users->setBlocked($targetId, false);
                    $this->tg->sendMessage($chatId, "✅ Foydalanuvchi <code>{$targetId}</code> blokdan chiqarildi.");
                    return;
                }
                if (str_starts_with($text, '/set_limit ')) {
                    $p = preg_split('/\s+/', trim($text));
                    if (count($p) >= 3) {
                        $targetId = (int)$p[1];
                        $limit = (int)$p[2];
                        $this->users->setLimitOverride($targetId, $limit);
                        $this->tg->sendMessage($chatId, "✅ Foydalanuvchi <code>{$targetId}</code> uchun kunlik limit <b>{$limit}</b> ta etib belgilandi.");
                        return;
                    }
                }
                if (str_starts_with($text, '/set_daily_limit ')) {
                    $newLimit = (int)trim(substr($text, 17));
                    if ($newLimit > 0) {
                        $this->botSettings->set('FREE_DAILY_LIMIT', (string)$newLimit);
                        $this->settings->FREE_DAILY_LIMIT = $newLimit;
                        $this->tg->sendMessage($chatId, "✅ Barcha foydalanuvchilar uchun umumiy kunlik limit <b>{$newLimit}</b> ta etib yangilandi.");
                        return;
                    }
                }
                if (str_starts_with($text, '/user ')) {
                    $targetId = (int)trim(substr($text, 6));
                    $this->cmdUserInspect($chatId, $targetId);
                    return;
                }
                if ($text === '/cleanup') {
                    $this->cmdCleanup($chatId);
                    return;
                }
            }

            // Menyu tugmalari
            switch ($text) {
                case '🤖 AI Assistant':
                    $this->states->clear($uid);
                    $this->tg->sendMessage($chatId, "🤖 <b>AI Assistant bo'limi faol!</b>\n\nMenga istalgan savol, g'oya, reja, matematika, tarjima yoki dasturlash bo'yicha vazifangizni yozing.\nSuhbatni noldan boshlash uchun: <b>🗑 Yangi suhbat</b>");
                    return;
                case '🎙 Ovozli AI':
                    $this->states->clear($uid);
                    $this->tg->sendMessage($chatId, "🎙 <b>Ovozli AI bo'limi!</b>\n\nMenga Telegram ovozli xabarini (voice note) yuboring.");
                    return;
                case '📄 Hujjat tahlili':
                    $this->states->clear($uid);
                    $this->tg->sendMessage($chatId,
                        "📄 <b>Hujjat tahlili bo'limi!</b>\n\nMenga PDF, Word (.docx), Excel (.xlsx), CSV, TXT yoki PPTX faylini yuboring.\n<i>Maksimal hajm: {$this->settings->MAX_FILE_SIZE_MB} MB</i>",
                        Keyboards::cancel());
                    return;
                case '🖼 Rasm tahlili':
                    $this->states->clear($uid);
                    $this->tg->sendMessage($chatId,
                        "🖼 <b>Rasm tahlili bo'limi!</b>\n\nMenga rasm, skrinshot, grafik yoki jadval yuboring. Izoh (caption) qo'shishingiz mumkin.",
                        Keyboards::cancel());
                    return;
                case '💻 Kod yordamchisi':
                    $this->states->clear($uid);
                    $this->tg->sendMessage($chatId,
                        "💻 <b>Professional Kod Yordamchisi!</b>\n\nKerakli yo'nalishni tanlang 👇",
                        Keyboards::codingModes());
                    return;
                case '🛠 AI Tools':
                    $this->states->clear($uid);
                    $this->tg->sendMessage($chatId,
                        "🛠 <b>AI Tools bo'limi!</b>\n\nKerakli vositani tanlang 👇",
                        Keyboards::aiTools());
                    return;
                case '🗑 Yangi suhbat':
                    $this->cmdNewChat($uid, $chatId, $user);
                    return;
                case '📚 Tarix':
                    $this->cmdHistory($uid, $chatId);
                    return;
                case '⬅️ Asosiy menyu':
                    $this->states->clear($uid);
                    $this->tg->sendMessage($chatId, "Asosiy menyuga qaytdingiz 👇", Keyboards::main());
                    return;
                case '❓ Yordam':
                    $this->cmdStart($chatId, true);
                    return;
            }
        }

        // FSM states bo'yicha yo'naltirish
        $state = $this->states->getState($uid);

        if ($state === 'admin_set_webhook' && $text !== '') { $this->handleSetWebhook($uid, $chatId, $text); return; }
        if ($state === 'doc_question' && $text !== '') { $this->handleDocQuestion($uid, $chatId, $text, $user); return; }
        if ($state === 'coding_wait' && $text !== '')  { $this->handleCodingInput($uid, $chatId, $text, $user); return; }
        if ($state === 'tool_wait' && $text !== '')    { $this->handleToolInput($uid, $chatId, $text, $user); return; }
        if ($state === 'admin_broadcast')              { $this->handleBroadcast($uid, $msg); return; }

        // Fayllar va media
        if (isset($msg['voice']) || isset($msg['audio'])) { $this->handleVoice($uid, $chatId, $msg); return; }
        if (isset($msg['photo']))                          { $this->handlePhoto($uid, $chatId, $msg); return; }
        if (isset($msg['document']))                       { $this->handleDocument($uid, $chatId, $msg); return; }

        // Umumiy chat
        if ($text !== '' && !str_starts_with($text, '/')) {
            $this->handleGeneralChat($uid, $chatId, $text, $user);
        }
    }

    // ------------------------------------------------------------------
    // CALLBACK QUERY
    // ------------------------------------------------------------------
    private function handleCallback(array $cb): void {
        $user = $cb['from']; $uid = (int)$user['id'];
        $chatId = (int)$cb['message']['chat']['id'];
        $msgId  = (int)$cb['message']['message_id'];
        $data   = $cb['data'] ?? '';

        $this->users->getOrCreate($uid, $user['username'] ?? null, $user['first_name'] ?? null);

        // Document actions
        if (str_starts_with($data, 'doc_action:'))  { $this->cbDocAction($uid, $chatId, $msgId, $cb['id'], $data, $user); return; }
        if (str_starts_with($data, 'code_mode:'))   { $this->cbCodeMode($uid, $chatId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'tool:'))        { $this->cbTool($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'tr_lang:'))     { $this->cbTrLang($uid, $chatId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'style:'))       { $this->cbStyle($uid, $chatId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'switch_conv:')) { $this->cbSwitchConv($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if ($data === 'settings:language')          { $this->tg->editMessageText($chatId, $msgId, "🌐 Tilni tanlang:", Keyboards::languageSelect()); $this->tg->answerCallback($cb['id']); return; }
        if (str_starts_with($data, 'lang:'))        { $this->cbLang($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if ($data === 'settings:model')             { $this->cbModel($chatId, $msgId, $cb['id']); return; }
        if ($data === 'settings:clear_history')     { $this->cbClearHistory($uid, $chatId, $msgId, $cb['id'], $user); return; }
        if ($data === 'settings:about')             { $this->cbAbout($chatId, $msgId, $cb['id']); return; }
        if ($data === 'settings:export')            { $this->tg->editMessageText($chatId, $msgId, "📥 <b>Suhbat yoki amallar tarixingizni yuklab olish:</b>\n\nKerakli formatni tanlang 👇", Keyboards::userExport()); $this->tg->answerCallback($cb['id']); return; }
        if ($data === 'export:back_to_settings')    { $this->tg->editMessageText($chatId, $msgId, "⚙️ <b>Bot Sozlamalari:</b>", Keyboards::settings()); $this->tg->answerCallback($cb['id']); return; }
        if (str_starts_with($data, 'export:'))      { $this->cbUserExport($uid, $chatId, $msgId, $cb['id'], $data, $user); return; }
        if (str_starts_with($data, 'admin_exp:'))   { $this->cbAdminExport($uid, $chatId, $msgId, $cb['id'], $data); return; }

        // Admin & Sheets
        if (str_starts_with($data, 'admin:'))  { $this->cbAdmin($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'sheets:')) { $this->cbSheets($uid, $chatId, $msgId, $cb['id'], $data); return; }

        $this->tg->answerCallback($cb['id']);
    }

    // ------------------------------------------------------------------
    // BUYRUQ HANDLERLARI
    // ------------------------------------------------------------------
    private function cmdStart(int $chatId, bool $help = false): void {
        $text = $help
            ? "📖 <b>Botdan foydalanish bo'yicha qo'llanma:</b>\n\n" .
              "1️⃣ <b>🤖 AI Assistant:</b> To'g'ridan-to'g'ri xabar yozing.\n" .
              "2️⃣ <b>🎙 Ovozli AI:</b> Telegram ovozli xabarlarini yuboring.\n" .
              "3️⃣ <b>📄 Hujjat tahlili:</b> PDF, DOCX, XLSX, TXT, PPTX yuboring.\n" .
              "4️⃣ <b>🖼 Rasm tahlili:</b> Skrinshot yoki rasm yuboring.\n" .
              "5️⃣ <b>💻 Kod yordamchisi:</b> Dasturlash savollari.\n" .
              "6️⃣ <b>🛠 AI Tools:</b> Tarjima, marketing va boshqa vositalar.\n\n" .
              "Buyruqlar: /start, /newchat, /history, /settings, /help"
            : "👋 <b>Assalomu alaykum!</b>\n\n" .
              "Men sizning shaxsiy <b>Universal AI Assistant</b> botingizman.\n\n" .
              "🤖 <b>AI Assistant</b> — Savollarga javoblar\n" .
              "🎙 <b>Ovozli AI</b> — Ovozli xabarlarni tushunish\n" .
              "📄 <b>Hujjat tahlili</b> — PDF, Word, Excel tahlili\n" .
              "🖼 <b>Rasm tahlili</b> — Rasmlar va grafiklar\n" .
              "💻 <b>Kod yordamchisi</b> — Kod yozish va tuzatish\n" .
              "🛠 <b>AI Tools</b> — Tarjimon, rewriter, marketing\n\n" .
              "Quyidagi menyudan kerakli bo'limni tanlang 👇";

        $this->tg->sendMessage($chatId, $text, Keyboards::main());
    }

    private function cmdNewChat(int $uid, int $chatId, ?array $user = null): void {
        $this->states->clear($uid);
        $this->convs->createNew($uid);
        $uname = $user['username'] ?? null;
        $fname = trim(($user['first_name'] ?? '') . ' ' . ($user['last_name'] ?? ''));
        $this->logAction($uid, $uname, $fname, '🗑 Yangi suhbat', '/newchat', 'Yangi suhbat ochildi, xotira tozalandi.', 0, 'OK');
        $this->tg->sendMessage($chatId, "✨ <b>Yangi suhbat boshlandi!</b> Eski xotira tozalandi.", Keyboards::main());
    }

    private function cmdHistory(int $uid, int $chatId): void {
        $convs = $this->convs->getUserConversations($uid, 8);
        if (!$convs) { $this->tg->sendMessage($chatId, "📚 Suhbatlar tarixi mavjud emas."); return; }
        $this->tg->sendMessage($chatId, "📚 <b>Suhbatlar tarixi:</b>", Keyboards::history($convs));
    }

    private function cmdSettings(int $chatId): void {
        $this->tg->sendMessage($chatId, "⚙️ <b>Bot Sozlamalari:</b>", Keyboards::settings());
    }

    private function cmdAdmin(int $chatId): void {
        $this->tg->sendMessage($chatId, "👑 <b>Admin Paneli:</b>", Keyboards::admin());
    }

    // ------------------------------------------------------------------
    // OVOZLI XABAR
    // ------------------------------------------------------------------
    private function handleVoice(int $uid, int $chatId, array $msg): void {
        [$ok, $used, $limit] = $this->usages->canRequest($uid);
        if (!$ok) { $this->tg->sendMessage($chatId, "⚠️ Bugungi bepul limit tugadi ({$used}/{$limit})."); return; }

        $statusMsg = $this->tg->sendMessage($chatId, "🎙 <i>Ovoz qabul qilindi...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;
        $this->tg->sendChatAction($chatId, 'record_voice');

        $fileId = $msg['voice']['file_id'] ?? $msg['audio']['file_id'];
        $ext = isset($msg['voice']) ? 'ogg' : 'mp3';

        [$transcription, $err] = $this->speechService->processVoice($fileId, $ext);
        if ($err || !$transcription) {
            $this->tg->editMessageText($chatId, $statusId, $err ?? "❌ Ovoz tushunilmadi.");
            return;
        }

        $this->tg->editMessageText($chatId, $statusId,
            "🎙 <b>Ovoz qabul qilindi.</b>\n\n📝 <b>Transkripsiya:</b>\n<i>\"{$transcription}\"</i>\n\n🤔 <i>AI javob tayyorlamoqda...</i>");

        $conv = $this->convs->getOrCreateActive($uid);
        $history = $this->msgs->recent((int)$conv['id'], 6);
        $payload = $history;
        $payload[] = ['role' => 'user', 'content' => $transcription];

        try {
            $res = $this->gemini->chat($payload);
            $this->msgs->add((int)$conv['id'], 'user', "[Voice]: {$transcription}");
            $this->msgs->add((int)$conv['id'], 'assistant', $res['content']);
            $this->convs->updateTitleIfDefault((int)$conv['id'], $transcription);
            $this->usages->record($uid, 'voice', $res['tokens']);

            $u = $msg['from'] ?? [];
            $fullName = trim(($u['first_name'] ?? '') . ' ' . ($u['last_name'] ?? ''));
            $this->logAction($uid, $u['username'] ?? null, $fullName, '🎙 Ovozli xabar', $transcription, $res['content'], $res['tokens'], 'OK');

            $this->tg->editMessageText($chatId, $statusId,
                "📝 <b>Transkripsiya:</b>\n<i>\"{$transcription}\"</i>");

            foreach (Helpers::splitText("🤖 <b>AI:</b>\n\n" . $res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk);
            }
        } catch (Throwable $e) {
            $this->logger->error("Voice response error: " . $e->getMessage());
            $u = $msg['from'] ?? [];
            $fullName = trim(($u['first_name'] ?? '') . ' ' . ($u['last_name'] ?? ''));
            $this->logAction($uid, $u['username'] ?? null, $fullName, '🎙 Ovozli xabar', $transcription, $e->getMessage(), 0, 'ERROR');
            $this->tg->sendMessage($chatId, "❌ Ovoz tahlil qilindi, ammo javob berishda xatolik.");
        }
    }

    // ------------------------------------------------------------------
    // RASM
    // ------------------------------------------------------------------
    private function handlePhoto(int $uid, int $chatId, array $msg): void {
        [$ok, $used, $limit] = $this->usages->canRequest($uid);
        if (!$ok) { $this->tg->sendMessage($chatId, "⚠️ Bugungi bepul limit tugadi ({$used}/{$limit})."); return; }

        $photos = $msg['photo'];
        $photo = end($photos);
        $prompt = $msg['caption'] ?? "Ushbu rasmni tahlil qiling va undagi barcha ma'lumotlarni tushuntirib bering.";

        $statusMsg = $this->tg->sendMessage($chatId, "🤔 <i>Rasm tahlil qilinmoqda...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;
        $this->tg->sendChatAction($chatId, 'upload_photo');

        try {
            $res = $this->imageService->processPhoto($photo['file_id'], $prompt);
            $this->usages->record($uid, 'image', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);

            $u = $msg['from'] ?? [];
            $fullName = trim(($u['first_name'] ?? '') . ' ' . ($u['last_name'] ?? ''));
            $this->logAction($uid, $u['username'] ?? null, $fullName, '🖼 Rasm tahlili', $prompt, $res['content'], $res['tokens'], 'OK');

            $prefix = isset($msg['caption']) ? "📝 <i>Izoh: \"{$msg['caption']}\"</i>\n\n" : "";
            foreach (Helpers::splitText("🖼 <b>Rasm tahlili:</b>\n\n{$prefix}{$res['content']}") as $chunk) {
                $this->tg->sendMessage($chatId, $chunk);
            }
        } catch (Throwable $e) {
            $this->logger->error("Image error: " . $e->getMessage());
            $u = $msg['from'] ?? [];
            $fullName = trim(($u['first_name'] ?? '') . ' ' . ($u['last_name'] ?? ''));
            $this->logAction($uid, $u['username'] ?? null, $fullName, '🖼 Rasm tahlili', $prompt, $e->getMessage(), 0, 'ERROR');
            $this->tg->editMessageText($chatId, $statusId, "❌ Rasmni tahlil qilishda xatolik.");
        }
    }

    // ------------------------------------------------------------------
    // HUJJAT
    // ------------------------------------------------------------------
    private function handleDocument(int $uid, int $chatId, array $msg): void {
        $doc = $msg['document'];
        $filename = $doc['file_name'] ?? 'document.txt';
        $fileSize = (int)($doc['file_size'] ?? 0);

        if ($fileSize > $this->settings->maxFileSizeBytes()) {
            $this->tg->sendMessage($chatId, "❌ Fayl hajmi limitdan ({$this->settings->MAX_FILE_SIZE_MB} MB) oshib ketdi.");
            return;
        }
        if (!Helpers::validateExt($filename, Helpers::ALLOWED_DOC_EXT)) {
            $this->tg->sendMessage($chatId, "❌ Fayl formati qo'llab-quvvatlanmaydi. Ruxsat: " . implode(', ', Helpers::ALLOWED_DOC_EXT));
            return;
        }

        [$ok, $used, $limit] = $this->usages->canRequest($uid);
        if (!$ok) { $this->tg->sendMessage($chatId, "⚠️ Bugungi bepul limit tugadi ({$used}/{$limit})."); return; }

        $statusMsg = $this->tg->sendMessage($chatId, "⏳ <i>Hujjat tekshirilmoqda va matn olinmoqda...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;

        $safeName = Helpers::sanitizeFilename($filename);
        $tempPath = $this->settings->TEMP_DIR . '/' . Helpers::uuid() . '_' . $safeName;

        try {
            $filePath = $this->tg->getFile($doc['file_id']);
            if (!$filePath) throw new RuntimeException('Faylni yuklab bo\'lmadi.');
            if (!$this->tg->downloadFile($filePath, $tempPath)) throw new RuntimeException('Faylni saqlab bo\'lmadi.');

            $extracted = $this->docService->extractText($tempPath);
            if (mb_strlen(trim($extracted)) < 10) {
                $this->tg->editMessageText($chatId, $statusId, "❌ Hujjatdan o'qish uchun yaroqli matn topilmadi.");
                return;
            }

            $this->files->record($uid, $filename, '.' . strtolower(pathinfo($filename, PATHINFO_EXTENSION)), $fileSize);

            $this->states->updateData($uid, ['doc_filename' => $filename, 'doc_text' => $extracted]);

            $size = Helpers::formatSize($fileSize);
            $chars = number_format(mb_strlen($extracted));
            $this->tg->editMessageText($chatId, $statusId,
                "✅ <b>Hujjat qabul qilindi!</b>\n\n" .
                "📁 <b>Nomi:</b> <code>{$filename}</code>\n" .
                "📦 <b>Hajmi:</b> {$size}\n" .
                "📝 <b>Belgilar soni:</b> {$chars}\n\n" .
                "Quyidagi tahlil turini tanlang 👇",
                Keyboards::documentActions());
        } catch (Throwable $e) {
            $this->logger->error("Doc error: " . $e->getMessage());
            $this->tg->editMessageText($chatId, $statusId, "❌ Xatolik yuz berdi: " . $e->getMessage());
        } finally {
            @unlink($tempPath);
        }
    }

    // ------------------------------------------------------------------
    // DOC ACTION CALLBACK
    // ------------------------------------------------------------------
    private function cbDocAction(int $uid, int $chatId, int $msgId, string $cbId, string $data, ?array $user = null): void {
        $action = explode(':', $data, 2)[1];
        $stateData = $this->states->getData($uid);
        $docText = $stateData['doc_text'] ?? null;
        $filename = $stateData['doc_filename'] ?? 'Hujjat';

        if (!$docText) {
            $this->tg->answerCallback($cbId, "Hujjat muddati o'tgan, qaytadan yuboring.", true);
            return;
        }

        if ($action === 'question') {
            $this->states->setState($uid, 'doc_question');
            $this->tg->answerCallback($cbId);
            $this->tg->sendMessage($chatId, "❓ <b>\"{$filename}\" bo'yicha savolingizni yozing:</b>");
            return;
        }

        $this->tg->answerCallback($cbId, "Tahlil boshlandi...");
        $statusMsg = $this->tg->sendMessage($chatId, "🤔 <i>Tahlil qilinmoqda, kuting...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;

        $uname = $user['username'] ?? null;
        $fname = trim(($user['first_name'] ?? '') . ' ' . ($user['last_name'] ?? ''));

        try {
            $res = $this->gemini->analyzeDocument($docText, $action);
            $this->usages->record($uid, 'document', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);
            $this->logAction($uid, $uname, $fname, "📄 Hujjat: {$action} ({$filename})", "Hujjat tahlili", $res['content'], $res['tokens'], 'OK');

            $titles = [
                'summary' => '📌 Qisqa xulosa', 'deep_analysis' => '🔍 Chuqur tahlil',
                'issues' => '⚠️ Muammolar', 'recommendations' => '💡 Tavsiyalar', 'data' => '📊 Asosiy ma\'lumotlar',
            ];
            $title = $titles[$action] ?? 'Tahlil';
            foreach (Helpers::splitText("<b>{$title} ({$filename}):</b>\n\n" . $res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk);
            }
        } catch (Throwable $e) {
            $this->logger->error("Doc action error: " . $e->getMessage());
            $this->logAction($uid, $uname, $fname, "📄 Hujjat: {$action} ({$filename})", "Hujjat tahlili", $e->getMessage(), 0, 'ERROR');
            $this->tg->editMessageText($chatId, $statusId, "❌ Tahlilda xatolik yuz berdi.");
        }
    }

    private function handleDocQuestion(int $uid, int $chatId, string $question, ?array $user = null): void {
        $data = $this->states->getData($uid);
        $docText = $data['doc_text'] ?? null;
        if (!$docText) {
            $this->tg->sendMessage($chatId, "Hujjat topilmadi.");
            $this->states->clear($uid);
            return;
        }

        $statusMsg = $this->tg->sendMessage($chatId, "🤔 <i>Hujjatdan javob qidirilmoqda...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;

        $uname = $user['username'] ?? null;
        $fname = trim(($user['first_name'] ?? '') . ' ' . ($user['last_name'] ?? ''));

        try {
            $res = $this->gemini->analyzeDocument($docText, 'question', $question);
            $this->usages->record($uid, 'document', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);
            $this->logAction($uid, $uname, $fname, "📄 Hujjat savoli", $question, $res['content'], $res['tokens'], 'OK');
            foreach (Helpers::splitText("❓ <b>Savol:</b> <i>{$question}</i>\n\n📌 <b>Javob:</b>\n\n" . $res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk);
            }
        } catch (Throwable $e) {
            $this->logAction($uid, $uname, $fname, "📄 Hujjat savoli", $question, $e->getMessage(), 0, 'ERROR');
            $this->tg->editMessageText($chatId, $statusId, "❌ Savolga javob berishda xatolik.");
        }
        $this->states->setState($uid, null);
    }

    // ------------------------------------------------------------------
    // CODING
    // ------------------------------------------------------------------
    private function cbCodeMode(int $uid, int $chatId, string $cbId, string $data): void {
        $mode = explode(':', $data, 2)[1];
        $this->states->setState($uid, 'coding_wait');
        $this->states->updateData($uid, ['coding_mode' => $mode]);
        $this->tg->answerCallback($cbId);
        $this->tg->sendMessage($chatId, "Endi menga talablaringiz yoki kodingizni yuboring:", Keyboards::cancel());
    }

    private function handleCodingInput(int $uid, int $chatId, string $text, ?array $user = null): void {
        if ($text === '⬅️ Asosiy menyu') {
            $this->states->clear($uid);
            $this->tg->sendMessage($chatId, "Asosiy menyuga qaytdingiz.", Keyboards::main());
            return;
        }

        [$ok, $used, $limit] = $this->usages->canRequest($uid);
        if (!$ok) { $this->tg->sendMessage($chatId, "⚠️ Bugungi limit tugadi ({$used}/{$limit})."); return; }

        $mode = $this->states->getData($uid)['coding_mode'] ?? 'general';
        $statusMsg = $this->tg->sendMessage($chatId, "🤔 <i>Kod tayyorlanmoqda...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;
        $this->tg->sendChatAction($chatId, 'typing');

        $uname = $user['username'] ?? null;
        $fname = trim(($user['first_name'] ?? '') . ' ' . ($user['last_name'] ?? ''));

        try {
            $res = $this->gemini->coding($text, $mode);
            $this->usages->record($uid, 'coding', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);
            $this->logAction($uid, $uname, $fname, "💻 Kod: {$mode}", $text, $res['content'], $res['tokens'], 'OK');
            foreach (Helpers::splitText($res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk, null, 'Markdown');
            }
        } catch (Throwable $e) {
            $this->logger->error("Coding error: " . $e->getMessage());
            $this->logAction($uid, $uname, $fname, "💻 Kod: {$mode}", $text, $e->getMessage(), 0, 'ERROR');
            $this->tg->editMessageText($chatId, $statusId, "❌ Kodni qayta ishlashda xatolik.");
        }
        $this->states->setState($uid, null);
    }

    // ------------------------------------------------------------------
    // AI TOOLS
    // ------------------------------------------------------------------
    private function cbTool(int $uid, int $chatId, int $msgId, string $cbId, string $data): void {
        $tool = explode(':', $data, 2)[1];

        if ($tool === 'translator') {
            $this->states->updateData($uid, ['active_tool' => 'translator']);
            $this->tg->editMessageText($chatId, $msgId, "🌐 Tarjima tilini tanlang:", Keyboards::translatorLanguages());
            $this->tg->answerCallback($cbId);
            return;
        }
        if ($tool === 'rewrite') {
            $this->states->updateData($uid, ['active_tool' => 'rewrite']);
            $this->tg->editMessageText($chatId, $msgId, "✍️ Uslubni tanlang:", Keyboards::rewriteStyles());
            $this->tg->answerCallback($cbId);
            return;
        }

        $this->states->setState($uid, 'tool_wait');
        $this->states->updateData($uid, ['active_tool' => $tool]);
        $this->tg->answerCallback($cbId);
        $this->tg->sendMessage($chatId, "Iltimos, matn yoki mavzuni yuboring:", Keyboards::cancel());
    }

    private function cbTrLang(int $uid, int $chatId, string $cbId, string $data): void {
        $lang = explode(':', $data, 2)[1];
        $this->states->setState($uid, 'tool_wait');
        $this->states->updateData($uid, ['active_tool' => 'translator', 'target_lang' => $lang]);
        $this->tg->answerCallback($cbId);
        $this->tg->sendMessage($chatId, "🌐 Til: " . ucfirst($lang) . ". Matnni yuboring:", Keyboards::cancel());
    }

    private function cbStyle(int $uid, int $chatId, string $cbId, string $data): void {
        $style = explode(':', $data, 2)[1];
        $this->states->setState($uid, 'tool_wait');
        $this->states->updateData($uid, ['active_tool' => 'rewrite', 'rewrite_style' => $style]);
        $this->tg->answerCallback($cbId);
        $this->tg->sendMessage($chatId, "✍️ Uslub: " . ucfirst($style) . ". Matnni yuboring:", Keyboards::cancel());
    }

    private function handleToolInput(int $uid, int $chatId, string $text, ?array $user = null): void {
        if ($text === '⬅️ Asosiy menyu') {
            $this->states->clear($uid);
            $this->tg->sendMessage($chatId, "Asosiy menyuga qaytdingiz.", Keyboards::main());
            return;
        }

        [$ok, $used, $limit] = $this->usages->canRequest($uid);
        if (!$ok) { $this->tg->sendMessage($chatId, "⚠️ Bugungi limit tugadi ({$used}/{$limit})."); return; }

        $data = $this->states->getData($uid);
        $tool = $data['active_tool'] ?? 'text_gen';
        $params = [
            'target_lang' => $data['target_lang'] ?? "o'zbek",
            'style'       => $data['rewrite_style'] ?? 'professional',
        ];

        $statusMsg = $this->tg->sendMessage($chatId, "🤔 <i>Bajarilmoqda...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;
        $this->tg->sendChatAction($chatId, 'typing');

        $uname = $user['username'] ?? null;
        $fname = trim(($user['first_name'] ?? '') . ' ' . ($user['last_name'] ?? ''));

        try {
            $res = $this->gemini->runTool($tool, $text, $params);
            $this->usages->record($uid, 'tool', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);
            $this->logAction($uid, $uname, $fname, "🛠 Tool: {$tool}", $text, $res['content'], $res['tokens'], 'OK');
            foreach (Helpers::splitText($res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk, null, 'Markdown');
            }
        } catch (Throwable $e) {
            $this->logger->error("Tool error: " . $e->getMessage());
            $this->logAction($uid, $uname, $fname, "🛠 Tool: {$tool}", $text, $e->getMessage(), 0, 'ERROR');
            $this->tg->editMessageText($chatId, $statusId, "❌ Xatolik yuz berdi.");
        }
        $this->states->setState($uid, null);
    }

    // ------------------------------------------------------------------
    // SWITCH CONVERSATION
    // ------------------------------------------------------------------
    private function cbSwitchConv(int $uid, int $chatId, int $msgId, string $cbId, string $data): void {
        $cid = (int)explode(':', $data, 2)[1];
        $conv = $this->convs->setActive($uid, $cid);
        if ($conv) {
            $this->tg->answerCallback($cbId, "Faol suhbat: " . $conv['title']);
            $this->tg->editMessageText($chatId, $msgId, "✅ <b>Faol suhbat o'zgartirildi:</b> <i>{$conv['title']}</i>");
        } else {
            $this->tg->answerCallback($cbId, "Suhbat topilmadi.", true);
        }
    }

    // ------------------------------------------------------------------
    // SETTINGS CALLBACKS
    // ------------------------------------------------------------------
    private function cbLang(int $uid, int $chatId, int $msgId, string $cbId, string $data): void {
        $lang = explode(':', $data, 2)[1];
        $this->users->updateLanguage($uid, $lang);
        $this->tg->answerCallback($cbId, "Til saqlandi!");
        $this->tg->editMessageText($chatId, $msgId, "✅ Til o'zgartirildi: " . strtoupper($lang));
    }

    private function cbModel(int $chatId, int $msgId, string $cbId): void {
        $t = "🧠 <b>Model:</b> {$this->settings->GEMINI_MODEL}\n" .
             "🎙 <b>Ovoz:</b> Gemini Native Multimodal\n" .
             "Limit: {$this->settings->FREE_DAILY_LIMIT} ta/kun";
        $this->tg->editMessageText($chatId, $msgId, $t, Keyboards::settings());
        $this->tg->answerCallback($cbId);
    }

    private function cbClearHistory(int $uid, int $chatId, int $msgId, string $cbId, ?array $user = null): void {
        $this->convs->createNew($uid, 'Yangi suhbat');
        $uname = $user['username'] ?? null;
        $fname = trim(($user['first_name'] ?? '') . ' ' . ($user['last_name'] ?? ''));
        $this->logAction($uid, $uname, $fname, '🗑 Tarix tozalash', 'settings:clear_history', 'Faol suhbat tozalandi.', 0, 'OK');
        $this->tg->answerCallback($cbId, "Tarix tozalandi!", true);
        $this->tg->editMessageText($chatId, $msgId, "🗑 Faol suhbat tozalandi!", Keyboards::settings());
    }

    private function cbAbout(int $chatId, int $msgId, string $cbId): void {
        $this->tg->editMessageText($chatId, $msgId,
            "ℹ️ <b>Universal AI Assistant Bot</b>\n\nGoogle Gemini & PHP 8.2 asosida yaratilgan professional Telegram yordamchi.",
            Keyboards::settings());
        $this->tg->answerCallback($cbId);
    }

    // ------------------------------------------------------------------
    // ADMIN CALLBACKS
    // ------------------------------------------------------------------
    private function cbAdmin(int $uid, int $chatId, int $msgId, string $cbId, string $data): void {
        if (!$this->settings->isAdmin($uid)) { $this->tg->answerCallback($cbId); return; }
        $action = explode(':', $data, 2)[1] ?? '';

        if ($action === 'stats') {
            $stats = $this->usages->statsSummary();
            $total = $this->users->totalCount();
            $active = $this->users->activeCount();
            $b = $stats['breakdown'];
            $t = "📊 <b>Statistika:</b>\n\n" .
                 "👥 Foydalanuvchilar: <b>{$total}</b> (Faol: {$active})\n" .
                 "Bugungi so'rovlar: <b>{$stats['today_requests']}</b>\n" .
                 "Oylik so'rovlar: <b>{$stats['month_requests']}</b>\n" .
                 "🤖 Chat: " . ($b['chat'] ?? 0) . " | 🎙 Ovoz: " . ($b['voice'] ?? 0) . " | 📄 Hujjat: " . ($b['document'] ?? 0) . "\n" .
                 "🖼 Rasm: " . ($b['image'] ?? 0) . " | 💻 Kod: " . ($b['coding'] ?? 0) . " | 🛠 Tools: " . ($b['tool'] ?? 0) . "\n" .
                 "💰 Tokenlar: " . number_format($stats['total_tokens']);
            $this->tg->editMessageText($chatId, $msgId, $t, Keyboards::admin());
        } elseif ($action === 'users') {
            $total = $this->users->totalCount();
            $active = $this->users->activeCount(1);
            $blocked = $this->users->blockedCount();
            $recents = $this->users->getRecentUsers(8);

            $userList = '';
            foreach ($recents as $ru) {
                $uname = $ru['username'] ? '@' . $ru['username'] : '—';
                $name = htmlspecialchars(mb_substr($ru['first_name'] ?? 'User', 0, 15), ENT_QUOTES, 'UTF-8');
                $st = !empty($ru['is_blocked']) ? '🔴' : '🟢';
                $userList .= "• {$st} <b>{$name}</b> ({$uname}) — <code>{$ru['telegram_id']}</code>\n";
            }

            $t = "👥 <b>Foydalanuvchilar Boshqaruvi:</b>\n\n" .
                 "📊 Jami: <b>{$total}</b> ta | 24 soatda faol: <b>{$active}</b> ta | Bloklangan: <b>{$blocked}</b> ta\n\n" .
                 "<b>Oxirgi faol foydalanuvchilar:</b>\n{$userList}\n" .
                 "💡 <b>Tezkor buyruqlar:</b>\n" .
                 "• <code>/user &lt;id&gt;</code> — Foydalanuvchi ma'lumoti & amallarini PDF olish\n" .
                 "• <code>/block &lt;id&gt;</code> — Bloklash\n" .
                 "• <code>/unblock &lt;id&gt;</code> — Blokdan chiqarish\n" .
                 "• <code>/set_limit &lt;id&gt; &lt;soni&gt;</code> — Maxsus limit berish";

            $this->tg->editMessageText($chatId, $msgId, $t, Keyboards::admin());
        } elseif ($action === 'exports') {
            $this->tg->editMessageText($chatId, $msgId, "📥 <b>Hisobotlarni Eksport Qilish Markazi:</b>\n\nKerakli formatni tanlang 👇", Keyboards::adminExports());
        } elseif ($action === 'cleanup') {
            $this->cmdCleanup($chatId, $msgId);
        } elseif ($action === 'health') {
            $this->tg->editMessageText($chatId, $msgId,
                "⚙️ <b>Holat:</b> 🟢 Faol\nModel: {$this->settings->GEMINI_MODEL}\nAudio/Vision: 🟢 Faol (Gemini Native)",
                Keyboards::admin());
        } elseif ($action === 'sheets') {
            $configured = $this->sheets->isConfigured();
            $statusText = $configured ? "🟢 <b>Uланган (Faol)</b>" : "⚪ <b>Ulanmagan (Sozlash kutilmoqda)</b>";
            $count = $this->activityLogs->countTotal();
            $sheetUrl = !empty($this->settings->GOOGLE_SHEET_URL) ? $this->settings->GOOGLE_SHEET_URL : null;

            $text = "📈 <b>Google Sheets — Foydalanuvchilar Harakati</b>\n\n" .
                    "Ushbu bo'limda foydalanuvchilarning barcha so'rovlari, ovozlari, rasmlari va AI javoblari real vaqtda Google Jadvallariga tushib boradi.\n\n" .
                    "⚙️ <b>Holat:</b> {$statusText}\n" .
                    "📊 <b>Bazadagi jami harakatlar:</b> <b>{$count}</b> ta\n\n" .
                    ($configured 
                        ? "✅ Har bir harakat avtomatik ravishda jadvalga yozilmoqda." 
                        : "⚠️ <i>Eslatma:</i> Google Sheets webhook sozlanmagan bo'lsa ham, barcha harakatlar SQLite bazasida xavfsiz saqlanib boradi.");

            $this->tg->editMessageText($chatId, $msgId, $text, Keyboards::sheets($sheetUrl));
        } elseif ($action === 'back') {
            $this->tg->editMessageText($chatId, $msgId, "👑 <b>Admin Paneli:</b>", Keyboards::admin());
        } elseif ($action === 'broadcast') {
            $this->states->setState($uid, 'admin_broadcast');
            $this->tg->sendMessage($chatId, "📨 Yuboriladigan xabarni kiriting:", Keyboards::cancel());
        }
        $this->tg->answerCallback($cbId);
    }

    private function cbSheets(int $uid, int $chatId, int $msgId, string $cbId, string $data): void {
        if (!$this->settings->isAdmin($uid)) { $this->tg->answerCallback($cbId); return; }
        $action = explode(':', $data, 2)[1] ?? '';

        if ($action === 'ping') {
            if (!$this->sheets->isConfigured()) {
                $this->tg->answerCallback($cbId, "⚠️ Google Sheets Webhook sozlanmagan!", true);
                return;
            }
            $this->tg->answerCallback($cbId, "Sinov yuborilmoqda...");
            $ok = $this->sheets->log(
                $uid,
                'admin_test',
                'Admin Test',
                '🧪 Sinov (Ping)',
                'Google Sheets integratsiyasini sinovdan o\'tkazish',
                'Test muvaffaqiyatli qabul qilindi!',
                10,
                'OK'
            );
            if ($ok) {
                $this->tg->sendMessage($chatId, "✅ <b>Sinov yozuvi Google Sheets jadvaliga muvaffaqiyatli yuborildi!</b>\nJadvalingizni tekshirishingiz mumkin.");
            } else {
                $this->tg->sendMessage($chatId, "❌ <b>Sinov yozuvini yuborishda xatolik yuz berdi.</b>\nIltimos, Google Apps Script Webhook URL manzili to'g'ri sozlanganini tekshiring.");
            }
        } elseif ($action === 'sync') {
            if (!$this->sheets->isConfigured()) {
                $this->tg->answerCallback($cbId, "⚠️ Google Sheets Webhook sozlanmagan!", true);
                return;
            }
            $this->tg->answerCallback($cbId, "Sinxronlash boshlandi...");
            $logs = $this->activityLogs->getRecent(50);
            if (empty($logs)) {
                $this->tg->sendMessage($chatId, "ℹ️ Bazada hozircha yozuvlar mavjud emas.");
                return;
            }
            $synced = $this->sheets->syncBatch($logs);
            $this->tg->sendMessage($chatId, "📥 <b>Sinxronizatsiya yakunlandi!</b>\nGoogle Sheets ga <b>{$synced}</b> ta yozuv yuklandi.");
        } elseif ($action === 'guide') {
            $guide = "📋 <b>Google Sheets'ni ulash bo'yicha qo'llanma (1 daqiqa):</b>\n\n" .
                     "1️⃣ Yangi Google Sheet yarating (<code>sheets.new</code>).\n" .
                     "2️⃣ Menyudan <b>Extensions -> Apps Script</b> ni bosing.\n" .
                     "3️⃣ Loyihamizdagi <code>google_sheets_script.js</code> fayli ichidagi kodni nusxalab, Apps Script oynasiga qo'ying va saqlang.\n" .
                     "4️⃣ <b>Deploy -> New deployment</b> ni bosing:\n" .
                     "   • Type: <i>Web app</i>\n" .
                     "   • Execute as: <i>Me</i>\n" .
                     "   • Who has access: <i>Anyone</i>\n" .
                     "5️⃣ Olingan <b>Web App URL</b> manzilini <code>.env</code> faylidagi yoki Railway Variables dagi <code>GOOGLE_SHEETS_WEBHOOK_URL</code> ga qo'ying!\n\n" .
                     "Shuningdek, o'zingiz oson kirishingiz uchun jadval havolasini <code>GOOGLE_SHEET_URL</code> ga qo'yishingiz mumkin.";
            $sheetUrl = !empty($this->settings->GOOGLE_SHEET_URL) ? $this->settings->GOOGLE_SHEET_URL : null;
            $this->tg->editMessageText($chatId, $msgId, $guide, Keyboards::sheets($sheetUrl));
            $this->tg->answerCallback($cbId);
        } elseif ($action === 'set_webhook') {
            $this->states->setState($uid, 'admin_set_webhook');
            $this->tg->sendMessage($chatId,
                "🔗 <b>Google Apps Script Webhook URL manzilini yuboring:</b>\n\n" .
                "Masalan:\n<code>https://script.google.com/macros/s/AKfycb.../exec</code>\n\n" .
                "Bekor qilish uchun: <b>⬅️ Asosiy menyu</b>",
                Keyboards::cancel());
            $this->tg->answerCallback($cbId);
        }
    }

    private function handleSetWebhook(int $uid, int $chatId, string $text): void {
        if (!$this->settings->isAdmin($uid)) return;
        if ($text === '⬅️ Asosiy menyu') {
            $this->states->clear($uid);
            $this->tg->sendMessage($chatId, "Bekor qilindi.", Keyboards::main());
            return;
        }
        $url = trim($text);
        if (!str_starts_with($url, 'https://script.google.com/macros/s/')) {
            $this->tg->sendMessage($chatId, "❌ Noto'g'ri havola!\nManzil <code>https://script.google.com/macros/s/.../exec</code> shaklida bo'lishi kerak.\n\nQaytadan yuboring yoki bekor qilish uchun '⬅️ Asosiy menyu' ni bosing.");
            return;
        }

        $this->botSettings->set('GOOGLE_SHEETS_WEBHOOK_URL', $url);
        $this->settings->GOOGLE_SHEETS_WEBHOOK_URL = $url;
        $this->states->clear($uid);

        $statusMsg = $this->tg->sendMessage($chatId, "⏳ Webhook saqlandi, sinov qatori yuborilmoqda...");
        $statusId = $statusMsg['result']['message_id'] ?? 0;

        $ok = $this->sheets->log(
            $uid,
            'admin',
            'Admin',
            '🧪 Webhook Ulandi',
            'Google Sheets muvaffaqiyatli ulandi!',
            'Integratsiya faol va ishlayapti!',
            5,
            'OK'
        );

        if ($ok) {
            $sheetUrl = !empty($this->settings->GOOGLE_SHEET_URL) ? $this->settings->GOOGLE_SHEET_URL : null;
            $this->tg->editMessageText($chatId, $statusId,
                "✅ <b>Google Sheets Webhook muvaffaqiyatli sozlandi va faollashtirildi!</b>\n\n" .
                "Google Sheet jadvalingizga sinov qatori yozildi. Endi barcha amallar jadvalingizda avtomatik ko'rinib boradi.",
                Keyboards::sheets($sheetUrl));
        } else {
            $this->tg->editMessageText($chatId, $statusId,
                "⚠️ Webhook saqlandi, ammo Google Sheets test so'roviga javob bermadi.\n\n" .
                "Mumkin bo'lgan sabablar:\n" .
                "1. Apps Script da <b>Deploy -> New deployment</b> qilinganda <b>Who has access: Anyone</b> tanlanmagan.\n" .
                "2. URL to'liq nusxalanmagan.\n\n" .
                "Iltimos, tekshirib qaytadan kiriting.");
        }
    }

    private function handleBroadcast(int $uid, array $msg): void {
        if (!$this->settings->isAdmin($uid)) return;
        $chatId = (int)$msg['chat']['id'];
        if (($msg['text'] ?? '') === '⬅️ Asosiy menyu') {
            $this->states->clear($uid);
            $this->tg->sendMessage($chatId, "Bekor qilindi.", Keyboards::main());
            return;
        }

        $this->states->clear($uid);
        $statusMsg = $this->tg->sendMessage($chatId, "⏳ <i>Tarqatilmoqda...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;

        $ids = $this->users->getAllActiveIds();
        $delivered = 0; $failed = 0;

        foreach ($ids as $tid) {
            $r = $this->tg->call('copyMessage', [
                'chat_id' => $tid,
                'from_chat_id' => $chatId,
                'message_id' => $msg['message_id'],
            ]);
            if ($r['ok'] ?? false) $delivered++; else $failed++;
            usleep(50000);
        }

        $this->tg->editMessageText($chatId, $statusId,
            "✅ Yetkazildi: {$delivered} ta | Xato/Blok: {$failed} ta\nJami: " . count($ids) . " ta");
    }

    // ------------------------------------------------------------------
    // EKSPORT VA ADMIN BOSHQARUV HANDLERLARI
    // ------------------------------------------------------------------
    private function cbUserExport(int $uid, int $chatId, int $msgId, string $cbId, string $data, array $user): void {
        $action = explode(':', $data, 2)[1] ?? '';
        $uname = $user['username'] ? '@' . $user['username'] : ($user['first_name'] ?? 'Foydalanuvchi');

        if ($action === 'chat_pdf' || $action === 'chat_md') {
            $conv = $this->convs->getOrCreateActive($uid);
            $messages = $this->msgs->getAllForConversation((int)$conv['id']);
            if (empty($messages)) {
                $this->tg->answerCallback($cbId, "Faol suhbatda hali xabarlar yo'q.", true);
                return;
            }

            $this->tg->answerCallback($cbId, "Hujjat tayyorlanmoqda...");

            if ($action === 'chat_pdf') {
                $file = $this->exporter->exportUserMessagesPdf($uid, $uname, $messages, $conv['title'] ?? 'AI Suhbat');
                if ($file) {
                    $this->tg->sendDocument($chatId, $file, "📑 <b>\"{$conv['title']}\" suhbat hisoboti (PDF)</b>\n\nUniversal AI Assistant Bot tomonidan tayyorlandi.");
                    @unlink($file);
                } else {
                    $this->tg->sendMessage($chatId, "❌ PDF faylni shakllantirishda xatolik yuz berdi.");
                }
            } else {
                $file = $this->exporter->exportUserMessagesMd($uid, $uname, $messages, $conv['title'] ?? 'AI Suhbat');
                if ($file) {
                    $this->tg->sendDocument($chatId, $file, "📝 <b>\"{$conv['title']}\" suhbat hisoboti (Markdown)</b>\n\nUniversal AI Assistant Bot tomonidan tayyorlandi.");
                    @unlink($file);
                } else {
                    $this->tg->sendMessage($chatId, "❌ Markdown faylni shakllantirishda xatolik yuz berdi.");
                }
            }
        } elseif ($action === 'user_all_pdf' || $action === 'user_all_md') {
            $activities = $this->activityLogs->getUserActivities($uid, 100);
            if (empty($activities)) {
                $this->tg->answerCallback($cbId, "Hali birorta amal qayd etilmagan.", true);
                return;
            }

            $this->tg->answerCallback($cbId, "Hujjat tayyorlanmoqda...");

            if ($action === 'user_all_pdf') {
                $file = $this->exporter->exportUserActivitiesPdf($uid, $uname, $activities);
                if ($file) {
                    $this->tg->sendDocument($chatId, $file, "📊 <b>Sizning barcha amallaringiz hisoboti (PDF)</b>\n\nUniversal AI Assistant Bot tomonidan tayyorlandi.");
                    @unlink($file);
                } else {
                    $this->tg->sendMessage($chatId, "❌ PDF faylni shakllantirishda xatolik yuz berdi.");
                }
            } else {
                $file = $this->exporter->exportUserActivitiesMd($uid, $uname, $activities);
                if ($file) {
                    $this->tg->sendDocument($chatId, $file, "📝 <b>Sizning barcha amallaringiz hisoboti (Markdown)</b>\n\nUniversal AI Assistant Bot tomonidan tayyorlandi.");
                    @unlink($file);
                } else {
                    $this->tg->sendMessage($chatId, "❌ Markdown faylni shakllantirishda xatolik yuz berdi.");
                }
            }
        }
    }

    private function cbAdminExport(int $uid, int $chatId, int $msgId, string $cbId, string $data): void {
        if (!$this->settings->isAdmin($uid)) { $this->tg->answerCallback($cbId); return; }
        $action = explode(':', $data, 2)[1] ?? '';

        $this->tg->answerCallback($cbId, "Hisobot tayyorlanmoqda...");

        if ($action === 'all_pdf') {
            $activities = $this->activityLogs->getAllForAudit(1000);
            if (empty($activities)) {
                $this->tg->sendMessage($chatId, "ℹ️ Bazada yozuvlar mavjud emas.");
                return;
            }
            $file = $this->exporter->exportAdminAuditPdf($activities, "Barcha Foydalanuvchilar Harakatlari Audit Hisoboti");
            if ($file) {
                $this->tg->sendDocument($chatId, $file, "👑 <b>Barcha harakatlar hisoboti (PDF)</b>\n\nJami: " . count($activities) . " ta yozuv.");
                @unlink($file);
            }
        } elseif ($action === 'all_csv') {
            $activities = $this->activityLogs->getAllForAudit(1500);
            if (empty($activities)) {
                $this->tg->sendMessage($chatId, "ℹ️ Bazada yozuvlar mavjud emas.");
                return;
            }
            $file = $this->exporter->exportAdminAuditCsv($activities);
            if ($file) {
                $this->tg->sendDocument($chatId, $file, "📊 <b>Excel / CSV hisoboti</b>\n\nUshbu faylni Excel yoki Google Sheets'da to'g'ridan-to'g'ri ochishingiz mumkin.");
                @unlink($file);
            }
        } elseif ($action === 'all_md') {
            $activities = $this->activityLogs->getAllForAudit(1000);
            if (empty($activities)) {
                $this->tg->sendMessage($chatId, "ℹ️ Bazada yozuvlar mavjud emas.");
                return;
            }
            $file = $this->exporter->exportAdminAuditMd($activities, "Foydalanuvchilar Harakatlari Hisoboti");
            if ($file) {
                $this->tg->sendDocument($chatId, $file, "📝 <b>Markdown hisoboti (.md)</b>");
                @unlink($file);
            }
        } elseif ($action === 'today_pdf') {
            $activities = $this->activityLogs->getAllForAudit(500, true);
            if (empty($activities)) {
                $this->tg->sendMessage($chatId, "ℹ️ Bugun hali birorta harakat qayd etilmagan.");
                return;
            }
            $file = $this->exporter->exportAdminAuditPdf($activities, "Bugungi Harakatlar Audit Hisoboti", true);
            if ($file) {
                $this->tg->sendDocument($chatId, $file, "📅 <b>Bugungi kunlik audit hisoboti (PDF)</b>\n\nJami: " . count($activities) . " ta yozuv.");
                @unlink($file);
            }
        }
    }

    private function cmdCleanup(int $chatId, ?int $msgId = null): void {
        $deleted = 0;
        $freedBytes = 0;
        foreach ([$this->settings->TEMP_DIR, $this->settings->DOWNLOADS_DIR] as $dir) {
            if (is_dir($dir)) {
                foreach (glob($dir . '/*') as $f) {
                    if (is_file($f)) {
                        $freedBytes += filesize($f);
                        @unlink($f);
                        $deleted++;
                    }
                }
            }
        }
        $freedMb = round($freedBytes / (1024 * 1024), 2);
        $text = "🧹 <b>Tizim keshini tozalash yakunlandi!</b>\n\n" .
                "🗑 O'chirilgan vaqtinchalik fayllar: <b>{$deleted}</b> ta\n" .
                "💾 Bo'shatilgan joy: <b>{$freedMb} MB</b>";
        if ($msgId) {
            $this->tg->editMessageText($chatId, $msgId, $text, Keyboards::admin());
        } else {
            $this->tg->sendMessage($chatId, $text, Keyboards::admin());
        }
    }

    private function cmdUserInspect(int $chatId, int $targetId): void {
        $u = $this->users->getByTelegramId($targetId);
        if (!$u) {
            $this->tg->sendMessage($chatId, "❌ Foydalanuvchi `{$targetId}` topilmadi.");
            return;
        }

        $activities = $this->activityLogs->getUserActivities($targetId, 50);
        $actCount = count($activities);
        $status = !empty($u['is_blocked']) ? "🔴 Bloklangan" : "🟢 Faol";
        $limit = $u['daily_limit_override'] !== null ? "{$u['daily_limit_override']} ta (maxsus)" : "Standart ({$this->settings->FREE_DAILY_LIMIT})";

        $text = "👤 <b>Foydalanuvchi Ma'lumotlari:</b>\n\n" .
                "🆔 Telegram ID: <code>{$u['telegram_id']}</code>\n" .
                "👤 Ismi: <b>" . htmlspecialchars($u['first_name'] ?? '—', ENT_QUOTES, 'UTF-8') . "</b>\n" .
                "🌐 Username: @" . ($u['username'] ?? '—') . "\n" .
                "🛡 Holati: {$status}\n" .
                "⚡ Kunlik limit: <b>{$limit}</b>\n" .
                "📅 Ro'yxatdan o'tgan: {$u['created_at']}\n" .
                "🕒 Oxirgi faollik: {$u['last_active']}\n" .
                "📊 Jami amallari: <b>{$actCount}</b> ta\n\n" .
                "<b>Boshqaruv buyruqlari:</b>\n" .
                "• <code>/block {$targetId}</code> — Bloklash\n" .
                "• <code>/unblock {$targetId}</code> — Blokdan chiqarish\n" .
                "• <code>/set_limit {$targetId} 50</code> — Limit belgilash";

        $this->tg->sendMessage($chatId, $text);

        if (!empty($activities)) {
            $name = $u['username'] ? '@' . $u['username'] : ($u['first_name'] ?? 'User');
            $file = $this->exporter->exportUserActivitiesPdf($targetId, $name, $activities);
            if ($file) {
                $this->tg->sendDocument($chatId, $file, "📑 <b>Foydalanuvchi `{$targetId}` amallari to'liq hisoboti (PDF)</b>");
                @unlink($file);
            }
        }
    }

    // ------------------------------------------------------------------
    // UMUMIY CHAT
    // ------------------------------------------------------------------
    private function handleGeneralChat(int $uid, int $chatId, string $text, ?array $user = null): void {
        [$ok, $used, $limit] = $this->usages->canRequest($uid);
        if (!$ok) { $this->tg->sendMessage($chatId, "⚠️ Bugungi bepul limit tugadi ({$used}/{$limit})."); return; }

        $this->tg->sendChatAction($chatId, 'typing');

        $conv = $this->convs->getOrCreateActive($uid);
        $history = $this->msgs->recent((int)$conv['id'], 8);
        $payload = $history;
        $payload[] = ['role' => 'user', 'content' => $text];

        $uname = $user['username'] ?? null;
        $fname = trim(($user['first_name'] ?? '') . ' ' . ($user['last_name'] ?? ''));

        try {
            $res = $this->gemini->chat($payload);
            $this->msgs->add((int)$conv['id'], 'user', $text);
            $this->msgs->add((int)$conv['id'], 'assistant', $res['content']);
            $this->convs->updateTitleIfDefault((int)$conv['id'], $text);
            $this->usages->record($uid, 'chat', $res['tokens']);

            $this->logAction($uid, $uname, $fname, '🤖 AI Chat', $text, $res['content'], $res['tokens'], 'OK');

            foreach (Helpers::splitText($res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk, null, 'Markdown');
            }
        } catch (Throwable $e) {
            $this->logger->error("General chat error: " . $e->getMessage());
            $this->logAction($uid, $uname, $fname, '🤖 AI Chat', $text, $e->getMessage(), 0, 'ERROR');
            $this->tg->sendMessage($chatId, "❌ Xizmatda texnik muammo yuz berdi.");
        }
    }
}

// ============================================================================
// 10. ISHGA TUSHIRISH
// ============================================================================

try {
    $pdo = Database::pdo($settings->DATABASE_PATH);

    // Jadvallar mavjudligini tekshirish
    $tableCheck = $pdo->query("SELECT name FROM sqlite_master WHERE type='table' AND name='users'")->fetchColumn();
    if (!$tableCheck) {
        $logger->error("❌ Ma'lumotlar bazasi jadvallari topilmadi. Avval 'php install.php' buyrug'ini ishga tushiring.");
        exit(1);
    }

    if (empty($settings->BOT_TOKEN) || $settings->BOT_TOKEN === 'YOUR_TELEGRAM_BOT_TOKEN_HERE') {
        $logger->error("❌ BOT_TOKEN .env faylda sozlanmagan.");
        exit(1);
    }
    if (empty($settings->GEMINI_API_KEY) || $settings->GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
        $logger->error("❌ GEMINI_API_KEY .env faylda sozlanmagan.");
        exit(1);
    }

    $tg = new TelegramAPI($settings->BOT_TOKEN);
    $bot = new TelegramBot($settings, $pdo, $tg, $logger);
    $bot->run();
} catch (Throwable $e) {
    fwrite(STDERR, "❌ FATAL: " . $e->getMessage() . "\n");
    exit(1);
}