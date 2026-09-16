<?php
/**
 * ============================================================================
 * UNIVERSAL TELEGRAM AI ASSISTANT BOT — ALL-IN-ONE (PHP 8.1+)
 * ============================================================================
 * Texnologiyalar:
 *  - PHP 8.1+
 *  - GuzzleHTTP (Telegram & OpenAI)
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

// ============================================================================
// 0. PROMPTS
// ============================================================================

class Prompts {
    public const SYSTEM = <<<'EOT'
You are a professional multilingual AI assistant.
Your goal is to provide accurate, useful, clear and practical answers.
Never fabricate facts. If uncertain, state it clearly.
Adapt to user's language (Uzbek, Russian, English).
For programming, provide clean, runnable code.
Be concise when simple, detailed when depth is needed.
EOT;

    public const DOCUMENT = <<<'EOT'
You are an expert Document Analyst AI Assistant.
Analyze documents thoroughly: core message, structure, facts, numbers, dates.
Distinguish facts from interpretation. Identify risks and contradictions.
Provide actionable recommendations. Never fabricate.
Adapt to user's language. Format cleanly for Telegram.
EOT;

    public const CODING = <<<'EOT'
You are a Principal Software Engineer and Code Architect.
Provide production-grade, maintainable, secure code.
Include complete runnable code blocks with syntax highlighting.
For debugging: root cause, why, fix, prevention.
Respond in user's language (Uzbek, Russian, English).
EOT;

    public const TOOLS = <<<'EOT'
You are a specialized AI Multi-Tool Engine.
Execute tasks with maximum quality: text generation, translation,
summarization, ideas, rewriting, marketing, study assistance, data analysis.
Adapt to user's language.
EOT;
}


// ============================================================================
// 1. SOZLAMALAR
// ============================================================================

$dotenv = Dotenv::createImmutable(__DIR__);
if (file_exists(__DIR__ . '/.env')) $dotenv->load();

class Settings {
    public string $BOT_TOKEN;
    public int $ADMIN_ID;
    public string $OPENAI_API_KEY;
    public string $OPENAI_MODEL;
    public string $OPENAI_WHISPER_MODEL;
    public float $OPENAI_TEMPERATURE;
    public string $DATABASE_PATH;
    public int $FREE_DAILY_LIMIT;
    public int $MAX_FILE_SIZE_MB;
    public string $TEMP_DIR, $UPLOADS_DIR, $DOWNLOADS_DIR, $DATA_DIR, $LOGS_DIR, $LOG_LEVEL;

    public function __construct() {
        $this->BOT_TOKEN            = $_ENV['BOT_TOKEN']            ?? '';
        $this->ADMIN_ID             = (int)($_ENV['ADMIN_ID']        ?? 0);
        $this->OPENAI_API_KEY       = $_ENV['OPENAI_API_KEY']        ?? '';
        $this->OPENAI_MODEL         = $_ENV['OPENAI_MODEL']          ?? 'gpt-4o-mini';
        $this->OPENAI_WHISPER_MODEL = $_ENV['OPENAI_WHISPER_MODEL']  ?? 'whisper-1';
        $this->OPENAI_TEMPERATURE   = (float)($_ENV['OPENAI_TEMPERATURE'] ?? 0.7);
        $this->DATABASE_PATH        = $_ENV['DATABASE_PATH']         ?? './data/bot.db';
        $this->FREE_DAILY_LIMIT     = (int)($_ENV['FREE_DAILY_LIMIT'] ?? 30);
        $this->MAX_FILE_SIZE_MB     = (int)($_ENV['MAX_FILE_SIZE_MB'] ?? 20);
        $this->TEMP_DIR             = $_ENV['TEMP_DIR']              ?? './temp';
        $this->UPLOADS_DIR          = $_ENV['UPLOADS_DIR']           ?? './uploads';
        $this->DOWNLOADS_DIR        = $_ENV['DOWNLOADS_DIR']         ?? './downloads';
        $this->DATA_DIR             = $_ENV['DATA_DIR']              ?? './data';
        $this->LOGS_DIR             = $_ENV['LOGS_DIR']              ?? './logs';
        $this->LOG_LEVEL            = $_ENV['LOG_LEVEL']             ?? 'INFO';
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
}

class OpenAIService {
    private Client $http;
    public function __construct(private Settings $settings) {
        $this->http = new Client([
            'base_uri' => 'https://api.openai.com/v1/',
            'timeout' => 120,
            'headers' => [
                'Authorization' => 'Bearer ' . $settings->OPENAI_API_KEY,
                'Content-Type' => 'application/json',
            ],
        ]);
    }

    private function post(string $path, array $body): array {
        try {
            $r = $this->http->post($path, ['json' => $body]);
            return json_decode((string)$r->getBody(), true) ?? [];
        } catch (GuzzleException $e) {
            throw new RuntimeException('OpenAI xatosi: ' . $e->getMessage());
        }
    }

    public function chat(array $messages, ?string $systemPrompt = null, ?float $temp = null): array {
        $full = [['role' => 'system', 'content' => $systemPrompt ?? Prompts::SYSTEM]];
        foreach ($messages as $m) $full[] = $m;

        $data = $this->post('chat/completions', [
            'model' => $this->settings->OPENAI_MODEL,
            'messages' => $full,
            'temperature' => $temp ?? $this->settings->OPENAI_TEMPERATURE,
        ]);
        return [
            'content' => $data['choices'][0]['message']['content'] ?? '',
            'tokens'  => $data['usage']['total_tokens'] ?? 0,
        ];
    }

    public function transcribe(string $audioPath): string {
        try {
            $r = $this->http->post('audio/transcriptions', [
                'multipart' => [
                    ['name' => 'file', 'contents' => fopen($audioPath, 'r')],
                    ['name' => 'model', 'contents' => $this->settings->OPENAI_WHISPER_MODEL],
                ],
            ]);
            $d = json_decode((string)$r->getBody(), true);
            return trim($d['text'] ?? '');
        } catch (GuzzleException $e) {
            throw new RuntimeException('Whisper xatosi: ' . $e->getMessage());
        }
    }

    public function analyzeImage(string $base64, string $prompt): array {
        $messages = [
            ['role' => 'system', 'content' => Prompts::SYSTEM],
            ['role' => 'user', 'content' => [
                ['type' => 'text', 'text' => $prompt],
                ['type' => 'image_url', 'image_url' => ['url' => "data:image/jpeg;base64,{$base64}"]],
            ]],
        ];
        $data = $this->post('chat/completions', [
            'model' => $this->settings->OPENAI_MODEL,
            'messages' => $messages,
            'temperature' => $this->settings->OPENAI_TEMPERATURE,
        ]);
        return [
            'content' => $data['choices'][0]['message']['content'] ?? '',
            'tokens'  => $data['usage']['total_tokens'] ?? 0,
        ];
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
        private OpenAIService $openai
    ) {}

    public function processVoice(string $fileId, string $ext = 'ogg'): array {
        $tempPath = rtrim($_ENV['TEMP_DIR'] ?? './temp', '/') . '/voice_' . Helpers::uuid() . '.' . $ext;
        try {
            $filePath = $this->tg->getFile($fileId);
            if (!$filePath) return [null, "Ovozli faylni yuklab bo'lmadi."];
            if (!$this->tg->downloadFile($filePath, $tempPath))
                return [null, "Ovozli faylni saqlab bo'lmadi."];

            $text = $this->openai->transcribe($tempPath);
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
        private OpenAIService $openai
    ) {}

    public function processPhoto(string $fileId, string $prompt): array {
        $tempPath = rtrim($_ENV['TEMP_DIR'] ?? './temp', '/') . '/photo_' . Helpers::uuid() . '.jpg';
        try {
            $filePath = $this->tg->getFile($fileId);
            if (!$filePath) throw new RuntimeException('Rasmni yuklab bo\'lmadi.');
            if (!$this->tg->downloadFile($filePath, $tempPath))
                throw new RuntimeException('Rasmni saqlab bo\'lmadi.');
            $b64 = base64_encode(file_get_contents($tempPath));
            return $this->openai->analyzeImage($b64, $prompt);
        } finally {
            @unlink($tempPath);
        }
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
            [['text' => '🗑 Tarixni tozalash', 'callback_data' => 'settings:clear_history'],
             ['text' => 'ℹ️ Bot haqida', 'callback_data' => 'settings:about']],
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
        return ['inline_keyboard' => $rows];
    }

    public static function admin(): array {
        return ['inline_keyboard' => [
            [['text' => '📊 Statistika', 'callback_data' => 'admin:stats'],
             ['text' => '📨 Xabar tarqatish', 'callback_data' => 'admin:broadcast']],
            [['text' => '👥 Foydalanuvchilar', 'callback_data' => 'admin:users'],
             ['text' => '⚙️ Tizim holati', 'callback_data' => 'admin:health']],
        ]];
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
    private OpenAIService $openai;
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

        $this->openai        = new OpenAIService($settings);
        $this->docService    = new DocumentService();
        $this->speechService = new SpeechService($tg, $this->openai);
        $this->imageService  = new ImageService($tg, $this->openai);

        $this->states    = new StateManager($settings->DATA_DIR . '/states');
        $this->rateLimit = new RateLimiter(0.8);
    }

    // ------------------------------------------------------------------
    // ASOSIY POLLING LOOP
    // ------------------------------------------------------------------
    public function run(): void {
        $this->tg->call('deleteWebhook', ['drop_pending_updates' => true]);
        $this->registerCommands();

        $me = $this->tg->call('getMe');
        $username = $me['result']['username'] ?? 'unknown';
        $this->logger->info("Bot started: @{$username} | Model: {$this->settings->OPENAI_MODEL}");

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
        $this->users->getOrCreate($uid, $user['username'] ?? null, $user['first_name'] ?? null,
                                  $user['language_code'] ?? 'uz');

        if ($this->rateLimit->check($uid) && !$this->settings->isAdmin($uid)) {
            $this->tg->sendMessage($chatId, "⚠️ Iltimos, biroz kuting.");
            return;
        }

        // Buyruqlar
        $text = $msg['text'] ?? '';
        if ($text !== '') {
            if ($text === '/start' || $text === '/help')    { $this->cmdStart($chatId, $text === '/help'); return; }
            if ($text === '/newchat')                       { $this->cmdNewChat($uid, $chatId); return; }
            if ($text === '/history')                       { $this->cmdHistory($uid, $chatId); return; }
            if ($text === '/settings' || $text === '⚙️ Sozlamalar') { $this->cmdSettings($chatId); return; }
            if ($text === '/admin' && $this->settings->isAdmin($uid)) { $this->cmdAdmin($chatId); return; }

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
                    $this->cmdNewChat($uid, $chatId);
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

        if ($state === 'doc_question' && $text !== '') { $this->handleDocQuestion($uid, $chatId, $text); return; }
        if ($state === 'coding_wait' && $text !== '')  { $this->handleCodingInput($uid, $chatId, $text); return; }
        if ($state === 'tool_wait' && $text !== '')    { $this->handleToolInput($uid, $chatId, $text); return; }
        if ($state === 'admin_broadcast')              { $this->handleBroadcast($uid, $msg); return; }

        // Fayllar va media
        if (isset($msg['voice']) || isset($msg['audio'])) { $this->handleVoice($uid, $chatId, $msg); return; }
        if (isset($msg['photo']))                          { $this->handlePhoto($uid, $chatId, $msg); return; }
        if (isset($msg['document']))                       { $this->handleDocument($uid, $chatId, $msg); return; }

        // Umumiy chat
        if ($text !== '' && !str_starts_with($text, '/')) {
            $this->handleGeneralChat($uid, $chatId, $text);
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
        if (str_starts_with($data, 'doc_action:'))  { $this->cbDocAction($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'code_mode:'))   { $this->cbCodeMode($uid, $chatId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'tool:'))        { $this->cbTool($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'tr_lang:'))     { $this->cbTrLang($uid, $chatId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'style:'))       { $this->cbStyle($uid, $chatId, $cb['id'], $data); return; }
        if (str_starts_with($data, 'switch_conv:')) { $this->cbSwitchConv($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if ($data === 'settings:language')          { $this->tg->editMessageText($chatId, $msgId, "🌐 Tilni tanlang:", Keyboards::languageSelect()); $this->tg->answerCallback($cb['id']); return; }
        if (str_starts_with($data, 'lang:'))        { $this->cbLang($uid, $chatId, $msgId, $cb['id'], $data); return; }
        if ($data === 'settings:model')             { $this->cbModel($chatId, $msgId, $cb['id']); return; }
        if ($data === 'settings:clear_history')     { $this->cbClearHistory($uid, $chatId, $msgId, $cb['id']); return; }
        if ($data === 'settings:about')             { $this->cbAbout($chatId, $msgId, $cb['id']); return; }

        // Admin
        if (str_starts_with($data, 'admin:')) { $this->cbAdmin($uid, $chatId, $msgId, $cb['id'], $data); return; }

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

    private function cmdNewChat(int $uid, int $chatId): void {
        $this->states->clear($uid);
        $this->convs->createNew($uid);
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
            $res = $this->openai->chat($payload);
            $this->msgs->add((int)$conv['id'], 'user', "[Voice]: {$transcription}");
            $this->msgs->add((int)$conv['id'], 'assistant', $res['content']);
            $this->convs->updateTitleIfDefault((int)$conv['id'], $transcription);
            $this->usages->record($uid, 'voice', $res['tokens']);

            $this->tg->editMessageText($chatId, $statusId,
                "📝 <b>Transkripsiya:</b>\n<i>\"{$transcription}\"</i>");

            foreach (Helpers::splitText("🤖 <b>AI:</b>\n\n" . $res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk);
            }
        } catch (Throwable $e) {
            $this->logger->error("Voice response error: " . $e->getMessage());
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

            $prefix = isset($msg['caption']) ? "📝 <i>Izoh: \"{$msg['caption']}\"</i>\n\n" : "";
            foreach (Helpers::splitText("🖼 <b>Rasm tahlili:</b>\n\n{$prefix}{$res['content']}") as $chunk) {
                $this->tg->sendMessage($chatId, $chunk);
            }
        } catch (Throwable $e) {
            $this->logger->error("Image error: " . $e->getMessage());
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
    private function cbDocAction(int $uid, int $chatId, int $msgId, string $cbId, string $data): void {
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

        try {
            $res = $this->openai->analyzeDocument($docText, $action);
            $this->usages->record($uid, 'document', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);

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
            $this->tg->editMessageText($chatId, $statusId, "❌ Tahlilda xatolik yuz berdi.");
        }
    }

    private function handleDocQuestion(int $uid, int $chatId, string $question): void {
        $data = $this->states->getData($uid);
        $docText = $data['doc_text'] ?? null;
        if (!$docText) {
            $this->tg->sendMessage($chatId, "Hujjat topilmadi.");
            $this->states->clear($uid);
            return;
        }

        $statusMsg = $this->tg->sendMessage($chatId, "🤔 <i>Hujjatdan javob qidirilmoqda...</i>");
        $statusId = $statusMsg['result']['message_id'] ?? 0;

        try {
            $res = $this->openai->analyzeDocument($docText, 'question', $question);
            $this->usages->record($uid, 'document', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);
            foreach (Helpers::splitText("❓ <b>Savol:</b> <i>{$question}</i>\n\n📌 <b>Javob:</b>\n\n" . $res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk);
            }
        } catch (Throwable $e) {
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

    private function handleCodingInput(int $uid, int $chatId, string $text): void {
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

        try {
            $res = $this->openai->coding($text, $mode);
            $this->usages->record($uid, 'coding', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);
            foreach (Helpers::splitText($res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk, null, 'Markdown');
            }
        } catch (Throwable $e) {
            $this->logger->error("Coding error: " . $e->getMessage());
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

    private function handleToolInput(int $uid, int $chatId, string $text): void {
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

        try {
            $res = $this->openai->runTool($tool, $text, $params);
            $this->usages->record($uid, 'tool', $res['tokens']);
            $this->tg->deleteMessage($chatId, $statusId);
            foreach (Helpers::splitText($res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk, null, 'Markdown');
            }
        } catch (Throwable $e) {
            $this->logger->error("Tool error: " . $e->getMessage());
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
        $t = "🧠 <b>Model:</b> {$this->settings->OPENAI_MODEL}\n" .
             "🎙 <b>Whisper:</b> {$this->settings->OPENAI_WHISPER_MODEL}\n" .
             "Limit: {$this->settings->FREE_DAILY_LIMIT} ta/kun";
        $this->tg->editMessageText($chatId, $msgId, $t, Keyboards::settings());
        $this->tg->answerCallback($cbId);
    }

    private function cbClearHistory(int $uid, int $chatId, int $msgId, string $cbId): void {
        $this->convs->createNew($uid, 'Yangi suhbat');
        $this->tg->answerCallback($cbId, "Tarix tozalandi!", true);
        $this->tg->editMessageText($chatId, $msgId, "🗑 Faol suhbat tozalandi!", Keyboards::settings());
    }

    private function cbAbout(int $chatId, int $msgId, string $cbId): void {
        $this->tg->editMessageText($chatId, $msgId,
            "ℹ️ <b>Universal AI Assistant Bot</b>\n\nGPT-4o & PHP 8.1 asosida yaratilgan professional Telegram yordamchi.",
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
            $this->tg->editMessageText($chatId, $msgId, "👥 Jami: {$total} ta | 24 soatda faol: {$active} ta", Keyboards::admin());
        } elseif ($action === 'health') {
            $this->tg->editMessageText($chatId, $msgId,
                "⚙️ <b>Holat:</b> 🟢 Faol\nModel: {$this->settings->OPENAI_MODEL}\nWhisper: {$this->settings->OPENAI_WHISPER_MODEL}",
                Keyboards::admin());
        } elseif ($action === 'broadcast') {
            $this->states->setState($uid, 'admin_broadcast');
            $this->tg->sendMessage($chatId, "📨 Yuboriladigan xabarni kiriting:", Keyboards::cancel());
        }
        $this->tg->answerCallback($cbId);
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
    // UMUMIY CHAT
    // ------------------------------------------------------------------
    private function handleGeneralChat(int $uid, int $chatId, string $text): void {
        [$ok, $used, $limit] = $this->usages->canRequest($uid);
        if (!$ok) { $this->tg->sendMessage($chatId, "⚠️ Bugungi bepul limit tugadi ({$used}/{$limit})."); return; }

        $this->tg->sendChatAction($chatId, 'typing');

        $conv = $this->convs->getOrCreateActive($uid);
        $history = $this->msgs->recent((int)$conv['id'], 8);
        $payload = $history;
        $payload[] = ['role' => 'user', 'content' => $text];

        try {
            $res = $this->openai->chat($payload);
            $this->msgs->add((int)$conv['id'], 'user', $text);
            $this->msgs->add((int)$conv['id'], 'assistant', $res['content']);
            $this->convs->updateTitleIfDefault((int)$conv['id'], $text);
            $this->usages->record($uid, 'chat', $res['tokens']);

            foreach (Helpers::splitText($res['content']) as $chunk) {
                $this->tg->sendMessage($chatId, $chunk, null, 'Markdown');
            }
        } catch (Throwable $e) {
            $this->logger->error("General chat error: " . $e->getMessage());
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
    if (empty($settings->OPENAI_API_KEY) || $settings->OPENAI_API_KEY === 'YOUR_OPENAI_API_KEY_HERE') {
        $logger->error("❌ OPENAI_API_KEY .env faylda sozlanmagan.");
        exit(1);
    }

    $tg = new TelegramAPI($settings->BOT_TOKEN);
    $bot = new TelegramBot($settings, $pdo, $tg, $logger);
    $bot->run();
} catch (Throwable $e) {
    fwrite(STDERR, "❌ FATAL: " . $e->getMessage() . "\n");
    exit(1);
}