# Universal Telegram AI Bot (PHP)

Ko'p funksiyali Telegram AI yordamchi bot (PHP + OpenAI).

## Imkoniyatlar

- Telegram bot orqali OpenAI GPT modellari bilan muloqot
- Ovozli xabarlarni Whisper orqali matnga o'girish va javob berish
- Hujjatlar bilan ishlash (PDF, Word, Excel, PowerPoint)
- SQLite ma'lumotlar bazasi orqali foydalanuvchilar va kunlik limitlar boshqaruvi

## O'rnatish

1. Repozitoriyani klon qiling:
```bash
git clone <repo-url>
cd telegram-ai-bot
```

2. Bog'liqliklarni o'rnating:
```bash
composer install
```

3. Konfiguratsiya faylini tayyorlang:
```bash
cp .env.example .env
```
`.env` fayliga o'zingizning Telegram `BOT_TOKEN` va `OPENAI_API_KEY` qiymatlaringizni kiriting.

4. Botni sozlash va ishga tushirish:
```bash
php install.php
php bot.php
```
