# Universal Telegram AI Bot (PHP + Google Gemini)

Ko'p funksiyali, 100% **Google Gemini API** orqali ishlaydigan zamonaviy Telegram AI yordamchi bot (PHP 8.2+).

---

## 🚀 Imkoniyatlar

- 🤖 **100% Google Gemini API**: Eng so'nggi va tezkor `gemini-2.5-flash` modeli orqali chat, savol-javob va tavsiyalar.
- 🎙 **Ovozli xabarlar (Gemini Native Audio)**: Whisper o'rniga Gemini'ning ichki multimodal ovoz tahlili orqali ovozli xabarlarni matnga o'girish va ovozga javob qaytarish.
- 🖼 **Rasm tahlili (Vision)**: Foydalanuvchi yuborgan rasmlarni Gemini Vision orqali batafsil tahlil qilish.
- 📄 **Hujjatlar tahlili**: PDF, Word (DOCX), Excel (XLSX), PowerPoint (PPTX) va matnli fayllardan ma'lumotlarni o'qish, xulosa qilish va savollarga javob berish.
- 💻 **Dasturlash yordamchisi**: Kod yozish, xatolarni tuzatish, refaktoring va testlar yaratish.
- 🛠 **AI Tools**: Professional matn yozish, tarjimon, xulosalash, g'oyalar yaratish, marketing postlari va tahlil.
- 🗄 **SQLite & Limitlar**: Foydalanuvchilar, suhbatlar tarixi va kunlik bepul limitlar boshqaruvi.

---

## 🛠 Mahalliy O'rnatish (Local Setup)

1. **Repozitoriyani klon qiling va papkaga kiring:**
   ```bash
   git clone <repo-url>
   cd telegram-ai-bot
   ```

2. **Bog'liqliklarni o'rnating:**
   ```bash
   composer install
   ```

3. **Muhit sozlamalari (`.env`):**
   ```bash
   cp .env.example .env
   ```
   `.env` faylini ochib, quyidagilarni kiriting:
   - `BOT_TOKEN` — @BotFather'dan olingan bot tokeni
   - `ADMIN_ID` — Administrator Telegram ID raqami
   - `GEMINI_API_KEY` — Google AI Studio'dan olingan Gemini API kaliti
   - `GEMINI_MODEL=gemini-2.5-flash`

4. **Ma'lumotlar bazasini yaratish va botni ishga tushirish:**
   ```bash
   php install.php
   php bot.php
   ```

---

## 🚂 Railway Platformasiga Joylash (Railway Deployment)

Ushbu loyiha Railway uchun tayyor `Dockerfile` va `railway.json` konfiguratsiyasiga ega.

### 1-qadam: GitHub orqali ulash
1. Loyihani GitHub repozitoriyangizga push qiling.
2. [Railway.app](https://railway.app) ga kiring va **"New Project"** -> **"Deploy from GitHub repo"** ni tanlang.

### 2-qadam: Railway Variables sozlash
Railway boshqaruv panelida loyihangizning **Variables** (muhit o'zgaruvchilari) bo'limiga kiring va quyidagi kalitlarni qo'shing:

| Kalit (Variable) | Tavsif | Namuna qiymat |
| :--- | :--- | :--- |
| `BOT_TOKEN` | Telegram Bot Tokeni (@BotFather) | `8284934602:AA...` |
| `ADMIN_ID` | Telegram hisobingiz ID si | `8536958944` |
| `GEMINI_API_KEY` | Google Gemini API kaliti | `AIzaSy...` |
| `GEMINI_MODEL` | Gemini modeli | `gemini-2.5-flash` |
| `GEMINI_TEMPERATURE` | Kreativlik darajasi | `0.7` |
| `FREE_DAILY_LIMIT` | Bepul kunlik limit | `30` |

### 3-qadam: Ma'lumotlar saqlanishi (Railway Volume)
SQLite bazasi (`data/bot.db`) deploylar vaqtida o'chib ketmasligi uchun:
- Railway loyihangizga **"Add Volume"** bosing.
- Mount path qilib `/app/data` ni belgilang.

### 4-qadam: Deploy
Railway avtomatik tarzda `Dockerfile` orqali konteynerni yig'adi, `install.php` ni ishga tushiradi va botni uzluksiz fonda yurgizadi.
