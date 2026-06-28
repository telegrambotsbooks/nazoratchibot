# Savdo Nazorat Telegram Bot — Railway tayyor

## Local ishga tushirish

```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
python main.py
```

`.env` ichiga `BOT_TOKEN` yozing. Localda `DATABASE_URL` yozilmasa bot `bot.db` SQLite bilan ishlaydi.

## Railway deploy

1. Papkani GitHub'ga push qiling.
2. Railway'da **New Project → Deploy from GitHub repo** qiling.
3. Railway'da **PostgreSQL** qo'shing.
4. Bot service Variables ichiga quyidagilarni yozing:

```env
BOT_TOKEN=BotFatherdan_olingan_token
SELLER_LOGIN=sotuvchi
SELLER_PASSWORD=12345
CONTROLLER_LOGIN=nazoratchi
CONTROLLER_PASSWORD=54321
```

5. PostgreSQL qo'shilgandan keyin Railway `DATABASE_URL` ni o'zi beradi. Bot avtomatik PostgreSQL ishlatadi.
6. Deploy qiling.

Start command: `python main.py`

## Muhim

- `.env` faylni GitHub'ga chiqarmang.
- Tokenni chatga yoki GitHub'ga tashlamang.
- Railway'da ma'lumotlar PostgreSQL'da saqlanadi.
