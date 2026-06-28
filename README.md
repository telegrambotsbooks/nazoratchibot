# Savdo Nazorat Bot — Railway Ready

Bu versiya Railway uchun moslangan.

## Muhim
- Railway'da PostgreSQL qo'shilsa `DATABASE_URL` avtomatik chiqadi.
- Bot barcha ma'lumotlarni PostgreSQL bazaga saqlaydi.
- Railway qayta deploy bo'lsa ham ma'lumotlar o'chib ketmaydi.
- `bot.db` SQLite faqat local test uchun ishlaydi. Railway'da PostgreSQL ishlatish shart.

## Railway Variables
Railway → Variables ichiga yozing:

```env
BOT_TOKEN=telegram_bot_token
SELLER_LOGIN=sotuvchi
SELLER_PASSWORD=12345
CONTROLLER_LOGIN=nazoratchi
CONTROLLER_PASSWORD=54321
```

`DATABASE_URL` ni qo'lda yozmang. PostgreSQL qo'shilganda Railway o'zi beradi.

## Hisobotlar
- Kunlik Excel — faqat bugungi yozuvlarni chiqaradi.
- Oylik Excel — shu oyning yozuvlarini chiqaradi.
- Umumiy Excel — barcha yozuvlarni chiqaradi.
- Suv berish sotuvga qo'shilmaydi, lekin hisobotga chiqadi.
- Qarzdorlik PostgreSQL'da `payments` jadvalida saqlanadi.

## Ishga tushirish local
```bash
pip install -r requirements.txt
python main.py
```
