# Davomiylik — SVOD TIZIMI

CSV fayllarni tekshiradigan va Excel svod hisobotini yaratadigan dastur. Windows'da lokal, Vercel'da veb rejimda ishlaydi.

## Imkoniyatlar

- Bir yoki bir nechta CSV faylni qabul qiladi
- Majburiy ustunlar va sana formatlarini tekshiradi
- Realizatsiyalar orasidagi kun farqini hisoblaydi
- Dublikat abonent kodlarini alohida varaqda ko‘rsatadi
- Raygaz va mahalla kesimida Excel hisoboti yaratadi
- Lokal rejimda ma’lumotlarni faqat kompyuterning o‘zida qayta ishlaydi

## Windows'da lokal ishga tushirish

Windows’da `ISHGA_TUSHIRISH.bat` faylini ikki marta bosing. Dastur brauzerda `http://127.0.0.1:8765` manzilida ochiladi.

Talablar birinchi ishga tushirishda avtomatik o‘rnatiladi:

- Python 3.9 yoki yangi versiyasi
- Flask
- XlsxWriter

`NAMUNA.csv` fayli orqali dastur ishini sinab ko‘rish mumkin.

Lokal rejimda CSV fayllar kompyuterdan tashqariga yuborilmaydi. Jami yuklash limiti 350 MB.

## Vercel'da ishga tushirish

GitHub repozitoriysini Vercel loyihasiga ulang. Vercel `app.py` ichidagi `app` Flask obyektini avtomatik topadi; alohida Build Command yoki Output Directory kiritmang. `requirements.txt` dagi kutubxonalar o‘rnatiladi. GitHub'ga yangi commit yuborilganda Vercel qayta deploy qiladi.

Vercel'da CSV fayllar serverga yuborilib, so‘rov davomida qayta ishlanadi. Serverga jami 4 MB gacha CSV yuboring; tayyor Excel ham 4 MB dan kichik bo‘lishi kerak. Bu Vercel funksiyalaridagi [4.5 MB request va response cheklovi](https://vercel.com/docs/functions/limitations) uchun xavfsiz zaxira. Katta yoki maxfiy fayllar uchun lokal rejimdan foydalaning.

Vercel funksiyasi loyiha papkasiga log yozmaydi; xabarlar Vercel Runtime Logs'da ko‘rinadi. Lokal rejimda log `logs/app.log` fayliga yoziladi.
