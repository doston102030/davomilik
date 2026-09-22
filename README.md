# Davomiylik — SVOD TIZIMI

CSV fayllarni tekshiradigan va Excel svod hisobotini yaratadigan dastur. Windows'da lokal, Vercel'da veb rejimda ishlaydi.

## Imkoniyatlar

- Bir yoki bir nechta CSV faylni, ichida CSV fayllari bo‘lgan ZIP’ni yoki butun papkani qabul qiladi
- Majburiy ustunlar va sana formatlarini tekshiradi
- Realizatsiyalar orasidagi kun farqini hisoblaydi
- Dublikat abonent kodlarini alohida varaqda ko‘rsatadi
- Raygaz va mahalla kesimida Excel hisoboti yaratadi
- “Давомилик” varag‘ida rasmdagi kabi `Райгаз | 31-35 ... 71+ | Жами` jadvalini beradi
- Ixtiyoriy .xlsx etalonning yangi (`Райгаз` sarlavhali) yoki eski (`ХГТ булими` sarlavhali) shaklini qabul qiladi; “Солиштириш” varag‘ida etalon sonining yoniga yangi jadvaldagi ayiriladigan sonni qizil yozadi (masalan `5829 -1383`), “Фарқлар” varag‘ida esa etalon, dastur va qoldiqning sonli jadvallarini beradi
- Lokal rejimda ma’lumotlarni faqat kompyuterning o‘zida qayta ishlaydi

## Windows'da lokal ishga tushirish

Windows’da `ISHGA_TUSHIRISH.bat` faylini ikki marta bosing. Dastur brauzerda `http://127.0.0.1:8765` manzilida ochiladi.

Talablar birinchi ishga tushirishda avtomatik o‘rnatiladi:

- Python 3.9 yoki yangi versiyasi
- Flask
- XlsxWriter

`NAMUNA.csv` fayli orqali dastur ishini sinab ko‘rish mumkin. Katta maydon alohida CSV, CSVlar joylangan `.zip` yoki butun papkani qabul qiladi; papkaning ichki papkalari ham ko‘riladi. ZIP diskka ochilmaydi, ichidagi CSVlar bevosita o‘qiladi. Etalon `.xlsx` faylni ham solishtirish blokining ustiga sudrab tashlash mumkin. Tanlangan CSV/ZIP fayllarini va etalonni “Olib tashlash” tugmasi bilan almashtirish mumkin.

Solishtirish uchun CSV fayllar bilan birga tayyor davomat `.xlsx` faylini ham tanlang. CSV va etalonning sanalarini tekshiring: turli kunlar yuklansa, farq tabiiy ravishda paydo bo‘ladi. “Давомилик” jadvaliga 0–30 kunlik yozuvlar kiritilmaydi; ular “Свод” va “База”da qoladi. Eski etalonda `1-30` ustuni bo‘lsa, “Солиштириш”da u ham tekshiriladi (0 kunlik yozuvlarsiz). Etalondagi “umuman gaz olmaganlar” va “muddatida almashtirishga ehtiyoj yo‘q” kabi maxsus toifalar joriy CSV formatida alohida belgilanmagan, shuning uchun dastur ularni taxmin qilmaydi.

Lokal rejimda yuklangan fayllar kompyuterdan tashqariga yuborilmaydi. Jami yuklash limiti 350 MB.

## Vercel'da ishga tushirish

GitHub repozitoriysini Vercel loyihasiga ulang. Vercel `app.py` ichidagi `app` Flask obyektini avtomatik topadi; alohida Build Command yoki Output Directory kiritmang. `requirements.txt` dagi kutubxonalar o‘rnatiladi. GitHub'ga yangi commit yuborilganda Vercel qayta deploy qiladi.

Vercel'da yuklangan fayllar serverga yuborilib, so‘rov davomida qayta ishlanadi. CSV va ixtiyoriy etalon Excel birga jami 4 MB dan oshmasin; tayyor Excel ham 4 MB dan kichik bo‘lishi kerak. Bu Vercel funksiyalaridagi [4.5 MB request va response cheklovi](https://vercel.com/docs/functions/limitations) uchun xavfsiz zaxira. Katta yoki maxfiy fayllar uchun lokal rejimdan foydalaning.

Vercel funksiyasi loyiha papkasiga log yozmaydi; xabarlar Vercel Runtime Logs'da ko‘rinadi. Lokal rejimda log `logs/app.log` fayliga yoziladi.
