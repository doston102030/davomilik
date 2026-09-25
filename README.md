# Davomiylik — SVOD TIZIMI

CSV/XLSX fayllarni tekshiradigan va Excel svod hisobotini yaratadigan dastur. Windows'da lokal, Vercel'da veb rejimda ishlaydi.

## Imkoniyatlar

- Bir yoki bir nechta CSV faylni, ichida CSV fayllari bo‘lgan ZIP’ni yoki butun papkani qabul qiladi
- `Учреждение / Инспектор / Принял / Реализовал / Вернул` formatidagi XLSX hisobotini avtomatik tanib, qabul-sotuv svodini yaratadi
- Majburiy ustunlar va sana formatlarini tekshiradi
- Ustunlar soni buzilgan qatorlarni hamda aynan bir xil faylning takror yuklanishini bloklaydi
- Realizatsiyalar orasidagi kun farqini hisoblaydi
- Dublikat abonent kodlarini alohida varaqda ko‘rsatadi
- Raygaz va mahalla kesimida Excel hisoboti yaratadi
- “Давомилик” varag‘ida rasmdagi kabi `Райгаз | 31-35 ... 71+ | Жами` jadvalini beradi
- Ixtiyoriy .xlsx etalonning yangi (`Райгаз` sarlavhali) yoki eski (`ХГТ булими` sarlavhali) shaklini qabul qiladi; “Солиштириш” varag‘ida etalon sonining yoniga yangi jadvaldagi ayiriladigan sonni qizil yozadi (masalan `5829 -1383`), “Фарқлар” varag‘ida esa etalon, dastur va qoldiqning sonli jadvallarini beradi
- Lokal rejimda ma’lumotlarni faqat kompyuterning o‘zida qayta ishlaydi
- Alohida “SOTILGAN GAZLAR” sahifasida kechagi va yangi holatni abonent kodi bo‘yicha solishtiradi
- Shu sahifada ikki `Принял / Реализовал / Вернул` XLSX hisobotini raygaz va chilangar kesimida solishtirib, qabul, sotuv va qaytarilgan miqdor farqlarini ko‘rsatadi
- “So‘nggi realizatsiya” sanasi yangilangan abonentlarni sotuv sifatida Raygaz kesimida va kodlari bilan ko‘rsatadi
- Yangi holatda yo‘qolgan, yangi qo‘shilgan yoki noaniq o‘zgargan kodlarni sotuvga qo‘shmasdan alohida nazorat qiladi
- Bir xil kodning Raygaz, mahalla, abonent yoki ehtiyoj ma’lumoti o‘zgarsa, uni ham nazoratga chiqaradi
- Sahifalar maxfiy operatsion ma’lumotlarni brauzer keshida saqlamaslik uchun `no-store` va qo‘shimcha xavfsizlik sarlavhalarini yuboradi
- Fayl tanlanganda yuqoridan qisqa xabar va ovoz chiqadi; hisob tugaganda alohida yakuniy xabar hamda ovoz beriladi

## Windows'da lokal ishga tushirish

Windows’da `ISHGA_TUSHIRISH.bat` faylini ikki marta bosing. Dastur brauzerda `http://127.0.0.1:8765` manzilida ochiladi.

`E-GAZ HISOBOTI` sahifasida davr fayl nomidan olinadi. E-GAZ’ning asl `YYYY-MM-DD` sanali nomi ham, `1-31 Avgust.xlsx` kabi sodda nom ham qabul qilinadi; ikkinchi holatda hisobot sarlavhasi `1-31 Avgust holatiga` bo‘ladi.

Sotuvlarni aniqlash uchun yuqoridagi `SOTILGAN GAZLAR` tugmasini bosing. Chap tomonga kechagi holat CSV/XLSX/ZIP fayllarini, o‘ng tomonga yangi holat fayllarini joylang va `Sotuvlarni hisoblash` tugmasini bosing. Abonent fayllari kod va realizatsiya sanasi bo‘yicha, `Принял / Реализовал / Вернул` XLSX hisobotlari esa raygaz va chilangar miqdorlari bo‘yicha avtomatik solishtiriladi. Ikkinchi formatda sotilmagan gaz har bir qatorda `Принял − Реализовал` orqali hisoblanadi; manbadagi `Вернул` 0 yoki noto‘liq bo‘lsa ham qoldiq yo‘qolmaydi. Gaz natijasida kechagi chilangar satrlari va har bir raygazning `ЖАМИ` qatori bitta jadvalda ko‘rsatiladi; alohida takroriy Raygaz yoki `Nazorat` jadvali chiqarilmaydi. `Rasmdagidek Excel yuklash` tugmasi kechagi jadval ko‘rinishini saqlaydi, ko‘k sarlavhada kechagi hisobot davrini ko‘rsatadi va `G` ustunidagi har bir `ЖАМИ` hamda umumiy `Жами` qatoriga `yangi sotilmagan gaz − kechagi sotilmagan gaz` farqini yozadi (`1400 → 1000 = -400`). Natija tafsilotini CSV ko‘rinishida ham yuklab olish mumkin. Eski `.xls` faylni avval Excel orqali `.xlsx` formatida saqlash kerak.

Talablar birinchi ishga tushirishda avtomatik o‘rnatiladi:

- Python 3.9 yoki yangi versiyasi
- Flask
- XlsxWriter

`NAMUNA.csv` fayli orqali dastur ishini sinab ko‘rish mumkin. Katta maydon alohida CSV/XLSX, ular joylangan `.zip` yoki butun papkani qabul qiladi; papkaning ichki papkalari ham ko‘riladi. ZIP diskka ochilmaydi, ichidagi fayllar bevosita o‘qiladi. `Принял / Реализовал / Вернул` XLSX yuklansa, natijada `Свод`, `Райгаз`, `Чилангарлар` va `Назорат` varaqlari yaratiladi. Etalon `.xlsx` faylni solishtirish blokining ustiga sudrab tashlash mumkin. Tanlangan fayllarni va etalonni “Olib tashlash” tugmasi bilan almashtirish mumkin.

Solishtirish uchun CSV fayllar bilan birga tayyor davomat `.xlsx` faylini ham tanlang. CSV va etalonning sanalarini tekshiring: turli kunlar yuklansa, farq tabiiy ravishda paydo bo‘ladi. “Давомилик” jadvaliga 0–30 kunlik yozuvlar kiritilmaydi; ular “Свод” va “База”da qoladi. Eski etalonda `1-30` ustuni bo‘lsa, “Солиштириш”da u ham tekshiriladi (0 kunlik yozuvlarsiz). Etalondagi “umuman gaz olmaganlar” va “muddatida almashtirishga ehtiyoj yo‘q” kabi maxsus toifalar joriy CSV formatida alohida belgilanmagan, shuning uchun dastur ularni taxmin qilmaydi.

Lokal rejimda yuklangan fayllar kompyuterdan tashqariga yuborilmaydi. Jami yuklash limiti 350 MB.

## Vercel'da ishga tushirish

GitHub repozitoriysini Vercel loyihasiga ulang. Vercel `app.py` ichidagi `app` Flask obyektini avtomatik topadi; alohida Build Command yoki Output Directory kiritmang. `requirements.txt` dagi kutubxonalar o‘rnatiladi. GitHub'ga yangi commit yuborilganda Vercel qayta deploy qiladi.

Vercel'da yuklangan fayllar serverga yuborilib, so‘rov davomida qayta ishlanadi. CSV va ixtiyoriy etalon Excel birga jami 4 MB dan oshmasin; tayyor Excel ham 4 MB dan kichik bo‘lishi kerak. Bu Vercel funksiyalaridagi [4.5 MB request va response cheklovi](https://vercel.com/docs/functions/limitations) uchun xavfsiz zaxira. Katta yoki maxfiy fayllar uchun lokal rejimdan foydalaning.

Vercel funksiyasi loyiha papkasiga log yozmaydi; xabarlar Vercel Runtime Logs'da ko‘rinadi. Lokal rejimda log `logs/app.log` fayliga yoziladi.
