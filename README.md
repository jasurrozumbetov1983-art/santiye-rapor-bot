# Şantiye Rapor Kontrol Botu — V2

Bu sürüm, gönderdiğiniz örneklerdeki gibi serbest metin raporlarını okuyacak şekilde tasarlanmıştır.

Örnek:
"04.09.2026 ... Stadyum ... Toplam 39 kişi"
→ Şantiye: STADYUM
→ Toplam: 39 kişi
→ Excel'e kaydedilir.

"ŞANTİYE: Piramit Tower ... Genel Toplam: 69 Kişi"
→ Şantiye: PİRAMİT
→ Toplam: 69 kişi

Bot ayrıca saat 15:00'te hangi şantiyelerin rapor verdiğini/vermeyeni Türkçe olarak gönderir.

ÖNEMLİ:
Botun Telegram grubundaki tüm rapor mesajlarını okuyabilmesi için BotFather'da Privacy Mode kapatılmalıdır.
Bot gruba eklenmelidir.

Kurulum:
1. @BotFather → /newbot → token alın.
2. .env.example dosyasını .env yapın.
3. BOT_TOKEN ve ADMIN_CHAT_ID girin.
4. pip install -r requirements.txt
5. python bot.py

Bu MVP'de şantiye ve toplam kişi otomatik çıkarılır. Sonraki aşamada görev bazında kişi sayıları, izinli/hasta/gece/depo/mobilizasyon ayrımı ve günlük/haftalık Excel özetleri eklenebilir.
