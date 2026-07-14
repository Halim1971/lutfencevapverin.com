# lutfencevapverin.com MVP

Mobil odaklı LCV / RSVP takip sistemi.

## Kurulum

```bash
cd /home/halim/Documents/CALISMA/LCV
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
flask seed
python app.py
```

Uygulama varsayılan olarak:

```text
http://localhost:5062
```

## Demo girişleri

Süper admin:

```text
admin@lutfencevapverin.com
123456
```

Düğün sahibi:

```text
demo@lutfencevapverin.com
123456
```

## JSON aktarım

Başka lutfencevapverin.com kullanıcısına davetli listesi aktarmak için JSON dışarı aktar / JSON içeri aktar kullanılır.

```json
{
  "guests": [
    {"full_name": "Ayşe Yılmaz", "phone": "90532..."}
  ]
}
```

## WhatsApp mantığı

WhatsApp Business API kullanılmaz. Sistem `wa.me` linki üretir. Mesaj içinde kişiye özel LCV linki vardır.

Toplu gönderim mobil tarayıcı ve WhatsApp izinlerine bağlıdır; ilk sürümde yarı otomatik test ekranı olarak çalışır.

## Ana akış

1. Düğün sahibi giriş yapar.
2. Davetiyesini yükler.
3. Davetlileri manuel veya Excel ile ekler.
4. Başka kullanıcıya aktarım gerekiyorsa JSON dışarı aktarır; gelen JSON dosyasını JSON içeri aktar ekranından alır.
5. WhatsApp ile kişisel LCV linki gönderir.
6. Davetli linkte Katılıyorum / Katılamıyorum cevabı verir.
7. Katılımcılar, Katılamayanlar ve Cevap Vermeyenler listeleri güncellenir.
