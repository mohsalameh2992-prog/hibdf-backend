# 📑 مشروع hibdf.com — الذاكرة الكاملة ودليل التشغيل والإنتاج (Project Memory & Production Guide)

تاريخ التحديث: 2 نوفمبر 2026  
الحالة: **مباشر، متصل 100%، مجاني بالكامل 0.00$، وبدون أي بطاقات ائتمان.**

---

## 1. المعمارية الفنية المعتمدة (Zero-Cost Production Architecture)

```
[ الزائر في المتصفح ] 
         │ (HTTPS)
         ▼
[ Cloudflare Pages / Workers: pdf-pro-web ] ──► النطاق: https://hibdf.com
         │ (API Fetch)
         ▼
[ Cloudflare Tunnel (cloudflared) ] ──────────► النطاق الفرعي: https://backend.hibdf.com/api
         │ (Tunnel Loopback)
         ▼
[ جهازك المحلي (Localhost:8000) ] ────────────► سيرفر Flask (python app.py)
         ├── محرك Gemini Vision AI (مفاتيح مدمجة ومحمية)
         ├── محرك OCR المحلي (Tesseract Engine بديل سريع)
         └── معالجة النصوص العربية وحفظ التنسيق والجداول
         │ (PostgreSQL Connection)
         ▼
[ قاعدة بيانات Supabase المجانية 100% ] ───────► (pooler.supabase.com:5432)
```

---

## 2. تفاصيل الحسابات والروابط الأساسية

| العنصر | القيمة / الرابط | الملاحظات |
|---|---|---|
| **الموقع الرئيسي (Frontend)** | `https://hibdf.com` | مستضاف على Cloudflare (مشروع `pdf-pro-web`) |
| **رابط الـ API والباك إند** | `https://backend.hibdf.com/api` | مشفر ومحمي بشهادة SSL تلقائية من Cloudflare |
| **معرف النفق (Tunnel ID)** | `624b56f4-113c-4728-9e18-8d50e638ee49` | حساب Cloudflare: `5a486020f5e29e1eeb6ac535ace46705` |
| **قاعدة البيانات** | `Supabase PostgreSQL (AWS Sydney Pooler)` | باقة مجانية دائمة، صفر تكلفة |
| **مستودع الكود (GitHub)** | `https://github.com/mohsalameh2992-prog/hibdf-backend` | يحتوي أحدث كود مصحح وشغال 100% |
| **حساب الأدمن المميز** | `moh.salameh2992@gmail.com` | خطة Premium مفعلة، تحويلات غير محدودة، صلاحية Superadmin |

---

## 3. سبب مشكلة سحب الـ 72 دولار السابقة من Google Cloud

- **السبب الفعلي:** كان المشروع في نسخته القديمة معتمداً على 3 خدمات مدفوعة على Google Cloud:
  1. `Cloud SQL for PostgreSQL` (تكلفته الشهرية 45$ - 60$ بدون باقة مجانية).
  2. `Cloud Run` (استضافة الحاويات).
  3. `Google Cloud Storage (GCS)` (تخزين الملفات المؤقتة).
- **الحل الجذري:**
  - تم استبدال `Cloud SQL` بقاعدة **Supabase** المجانية.
  - تم استبدال `Cloud Run` بسيرفر محلي سريع يعمل على معالج ورام جهازك ويوفر أعلى سرعة تحويل.
  - تم استبدال `Cloud Storage` بنظام التخزين المحلي المؤقت مع التنظيف الذاتي التلقائي للملفات.
  - **النتيجة:** التكلفة أصبحت **0.00$ شهرياً دون الحاجة لإدخال أي بطاقة دفع**.

---

## 4. كيفية تشغيل الموقع بعد إعادة تشغيل الكمبيوتر (بنقرة واحدة)

تم إنشاء ملف تشغيل سريع على المسار:
`D:\pdf to word\RUN_HIBDF_SERVER_AND_TUNNEL.bat`

**ما يفعله الملف تلقائياً:**
1. يشغّل سيرفر Flask على المنفذ `8000` في الخلفية.
2. يشغّل نفق Cloudflare Tunnel لربط جهازك بالدومين الفرعي `backend.hibdf.com`.
3. يضمن بقاء الموقع متصلاً ومتاحاً لجميع الزوار حول العالم.

---

## 5. الإصلاحات الفنية التي تم تطبيقها على الكود:

1. **الفرونت إند (`js/app.js` و `js/app_dev.js`):**
   - تم تعديل `const API_BASE = 'https://backend.hibdf.com/api';` بدلاً من رابط جوجل كلاود المحذوف.
   - تم تحديث ملف الحزمة ورفعها على Cloudflare Worker بنجاح.
2. **محرك الذكاء الاصطناعي (`converter.py`):**
   - حل خطأ `_api_keys_list is not defined` مما أعاد تفعيل التعرف البصري عبر Gemini Vision بنجاح تام.
3. **التخزين والتحميل (`app.py`):**
   - إزالة الاعتماد الإجباري على Google Cloud Storage (GCS).
   - تفعيل التنزيل المباشر للملفات الناتجة بصيغة `.docx` محلياً عبر المسار `/api/download/<id>`.
4. **تسجيل الدخول والأدمن:**
   - تفعيل مسار التحقق من التوكن لحساب `moh.salameh2992@gmail.com` ليعمل بخطة Unlimited فوراً.
