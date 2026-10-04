# 🧠 آزمون N-Back — نسخه‌ی چهره‌ای

> یک آزمون شناختی **رایگان** برای سنجش **حافظه‌ی کاری** با استفاده از
> تصاویر چهره‌های انسانی. ساخته‌شده با پایتون برای **دانشجویان روان‌شناسی
> و علوم شناختی** فارسی‌زبان.

[![Download](https://img.shields.io/github/v/release/Hossein-aliian/N_Back?label=Download&style=for-the-badge&color=blue)](https://github.com/Hossein-aliian/N_Back/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

---

## ✨ این برنامه چیکار می‌کنه؟

آزمون **N-Back** یکی از معتبرترین تسک‌های روان‌شناسی برای سنجش
**حافظه‌ی کاری** است. در این نسخه:

- تصاویر چهره به‌صورت متوالی نمایش داده می‌شن
- شرکت‌کننده باید تشخیص بده که آیا چهره‌ی فعلی با **N مرحله قبل** یکسانه یا نه
- سه مرحله با ترتیب تصادفی: **جذاب / خنثی / غیرجذاب**
- نتایج به‌صورت خودکار در فایل **Excel فارسی** ذخیره می‌شن
- پشتیبانی کامل از **ورودی متن فارسی** (نام و رشته‌ی تحصیلی)

---

## 🚀 شروع سریع (بدون نیاز به پایتون)

۱. به صفحه‌ی [**Releases**](../../releases/latest) برو
۲. آخرین فایل **`N_Back.rar`** رو دانلود کن
۳. از حالت فشرده خارج کن
۴. روی **`N_Back.exe`** دابل-کلیک کن
۵. تمام! ✅

> ⚠️ بار اول ویندوز ممکنه هشدار بده — روی **More info → Run anyway** بزن.
> این یه هشدار کاذبه، چون فایل با PyInstaller ساخته شده و کد کاملش
> توی همین مخزن موجوده.

---

## 🖼️ عکس‌های خودت رو بذار

توی پوشه‌ی `images/` سه زیرپوشه هست:

```
images/
├── high/       → چهره‌های جذاب
├── neutral/    → چهره‌های خنثی
└── low/        → چهره‌های غیرجذاب
```

عکس‌های نمونه رو با عکس‌های خودت جایگزین کن (JPG یا PNG).

- **حداقل:** ۲ عکس در هر پوشه
- **توصیه‌شده:** ۲۰ تا ۴۰ عکس در هر پوشه

---

## ⚙️ تنظیمات

با کلیک روی **آیکون چرخ‌دنده** (گوشه‌ی بالا-چپ فرم) می‌تونی این‌ها رو تنظیم کنی:

| تنظیم | پیش‌فرض |
|---|---|
| مدت نمایش هر تصویر | ۲ ثانیه |
| استراحت بین تصاویر | ۱ ثانیه |
| تعداد سوال هر مرحله | ۲۲ |
| سطح N | ۲ |
| نرخ تطابق | ۲۰٪ |
| حداقل فاصله‌ی match | ۳ |

---

## 📊 خروجی‌ها

بعد از هر جلسه، این فایل‌ها کنار `N_Back.exe` ساخته می‌شن:

| فایل | توضیح |
|---|---|
| `participants.xlsx` | خلاصه‌ی نتایج هر شرکت‌کننده (سرصفحه‌های فارسی، ۴۷ ستون آماری) |
| `results/nback_*.csv` | لاگ کامل تریال-به-تریال برای تحلیل |
| `settings.json` | تنظیمات فعلی (خودکار) |

---

## 💻 برای توسعه‌دهنده‌ها

```bash
git clone https://github.com/Hossein-aliian/N_Back.git
cd N_Back
pip install -r requirements.txt
python N_Back.py
```

**وابستگی‌ها:** `pygame`, `openpyxl`, `arabic-reshaper`, `python-bidi`

**ساخت فایل EXE:**

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "N_Back" --clean --icon="icon.ico" N_Back.py
```

---

## رایگان برای دانشجویان

این ابزار تحت **مجوز MIT** منتشر شده و استفاده از اون برای **پژوهش‌های
دانشجویی، پایان‌نامه‌ها و پروژه‌های کلاسی کاملاً رایگان و آزاد** است.
اگر ازش توی پژوهش خودت استفاده کردی، خوشحال می‌شم بهم خبر بدی.

---

## 📄 استناد

اگه از این ابزار توی پژوهش خودت استفاده کردی، لطفاً این‌طور ارجاع بده:

```bibtex
@software{nback_face_2026,
  author    = {Hossein Alian},
  title     = {N-Back Test — Face Stimuli Edition},
  year      = {2026},
  publisher = {GitHub},
  url       = {https://github.com/Hossein-aliian/N_Back}
}
```

---

## 👤 نویسنده

**Hossein Alian**
- GitHub: [@Hossein-aliian](https://github.com/Hossein-aliian)
- Email: mr.alian1997@yahoo.com

---

⭐ اگه این پروژه بهت کمک کرد، یه **Star** بده تا بقیه هم پیداش کنن!