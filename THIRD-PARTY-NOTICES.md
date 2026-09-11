# یادداشتِ حق‌نشرِ شخصِ ثالث

این اسکیل بخش‌هایی از کد/فونتِ پروژه‌ی متن‌بازِ زیر رو استفاده می‌کنه:

## persian-writing

منبع: https://github.com/ali2000hos/persian-writing
مجوز: MIT
Copyright (c) 2026 persian-writing contributors

فایل‌های زیر عیناً کپی یا با کمترین تغییرِ فنی (نه محتوایی/سلیقه‌ای) از این پروژه اومدن:

| فایلِ ما | منبع |
|---|---|
| `scripts/verify_docx.py` | `scripts/verify_docx.py` |
| `scripts/verify_pdf.py` | `scripts/verify_pdf.py` |
| `scripts/install_fonts.sh` | `scripts/install_fonts.sh` |
| `assets/fonts/Vazirmatn-*.ttf` | `assets/fonts/Vazirmatn-*.ttf` |
| `assets/fonts/Lalezar-Regular.ttf` | `assets/fonts/Lalezar-Regular.ttf` |
| `references/documents.md` (توابعِ پایتون/CSS) | `references/docx-pdf.md`, `references/pptx.md`, `references/html-css.md` |

**دو اصلاح نسبت به مبدأ، هر دو با تستِ سرتاسریِ خودمون پیدا شدن:**
۱. تابعِ `rtl_paragraph` (نسخه‌ی docx) مقدار برنمی‌گردوند، درحالی‌که خودِ مثالِ مبدأ
   الگوی `p = rtl_paragraph(doc.add_paragraph())` رو نشون می‌ده که بدونِ `return p`
   همیشه `p` رو `None` می‌کنه. `return p` اضافه شد.
۲. تابعِ `rtl_section` با `sectPr.append(...)` کار می‌کرد که `w:bidi` رو **آخرین**
   فرزند می‌ذاره؛ درحالی‌که OOXML لازم داره اولین فرزند باشه (نکته‌ای که خودِ همون
   فایلِ مبدأ، بخشِ تأییدِ نهایی، بهش اشاره می‌کنه — تناقضِ داخلی). با `insert(0, …)`
   درست شد.

متنِ کاملِ مجوزِ MIT:

```
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## فونت‌های باندل‌شده

فونت‌های `assets/fonts/` تحتِ MIT نیستن؛ تحتِ **SIL Open Font License 1.1** هستن:

- **Vazirmatn** — The Vazirmatn Project Authors، github.com/rastikerdar/vazirmatn
- **Lalezar** — Borna Izadpanah، github.com/BornaIz/Lalezar

هر دو برایِ استفاده‌ی تجاری آزادن.
