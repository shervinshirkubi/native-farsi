# اسنادِ فارسی: docx، pptx، PDF، HTML

> منبع: بخشِ فنیِ این فایل (توابعِ پایتون/CSS) از پروژه‌ی متن‌بازِ
> [persian-writing](https://github.com/ali2000hos/persian-writing) گرفته شده،
> تحتِ MIT. جزئیات در `THIRD-PARTY-NOTICES.md`. کدِ فنیه، بدون لحن یا سلیقه؛
> تغییری در منطقش داده نشده.

**کِی این فایل رو بخون:** مقصدِ متن یه فایلِ Word، PowerPoint، PDF یا HTMLِ راست‌چینه.
برای خودِ نثرِ فارسی، همیشه `core.md` + رجیسترِ مناسب رو هم بخون؛ این فایل فقط مکانیکِ فایله.

قبل از هر تولید: `bash scripts/install_fonts.sh`؛ وگرنه فونت خاموش به DejaVu می‌افته
و کسی نمی‌فهمه چرا.

---

## ۱. اصلِ حاکم: RTL یعنی START، نه RIGHT

تو OOXML (فایلِ Word)، وقتی `bidi: true` هست، `AlignmentType.RIGHT` در **چپِ بصری**
رندر می‌شه. همیشه `START` رو تنظیم کن، نه `RIGHT`.

هر پاراگرافِ فارسی: `rtl: true` + فونت با `hint: "cs"` (اسکریپتِ پیچیده).
هر سکشن: `bidi: true`. جدولی که باید راست‌چین جاری بشه: `bidiVisual: true`.

**اعدادِ فارسی در محتوای شماره‌دار اجباریه:** یه `1.` لاتین کلِ پاراگراف رو چپ‌چین می‌کنه.
به همین دلیل: **هیچ‌وقت از استایلِ آماده‌ی `List Number`/`List Bullet` تو Word استفاده نکن.**
نشانه‌گذارشون تو `numbering.xml`ه که bidi نداره و `1.` لاتین تو سمتِ غلط رندر می‌شه، به‌علاوه‌ی
بولتِ OpenSymbol که فونت رو می‌شکنه. نشانه رو به‌صورتِ رانِ معمولی بنویس: «۱.  »، «•  ».

---

## ۲. Word (python-docx)

python-docx واسطِ درجه‌یکی برای RTL نداره؛ باید المان‌هایِ OOXML رو مستقیم تنظیم کنی.

```python
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

def _set(el, parent):
    parent.append(el); return el

def rtl_paragraph(p):
    """پاراگرافِ bidi؛ w:jc رو دستکاری نکن (پیش‌فرضِ bidi = راستِ بصری)،
    یا صریح start بذار. هیچ‌وقت روی پاراگرافِ bidi، جهت رو RIGHT نذار."""
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn('w:bidi')) is None:
        bidi = OxmlElement('w:bidi'); bidi.set(qn('w:val'), '1'); pPr.append(bidi)
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = _set(OxmlElement('w:jc'), pPr)
    jc.set(qn('w:val'), 'start')
    return p  # اصلاحِ ما: مبدأ برگشت نمی‌داد، ولی مثالِ خودشون
              # «p = rtl_paragraph(doc.add_paragraph())» رو فرض می‌کنه — بدونش p می‌شد None

def fa_run(p, text, font='Vazirmatn', size=11, bold=False):
    run = p.add_run(text)
    run.bold = bold
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts')) or _set(OxmlElement('w:rFonts'), rPr)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rFonts.set(qn(attr), font)
    _set(OxmlElement('w:rtl'), rPr)                       # رانِ اسکریپتِ پیچیده
    szCs = _set(OxmlElement('w:szCs'), rPr)               # سایزِ CS (وگرنه ریز/درشتِ غلط)
    szCs.set(qn('w:val'), str(int(size * 2)))
    run.font.size = Pt(size)
    if bold:
        _set(OxmlElement('w:bCs'), rPr)                    # بولدِ CS
    return run

def rtl_table(table):
    tblPr = table._tbl.tblPr
    if tblPr.find(qn('w:bidiVisual')) is None:
        tblPr.append(OxmlElement('w:bidiVisual'))          # ستونِ اول = راستِ بصری

def rtl_section(section):
    """اصلاحِ ما: مبدأ با append اضافه می‌کرد که bidi رو آخرین فرزند می‌ذاره؛
    OOXML لازم داره w:bidi اولین فرزندِ w:sectPr باشه، وگرنه رندرر نادیده‌ش می‌گیره
    (خودِ بخشِ ۷ همین فایل به این نکته اشاره می‌کنه). با insert(0, ...) درست شد."""
    sectPr = section._sectPr
    if sectPr.find(qn('w:bidi')) is None:
        sectPr.insert(0, OxmlElement('w:bidi'))

def keep_with_next(p):
    pPr = p._p.get_or_add_pPr()
    pPr.append(OxmlElement('w:keepNext'))
    pPr.append(OxmlElement('w:keepLines'))

def cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), 'true')
    trPr.append(tblHeader)
```

نکاتی که گاز می‌گیرن:
- `w:szCs` بدونش، متنِ اسکریپتِ پیچیده سایز رو نادیده می‌گیره.
- `w:bCs` بدونش، فقط رانِ لاتین بولد می‌شه.
- کارتِ غیرقابلِ‌شکست: جدولِ ۱×۱ + `cant_split(row)`.

### ۲-۱. سه شکافی که هلپرهای بالا نمی‌بندن

با اینکه هر پاراگراف رو bidi کردی، سه جا هنوز می‌شکنه، چون Word/LibreOffice از
`styles.xml` و `numbering.xml` می‌خونن که python-docx از یه قالبِ لاتین می‌سازتشون:
شماره‌ی لیستِ لاتین، تیترِ آبیِ ناخواسته، و افتادن به DejaVu/OpenSymbol.

**این رو یه‌بار، بلافاصله بعدِ `Document()`، قبل از هر محتوایی اجرا کن:**

```python
def persianize_styles(doc, font='Vazirmatn', size=11):
    """پیش‌فرض‌هایی که python-docx از قالبِ لاتینش به ارث می‌بره رو اصلاح می‌کنه."""
    st = doc.styles['Normal']
    st.font.name = font
    st.font.size = Pt(size)
    rPr = st.element.get_or_add_rPr()
    rF = rPr.find(qn('w:rFonts'))
    if rF is None:
        rF = _set(OxmlElement('w:rFonts'), rPr)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rF.set(qn(a), font)
    _set(OxmlElement('w:rtl'), rPr)
    _set(OxmlElement('w:szCs'), rPr).set(qn('w:val'), str(int(size * 2)))
    _set(OxmlElement('w:bidi'), st.element.get_or_add_pPr())

    # تیترهای آماده: فونتِ فارسی، bidi، و بدونِ رنگِ آبیِ به‌ارث‌رسیده.
    for i in range(1, 5):
        try:
            h = doc.styles[f'Heading {i}']
        except KeyError:
            continue
        h.font.name = font
        h.font.color.rgb = RGBColor(0, 0, 0)      # خنثی؛ رنگ رو کاربر یا برند تعیین می‌کنه
        hr = h.element.get_or_add_rPr()
        hf = hr.find(qn('w:rFonts'))
        if hf is None:
            hf = _set(OxmlElement('w:rFonts'), hr)
        for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
            hf.set(qn(a), font)
        _set(OxmlElement('w:rtl'), hr)
        _set(OxmlElement('w:bidi'), h.element.get_or_add_pPr())
```

**لیستِ شماره‌دار و بولت: استایلِ آماده استفاده نکن.** به‌جاش دستی شماره بزن:

```python
FA_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
to_fa = lambda n: str(n).translate(str.maketrans("0123456789", FA_DIGITS))

for i, item in enumerate(items, 1):
    p = rtl_paragraph(doc.add_paragraph())
    fa_run(p, f"{to_fa(i)}.  ")     # ۱.  ۲.  ۳. (راست‌چین می‌مونه)
    fa_run(p, item)

for item in bullets:
    p = rtl_paragraph(doc.add_paragraph())
    fa_run(p, "•  ")                # • در Vazirmatn تأییدشده؛ ▪ نه
    fa_run(p, item)
```
تورفتگی با `p.paragraph_format.right_indent` (سمتِ RTL)، نه `left_indent`.

**هدر و فوتر بخشِ XMLِ جدان** و چیزی از بدنه به ارث نمی‌برن؛ `rtl_paragraph`
و `fa_run` رو مستقیم روشون اعمال کن.

**تأییدِ نهایی رو از خروجی بگیر، نه از کد:** `pdffonts out.pdf` باید فقط فونتِ
فارسی رو نشون بده. ردیفِ `DejaVu` یا `OpenSymbol` یعنی یه گلیف افتاده به fallback.
`scripts/verify_pdf.py` این رو خودکار تشخیص می‌ده.

---

## ۳. PowerPoint (python-pptx)

PowerPoint سوییچِ RTLِ سندی نداره؛ جهت پاراگراف‌به‌پاراگراف و ران‌به‌ران تنظیم می‌شه.

```python
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

def rtl_paragraph(paragraph, align=PP_ALIGN.RIGHT):
    """جهتِ RTL + راست‌چینیِ بصری. در DrawingML (برخلافِ docx!) algn='r' راستِ فیزیکیه."""
    paragraph.alignment = align
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('rtl', '1')

def fa_run(run, font='Vazirmatn', size=18, bold=False, color=None):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color: run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rPr.set('lang', 'fa-IR')
    cs = rPr.find(qn('a:cs'))
    if cs is None:
        cs = etree.SubElement(rPr, qn('a:cs'))
    cs.set('typeface', font)      # این باعثِ رندرِ واقعیِ گلیفِ فارسیه

def fa_text_frame(tf, font='Vazirmatn', size=18):
    for p in tf.paragraphs:
        rtl_paragraph(p)
        for r in p.runs:
            fa_run(r, font=font, size=size)
```

**نکته‌ی مهم:** تو pptx برخلافِ docx، `algn="r"` راستِ فیزیکیه. تلهٔ START/RIGHTِ
docx اینجا صدق نمی‌کنه. متن/تیترِ فارسی: `rtl='1'` + `PP_ALIGN.RIGHT`.

**چیدمان:** تیتر راست‌چین یا وسط، محتوا از بالا-راست شروع. عکس+متن: عکس چپ، متن راست
(خواننده از راست وارد می‌شه). فلش‌های روند: راست‌به‌چپ، گلیفِ فلش رو برعکس کن (→ بشه ←).

**بولت:**
```python
pPr = paragraph._p.get_or_add_pPr()
buFont = etree.SubElement(pPr, qn('a:buFont')); buFont.set('typeface', 'Vazirmatn')
buChar = etree.SubElement(pPr, qn('a:buChar')); buChar.set('char', '•')
```
`buAutoNum` فقط رقمِ لاتین رندر می‌کنه؛ برای شماره‌ی فارسی `a:buNone` بذار و دستی
پیشوند بزن: «۱. »، «۲. ».

**جدول:** `a:tbl` معادلِ bidiVisual نداره؛ **خودت ستون‌ها رو برعکس بچین**
(اولین ستونِ داده = راست‌ترین سلول) و `rtl_paragraph`/`fa_run` رو روی هر سلول بزن.

**فونت:** جاسازیِ فونتِ pptx بین پلتفرم‌ها ناپایداره. اگه دک از دستگاهِ بدونِ فونت
عبور می‌کنه، همیشه یه PDF هم کنارش بده:
```
bash scripts/install_fonts.sh
soffice --headless --convert-to pdf deck.pptx
python3 scripts/verify_pdf.py deck.pdf --expect-font Vazirmatn
```
بدونِ letter-spacing، بولدِ واقعی نه فیک، فاصله‌ی خط ≥۱.۳.

---

## ۴. HTML و HTML→PDF

```css
body {
  direction: rtl; text-align: right;
  font-family: "Vazirmatn", Tahoma, sans-serif;
  line-height: 1.8;         /* حداقل ۱.۶؛ زیرش فارسی بریده می‌شه */
  letter-spacing: 0;        /* هیچ‌وقت روی فارسی track نذار (اتصالِ حروف می‌شکنه) */
}
h1, h2, h3 { line-height: 1.5; font-weight: 700; }
```

**متنِ ترکیبی (لاتین وسطِ فارسی):**
```html
<p>افزونه‌ی <bdi>WooCommerce 9.5</bdi> رو نصب کن.</p>
<pre dir="ltr">npm install docx</pre>
<span dir="ltr">+98 912 345 6789</span>
```
معادلِ CSS: `unicode-bidi: isolate`. اعداد تو متن: فارسی ۰-۹؛ فیلدِ شماره/URL: `dir="ltr"`.

**چیدمان:** فلکس/گرید زیرِ `dir=rtl` خودکار آینه می‌شن؛ `row-reverse` ننویس، خودش
دوبار flip می‌شه. آیکونِ جهت‌دار (فلش، شورون): `[dir="rtl"] .icon { transform: scaleX(-1); }`.

**فونتِ آنلاین:**
```html
<link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100..900&family=Lalezar&display=swap" rel="stylesheet">
```
آفلاین/ایمیل/HTML→PDF: از `assets/fonts/` با `@font-face` جاسازی کن. هیچ‌وقت رویِ
فونتِ سیستم حساب نکن؛ مک به Geeza Pro (لحنِ غلط) و لینوکس به DejaVu (شکسته) می‌افته.

**ایمیل:** کلاینت‌ها `<style>` رو حذف می‌کنن. `dir="rtl"` رو **صفتِ** هر `<table>`/`<td>`
بذار، نه فقط CSS. استایل فقط inline. فونتِ امنِ ایمیل بدونِ وب‌فونت: Tahoma.

**HTML→PDF:**
```css
@page { size: A4; margin: 2cm; }
h1, h2, h3, h4 { break-after: avoid; }
.card, .price-box, figure { break-inside: avoid; }
p { orphans: 2; widows: 2; }
```
weasyprint یا Chromeِ headless (`--print-to-pdf`)، بعدش حتماً `verify_pdf.py`.

---

## ۵. صفحه‌بندی (هر فرمتی)

- تیترها: `keepNext` + `keepLines` (تیترِ یتیم تهِ صفحه ممنوع)
- کارت/باکس: تو یه جدولِ تک‌سلولی با `cantSplit: true` (هیچ‌وقت بینِ صفحه پاره نشه)
- بعدِ آخرین آیتمِ لیست جداکننده نذار.
- `PageBreak`ِ ولگرد قبلِ سکشن‌بریک ممنوع (صفحه‌ی خالی می‌سازه).

---

## ۶. نمادها

فونت‌های فارسی خیلی گلیف کم دارن. تو Vazirmatn/Lalezار فقط `•` و `·` تأییدشده؛
`▪ ■ ✓ ✕ ● ◆ ⊙` به DejaVu می‌افتن. برای هر نمادِ دیگه، اول cmapِ فونت رو چک کن.

---

## ۷. تأییدِ نهایی قبلِ تحویل

دو دروازه. اول docx رو چک کن، بعد PDF رو:

```bash
# ۱. DOCX: یکپارچگیِ پکیج اول (فایلی که Word بازش نمی‌کنه، RTL‌چکش بی‌معنیه)،
#    بعد bidiِ سکشن، تله‌ی jc=right، فونتِ cs، شماره‌گذاریِ لیست، رنگِ تیتر.
#    --fix با میسینگِ <w:bidi/> رو ترمیم می‌کنه؛ --sanitize آرتیفکتِ Word-for-Mac رو پاک می‌کنه
python3 scripts/verify_docx.py output.docx --expect-font Vazirmatn --fix --sanitize

# ۲. PDF: فونتِ fallback، صفحه‌ی خالی، لیکِ قالب، حروفِ عربی
python3 scripts/verify_pdf.py output.pdf --expect-font Vazirmatn
```

`verify_docx.py` هست چون تنظیم‌کردنِ فلگِ bidi تو کد و **رسیدنش به XML** دو چیزِ
جدان: کتابخونه‌ها گاهی `bidi` رو از پراپرتیِ سکشن می‌ندازن، و OOXML لازم داره
`<w:bidi/>` **اولین فرزند**ِ `<w:sectPr>` باشه؛ هرجایِ دیگه که بیاد، رندرر نادیده‌ش می‌گیره.

چک می‌کنه: صفحه‌ی تقریباً خالی، لیکِ «undefined»/قالب، فونتِ embed‌نشده یا fallback،
حروفِ عربیِ نادرست (`ي`/`ك`) در متنِ استخراج‌شده، و تعدادِ صفحه. هر وارنینگ رو رفع کن، دوباره بساز،
دوباره چک کن.

برای نثر، یه بار `persian_cleanup.py`‌ی اون پروژه یا `farsi_check.py` خودمون +
یه بلندخوانیِ صادقانه بیشترِ فاجعه‌ها رو می‌گیره.
