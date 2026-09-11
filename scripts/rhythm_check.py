#!/usr/bin/env python3
"""سنجه‌ی ریتم: واریانسِ طولِ جمله رو عدد می‌ده. ابزارِ تشخیصیه، نه گیت.

فقط برای متنِ بلند (>۳۰۰ کلمه یا روایت) در گامِ چهارمِ ورک‌فلو صدا زده می‌شه.
craft.md بخشِ ۳-۱ («طول باید موج داشته باشه») رو اندازه می‌گیره، جایگزینش نمی‌کنه.

اجرا:
    python3 rhythm_check.py متن.md
"""
import re
import sys

SENT_SPLIT = re.compile(r"[.؟!]")
CODE_FENCE = re.compile(r"^```")
QUOTE_LINE = re.compile(r"^\s*>")


def strip_noise(text: str) -> str:
    out, in_code = [], False
    for line in text.split("\n"):
        if CODE_FENCE.match(line.lstrip()):
            in_code = not in_code
            continue
        if in_code or QUOTE_LINE.match(line):
            continue
        out.append(line)
    return "\n".join(out)


def sentence_lengths(text: str):
    lens = []
    for s in SENT_SPLIT.split(text):
        n = len(s.split())
        if n >= 2:  # جمله‌ی واقعی، نه باقیمانده‌ی خالی
            lens.append(n)
    return lens


def variance_coefficient(lens):
    """ضریبِ تغییرات (CV) = انحرافِ معیار / میانگین. هرچی بالاتر، ریتم متنوع‌تر."""
    if len(lens) < 3:
        return None
    mean = sum(lens) / len(lens)
    if mean == 0:
        return None
    var = sum((x - mean) ** 2 for x in lens) / len(lens)
    std = var ** 0.5
    return std / mean


def main():
    if len(sys.argv) != 2:
        print("اجرا: python3 rhythm_check.py متن.md")
        return 1
    path = sys.argv[1]
    text = strip_noise(open(path, encoding="utf-8").read())
    lens = sentence_lengths(text)

    if len(lens) < 5:
        print("متن خیلی کوتاهه؛ سنجه‌ی ریتم برای متنِ بلند معناداره (>۳۰۰ کلمه).")
        return 0

    cv = variance_coefficient(lens)
    print(f"تعدادِ جمله: {len(lens)}")
    print(f"کوتاه‌ترین: {min(lens)} کلمه  |  بلندترین: {max(lens)} کلمه  |  میانگین: {sum(lens)//len(lens)} کلمه")
    print(f"ضریبِ تنوع (CV): {cv:.2f}")

    if cv < 0.35:
        print("⚠️  ریتم صافه. جمله‌ها هم‌قدن — نشانه‌ی ماشینی‌بودن (craft.md ۳-۱).")
        print("   این حکم نیست، یادآوره: چند جمله رو کوتاه‌تر یا بلندتر کن، بر اساسِ محتوا نه فرمول.")
    else:
        print("✅ ریتم موج داره.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
