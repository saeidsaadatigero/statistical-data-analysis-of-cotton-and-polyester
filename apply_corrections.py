# -*- coding: utf-8 -*-
"""اعمال اصلاحات استاد روی دو سند فارسی.

فایل‌های اصلی («گزارش_تحلیل_آماری.docx» و «راهنمای_کامل_پروژه.docx») فقط خوانده
می‌شوند؛ خروجی در دو فایل تازه ذخیره می‌گردد:
    Persian/گزارش_تحلیل_آماری_اصلاح‌شده.docx
    Persian/راهنمای_کامل_پروژه_اصلاح‌شده.docx

هیچ جدول یا نموداری حذف نمی‌شود. ساختار راست‌به‌چپ (w:bidi / w:rtl / w:bidiVisual)
در تمام سند حفظ می‌گردد، زیرا متن‌های تازه با همان توابع کمکی فارسی ساخته می‌شوند.

اجرا:
    python apply_corrections.py            # اعمال اصلاحات (متن‌های تازه هایلایت زرد می‌شوند)
    python apply_corrections.py --verify   # بازرسی سندهای ساخته‌شده و فهرست اصلاح‌های اعمال‌شده
    python apply_corrections.py --dump     # نمایش کاربرگ‌های پارچه برای بازبینی
"""
import os, sys, copy, re
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

import corrections_text as T

OUT = 'Persian'
SRC_REPORT = os.path.join(OUT, 'گزارش_تحلیل_آماری.docx')
SRC_GUIDE = os.path.join(OUT, 'راهنمای_کامل_پروژه.docx')
DST_REPORT = os.path.join(OUT, 'گزارش_تحلیل_آماری_اصلاح‌شده.docx')
DST_GUIDE = os.path.join(OUT, 'راهنمای_کامل_پروژه_اصلاح‌شده.docx')

NAVY = RGBColor(0x0D, 0x36, 0x6B)
BLUEC = RGBColor(0x1C, 0x5C, 0xAB)
GREY = RGBColor(0x5C, 0x5C, 0x58)
DARKRED = RGBColor(0x9B, 0x1C, 0x1C)
FA_FONT, FA_TITLE_FONT = 'B Nazanin', 'B Titr'
CHANGE_LOG = []

# وقتی روشن باشد، هر متن تازه‌ساخته یا ویرایش‌شده با هایلایت زرد مشخص می‌شود
# تا خواننده بداند کجا باید تصاویر واقعی و مراجع را تکمیل کند.
HL = {'on': True, 'suppress': 0, 'runs': 0}


def hl_off():
    HL['on'] = False


def hl_on():
    HL['on'] = True


def hl_count():
    return HL['runs']


def log(kind, text):
    CHANGE_LOG.append((kind, text))


# ===========================================================================
# ابزارهای پایه‌ی سند فارسی (راست‌به‌چپ)
# ===========================================================================
def _el(tag, val='1'):
    e = OxmlElement(tag)
    e.set(qn('w:val'), val)
    return e


def _strip(rPr, tag):
    """حذف عنصر ویژگی تکراری تا در بازنویسی متن، w:rtl و w:szCs دوباره اضافه نشوند."""
    for e in rPr.findall(qn(tag)):
        rPr.remove(e)


def highlight(run):
    """هایلایت زرد روی متن‌های تازه یا ویرایش‌شده.

    جدول‌های تازه فقط در ردیف سرستون و در عنوان (caption) هایلایت می‌شوند تا
    خواندن عددهای داخل جدول دشوار نشود.
    """
    if HL['on'] and not HL['suppress'] and run.text.strip():
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        HL['runs'] += 1
    return run


def rtl_run(run, font=FA_FONT, size=13):
    run.font.name = font
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rPr.rFonts.set(qn('w:cs'), font)
    rPr.rFonts.set(qn('w:ascii'), font)
    rPr.rFonts.set(qn('w:hAnsi'), font)
    _strip(rPr, 'w:rtl')
    rPr.append(_el('w:rtl'))
    _strip(rPr, 'w:szCs')
    szcs = OxmlElement('w:szCs')
    szcs.set(qn('w:val'), str(int(size * 2)))
    rPr.append(szcs)
    highlight(run)
    return run


def rtl_par(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    pPr = p._p.get_or_add_pPr()
    for e in pPr.findall(qn('w:bidi')):
        pPr.remove(e)
    pPr.append(_el('w:bidi'))
    p.alignment = align
    return p


def rtl_table(t):
    tblPr = t._tbl.tblPr
    for e in tblPr.findall(qn('w:bidiVisual')):
        tblPr.remove(e)
    tblPr.append(_el('w:bidiVisual'))
    return t


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def P(doc, text, size=13, bold=False, italic=False, color=None,
      align=WD_ALIGN_PARAGRAPH.JUSTIFY, font=FA_FONT, before=0, after=6):
    p = doc.add_paragraph()
    r = p.add_run(text)
    rtl_run(r, font, size)
    r.bold, r.italic = bold, italic
    if color is not None:
        r.font.color.rgb = color
    rtl_par(p, align)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    return p


def H(doc, text, level=1):
    size = {1: 16, 2: 14, 3: 13.5}[level]
    return P(doc, text, size=size, bold=True,
             color=NAVY if level == 1 else BLUEC,
             align=WD_ALIGN_PARAGRAPH.RIGHT,
             font=FA_TITLE_FONT if level == 1 else FA_FONT,
             before=16 if level == 1 else 12, after=6)


def BUL(doc, text, size=13):
    p = doc.add_paragraph(style='List Bullet')
    r = p.add_run(text)
    rtl_run(r, FA_FONT, size)
    rtl_par(p, WD_ALIGN_PARAGRAPH.JUSTIFY)
    p.paragraph_format.space_after = Pt(3)
    return p


def REF(doc, text):
    """یک مدخل استناد، وسط‌چین با تورفتگی."""
    p = doc.add_paragraph()
    r = p.add_run(text)
    rtl_run(r, FA_FONT, 11.5)
    r.font.color.rgb = GREY
    rtl_par(p, WD_ALIGN_PARAGRAPH.JUSTIFY)
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(2)
    return p


def FORMULA(doc, text, note=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = 'Cambria Math'
    r.font.size = Pt(13)
    rPr = r._element.get_or_add_rPr()
    for k in ('w:cs', 'w:ascii', 'w:hAnsi'):
        rPr.rFonts.set(qn(k), 'Cambria Math')
    r.bold = True
    r.font.color.rgb = NAVY
    plain = OxmlElement('w:rtl')
    plain.set(qn('w:val'), '0')
    rPr.append(plain)
    highlight(r)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    if note:
        P(doc, note, size=11, italic=True, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
    return p


def PLACEHOLDER(doc, text, note=None):
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    rtl_table(t)
    c = t.rows[0].cells[0]
    c.text = ''
    pad = OxmlElement('w:tcMar')
    for side in ('top', 'bottom', 'left', 'right'):
        e = OxmlElement(f'w:{side}')
        e.set(qn('w:w'), '200')
        e.set(qn('w:type'), 'dxa')
        pad.append(e)
    c._tc.get_or_add_tcPr().append(pad)
    r = c.paragraphs[0].add_run(text)
    rtl_run(r, FA_FONT, 12)
    r.bold = True
    r.font.color.rgb = DARKRED
    rtl_par(c.paragraphs[0], WD_ALIGN_PARAGRAPH.CENTER)
    shade(c, 'FFF4F4')
    if note:
        P(doc, note, size=11, italic=True, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=14)
    return t


def CAPTION(doc, text, size=11.5):
    return P(doc, text, size=size, bold=True, color=GREY,
             align=WD_ALIGN_PARAGRAPH.RIGHT, before=10, after=4)


def NOTE(doc, text, size=10.5):
    return P(doc, text, size=size, italic=True, color=GREY, before=3, after=10)


def TABLE(doc, rows, size=11, widths=None, highlight_data=False):
    t = doc.add_table(rows=0, cols=len(rows[0]))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    rtl_table(t)
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        # ردیف سرستون و (در صورت درخواست) همه‌ی سلول‌ها هایلایت می‌شوند؛
        # سلول‌های داده به‌طور پیش‌فرض بدون هایلایت می‌مانند تا خوانا بمانند.
        mute = (i > 0) and not highlight_data
        if mute:
            HL['suppress'] += 1
        for j, v in enumerate(row):
            cells[j].text = ''
            r = cells[j].paragraphs[0].add_run(str(v))
            rtl_run(r, FA_FONT, size)
            rtl_par(cells[j].paragraphs[0], WD_ALIGN_PARAGRAPH.CENTER)
            if i == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                shade(cells[j], '1C5CAB')
            elif i % 2 == 0:
                shade(cells[j], 'EEF4FC')
        if mute:
            HL['suppress'] -= 1
    return t


def PAGEBREAK(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    return p


# ===========================================================================
# ابزارهای کار روی سند موجود
# ===========================================================================
def capture(doc, fn):
    """عناصر تازه‌ساخته‌شده در انتهای body را برمی‌گرداند (بدون sectPr)."""
    body = doc.element.body
    n = len(body)
    fn(doc)
    return [el for el in list(body)[n:] if el.tag != qn('w:sectPr')]


def place_after(anchor_par, els):
    cur = anchor_par._p
    for el in els:
        cur.addnext(el)
        cur = el


def place_before(anchor_par, els):
    cur = anchor_par._p
    for el in els:
        cur.addprevious(el)


def find_par(doc, text, exact=True, style=None, start=0):
    for i in range(start, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        if style is not None and p.style.name != style:
            continue
        if (p.text == text) if exact else (text in p.text):
            return p
    return None


def delete_par(p):
    p._p.getparent().remove(p._p)


def set_par_text(p, text):
    """متن پاراگراف موجود را جایگزین می‌کند و راست‌به‌چپ بودن را حفظ می‌نماید."""
    runs = p.runs
    if not runs:
        r = p.add_run(text)
        rtl_run(r, FA_FONT, 13)
        rtl_par(p, WD_ALIGN_PARAGRAPH.JUSTIFY)
        return p
    keep = runs[0]
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    size = keep.font.size.pt if keep.font.size else 13
    keep.text = text
    rtl_run(keep, keep.font.name or FA_FONT, size)
    return p


def replace_sub(doc, old, new, count=1):
    """جایگزینی یک زیررشته در متن پاراگراف‌های سطح سند."""
    done = 0
    for p in doc.paragraphs:
        if old in p.text:
            set_par_text(p, p.text.replace(old, new))
            done += 1
            if done >= count:
                return done
    return done


def apply_rename_map(doc, mapping):
    """تغییر متن بر پایه‌ی نقشه‌ی «متن اصلی ← متن تازه».

    متن اصلی یک‌بار خوانده می‌شود، بنابراین تغییرها روی هم اثر نمی‌گذارند.
    """
    for p in doc.paragraphs:
        t = p.text
        if t in mapping:
            set_par_text(p, mapping[t])
            log('تغییر متن', f'{t[:48]} ← {mapping[t][:48]}')


def par_text_of(el):
    return Paragraph(el, None).text


def clear_section(doc, start_text, stop_text):
    """همه‌ی پاراگراف‌های بین دو سرتیتر را پاک می‌کند (جدول‌ها و تصاویر می‌مانند)."""
    body = doc.element.body
    kids = list(body)

    def idx(txt):
        for k, el in enumerate(kids):
            if el.tag == qn('w:p') and Paragraph(el, None).text == txt:
                return k
        raise KeyError(txt)

    i, j = idx(start_text), idx(stop_text)
    n = 0
    for el in kids[i + 1:j]:
        if el.tag == qn('w:p'):
            body.remove(el)
            n += 1
    return n


def last_body_par(doc):
    for el in reversed(list(doc.element.body)):
        if el.tag == qn('w:p'):
            return Paragraph(el, None)
    return doc.paragraphs[-1]


def save(doc, path):
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
    try:
        doc.save(path)
    except OSError as e:
        raise SystemExit(f'[خطا] ذخیره‌ی «{path}» ممکن نشد ({e.strerror}). '
                         'اگر فایل در Word باز است آن را ببندید و دوباره اجرا کنید.')
    print('[سند اصلاح‌شده]', path)


# ===========================================================================
# داده‌ها (برای جدول‌های تفکیکی، کدهای DOE و مشخصات پارچه)
# ===========================================================================
SRC = [f for f in os.listdir('.')
       if f.lower().endswith('.xlsx') and not f.startswith(('Descriptive', '~$'))
       and not f.startswith('آمار_')][0]

C_TENS = 'کشش'
C_COUNT = 'نمره نخ قرره\nپنبه (ne)'
C_W1 = 'درصد آبرفت کل در شستشو اول'
C_W2 = 'درصد آبرفت کل در شستشو دوم'
C_REP = 'تکرار'
C_CODE = 'کد نمونه'
TENS_MAP = {'کم': 'کم', 'متوسط': 'متوسط', 'زیاد': 'زیاد'}
TENS_ORDER_FA = ['کم', 'متوسط', 'زیاد']


def clean_code(s):
    """کد نمونه را یکدست می‌کند: «C20L - 1» ← «C20L-1»."""
    return re.sub(r'\s*-\s*', '-', str(s).strip())


def load_samples(sheet, fabric_fa):
    df = pd.read_excel(SRC, sheet_name=sheet)
    df = df[df[C_TENS].isin(TENS_MAP)].copy()
    out = pd.DataFrame({
        'code': df[C_CODE].map(clean_code),
        'fabric': fabric_fa,
        'yarn': pd.to_numeric(df[C_COUNT], errors='coerce').astype('Int64'),
        'tension': df[C_TENS].astype(str).str.strip(),
        'rep': pd.to_numeric(df[C_REP], errors='coerce').astype('Int64'),
        'w1': pd.to_numeric(df[C_W1], errors='coerce'),
        'w2': pd.to_numeric(df[C_W2], errors='coerce'),
    })
    return out


def specimen_table(df):
    """جدول نتیجه‌ی یک پارچه در سطح نمونه، با ستون آبرفت افزوده (RTIΔ)."""
    d = df.copy()
    d['delta'] = (d['w2'] - d['w1']).round(2)
    d = d.sort_values(['yarn', 'tension', 'rep'],
                      key=lambda s: s.map({t: i for i, t in enumerate(TENS_ORDER_FA)})
                      if s.name == 'tension' else s)
    rows = [['کد نمونه', 'نمره نخ (Ne)', 'سطح کشش', 'تکرار',
             'RTI شستشوی اول (%)', 'RTI شستشوی دوم (%)', 'RTIΔ (واحد درصد)']]
    for _, r in d.iterrows():
        rows.append([r['code'], int(r['yarn']), r['tension'], int(r['rep']),
                     f'{r["w1"]:.1f}', f'{r["w2"]:.1f}', f'{r["delta"]:+.1f}'])
    return rows


def doe_rows(wide):
    """طرح کامل DOE: ۱۲ خانه‌ی آزمایشی با کد نمونه‌های هر خانه."""
    rows = [['ردیف', 'نوع پارچه', 'نمره نخ دوخت', 'سطح کشش دوخت', 'کد نمونه‌ها', 'تعداد']]
    n = 0
    for fabric in ['پنبه', 'پلی‌استر']:
        for yarn in (20, 40):
            for tens in TENS_ORDER_FA:
                sub = wide[(wide['fabric'] == fabric) & (wide['yarn'] == yarn)
                           & (wide['tension'] == tens)]
                if sub.empty:
                    continue
                n += 1
                codes = '، '.join(sorted(sub['code'].unique()))
                rows.append([n, fabric, f'Ne {yarn}', tens, codes, len(sub)])
    return rows


def sheet_rows(sheet):
    df = pd.read_excel(SRC, sheet_name=sheet, header=None)
    rows = []
    for _, r in df.iterrows():
        vals = [('' if pd.isna(v) else str(v).strip()) for v in r.tolist()]
        vals = [v for v in vals if v]
        if vals:
            rows.append(vals)
    return rows


def lookup(rows, keywords):
    for row in rows:
        head = row[0]
        if any(k.lower() in head.lower() for k in keywords):
            return ' ، '.join(row[1:]) if len(row) > 1 else row[0]
    return '—'


def fabric_spec_table():
    """جدول مشخصات فیزیکی دو پارچه‌ی پایه، استخراج‌شده از کاربرگ‌های پارچه."""
    cot = sheet_rows('پارچه‌ی پنبه')
    pol = sheet_rows('پارچه پلی استر')
    spec = [
        ('وزن سطحی (GSM)', ['گرمی', 'gsm', 'وزن']),
        ('تراکم تار (نخ/سانتی‌متر)', ['تراکم تار', 'تار']),
        ('تراکم پود (نخ/سانتی‌متر)', ['تراکم پود', 'پود']),
        ('نمره خطی نخ تار (Tex)', ['tex', 'نمره نخ تار', 'نمره تار']),
        ('نمره خطی نخ پود (Tex)', ['نمره نخ پود', 'نمره پود']),
        ('آبرفت پارچه‌ی شاهد (%)', ['آبرفت', 'شاهد']),
    ]
    rows = [['مشخصه', 'پارچه‌ی پنبه', 'پارچه‌ی پلی‌استر', 'روش / مرجع']]
    for label, keys in spec:
        rows.append([label, lookup(cot, keys), lookup(pol, keys), 'کاربرگ پارچه — فایل داده‌ی اولیه'])
    return rows, cot, pol


# ===========================================================================
# ============================ سند ۱: گزارش =================================
# ===========================================================================
def build_report():
    if '--dump' in sys.argv:
        for s in ('پارچه‌ی پنبه', 'پارچه پلی استر'):
            print(f'\n========== {s} ==========')
            for r in sheet_rows(s):
                print(' | '.join(r))
        spec, _, _ = fabric_spec_table()
        print('\n========== نتیجه‌ی استخراج جدول ۴ ==========')
        for r in spec:
            print(' | '.join(str(c) for c in r))
        return

    wide = pd.concat([load_samples('نمونه پنبه', 'پنبه'),
                      load_samples('نمونه پلی استر', 'پلی‌استر')], ignore_index=True)
    print(f'[داده] {len(wide)} نمونه‌ی دوخته‌شده از «{SRC}»')

    doc = Document(SRC_REPORT)
    log('سند مبنا', f'{SRC_REPORT} → {DST_REPORT}')

    # ---------------------------------------------------------------- ۱) عنوان
    t = find_par(doc, 'اثر پارامترهای نخ و پارچه بر آبرفتگی دوخت', exact=True)
    if t is not None:
        set_par_text(t, T.TITLE)
        log('اصلاح ۱ — عنوان', f'«اثر پارامترهای نخ و پارچه بر آبرفتگی دوخت» ← «{T.TITLE}»')

    # ---------------------------------------------------------------- ۲) شماره‌گذاری سرتیترها و جدول‌ها
    apply_rename_map(doc, {
        '۱. مقدمه و هدف پژوهش': '۱. مقدمه',
        '۲. مواد، طرح آزمایش و آماده‌سازی داده‌ها': '۳. متدولوژی و تعریف متغیرها',
        '۳. آمار توصیفی': '۴. آمار توصیفی و نتایج تفکیکی',
        '۴. مقایسه‌ی نموداری': '۵. مقایسه‌ی نموداری',
        '۵. تحلیل واریانس': '۶. تحلیل واریانس',
        '۶. نتیجه‌گیری': '۷. نتیجه‌گیری',
        '۵-۱. اثر نوع پارچه': '۶-۱. اثر نوع پارچه',
        '۵-۲. اثر سطح کشش دوخت': '۶-۲. اثر سطح کشش دوخت',
        '۵-۳. اثر نوبت شستشو': '۶-۳. اثر نوبت شستشو',
        '۵-۴. اثر نمره نخ دوخت': '۶-۴. اثر نمره نخ دوخت',
        '۵-۵. بررسی مفروضات و تأیید ناپارامتری': '۶-۵. بررسی مفروضات و تأیید ناپارامتری',
        'جدول ۱. طرح آزمایش پژوهش آبرفتگی دوخت.': 'جدول ۲. طرح آزمایش پژوهش.',
        'جدول ۲. آبرفت دوخت به تفکیک نوع پارچه (تجمیع هر دو شستشو).':
            'جدول ۵. آبرفت دوخت به تفکیک نوع پارچه (تجمیع هر دو شستشو).',
        'جدول ۳. آبرفت دوخت به تفکیک نوع پارچه و نوبت شستشو.':
            'جدول ۶. آبرفت دوخت به تفکیک نوع پارچه و نوبت شستشو.',
        'جدول ۴. آبرفت دوخت به تفکیک سطح کشش دوخت و نوبت شستشو.':
            'جدول ۷. آبرفت دوخت به تفکیک سطح کشش دوخت و نوبت شستشو.',
        'جدول ۵. آبرفت دوخت به تفکیک نوع پارچه، سطح کشش و نوبت شستشو.':
            'جدول ۸. آبرفت دوخت به تفکیک نوع پارچه، سطح کشش و نوبت شستشو.',
        'جدول ۶. آبرفت دوخت به تفکیک نمره نخ دوخت و نوبت شستشو.':
            'جدول ۹. آبرفت دوخت به تفکیک نمره نخ دوخت و نوبت شستشو.',
        'جدول ۷. آبرفت دوخت به تفکیک نوبت شستشو (تجمیع همه‌ی نمونه‌ها).':
            'جدول ۱۰. آبرفت دوخت به تفکیک نوبت شستشو (تجمیع همه‌ی نمونه‌ها).',
        'جدول ۸. آمار همه‌ی خانه‌های طرح عاملی (پارچه × نمره نخ × کشش × شستشو).':
            'جدول ۱۱. آمار همه‌ی خانه‌های طرح عاملی (پارچه × نمره نخ × کشش × شستشو).',
        'جدول ۹. نتایج تحلیل واریانس یک‌طرفه برای همه‌ی عامل‌های بررسی‌شده.':
            'جدول ۱۴. نتایج تحلیل واریانس یک‌طرفه برای همه‌ی عامل‌های بررسی‌شده.',
        'جدول ۱۰. مقایسه‌های دوبه‌دوی سطوح کشش با تصحیح بونفرونی.':
            'جدول ۱۵. مقایسه‌های دوبه‌دوی سطوح کشش با تصحیح بونفرونی.',
        'جدول ۱۱. تأیید ناپارامتری نتایج با آزمون کروسکال-والیس.':
            'جدول ۱۶. تأیید ناپارامتری نتایج با آزمون کروسکال-والیس.',
    })
    if replace_sub(doc, 'جدول‌های ۲ تا ۸', 'جدول‌های ۵ تا ۱۱'):
        log('ارجاع متنی', '«جدول‌های ۲ تا ۸» ← «جدول‌های ۵ تا ۱۱»')

    # ---------------------------------------------------------------- ۳) مقدمه‌ی قیفی
    h_intro = find_par(doc, '۱. مقدمه', exact=True)
    h_meth = find_par(doc, '۳. متدولوژی و تعریف متغیرها', exact=True)
    removed = clear_section(doc, '۱. مقدمه', '۳. متدولوژی و تعریف متغیرها')

    els = capture(doc, lambda d: (
        [P(d, s) for s in T.R_INTRO]
        + [P(d, T.R_INTRO_LEAD)]
        + [BUL(d, q) for q in T.R_QUESTIONS]
    ))
    place_after(h_intro, els)
    log('اصلاح ۲ — مقدمه',
        f'بازنویسی کامل با ساختار قیفی (اهمیت پایداری ابعادی ← سازوکارهای فیزیکی ← '
        f'ضرورت بررسی متغیرها)؛ {removed} پاراگراف قدیمی جایگزین و ۴ پرسش پژوهش درج شد.')

    # ---------------------------------------------------------------- راهنمای نشانه‌گذاری
    def _legend(d):
        H(d, T.LEGEND_LEAD, 2)
        for s in T.LEGEND:
            P(d, s)
    els = capture(doc, _legend)
    place_before(h_intro, els)
    log('نشانه‌گذاری', 'راهنمای نشانه‌گذاری نسخه (هایلایت زرد) پیش از بخش ۱ درج شد.')

    # ---------------------------------------------------------------- ۴) پیشینه‌ی تحقیق
    lit = []
    lit.append(lambda d: PAGEBREAK(d))
    lit.append(lambda d: H(d, '۲. پیشینه تحقیق', 1))
    lit.append(lambda d: P(d, T.R_LIT_INTRO))
    lit.append(lambda d: NOTE(d, T.R_LIT_NOTE))
    for head, paras, refs, kw in T.R_LIT:
        lit.append((lambda h: (lambda d: H(d, h, 2)))(head))
        for pa in paras:
            lit.append((lambda x: (lambda d: P(d, x)))(pa))
        for rf in refs:
            lit.append((lambda x: (lambda d: REF(d, x)))(rf))
        if kw:
            lit.append((lambda x: (lambda d: P(d, '◄ ' + x, size=10.5, italic=True,
                                                color=BLUEC, after=8)))(kw))

    def _lit(d):
        for fn in lit:
            fn(d)
    els = capture(doc, _lit)
    place_before(h_meth, els)
    nrefs = sum(len(r) for _, _, r, _ in T.R_LIT)
    log('اصلاح ۳ — پیشینه تحقیق',
        f'بخش تازه «۲. پیشینه تحقیق» در ۴ زیربخش با {nrefs} مدخل استناد (TRJ / JTI / AUTEX) '
        f'و ارجاع به AATCC 88B و ISO 7770 درج شد؛ کلیدواژه‌ی جست‌وجو برای هر محور اضافه شد.')

    # ---------------------------------------------------------------- ۵) متدولوژی: متغیرها
    iv_rows = [
        ['متغیر مستقل', 'نوع', 'سطوح', 'نحوه‌ی کنترل / اندازه‌گیری'],
        ['نوع پارچه', 'کیفی اسمی', '۲ سطح: پنبه، پلی‌استر',
         'انتخاب دو پارچه‌ی پایه‌ی متمایز؛ جنس الیاف کنترل‌شده و ثابت در سراسر آزمایش'],
        ['نمره‌ی نخ دوخت', 'کیفی رتبه‌ای', '۲ سطح: Ne 20، Ne 40',
         'تأمین از قرقره‌های شماره‌دار؛ نمره‌ی واقعی نخ در کاربرگ‌های دوک اندازه‌گیری شده است'],
        ['سطح کشش دوخت', 'کیفی رتبه‌ای', '۳ سطح: کم، متوسط، زیاد',
         'تنظیم کشش چرخ خیاطی و ثابت نگه‌داشتن آن در هر سطح برای همه‌ی نمونه‌های آن سطح'],
        ['تعداد شستشو', 'کیفی ترتیبی', '۲ سطح: شستشوی اول، شستشوی دوم',
         'برنامه‌ی شستشوی یکسان با دمای ثابت؛ اندازه‌گیری مکرر روی همان نمونه'],
    ]
    dv_rows = [
        ['متغیر وابسته', 'نماد', 'واحد', 'دامنه', 'روش تعیین'],
        ['شاخص آبرفت درز پس از شستشوی اول', 'RTI_W1', 'درصد', '۰ تا ۱۰۰', 'رابطه‌ی (۱)'],
        ['شاخص آبرفت درز پس از شستشوی دوم', 'RTI_W2', 'درصد', '۰ تا ۱۰۰', 'رابطه‌ی (۲)'],
        ['آبرفت افزوده‌ی شستشوی دوم', 'RTIΔ', 'واحد درصد', '−۱۰۰ تا +۱۰۰', 'رابطه‌ی (۳)'],
        ['نمره‌ی بصری چروکیدگی درز', 'Visual Grade', 'امتیاز', '۱ تا ۵',
         'ارزیابی چشمی طبق AATCC 88B / ISO 7770 (فقط نمونه‌های شاهد)'],
    ]

    def _vars(d):
        H(d, T.R_IV_LEAD, 2)
        P(d, 'عامل‌هایی که پژوهشگر آن‌ها را تغییر می‌دهد و اثرشان بر شاخص آبرفت سنجیده می‌شود:')
        TABLE(d, iv_rows, size=10.5)
        P(d, '')
        H(d, T.R_DV_LEAD, 2)
        P(d, 'کمیت‌هایی که به‌عنوان پاسخ اندازه‌گیری می‌شوند:')
        TABLE(d, dv_rows, size=10.5)
        NOTE(d, T.R_DV_NOTE)
        NOTE(d, T.R_VG_NOTE)
    els = capture(doc, _vars)
    cap_design = find_par(doc, 'جدول ۲. طرح آزمایش پژوهش.', exact=True)
    place_before(cap_design, els)
    log('اصلاح ۴ — متغیرها',
        'دو جدول تعریف متغیرها درج شد: جدول مستقل‌ها (۴ عامل با سطوح و نحوه‌ی کنترل) و '
        'جدول وابسته‌ها (RTI_W1، RTI_W2، RTIΔ و Visual Grade با واحد و دامنه).')

    # ---------------------------------------------------------------- ۶) متدولوژی: DOE، فرمول‌ها، استراحت
    doe = doe_rows(wide)
    rest = []

    def _mid(d):
        H(d, T.R_FORMULA_LEAD, 2)
        CAPTION(d, 'جدول ۳. طرح کامل DOE — ۱۲ خانه‌ی آزمایشی و کد نمونه‌های هر خانه.')
        TABLE(d, doe, size=10.5)
        NOTE(d, T.R_DOE_NOTE)
        P(d, 'شاخص‌های وابسته از روابط زیر محاسبه می‌شوند:')
        FORMULA(d, 'RTI_W1 = ((L0 − L_washed1) / L0) × 100        (۱)')
        FORMULA(d, 'RTI_W2 = ((L0 − L_washed2) / L0) × 100        (۲)')
        FORMULA(d, 'RTIΔ  = RTI_W2 − RTI_W1                        (۳)')
        NOTE(d, T.R_L0_NOTE)
        H(d, T.R_REST_LEAD, 2)
        for s in T.R_REST:
            P(d, s)
        P(d, T.R_VG_SIDE)
    els = capture(doc, _mid)
    anchor = find_par(doc, 'هر یک از ۱۲ خانه‌ی طرح', exact=False)
    if anchor is None:
        anchor = find_par(doc, 'متغیر پاسخ، درصد آبرفت کل دوخت', exact=False)
    place_after(anchor, els)
    log('اصلاح ۴ — DOE و فرمول‌ها',
        f'جدول کامل DOE با {len(doe)-1} خانه‌ی آزمایشی و کد نمونه‌ها، سه رابطه‌ی '
        f'RTI_W1 / RTI_W2 / RTIΔ و پروتکل استراحت ۲۴ ساعته (ISO 139) درج شد.')

    old_resp = find_par(doc, 'متغیر پاسخ، درصد آبرفت کل دوخت', exact=False)
    if old_resp is not None:
        delete_par(old_resp)
        log('حذف/ادغام', 'پاراگراف قدیمی «متغیر پاسخ، درصد آبرفت کل دوخت...» با بخش '
                         'شاخص‌های وابسته و روابط محاسباتی جایگزین شد.')

    # ---------------------------------------------------------------- ۷) مشخصات فیزیکی پارچه
    spec_rows, cot_raw, pol_raw = fabric_spec_table()

    def _spec(d):
        H(d, T.R_FAB_SPEC_LEAD, 2)
        P(d, 'مشخصات فیزیکی دو پارچه‌ی پایه، که تعیین‌کننده‌ی سازوکار غالب جمع‌شدگی است:')
        CAPTION(d, 'جدول ۴. مشخصات فیزیکی پارچه‌های پایه.')
        TABLE(d, spec_rows, size=10.5)
        NOTE(d, T.R_FAB_SPEC_NOTE)
    els = capture(doc, _spec)
    cap_t1 = find_par(doc, 'جدول ۵. آبرفت دوخت به تفکیک نوع پارچه', exact=False)
    place_before(cap_t1, els)
    log('اصلاح ۵ — مشخصات پارچه',
        'جدول ۴ (مشخصات فیزیکی پارچه) با GSM، تراکم تار و پود، نمره‌ی خطی Tex و '
        'آبرفت پارچه‌ی شاهد درج شد.')

    # ---------------------------------------------------------------- ۸) جدول‌های تفکیکی پنبه / پلی‌استر
    cot_tbl = specimen_table(wide[wide['fabric'] == 'پنبه'])
    pol_tbl = specimen_table(wide[wide['fabric'] == 'پلی‌استر'])

    def _sep(d):
        H(d, T.R_DESC_SEP_LEAD, 2)
        P(d, T.R_DESC_SEP)
        CAPTION(d, 'جدول ۱۲. نتایج تفکیکی نمونه‌های پارچه‌ی پنبه (۱۸ نمونه).')
        TABLE(d, cot_tbl, size=9.5)
        NOTE(d, 'مقادیر بر حسب درصد طول اولیه‌ی درز (L0 = ۲۰۰ میلی‌متر) هستند. RTIΔ مثبت یعنی '
                'نمونه در شستشوی دوم بیشتر جمع شده است.')
        CAPTION(d, 'جدول ۱۳. نتایج تفکیکی نمونه‌های پارچه‌ی پلی‌استر (۱۸ نمونه).')
        TABLE(d, pol_tbl, size=9.5)
        NOTE(d, 'همان قالب جدول ۱۲. مقایسه‌ی دو جدول، تفاوت سطح مقادیر میان دو پارچه را '
                'به‌روشنی نشان می‌دهد.')
    els = capture(doc, _sep)
    concl_desc = find_par(doc, 'تصویر توصیفی از همین‌جا روشن است', exact=False)
    h33 = capture(doc, lambda d: H(d, T.R_DESC_CONCL_LEAD, 2))
    place_before(concl_desc, els + h33)
    log('اصلاح ۵ — جدول‌های تفکیکی',
        f'جدول ۱۲ (پنبه، {len(cot_tbl)-1} نمونه) و جدول ۱۳ (پلی‌استر، {len(pol_tbl)-1} نمونه) '
        f'با ستون‌های RTI شستشوی اول، RTI شستشوی دوم و RTIΔ درج شد.')

    # ---------------------------------------------------------------- ۹) جای‌نگهدار تصاویر
    def _ph(d):
        H(d, '۵-۱. تصاویر آزمایشگاهی (جای‌نگهدار)', 2)
        P(d, T.R_PLACEHOLDER_LEAD)
        for label, note in T.R_PLACEHOLDERS:
            PLACEHOLDER(d, label, note)
    els = capture(doc, _ph)
    h_anova = find_par(doc, '۶. تحلیل واریانس', exact=True)
    place_before(h_anova, els)
    log('اصلاح ۶ — جای‌نگهدار تصاویر',
        f'{len(T.R_PLACEHOLDERS)} کادر جای‌نگهدار (تصویر ۱ تا ۴) با برچسب و شرح درج شد.')

    # ---------------------------------------------------------------- ۱۰) پیوست
    def _app(d):
        PAGEBREAK(d)
        H(d, T.R_APPENDIX_LEAD, 1)
        for s in T.R_APPENDIX:
            P(d, s)
    anchor_end = last_body_par(doc)
    els = capture(doc, _app)
    place_after(anchor_end, els)
    log('پیوست', 'بخش «۸. پیوست: فهرست تصاویر موردنیاز» شامل قالب نام‌گذاری ۱۰۸ تصویر درج شد.')

    save(doc, DST_REPORT)


# ===========================================================================
# ============================ سند ۲: راهنما ================================
# ===========================================================================
GUIDE_TOC = [
    'این پروژه درباره‌ی چیست؟', 'پیشینه تحقیق و جایگاه این پژوهش',
    'فایل اکسل اولیه چه چیزی در خود داشت؟', 'طرح آزمایش: چرا دقیقاً ۳۶ نمونه؟',
    'متغیرهای پژوهش و شاخص RTI چگونه محاسبه می‌شوند؟', 'چه داده‌هایی کنار گذاشته شد و چرا؟',
    'مفاهیم آماری به زبان ساده', 'نمودارها یکی‌یکی: هر کدام چه می‌گویند؟',
    'جدول‌ها چه چیزی نشان می‌دهند؟', 'نتایج اصلی و تفسیر فنی آن‌ها',
    'محدودیت‌های پژوهش', 'راهنمای ارائه: چه بگویید و به چه سؤال‌هایی آماده باشید',
    'فایل‌های پروژه و نحوه‌ی اجرای دوباره',
]

FA_DIGITS = '۰۱۲۳۴۵۶۷۸۹'


def fa_num(n):
    return ''.join(FA_DIGITS[int(c)] if c.isdigit() else c for c in str(n))


def build_guide():
    if '--dump' in sys.argv:
        return

    doc = Document(SRC_GUIDE)
    log('سند مبنا', f'{SRC_GUIDE} → {DST_GUIDE}')

    # ---------------------------------------------------------------- عنوان
    sub = find_par(doc, 'اثر پارامترهای نخ و پارچه بر آبرفتگی دوخت در پارچه‌های پنبه و پلی‌استر',
                   exact=True)
    if sub is not None:
        set_par_text(sub, T.TITLE)
        log('اصلاح ۱ — عنوان', f'زیرعنوان راهنما به «{T.TITLE}» تغییر یافت.')

    # ---------------------------------------------------------------- شماره‌گذاری بخش‌ها (سرتیترها)
    renames = {
        'بخش ۲: فایل اکسل اولیه چه چیزی در خود داشت؟': 'بخش ۳: فایل اکسل اولیه چه چیزی در خود داشت؟',
        'بخش ۳: طرح آزمایش: چرا دقیقاً ۳۶ نمونه؟': 'بخش ۴: طرح آزمایش: چرا دقیقاً ۳۶ نمونه؟',
        'بخش ۴: متغیر پاسخ — «درصد آبرفت» چگونه محاسبه می‌شود؟':
            'بخش ۵: متغیرهای پژوهش و شاخص RTI چگونه محاسبه می‌شوند؟',
        'بخش ۵: چه داده‌هایی کنار گذاشته شد و چرا؟': 'بخش ۶: چه داده‌هایی کنار گذاشته شد و چرا؟',
        'بخش ۶: مفاهیم آماری به زبان ساده': 'بخش ۷: مفاهیم آماری به زبان ساده',
        'بخش ۷: نمودارها یکی‌یکی — هر کدام چه می‌گویند؟':
            'بخش ۸: نمودارها یکی‌یکی — هر کدام چه می‌گویند؟',
        'بخش ۸: جدول‌ها چه چیزی نشان می‌دهند؟': 'بخش ۹: جدول‌ها چه چیزی نشان می‌دهند؟',
        'بخش ۹: نتایج اصلی و تفسیر فنی آن‌ها': 'بخش ۱۰: نتایج اصلی و تفسیر فنی آن‌ها',
        'بخش ۱۰: محدودیت‌های پژوهش': 'بخش ۱۱: محدودیت‌های پژوهش',
        'بخش ۱۱: راهنمای ارائه': 'بخش ۱۲: راهنمای ارائه',
        'بخش ۱۲: فایل‌های پروژه و اجرای دوباره': 'بخش ۱۳: فایل‌های پروژه و اجرای دوباره',
    }
    # فقط سرتیترهای واقعی (سبک Normal) تغییر می‌کنند؛ فهرست مطالب جدا بازنویسی می‌شود.
    for p in doc.paragraphs:
        if p.style.name != 'Normal':
            continue
        if p.text in renames:
            set_par_text(p, renames[p.text])
    log('شماره‌گذاری', 'بخش‌های راهنما از ۲..۱۲ به ۳..۱۳ جابه‌جا شد تا جای بخش تازه‌ی '
                       '«پیشینه تحقیق» باز شود.')

    # ---------------------------------------------------------------- فهرست مطالب
    toc_head = find_par(doc, 'مطالب', exact=False)
    old_bullets = [p for p in doc.paragraphs
                   if p.style.name == 'List Bullet' and p.text.startswith('بخش ')]
    for b in old_bullets:
        delete_par(b)
    # اقلام فهرست مطالب عمداً هایلایت نمی‌شوند؛ نشانه‌گذاری زرد باید جای تصاویر
    # و مراجع را نشان دهد، نه آنکه کل فهرست را زرد کند.
    hl_off()
    els = capture(doc, lambda d: [BUL(d, f'بخش {fa_num(i)}: {t}')
                                  for i, t in enumerate(GUIDE_TOC, start=1)])
    hl_on()
    place_after(toc_head, els)
    log('فهرست مطالب', f'فهرست مطالب با {len(GUIDE_TOC)} بخش بازنویسی شد '
                       f'(بدون هایلایت، تا نشانه‌گذاری زرد گویا بماند).')
    # ---------------------------------------------------------------- بازنویسی بخش ۱
    # سرتیترها سبک Normal دارند و فهرست مطالب سبک List Bullet؛ بنابراین جست‌وجوی
    # سرتیتر باید با فیلتر سبک انجام شود تا با اقلام فهرست اشتباه گرفته نشود.
    h1 = find_par(doc, 'بخش ۱: این پروژه درباره‌ی چیست؟', exact=True, style='Normal')
    h3 = find_par(doc, 'بخش ۳: فایل اکسل اولیه چه چیزی در خود داشت؟', exact=True, style='Normal')
    assert h1 is not None and h3 is not None, 'سرتیتر بخش‌های ۱ یا ۳ پیدا نشد.'

    # پاک‌کردن محتوای قبلی بخش ۱ (سرتیتر و جدول‌ها می‌مانند)
    body = doc.element.body
    kids = list(body)
    i = kids.index(h1._p)
    j = kids.index(h3._p)
    removed = 0
    for el in kids[i + 1:j]:
        if el.tag == qn('w:p'):
            body.remove(el)
            removed += 1

    def _g1(d):
        for s in T.G_INTRO:
            P(d, s)
        P(d, 'پرسش‌های پژوهش:')
        for q in T.R_QUESTIONS:
            BUL(d, q)
        P(d, 'پاسخ کوتاه پروژه، که در بخش‌های بعد مستند می‌شود، این است: جنس پارچه بسیار مهم است، '
             'کشش دوخت فقط وقتی زیاد باشد و آن هم عمدتاً روی پلی‌استر مهم است، شستشوی دوم قطعاً '
             'آبرفت را افزایش می‌دهد، و نمره‌ی نخ در بازه‌ی بررسی‌شده اثر معناداری ندارد.')
    els = capture(doc, _g1)
    place_after(h1, els)
    log('اصلاح ۲ — مقدمه‌ی راهنما',
        f'بخش ۱ با ساختار قیفی بازنویسی شد ({removed} پاراگراف قدیمی جایگزین شد).')

    # ---------------------------------------------------------------- راهنمای نشانه‌گذاری
    def _legend(d):
        H(d, T.LEGEND_LEAD, 2)
        for s in T.LEGEND:
            P(d, s)
    els = capture(doc, _legend)
    place_before(h1, els)
    log('نشانه‌گذاری (راهنما)', 'راهنمای نشانه‌گذاری نسخه پیش از بخش ۱ درج شد.')

    # ---------------------------------------------------------------- بخش تازه‌ی پیشینه
    h3 = find_par(doc, 'بخش ۳: فایل اکسل اولیه چه چیزی در خود داشت؟', exact=True, style='Normal')
    assert h3 is not None, 'سرتیتر بخش ۳ پیدا نشد.'

    def _gl(d):
        H(d, 'بخش ۲: پیشینه تحقیق و جایگاه این پژوهش', 1)
        P(d, 'سه محور پژوهشی، پشتوانه‌ی نظری این پروژه را می‌سازند. در ارائه، دانستن این سه محور '
             'به شما کمک می‌کند نشان دهید کارتان در ادامه‌ی ادبیات موجود است، نه جدا از آن.')
        for head, paras, refs, kw in T.R_LIT:
            H(d, head.replace('۲-', '۲-'), 2)
            for pa in paras:
                P(d, pa)
            for rf in refs:
                REF(d, rf)
            if kw:
                P(d, '◄ ' + kw, size=10.5, italic=True, color=BLUEC, after=8)
        NOTE(d, T.R_LIT_NOTE)
    els = capture(doc, _gl)
    place_before(h3, els)
    log('اصلاح ۳ — پیشینه تحقیق (راهنما)',
        'بخش تازه‌ی «۲. پیشینه تحقیق» با همان شش مدخل استناد و توضیح ساده درج شد.')
    return doc


def finish_guide(doc):
    """افزودن شاخص RTI، پروتکل استراحت و جای‌نگهدارها به راهنما."""
    # --------- بخش ۵ (متغیرها): درج روابط RTI و پروتکل استراحت
    anchor = find_par(doc, 'یک نکته‌ی مهم که حتماً باید بدانید', exact=False)
    if anchor is None:
        anchor = find_par(doc, 'درصد آبرفت = (طول اولیه منهای طول پس از شستشو)', exact=False)

    def _f(d):
        P(d, 'در این گزارش، همین کمیت با نام «شاخص آبرفت درز» و نماد RTI گزارش می‌شود. روابط '
             'دقیق محاسبه چنین است:')
        FORMULA(d, 'RTI_W1 = ((L0 − L_washed1) / L0) × 100        (۱)')
        FORMULA(d, 'RTI_W2 = ((L0 − L_washed2) / L0) × 100        (۲)')
        FORMULA(d, 'RTIΔ  = RTI_W2 − RTI_W1                        (۳)')
        NOTE(d, 'L0 طول نشانه‌گذاری‌شده‌ی درز پس از دوختن و پیش از شستشو است (۲۰۰ میلی‌متر)؛ '
                'L_washed1 و L_washed2 همان طول پس از شستشوی اول و دوم‌اند. RTI_W1 و RTI_W2 بر حسب '
                'درصد و RTIΔ بر حسب واحد درصد گزارش می‌شوند.')
        H(d, 'پروتکل استراحت نمونه‌ها', 2)
        P(d, 'پس از هر مرحله‌ی شستشو، نمونه‌ها پیش از اندازه‌گیری به‌مدت ۲۴ ساعت روی سطح صاف و در '
             'شرایط استاندارد آزمایشگاه (۲۰±۲ درجه‌ی سلسیوس، رطوبت نسبی ۶۵±۲ درصد، مطابق ISO 139) '
             'استراحت داده می‌شوند تا رطوبت به حالت تعادل برسد و تنش‌های بازآرایی‌شده تثبیت شوند. '
             'اگر داور پرسید «چرا بلافاصله بعد از شستشو اندازه نگرفتید؟»، پاسخ همین است: اندازه‌گیری '
             'در حالت مرطوب آبرفت را بیش از واقع نشان می‌دهد.')
    els = capture(doc, _f)
    place_after(anchor, els)
    log('اصلاح ۴ — روابط و پروتکل (راهنما)',
        'سه رابطه‌ی RTI و پروتکل استراحت ۲۴ ساعته به بخش متغیرها افزوده شد.')

    # --------- بخش ۸ (نمودارها): درج جای‌نگهدارها پیش از بخش ۹
    h9 = find_par(doc, 'بخش ۹: جدول‌ها چه چیزی نشان می‌دهند؟', exact=True, style='Normal')
    assert h9 is not None, 'سرتیتر بخش ۹ پیدا نشد.'

    def _ph(d):
        PAGEBREAK(d)
        H(d, 'بخش ۸-۱: تصاویر آزمایشگاهی موردنیاز (جای‌نگهدار)', 2)
        P(d, T.R_PLACEHOLDER_LEAD)
        for label, note in T.R_PLACEHOLDERS:
            PLACEHOLDER(d, label, note)
        P(d, 'افزون بر این چهار تصویر، برای هر نمونه در هر تکرار سه عکس لازم است: پس از دوختن، '
             'پس از شستشوی اول و پس از شستشوی دوم. با ۳۶ نمونه، در مجموع ۱۰۸ تصویر به دست می‌آید. '
             'قالب نام‌گذاری پیشنهادی: [کد نمونه]_[فاز].jpg؛ برای نمونه C20L-1_wash1.jpg. تصاویر '
             'برای مستندسازی اجرای آزمایش لازم‌اند و به اعتبار محاسبات آماری که در بخش‌های بعد '
             'آمده، خللی وارد نمی‌کنند.')
        H(d, 'پاسخ آماده برای داور', 3)
        P(d, 'اگر پرسیدند «چرا تصویر میکروسکوپی یا عکس نمونه‌ها را نیاوردی؟»، پاسخ این است که '
             'محاسبات این گزارش بر پایه‌ی اندازه‌گیری طول نشانه‌گذاری‌شده است و به تصویر وابسته '
             'نیست؛ اما جای تصاویر در سند مشخص و برچسب‌گذاری شده است تا در نسخه‌ی نهایی تکمیل شود.')
    els = capture(doc, _ph)
    place_before(h9, els)
    log('اصلاح ۶ — جای‌نگهدار تصاویر (راهنما)',
        'زیربخش «۸-۱» با چهار کادر جای‌نگهدار و توضیح ۱۰۸ تصویر پیوست درج شد.')

    save(doc, DST_GUIDE)


# ===========================================================================
# بازرسی سندهای تولیدشده
# ===========================================================================
R_CHECKS = [
 ('اصلاح ۱ — عنوان تازه', T.TITLE),
 ('اصلاح ۲ — مقدمه‌ی قیفی (راهنمای نشانه‌گذاری)', T.LEGEND_LEAD),
 ('اصلاح ۲ — پرسش پژوهش ۱', T.R_QUESTIONS[0]),
 ('اصلاح ۳ — بخش ۲. پیشینه تحقیق', '۲. پیشینه تحقیق'),
 ('اصلاح ۳ — استناد TRJ', 'Textile Research Journal'),
 ('اصلاح ۳ — استناد JTI', 'Journal of the Textile Institute'),
 ('اصلاح ۳ — استناد AUTEX', 'AUTEX Research Journal'),
 ('اصلاح ۳ — استاندارد AATCC 88B', 'AATCC 88B'),
 ('اصلاح ۳ — استاندارد ISO 7770', 'ISO 7770'),
 ('اصلاح ۴ — سرتیتر متغیرهای مستقل', T.R_IV_LEAD),
 ('اصلاح ۴ — سرتیتر متغیرهای وابسته', T.R_DV_LEAD),
 ('اصلاح ۴ — طرح کامل DOE', 'طرح کامل DOE'),
 ('اصلاح ۴ — رابطه‌ی RTI_W1', 'RTI_W1 = ((L0'),
 ('اصلاح ۴ — رابطه‌ی RTI_W2', 'RTI_W2 = ((L0'),
 ('اصلاح ۴ — رابطه‌ی RTIΔ', 'RTIΔ'),
 ('اصلاح ۴ — پروتکل استراحت ۲۴ ساعته', '۲۴ ساعت'),
 ('اصلاح ۵ — مشخصات فیزیکی پارچه', T.R_FAB_SPEC_LEAD),
 ('اصلاح ۵ — جدول تفکیکی پنبه', 'جدول ۱۲'),
 ('اصلاح ۵ — جدول تفکیکی پلی‌استر', 'جدول ۱۳'),
 ('اصلاح ۶ — جای‌نگهدار تصویر ۱', '[تصویر ۱:'),
 ('اصلاح ۶ — جای‌نگهدار تصویر ۲', '[تصویر ۲:'),
 ('اصلاح ۶ — جای‌نگهدار تصویر ۳', '[تصویر ۳:'),
 ('اصلاح ۶ — جای‌نگهدار تصویر ۴', '[تصویر ۴:'),
 ('پیوست تصاویر', T.R_APPENDIX_LEAD),
 ('متن راهبری: مرجع ناتکمیل', '[نام خانوادگی]'),
]

G_CHECKS = [
 ('اصلاح ۱ — عنوان تازه', T.TITLE),
 ('راهنمای نشانه‌گذاری', T.LEGEND_LEAD),
 ('اصلاح ۲ — بخش ۱ قیفی', T.G_INTRO[0][:40]),
 ('اصلاح ۳ — بخش ۲: پیشینه تحقیق', 'بخش ۲: پیشینه تحقیق'),
 ('اصلاح ۳ — استناد TRJ', 'Textile Research Journal'),
 ('اصلاح ۴ — بخش ۵: متغیرها و شاخص RTI', 'بخش ۵: متغیرهای پژوهش و شاخص RTI'),
 ('اصلاح ۴ — رابطه‌ی RTI_W1', 'RTI_W1 = ((L0'),
 ('اصلاح ۴ — پروتکل استراحت', 'پروتکل استراحت نمونه‌ها'),
 ('اصلاح ۶ — بخش ۸-۱ جای‌نگهدارها', 'بخش ۸-۱:'),
 ('اصلاح ۶ — جای‌نگهدار تصویر ۱', '[تصویر ۱:'),
 ('فهرست مطالب ۱۳ بخشی', 'بخش ۱۳:'),
]

APPLIED = '  → اعمال‌شده: {} از {}'
UNFINISHED = ' | ناتمام: {}'


def doc_text(doc):
    """متن کل سند، شامل متن داخل سلول‌های جدول."""
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                parts.append(c.text)
    return '\n'.join(parts)


def table_texts(doc):
    return ['\n'.join(c.text for row in t.rows for c in row.cells)
            for t in doc.tables]


def verify():
    """بازرسی سندهای تولیدشده: کدام اصلاح اعمال شده و کدام ناتمام است."""
    for path, checks, dump_tables in ((DST_REPORT, R_CHECKS, True),
                                      (DST_GUIDE, G_CHECKS, False)):
        if not os.path.exists(path):
            print(f'[نبود] {path}')
            continue
        doc = Document(path)
        paras = doc.paragraphs
        txt = doc_text(doc)
        hl_runs = hl_paras = 0
        for p in paras:
            hit = False
            for r in p.runs:
                if r.font.highlight_color is not None:
                    hl_runs += 1
                    hit = True
            hl_paras += hit

        print('\n' + '=' * 78)
        print(f'بازرسی سند: {path}')
        print('=' * 78)
        print(f'  پاراگراف: {len(paras)} | جدول: {len(doc.tables)} | '
              f'تصویر: {len(doc.inline_shapes)} | '
              f'پاراگراف هایلایت‌شده: {hl_paras} ({hl_runs} ران)')
        missing = []
        for label, needle in checks:
            ok = needle in txt
            print(f'  [{"OK " if ok else "—  "}] {label}')
            if not ok:
                missing.append(label)
        print(APPLIED.format(len(checks) - len(missing), len(checks))
              + (UNFINISHED.format('، '.join(missing)) if missing else ''))

        print('\n  --- ۲۵ پاراگراف نخست سند (شماره، علامت * = هایلایت‌شده) ---')
        for i, p in enumerate(paras[:25]):
            mark = '*' if any(r.font.highlight_color is not None for r in p.runs) else ' '
            print(f'   {i:>3} {mark} {p.text.strip()[:88]!r}')

        for lab, needle in [('جای‌نگهدار تصویر', '[تصویر '),
                            ('مرجع ناتکمیل', '[نام خانوادگی]')]:
            hits = [(i, t) for i, t in enumerate(table_texts(doc), start=1)
                    if needle in t]
            print(f'\n  --- «{lab}» در جدول‌ها: {len(hits)} مورد ---')
            for i, t in hits[:8]:
                first = [ln for ln in t.split('\n') if needle in ln]
                print(f'   جدول {i}: {first[0][:96] if first else ""}')

        if dump_tables:
            keys = ('مشخصه', 'متغیر مستقل', 'ردیف', 'کد نمونه')
            for i, t in enumerate(doc.tables, start=1):
                head = ' | '.join(c.text.strip() for c in t.rows[0].cells)
                if not any(k in head for k in keys):
                    continue
                print(f'\n  --- جدول شماره‌ی {i} در سند: {head} ---')
                for row in t.rows[1:]:
                    print('   ' + ' | '.join(c.text.strip() for c in row.cells))


# ===========================================================================
if __name__ == '__main__':
    if '--verify' in sys.argv:
        verify()
        raise SystemExit(0)

    build_report()
    n_report = hl_count()
    HL['runs'] = 0
    guide = build_guide()
    if guide is not None:
        finish_guide(guide)
    n_guide = hl_count()

    if '--dump' not in sys.argv:
        print('\n' + '=' * 78)
        print('گزارش تغییرات')
        print('=' * 78)
        for kind, text in CHANGE_LOG:
            print(f'  • [{kind}] {text}')
        print(f'\n  ► هایلایت زرد: {n_report} ران در گزارش، {n_guide} ران در راهنما.')
        print('  ► برای بازرسی سندهای تولیدشده: python apply_corrections.py --verify')
