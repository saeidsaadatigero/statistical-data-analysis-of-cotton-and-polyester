# -*- coding: utf-8 -*-
"""ساخت نسخه‌ی کامل فارسی: نمودارها، جدول اکسل، گزارش و راهنمای مطالعه/ارائه."""
import os, io, contextlib, datetime
import numpy as np, pandas as pd

# --- محاسبات را از اسکریپت اصلی می‌گیریم تا اعداد دقیقاً یکسان باشند ---
# اسکریپت را در فضای‌نام خودمان اجرا می‌کنیم تا اگر فایل خروجی انگلیسی در Word باز
# و قفل باشد، محاسبات از دست نرود (همه‌ی محاسبات پیش از ذخیره‌سازی انجام شده است).
class _A:  # نگه‌دارنده‌ی فضای‌نام اسکریپت اصلی
    pass

_ns = {'__name__': '__analysis_core__', '__file__': 'analysis.py'}
with contextlib.redirect_stdout(io.StringIO()):
    try:
        exec(compile(open('analysis.py', encoding='utf-8').read(), 'analysis.py', 'exec'), _ns)
    except PermissionError as e:
        print(f'[هشدار] بازنویسی خروجی انگلیسی ممکن نشد (فایل باز است): {e.filename}')
A = _A()
A.__dict__.update(_ns)

long, wide = A.long, A.wide
TABLES, anova_df, posthoc_df, NP_RESULTS, paired = A.TABLES, A.anova_df, A.posthoc_df, A.NP_RESULTS, A.paired

OUT = 'Persian'
os.makedirs(OUT, exist_ok=True)

import arabic_reshaper
from bidi.algorithm import get_display

def fa(t):
    """متن فارسی برای matplotlib.

    matplotlib نسخه‌ی ۳.۱۱ خودش حروف فارسی را به هم می‌چسباند و ترتیب راست‌به‌چپ را
    اعمال می‌کند؛ بنابراین متن باید خام و بدون پردازش داده شود. اگر با نسخه‌ی قدیمی‌تر
    matplotlib اجرا شد، متغیر محیطی RESHAPE=1 را تنظیم کنید تا شکل‌دهی دستی انجام شود.
    """
    if os.environ.get('RESHAPE') == '1':
        return get_display(arabic_reshaper.reshape(str(t)))
    return str(t)

# ============================ ترجمه‌ی برچسب‌ها ============================
FAB_FA  = {'Cotton': 'پنبه', 'Polyester': 'پلی‌استر'}
TENS_FA = {'Low': 'کم', 'Medium': 'متوسط', 'High': 'زیاد'}
WASH_FA = {'Wash 1': 'شستشوی اول', 'Wash 2': 'شستشوی دوم'}
COL_FA  = {'Fabric': 'نوع پارچه', 'YarnCount': 'نمره نخ (Ne)', 'Tension': 'سطح کشش دوخت',
           'Wash': 'نوبت شستشو', 'Count': 'تعداد', 'Mean': 'میانگین', 'Std': 'انحراف معیار',
           'Min': 'کمینه', 'Max': 'بیشینه', 'Replicate': 'تکرار', 'SampleCode': 'کد نمونه',
           'Shrinkage': 'درصد آبرفت'}

def fa_frame(df):
    """یک جدول انگلیسی را به جدول کاملاً فارسی تبدیل می‌کند"""
    d = df.copy()
    for c in d.columns:
        if c == 'Fabric':  d[c] = d[c].astype(str).map(FAB_FA)
        if c == 'Tension': d[c] = d[c].astype(str).map(TENS_FA)
        if c == 'Wash':    d[c] = d[c].astype(str).map(WASH_FA)
    return d.rename(columns=COL_FA)

# ============================ نمودارهای فارسی ============================
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator

FONT = ['B Nazanin', 'Tahoma', 'DejaVu Sans']  # جایگزینی خودکار برای حروف لاتین
SURFACE, INK, INK2, GRID = '#fcfcfb', '#1a1a19', '#5c5c58', '#e4e4e0'
BLUE, ORANGE = '#2a78d6', '#eb6834'
plt.rcParams.update({
    'figure.dpi': 300, 'savefig.dpi': 300, 'font.family': FONT, 'font.size': 13,
    'figure.facecolor': SURFACE, 'axes.facecolor': SURFACE, 'savefig.facecolor': SURFACE,
    'text.color': INK, 'axes.labelcolor': INK2, 'xtick.color': INK2, 'ytick.color': INK2,
    'axes.edgecolor': GRID, 'axes.linewidth': 0.8, 'axes.unicode_minus': False,
})

YLAB = fa('درصد آبرفت کل دوخت')
FA_CHARTS = []

def style(ax, title, sub=None):
    ax.set_title(fa(title), fontsize=17, fontweight='bold', color=INK, loc='right', pad=24 if sub else 12)
    if sub:
        ax.text(1, 1.02, fa(sub), transform=ax.transAxes, fontsize=12, color=INK2, va='bottom', ha='right')
    ax.set_ylabel(YLAB, fontsize=13)
    ax.yaxis.grid(True, color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    for s in ('top', 'right', 'left'):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis='both', length=0)

def bars(ax, n, means, sds, color, width, offset=0.0, label=None):
    pos = np.arange(n) + offset
    ax.bar(pos, means, width, color=color, label=label, zorder=3, edgecolor=SURFACE, linewidth=2)
    ax.errorbar(pos, means, yerr=sds, fmt='none', ecolor=INK2, elinewidth=1.1,
                capsize=4, capthick=1.1, zorder=4)
    for p, m, s in zip(pos, means, sds):
        ax.text(p, m + s + 0.13, f'{m:.2f}', ha='center', va='bottom',
                fontsize=12, color=INK, fontweight='bold', zorder=5, family='Tahoma')

def cell(keys, vals):
    sub = long.copy()
    for k, v in zip(keys, vals):
        sub = sub[sub[k] == v]
    return sub['Shrinkage'].mean(), sub['Shrinkage'].std()

def finish(fig, ax, fname, caption, legend=True):
    if legend:
        ax.legend(frameon=False, fontsize=12.5, loc='upper right',
                  bbox_to_anchor=(1, -0.19), ncol=3, handlelength=1.1)
    ax.yaxis.set_major_locator(MultipleLocator(0.5))
    for lb in ax.get_yticklabels():
        lb.set_family('Tahoma'); lb.set_fontsize(11)
    path = os.path.join(OUT, fname)
    fig.tight_layout()
    # اگر فایل اجرای قبلی باقی مانده یا قفل باشد، ابتدا حذف می‌شود تا ذخیره‌سازی
    # با خطای «Invalid argument» متوقف نشود.
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
    try:
        fig.savefig(path, dpi=300, bbox_inches='tight')
    except OSError as e:
        plt.close(fig)
        raise SystemExit(f'[خطا] ذخیره‌ی «{path}» ممکن نشد ({e.strerror}). '
                         'اگر فایل در نرم‌افزار دیگری باز است آن را ببندید و دوباره اجرا کنید.')
    plt.close(fig)
    FA_CHARTS.append((path, caption))
    print('[نمودار]', path)

FAB_ORDER, TENS_ORDER = A.FAB_ORDER, A.TENS_ORDER
WASHES = ['Wash 1', 'Wash 2']

# --- نمودار ۱ و ۲: پنبه در برابر پلی‌استر، هر شستشو جداگانه ---
for i, w in enumerate(WASHES, start=1):
    fig, ax = plt.subplots(figsize=(6.6, 4.4))
    m = [cell(['Fabric', 'Wash'], [f, w])[0] for f in FAB_ORDER]
    s = [cell(['Fabric', 'Wash'], [f, w])[1] for f in FAB_ORDER]
    bars(ax, 2, m, s, BLUE, 0.42)
    for rect, c in zip(ax.patches, [BLUE, ORANGE]):
        rect.set_color(c)
    ax.set_xticks(range(2))
    ax.set_xticklabels([fa(FAB_FA[f]) for f in FAB_ORDER], fontsize=15)
    ax.set_ylim(0, max(np.array(m) + np.array(s)) * 1.34)
    style(ax, f'آبرفت دوخت پس از {WASH_FA[w]}: پنبه در برابر پلی‌استر',
          'میانگین ۱۸ نمونه برای هر پارچه؛ میله‌های خطا = یک انحراف معیار')
    finish(fig, ax, f'نمودار{i}_نوع-پارچه_{"شستشوی-اول" if i == 1 else "شستشوی-دوم"}.png',
           f'نمودار {i}. میانگین درصد آبرفت کل دوخت در پارچه‌های پنبه و پلی‌استر پس از {WASH_FA[w]} '
           f'(میله‌ی خطا = یک انحراف معیار، ۱۸ نمونه در هر ستون).', legend=False)

# --- نمودار ۳ و ۴: سطح کشش، هر شستشو، تفکیک‌شده بر حسب پارچه ---
for i, w in enumerate(WASHES, start=3):
    fig, ax = plt.subplots(figsize=(7.4, 4.6))
    W = 0.36
    for j, (f, col) in enumerate(zip(FAB_ORDER, [BLUE, ORANGE])):
        m = [cell(['Fabric', 'Tension', 'Wash'], [f, t, w])[0] for t in TENS_ORDER]
        s = [cell(['Fabric', 'Tension', 'Wash'], [f, t, w])[1] for t in TENS_ORDER]
        bars(ax, 3, m, s, col, W, offset=(j - 0.5) * W, label=fa(FAB_FA[f]))
    ax.set_xticks(range(3))
    ax.set_xticklabels([fa(TENS_FA[t]) for t in TENS_ORDER], fontsize=15)
    ax.set_xlabel(fa('سطح کشش دوخت'), fontsize=13)
    ax.set_ylim(0, 4.9)
    style(ax, f'اثر کشش دوخت پس از {WASH_FA[w]}',
          'میانگین ۶ نمونه در هر ستون؛ میله‌های خطا = یک انحراف معیار')
    finish(fig, ax, f'نمودار{i}_سطح-کشش_{"شستشوی-اول" if i == 3 else "شستشوی-دوم"}.png',
           f'نمودار {i}. میانگین درصد آبرفت کل دوخت در سه سطح کشش کم، متوسط و زیاد پس از {WASH_FA[w]}، '
           f'به تفکیک نوع پارچه (میله‌ی خطا = یک انحراف معیار، ۶ نمونه در هر ستون).')

# --- نمودار ۵: کشش، تجمیع دو پارچه، هر دو شستشو ---
fig, ax = plt.subplots(figsize=(7.4, 4.6))
for j, (w, col) in enumerate(zip(WASHES, [BLUE, ORANGE])):
    m = [cell(['Tension', 'Wash'], [t, w])[0] for t in TENS_ORDER]
    s = [cell(['Tension', 'Wash'], [t, w])[1] for t in TENS_ORDER]
    bars(ax, 3, m, s, col, 0.36, offset=(j - 0.5) * 0.36, label=fa(WASH_FA[w]))
ax.set_xticks(range(3))
ax.set_xticklabels([fa(TENS_FA[t]) for t in TENS_ORDER], fontsize=15)
ax.set_xlabel(fa('سطح کشش دوخت'), fontsize=13)
ax.set_ylim(0, 5.0)
style(ax, 'اثر همزمان کشش دوخت و نوبت شستشو',
      'تجمیع هر دو پارچه؛ میانگین ۱۲ نمونه در هر ستون؛ میله‌های خطا = یک انحراف معیار')
finish(fig, ax, 'نمودار5_سطح-کشش_هر-دو-شستشو.png',
       'نمودار ۵. میانگین درصد آبرفت کل دوخت بر حسب سطح کشش پس از شستشوی اول و دوم، با تجمیع هر دو پارچه '
       '(میله‌ی خطا = یک انحراف معیار، ۱۲ نمونه در هر ستون).')

# --- نمودار ۶: اثر نوبت شستشو به تفکیک پارچه ---
fig, ax = plt.subplots(figsize=(6.9, 4.6))
for j, (w, col) in enumerate(zip(WASHES, [BLUE, ORANGE])):
    m = [cell(['Fabric', 'Wash'], [f, w])[0] for f in FAB_ORDER]
    s = [cell(['Fabric', 'Wash'], [f, w])[1] for f in FAB_ORDER]
    bars(ax, 2, m, s, col, 0.36, offset=(j - 0.5) * 0.36, label=fa(WASH_FA[w]))
ax.set_xticks(range(2))
ax.set_xticklabels([fa(FAB_FA[f]) for f in FAB_ORDER], fontsize=15)
ax.set_ylim(0, 4.7)
style(ax, 'رشد آبرفت از شستشوی اول به شستشوی دوم',
      'میانگین ۱۸ نمونه در هر ستون؛ میله‌های خطا = یک انحراف معیار')
finish(fig, ax, 'نمودار6_اثر-نوبت-شستشو.png',
       'نمودار ۶. میانگین درصد آبرفت کل دوخت هر پارچه پس از شستشوی اول و دوم '
       '(میله‌ی خطا = یک انحراف معیار، ۱۸ نمونه در هر ستون).')

# --- نمودار ۷: اثر نمره نخ دوخت ---
fig, ax = plt.subplots(figsize=(6.9, 4.6))
for j, (w, col) in enumerate(zip(WASHES, [BLUE, ORANGE])):
    m = [cell(['YarnCount', 'Wash'], [c, w])[0] for c in [20, 40]]
    s = [cell(['YarnCount', 'Wash'], [c, w])[1] for c in [20, 40]]
    bars(ax, 2, m, s, col, 0.36, offset=(j - 0.5) * 0.36, label=fa(WASH_FA[w]))
ax.set_xticks([0, 1])
ax.set_xticklabels(['Ne 20', 'Ne 40'], fontsize=14, family='Tahoma')
ax.set_xlabel(fa('نمره نخ دوخت'), fontsize=13)
ax.set_ylim(0, 4.7)
style(ax, 'اثر نمره نخ دوخت بر آبرفت',
      'تجمیع هر دو پارچه؛ میانگین ۱۸ نمونه در هر ستون؛ میله‌های خطا = یک انحراف معیار')
finish(fig, ax, 'نمودار7_نمره-نخ.png',
       'نمودار ۷. میانگین درصد آبرفت کل دوخت برای نخ دوخت Ne 20 و Ne 40 پس از هر شستشو، '
       'با تجمیع هر دو پارچه (میله‌ی خطا = یک انحراف معیار، ۱۸ نمونه در هر ستون).')

# ============================ جدول اکسل فارسی ============================
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

XLS_FA = os.path.join(OUT, 'آمار_توصیفی_و_نتایج_آزمون‌ها.xlsx')
HDR_FILL = PatternFill('solid', fgColor='1C5CAB')
HDR_FONT = Font(name='B Nazanin', bold=True, color='FFFFFF', size=12)
BODY_FONT = Font(name='B Nazanin', size=12)
TTL_FONT = Font(name='B Titr', bold=True, size=13, color='0D366B')
BAND = PatternFill('solid', fgColor='EEF4FC')
THIN = Side(style='thin', color='B7D3F6')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

SHEET_FA = {
    'T1_Overall': ('آبرفت به تفکیک پارچه', 'آبرفت دوخت به تفکیک نوع پارچه (تجمیع هر دو شستشو)'),
    'T2_Fabric_x_Wash': ('پارچه و شستشو', 'آبرفت دوخت به تفکیک نوع پارچه و نوبت شستشو'),
    'T3_Tension_x_Wash': ('کشش و شستشو', 'آبرفت دوخت به تفکیک سطح کشش دوخت و نوبت شستشو'),
    'T4_Fabric_Tension_Wash': ('پارچه، کشش و شستشو', 'آبرفت دوخت به تفکیک نوع پارچه، سطح کشش و نوبت شستشو'),
    'T5_YarnCount_x_Wash': ('نمره نخ و شستشو', 'آبرفت دوخت به تفکیک نمره نخ دوخت و نوبت شستشو'),
    'T7_Wash': ('نوبت شستشو', 'آبرفت دوخت به تفکیک نوبت شستشو (تجمیع همه‌ی نمونه‌ها)'),
    'T6_Full_Design': ('طرح کامل عاملی', 'آمار همه‌ی خانه‌های طرح عاملی (پارچه × نمره نخ × کشش × شستشو)'),
}

anova_fa = anova_df[['Test', 'Factor', 'Scope', 'df', 'F', 'p_str', 'eta2', 'Verdict']].copy()
FACTOR_FA = {'Fabric type': 'نوع پارچه', 'Seam tension level': 'سطح کشش دوخت',
             'Wash number': 'نوبت شستشو', 'Sewing-thread yarn count': 'نمره نخ دوخت'}
SCOPE_FA = {
    'Both washes pooled (N = 72)': 'تجمیع هر دو شستشو (۷۲ داده)',
    'Wash 1 only (N = 36)': 'فقط شستشوی اول (۳۶ داده)',
    'Wash 2 only (N = 36)': 'فقط شستشوی دوم (۳۶ داده)',
    'Both fabrics & washes pooled (N = 72)': 'تجمیع هر دو پارچه و هر دو شستشو (۷۲ داده)',
    'Cotton only (N = 36)': 'فقط پنبه (۳۶ داده)',
    'Polyester only (N = 36)': 'فقط پلی‌استر (۳۶ داده)',
    'All samples pooled (N = 72)': 'تجمیع همه‌ی نمونه‌ها (۷۲ داده)',
    'Both fabrics pooled': 'تجمیع هر دو پارچه', 'Wash 1 only': 'فقط شستشوی اول',
    'Wash 2 only': 'فقط شستشوی دوم', 'Both washes pooled': 'تجمیع هر دو شستشو',
    'All samples pooled': 'تجمیع همه‌ی نمونه‌ها', 'Cotton only': 'فقط پنبه', 'Polyester only': 'فقط پلی‌استر',
}
VERDICT_FA = {'Significant difference exists': 'تفاوت معنادار وجود دارد',
              'No significant difference': 'تفاوت معنادار وجود ندارد'}
anova_fa['Factor'] = anova_fa['Factor'].map(FACTOR_FA)
anova_fa['Scope'] = anova_fa['Scope'].map(SCOPE_FA)
anova_fa['Verdict'] = anova_fa['Verdict'].map(VERDICT_FA)
anova_fa['p_str'] = anova_fa['p_str'].replace('< 0.001', '‏< 0.001')
anova_fa.columns = ['شناسه', 'عامل بررسی‌شده', 'دامنه‌ی داده', 'درجات آزادی', 'آماره F',
                    'مقدار p', 'اندازه اثر (اتا مربع)', 'تفسیر']

ph_fa = posthoc_df.copy()
ph_fa['Scope'] = ph_fa['Scope'].map({'Both fabrics pooled': 'تجمیع هر دو پارچه',
                                     'Cotton only': 'فقط پنبه', 'Polyester only': 'فقط پلی‌استر'})
ph_fa['Comparison'] = ph_fa['Comparison'].replace({
    'Low vs Medium': 'کم در برابر متوسط', 'Low vs High': 'کم در برابر زیاد',
    'Medium vs High': 'متوسط در برابر زیاد'})
ph_fa['Result'] = ph_fa['Result'].map({'Significant': 'معنادار', 'Not significant': 'غیرمعنادار'})
ph_fa.columns = ['دامنه‌ی داده', 'مقایسه', 'اختلاف میانگین', 'آماره t', 'مقدار p خام',
                 'مقدار p تصحیح‌شده', 'اندازه اثر (d کوهن)', 'نتیجه']

np_fa = NP_RESULTS.copy()
np_fa['Factor'] = np_fa['Factor'].map(FACTOR_FA)
np_fa['Scope'] = np_fa['Scope'].map(SCOPE_FA)
np_fa['Verdict'] = np_fa['Verdict'].map(VERDICT_FA)
np_fa['p_str'] = np_fa['p_str'].replace('< 0.001', '‏< 0.001')
np_fa.columns = ['عامل بررسی‌شده', 'دامنه‌ی داده', 'آماره H', 'مقدار p', 'تفسیر']

long_fa = fa_frame(long)

def write_sheet(xw, df, sheet, title):
    df.to_excel(xw, sheet_name=sheet[:31], index=False, startrow=2)
    ws = xw.sheets[sheet[:31]]
    ws.sheet_view.rightToLeft = True
    ws['A1'] = title
    ws['A1'].font = TTL_FONT
    for c in range(1, df.shape[1] + 1):
        cl = ws.cell(row=3, column=c)
        cl.fill, cl.font = HDR_FILL, HDR_FONT
        cl.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for r in range(4, 4 + len(df)):
        for c in range(1, df.shape[1] + 1):
            cl = ws.cell(row=r, column=c)
            cl.border, cl.font = BORDER, BODY_FONT
            cl.alignment = Alignment(horizontal='center', vertical='center')
            if r % 2 == 0:
                cl.fill = BAND
            if isinstance(cl.value, float):
                cl.number_format = '0.000'
    for c in range(1, df.shape[1] + 1):
        w = max([len(str(df.columns[c - 1]))] + [len(str(v)) for v in df.iloc[:, c - 1]]) + 5
        ws.column_dimensions[get_column_letter(c)].width = min(w, 42)
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = 'A4'

try:
    _xw_cm = pd.ExcelWriter(XLS_FA, engine='openpyxl')
except OSError as e:
    raise SystemExit(f'[خطا] نوشتن «{XLS_FA}» ممکن نشد ({e.strerror}). '
                     'اگر فایل در Excel باز است آن را ببندید و دوباره اجرا کنید.')
with _xw_cm as xw:
    for key, (sheet, title) in SHEET_FA.items():
        write_sheet(xw, fa_frame(TABLES[key]), sheet, title)
    write_sheet(xw, anova_fa, 'نتایج آنالیز واریانس', 'نتایج تحلیل واریانس یک‌طرفه برای همه‌ی عامل‌ها')
    write_sheet(xw, ph_fa, 'آزمون تعقیبی کشش', 'مقایسه‌های دوبه‌دوی سطوح کشش با تصحیح بونفرونی')
    write_sheet(xw, np_fa, 'آزمون کروسکال والیس', 'تأیید ناپارامتری نتایج با آزمون کروسکال-والیس')
    write_sheet(xw, long_fa, 'داده‌های پاک‌سازی‌شده', 'داده‌های نهایی تحلیل؛ هر سطر یک مشاهده (۷۲ مشاهده)')

print('[اکسل]', XLS_FA)

# ============================ ابزارهای ساخت سند فارسی (راست‌به‌چپ) ============================
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY, BLUEC, GREY = RGBColor(0x0D, 0x36, 0x6B), RGBColor(0x1C, 0x5C, 0xAB), RGBColor(0x5C, 0x5C, 0x58)
FA_FONT, FA_TITLE_FONT = 'B Nazanin', 'B Titr'


def _el(tag, val='1'):
    e = OxmlElement(tag)
    e.set(qn('w:val'), val)
    return e


def rtl_run(run, font=FA_FONT, size=13):
    """یک اجرای متنی را راست‌به‌چپ و با فونت فارسی تنظیم می‌کند."""
    run.font.name = font
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rPr.rFonts.set(qn('w:cs'), font)
    rPr.rFonts.set(qn('w:ascii'), font)
    rPr.rFonts.set(qn('w:hAnsi'), font)
    rPr.append(_el('w:rtl'))
    szcs = OxmlElement('w:szCs')
    szcs.set(qn('w:val'), str(int(size * 2)))
    rPr.append(szcs)
    return run


def rtl_par(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    pPr = p._p.get_or_add_pPr()
    pPr.append(_el('w:bidi'))
    p.alignment = align
    return p


def rtl_table(t):
    tblPr = t._tbl.tblPr
    tblPr.append(_el('w:bidiVisual'))
    return t


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def new_doc():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = FA_FONT
    st.font.size = Pt(13)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.4
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2.2)
        s.left_margin = s.right_margin = Cm(2.3)
        s._sectPr.append(_el('w:bidi'))
    return doc


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
    return P(doc, text, size=size, bold=True, color=NAVY if level == 1 else BLUEC,
             align=WD_ALIGN_PARAGRAPH.RIGHT, font=FA_TITLE_FONT if level == 1 else FA_FONT,
             before=16 if level == 1 else 12, after=6)


def BUL(doc, text, size=13):
    p = doc.add_paragraph(style='List Bullet')
    r = p.add_run(text)
    rtl_run(r, FA_FONT, size)
    rtl_par(p, WD_ALIGN_PARAGRAPH.JUSTIFY)
    p.paragraph_format.space_after = Pt(3)
    return p


def TBL(doc, df, caption=None, note=None, size=11):
    if caption:
        P(doc, caption, size=11.5, bold=True, color=GREY,
          align=WD_ALIGN_PARAGRAPH.RIGHT, before=10, after=4)
    t = doc.add_table(rows=1, cols=df.shape[1])
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    rtl_table(t)
    for j, col in enumerate(df.columns):
        c = t.rows[0].cells[j]
        c.text = ''
        r = c.paragraphs[0].add_run(str(col))
        rtl_run(r, FA_FONT, size)
        r.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        rtl_par(c.paragraphs[0], WD_ALIGN_PARAGRAPH.CENTER)
        shade(c, '1C5CAB')
    for i, (_, row) in enumerate(df.iterrows()):
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ''
            txt = f'{v:.3f}' if isinstance(v, float) else str(v)
            r = cells[j].paragraphs[0].add_run(txt)
            rtl_run(r, FA_FONT, size)
            rtl_par(cells[j].paragraphs[0], WD_ALIGN_PARAGRAPH.CENTER)
            if i % 2 == 1:
                shade(cells[j], 'EEF4FC')
    if note:
        P(doc, note, size=10.5, italic=True, color=GREY, before=3, after=10)
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
    return t


def FIG(doc, path, caption, width_cm=14.2):
    doc.add_picture(path, width=Cm(width_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    P(doc, caption, size=11, italic=True, color=GREY,
      align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=14)


def PAGEBREAK(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


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
    print('[سند]', path)


# اعداد کلیدی که در متن‌ها بارها استفاده می‌شوند
def val(test, field):
    return anova_df.loc[anova_df.Test == test, field].iloc[0]

def pstr(test):
    s = val(test, 'p_str')
    return f'p {s}' if s.startswith('<') else f'p = {s}'

M_COTTON = TABLES['T1_Overall'].loc[0, 'Mean']
M_POLY   = TABLES['T1_Overall'].loc[1, 'Mean']
M_W1     = TABLES['T7_Wash'].loc[0, 'Mean']
M_W2     = TABLES['T7_Wash'].loc[1, 'Mean']
TODAY    = datetime.date.today().strftime('%Y/%m/%d')

# ============================ سند ۱: گزارش تحلیل آماری ============================
doc = new_doc()
for _ in range(4):
    doc.add_paragraph()
P(doc, 'گزارش تحلیل آماری', size=12, bold=True, color=BLUEC, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
P(doc, 'اثر پارامترهای نخ و پارچه بر آبرفتگی دوخت', size=24, bold=True, color=NAVY,
  align=WD_ALIGN_PARAGRAPH.CENTER, font=FA_TITLE_FONT, after=8)
P(doc, 'مقایسه‌ی پارچه‌های پنبه و پلی‌استر پس از شستشوی اول و دوم', size=15, italic=True,
  color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=36)

info = pd.DataFrame({
    'شرح': ['فایل داده‌ی اولیه', 'طرح آزمایش', 'حجم نمونه', 'متغیر پاسخ',
            'روش‌های آماری', 'سطح معناداری', 'نرم‌افزار', 'تاریخ گزارش'],
    'مقدار': [
        A.SRC,
        'عاملی کامل: ۲ پارچه × ۲ نمره نخ × ۳ سطح کشش × ۳ تکرار',
        f'{len(wide)} نمونه‌ی دوخته‌شده ({len(long)} مشاهده‌ی آبرفت)',
        'درصد آبرفت کل دوخت',
        'آمار توصیفی، تحلیل واریانس یک‌طرفه، آزمون تعقیبی بونفرونی، کروسکال-والیس',
        'آلفا = ۰٫۰۵',
        'پایتون همراه با کتابخانه‌های pandas، SciPy، matplotlib و python-docx',
        TODAY]})
TBL(doc, info, size=11.5)
PAGEBREAK(doc)

H(doc, '۱. مقدمه و هدف پژوهش', 1)
P(doc, 'آبرفتگی دوخت، جمع‌شدگی ابعادی است که پس از شستشو در امتداد خط دوخت پدید می‌آید. این پدیده '
       'حتی وقتی قطعات پارچه به‌تنهایی پایداری ابعادی دارند نیز ظاهر لباس را خراب می‌کند، زیرا نخ دوخت، '
       'پارچه و هندسه‌ی بخیه با نرخ‌های متفاوتی رها می‌شوند. این گزارش نشان می‌دهد که جنس الیاف پارچه، '
       'نمره‌ی نخ دوخت و کشش اعمال‌شده هنگام دوخت چه اثری بر میزان آبرفتگی دوخت دارند و این آبرفتگی '
       'بین شستشوی اول و دوم چگونه رشد می‌کند.')
P(doc, 'سه پرسش اصلی به‌صورت آماری بررسی شده است:')
for q in ['آیا نوع پارچه (پنبه در برابر پلی‌استر) میزان آبرفتگی دوخت را تغییر می‌دهد؟',
          'آیا سطح کشش دوخت (کم، متوسط، زیاد) میزان آبرفتگی دوخت را تغییر می‌دهد؟',
          'آیا نوبت شستشو (اول در برابر دوم) میزان آبرفتگی دوخت را تغییر می‌دهد؟']:
    BUL(doc, q)
P(doc, 'افزون بر این، آزمون تکمیلی روی نمره‌ی نخ دوخت (Ne 20 در برابر Ne 40) نیز گزارش شده است، '
       'زیرا نمره‌ی نخ یکی از پارامترهای نخ است که در عنوان پژوهش آمده است.')

H(doc, '۲. مواد، طرح آزمایش و آماده‌سازی داده‌ها', 1)
P(doc, 'فایل اکسل اولیه شامل شش کاربرگ است. دو کاربرگ «نمونه پنبه» و «نمونه پلی استر» اندازه‌گیری‌های '
       'نمونه‌های دوخته‌شده را در بر دارند و مبنای این تحلیل هستند؛ چهار کاربرگ دیگر مشخصات دوک‌های نخ '
       '(نمره‌های ۲۰٫۲ و ۴۰٫۲) و ساختار دو پارچه‌ی پایه را توصیف می‌کنند و تنها برای شناسایی مواد به کار '
       'رفته‌اند. سطرهای مربوط به نمونه‌های شاهدِ بدون دوخت («Control 1 تا 3») و سطرهای دوک خام پلی‌استر '
       'از تحلیل آماری کنار گذاشته شدند، زیرا این سطرها به‌جای درصد آبرفت اندازه‌گیری‌شده، نمره‌ی ارزیابان '
       'را در خود دارند.')
P(doc, 'مجموعه‌ی داده‌ی حاصل یک طرح عاملی کامل و کاملاً متوازن است:')
design = pd.DataFrame({
    'عامل': ['نوع پارچه', 'نمره نخ دوخت', 'سطح کشش دوخت', 'تکرار', 'نوبت شستشو (اندازه‌گیری تکراری)'],
    'تعداد سطوح': ['۲', '۲', '۳', '۳', '۲'],
    'سطوح': ['پنبه، پلی‌استر', 'Ne 20، Ne 40', 'کم، متوسط، زیاد', '۱، ۲، ۳', 'شستشوی اول، شستشوی دوم']})
TBL(doc, design, 'جدول ۱. طرح آزمایش پژوهش آبرفتگی دوخت.',
    'هر یک از ۱۲ خانه‌ی طرح (۲×۲×۳) شامل ۳ نمونه‌ی تکراری است و در مجموع ۳۶ نمونه‌ی دوخته‌شده به دست می‌آید. '
    'هر نمونه پس از شستشوی اول و پس از شستشوی دوم اندازه‌گیری شده و ۷۲ مشاهده حاصل شده است. هیچ داده‌ی گمشده‌ای وجود ندارد.')
P(doc, 'متغیر پاسخ، درصد آبرفت کل دوخت است که در فایل اولیه از روی طول نشانه‌گذاری‌شده‌ی ۲۰۰ میلی‌متری '
       'دوخت، پیش و پس از شستشو محاسبه شده است. مقادیر روی شبکه‌ی اندازه‌گیری ۰٫۵ درصد ثبت شده‌اند که '
       'دقت عملی روش اندازه‌گیری است.')

PAGEBREAK(doc)
H(doc, '۳. آمار توصیفی', 1)
P(doc, 'جدول‌های ۲ تا ۸ میانگین، انحراف معیار، کمینه، بیشینه و تعداد نمونه‌ی آبرفت دوخت را برای همه‌ی '
       'گروه‌بندی‌های مورد نیاز پژوهش خلاصه می‌کنند. همه‌ی مقادیر بر حسب درصد هستند.')
CAPS = {
    'T1_Overall': ('جدول ۲. آبرفت دوخت به تفکیک نوع پارچه (تجمیع هر دو شستشو).'),
    'T2_Fabric_x_Wash': ('جدول ۳. آبرفت دوخت به تفکیک نوع پارچه و نوبت شستشو.'),
    'T3_Tension_x_Wash': ('جدول ۴. آبرفت دوخت به تفکیک سطح کشش دوخت و نوبت شستشو.'),
    'T4_Fabric_Tension_Wash': ('جدول ۵. آبرفت دوخت به تفکیک نوع پارچه، سطح کشش و نوبت شستشو.'),
    'T5_YarnCount_x_Wash': ('جدول ۶. آبرفت دوخت به تفکیک نمره نخ دوخت و نوبت شستشو.'),
    'T7_Wash': ('جدول ۷. آبرفت دوخت به تفکیک نوبت شستشو (تجمیع همه‌ی نمونه‌ها).'),
    'T6_Full_Design': ('جدول ۸. آمار همه‌ی خانه‌های طرح عاملی (پارچه × نمره نخ × کشش × شستشو).'),
}
for k, cap in CAPS.items():
    TBL(doc, fa_frame(TABLES[k]), cap,
        'انحراف معیار نمونه‌ای گزارش شده است؛ «تعداد» یعنی شمار مشاهده‌های آن گروه.')
P(doc, f'تصویر توصیفی از همین‌جا روشن است: آبرفت دوخت در پنبه حدود دو و نیم برابر پلی‌استر است '
       f'({M_COTTON:.2f} درصد در برابر {M_POLY:.2f} درصد)، همه‌ی گروه‌ها در شستشوی دوم بیشتر جمع می‌شوند '
       'و آبرفت با افزایش کشش دوخت از کم به زیاد به‌طور یکنواخت بالا می‌رود.')

PAGEBREAK(doc)
H(doc, '۴. مقایسه‌ی نموداری', 1)
P(doc, 'در همه‌ی نمودارها میانگین هر گروه به‌همراه میله‌ی خطا به اندازه‌ی یک انحراف معیار رسم شده و '
       'مقدار عددی میانگین بالای هر ستون چاپ شده است.')
for path, cap in FA_CHARTS:
    FIG(doc, path, cap)

PAGEBREAK(doc)
H(doc, '۵. تحلیل واریانس', 1)
P(doc, 'برای هر یک از سه عامل مطرح‌شده در هدف پژوهش، تحلیل واریانس یک‌طرفه هم روی کل داده‌ها و هم درون '
       'زیرگروه‌های مرتبط اجرا شد. قاعده‌ی تصمیم همان قاعده‌ی مرسوم است: اگر p کوچک‌تر از ۰٫۰۵ باشد، '
       'تفاوت معناداری میان میانگین گروه‌ها وجود دارد؛ در غیر این صورت داده‌ها شواهدی بر وجود تفاوت به دست نمی‌دهند.')
TBL(doc, anova_fa, 'جدول ۹. نتایج تحلیل واریانس یک‌طرفه برای همه‌ی عامل‌های بررسی‌شده.',
    'درجات آزادی به‌صورت (بین‌گروهی، درون‌گروهی) گزارش شده است. اتا مربع سهم عامل از کل تغییرپذیری آبرفت را '
    'نشان می‌دهد: ۰٫۰۱ کوچک، ۰٫۰۶ متوسط و ۰٫۱۴ بزرگ. ستون تفسیر، قاعده‌ی p < 0.05 را اعمال کرده است.', size=10)

H(doc, '۵-۱. اثر نوع پارچه', 2)
P(doc, f'نوع پارچه به‌مراتب قوی‌ترین عامل است. با تجمیع هر دو شستشو، F({val("A1","df")}) = {val("A1","F")} و '
       f'{pstr("A1")} با اتا مربع برابر {val("A1","eta2")} به دست آمد؛ یعنی تفاوت معنادار وجود دارد و حدود '
       f'{val("A1","eta2")*100:.0f} درصد از کل تغییرپذیری آبرفت دوخت تنها به جنس الیاف مربوط است. این اثر هم '
       f'پس از شستشوی اول (F({val("A2","df")}) = {val("A2","F")}، {pstr("A2")}) و هم پس از شستشوی دوم '
       f'(F({val("A3","df")}) = {val("A3","F")}، {pstr("A3")}) به‌تنهایی معنادار است. پنبه در هر دو حالت بیشتر '
       'جمع می‌شود؛ این با تورم رطوبتی الیاف سلولزی و رهایش تنش‌های ذخیره‌شده در آن‌ها هنگام شستشو سازگار '
       'است، حال آنکه نخ پلی‌استر گرمانرم ابعاد تثبیت‌شده در فرایند تثبیت حرارتی را تا حد زیادی حفظ می‌کند.')

H(doc, '۵-۲. اثر سطح کشش دوخت', 2)
P(doc, f'با تجمیع هر دو پارچه و هر دو شستشو، کشش دوخت اثر معناداری دارد (F({val("B1","df")}) = {val("B1","F")}، '
       f'{pstr("B1")}، اتا مربع = {val("B1","eta2")})، اما این اثر بسیار ضعیف‌تر از اثر جنس الیاف است و میان دو '
       f'پارچه یکسان نیست. تفکیک بر حسب پارچه علت را روشن می‌کند: در پلی‌استر اثر کشش قوی و معنادار است '
       f'(F({val("B5","df")}) = {val("B5","F")}، {pstr("B5")}، اتا مربع = {val("B5","eta2")})، در حالی که در پنبه '
       f'اصلاً دیده نمی‌شود (F({val("B4","df")}) = {val("B4","F")}، {pstr("B4")}). در پنبه آبرفت ناشی از خود الیاف '
       'چنان بزرگ است که سهم کشش بخیه را می‌پوشاند؛ اما در پلی‌استرِ ابعادپایدار، کرنش کشسانِ ذخیره‌شده در '
       'دوختِ بیش از حد کشیده به سازوکار اصلی تبدیل می‌شود که هنگام شستشو آزاد می‌گردد.')
TBL(doc, ph_fa, 'جدول ۱۰. مقایسه‌های دوبه‌دوی سطوح کشش با تصحیح بونفرونی.',
    'اختلاف میانگین منفی یعنی سطح نخست کمتر از سطح دوم جمع شده است. تفسیر d کوهن: ۰٫۲ کوچک، ۰٫۵ متوسط و ۰٫۸ بزرگ.', size=10)
P(doc, 'آزمون‌های تعقیبی نشان می‌دهند که اثر کشش دقیقاً در سطح «زیاد» متمرکز است: در داده‌های تجمیعی تنها '
       'مقایسه‌ی «کم در برابر زیاد» معنادار است و در پلی‌استر هر دو مقایسه‌ی «کم در برابر زیاد» و «متوسط در '
       'برابر زیاد» با اندازه اثر بسیار بزرگ معنادار می‌شوند، در حالی که «کم در برابر متوسط» هرگز تفاوت '
       'معناداری ندارد. از نظر عملی یعنی تغییر ملایم کشش دوخت بی‌ضرر است و آنچه آسیب می‌زند، تنظیم کشش در '
       'حد زیاد است.')

H(doc, '۵-۳. اثر نوبت شستشو', 2)
P(doc, f'آبرفت دوخت از شستشوی اول به دوم به‌طور معنادار افزایش می‌یابد (F({val("C1","df")}) = {val("C1","F")}، '
       f'{pstr("C1")}، اتا مربع = {val("C1","eta2")}) و از {M_W1:.2f} درصد به {M_W2:.2f} درصد می‌رسد. این افزایش '
       f'درون هر پارچه نیز جداگانه معنادار است (پنبه: F({val("C2","df")}) = {val("C2","F")}، {pstr("C2")}؛ '
       f'پلی‌استر: F({val("C3","df")}) = {val("C3","F")}، {pstr("C3")}).')
P(doc, 'از آنجا که دو شستشو اندازه‌گیری تکراری روی نمونه‌های یکسان هستند، تحلیل واریانس بالا دو نمونه‌ی در '
       'واقع جفت‌شده را مستقل فرض می‌کند که فرضی محافظه‌کارانه است. آزمون t زوجی روی همین داده‌ها نتیجه را '
       f'تأیید و تقویت می‌کند: میانگین افزایش درون‌نمونه‌ای برابر {paired["mean_increase"]} واحد درصد با '
       f'انحراف معیار {paired["sd_increase"]} است و t({paired["df"]}) = {paired["t"]} با '
       f'p {paired["p"]} به دست می‌آید. هر ۳۶ نمونه بدون استثنا در شستشوی دوم بیشتر جمع شده‌اند؛ بنابراین یک '
       'چرخه‌ی شستشو برای توصیف رفتار دوخت هیچ‌یک از دو پارچه کافی نیست.')

H(doc, '۵-۴. اثر نمره نخ دوخت', 2)
P(doc, f'نمره‌ی نخ دوخت (Ne 20 در برابر Ne 40) اثر معناداری بر آبرفت دوخت نداشت '
       f'(F({val("D1","df")}) = {val("D1","F")}، {pstr("D1")}، اتا مربع = {val("D1","eta2")}). نخ ظریف‌تر Ne 40 در '
       'هر دو شستشو میانگین اندکی کمتر داد، اما این تفاوت کاملاً درون پراکندگی تکرارها قرار می‌گیرد.')

H(doc, '۵-۵. بررسی مفروضات و تأیید ناپارامتری', 2)
P(doc, 'آزمون لِوین روی همگنی واریانس‌ها برای همه‌ی مدل‌های گزارش‌شده غیرمعنادار بود، جز مقایسه‌ی کشش درون '
       'پلی‌استر که واریانس نزدیک به صفرِ خانه‌های کشش کم و متوسط، آماره را متورم می‌کند. آزمون شاپیرو-ویلک '
       'در بیشتر گروه‌ها نرمال بودن را رد می‌کند؛ این نتیجه بیش از آنکه نگران‌کننده باشد، انتظار می‌رفت: آبرفت '
       'روی شبکه‌ی گسسته‌ی ۰٫۵ درصد ثبت شده، پس مقادیر گره‌های تکراری فراوان دارند و نمی‌توانند دقیقاً نرمال '
       'باشند. چون تحلیل واریانس در طرح متوازن نسبت به این موضوع مقاوم است و برای آنکه نتیجه‌گیری‌ها مستقل '
       'از فرض نرمال بودن باشند، همه‌ی آزمون‌ها با آزمون رتبه‌ای کروسکال-والیس تکرار شدند.')
TBL(doc, np_fa, 'جدول ۱۱. تأیید ناپارامتری نتایج با آزمون کروسکال-والیس.',
    'همه‌ی نتیجه‌گیری‌های حاصل از تحلیل واریانس با آزمون رتبه‌ای بازتولید می‌شوند؛ بنابراین هیچ‌یک از یافته‌ها به '
    'فرض نرمال بودن وابسته نیست.', size=10.5)

PAGEBREAK(doc)
H(doc, '۶. نتیجه‌گیری', 1)
P(doc, f'این تحلیل روی ۳۶ نمونه‌ی دوخته‌شده که هر کدام پس از دو چرخه‌ی شستشو اندازه‌گیری شده‌اند، پاسخی روشن و '
       f'سازگار به سه پرسش پژوهش می‌دهد. جنس الیاف تعیین‌کننده‌ی اصلی آبرفت دوخت است: دوخت روی پنبه به‌طور '
       f'میانگین {M_COTTON:.2f} درصد در برابر {M_POLY:.2f} درصد پلی‌استر جمع شد؛ تفاوتی بسیار معنادار '
       f'(p < 0.001) که به‌تنهایی حدود {val("A1","eta2")*100:.0f} درصد از کل تغییرپذیری را توضیح می‌دهد و هم پس '
       'از شستشوی اول و هم پس از دوم برقرار است. کشش دوخت عامل دوم و ضعیف‌تر اما از نظر صنعتی کنترل‌پذیر '
       'است: آبرفت از کشش کم به متوسط و زیاد به‌طور یکنواخت افزایش یافت و اثر تجمیعی معنادار شد، ولی تحلیل '
       'تعقیبی نشان می‌دهد تنها تنظیم کشش زیاد به‌طور قابل‌اتکا با بقیه تفاوت دارد و این اثر تقریباً به‌طور کامل '
       'در پارچه‌ی پلی‌استر رخ می‌دهد، جایی که کرنش ذخیره‌شده در دوخت بیش از حد کشیده هنگام شستشو آزاد '
       f'می‌شود و زیر آبرفت الیاف پنهان نمی‌ماند. نوبت شستشو برای هر دو پارچه معنادار است: آبرفت از '
       f'{M_W1:.2f} درصد پس از شستشوی اول به {M_W2:.2f} درصد پس از شستشوی دوم رسید؛ افزایشی درون‌نمونه‌ای '
       'در حدود یک واحد درصد که در تک‌تک نمونه‌ها مشاهده شد. در مقابل، نمره‌ی نخ دوخت در بازه‌ی Ne 20 تا '
       'Ne 40 اثر معناداری نداشت. از نظر کاربردی، تولیدکننده‌ای که دوخت پایدار می‌خواهد باید جنس الیاف را '
       'اهرم اصلی بداند، کشش دوخت را از تنظیم «زیاد» دور نگه دارد — که به‌ویژه روی پارچه‌های مصنوعی اهمیت '
       'دارد، چون در آن‌ها کشش مهم‌ترین منبع باقی‌مانده‌ی آبرفت دوخت است — و کیفیت دوخت را دست‌کم طی دو '
       'چرخه‌ی شستشو ارزیابی کند، زیرا یک شستشو به‌طور نظام‌مند آبرفت نهایی را کمتر از واقع نشان می‌دهد. این '
       'نتیجه‌گیری‌ها هم با آزمون‌های پارامتری و هم ناپارامتری پشتیبانی می‌شوند. دامنه‌ی اعتبار آن‌ها به اندازه‌ی '
       'آزمایش محدود است: سه تکرار در هر خانه، یک نوع بخیه و دو نمره نخ، با پاسخی که روی شبکه‌ی ۰٫۵ درصد '
       'ثبت شده است؛ برای کمّی‌کردن اثر متقابل پارچه و کشش که تحلیل‌های زیرگروهی به‌شدت به آن اشاره می‌کنند، '
       'آزمایشی بزرگ‌تر همراه با تحلیل واریانس عاملی کامل لازم است.')

save(doc, os.path.join(OUT, 'گزارش_تحلیل_آماری.docx'))

# ============================ سند ۲: راهنمای کامل پروژه برای دانشجو ============================
g = new_doc()
for _ in range(4):
    g.add_paragraph()
P(g, 'راهنمای مطالعه و ارائه', size=12, bold=True, color=BLUEC, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
P(g, 'توضیح کامل پروژه به زبان ساده', size=24, bold=True, color=NAVY,
  align=WD_ALIGN_PARAGRAPH.CENTER, font=FA_TITLE_FONT, after=8)
P(g, 'اثر پارامترهای نخ و پارچه بر آبرفتگی دوخت در پارچه‌های پنبه و پلی‌استر', size=15,
  italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
P(g, 'این سند برای آن نوشته شده است که پیش از ارائه، دقیقاً بدانید هر عدد، هر جدول و هر نمودار '
     'این پروژه چه می‌گوید، چگونه به دست آمده و اگر داور درباره‌ی آن پرسید چه باید پاسخ داد.',
  size=13, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=30)
P(g, f'تاریخ تهیه: {TODAY}', size=11.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
PAGEBREAK(g)

# ---------------------------------------------------------------- فهرست
H(g, 'فهرست مطالب', 1)
for i, t in enumerate([
        'این پروژه درباره‌ی چیست؟', 'فایل اکسل اولیه چه چیزی در خود داشت؟',
        'طرح آزمایش: چرا دقیقاً ۳۶ نمونه؟', 'متغیر پاسخ: «درصد آبرفت» چگونه محاسبه می‌شود؟',
        'چه داده‌هایی کنار گذاشته شد و چرا؟', 'مفاهیم آماری به زبان ساده',
        'نمودارها یکی‌یکی: هر کدام چه می‌گویند؟', 'جدول‌ها چه چیزی نشان می‌دهند؟',
        'نتایج اصلی و تفسیر فنی آن‌ها', 'محدودیت‌های پژوهش',
        'راهنمای ارائه: چه بگویید و به چه سؤال‌هایی آماده باشید',
        'فایل‌های پروژه و نحوه‌ی اجرای دوباره'], start=1):
    BUL(g, f'بخش {i}: {t}')

# ---------------------------------------------------------------- ۱
PAGEBREAK(g)
H(g, 'بخش ۱: این پروژه درباره‌ی چیست؟', 1)
P(g, 'وقتی یک لباس دوخته‌شده را می‌شویید، ممکن است خود پارچه تقریباً ثابت بماند اما خط دوخت جمع شود و '
     'حالت چین‌خورده یا موج‌دار پیدا کند. به این جمع‌شدگی در امتداد درز، «آبرفتگی دوخت» گفته می‌شود. '
     'علت آن ساده است: در محل درز سه چیز کنار هم قرار دارند — پارچه، نخ دوخت و هندسه‌ی بخیه — و این سه '
     'هنگام شستشو با سرعت و میزان متفاوتی جمع می‌شوند یا تنش خود را آزاد می‌کنند. اختلاف همین رفتارها '
     'است که درز را جمع می‌کند.')
P(g, 'در این پروژه می‌خواهیم بدانیم کدام عامل‌ها این پدیده را بیشتر یا کمتر می‌کنند. سه پرسش اصلی داریم:')
BUL(g, 'آیا جنس پارچه مهم است؟ یعنی آیا درز روی پنبه بیشتر از پلی‌استر جمع می‌شود؟')
BUL(g, 'آیا کشش نخ هنگام دوخت مهم است؟ یعنی اگر چرخ خیاطی نخ را محکم‌تر بکشد، درز بیشتر جمع می‌شود؟')
BUL(g, 'آیا تعداد شستشو مهم است؟ یعنی آیا شستشوی دوم چیزی به آبرفت اضافه می‌کند یا همه‌چیز در شستشوی اول رخ می‌دهد؟')
P(g, 'یک پرسش تکمیلی هم بررسی شده است: آیا ضخامت (نمره‌ی) نخ دوخت اثر دارد؟ نمره‌ی نخ با واحد Ne بیان '
     'می‌شود و نکته‌ی مهم این است که هرچه عدد Ne بزرگ‌تر باشد، نخ ظریف‌تر است. پس Ne 40 از Ne 20 نازک‌تر است.')
P(g, 'پاسخ کوتاه پروژه، که در بخش‌های بعد مستند می‌شود، این است: جنس پارچه بسیار مهم است، کشش دوخت '
     'فقط وقتی زیاد باشد و آن هم عمدتاً روی پلی‌استر مهم است، شستشوی دوم قطعاً آبرفت را افزایش می‌دهد، '
     'و نمره‌ی نخ در بازه‌ی بررسی‌شده اثر معناداری ندارد.')

# ---------------------------------------------------------------- ۲
H(g, 'بخش ۲: فایل اکسل اولیه چه چیزی در خود داشت؟', 1)
P(g, f'فایل اولیه «{A.SRC}» شش کاربرگ دارد. دانستن نقش هر کاربرگ برای ارائه لازم است، چون داور معمولاً '
     'می‌پرسد «از کدام داده استفاده کردی؟».')
sheets_df = pd.DataFrame({
    'نام کاربرگ': ['دوک 20.2', 'دوک40.2', 'پارچه‌ی پنبه', 'پارچه پلی استر', 'نمونه پنبه', 'نمونه پلی استر'],
    'محتوا': ['اندازه‌گیری نمره‌ی واقعی نخ دوک ۲۰', 'اندازه‌گیری نمره‌ی واقعی نخ دوک ۴۰',
              'تراکم و نمره‌ی نخ تار و پود پارچه‌ی پنبه', 'تراکم و نمره‌ی نخ تار و پود پارچه‌ی پلی‌استر',
              'اندازه‌گیری آبرفت نمونه‌های دوخته‌شده‌ی پنبه', 'اندازه‌گیری آبرفت نمونه‌های دوخته‌شده‌ی پلی‌استر'],
    'نقش در تحلیل': ['شناسایی ماده', 'شناسایی ماده', 'شناسایی ماده', 'شناسایی ماده',
                     'داده‌ی اصلی تحلیل', 'داده‌ی اصلی تحلیل']})
TBL(g, sheets_df, 'جدول الف. شش کاربرگ فایل اولیه و نقش هر کدام.',
    'چهار کاربرگ نخست مشخصات مواد را توصیف می‌کنند و وارد آزمون‌های آماری نشده‌اند؛ تحلیل آماری فقط روی دو '
    'کاربرگ آخر انجام شده است.', size=11)
P(g, 'در دو کاربرگ اصلی، ستون‌های زیر اهمیت دارند:')
cols_df = pd.DataFrame({
    'ستون در فایل اصلی': ['کد نمونه', 'پارچه', 'نمره نخ قرقره (Ne)', 'کشش', 'تکرار',
                          'درصد آبرفت کل در شستشو اول', 'درصد آبرفت کل در شستشو دوم'],
    'معنی': ['شناسه‌ی یکتای هر نمونه، مثل C20L-1', 'جنس پارچه: پنبه یا پلی‌استر',
             'ضخامت نخ دوخت: ۲۰ یا ۴۰', 'کشش هنگام دوخت: کم، متوسط یا زیاد',
             'شماره‌ی تکرار: ۱، ۲ یا ۳', 'متغیر پاسخ پس از شستشوی اول',
             'متغیر پاسخ پس از شستشوی دوم']})
TBL(g, cols_df, 'جدول ب. ستون‌های کلیدی و معنی آن‌ها.', size=11)
P(g, 'کد نمونه ساختار معناداری دارد و بهتر است در ارائه آن را توضیح دهید. برای نمونه در کد C20L-1 '
     'حرف اول جنس پارچه است (C برای پنبه و P برای پلی‌استر)، عدد بعدی نمره‌ی نخ است (۲۰ یا ۴۰)، حرف '
     'سوم سطح کشش است (L برای کم، S برای متوسط و H برای زیاد) و عدد پس از خط تیره شماره‌ی تکرار است. '
     'پس C20L-1 یعنی «پارچه‌ی پنبه، نخ Ne 20، کشش کم، تکرار اول».')

# ---------------------------------------------------------------- ۳
H(g, 'بخش ۳: طرح آزمایش — چرا دقیقاً ۳۶ نمونه؟', 1)
P(g, 'این پژوهش از یک «طرح عاملی کامل» استفاده کرده است. یعنی همه‌ی ترکیب‌های ممکن از عامل‌ها ساخته '
     'شده‌اند، نه فقط بعضی از آن‌ها. حساب آن ساده است:')
BUL(g, '۲ جنس پارچه (پنبه، پلی‌استر)')
BUL(g, 'ضربدر ۲ نمره نخ (Ne 20، Ne 40)')
BUL(g, 'ضربدر ۳ سطح کشش (کم، متوسط، زیاد)')
BUL(g, 'مساوی ۱۲ ترکیب یا «خانه»ی آزمایشی')
BUL(g, 'ضربدر ۳ تکرار در هر خانه، مساوی ۳۶ نمونه‌ی دوخته‌شده')
BUL(g, 'و چون هر نمونه دو بار (پس از شستشوی اول و دوم) اندازه‌گیری شده است، در مجموع ۷۲ عدد آبرفت داریم.')
P(g, 'دو نکته‌ای که ارزش گفتن در ارائه را دارند: نخست اینکه طرح «متوازن» است، یعنی در همه‌ی خانه‌ها '
     'دقیقاً ۳ تکرار وجود دارد و هیچ خانه‌ای کم و زیاد نیست. طرح متوازن باعث می‌شود آزمون‌های آماری '
     'قابل‌اعتمادتر و نسبت به نقض مفروضات مقاوم‌تر باشند. دوم اینکه هیچ داده‌ی گمشده‌ای وجود ندارد؛ '
     'هر ۷۲ عدد موجود است.')
P(g, 'چرا تکرار لازم است؟ چون اگر فقط یک نمونه از هر ترکیب می‌ساختیم، هیچ راهی نداشتیم بفهمیم تفاوت '
     'بین دو عدد به‌خاطر عامل مورد بررسی است یا صرفاً پراکندگی طبیعی کار آزمایشگاهی. تکرار به ما «خطای '
     'آزمایش» را می‌دهد و همه‌ی آزمون‌های آماری در واقع اثر عامل را با همین خطا مقایسه می‌کنند.')

# ---------------------------------------------------------------- ۴
H(g, 'بخش ۴: متغیر پاسخ — «درصد آبرفت» چگونه محاسبه می‌شود؟', 1)
P(g, 'روی هر نمونه، طول مشخصی از درز (۲۰۰ میلی‌متر) پیش از شستشو علامت‌گذاری شده است. پس از شستشو '
     'همان فاصله دوباره اندازه‌گیری می‌شود و درصد کاهش طول محاسبه می‌گردد:')
P(g, 'درصد آبرفت = (طول اولیه منهای طول پس از شستشو) تقسیم بر طول اولیه، ضربدر ۱۰۰',
  size=13.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY, before=6, after=6)
P(g, 'برای نمونه اگر فاصله‌ی ۲۰۰ میلی‌متری پس از شستشو به ۱۹۵ میلی‌متر برسد، آبرفت برابر پنج تقسیم بر '
     'دویست ضربدر صد، یعنی ۲٫۵ درصد است. عدد بزرگ‌تر یعنی آبرفت بیشتر و کیفیت بدتر.')
P(g, 'یک نکته‌ی مهم که حتماً باید بدانید: چون اندازه‌گیری با دقت نیم‌میلی‌متری روی طول ۲۰۰ میلی‌متر '
     'انجام شده، همه‌ی مقادیر آبرفت مضربی از ۰٫۵ درصد هستند (۰، ۰٫۵، ۱، ۱٫۵ و …). به این می‌گویند '
     'داده‌ی «شبکه‌ای» یا گسسته. این موضوع بعداً در بخش مفروضات آماری دوباره مطرح می‌شود و دلیل آن است '
     'که آزمون نرمال بودن رد می‌شود. اگر داور پرسید «چرا داده‌هایت نرمال نیستند؟»، پاسخ درست همین است.')

# ---------------------------------------------------------------- ۵
H(g, 'بخش ۵: چه داده‌هایی کنار گذاشته شد و چرا؟', 1)
P(g, 'در دو کاربرگ اصلی، در انتهای جدول‌ها چند سطر اضافه وجود دارد که وارد تحلیل آماری نشده‌اند. '
     'کنار گذاشتن آن‌ها تصمیمی آگاهانه و قابل‌دفاع است:')
BUL(g, 'سطرهای «Control 1» تا «Control 3» (نمونه‌های شاهد): این نمونه‌ها دوخت ندارند و عددی که برای '
       'آن‌ها ثبت شده «نمره‌ی ارزیاب‌ها» است، یعنی امتیازی چشمی بین ۱ تا ۵، نه درصد آبرفت اندازه‌گیری‌شده. '
       'مخلوط‌کردن یک امتیاز چشمی با یک کمیت فیزیکی از نظر آماری نادرست است.')
BUL(g, 'سطرهای «نخ دوک پلی استر»: این‌ها مربوط به خود دوک نخ هستند، نه نمونه‌ی دوخته‌شده، و ستون آبرفت '
       'آن‌ها خالی است.')
P(g, 'اگر داور بپرسد «چرا نمونه‌های شاهد را در ANOVA نیاوردی؟»، پاسخ این است که واحد اندازه‌گیری آن‌ها '
     'با متغیر پاسخ یکی نیست؛ اما می‌توان به‌عنوان اطلاعات جانبی گفت که میانگین نمره‌ی ارزیابان برای '
     'شاهد پنبه حدود ۳٫۶۷ و برای شاهد پلی‌استر حدود ۴٫۶۷ از ۵ بوده است؛ یعنی حتی ارزیابی چشمی هم همان '
     'جهت‌گیری نتایج آماری را تأیید می‌کند: پلی‌استر وضعیت بهتری دارد.')

# ---------------------------------------------------------------- ۶
PAGEBREAK(g)
H(g, 'بخش ۶: مفاهیم آماری به زبان ساده', 1)
P(g, 'این بخش مهم‌ترین قسمت برای آمادگی ارائه است. هر مفهوم را با یک جمله‌ی ساده و یک مثال از همین '
     'پروژه توضیح می‌دهیم.')

H(g, '۶-۱. میانگین و انحراف معیار', 2)
P(g, 'میانگین، عدد نماینده‌ی یک گروه است. انحراف معیار می‌گوید داده‌های آن گروه چقدر حول میانگین '
     'پراکنده‌اند. انحراف معیار کوچک یعنی نمونه‌ها شبیه هم رفتار کرده‌اند و نتیجه پایدار است؛ انحراف '
     f'معیار بزرگ یعنی پراکندگی زیاد بوده است. برای مثال میانگین آبرفت پنبه {M_COTTON:.2f} درصد است، '
     'یعنی به‌طور متوسط درزهای پنبه این‌قدر جمع شده‌اند.')

H(g, '۶-۲. میله‌ی خطا روی نمودارها', 2)
P(g, 'خط عمودی کوچکی که روی هر ستون نمودار دیده می‌شود، «میله‌ی خطا» است و در این پروژه اندازه‌ی آن '
     'یک انحراف معیار است. کاربرد عملی آن در ارائه این است: اگر میله‌های خطای دو ستون هم‌پوشانی زیادی '
     'داشته باشند، تفاوت آن دو ستون احتمالاً معنادار نیست؛ و اگر کاملاً از هم جدا باشند، تفاوت احتمالاً '
     'واقعی است. البته قضاوت نهایی با آزمون آماری است، نه با چشم.')

H(g, '۶-۳. فرض صفر و فرض مقابل', 2)
P(g, 'هر آزمون آماری با یک «فرض صفر» شروع می‌شود که همیشه می‌گوید «تفاوتی وجود ندارد». مثلاً: میانگین '
     'آبرفت پنبه و پلی‌استر برابر است. فرض مقابل می‌گوید تفاوت وجود دارد. کار آزمون این است که ببیند '
     'داده‌ها چقدر با فرض صفر ناسازگارند. اگر ناسازگاری به قدر کافی زیاد باشد، فرض صفر رد می‌شود.')

H(g, '۶-۴. مقدار p و معنی درست آن', 2)
P(g, 'مقدار p یعنی: «اگر واقعاً هیچ تفاوتی وجود نداشت، احتمال اینکه تصادفاً تفاوتی به بزرگی آنچه دیدیم '
     'مشاهده کنیم چقدر بود؟» اگر این احتمال خیلی کم باشد (کمتر از ۰٫۰۵)، نتیجه می‌گیریم تفاوت واقعی '
     'است. عدد ۰٫۰۵ یک قرارداد رایج است، نه یک قانون طبیعت.')
P(g, 'دو اشتباه رایج که نباید در ارائه مرتکب شوید: نخست، p احتمال درست بودن فرض صفر نیست. دوم، '
     '«معنادار» به معنی «مهم» یا «بزرگ» نیست؛ معناداری فقط می‌گوید تفاوت احتمالاً تصادفی نیست. برای '
     'اینکه بفهمیم تفاوت چقدر بزرگ است، باید به «اندازه اثر» نگاه کنیم که در ادامه می‌آید.')

H(g, '۶-۵. تحلیل واریانس (ANOVA) و آماره F', 2)
P(g, 'تحلیل واریانس روشی است برای مقایسه‌ی میانگین دو گروه یا بیشتر. منطق آن را می‌توان با یک کسر '
     'توضیح داد: آماره‌ی F برابر است با «تفاوت بین گروه‌ها» تقسیم بر «پراکندگی درون گروه‌ها». اگر '
     'تفاوت بین گروه‌ها نسبت به نوسان طبیعی درون هر گروه بزرگ باشد، F بزرگ می‌شود و p کوچک. پس F بزرگ '
     'یعنی سیگنال قوی‌تر از نویز.')
P(g, 'به همراه F دو عدد به شکل «درجات آزادی» گزارش می‌شود، مثلاً F(1, 70). عدد اول به تعداد گروه‌ها '
     'مربوط است (تعداد گروه‌ها منهای یک) و عدد دوم به تعداد داده‌ها (تعداد کل داده منهای تعداد گروه‌ها). '
     'در این پروژه برای مقایسه‌ی دو پارچه با ۷۲ داده، درجات آزادی برابر ۱ و ۷۰ است.')

H(g, '۶-۶. اندازه اثر: اتا مربع و d کوهن', 2)
P(g, 'اتا مربع می‌گوید چند درصد از کل تغییرپذیری متغیر پاسخ توسط آن عامل توضیح داده می‌شود. مقیاس '
     'مرسوم آن چنین است: حدود ۰٫۰۱ اثر کوچک، ۰٫۰۶ متوسط و ۰٫۱۴ به بالا اثر بزرگ. در این پروژه اتا مربع '
     f'نوع پارچه {val("A1","eta2")} است، یعنی حدود {val("A1","eta2")*100:.0f} درصد از تغییرات آبرفت فقط '
     'به جنس الیاف مربوط می‌شود — اثری بسیار بزرگ.')
P(g, 'd کوهن همین ایده را برای مقایسه‌ی دو گروه بیان می‌کند: فاصله‌ی دو میانگین بر حسب انحراف معیار. '
     'حدود ۰٫۲ کوچک، ۰٫۵ متوسط و ۰٫۸ به بالا بزرگ است.')
P(g, 'چرا این مهم است؟ چون با نمونه‌ی خیلی بزرگ، حتی تفاوت‌های بی‌اهمیت هم معنادار می‌شوند. گزارش '
     'اندازه اثر در کنار p نشان می‌دهد که تفاوت علاوه بر واقعی بودن، از نظر عملی هم اهمیت دارد.')

H(g, '۶-۷. آزمون تعقیبی و مسئله‌ی مقایسه‌های چندگانه', 2)
P(g, 'وقتی ANOVA برای عاملی با سه سطح (مثل کشش کم، متوسط و زیاد) معنادار می‌شود، فقط می‌فهمیم «دست‌کم '
     'یکی از سه گروه با بقیه فرق دارد»، اما نمی‌دانیم کدام. برای یافتن آن، مقایسه‌های دوبه‌دو انجام '
     'می‌دهیم که به آن آزمون تعقیبی می‌گویند.')
P(g, 'اما یک مشکل وجود دارد: هر بار که آزمونی انجام می‌دهیم، پنج درصد احتمال هشدار کاذب داریم. با سه '
     'مقایسه، این خطر جمع می‌شود. «تصحیح بونفرونی» ساده‌ترین راه‌حل است: مقدار p هر مقایسه را در تعداد '
     'مقایسه‌ها (اینجا ۳) ضرب می‌کنیم و سپس با ۰٫۰۵ می‌سنجیم. این کار آزمون را سخت‌گیرتر می‌کند تا نتیجه‌ی '
     'کاذب نگیریم.')

H(g, '۶-۸. بررسی مفروضات: آزمون لِوین و شاپیرو-ویلک', 2)
P(g, 'تحلیل واریانس دو فرض اصلی دارد: پراکندگی گروه‌ها تقریباً برابر باشد (همگنی واریانس) و داده‌ها '
     'تقریباً نرمال توزیع شده باشند. آزمون لِوین فرض اول و آزمون شاپیرو-ویلک فرض دوم را بررسی می‌کند. '
     'در این دو آزمون، p بزرگ‌تر از ۰٫۰۵ خبر خوب است، چون یعنی فرض نقض نشده است.')
P(g, 'در این پروژه لِوین تقریباً همه‌جا خوب بود، اما شاپیرو-ویلک در بیشتر گروه‌ها نرمال بودن را رد کرد. '
     'همان‌طور که در بخش ۴ توضیح داده شد، علت آن ثبت داده روی شبکه‌ی ۰٫۵ درصد است. به همین دلیل تحلیل '
     'با یک آزمون جایگزین هم تکرار شد.')

H(g, '۶-۹. آزمون کروسکال-والیس (جایگزین ناپارامتری)', 2)
P(g, 'این آزمون همان کار ANOVA را انجام می‌دهد اما به‌جای خود اعداد، با رتبه‌ی آن‌ها کار می‌کند و '
     'بنابراین هیچ فرضی درباره‌ی نرمال بودن ندارد. در این پروژه همه‌ی آزمون‌ها با کروسکال-والیس تکرار '
     'شدند و نتیجه دقیقاً همان بود. این یعنی نتایج پروژه محکم هستند و به فرض نرمال بودن وابسته نیستند. '
     'این نکته یکی از قوی‌ترین حرف‌هایی است که می‌توانید در دفاع بزنید.')

H(g, '۶-۱۰. آزمون t زوجی و چرا اینجا لازم بود', 2)
P(g, 'شستشوی اول و دوم روی نمونه‌های متفاوتی انجام نشده‌اند؛ هر دو روی همان ۳۶ نمونه انجام شده‌اند. به '
     'این حالت «اندازه‌گیری تکراری» می‌گویند. ANOVA معمولی این دو گروه را مستقل فرض می‌کند که کاملاً '
     'دقیق نیست (هرچند محافظه‌کارانه است و نتیجه را ضعیف‌تر نشان می‌دهد، نه قوی‌تر). آزمون t زوجی برای '
     'همین حالت ساخته شده است: به‌جای مقایسه‌ی دو میانگین، تغییر درون هر نمونه را می‌سنجد. در گزارش هر '
     'دو آزمون آمده‌اند و هر دو یک نتیجه می‌دهند. اگر داور روی این نکته دست گذاشت، شما پیشاپیش پاسخ '
     'داده‌اید — و این امتیاز مثبتی است.')

# ---------------------------------------------------------------- ۷
PAGEBREAK(g)
H(g, 'بخش ۷: نمودارها یکی‌یکی — هر کدام چه می‌گویند؟', 1)
P(g, 'پیش از توضیح تک‌تک نمودارها، چند نکته‌ی مشترک درباره‌ی همه‌ی آن‌ها را بدانید. در همه‌ی نمودارها '
     'محور عمودی درصد آبرفت کل دوخت است و هرچه ستون بلندتر باشد، آبرفت بیشتر و کیفیت بدتر است. ارتفاع '
     'هر ستون میانگین آن گروه است و عدد چاپ‌شده بالای ستون همان میانگین است. خط عمودی روی هر ستون '
     'میله‌ی خطا به اندازه‌ی یک انحراف معیار است. رنگ‌ها معنی ثابتی دارند و در راهنمای زیر هر نمودار '
     'توضیح داده شده‌اند.')

GUIDE = {
 'نمودار1': (
    'این نمودار ساده‌ترین و مهم‌ترین پیام پروژه را نشان می‌دهد: پس از یک بار شستشو، درز روی پارچه‌ی پنبه '
    'به‌طور میانگین حدود {c1:.2f} درصد جمع شده، در حالی که همین عدد برای پلی‌استر {p1:.2f} درصد است؛ یعنی '
    'تقریباً چند برابر. محور افقی فقط دو گروه دارد و مقایسه مستقیم است. اگر بخواهید ارائه را با یک '
    'تصویر شروع کنید، همین نمودار بهترین گزینه است.',
    'وقتی این نمودار را نشان می‌دهید بگویید: «تفاوت این دو ستون با ANOVA بررسی شده و معنادار است، '
    'F({a2df}) = {a2f} و {a2p}.»'),
 'نمودار2': (
    'همان مقایسه، اما این بار پس از شستشوی دوم. دو نکته را باید ببینید: نخست اینکه هر دو ستون نسبت به '
    'نمودار قبل بلندتر شده‌اند (یعنی شستشوی دوم آبرفت را اضافه کرده است) و دوم اینکه ترتیب عوض نشده و '
    'پنبه همچنان بالاتر است. یعنی برتری پلی‌استر با تکرار شستشو از بین نمی‌رود.',
    'نکته‌ی مهم برای ارائه: مقایسه‌ی این دو نمودار کنار هم نشان می‌دهد که نتیجه‌گیری ما به تعداد شستشو '
    'وابسته نیست و در هر دو نوبت برقرار است.'),
 'نمودار3': (
    'اینجا محور افقی سه سطح کشش دوخت است و در هر سطح دو ستون داریم: آبی برای پنبه و نارنجی برای '
    'پلی‌استر. این نمودار مهم‌ترین یافته‌ی ظریف پروژه را نشان می‌دهد. به ستون‌های نارنجی نگاه کنید: در '
    'کشش کم و متوسط تقریباً صفر هستند، اما در کشش زیاد ناگهان بالا می‌روند. حالا به ستون‌های آبی نگاه '
    'کنید: تقریباً هم‌ارتفاع‌اند و با تغییر کشش تغییر چندانی نمی‌کنند.',
    'پیام: کشش دوخت روی پلی‌استر اثر دارد ولی روی پنبه نه. دلیل فنی آن در بخش ۹ آمده است. این همان '
    'چیزی است که در آمار به آن «اثر متقابل» می‌گویند و نکته‌ی برجسته‌ی ارائه‌ی شماست.'),
 'نمودار4': (
    'همان ساختار نمودار قبل، اما پس از شستشوی دوم. الگو تکرار می‌شود و حتی واضح‌تر می‌شود: ستون '
    'نارنجی کشش زیاد به‌شکل چشمگیری از دو ستون نارنجی دیگر بالاتر است.',
    'اینکه یک الگو در دو شستشوی مستقل تکرار شود، شاهد خوبی بر واقعی بودن آن است. این جمله را در ارائه '
    'بگویید.'),
 'نمودار5': (
    'در این نمودار دو پارچه با هم ادغام شده‌اند و محور افقی سه سطح کشش است؛ ستون آبی شستشوی اول و ستون '
    'نارنجی شستشوی دوم را نشان می‌دهد. دو روند همزمان دیده می‌شود: از چپ به راست (کم به زیاد) ستون‌ها '
    'بلندتر می‌شوند، و در هر جفت، ستون نارنجی از آبی بلندتر است.',
    'یعنی هم کشش بیشتر و هم شستشوی بیشتر، هر دو آبرفت را افزایش می‌دهند و اثرشان روی هم جمع می‌شود.'),
 'نمودار6': (
    'این نمودار مستقیماً به پرسش سوم پروژه پاسخ می‌دهد. برای هر پارچه دو ستون داریم: شستشوی اول و دوم. '
    'در هر دو پارچه ستون دوم بلندتر است، یعنی آبرفت پس از شستشوی دوم بیشتر شده است.',
    'عدد کلیدی برای گفتن: میانگین کل از {w1:.2f} درصد به {w2:.2f} درصد رسیده و مهم‌تر از آن، هر ۳۶ نمونه '
    'بدون استثنا در شستشوی دوم بیشتر جمع شده‌اند. نتیجه‌ی کاربردی: ارزیابی کیفیت دوخت با یک بار شستشو '
    'گمراه‌کننده است.'),
 'نمودار7': (
    'آخرین نمودار به پرسش تکمیلی می‌پردازد: آیا ضخامت نخ دوخت مهم است؟ محور افقی دو نمره‌ی نخ Ne 20 و '
    'Ne 40 است و رنگ‌ها دو شستشو را از هم جدا می‌کنند. ستون‌های Ne 40 اندکی کوتاه‌تر از Ne 20 هستند، '
    'یعنی نخ ظریف‌تر کمی بهتر عمل کرده است.',
    'اما توجه کنید که میله‌های خطا هم‌پوشانی زیادی دارند و آزمون آماری هم این تفاوت را معنادار ندانست '
    '({d1p}). پس پاسخ درست این است: «در بازه‌ی بررسی‌شده تفاوت معناداری دیده نشد.» گفتن اینکه «نخ '
    'ظریف‌تر بهتر است» بدون این قید، خطای علمی است.'),
}
fmt = dict(c1=TABLES['T2_Fabric_x_Wash'].loc[0, 'Mean'], p1=TABLES['T2_Fabric_x_Wash'].loc[2, 'Mean'],
           a2df=val('A2', 'df'), a2f=val('A2', 'F'), a2p=pstr('A2'),
           w1=M_W1, w2=M_W2, d1p=pstr('D1'))
for path, cap in FA_CHARTS:
    key = os.path.basename(path).split('_')[0]
    body, tip = GUIDE[key]
    H(g, cap.split('.')[0] + f' — {os.path.basename(path)}', 2)
    FIG(g, path, cap, width_cm=13.2)
    P(g, body.format(**fmt))
    P(g, 'نکته‌ی ارائه: ' + tip.format(**fmt), italic=True, color=BLUEC)

# ---------------------------------------------------------------- ۸
PAGEBREAK(g)
H(g, 'بخش ۸: جدول‌ها چه چیزی نشان می‌دهند؟', 1)
P(g, 'در فایل اکسل خروجی و در گزارش، چند دسته جدول وجود دارد. لازم نیست همه‌ی اعداد را حفظ کنید؛ کافی '
     'است بدانید هر جدول برای چیست.')
tbl_guide = pd.DataFrame({
    'جدول': ['آبرفت به تفکیک پارچه', 'پارچه و شستشو', 'کشش و شستشو', 'پارچه، کشش و شستشو',
             'نمره نخ و شستشو', 'نوبت شستشو', 'طرح کامل عاملی', 'نتایج آنالیز واریانس',
             'آزمون تعقیبی کشش', 'آزمون کروسکال والیس', 'داده‌های پاک‌سازی‌شده'],
    'به چه پرسشی پاسخ می‌دهد': [
        'به‌طور کلی کدام پارچه بیشتر جمع می‌شود؟',
        'این تفاوت در هر شستشو چقدر است؟',
        'با افزایش کشش، آبرفت چگونه تغییر می‌کند؟',
        'اثر کشش در هر پارچه جداگانه چگونه است؟',
        'نمره نخ چه تفاوتی ایجاد می‌کند؟',
        'شستشوی دوم چقدر آبرفت اضافه می‌کند؟',
        'میانگین هر یک از ۲۴ حالت ممکن چقدر است؟',
        'کدام عامل‌ها از نظر آماری معنادارند؟',
        'دقیقاً کدام دو سطح کشش با هم تفاوت دارند؟',
        'آیا نتایج بدون فرض نرمال بودن هم برقرارند؟',
        'داده‌ی نهایی که همه‌ی محاسبات روی آن انجام شده چیست؟']})
TBL(g, tbl_guide, 'جدول پ. راهنمای جدول‌های خروجی.', size=11)
P(g, 'در جدول‌های آمار توصیفی، ستون‌ها همیشه یکسان‌اند: «تعداد» شمار مشاهده‌های آن گروه، «میانگین» '
     'مقدار نماینده، «انحراف معیار» پراکندگی، و «کمینه» و «بیشینه» کوچک‌ترین و بزرگ‌ترین مقدار مشاهده‌شده. '
     'در جدول ANOVA، ستون‌های کلیدی «آماره F»، «مقدار p» و «اندازه اثر» هستند و ستون آخر تفسیر آماده را '
     'می‌دهد.')

# ---------------------------------------------------------------- ۹
H(g, 'بخش ۹: نتایج اصلی و تفسیر فنی آن‌ها', 1)
H(g, '۹-۱. چرا پنبه بیشتر آبرفت دارد؟', 2)
P(g, f'یافته: پنبه {M_COTTON:.2f} درصد در برابر {M_POLY:.2f} درصد پلی‌استر، با {pstr("A1")} و اتا مربع '
     f'{val("A1","eta2")}؛ یعنی بزرگ‌ترین اثر در کل پروژه.')
P(g, 'تفسیر فنی: الیاف پنبه سلولزی و آب‌دوست هستند. هنگام شستشو آب را جذب می‌کنند و متورم می‌شوند؛ قطر '
     'الیاف افزایش می‌یابد و در نتیجه نخ‌ها در ساختار پارچه مسیر پرپیچ‌وخم‌تری طی می‌کنند و طول کلی کم '
     'می‌شود. افزون بر این، تنش‌هایی که در مراحل ریسندگی، بافندگی و تکمیل در الیاف ذخیره شده‌اند، در '
     'محیط مرطوب و گرم آزاد می‌شوند و الیاف به حالت آسوده‌تر خود بازمی‌گردند. پلی‌استر در مقابل، الیافی '
     'گرمانرم و آب‌گریز است که در فرایند تثبیت حرارتی ابعادش تثبیت شده و در دمای شستشوی معمولی تقریباً '
     'هیچ تغییری نمی‌کند.')
H(g, '۹-۲. چرا کشش دوخت فقط در پلی‌استر اثر دارد؟', 2)
P(g, f'یافته: در پلی‌استر {pstr("B5")} با اتا مربع {val("B5","eta2")}، اما در پنبه {pstr("B4")}؛ یعنی هیچ '
     'اثری دیده نشد.')
P(g, 'تفسیر فنی: وقتی نخ دوخت با کشش زیاد به پارچه دوخته می‌شود، نخ در حالت کشیده و تحت تنش تثبیت '
     'می‌شود؛ مثل فنری که جمع شده باشد. شستشو با گرما و رطوبت این تنش را آزاد می‌کند و نخ می‌خواهد به '
     'طول آسوده‌ی خود برگردد و در نتیجه درز را جمع می‌کند. این سازوکار در هر دو پارچه وجود دارد، اما در '
     'پنبه آبرفت خود الیاف آن‌قدر بزرگ است که سهم کشش نخ زیر آن گم می‌شود؛ در پلی‌استر که آبرفت الیاف '
     'تقریباً صفر است، همین سازوکار به عامل اصلی و قابل مشاهده تبدیل می‌شود.')
P(g, 'این نکته پیام کاربردی مهمی دارد: تنظیم درست کشش چرخ خیاطی، دقیقاً روی پارچه‌های مصنوعی بیشترین '
     'اهمیت را دارد؛ یعنی جایی که معمولاً تصور می‌شود «پارچه پایدار است پس مشکلی پیش نمی‌آید».')
H(g, '۹-۳. چرا شستشوی دوم مهم است؟', 2)
P(g, f'یافته: میانگین از {M_W1:.2f} به {M_W2:.2f} درصد رسید، {pstr("C1")}، و آزمون t زوجی افزایش '
     f'{paired["mean_increase"]} واحد درصدی را با p {paired["p"]} تأیید کرد. هر ۳۶ نمونه بدون استثنا '
     'افزایش نشان دادند.')
P(g, 'تفسیر فنی: آزادسازی تنش و بازآرایی ساختار در یک چرخه‌ی شستشو کامل نمی‌شود و بخشی از آن به '
     'چرخه‌های بعد موکول می‌گردد. نتیجه‌ی عملی: استانداردهای ارزیابی پایداری ابعادی که تنها یک شستشو را '
     'مبنا قرار می‌دهند، آبرفت نهایی را کمتر از واقع برآورد می‌کنند.')
H(g, '۹-۴. چرا نمره نخ اثر نداشت؟', 2)
P(g, f'یافته: {pstr("D1")} و اتا مربع {val("D1","eta2")}، یعنی عملاً بدون اثر.')
P(g, 'تفسیر: بازه‌ی بررسی‌شده (Ne 20 تا Ne 40) احتمالاً برای آشکارشدن اثر ضخامت نخ باریک است، و اثر آن '
     'در برابر دو عامل بزرگ‌تر (جنس الیاف و کشش زیاد) ناچیز می‌ماند. توجه کنید که «معنادار نبودن» به '
     'معنی «اثبات نبودن اثر» نیست، بلکه یعنی با این حجم نمونه شواهدی برای آن نیافتیم.')

# ---------------------------------------------------------------- ۱۰
H(g, 'بخش ۱۰: محدودیت‌های پژوهش', 1)
P(g, 'گفتن محدودیت‌ها ضعف نیست؛ نشانه‌ی بلوغ علمی است و معمولاً نمره‌ی مثبت دارد. محدودیت‌های اصلی این پروژه:')
BUL(g, 'حجم نمونه کوچک است: تنها ۳ تکرار در هر خانه‌ی آزمایشی.')
BUL(g, 'دقت اندازه‌گیری ۰٫۵ درصد است؛ تغییرات ریزتر از این قابل تشخیص نبوده‌اند.')
BUL(g, 'تنها دو نمره نخ و یک نوع بخیه بررسی شده است؛ تعمیم نتیجه به نمره‌ها و بخیه‌های دیگر نیازمند آزمایش بیشتر است.')
BUL(g, 'اثر متقابل پارچه و کشش با تحلیل زیرگروهی نشان داده شده، نه با تحلیل واریانس دوعاملی؛ برای '
       'کمّی‌کردن دقیق آن به طرح و نمونه‌ی بزرگ‌تر نیاز است.')
BUL(g, 'شرایط شستشو (دما، مواد شوینده، تعداد دور) ثابت فرض شده و اثر آن‌ها بررسی نشده است.')

# ---------------------------------------------------------------- ۱۱
PAGEBREAK(g)
H(g, 'بخش ۱۱: راهنمای ارائه', 1)
H(g, '۱۱-۱. ساختار پیشنهادی اسلایدها', 2)
slides = pd.DataFrame({
    'اسلاید': ['۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹', '۱۰'],
    'موضوع': ['عنوان و معرفی', 'آبرفتگی دوخت چیست و چرا مهم است',
              'طرح آزمایش و نحوه‌ی اندازه‌گیری', 'روش آماری در یک نگاه',
              'یافته ۱: اثر جنس پارچه', 'یافته ۲: اثر کشش دوخت و اثر متقابل',
              'یافته ۳: اثر شستشوی دوم', 'یافته تکمیلی: نمره نخ',
              'نتیجه‌گیری و پیام کاربردی', 'محدودیت‌ها و پیشنهاد ادامه'],
    'چه چیزی نشان دهید': ['—', 'یک تصویر از درز جمع‌شده', 'جدول ۱ و فرمول درصد آبرفت',
                          'یک جمله درباره‌ی ANOVA و p', 'نمودار ۱ و ۲', 'نمودار ۳ و ۴',
                          'نمودار ۶', 'نمودار ۷', 'سه جمله‌ی کلیدی', '—']})
TBL(g, slides, 'جدول ت. ساختار پیشنهادی ارائه در ده اسلاید.', size=11)
H(g, '۱۱-۲. سه جمله‌ای که باید حفظ باشید', 2)
P(g, f'۱) جنس الیاف تعیین‌کننده‌ی اصلی است: پنبه {M_COTTON:.2f} درصد در برابر {M_POLY:.2f} درصد پلی‌استر، '
     f'{pstr("A1")}، و به‌تنهایی حدود {val("A1","eta2")*100:.0f} درصد تغییرپذیری را توضیح می‌دهد.', bold=True)
P(g, '۲) کشش دوخت تنها در تنظیم «زیاد» و تقریباً فقط روی پلی‌استر اثرگذار است؛ این یک اثر متقابل است.', bold=True)
P(g, f'۳) یک شستشو کافی نیست: آبرفت از {M_W1:.2f} به {M_W2:.2f} درصد رسید و هر ۳۶ نمونه افزایش نشان دادند.', bold=True)
H(g, '۱۱-۳. پرسش‌های محتمل داور و پاسخ آماده', 2)
qa = pd.DataFrame({
    'پرسش احتمالی': [
        'چرا از ANOVA استفاده کردی و نه آزمون t؟',
        'داده‌هایت نرمال نیستند، پس نتیجه‌ات معتبر است؟',
        'چرا نمونه‌های شاهد را وارد تحلیل نکردی؟',
        'تعداد نمونه کم نیست؟',
        'چرا اثر کشش در پنبه دیده نشد؟',
        'دو شستشو روی یک نمونه بوده؛ آیا ANOVA درست است؟',
        'اثر متقابل را چرا با ANOVA دوعاملی بررسی نکردی؟',
        'اندازه اثر یعنی چه و چرا گزارشش کردی؟'],
    'پاسخ کوتاه': [
        'برای عامل کشش که سه سطح دارد آزمون t کافی نیست؛ ANOVA همه‌ی سطوح را همزمان می‌سنجد. برای عامل‌های دوسطحی نیز ANOVA و t نتیجه‌ی یکسان می‌دهند.',
        'بله؛ چون داده روی شبکه‌ی ۰٫۵ درصد ثبت شده و طبعاً نرمال نیست. به همین دلیل همه‌ی آزمون‌ها با کروسکال-والیس تکرار شد و نتایج تغییر نکرد.',
        'چون مقدار ثبت‌شده برای آن‌ها نمره‌ی چشمی ارزیابان است، نه درصد آبرفت اندازه‌گیری‌شده؛ واحدشان با متغیر پاسخ یکی نیست.',
        'حجم نمونه محدود است، اما طرح کاملاً متوازن است و اثرهای یافت‌شده اندازه اثر بزرگی دارند که با نمونه‌ی کوچک هم قابل اتکاست.',
        'چون آبرفت خود الیاف پنبه بزرگ است و سهم کشش نخ را می‌پوشاند؛ در پلی‌استر که آبرفت الیاف تقریباً صفر است این سهم آشکار می‌شود.',
        'به‌طور دقیق باید تحلیل اندازه‌گیری تکراری کرد. ANOVA در این حالت محافظه‌کارانه است و افزون بر آن، آزمون t زوجی هم اجرا و گزارش شده که همان نتیجه را می‌دهد.',
        'در این حجم نمونه، ANOVA دوعاملی توان کافی برای برآورد دقیق اثر متقابل ندارد؛ به‌جای آن تحلیل زیرگروهی گزارش شده و به‌عنوان محدودیت و پیشنهاد پژوهش آینده ذکر شده است.',
        'اندازه اثر می‌گوید تفاوت چقدر بزرگ است، در حالی که p فقط می‌گوید تفاوت تصادفی نیست. گزارش هر دو، تصویر کاملی می‌دهد.']})
TBL(g, qa, 'جدول ث. پرسش‌های محتمل و پاسخ‌های آماده.', size=10.5)

# ---------------------------------------------------------------- ۱۲
H(g, 'بخش ۱۲: فایل‌های پروژه و اجرای دوباره', 1)
files = pd.DataFrame({
    'نام فایل': ['راهنمای_کامل_پروژه.docx', 'گزارش_تحلیل_آماری.docx',
                 'آمار_توصیفی_و_نتایج_آزمون‌ها.xlsx', 'نمودار۱ تا نمودار۷ (png)',
                 'analysis.py', 'persian_outputs.py'],
    'توضیح': ['همین سند: توضیح کامل پروژه برای مطالعه و ارائه',
              'گزارش رسمی فارسی همراه با جدول‌ها، نمودارها و نتیجه‌گیری',
              'همه‌ی جدول‌های توصیفی، نتایج ANOVA، آزمون تعقیبی و کروسکال-والیس',
              'نمودارهای فارسی با کیفیت چاپ (۳۰۰ نقطه بر اینچ)',
              'اسکریپت اصلی محاسبات و خروجی انگلیسی',
              'اسکریپت ساخت همه‌ی خروجی‌های فارسی']})
TBL(g, files, 'جدول ج. فایل‌های پروژه.', size=11)
P(g, 'برای تولید دوباره‌ی همه‌ی خروجی‌ها کافی است در پوشه‌ی پروژه دستور python persian_outputs.py را '
     'اجرا کنید. اگر یکی از فایل‌های خروجی در Word یا Excel باز باشد، برنامه پیام روشنی می‌دهد و از شما '
     'می‌خواهد آن را ببندید. توجه داشته باشید که همه‌ی اعداد گزارش و راهنما به‌طور خودکار از دل همان '
     'محاسبات بیرون می‌آیند؛ بنابراین اگر داده‌ی اولیه اصلاح شود، کافی است اسکریپت دوباره اجرا شود و '
     'تمام جدول‌ها، نمودارها و متن‌ها به‌روز می‌شوند.')

save(g, os.path.join(OUT, 'راهنمای_کامل_پروژه.docx'))
