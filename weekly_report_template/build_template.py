#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
مولّد قالب «التقرير الأسبوعي للمهام» (A4 عمودي، عربي RTL، مستوحى من هوية سدايا).

يبني ملف PPTX فيه:
  - تخطيط (Layout) مخصص اسمه «التقرير الأسبوعي»: العناصر الثابتة (الترويسة، عناوين الأقسام، الصناديق)
    + حقول إرشادية (Placeholders) تظهر في وضع التحرير فقط ولا تُطبع إن بقيت فارغة.
  - شريحة 1: نموذج التقرير (فارغ) مع عداد «تقرير رقم N» مربوط برقم الشريحة.
  - شريحة أخيرة مخفية: دليل الاستخدام (لا تُصدَّر إلى PDF افتراضيًا).

الاستخدام:
  python3 build_template.py --out "out/التقرير_الأسبوعي_قالب.pptx"
  python3 build_template.py --out out/example.pptx --example [--weeks 2]   # نسخة معبأة ببيانات تجريبية
  python3 build_template.py --out x.pptx --font "Tajawal" --body-font "Tajawal"   # تجربة خط آخر
  python3 build_template.py --out qa.pptx --example --table-ltr   # للمعاينة في LibreOffice (يتجاهل rtl الجدول)
"""
import argparse
import copy
import datetime as _dt
import os
import tempfile
from xml.sax.saxutils import escape

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu

# ----------------------------------------------------------------------------
# الإعدادات
# ----------------------------------------------------------------------------
MM = 36000  # EMU per mm
PAGE_W, PAGE_H = 210, 297
ML, MR = 12, 198  # حدود المحتوى (مم)
NS = nsdecls("a", "p", "r")

DEFAULT_FONT = "IBM Plex Sans Arabic"
DEFAULT_BODY_FONT = "IBM Plex Sans Arabic Medium"  # وزن أثقل قليلًا للمتن: أوضح على الهاتف
TABLE_RTL = True  # tblPr rtl="1" + ترتيب منطقي للأعمدة (المهمة أولًا). PowerPoint يعكس الأعمدة بصريًا.

# الألوان الأساسية مأخوذة من ملف شعار سدايا (SVG) ومن موقعها: كحلي #273470، نعناعي #33CC99، ليموني #99D633،
# بنفسجي #6459A7، سماوي #66CCFF. الباقي درجات مشتقة للنص والخلفيات مع تباين ≥ 4.5:1 للنص الصغير.
PALETTE = {
    "navy": "273470",        # لوحة الترويسة (كحلي سدايا)
    "navy2": "1B2452",       # نهاية تدرج الترويسة
    "ink": "1F2A37",         # نص أساسي
    "muted": "5F6B7A",       # نص ثانوي (5.4:1)
    "quiet": "4F5B6A",       # نص قسم التحديات (6.3:1)
    "faint": "6C727E",       # نص التذييل
    "rule": "BFC6D1",        # خطوط الجدول والفواصل (تُطبع بوضوح)
    "edge": "C9CFD9",        # إطار الصناديق الهادئة
    "mint": "33CC99",        # نعناعي: على الكحلي فقط (تسميات/شارة/أيقونات الترويسة)
    "lime": "99D633",        # ليموني: نهاية تدرج الشارة
    "sky": "66CCFF",         # سماوي: زخرفة على الكحلي
    "purple": "6459A7",      # بنفسجي سدايا (قسم «المقبل»)
    "teal_dark": "0C7373",   # تركوازي داكن للتسميات على الأبيض (5.6:1)
    "green_text": "127A5B",  # أخضر داكن للأرقام/الأيقونات على الأبيض (5.3:1)
    "amber": "B54708",       # متأخر/خطر (5.4:1)
    "tint_navy": "EEF0F8",   # خلفية بطاقات الأرقام
    "tint_mint": "E8F7F1",   # صندوق القرار المطلوب
    "tint_gray": "F4F5F7",   # صندوق التحديات (هادئ)
    "white": "FFFFFF",
}

TXT = {
    "title": "التقرير الأسبوعي للمهام",
    "entity_prompt": "اسم الجهة — الإدارة / القسم",
    "logo_prompt": "الشعار",
    "counter_label": "تقرير رقم",
    "period_cap": "الفترة",
    "author_cap": "إعداد",
    "to_cap": "مقدَّم إلى",
    "period_prompt": "من … إلى … (الأسبوع …)",
    "author_prompt": "الاسم — المسمى الوظيفي",
    "to_prompt": "اسم المدير — المسمى الوظيفي",
    "summary_cap": "خلاصة الأسبوع",
    "summary_prompt": "ثلاثة أسطر: أبرز ما أُنجز، وما يجري العمل عليه، وما يحتاج تدخّلًا. الأعداد هنا تطابق البطاقات.",
    "kpi_done": "مهام منجزة",
    "kpi_doing": "مهام جارية",
    "kpi_next": "مهام مقبلة",
    "kpi_prompt": "العدد",
    "sec_done": "المهام المنجزة هذا الأسبوع",
    "done_prompt": "مهمة في كل سطر: اسم المهمة — النتيجة أو الأثر (برقم إن أمكن). خمسة بنود كحد أقصى.",
    "sec_doing": "المهام الجارية",
    "tbl_task": "المهمة",
    "tbl_pct": "نسبة الإنجاز (٪)",
    "tbl_due": "الموعد المستهدف",
    "tbl_status": "الحالة",
    "sec_next": "خطة الأسبوع المقبل",
    "next_prompt": "مهام الأسبوع المقبل، مهمة في كل سطر: النتيجة المتوقعة + اليوم. خمسة بنود كحد أقصى.",
    "decision_cap": "قرار أو دعم مطلوب من المدير",
    "decision_prompt": "المطلوب: … | قبل: (التاريخ) | أثر التأخير: … (اتركه فارغًا إن لم يوجد)",
    "challenges_cap": "التحديات",
    "challenges_prompt": "جملة واحدة: السبب ← الأثر ← الموعد الجديد ← هل يحتاج تدخّلك؟ (اتركه فارغًا إن لم يوجد)",
    "sign_name_cap": "إعداد:",
    "sign_name_prompt": "الاسم",
    "sign_cap": "التوقيع:",
    "sign_date_cap": "التاريخ:",
    "sign_date_prompt": "اليوم الشهر السنة",
    "footer_fixed": "تقرير أسبوعي — للاستخدام الداخلي",
    "layout_name": "التقرير الأسبوعي",
    "blank_layout_name": "صفحة فارغة",
}

EXAMPLE = {
    "entity": "الإدارة العامة للتحول الرقمي — إدارة البيانات والتحليلات",
    "period": "من 25 إلى 29 أكتوبر 2026 (الأسبوع 44)",
    "author": "أحمد العتيبي — محلل بيانات أول",
    "to": "م. خالد القحطاني — مدير إدارة البيانات",
    "summary": "أُطلقت لوحة مؤشرات الأداء واعتمدتها 3 إدارات، وأُنجزت 4 مهام. مهمتان ضمن الجدول، ومهمة "
               "(تكامل الموارد البشرية) متأخرة أسبوعًا. مطلوب قرار واحد قبل 2 نوفمبر (أدناه).",
    "kpis": ("4", "3", "4"),
    "done": [
        ("إطلاق لوحة مؤشرات الأداء", "نُشرت على بيئة الإنتاج واعتمدتها 3 إدارات."),
        ("تنقية بيانات المستفيدين", "انخفضت السجلات المكررة بنسبة 18٪ (من 9,400 إلى 7,700 سجل)."),
        ("ورشة سياسة حوكمة البيانات", "24 مشاركًا، تقييم 4.6 من 5، واعتُمدت 3 بنود من السياسة."),
        ("تقرير الامتثال الشهري", "سُلِّم قبل الموعد بيومين دون ملاحظات."),
    ],
    "doing": [
        ("أتمتة تقارير الالتزام الأسبوعية", "70", "5 نوفمبر 2026", "على المسار"),
        ("تكامل بيانات الموارد البشرية مع مستودع البيانات", "45", "12 نوفمبر 2026", "متأخر أسبوعًا"),
        ("تحديث دليل تصنيف البيانات", "30", "20 نوفمبر 2026", "على المسار"),
    ],
    "next": [
        "اكتمال اختبار أتمتة التقارير مع الفريق التقني — الثلاثاء.",
        "اعتماد حقول التكامل مع إدارة الموارد البشرية بمحضر — الأربعاء.",
        "تسليم المسودة الأولى لدليل التصنيف للمراجعة — الخميس.",
        "إغلاق طلبات الوصول المعلّقة (6 طلبات) — الأربعاء.",
    ],
    "decision": [
        ("المطلوب:", " اعتماد صلاحية «قراءة» لبيانات المستفيدين لفريق التحليلات (3 موظفين)، طلب رقم ACC-214."),
        ("قبل:", " الأحد 2 نوفمبر، ليبقى موعد 12 نوفمبر قابلًا للتحقيق."),
        ("أثر التأخير:", " يتأخر التكامل أسبوعًا عن كل أسبوع تأخير."),
    ],
    "challenges": "تأخر استلام بيانات الموارد البشرية (تغيّر مسؤول النظام لديهم) — الأثر: تأخر مهمة التكامل أسبوعًا؛ "
                  "الموعد الجديد للتسليم 9 نوفمبر؛ لا يحتاج تدخّلًا.",
    "sign_name": "أحمد العتيبي",
    "sign_date": "29 أكتوبر 2026",
}

# أسبوع ثانٍ تخيلي: يوضح زيادة العداد تلقائيًا وشكل «لا توجد تحديات» الهادئ
EXAMPLE2 = dict(EXAMPLE, **{
    "period": "من 1 إلى 5 نوفمبر 2026 (الأسبوع 45)",
    "summary": "اكتملت أتمتة تقارير الالتزام قبل موعدها واعتُمدت حقول تكامل الموارد البشرية رسميًا؛ 3 مهام منجزة. "
               "تسير المهام الثلاث الجارية ضمن الجدول، ولا يوجد قرار مطلوب هذا الأسبوع.",
    "kpis": ("3", "3", "4"),
    "done": [
        ("أتمتة تقارير الالتزام الأسبوعية", "تعمل آليًا كل أحد؛ وفّرت نحو 6 ساعات عمل أسبوعيًا."),
        ("اعتماد حقول تكامل الموارد البشرية", "اعتُمد 14 حقلًا بمحضر رسمي، وعاد التكامل إلى جدوله."),
        ("إغلاق طلبات الوصول المعلّقة", "أُغلقت 6 من 6 طلبات خلال المدة المستهدفة."),
    ],
    "doing": [
        ("تكامل بيانات الموارد البشرية مع مستودع البيانات", "65", "12 نوفمبر 2026", "على المسار"),
        ("تحديث دليل تصنيف البيانات", "55", "20 نوفمبر 2026", "على المسار"),
        ("إعداد عرض نتائج الربع للإدارة العليا", "40", "18 نوفمبر 2026", "على المسار"),
    ],
    "next": [
        "اكتمال اختبار التكامل الكامل مع الموارد البشرية — الأحد.",
        "تسليم النسخة شبه النهائية لدليل التصنيف — الثلاثاء.",
        "عرض نتائج الربع جاهز لمراجعتك قبل الاعتماد — الخميس.",
        "تدريب الفريق على لوحة المؤشرات (30 دقيقة) — الأربعاء.",
    ],
    "decision": [("", "لا يوجد قرار مطلوب هذا الأسبوع. قرار الأسبوع الماضي (صلاحية القراءة): اعتُمد في 2 نوفمبر.")],
    "challenges": "لا توجد تحديات هذا الأسبوع.",
    "sign_date": "5 نوفمبر 2026",
})

# معرّفات الحقول (idx) — بعيدة عن 10–12 المحجوزة للتاريخ/التذييل/رقم الشريحة
IDX = {"entity": 101, "period": 102, "author": 103, "to": 104, "summary": 105, "kpi_done": 106, "kpi_doing": 107,
       "kpi_next": 108, "done": 109, "next": 110, "decision": 111, "challenges": 112, "sign_name": 113,
       "sign_date": 114, "logo": 120}
PH_NAMES = {101: "اسم الجهة", 102: "الفترة", 103: "إعداد", 104: "مقدَّم إلى", 105: "خلاصة الأسبوع",
            106: "عدد المهام المنجزة", 107: "عدد المهام الجارية", 108: "عدد المهام المقبلة", 109: "المهام المنجزة",
            110: "خطة الأسبوع المقبل", 111: "القرار المطلوب", 112: "التحديات", 113: "اسم المُعِد", 114: "تاريخ التوقيع",
            120: "الشعار"}
CENTERED = {106, 107, 108}


# ----------------------------------------------------------------------------
# أدوات XML
# ----------------------------------------------------------------------------
def E(mm):
    return int(round(mm * MM))


def pt(v):
    return int(round(v * 100))


class Ids:
    def __init__(self, start=1):
        self.n = start

    def next(self):
        self.n += 1
        return self.n


def ids_from_tree(tree):
    return Ids(max([int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))] + [1]))


def BODY():
    return CONFIG.get("body_font") or CONFIG["font"]


def solid(color, alpha=None):
    a = f'<a:alpha val="{int(alpha * 1000)}"/>' if alpha is not None else ""
    return f'<a:solidFill><a:srgbClr val="{color}">{a}</a:srgbClr></a:solidFill>'


def grad(c1, c2, ang=0):
    return (f'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="{c1}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{c2}"/></a:gs></a:gsLst><a:lin ang="{ang * 60000}" scaled="0"/></a:gradFill>')


def rpr(sz, color, bold=False, font=None, alpha=None, lang="ar-SA", tag="a:rPr"):
    font = font or CONFIG["font"]
    b = ' b="1"' if bold else ""
    return (f'<{tag} lang="{lang}" sz="{pt(sz)}"{b} dirty="0">{solid(color, alpha)}'
            f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></{tag}>')


def run(text, sz, color, bold=False, font=None, alpha=None, lang="ar-SA"):
    return f'<a:r>{rpr(sz, color, bold, font, alpha, lang)}<a:t>{escape(text)}</a:t></a:r>'


def ppr(algn="r", rtl=True, ln=None, after=None, before=0, bullet=None, marL=0, indent=0,
        bullet_color=None, tag="a:pPr", extra=""):
    r = ' rtl="1"' if rtl else ' rtl="0"'
    s = f'<{tag} marL="{E(marL)}" indent="{E(indent)}" algn="{algn}"{r}>'
    if ln:
        s += f'<a:lnSpc><a:spcPct val="{int(ln * 100000)}"/></a:lnSpc>'
    if before is not None:
        s += f'<a:spcBef><a:spcPts val="{pt(before)}"/></a:spcBef>'
    if after is not None:
        s += f'<a:spcAft><a:spcPts val="{pt(after)}"/></a:spcAft>'
    if bullet:
        s += f'<a:buClr><a:srgbClr val="{bullet_color or CONFIG["pal"]["teal_dark"]}"/></a:buClr>'
        s += f'<a:buSzPct val="110000"/><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="{bullet}"/>'
    else:
        s += "<a:buNone/>"
    s += extra
    s += f"</{tag}>"
    return s


def para(runs_xml, **kw):
    return f"<a:p>{ppr(**kw)}{runs_xml}</a:p>"


def body_pr(anchor="t", insets=(0, 0, 0, 0), autofit="shrink", wrap=True):
    l, t, r, b = insets
    w = "square" if wrap else "none"
    fit = {"shrink": "<a:normAutofit/>", "resize": "<a:spAutoFit/>", None: ""}[autofit]
    return (f'<a:bodyPr wrap="{w}" lIns="{E(l)}" tIns="{E(t)}" rIns="{E(r)}" bIns="{E(b)}" '
            f'rtlCol="0" anchor="{anchor}">{fit}</a:bodyPr>')


def xfrm(x, y, w, h, flipV=False, flipH=False):
    f = ' flipV="1"' if flipV else ""
    f += ' flipH="1"' if flipH else ""
    return f'<a:xfrm{f}><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm>'


def shape(ids, name, x, y, w, h, fill=None, line=None, geom="rect", adj=None, alpha=None,
          grad_fill=None, line_w=0.75, text_xml="", body=None, txbox=False):
    if grad_fill:
        f = grad_fill
    elif fill:
        f = solid(fill, alpha)
    else:
        f = "<a:noFill/>"
    ln = f'<a:ln w="{pt(line_w) * 127}">{solid(line)}</a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    av = f'<a:gd name="adj" fmla="val {adj}"/>' if adj is not None else ""
    tb = ' txBox="1"' if txbox else ""
    bp = body or body_pr()
    tx = f"<p:txBody>{bp}<a:lstStyle/>{text_xml}</p:txBody>" if text_xml else ""
    return parse_xml(
        f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/><p:cNvSpPr{tb}/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="{geom}"><a:avLst>{av}</a:avLst></a:prstGeom>{f}{ln}</p:spPr>{tx}</p:sp>')


def textbox(ids, name, x, y, w, h, paras, anchor="t", insets=(0, 0, 0, 0), autofit=None, wrap=True, fill=None):
    return shape(ids, name, x, y, w, h, fill=fill, text_xml="".join(paras),
                 body=body_pr(anchor, insets, autofit, wrap), txbox=True)


def line(ids, name, x1, y1, x2, y2, color, w_pt=0.75, alpha=None):
    x, y = min(x1, x2), min(y1, y2)
    w, h = abs(x2 - x1), abs(y2 - y1)
    flipV = (x2 - x1) * (y2 - y1) < 0
    return parse_xml(
        f'<p:cxnSp {NS}><p:nvCxnSpPr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h, flipV=flipV)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        f'<a:ln w="{pt(w_pt) * 127}" cap="rnd">{solid(color, alpha)}<a:round/></a:ln></p:spPr></p:cxnSp>')


def lst_style(sz, color, bold=False, algn="r", bullet=None, ln=1.0, after=2, bullet_color=None, font=None):
    """lvl1–lvl5 كلها معرّفة حتى لا يرث المستخدم أحجام/اتجاه الشريحة الرئيسية عند Tab."""
    font = font or CONFIG["font"]
    out = ""
    for lvl in range(1, 6):
        if bullet:
            marL = 4.5 + (lvl - 1) * 5
            ch = bullet if lvl == 1 else "–"
            bc = bullet_color if lvl == 1 else CONFIG["pal"]["muted"]
        else:
            marL, ch, bc = (lvl - 1) * 5, None, None
        c = color if lvl == 1 else CONFIG["pal"]["muted"]
        s = sz if lvl == 1 else max(sz - 1, 8)
        out += ppr(algn=algn, ln=ln, after=after, bullet=ch, marL=marL, indent=-4.5 if ch else 0,
                   bullet_color=bc, tag=f"a:lvl{lvl}pPr", extra=rpr(s, c, bold and lvl == 1, font, tag="a:defRPr"))
    return f"<a:lstStyle>{out}</a:lstStyle>"


def placeholder(ids, name, idx, x, y, w, h, prompt, sz, color, bold=False, algn="r", bullet=None,
                ln=1.0, after=2, anchor="t", insets=(0, 0, 0, 0), autofit="shrink", bullet_color=None, font=None):
    """Body placeholder on a LAYOUT with custom prompt + inherited list style (RTL)."""
    prompt_xml = "".join(
        f'<a:p><a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(p)}</a:t></a:r></a:p>' for p in prompt.split("\n"))
    return parse_xml(
        f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr><p:ph type="body" sz="quarter" idx="{idx}" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody>{body_pr(anchor, insets, autofit)}'
        f'{lst_style(sz, color, bold, algn, bullet, ln, after, bullet_color, font)}{prompt_xml}</p:txBody></p:sp>')


def pic_placeholder(ids, name, idx, x, y, w, h, prompt):
    return parse_xml(
        f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr><p:ph type="pic" sz="quarter" idx="{idx}" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle><a:lvl1pPr algn="ctr" rtl="1">'
        f'<a:buNone/>{rpr(8, "FFFFFF", False, None, alpha=70, tag="a:defRPr")}</a:lvl1pPr></a:lstStyle>'
        f'<a:p><a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(prompt)}</a:t></a:r></a:p></p:txBody></p:sp>')


def table_cell(text, sz, color, bold=False, fill=None, algn="r", bottom=None, bottom_w=0.75, anchor="ctr",
               insets=(2.2, 1.0, 2.2, 1.0), font=None):
    font = font or CONFIG["font"]
    f = solid(fill) if fill else "<a:noFill/>"
    lnB = (f'<a:lnB w="{pt(bottom_w) * 127}" cap="flat" cmpd="sng" algn="ctr">{solid(bottom)}<a:prstDash val="solid"/></a:lnB>'
           if bottom else "<a:lnB><a:noFill/></a:lnB>")
    l, t, r, b = insets
    if text:
        p = para(run(text, sz, color, bold, font), algn=algn, ln=1.0, after=0)
    else:
        p = (f'<a:p>{ppr(algn=algn, ln=1.0, after=0)}<a:endParaRPr lang="ar-SA" sz="{pt(sz)}" dirty="0">{solid(color)}'
             f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:endParaRPr></a:p>')
    return (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{p}</a:txBody>'
            f'<a:tcPr marL="{E(l)}" marR="{E(r)}" marT="{E(t)}" marB="{E(b)}" anchor="{anchor}">'
            f'<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT>{lnB}{f}</a:tcPr></a:tc>')


STATUS_LATE = ("تأخر", "متأخر", "متوقف", "خطر")


def table(ids, name, x, y, w, col_mm, header, rows, hdr_h, row_h, sz_hdr, sz_body):
    """rows: list of (task, pct, due, status). الترتيب المنطقي: المهمة أولًا (أقصى اليمين)."""
    pal = CONFIG["pal"]
    aligns = ["r", "ctr", "ctr", "ctr"]
    task_i, status_i = 0, 3
    if not TABLE_RTL:  # ترتيب فعلي من اليسار إلى اليمين لبرامج تتجاهل rtl
        col_mm, header, rows = list(reversed(col_mm)), list(reversed(header)), [list(reversed(r)) for r in rows]
        aligns = list(reversed(aligns))
        task_i, status_i = 3, 0
    grid = "".join(f'<a:gridCol w="{E(c)}"/>' for c in col_mm)
    hdr = "".join(table_cell(t, sz_hdr, pal["navy"], True, fill=pal["tint_navy"], algn=a, bottom=pal["teal_dark"],
                             bottom_w=1.0) for t, a in zip(header, aligns))
    body = ""
    for r in rows:
        cells = ""
        for i, (t, a) in enumerate(zip(r, aligns)):
            late = i == status_i and any(k in t for k in STATUS_LATE)
            color = pal["ink"] if i == task_i else (pal["amber"] if late else pal["muted"])
            cells += table_cell(t, sz_body, color, bold=late, algn=a, bottom=pal["rule"], font=BODY())
        body += f'<a:tr h="{E(row_h)}">{cells}</a:tr>'
    h = hdr_h + row_h * len(rows)
    rtl = ' rtl="1"' if TABLE_RTL else ""
    style = "<a:tableStyleId>{2D5ABB26-0587-4C30-8999-92F81FD0307C}</a:tableStyleId>"  # No Style, No Grid
    return parse_xml(
        f'<p:graphicFrame {NS}><p:nvGraphicFramePr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/>'
        f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
        f'<p:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></p:xfrm>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        f'<a:tbl><a:tblPr firstRow="1"{rtl}>{style}</a:tblPr><a:tblGrid>{grid}</a:tblGrid>'
        f'<a:tr h="{E(hdr_h)}">{hdr}</a:tr>{body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')


def counter_box(ids, x, y, w, h):
    """«تقرير رقم» فوق رقم كبير؛ الرقم حقل رقم الشريحة فيزداد تلقائيًا عند تكرار الشريحة."""
    pal = CONFIG["pal"]
    font = CONFIG["font"]
    p1 = para(run(TXT["counter_label"], 9, pal["navy"], True), algn="ctr", ln=1.0, after=0)
    p2 = (f'<a:p>{ppr(algn="ctr", ln=0.9, after=0)}'
          f'<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
          f'{rpr(18, pal["navy"], True, font, lang="en-US")}<a:t>1</a:t></a:fld></a:p>')
    return shape(ids, "عداد التقرير", x, y, w, h, grad_fill=grad(pal["mint"], pal["lime"], 0),
                 geom="roundRect", adj=22000, text_xml=p1 + p2, body=body_pr("ctr", (1, 0.5, 1, 0.5), None), txbox=True)


# ----------------------------------------------------------------------------
# الترويسة والزخرفة (على التخطيط)
# ----------------------------------------------------------------------------
MOTIF_NODES = [(45.5, 17.5), (51, 11), (57, 14.5), (62.5, 6.5), (68.5, 10), (74, 2.5), (79.5, 7), (84.5, 0.8)]
MOTIF_BRANCH = [((57, 14.5), (60.5, 21.5)), ((68.5, 10), (71.5, 16.5))]


def network_motif(ids, color, alpha_line=30, alpha_node=55):
    """قوس متصل من العقد يصعد من خلف الشارة نحو العنوان (بلا تقاطعات)."""
    pal = CONFIG["pal"]
    els = []
    for a, b in zip(MOTIF_NODES, MOTIF_NODES[1:]):
        els.append(line(ids, "خط زخرفي", a[0], a[1], b[0], b[1], color, 0.5, alpha_line))
    for a, b in MOTIF_BRANCH:
        els.append(line(ids, "خط زخرفي", a[0], a[1], b[0], b[1], color, 0.5, alpha_line))
    nodes = MOTIF_NODES + [b for _, b in MOTIF_BRANCH]
    for i, (px, py) in enumerate(nodes):
        r = 1.6 if i % 3 == 0 else 1.0
        c = color
        if i == 3:
            c = pal["mint"]
        elif i == 6:
            c = pal["purple"]
        els.append(shape(ids, "عقدة زخرفية", px - r, py - r, 2 * r, 2 * r, fill=c, alpha=alpha_node, geom="ellipse"))
    return els


def section_header(ids, title, y, color, w=130):
    pal = CONFIG["pal"]
    d = 4.0
    return [
        shape(ids, "أيقونة قسم", MR - d, y + 1.6, d, d, fill=color, geom="ellipse"),
        shape(ids, "أيقونة قسم داخلية", MR - d + 1.25, y + 2.85, 1.5, 1.5, fill="FFFFFF", geom="ellipse"),
        textbox(ids, f"عنوان قسم: {title}", MR - d - 2.5 - w, y - 0.4, w, 8,
                [para(run(title, 14, pal["navy"], True), algn="r", ln=1.0, after=0)], anchor="ctr"),
    ]


# الميزانية الرأسية (مم) — محسوبة على مقاييس PowerPoint (ارتفاع سطر IBM Plex Arabic ≈ 1.73 من حجم الخط)
Y = dict(header_h=46, summary=52, done_hdr=90, done_list=97.5, done_h=38, doing_hdr=141, table=148.5,
         next_hdr=198, next_list=205.5, next_h=38, boxes=247, box_h=30, footer_rule=281.5, footer=283.5)
TABLE_COLS = [90, 24, 34, 38]  # المهمة، النسبة، الموعد، الحالة
TABLE_HDR_H, TABLE_ROW_H = 7.5, 9.0


def build_layout(layout, cfg):
    pal = cfg["pal"]
    tree = layout.shapes._spTree
    # احذف الحقول الموروثة من تخطيط «Blank» (تاريخ/تذييل/رقم شريحة) حتى لا تتعارض معرّفاتها
    for sp in list(tree):
        if sp.tag in (qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"), qn("p:cxnSp")):
            tree.remove(sp)
    ids = ids_from_tree(tree)
    els = []

    # --- الترويسة ---
    H = Y["header_h"]
    els.append(shape(ids, "لوحة الترويسة", 0, 0, PAGE_W, H, grad_fill=grad(pal["navy"], pal["navy2"], 20)))
    els += network_motif(ids, pal["sky"])
    els.append(counter_box(ids, ML, 6, 30, 14))
    els.append(pic_placeholder(ids, "الشعار", IDX["logo"], MR - 14, 5, 14, 14, TXT["logo_prompt"]))
    els.append(placeholder(ids, "اسم الجهة", IDX["entity"], 70, 7, 111, 8, TXT["entity_prompt"], 10, "FFFFFF",
                           algn="r", ln=1.0, after=0, anchor="ctr", font=BODY()))
    els.append(textbox(ids, "عنوان التقرير", 80, 20, MR - 80, 11.5,
                       [para(run(TXT["title"], 22, "FFFFFF", True), algn="r", ln=1.0, after=0)], anchor="ctr"))
    caps = [(TXT["period_cap"], IDX["period"], TXT["period_prompt"]), (TXT["author_cap"], IDX["author"], TXT["author_prompt"]),
            (TXT["to_cap"], IDX["to"], TXT["to_prompt"])]
    col_w = (MR - ML - 2 * 4) / 3
    for i, (cap, idx, prompt) in enumerate(caps):
        xr = MR - i * (col_w + 4)
        x = xr - col_w
        els.append(textbox(ids, f"تسمية: {cap}", x, 32, col_w, 4.5,
                           [para(run(cap, 9, pal["mint"], True), algn="r", ln=1.0, after=0)], anchor="b"))
        els.append(placeholder(ids, cap, idx, x, 36.5, col_w, 7.5, prompt, 10.5, "FFFFFF", algn="r", ln=1.0, after=0,
                               anchor="t", font=BODY()))

    # --- خلاصة الأسبوع + بطاقات الأرقام ---
    y = Y["summary"]
    els.append(textbox(ids, "تسمية الخلاصة", 83, y, MR - 83, 5,
                       [para(run(TXT["summary_cap"], 9.5, pal["teal_dark"], True), algn="r", ln=1.0, after=0)], anchor="b"))
    els.append(placeholder(ids, "خلاصة الأسبوع", IDX["summary"], 83, y + 5.5, MR - 83, 26.5, TXT["summary_prompt"], 12,
                           pal["ink"], ln=1.0, after=0, font=BODY()))
    tiles = [(TXT["kpi_done"], IDX["kpi_done"], pal["green_text"]), (TXT["kpi_doing"], IDX["kpi_doing"], pal["teal_dark"]),
             (TXT["kpi_next"], IDX["kpi_next"], pal["purple"])]
    tw, th, gap = 19, 32, 4
    for i, (label, idx, color) in enumerate(tiles):
        x = 77 - tw - i * (tw + gap)
        els.append(shape(ids, f"بطاقة {label}", x, y, tw, th, fill=pal["tint_navy"], line=pal["edge"], line_w=0.5,
                         geom="roundRect", adj=8000))
        els.append(placeholder(ids, f"عدد {label}", idx, x, y + 2, tw, 19, TXT["kpi_prompt"], 24, color, bold=True,
                               algn="ctr", ln=1.0, after=0, anchor="ctr", autofit="shrink"))
        els.append(textbox(ids, f"تسمية {label}", x - 1, y + 21.5, tw + 2, 9,
                           [para(run(label, 10, pal["navy"], False), algn="ctr", ln=1.0, after=0)], anchor="t"))

    # --- المهام المنجزة ---
    els += section_header(ids, TXT["sec_done"], Y["done_hdr"], pal["green_text"])
    els.append(placeholder(ids, "المهام المنجزة", IDX["done"], ML, Y["done_list"], MR - ML, Y["done_h"], TXT["done_prompt"],
                           11, pal["ink"], bullet="•", bullet_color=pal["green_text"], ln=1.0, after=2, font=BODY()))

    # --- المهام الجارية (الجدول يُضاف على الشريحة) ---
    els += section_header(ids, TXT["sec_doing"], Y["doing_hdr"], pal["teal_dark"])

    # --- خطة الأسبوع المقبل ---
    els += section_header(ids, TXT["sec_next"], Y["next_hdr"], pal["purple"])
    els.append(placeholder(ids, "خطة الأسبوع المقبل", IDX["next"], ML, Y["next_list"], MR - ML, Y["next_h"], TXT["next_prompt"],
                           11, pal["ink"], bullet="•", bullet_color=pal["purple"], ln=1.0, after=2, font=BODY()))

    # --- القرار المطلوب (يمين) + التحديات (يسار، هادئ) ---
    y, bh, gapb, wd = Y["boxes"], Y["box_h"], 4, 95
    wc = MR - ML - wd - gapb
    xd = MR - wd
    els.append(shape(ids, "صندوق القرار", xd, y, wd, bh, fill=pal["tint_mint"], line=pal["teal_dark"], line_w=0.75,
                     geom="roundRect", adj=5000))
    els.append(textbox(ids, "تسمية القرار", xd + 4, y + 2, wd - 8, 6.5,
                       [para(run(TXT["decision_cap"], 12, pal["teal_dark"], True), algn="r", ln=1.0, after=0)], anchor="ctr"))
    els.append(placeholder(ids, "القرار المطلوب", IDX["decision"], xd + 4, y + 9.5, wd - 8, bh - 11.5, TXT["decision_prompt"],
                           11.5, pal["ink"], ln=1.0, after=0, font=BODY()))
    els.append(shape(ids, "صندوق التحديات", ML, y, wc, bh, fill=pal["tint_gray"], line=pal["edge"], line_w=0.5,
                     geom="roundRect", adj=5000))
    els.append(textbox(ids, "تسمية التحديات", ML + 4, y + 2, wc - 8, 6.5,
                       [para(run(TXT["challenges_cap"], 10, pal["muted"], True), algn="r", ln=1.0, after=0)], anchor="ctr"))
    els.append(placeholder(ids, "التحديات", IDX["challenges"], ML + 4, y + 9.5, wc - 8, bh - 11.5, TXT["challenges_prompt"],
                           10.5, pal["quiet"], ln=1.0, after=0, font=BODY()))

    # --- التذييل: توقيع ثابت (يُطبع حتى في النموذج الفارغ) ---
    y = Y["footer_rule"]
    els.append(line(ids, "فاصل التذييل", ML, y, MR, y, pal["rule"], 0.75))
    fy, fh = Y["footer"], 7

    def lab(t, xr, w=13):
        return textbox(ids, f"تسمية {t}", xr - w, fy, w, fh,
                       [para(run(t, 9.5, pal["muted"], True), algn="r", ln=1.0, after=0)], anchor="ctr")

    els.append(lab(TXT["sign_name_cap"], MR))
    els.append(placeholder(ids, "اسم المُعِد", IDX["sign_name"], MR - 13 - 40, fy, 40, fh, TXT["sign_name_prompt"], 10,
                           pal["ink"], ln=1.0, after=0, anchor="ctr", font=BODY()))
    els.append(lab(TXT["sign_cap"], MR - 56, 15))
    els.append(line(ids, "خط التوقيع", MR - 56 - 15 - 34, fy + fh - 1.2, MR - 56 - 15 - 1, fy + fh - 1.2, pal["rule"], 0.75))
    els.append(lab(TXT["sign_date_cap"], MR - 108, 14))
    els.append(placeholder(ids, "تاريخ التوقيع", IDX["sign_date"], MR - 108 - 14 - 30, fy, 30, fh, TXT["sign_date_prompt"], 10,
                           pal["ink"], ln=1.0, after=0, anchor="ctr", font=BODY()))
    els.append(textbox(ids, "نص التذييل الثابت", ML, fy, 40, fh,
                       [para(run(TXT["footer_fixed"], 8.5, pal["faint"]), algn="l", rtl=True, ln=1.0, after=0)], anchor="ctr"))

    for el in els:
        tree.append(el)


# ----------------------------------------------------------------------------
# الشريحة (النموذج) والشريحة المعبأة
# ----------------------------------------------------------------------------
def add_report_slide(prs, layout, cfg, example=None, logo_png=None):
    slide = prs.slides.add_slide(layout)
    layout_ph = {sp.placeholder_format.idx: sp for sp in layout.placeholders}
    for sp in slide.placeholders:
        idx = sp.placeholder_format.idx
        sp.name = PH_NAMES.get(idx, sp.name)
        lp = layout_ph.get(idx)
        tb = sp.element.find(qn("p:txBody"))
        if lp is not None and tb is not None:
            tb.replace(tb.find(qn("a:bodyPr")), copy.deepcopy(lp.element.find(qn("p:txBody")).find(qn("a:bodyPr"))))
            for p in tb.findall(qn("a:p")):
                tb.remove(p)
            tb.append(parse_xml(f'<a:p {NS}><a:pPr rtl="1" algn="{"ctr" if idx in CENTERED else "r"}"/>'
                                f'<a:endParaRPr lang="ar-SA" dirty="0"/></a:p>'))
    ids = ids_from_tree(slide.shapes._spTree)
    header = [TXT["tbl_task"], TXT["tbl_pct"], TXT["tbl_due"], TXT["tbl_status"]]
    rows = list(example["doing"]) if example else [("", "", "", "")] * 4
    slide.shapes._spTree.append(table(ids, "جدول المهام الجارية", ML, Y["table"], MR - ML, TABLE_COLS, header, rows,
                                      hdr_h=TABLE_HDR_H, row_h=TABLE_ROW_H, sz_hdr=10, sz_body=10.5))
    slide.notes_slide.notes_text_frame.text = NOTES_TEXT
    if example:
        fill_example(slide, example, logo_png, cfg)
    return slide


def set_runs(sp, items, bullet=None, bullet_color=None, algn="r"):
    """items: str | (lead, rest). يكتب فقرات RTL ترث نمط التخطيط، مع pPr صريح للرصاصات (مستقل عن البرنامج)."""
    tb = sp.element.find(qn("p:txBody"))
    for p in tb.findall(qn("a:p")):
        tb.remove(p)
    f = CONFIG["font"]
    for it in items:
        if isinstance(it, tuple):
            lead, rest = it
            runs = ""
            if lead:
                runs += (f'<a:r><a:rPr lang="ar-SA" b="1" dirty="0"><a:latin typeface="{f}"/><a:cs typeface="{f}"/></a:rPr>'
                         f'<a:t>{escape(lead)}</a:t></a:r>')
            sep = " — " if lead and not rest.startswith(" ") else ""
            runs += f'<a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(sep + rest)}</a:t></a:r>'
        else:
            runs = f'<a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(it)}</a:t></a:r>'
        if bullet:
            pp = ppr(algn=algn, bullet=bullet, marL=4.5, indent=-4.5, bullet_color=bullet_color, before=None, after=None)
        else:
            pp = f'<a:pPr algn="{algn}" rtl="1"/>'
        tb.append(parse_xml(f'<a:p {NS}>{pp}{runs}</a:p>'))


def fill_example(slide, ex, logo_png, cfg):
    pal = cfg["pal"]
    ph = {sp.placeholder_format.idx: sp for sp in slide.placeholders}
    set_runs(ph[IDX["entity"]], [ex["entity"]])
    set_runs(ph[IDX["period"]], [ex["period"]])
    set_runs(ph[IDX["author"]], [ex["author"]])
    set_runs(ph[IDX["to"]], [ex["to"]])
    set_runs(ph[IDX["summary"]], [ex["summary"]])
    for i, key in enumerate(("kpi_done", "kpi_doing", "kpi_next")):
        set_runs(ph[IDX[key]], [ex["kpis"][i]], algn="ctr")
    set_runs(ph[IDX["done"]], ex["done"], bullet="•", bullet_color=pal["green_text"])
    set_runs(ph[IDX["next"]], ex["next"], bullet="•", bullet_color=pal["purple"])
    set_runs(ph[IDX["decision"]], ex["decision"])
    set_runs(ph[IDX["challenges"]], [ex["challenges"]])
    set_runs(ph[IDX["sign_name"]], [ex["sign_name"]])
    set_runs(ph[IDX["sign_date"]], [ex["sign_date"]])
    if logo_png and IDX["logo"] in ph:
        ph[IDX["logo"]].insert_picture(logo_png)


# ----------------------------------------------------------------------------
# شريحة الدليل (مخفية)
# ----------------------------------------------------------------------------
GUIDE = [
    ("كيف تستخدم القالب؟", None),
    ("الإدخال", "اكتب داخل الحقول ذات النص الإرشادي (يظهر بلون الحقل نفسه). يختفي النص الإرشادي فور الكتابة، وإن تركت الحقل فارغًا فلن يظهر في PDF ولا في الطباعة. استخدم الأرقام 0–9 في كل مكان، واكتب التواريخ باسم الشهر (5 نوفمبر 2026)."),
    ("العداد", "«تقرير رقم N» مرتبط تلقائيًا برقم الشريحة. لأسبوع جديد: حدّد شريحة الأسبوع السابق واضغط Ctrl+D، فتظهر نسخة برقم تالٍ تلقائيًا، وتبقى التقارير السابقة أرشيفًا في الملف نفسه. لا تحذف الشرائح القديمة حتى لا يتغيّر الترقيم؛ وإن حذفتها فاضبط «ترقيم الشرائح من» في: تصميم ← حجم الشريحة ← حجم شريحة مخصص."),
    ("الإخراج", "ملف ← تصدير ← إنشاء PDF/XPS ← خيارات ← النطاق: «الشريحة الحالية». هذه الشريحة مخفية فلا تُصدَّر إلى PDF. للطباعة الورقية: اطبع من ملف PDF، أو من PowerPoint مع التأكد أن «طباعة الشرائح المخفية» غير مفعّلة، واختر «ملاءمة الورقة» على A4 عمودي."),
    ("الجدول", "اكتب نسبة الإنجاز رقمًا فقط (الإشارة ٪ في رأس العمود). Tab ينتقل بين الخلايا من اليمين إلى اليسار، وفي آخر خلية يضيف صفًا. لحذف صف: زر الفأرة الأيمن ← حذف ← الصفوف. الحالة بأربع كلمات ثابتة: على المسار / تأخر بسيط / متأخر / متوقف، ولوّن المتأخر باللون الكهرماني العريض كما في المثال."),
    ("الشعار والجهة", "انقر أيقونة الصورة أعلى اليمين لإدراج شعار جهتك مرة واحدة؛ يُنسخ تلقائيًا مع تكرار الشريحة. إن حُذف حقل أو تحرّك بالخطأ: الصفحة الرئيسية ← إعادة تعيين (Reset)."),
    ("الخط", "القالب مبني على خط IBM Plex Sans Arabic بوزنيه Regular وBold للعناوين وMedium للمتن (مرفقة في مجلد fonts، ترخيص مفتوح). ثبّت الملفات الأربعة بالنقر المزدوج ← «تثبيت»؛ لا يحتاج صلاحيات مسؤول. دون تثبيتها يستبدلها أوفيس بخط قريب."),
    ("نصائح الكتابة", "ابدأ كل مهمة بالنتيجة لا بالنشاط («نُشرت اللوحة واعتمدتها 3 إدارات» بدل «عملت على اللوحة»). خمسة بنود كحد أقصى لكل قسم. بنود الأسبوع المقبل: نتيجة + يوم. الأعداد في الخلاصة تطابق البطاقات وحالات الجدول. التحديات: السبب ← الأثر ← الموعد الجديد ← هل يحتاج تدخّلك؟ القرار المطلوب: طلب واحد بموعد وأثر للتأخير."),
]

NOTES_TEXT = ("دليل سريع: اكتب داخل الحقول الإرشادية (لا تُطبع إن بقيت فارغة). لأسبوع جديد: Ctrl+D على هذه الشريحة فيزيد "
              "العداد تلقائيًا. للتصدير: ملف ← تصدير ← PDF ← خيارات ← الشريحة الحالية. التفاصيل في الشريحة المخفية الأخيرة.")


def add_guide_slide(prs, blank_layout, cfg):
    pal = cfg["pal"]
    slide = prs.slides.add_slide(blank_layout)
    slide._element.set("show", "0")  # مخفية
    ids = Ids(10)
    tree = slide.shapes._spTree
    tree.append(shape(ids, "خلفية الدليل", 0, 0, PAGE_W, 34, fill=pal["navy"]))
    tree.append(textbox(ids, "عنوان الدليل", 12, 8, 186, 10,
                        [para(run("دليل استخدام القالب", 20, "FFFFFF", True), algn="r")], anchor="ctr"))
    tree.append(textbox(ids, "تنبيه", 12, 19, 186, 8,
                        [para(run("هذه الشريحة مخفية: لا تظهر في العرض ولا في ملف PDF (وفي الطباعة تأكد من إلغاء «طباعة الشرائح المخفية»).", 10, pal["mint"]), algn="r")],
                        anchor="ctr"))
    y = 42
    for i, (head, body) in enumerate(GUIDE):
        if body is None:
            tree.append(textbox(ids, "عنوان فرعي", 12, y, 186, 8, [para(run(head, 14, pal["navy"], True), algn="r")], anchor="ctr"))
            y += 10
            continue
        tree.append(shape(ids, "رقم الخطوة", 190, y + 0.8, 7, 7, fill=pal["teal_dark"], geom="ellipse",
                          text_xml=para(run(str(i), 9, "FFFFFF", True, lang="en-US"), algn="ctr", rtl=False),
                          body=body_pr("ctr", (0, 0, 0, 0), None), txbox=True))
        tree.append(textbox(ids, f"دليل: {head}", 12, y, 174, 7, [para(run(head, 12, pal["teal_dark"], True), algn="r")], anchor="ctr"))
        lines_est = max(1, int(len(body) / 92) + 1)
        h = 5.6 * lines_est + 2
        tree.append(textbox(ids, f"نص: {head}", 12, y + 7, 174, h, [para(run(body, 10.5, pal["ink"]), algn="r", ln=1.15)], anchor="t"))
        y += 7 + h + 4
    tree.append(textbox(ids, "تذييل الدليل", 12, 284, 186, 6,
                        [para(run("القالب مستوحى من الهوية البصرية لسدايا (ألوانها وتدرجاتها وزخرفة شبكية) دون استخدام شعارها.", 8.5, pal["faint"]), algn="r")],
                        anchor="ctr"))
    return slide


# ----------------------------------------------------------------------------
# السمة (الخطوط والألوان)، الشريحة الرئيسية، وخصائص الملف
# ----------------------------------------------------------------------------
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def patch_theme(prs, cfg):
    theme_part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    font = cfg["font"]
    for tag in ("majorFont", "minorFont"):
        f = root.find(f".//{A}{tag}")
        f.find(f"{A}latin").set("typeface", font)
        f.find(f"{A}latin").attrib.pop("panose", None)
        f.find(f"{A}cs").set("typeface", font)
        for sf in f.findall(f"{A}font"):
            if sf.get("script") == "Arab":
                sf.set("typeface", font)
    pal = cfg["pal"]
    scheme = root.find(f".//{A}clrScheme")
    scheme.set("name", "SDAIA-inspired")
    mapping = {"dk1": pal["ink"], "lt1": "FFFFFF", "dk2": pal["navy"], "lt2": pal["tint_navy"], "accent1": pal["teal_dark"],
               "accent2": pal["purple"], "accent3": pal["green_text"], "accent4": pal["mint"], "accent5": pal["sky"],
               "accent6": pal["amber"], "hlink": pal["teal_dark"], "folHlink": pal["purple"]}
    for k, v in mapping.items():
        el = scheme.find(f"{A}{k}")
        for c in list(el):
            el.remove(c)
        el.append(etree.fromstring(f'<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="{v}"/>'))
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_master(prs, cfg):
    """أنماط الشريحة الرئيسية: عربي من اليمين بأحجام معقولة، حتى يرث أي نص جديد اتجاهًا وحجمًا سليمين."""
    master = prs.slide_master
    tx = master.element.find(qn("p:txStyles"))
    for style, sz in (("p:titleStyle", 2200), ("p:bodyStyle", 1100), ("p:otherStyle", 1100)):
        st = tx.find(qn(style))
        if st is None:
            continue
        for lvl in st:
            if lvl.tag.endswith("pPr"):
                lvl.set("algn", "r")
                lvl.set("rtl", "1")
                d = lvl.find(qn("a:defRPr"))
                if d is not None:
                    d.set("sz", str(sz))
                    d.set("lang", "ar-SA")
    pres = prs.part._element
    for d in pres.iter(qn("a:defRPr")):
        d.set("lang", "ar-SA")
    pres.find(qn("p:sldSz")).attrib.pop("type", None)  # حجم مخصص A4 عمودي
    now = _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0)
    cp = prs.core_properties
    cp.title = "التقرير الأسبوعي للمهام — قالب"
    cp.author = ""
    cp.last_modified_by = ""
    cp.created = cp.modified = now
    cp.revision = 1
    cp.language = "ar-SA"
    cp.subject = "قالب تقرير أسبوعي للمهام (A4 عمودي، عربي)"
    cp.comments = ""
    # أدلة العرض الموروثة من قالب 4:3
    for part in prs.part.package.iter_parts():
        if str(part.partname) == "/ppt/viewProps.xml":
            root = etree.fromstring(part.blob)
            for g in list(root.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}guideLst")):
                g.getparent().remove(g)
            part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def set_thumbnail(pptx_path, jpeg_path):
    """يستبدل المعاينة المصغرة الافتراضية (بيضاء) بصورة التقرير."""
    import zipfile
    tmp = pptx_path + ".tmp"
    with zipfile.ZipFile(pptx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/thumbnail.jpeg":
                data = open(jpeg_path, "rb").read()
            zout.writestr(item, data)
    os.replace(tmp, pptx_path)


def prune_layouts(prs, keep):
    master = prs.slide_master
    for lay in list(master.slide_layouts):
        if lay not in keep:
            master.slide_layouts.remove(lay)


def make_logo_png(path):
    """شعار تجريبي مجرد للمثال (ليس لأي جهة حقيقية)."""
    from PIL import Image, ImageDraw
    s = 512
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((16, 16, s - 16, s - 16), radius=90, fill=(255, 255, 255, 255))
    d.ellipse((120, 120, 392, 392), outline=(12, 115, 115, 255), width=34)
    d.ellipse((205, 205, 307, 307), fill=(51, 204, 153, 255))
    img.save(path)
    return path


def build(out, example=False, font=DEFAULT_FONT, palette=None, logo_png=None, body_font=None, weeks=1, thumbnail=None):
    global CONFIG
    CONFIG = {"font": font, "body_font": body_font or DEFAULT_BODY_FONT, "pal": dict(PALETTE, **(palette or {}))}
    prs = Presentation()
    prs.slide_width = Emu(E(PAGE_W))
    prs.slide_height = Emu(E(PAGE_H))
    master = prs.slide_master
    layouts = {l.name: l for l in master.slide_layouts}
    report_layout = layouts["Blank"]
    blank_layout = layouts["Title Only"]
    for sp in list(blank_layout.shapes._spTree):
        if sp.tag in (qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"), qn("p:cxnSp")):
            blank_layout.shapes._spTree.remove(sp)
    report_layout.name = TXT["layout_name"]
    blank_layout.name = TXT["blank_layout_name"]
    prune_layouts(prs, [report_layout, blank_layout])
    patch_theme(prs, CONFIG)
    patch_master(prs, CONFIG)
    build_layout(report_layout, CONFIG)
    if example and not logo_png:
        logo_png = make_logo_png(os.path.join(tempfile.gettempdir(), "sample_logo.png"))
    if example:
        for ex in [EXAMPLE, EXAMPLE2][:max(1, weeks)]:
            add_report_slide(prs, report_layout, CONFIG, ex, logo_png)
    else:
        add_report_slide(prs, report_layout, CONFIG, None, logo_png)
    add_guide_slide(prs, blank_layout, CONFIG)
    prs.save(out)
    if thumbnail and os.path.exists(thumbnail):
        set_thumbnail(out, thumbnail)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--example", action="store_true", help="نسخة معبأة ببيانات تجريبية")
    ap.add_argument("--weeks", type=int, default=1, help="عدد الأسابيع في المثال (1 أو 2)")
    ap.add_argument("--font", default=DEFAULT_FONT)
    ap.add_argument("--body-font", default=None, help="خط المتن (افتراضيًا IBM Plex Sans Arabic Medium)")
    ap.add_argument("--logo", default=None)
    ap.add_argument("--thumbnail", default=None, help="صورة JPEG تُستخدم معاينةً مصغرة للملف")
    ap.add_argument("--table-ltr", action="store_true", help="جدول بترتيب فعلي من اليسار (لبرامج تتجاهل rtl مثل LibreOffice)")
    a = ap.parse_args()
    if a.table_ltr:
        TABLE_RTL = False
    print(build(a.out, a.example, a.font, logo_png=a.logo, body_font=a.body_font, weeks=a.weeks, thumbnail=a.thumbnail))
