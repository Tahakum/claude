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
  python3 build_template.py --out "التقرير_الأسبوعي_قالب.pptx"
  python3 build_template.py --out example.pptx --example          # نسخة معبأة ببيانات تجريبية
  python3 build_template.py --out x.pptx --font "Tajawal"         # تجربة خط آخر
"""
import argparse
import copy
import io
import re
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
NS = nsdecls("a", "p", "r")

DEFAULT_FONT = "IBM Plex Sans Arabic"

# لوحة الألوان (مستوحاة من هوية سدايا: كحلي داكن + تدرج تركوازي/أخضر) — تُضبط من CONFIG
PALETTE = {
    "navy": "0B1F3F",        # لوحة الترويسة
    "navy2": "14315E",       # تدرج ثانٍ للترويسة
    "ink": "1C2B3A",         # نص أساسي
    "muted": "5F6B7A",       # نص ثانوي
    "faint": "8A95A3",       # نص خافت جدًا (التذييل)
    "rule": "D9DEE5",        # خطوط الجدول
    "teal": "12A5A8",        # اللون المميز A
    "teal_dark": "0E7C86",   # عناوين على خلفية فاتحة
    "green": "6CC04A",       # اللون المميز B (منجز)
    "amber": "D98B1F",       # متأخر/خطر (للمستخدم عند الحاجة)
    "tint_navy": "EEF2F8",   # خلفية بطاقات الأرقام
    "tint_teal": "E6F5F5",   # صندوق القرار المطلوب
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
    "to_cap": "مُقدَّم إلى",
    "period_prompt": "من … إلى … (الأسبوع …)",
    "author_prompt": "الاسم — المسمى الوظيفي",
    "to_prompt": "اسم المدير — المسمى",
    "summary_cap": "خلاصة الأسبوع",
    "summary_prompt": "لخّص الأسبوع في ثلاثة أسطر: أبرز ما أُنجز، وما يسير، وما يحتاج تدخّلًا.",
    "kpi_done": "مهام منجزة",
    "kpi_doing": "مهام جارية",
    "kpi_next": "مهام قادمة",
    "kpi_prompt": "0",
    "sec_done": "المهام المنجزة هذا الأسبوع",
    "done_prompt": "اكتب كل مهمة منجزة في سطر: اسم المهمة — النتيجة أو الأثر (مع رقم إن أمكن).",
    "sec_doing": "المهام الجارية",
    "tbl_task": "المهمة",
    "tbl_pct": "نسبة الإنجاز",
    "tbl_due": "الموعد المستهدف",
    "tbl_status": "الحالة",
    "sec_next": "خطة الأسبوع القادم",
    "next_prompt": "اكتب المهام المخطط لها الأسبوع القادم، مهمة في كل سطر، مع الموعد إن وُجد.",
    "decision_cap": "قرار أو دعم مطلوب من المدير",
    "decision_prompt": "اكتب طلبًا واضحًا واحدًا: ماذا تحتاج، ولماذا، وقبل أي تاريخ. (اتركه فارغًا إن لم يوجد)",
    "challenges_cap": "التحديات",
    "challenges_prompt": "لا توجد تحديات هذا الأسبوع — أو اكتب التحدي في جملة وما فعلته حياله.",
    "sign_prompt": "إعداد: الاسم        التوقيع: ______________        التاريخ: __ / __ / ____",
    "footer_fixed": "تقرير داخلي أسبوعي — يُقدَّم إلى المدير المباشر",
    "layout_name": "التقرير الأسبوعي",
    "blank_layout_name": "صفحة فارغة",
}

EXAMPLE = {
    "entity": "الإدارة العامة للتحول الرقمي — إدارة البيانات والتحليلات",
    "period": "من 25 إلى 29 أكتوبر 2026 (الأسبوع 44)",
    "author": "أحمد العتيبي — محلل بيانات أول",
    "to": "م. خالد القحطاني — مدير إدارة البيانات",
    "summary": "أُنجزت 4 مهام رئيسية هذا الأسبوع أبرزها إطلاق لوحة مؤشرات الأداء للإدارة، وتسير 3 مهام ضمن جدولها الزمني، "
               "ومهمة واحدة تحتاج قرارًا بشأن صلاحيات الوصول إلى بيانات المستفيدين قبل 12 نوفمبر.",
    "kpis": ("4", "3", "5"),
    "done": [
        ("إطلاق لوحة مؤشرات الأداء", "نُشرت على بيئة الإنتاج واعتمدتها 3 إدارات."),
        ("تنقية بيانات المستفيدين", "انخفضت السجلات المكررة بنسبة 18٪ (من 9,400 إلى 7,700 سجل)."),
        ("ورشة سياسة حوكمة البيانات", "24 مشاركًا، متوسط التقييم 4.6 من 5."),
        ("تقرير الامتثال الشهري", "سُلِّم قبل الموعد بيومين دون ملاحظات."),
    ],
    "doing": [
        ("أتمتة تقارير الالتزام الأسبوعية", "70٪", "5 نوفمبر", "على المسار"),
        ("تكامل بيانات الموارد البشرية مع مستودع البيانات", "45٪", "12 نوفمبر", "تأخر بسيط"),
        ("تحديث دليل تصنيف البيانات", "30٪", "20 نوفمبر", "على المسار"),
    ],
    "next": [
        "إكمال اختبار أتمتة التقارير مع الفريق التقني (الثلاثاء).",
        "اجتماع مع إدارة الموارد البشرية لاعتماد حقول التكامل.",
        "تسليم المسودة الأولى لدليل التصنيف للمراجعة.",
        "إغلاق طلبات الوصول المعلّقة (6 طلبات).",
        "إعداد عرض نتائج الربع للإدارة العليا.",
    ],
    "decision": "اعتماد منح صلاحية «قراءة» لبيانات المستفيدين لفريق التحليلات (3 موظفين) لإكمال التكامل قبل 12 نوفمبر.",
    "challenges": "تأخر استلام بيانات الموارد البشرية بسبب تغيّر مسؤول النظام لديهم؛ تمت المتابعة وحُدّد موعد جديد للتسليم.",
    "sign": "إعداد: أحمد العتيبي        التوقيع: ______________        التاريخ: 29 / 10 / 2026",
}


# ----------------------------------------------------------------------------
# أدوات XML
# ----------------------------------------------------------------------------
def E(mm):
    return int(round(mm * MM))


def pt(v):
    return int(round(v * 100))


class Ids:
    def __init__(self, start=2):
        self.n = start

    def next(self):
        self.n += 1
        return self.n


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


def ppr(algn="r", rtl=True, ln=None, after=None, before=None, bullet=None, marL=0, indent=0,
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
        s += f'<a:buClr><a:srgbClr val="{bullet_color or CONFIG["pal"]["teal"]}"/></a:buClr>'
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
          grad_fill=None, line_w=0.75, text_xml="", body=None, txbox=False, shadow=None):
    f = ""
    if grad_fill:
        f = grad_fill
    elif fill:
        f = solid(fill, alpha)
    else:
        f = "<a:noFill/>"
    ln = f'<a:ln w="{pt(line_w) * 127}">{solid(line)}</a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    av = f'<a:gd name="adj" fmla="val {adj}"/>' if adj is not None else ""
    eff = ""
    if shadow:
        eff = ('<a:effectLst><a:outerShdw blurRad="76200" dist="25400" dir="5400000" algn="t" rotWithShape="0">'
               f'<a:srgbClr val="{shadow[0]}"><a:alpha val="{int(shadow[1] * 1000)}"/></a:srgbClr></a:outerShdw></a:effectLst>')
    tb = ' txBox="1"' if txbox else ""
    bp = body or body_pr()
    tx = f"<p:txBody>{bp}<a:lstStyle/>{text_xml}</p:txBody>" if text_xml else ""
    return parse_xml(
        f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/><p:cNvSpPr{tb}/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="{geom}"><a:avLst>{av}</a:avLst></a:prstGeom>{f}{ln}{eff}</p:spPr>{tx}</p:sp>')


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


def placeholder(ids, name, idx, x, y, w, h, prompt, sz, color, bold=False, algn="r", bullet=None,
                ln=1.3, after=2, anchor="t", insets=(0, 0, 0, 0), autofit="shrink", lvl2=True,
                bullet_color=None, font=None):
    """Body placeholder on a LAYOUT with custom prompt + inherited list style (RTL)."""
    font = font or CONFIG["font"]
    marL = 4.5 if bullet else 0
    indent = -4.5 if bullet else 0
    lvl1 = ppr(algn=algn, ln=ln, after=after, bullet=bullet, marL=marL, indent=indent,
               bullet_color=bullet_color, tag="a:lvl1pPr",
               extra=rpr(sz, color, bold, font, tag="a:defRPr"))
    lvl2x = ""
    if lvl2:
        lvl2x = ppr(algn=algn, ln=ln, after=after, bullet="–" if bullet else None, marL=marL + 5, indent=indent,
                    bullet_color=CONFIG["pal"]["muted"], tag="a:lvl2pPr",
                    extra=rpr(max(sz - 1, 8), CONFIG["pal"]["muted"], False, font, tag="a:defRPr"))
    prompt_xml = "".join(
        f'<a:p><a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(p)}</a:t></a:r></a:p>' for p in prompt.split("\n"))
    return parse_xml(
        f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr><p:ph type="body" sz="quarter" idx="{idx}" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody>{body_pr(anchor, insets, autofit)}<a:lstStyle>{lvl1}{lvl2x}</a:lstStyle>{prompt_xml}</p:txBody></p:sp>')


def pic_placeholder(ids, name, idx, x, y, w, h, prompt):
    return parse_xml(
        f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr><p:ph type="pic" sz="quarter" idx="{idx}" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle><a:lvl1pPr algn="ctr" rtl="1">'
        f'<a:buNone/>{rpr(8, "FFFFFF", False, None, alpha=70, tag="a:defRPr")}</a:lvl1pPr></a:lstStyle>'
        f'<a:p><a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(prompt)}</a:t></a:r></a:p></p:txBody></p:sp>')


def table_cell(text, sz, color, bold=False, fill=None, algn="r", bottom=None, top=None, anchor="ctr",
               insets=(2.2, 1.2, 2.2, 1.2), font=None):
    f = solid(fill) if fill else "<a:noFill/>"
    lnB = f'<a:lnB w="6350" cap="flat" cmpd="sng" algn="ctr">{solid(bottom)}<a:prstDash val="solid"/></a:lnB>' if bottom else "<a:lnB><a:noFill/></a:lnB>"
    lnT = f'<a:lnT w="6350" cap="flat" cmpd="sng" algn="ctr">{solid(top)}<a:prstDash val="solid"/></a:lnT>' if top else "<a:lnT><a:noFill/></a:lnT>"
    l, t, r, b = insets
    if text:
        p = para(run(text, sz, color, bold, font), algn=algn, ln=1.2)
    else:
        p = f'<a:p>{ppr(algn=algn, ln=1.2)}<a:endParaRPr lang="ar-SA" sz="{pt(sz)}" dirty="0">{solid(color)}<a:latin typeface="{font or CONFIG["font"]}"/><a:cs typeface="{font or CONFIG["font"]}"/></a:endParaRPr></a:p>'
    return (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{p}</a:txBody>'
            f'<a:tcPr marL="{E(l)}" marR="{E(r)}" marT="{E(t)}" marB="{E(b)}" anchor="{anchor}">'
            f'<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR>{lnT}{lnB}{f}</a:tcPr></a:tc>')


def table(ids, name, x, y, w, col_mm, header, rows, hdr_h, row_h, sz_hdr, sz_body):
    pal = CONFIG["pal"]
    # الأعمدة مرتبة فعليًا من اليسار إلى اليمين (آخر عمود = «المهمة» في أقصى اليمين) بدل الاعتماد على rtl في tblPr
    # لأن بعض البرامج (LibreOffice مثلًا) تتجاهله؛ النص داخل الخلايا يبقى RTL.
    col_mm = list(reversed(col_mm)); header = list(reversed(header)); rows = [list(reversed(r)) for r in rows]
    grid = "".join(f'<a:gridCol w="{E(c)}"/>' for c in col_mm)
    aligns = ["ctr", "ctr", "ctr", "r"]
    hdr = "".join(table_cell(t, sz_hdr, pal["navy"], True, fill=pal["tint_navy"], algn=a, bottom=pal["teal"])
                  for t, a in zip(header, aligns))
    body = ""
    for r in rows:
        body += f'<a:tr h="{E(row_h)}">' + "".join(
            table_cell(t, sz_body, pal["ink"] if i == 3 else pal["muted"], False, algn=a, bottom=pal["rule"])
            for i, (t, a) in enumerate(zip(r, aligns))) + "</a:tr>"
    h = hdr_h + row_h * len(rows)
    return parse_xml(
        f'<p:graphicFrame {NS}><p:nvGraphicFramePr><p:cNvPr id="{ids.next()}" name="{escape(name)}"/>'
        f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
        f'<p:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></p:xfrm>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        f'<a:tbl><a:tblPr firstRow="1"/><a:tblGrid>{grid}</a:tblGrid>'
        f'<a:tr h="{E(hdr_h)}">{hdr}</a:tr>{body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')


def counter_box(ids, x, y, w, h):
    """«تقرير رقم N» — N is a slide-number field, so duplicating the slide increments it."""
    pal = CONFIG["pal"]
    font = CONFIG["font"]
    p = (f'<a:p>{ppr(algn="ctr", rtl=True)}'
         f'{run(TXT["counter_label"] + "  ", 9.5, "FFFFFF", False, alpha=88)}'
         f'<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
         f'{rpr(22, "FFFFFF", True, font, lang="en-US")}<a:t>1</a:t></a:fld></a:p>')
    return shape(ids, "عداد التقرير", x, y, w, h, grad_fill=grad(pal["teal"], pal["green"], 0),
                 geom="roundRect", adj=30000, text_xml=p, body=body_pr("ctr", (1, 0, 1, 0), None), txbox=True)


# ----------------------------------------------------------------------------
# الترويسة والزخرفة (على التخطيط)
# ----------------------------------------------------------------------------
def network_motif(ids, x0, y0, w, h, color, alpha_line=28, alpha_node=55, seed=7):
    """Subtle node/line network (SDAIA-like data motif), deterministic positions."""
    import random
    rnd = random.Random(seed)
    pts = []
    for i in range(9):
        pts.append((x0 + rnd.uniform(0.05, 0.95) * w, y0 + rnd.uniform(0.08, 0.92) * h))
    els = []
    # connect each node to its 2 nearest neighbours
    for i, (px, py) in enumerate(pts):
        d = sorted(((abs(px - qx) ** 2 + abs(py - qy) ** 2, j) for j, (qx, qy) in enumerate(pts) if j != i))
        for _, j in d[:2]:
            if j > i:
                qx, qy = pts[j]
                els.append(line(ids, "خط زخرفي", px, py, qx, qy, color, 0.5, alpha_line))
    for i, (px, py) in enumerate(pts):
        r = 0.9 if i % 3 else 1.5
        els.append(shape(ids, "عقدة زخرفية", px - r, py - r, 2 * r, 2 * r, fill=color, alpha=alpha_node, geom="ellipse"))
    return els


def section_header(ids, title, x_right, y, color, w=120):
    """Node icon + title; x_right is the right edge (RTL start)."""
    pal = CONFIG["pal"]
    els = []
    d = 3.6
    els.append(shape(ids, "أيقونة قسم", x_right - d, y + 1.3, d, d, fill=color, geom="ellipse"))
    els.append(shape(ids, "أيقونة قسم داخلية", x_right - d + 1.1, y + 2.4, 1.4, 1.4, fill="FFFFFF", geom="ellipse"))
    els.append(textbox(ids, f"عنوان قسم: {title}", x_right - d - 2 - w, y - 0.2, w, 6.4,
                       [para(run(title, 12.5, pal["navy"], True), algn="r")], anchor="ctr"))
    return els


def build_layout(layout, cfg):
    pal = cfg["pal"]
    ids = Ids(1)
    tree = layout.shapes._spTree
    els = []
    ML, MR = 12, 198  # content left/right edges (mm)

    # --- الترويسة (لوحة كحلية ممتدة) ---
    HEADER_H = 46
    els.append(shape(ids, "لوحة الترويسة", 0, 0, PAGE_W, HEADER_H, grad_fill=grad(pal["navy"], pal["navy2"], 20)))
    # زخرفة شبكية خفيفة (يسار الترويسة)
    els += network_motif(ids, 48, 3, 46, 24, pal["teal"], alpha_line=35, alpha_node=60)
    # الشعار (يمين أعلى)
    els.append(pic_placeholder(ids, "الشعار", 40, MR - 22, 8, 22, 22, TXT["logo_prompt"]))
    # اسم الجهة + العنوان
    els.append(placeholder(ids, "اسم الجهة", 11, 96, 8.5, 76, 7, TXT["entity_prompt"], 9.5, "FFFFFF",
                           algn="r", ln=1.1, after=0, anchor="t", autofit="shrink", lvl2=False))
    els.append(textbox(ids, "عنوان التقرير", 96, 15.5, 76, 11,
                       [para(run(TXT["title"], 20, "FFFFFF", True), algn="r", ln=1.0)], anchor="ctr"))
    # العداد (يسار أعلى)
    els.append(counter_box(ids, ML, 9, 34, 15))
    # صف البيانات الوصفية: الفترة | إعداد | مُقدَّم إلى
    caps = [(TXT["period_cap"], 12, TXT["period_prompt"]), (TXT["author_cap"], 13, TXT["author_prompt"]),
            (TXT["to_cap"], 14, TXT["to_prompt"])]
    col_w = (MR - ML - 2 * 4) / 3  # 59.33
    for i, (cap, idx, prompt) in enumerate(caps):
        xr = MR - i * (col_w + 4)
        x = xr - col_w
        els.append(textbox(ids, f"تسمية: {cap}", x, 30.5, col_w, 4,
                           [para(run(cap, 7.5, pal["teal"], True), algn="r")], anchor="b"))
        els.append(placeholder(ids, cap, idx, x, 35, col_w, 7.5, prompt, 10, "FFFFFF", algn="r", ln=1.1, after=0,
                               anchor="t", lvl2=False))

    # --- خلاصة الأسبوع + بطاقات الأرقام ---
    y = HEADER_H + 7  # 53
    els.append(textbox(ids, "تسمية الخلاصة", 80, y, MR - 80, 5,
                       [para(run(TXT["summary_cap"], 9, pal["teal_dark"], True), algn="r")], anchor="b"))
    els.append(placeholder(ids, "خلاصة الأسبوع", 20, 80, y + 6, MR - 80, 24, TXT["summary_prompt"], 12.5, pal["ink"],
                           ln=1.32, after=0, lvl2=False))
    # بطاقات الأرقام (ثلاث) — من اليمين: منجزة، جارية، قادمة
    tiles = [(TXT["kpi_done"], 21, pal["green"]), (TXT["kpi_doing"], 22, pal["teal"]), (TXT["kpi_next"], 23, pal["navy"])]
    tw, th, gap = 20, 27, 2.5
    x_right_tiles = 80 - 6  # 74
    for i, (label, idx, color) in enumerate(tiles):
        xr = x_right_tiles - i * (tw + gap)
        x = xr - tw
        els.append(shape(ids, f"بطاقة {label}", x, y + 3, tw, th, fill=pal["tint_navy"], geom="roundRect", adj=9000))
        els.append(placeholder(ids, f"عدد {label}", idx, x, y + 4.5, tw, 14, TXT["kpi_prompt"], 22, color, bold=True,
                               algn="ctr", ln=1.0, after=0, anchor="ctr", autofit=None, lvl2=False, font=cfg["font"]))
        els.append(textbox(ids, f"تسمية {label}", x, y + 19.5, tw, 8,
                           [para(run(label, 8, pal["muted"], False), algn="ctr", ln=1.0)], anchor="t"))

    # --- المهام المنجزة ---
    y = 91
    els += section_header(ids, TXT["sec_done"], MR, y, pal["green"])
    els.append(placeholder(ids, "المهام المنجزة", 24, ML, y + 8, MR - ML, 42, TXT["done_prompt"], 12, pal["ink"],
                           bullet="•", bullet_color=pal["green"], ln=1.28, after=2.5))

    # --- المهام الجارية (الجدول يُضاف على الشريحة) ---
    y = 146
    els += section_header(ids, TXT["sec_doing"], MR, y, pal["teal"])

    # --- خطة الأسبوع القادم ---
    y = 199.5
    els += section_header(ids, TXT["sec_next"], MR, y, pal["navy"])
    els.append(placeholder(ids, "خطة الأسبوع القادم", 25, ML, y + 7.5, MR - ML, 33, TXT["next_prompt"], 12, pal["ink"],
                           bullet="•", bullet_color=pal["navy"], ln=1.28, after=2.5))

    # --- القرار المطلوب (يمين) + التحديات (يسار، هادئ) ---
    y = 245
    bh = 29
    gapb = 5
    wd = 96   # صندوق القرار
    wc = MR - ML - wd - gapb  # 85
    xd = MR - wd
    els.append(shape(ids, "صندوق القرار", xd, y, wd, bh, fill=pal["tint_teal"], geom="roundRect", adj=6000))
    els.append(textbox(ids, "تسمية القرار", xd + 4, y + 2.5, wd - 8, 5.5,
                       [para(run(TXT["decision_cap"], 10.5, pal["teal_dark"], True), algn="r")], anchor="ctr"))
    els.append(placeholder(ids, "القرار المطلوب", 26, xd + 4, y + 9, wd - 8, bh - 11, TXT["decision_prompt"], 11,
                           pal["ink"], ln=1.25, after=0, lvl2=False))
    els.append(shape(ids, "صندوق التحديات", ML, y, wc, bh, fill=pal["tint_gray"], geom="roundRect", adj=6000))
    els.append(textbox(ids, "تسمية التحديات", ML + 4, y + 2.5, wc - 8, 5.5,
                       [para(run(TXT["challenges_cap"], 9.5, pal["muted"], True), algn="r")], anchor="ctr"))
    els.append(placeholder(ids, "التحديات", 27, ML + 4, y + 9, wc - 8, bh - 11, TXT["challenges_prompt"], 10,
                           pal["muted"], ln=1.25, after=0, lvl2=False))

    # --- التذييل والتوقيع ---
    y = 280.5
    els.append(line(ids, "فاصل التذييل", ML, y, MR, y, pal["rule"], 0.5))
    els.append(placeholder(ids, "التوقيع", 28, 70, y + 2, MR - 70, 7, TXT["sign_prompt"], 9, pal["muted"],
                           ln=1.1, after=0, anchor="ctr", lvl2=False))
    els.append(textbox(ids, "نص التذييل الثابت", ML, y + 2, 55, 7,
                       [para(run(TXT["footer_fixed"], 7, pal["faint"]), algn="l", rtl=True)], anchor="ctr"))

    for el in els:
        tree.append(el)


# ----------------------------------------------------------------------------
# الشريحة (النموذج) والشريحة المعبأة
# ----------------------------------------------------------------------------
PH_NAMES = {11: "اسم الجهة", 12: "الفترة", 13: "إعداد", 14: "مُقدَّم إلى", 20: "خلاصة الأسبوع", 21: "عدد المهام المنجزة",
            22: "عدد المهام الجارية", 23: "عدد المهام القادمة", 24: "المهام المنجزة", 25: "خطة الأسبوع القادم",
            26: "القرار المطلوب", 27: "التحديات", 28: "التوقيع", 40: "الشعار"}


def add_report_slide(prs, layout, cfg, example=None, logo_png=None):
    pal = cfg["pal"]
    slide = prs.slides.add_slide(layout)
    # أسماء عربية للحقول + bodyPr صريح (نسخ من التخطيط) حتى يعمل التصغير التلقائي على الشريحة أيضًا
    layout_ph = {sp.placeholder_format.idx: sp for sp in layout.placeholders}
    for sp in slide.placeholders:
        idx = sp.placeholder_format.idx
        sp.name = PH_NAMES.get(idx, sp.name)
        lp = layout_ph.get(idx)
        if lp is not None and sp.element.find(qn("p:txBody")) is not None:
            tb = sp.element.find(qn("p:txBody"))
            old = tb.find(qn("a:bodyPr"))
            tb.replace(old, copy.deepcopy(lp.element.find(qn("p:txBody")).find(qn("a:bodyPr"))))
            # فقرة فارغة بلغة عربية (نقطة الإدراج تبدأ عربيًا)
            for p in tb.findall(qn("a:p")):
                tb.remove(p)
            tb.append(parse_xml(f'<a:p {NS}><a:pPr rtl="1" algn="{"ctr" if idx in (21, 22, 23) else "r"}"/><a:endParaRPr lang="ar-SA" dirty="0"/></a:p>'))

    ids = Ids(max(int(e.get("id")) for e in slide.shapes._spTree.iter(qn("p:cNvPr"))) + 1)
    tree = slide.shapes._spTree

    # جدول المهام الجارية
    cols = [94, 26, 36, 30]  # الأول = الأقصى يمينًا (rtl)
    header = [TXT["tbl_task"], TXT["tbl_pct"], TXT["tbl_due"], TXT["tbl_status"]]
    rows = example["doing"] if example else [("", "", "", "")] * 4
    if example and len(rows) < 4:
        rows = list(rows) + [("", "", "", "")] * (4 - len(rows))
    tree.append(table(ids, "جدول المهام الجارية", 12, 154, 186, cols, header, rows, hdr_h=7.5, row_h=8,
                      sz_hdr=9.5, sz_body=11))

    # ملاحظات المتحدث: دليل مختصر
    slide.notes_slide.notes_text_frame.text = NOTES_TEXT

    if example:
        fill_example(slide, example, logo_png, cfg)
    return slide


def set_runs(sp, items, bold_lead=False):
    """items: list of str or (lead, rest). Writes RTL paragraphs inheriting layout style."""
    tb = sp.element.find(qn("p:txBody"))
    for p in tb.findall(qn("a:p")):
        tb.remove(p)
    for it in items:
        if isinstance(it, tuple):
            lead, rest = it
            runs = (f'<a:r><a:rPr lang="ar-SA" b="1" dirty="0"/><a:t>{escape(lead)}</a:t></a:r>'
                    f'<a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(" — " + rest)}</a:t></a:r>')
        else:
            runs = f'<a:r><a:rPr lang="ar-SA" dirty="0"/><a:t>{escape(it)}</a:t></a:r>'
        tb.append(parse_xml(f'<a:p {NS}>{runs}</a:p>'))


def fill_example(slide, ex, logo_png, cfg):
    ph = {sp.placeholder_format.idx: sp for sp in slide.placeholders}
    set_runs(ph[11], [ex["entity"]])
    set_runs(ph[12], [ex["period"]])
    set_runs(ph[13], [ex["author"]])
    set_runs(ph[14], [ex["to"]])
    set_runs(ph[20], [ex["summary"]])
    for i, idx in enumerate((21, 22, 23)):
        set_runs(ph[idx], [ex["kpis"][i]])
        # keep centred alignment on slide paragraphs
        for p in ph[idx].element.find(qn("p:txBody")).findall(qn("a:p")):
            p.insert(0, parse_xml(f'<a:pPr {NS} rtl="1" algn="ctr"/>'))
    set_runs(ph[24], ex["done"])
    set_runs(ph[25], ex["next"])
    set_runs(ph[26], [ex["decision"]])
    set_runs(ph[27], [ex["challenges"]])
    set_runs(ph[28], [ex["sign"]])
    if logo_png and 40 in ph:
        ph[40].insert_picture(logo_png)


# ----------------------------------------------------------------------------
# شريحة الدليل (مخفية)
# ----------------------------------------------------------------------------
GUIDE = [
    ("كيف تستخدم القالب؟", None),
    ("الإدخال", "اكتب مباشرة داخل الحقول ذات النص الإرشادي الرمادي. يختفي النص الإرشادي فور الكتابة، وإن تركت الحقل فارغًا فلن يظهر في الطباعة ولا في PDF."),
    ("العداد", "«تقرير رقم N» مرتبط تلقائيًا برقم الشريحة. لأسبوع جديد: حدّد شريحة الأسبوع السابق واضغط Ctrl+D، فتظهر نسخة برقم تالٍ تلقائيًا، وتبقى التقارير السابقة أرشيفًا في الملف نفسه. لا تحذف الشرائح القديمة حتى لا يتغيّر الترقيم؛ وإن حذفتها فاضبط «ترقيم الشرائح من» في: تصميم ← حجم الشريحة ← حجم مخصص."),
    ("الإخراج", "ملف ← تصدير ← إنشاء PDF ← خيارات ← «الشريحة الحالية» أو نطاق الشرائح. هذه الشريحة مخفية فلا تُصدَّر ولا تُطبع. للطباعة الورقية: طباعة ← A4 عمودي ← «ملاءمة الورقة»."),
    ("الجدول", "Tab في آخر خلية يضيف صفًا جديدًا. لحذف صف: زر الفأرة الأيمن ← حذف ← الصفوف. اكتب النسبة كرقم مع ٪ والموعد بتاريخ واضح."),
    ("الشعار والجهة", "انقر أيقونة الصورة أعلى اليمين لإدراج شعار جهتك مرة واحدة؛ يُنسخ تلقائيًا مع تكرار الشريحة."),
    ("الخط", "القالب مبني على خط IBM Plex Sans Arabic (مرفق في مجلد fonts، ترخيص مفتوح). ثبّته بالنقر المزدوج ← «تثبيت» (لا يحتاج صلاحيات مسؤول). دون تثبيته يستبدله أوفيس بخط قريب."),
    ("استعادة الحقول", "إن حُذف حقل أو تحرّك بالخطأ: الصفحة الرئيسية ← إعادة تعيين (Reset) فتعود الحقول إلى مواضعها دون فقدان النص."),
    ("نصائح الكتابة", "ابدأ كل مهمة بالنتيجة لا بالنشاط («نُشرت اللوحة واعتمدتها 3 إدارات» بدل «عملت على اللوحة»). خمسة عناصر كحد أقصى لكل قسم. أضف رقمًا أو تاريخًا حيث أمكن. التحديات: جملة واحدة وما فعلته حيالها. القرار المطلوب: طلب واحد واضح بموعد."),
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
                        [para(run("هذه الشريحة مخفية: لا تظهر في العرض ولا في الطباعة ولا في ملف PDF.", 10.5, pal["green"], False), algn="r")],
                        anchor="ctr"))
    y = 42
    for i, (head, body) in enumerate(GUIDE):
        if body is None:
            tree.append(textbox(ids, "عنوان فرعي", 12, y, 186, 8,
                                [para(run(head, 14, pal["navy"], True), algn="r")], anchor="ctr"))
            y += 10
            continue
        n = i  # رقم الخطوة
        tree.append(shape(ids, "رقم الخطوة", 190, y + 0.8, 7, 7, fill=pal["teal"], geom="ellipse",
                          text_xml=para(run(str(n), 9, "FFFFFF", True, lang="en-US"), algn="ctr", rtl=False),
                          body=body_pr("ctr", (0, 0, 0, 0), None), txbox=True))
        tree.append(textbox(ids, f"دليل: {head}", 12, y, 174, 7,
                            [para(run(head, 12, pal["teal_dark"], True), algn="r")], anchor="ctr"))
        lines_est = max(1, int(len(body) / 95) + 1)
        h = 5.2 * lines_est + 2
        tree.append(textbox(ids, f"نص: {head}", 12, y + 7, 174, h,
                            [para(run(body, 10.5, pal["ink"]), algn="r", ln=1.3)], anchor="t"))
        y += 7 + h + 4
    tree.append(textbox(ids, "تذييل الدليل", 12, 284, 186, 6,
                        [para(run("القالب مستوحى من الهوية البصرية لسدايا (ألوان وتدرجات وزخرفة شبكية) دون استخدام شعارها.", 8, pal["faint"]), algn="r")],
                        anchor="ctr"))
    return slide


# ----------------------------------------------------------------------------
# السمة (الخطوط والألوان) وتنظيف التخطيطات
# ----------------------------------------------------------------------------
def patch_theme(prs, cfg):
    theme_part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    font = cfg["font"]
    for tag in ("majorFont", "minorFont"):
        f = root.find(f".//{a}{tag}")
        f.find(f"{a}latin").set("typeface", font)
        f.find(f"{a}latin").attrib.pop("panose", None)
        f.find(f"{a}cs").set("typeface", font)
        for sf in f.findall(f"{a}font"):
            if sf.get("script") == "Arab":
                sf.set("typeface", font)
    pal = cfg["pal"]
    scheme = root.find(f".//{a}clrScheme")
    scheme.set("name", "SDAIA-inspired")
    mapping = {"dk1": pal["ink"], "lt1": "FFFFFF", "dk2": pal["navy"], "lt2": pal["tint_navy"], "accent1": pal["teal"],
               "accent2": pal["green"], "accent3": pal["navy2"], "accent4": pal["amber"], "accent5": pal["teal_dark"],
               "accent6": pal["muted"], "hlink": pal["teal_dark"], "folHlink": pal["muted"]}
    for k, v in mapping.items():
        el = scheme.find(f"{a}{k}")
        for c in list(el):
            el.remove(c)
        el.append(etree.fromstring(f'<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="{v}"/>'))
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_master(prs, cfg):
    """Master text styles: Arabic font + RTL defaults so anything the user adds inherits sensibly."""
    master = prs.slide_master
    font = cfg["font"]
    for el in master.element.iter():
        if el.tag in (qn("a:latin"), qn("a:cs")) and el.get("typeface", "").startswith("+"):
            pass  # theme-mapped, already fine
    # presentation-level default text style: language
    pres = prs.part._element
    for d in pres.iter(qn("a:defRPr")):
        d.set("lang", "ar-SA")
    # حجم مخصص (A4 عمودي) بدل وسم «شاشة 4:3» الموروث من القالب الافتراضي
    sz = pres.find(qn("p:sldSz"))
    sz.attrib.pop("type", None)
    # اسم الشرائح في اللوحة الجانبية
    prs.core_properties.title = "التقرير الأسبوعي للمهام — قالب"
    prs.core_properties.author = ""
    prs.core_properties.language = "ar-SA"
    prs.core_properties.subject = "قالب تقرير أسبوعي للمهام (A4 عمودي، عربي)"


def prune_layouts(prs, keep):
    master = prs.slide_master
    for lay in list(master.slide_layouts):
        if lay not in keep:
            master.slide_layouts.remove(lay)


# ----------------------------------------------------------------------------
def make_logo_png(path):
    """Sample abstract logo for the example file (not any real brand)."""
    from PIL import Image, ImageDraw
    s = 512
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((16, 16, s - 16, s - 16), radius=90, fill=(255, 255, 255, 255))
    d.ellipse((120, 120, 392, 392), outline=(18, 165, 168, 255), width=34)
    d.ellipse((205, 205, 307, 307), fill=(108, 192, 74, 255))
    img.save(path)
    return path


def build(out, example=False, font=DEFAULT_FONT, palette=None, logo_png=None):
    global CONFIG
    CONFIG = {"font": font, "pal": dict(PALETTE, **(palette or {}))}
    prs = Presentation()
    prs.slide_width = Emu(E(PAGE_W))
    prs.slide_height = Emu(E(PAGE_H))
    master = prs.slide_master
    layouts = {l.name: l for l in master.slide_layouts}
    report_layout = layouts["Blank"]
    blank_layout = layouts["Title Only"]
    # نظّف تخطيط «Title Only» ليكون فارغًا تمامًا
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
        import tempfile, os
        logo_png = make_logo_png(os.path.join(tempfile.gettempdir(), "sample_logo.png"))
    add_report_slide(prs, report_layout, CONFIG, EXAMPLE if example else None, logo_png)
    add_guide_slide(prs, blank_layout, CONFIG)
    prs.save(out)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--example", action="store_true", help="نسخة معبأة ببيانات تجريبية")
    ap.add_argument("--font", default=DEFAULT_FONT)
    ap.add_argument("--logo", default=None)
    a = ap.parse_args()
    print(build(a.out, a.example, a.font, logo_png=a.logo))
