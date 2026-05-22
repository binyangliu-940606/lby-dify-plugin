import io
import os
import re
import json
import zipfile

from collections.abc import Generator
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from datetime import datetime, timezone
from typing import Any, Optional, Tuple

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor
from lxml import etree as LET



# ---------------------------------------------------------------------------
# Tool 入口
# ---------------------------------------------------------------------------

class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        payload = json.loads(tool_parameters["payload_json"])
       
        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        final_data_json = payload["final_data_json"]
        mode = payload["mode"]
        file_name_clean = payload["file_name_clean"]
        file_name_track = payload["file_name_track"]
        old_data_json = payload["old_data_json"]

        out_bytes_clear,out_bytes_track = generate_title_page_files(
            final_data_json,                      # dict
            mode=mode,                 # 或其它字符串则只出 Clean
            old_data_json=old_data_json,         # revise_format 时用于生成批注对比
        )
        
        yield self.create_blob_message(
            blob=out_bytes_clear,
            meta={
                "mime_type": mime_type,
                "filename": file_name_clean,
            },
        )

        if mode=='revise_format':
            yield self.create_blob_message(
                blob=out_bytes_track,
                meta={
                    "mime_type": mime_type,
                    "filename": file_name_track,
                },
            )

# -*- coding: utf-8 -*-
"""
根据 final_data_json / old_data_json 生成论文标题页作者信息区 Word 文档。

- mode == "revise_format": 输出 Clean + 痕迹版（内容一致，痕迹版带 Word 批注说明相对 old 的增删改）
- 其他 mode: 仅输出 Clean

依赖: python-docx, lxml
"""

# Word 批注命名空间
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NSMAP = {"w": W_NS}
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# _DIGIT_SUPER = str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")
# LINK_BLUE_HEX = "0563C1"
_DIGIT_SUPER = str.maketrans("0123456789", "0123456789")
LINK_BLUE_HEX = "0000FF"


def _sanitize(s: str | None) -> str:
    if s is None:
        return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(s))


def _unicode_superscript_num(n: int) -> str:
    return str(n).translate(_DIGIT_SUPER)


def _format_affiliation_line_from_raw(raw: dict[str, Any]) -> str:
    """从 raw 生成单行英文机构地址（避免 affiliation_lines.text 含 dict repr）。"""
    if not raw:
        return ""

    dept_en = (raw.get("name_en") or "").strip()
    dept_cn = (raw.get("department") or "").strip()
    dept = dept_en or dept_cn

    inst = (raw.get("institution") or raw.get("english_name") or "").strip()

    addr_field = raw.get("address")
    street_part = ""
    if isinstance(addr_field, dict):
        street_part = (addr_field.get("street") or "").strip()
    elif isinstance(addr_field, str) and addr_field.strip():
        if addr_field.strip().isdigit() and len(addr_field.strip()) == 6:
            street_part = ""
        else:
            street_part = addr_field.strip()

    city = (raw.get("city") or "").strip()
    state = (raw.get("state") or "").strip()
    postal = (raw.get("postal_code") or "").strip()
    country = (raw.get("country") or "").strip()

    segments: list[str] = []
    if dept and inst:
        segments.append(f"{dept}, {inst}")
    elif dept:
        segments.append(dept)
    elif inst:
        segments.append(inst)

    tail_bits: list[str] = []
    if street_part:
        tail_bits.append(street_part)

    loc = ", ".join(x for x in [city, state] if x)
    if postal:
        if loc:
            loc = f"{loc} {postal}"
        else:
            loc = postal
    if loc:
        tail_bits.append(loc)
    if country:
        tail_bits.append(country)

    if tail_bits:
        segments.append(", ".join(tail_bits))

    return ", ".join(segments)


def _collect_affiliations_final(final_data: dict[str, Any]) -> list[tuple[int, str, dict]]:
    """返回 [(num, formatted_line, raw), ...] 按 num 排序。"""
    by_num: dict[int, tuple[str, dict]] = {}
    for line in final_data.get("affiliation_lines") or []:
        num = int(line.get("num") or 0)
        raw = line.get("raw") or {}

        department = raw.get('department','')
        institution = raw.get('institution','')
        address = raw.get('address','')
        city = raw.get('city','')
        state = raw.get('state','')
        postal_code = raw.get('postal_code','')
        country = raw.get('country','')
        
        txt = f'{department}, {institution}, {state} {city} {postal_code}, {country}'

        # txt = (line.get("text") or "").strip()
        # if "{'street'" in txt or "{" in txt and "'street'" in txt:
        #     txt = _format_affiliation_line_from_raw(raw)
        # elif not txt:
        #     txt = _format_affiliation_line_from_raw(raw)

        by_num[num] = (_sanitize(txt), raw)

    return [(n, by_num[n][0], by_num[n][1]) for n in sorted(by_num.keys())]


def _build_author_line_runs_final(final_data: dict[str, Any]) -> str:
    """纯文本作者行（用于 diff）；上标以 Unicode 数字与 *# 字符表示。"""
    authors = final_data.get("authors") or []
    chunks: list[str] = []
    for a in authors:
        name = _sanitize(a.get("full_name") or "")
        if not name:
            continue
        nums = a.get("affiliation_refs") or []
        sup = "".join(_unicode_superscript_num(int(x)) for x in nums if x is not None)
        marks = list(a.get("footnote_marks") or [])
        # if not marks and a.get("is_corresponding"):
        #     marks = ["*"]
        if a.get("is_corresponding") and '*' not in marks:
            marks.append("*")
        if a.get("is_cofirst") and '#' not in marks:
            marks.append("#")
        mark_s = ""
        for m in marks:
            m = str(m).strip()
            if m in ("*", "#", "†", "§"):
                mark_s += m
        chunks.append(f"{name}{sup}{mark_s}")
    return ", ".join(chunks)


def _add_hyperlink(paragraph, url: str, text: str, color_hex: str | None = LINK_BLUE_HEX) -> None:
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    r_pr.append(u)

    if color_hex:
        hx = color_hex.strip().lstrip("#")
        if len(hx) == 6:
            c = OxmlElement("w:color")
            c.set(qn("w:val"), hx.upper())
            r_pr.append(c)

    new_run.append(r_pr)

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


# def _set_run_superscript(run, on: bool = True) -> None:
#     run.font.superscript = on
def _set_run_superscript(run, on: bool = True) -> None:
    # 设置上标状态
    run.font.superscript = on
    
    run.font.bold = True
    # 上标颜色设为 RGB(0,0,255)
    run.font.color.rgb = RGBColor(0, 0, 255)
    # 如果需要同时支持下标（调用方可改用 run.font.subscript = True），
    # 也可在这里设置 run.font.subscript = True 并设置颜色。

def _populate_author_paragraph(paragraph, final_data: dict[str, Any]) -> None:
    authors = final_data.get("authors") or []
    for i, a in enumerate(authors):
        name = _sanitize(a.get("full_name") or "")
        if not name:
            continue
        if i:
            paragraph.add_run(", ")

        r_name = paragraph.add_run(name)
        r_name.bold = True

        nums = a.get("affiliation_refs") or []
        for num in nums:
            r = paragraph.add_run(_unicode_superscript_num(int(num)))
            _set_run_superscript(r)

        # marks = list(a.get("footnote_marks") or [])
        # if not marks and a.get("is_corresponding"):
        #     marks = ["*"]
        # for m in marks:
        #     m = str(m).strip()
        #     if not m:
        #         continue
        #     r = paragraph.add_run(m)
        #     _set_run_superscript(r)
        marks = list(a.get("footnote_marks") or [])
        # if not marks:
            # marks = []
        if a.get("is_corresponding") and '*' not in marks:
            marks.append("*")
        if a.get("is_cofirst") and '#' not in marks:
            marks.append("#")
        for m in marks:
            m = str(m).strip()
            if not m:
                continue
            r = paragraph.add_run(m)
            _set_run_superscript(r)


def _populate_affiliation_paragraph(paragraph, num: int, line_text: str) -> None:
    r_sup = paragraph.add_run(_unicode_superscript_num(num))
    _set_run_superscript(r_sup)
    paragraph.add_run(" ")
    paragraph.add_run(_sanitize(line_text))


def _correspondence_paragraphs_body(final_data: dict[str, Any]) -> tuple[str, str, str, str]:
    """返回 (通讯作者后的地址整句（不含姓名）, 邮箱, E-mail 标签, Tel 整行)。"""
    cb = final_data.get("correspondence_block") or {}
    name = _sanitize(cb.get("contact_person") or "")
    email = _sanitize(cb.get("email") or "")

    refs = _collect_affiliations_final(final_data)
    raw_by_num = {n: r for n, _t, r in refs}

    aff_nums = []
    # for a in final_data.get("authors") or []:
    #     if (a.get("contact_person") or "").strip() == name:
    #         aff_nums = a.get("affiliation_refs") or []
    #         break
    # if not aff_nums:
    for a in final_data.get("authors") or []:
        if _sanitize(a.get("full_name") or "") == name:
            aff_nums = a.get("affiliation_refs") or []
            break

    primary = int(aff_nums[0]) if aff_nums else None
    raw = raw_by_num.get(primary) if primary is not None else {}

    # inst = (cb.get("org") or raw.get("institution") or "").strip()
    # dept = (raw.get("name_en") or raw.get("department") or "").strip()

    inst = (raw.get("institution") or cb.get("org") or "").strip()
    dept = (raw.get("department") or raw.get("name_en") or "").strip()

    addr_field = cb.get("address")
    street = ""
    if isinstance(addr_field, dict):
        street = (addr_field.get("street") or "").strip()
    elif isinstance(addr_field, str) and addr_field.strip():
        if not addr_field.strip().isdigit():
            street = addr_field.strip()

    if not street:
        af = raw.get("address")
        if isinstance(af, dict):
            street = (af.get("street") or "").strip()

    city = _sanitize(cb.get("city") or raw.get("city"))
    state = _sanitize(raw.get("state"))
    postal = _sanitize(cb.get("postal_code") or raw.get("postal_code"))
    country = _sanitize(cb.get("country") or raw.get("country"))

    line_bits: list[str] = []
    org_line = ", ".join(x for x in [dept, inst] if x)
    if org_line:
        line_bits.append(org_line)
    tail: list[str] = []
    if street:
        tail.append(street)
    loc = ", ".join(x for x in [city, state] if x)
    if postal:
        loc = f"{loc} {postal}".strip() if loc else postal
    if loc:
        tail.append(loc)
    if country:
        tail.append(country)
    if tail:
        line_bits.append(", ".join(tail))

    line1_middle = ", ".join(line_bits) if line_bits else ""

    p_email_label = "E-mail:"
    p_tel = ""
    tel = _sanitize(cb.get("tel"))
    if tel:
        p_tel = f"Tel.: {tel}"
    else:
        p_tel = "Tel.:"

    return line1_middle, email, p_email_label, p_tel


def _funding_text_final(final_data: dict[str, Any]) -> str:
    items = (final_data.get("funding_block") or {}).get("items") or []
    parts = []
    for it in items:
        raw = _sanitize((it or {}).get("raw"))
        if raw:
            parts.append(raw)
    return "\n".join(parts) if parts else ""


# def build_title_page_document(final_data: dict[str, Any]) -> Document:
#     # doc = Document()
#     # style = doc.styles["Normal"]
#     # style.font.name = "Times New Roman"
#     # style.font.size = Pt(12)

#     doc = Document()

#     # 设置页边距为 2cm
#     section = doc.sections[0]
#     section.top_margin = Cm(2)
#     section.bottom_margin = Cm(2)
#     section.left_margin = Cm(2)
#     section.right_margin = Cm(2)

#     # 设置默认样式：英文 Times New Roman，中文（East Asia）宋体，字号 12pt，段落行距双倍
#     style = doc.styles["Normal"]
#     style.font.name = "Times New Roman"
#     style.font.size = Pt(12)
#     # 设置中文字体（East Asia）
#     style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
#     # 设置默认段落双倍行距
#     style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE


#     # 作者行
#     p_auth = doc.add_paragraph()
#     p_auth.paragraph_format.alignment = 0  # LEFT
#     p_auth.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
#     _populate_author_paragraph(p_auth, final_data)

#     # # 小间距
#     # doc.add_paragraph()

#     # 单位
#     for num, line_text, _raw in _collect_affiliations_final(final_data):
#         p = doc.add_paragraph()
#         p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
#         _populate_affiliation_paragraph(p, num, line_text)

#     doc.add_paragraph()
#     # doc.add_paragraph()

#     cb = final_data.get("correspondence_block") or {}
#     name = _sanitize(cb.get("contact_person") or "")
#     middle, email, email_label, tel_line = _correspondence_paragraphs_body(final_data)

#     p_corr = doc.add_paragraph()
#     p_corr.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
#     r_ast = p_corr.add_run("*")
#     _set_run_superscript(r_ast)
#     p_corr.add_run(" ")
#     r_lbl = p_corr.add_run("Correspondence to:")
#     r_lbl.bold = True
#     p_corr.add_run(" ")
#     r_nm = p_corr.add_run(name)
#     r_nm.bold = True
#     if middle:
#         p_corr.add_run(", ")
#         p_corr.add_run(middle)

#     p_mail = doc.add_paragraph()
#     p_mail.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
#     r_m0 = p_mail.add_run(email_label + " ")
#     r_m0.bold = True
#     if email:
#         href = email if email.lower().startswith("mailto:") else f"mailto:{email}"
#         _add_hyperlink(p_mail, href, email, LINK_BLUE_HEX)
#     else:
#         p_mail.add_run("")


#     p_tel = doc.add_paragraph()
#     p_tel.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
#     if tel_line.startswith("Tel."):
#         r_tl = p_tel.add_run("Tel.:")
#         r_tl.bold = True
#         rest = tel_line[5:].lstrip()
#         if rest:
#             p_tel.add_run(" " + rest)
#     else:
#         p_tel.add_run(tel_line)

#     doc.add_paragraph()
#     # doc.add_paragraph()

#     p_fund_h = doc.add_paragraph()
#     p_fund_h.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
#     r_f = p_fund_h.add_run("Funding")
#     r_f.bold = True

#     fund_txt = _funding_text_final(final_data)
#     if fund_txt:
#         p_fund_b = doc.add_paragraph()
#         p_fund_b.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
#         p_fund_b.add_run(fund_txt)

#     return doc



def build_title_page_document(final_data: dict[str, Any]) -> Document:
# helper to set paragraph spacing to 0
    def _zero_para_spacing(p):
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        return p

    doc = Document()

    # 设置页边距为 2cm
    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    # 设置默认样式：英文 Times New Roman，中文（East Asia）宋体，字号 12pt，段落行距双倍
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    # 设置中文字体（East Asia）
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    # 设置默认段落双倍行距
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    # 设置默认段前段后为 0
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # 作者行
    p_auth = _zero_para_spacing(doc.add_paragraph())
    p_auth.paragraph_format.alignment = 0  # LEFT
    p_auth.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    _populate_author_paragraph(p_auth, final_data)

    # 单位
    for num, line_text, _raw in _collect_affiliations_final(final_data):
        p = _zero_para_spacing(doc.add_paragraph())
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        _populate_affiliation_paragraph(p, num, line_text)


    p_jing = _zero_para_spacing(doc.add_paragraph())
    p_jing.paragraph_format.alignment = 0  # LEFT
    p_jing.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    r_jing = p_jing.add_run('#')
    _set_run_superscript(r_jing)
    p_jing.add_run(" ")
    p_jing.add_run(_sanitize('These authors are regarded as co-first authors'))


    # 小的空行（仍设为段前段后为0）
    tmp = _zero_para_spacing(doc.add_paragraph())

    cb = final_data.get("correspondence_block") or {}
    name = _sanitize(cb.get("contact_person") or "")
    middle, email, email_label, tel_line = _correspondence_paragraphs_body(final_data)

    p_corr = _zero_para_spacing(doc.add_paragraph())
    p_corr.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    r_ast = p_corr.add_run("*")
    _set_run_superscript(r_ast)
    p_corr.add_run(" ")
    r_lbl = p_corr.add_run("Correspondence to:")
    r_lbl.bold = True
    p_corr.add_run(" ")
    r_nm = p_corr.add_run(name)
    r_nm.bold = True
    if middle:
        p_corr.add_run(", ")
        p_corr.add_run(middle)

    p_mail = _zero_para_spacing(doc.add_paragraph())
    p_mail.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    r_m0 = p_mail.add_run(email_label + " ")
    r_m0.bold = True
    if email:
        href = email if email.lower().startswith("mailto:") else f"mailto:{email}"
        _add_hyperlink(p_mail, href, email, LINK_BLUE_HEX)
    else:
        p_mail.add_run("")

    p_tel = _zero_para_spacing(doc.add_paragraph())
    p_tel.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    if tel_line.startswith("Tel."):
        r_tl = p_tel.add_run("Tel.:")
        r_tl.bold = True
        rest = tel_line[5:].lstrip()
        if rest:
            p_tel.add_run(" " + rest)
    else:
        p_tel.add_run(tel_line)

    # 另一空行
    tmp2 = _zero_para_spacing(doc.add_paragraph())

    p_fund_h = _zero_para_spacing(doc.add_paragraph())
    p_fund_h.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    r_f = p_fund_h.add_run("Funding")
    r_f.bold = True

    fund_txt = _funding_text_final(final_data)
    if fund_txt:
        p_fund_b = _zero_para_spacing(doc.add_paragraph())
        p_fund_b.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        p_fund_b.add_run(fund_txt)

    return doc

# ---------------------------------------------------------------------------
# old_data_json → 纯文本快照（用于对比）
# ---------------------------------------------------------------------------


def _old_author_line(old: dict[str, Any]) -> str:
    authors = old.get("authors") or []
    chunks = []
    for au in authors:
        name = _sanitize(au.get("full_name"))
        nums = au.get("affiliation_refs") or []
        sup = "".join(_unicode_superscript_num(int(n)) for n in nums)
        marks = au.get("footnote_marks") or []
        ic = au.get("is_corresponding")
        if isinstance(ic, str) and ic.lower() in ("yes", "True", "1"):
            if "*" not in marks:
                marks = list(marks) + ["*"]
        elif ic is True and "*" not in marks:
            marks = list(marks) + ["*"]
        mark_s = "".join(str(m) for m in marks if m)
        chunks.append(f"{name}{sup}{mark_s}")
    return ", ".join(chunks)


def _old_affiliation_block(old: dict[str, Any]) -> str:
    lines = []
    for a in sorted(old.get("affiliations") or [], key=lambda x: int(x.get("idx") or 0)):
        num = int(a.get("idx") or 0)
        base = _sanitize(a.get("name_en") or a.get("institution"))
        addr = _sanitize(a.get("address"))
        city = _sanitize(a.get("city"))
        st = _sanitize(a.get("state"))
        pc = _sanitize(a.get("postal_code"))
        co = _sanitize(a.get("country"))
        tail = ", ".join(x for x in [addr, city, st, pc, co] if x)
        if base and tail:
            lines.append(f"{_unicode_superscript_num(num)} {base}, {tail}")
        elif base:
            lines.append(f"{_unicode_superscript_num(num)} {base}")
    return "\n".join(lines)


def _old_correspondence_block(old: dict[str, Any]) -> str:
    c = old.get("correspondence") or {}
    parts = []
    name = _sanitize(c.get("contact_person"))
    org = _sanitize(c.get("org"))
    addr = _sanitize(c.get("address"))
    city = _sanitize(c.get("city"))
    st = _sanitize(c.get("state"))
    pc = _sanitize(c.get("postal_code"))
    co = _sanitize(c.get("country"))
    email = _sanitize(c.get("email"))
    tel = _sanitize(c.get("tel"))
    line1 = f"* Correspondence to: {name}, {org}"
    if addr or city or pc:
        line1 += ", " + ", ".join(x for x in [addr, city, st, pc, co] if x)
    parts.append(line1)
    parts.append(f"E-mail: {email}")
    parts.append(f"Tel.: {tel}" if tel else "Tel.:")
    return "\n".join(parts)


def _old_funding_block(old: dict[str, Any]) -> str:
    items = old.get("funding") or []
    return "\n".join(_sanitize((x or {}).get("raw")) for x in items if (x or {}).get("raw"))


def snapshot_old_for_diff(old: dict[str, Any] | None) -> dict[str, str]:
    if not old:
        return {"authors": "", "affiliations": "", "correspondence": "", "funding": ""}
    return {
        "authors": _old_author_line(old),
        "affiliations": _old_affiliation_block(old),
        "correspondence": _old_correspondence_block(old),
        "funding": _old_funding_block(old),
    }


def snapshot_new_for_diff(final_data: dict[str, Any]) -> dict[str, str]:
    aff_lines = []
    for num, line_text, _r in _collect_affiliations_final(final_data):
        aff_lines.append(f"{_unicode_superscript_num(num)} {_sanitize(line_text)}")
    cb = final_data.get("correspondence_block") or {}
    middle, email, email_label, tel_line = _correspondence_paragraphs_body(final_data)
    name = _sanitize(cb.get("contact_person"))
    c_lines = [
        f"* Correspondence to: {name}, {middle}".replace(", ,", ",").strip().rstrip(","),
        (email_label + " " + email).strip(),
        tel_line,
    ]
    return {
        "authors": _build_author_line_runs_final(final_data),
        "affiliations": "\n".join(aff_lines),
        "correspondence": "\n".join(c_lines),
        "funding": _funding_text_final(final_data),
    }


def _diff_comment(old_s: str, new_s: str) -> str | None:
    o = (old_s or "").strip()
    n = (new_s or "").strip()
    if o == n:
        return None
    if not o and n:
        return "【增加】\n" + n
    if o and not n:
        return "【删除】\n" + o
    return "【修改】\n原：" + o + "\n新：" + n


def collect_section_comments(
    final_data: dict[str, Any], old_data: dict[str, Any] | None
) -> list[tuple[str, str]]:
    """返回 [(section_key, comment_text), ...]"""
    old_snap = snapshot_old_for_diff(old_data)
    new_snap = snapshot_new_for_diff(final_data)
    order = ["authors", "affiliations", "correspondence", "funding"]
    out: list[tuple[str, str]] = []
    for key in order:
        msg = _diff_comment(old_snap.get(key, ""), new_snap.get(key, ""))
        if msg:
            out.append((key, msg))
    return out


# ---------------------------------------------------------------------------
# OOXML 批注注入（痕迹版）
# ---------------------------------------------------------------------------


def _next_rid(rels_root: LET._Element) -> str:
    max_id = 0
    for rel in rels_root:
        rid = rel.get("Id") or ""
        if rid.startswith("rId"):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass
    return f"rId{max_id + 1}"


def _patch_content_types(zf: zipfile.ZipFile) -> bytes:
    data = zf.read("[Content_Types].xml")
    root = LET.fromstring(data)
    part = "/word/comments.xml"
    exists = False
    for ov in root.iter():
        if ov.tag.endswith("Override") and ov.get("PartName") == part:
            exists = True
            break
    if not exists:
        ov = LET.Element(f"{{{CT_NS}}}Override")
        ov.set("PartName", part)
        ov.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        )
        root.append(ov)
    return LET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _merge_comments_xml(existing_bytes: bytes | None, new_fragment: LET._Element) -> bytes:
    if not existing_bytes:
        return LET.tostring(new_fragment, xml_declaration=True, encoding="UTF-8", standalone=True)
    root = LET.fromstring(existing_bytes)
    for child in list(new_fragment):
        root.append(child)
    return LET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _paragraph_indices_for_sections(num_aff: int) -> dict[str, int]:
    """
    与 build_title_page_document 段落顺序一致（0-based 段落索引）。
    0 作者, 1 空, 2..2+num_aff-1 单位, ...
    """
    i = 0
    m: dict[str, int] = {}
    m["authors"] = i
    i += 1  # blank
    i += 1
    for _ in range(num_aff):
        m.setdefault("_aff_start", i)
        i += 1
    aff_start = m.pop("_aff_start", i - num_aff)
    m["affiliations"] = aff_start  # 批注挂在首条单位段
    i += 2  # two blanks
    m["correspondence"] = i
    i += 3  # corr + mail + tel
    i += 2  # blanks
    m["funding"] = i
    return m


def inject_comments_into_docx_bytes(
    docx_bytes: bytes,
    section_messages: list[tuple[str, str]],
    num_affiliations: int,
    author_display: str = "格式修订",
) -> bytes:
    """在内存中向 docx 注入批注；section_messages: [(authors|affiliations|correspondence|funding, text)]"""
    if not section_messages:
        return docx_bytes

    indices = _paragraph_indices_for_sections(num_affiliations)
    sec_to_idx = {k: indices[k] for k in ["authors", "affiliations", "correspondence", "funding"]}

    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        names = set(zin.namelist())
        rels_path = "word/_rels/document.xml.rels"

        rels_root = LET.fromstring(zin.read(rels_path))
        new_rid = _next_rid(rels_root)
        has_rel = False
        for rel in rels_root:
            if (rel.get("Type") or "").endswith("/comments"):
                has_rel = True
                new_rid = rel.get("Id")
                break
        if not has_rel:
            rel_el = LET.Element(f"{{{REL_NS}}}Relationship")
            rel_el.set("Id", new_rid)
            rel_el.set(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            )
            rel_el.set("Target", "comments.xml")
            rels_root.append(rel_el)

        doc_root = LET.fromstring(zin.read("word/document.xml"))
        body = doc_root.find(f".//{{{W_NS}}}body")
        if body is None:
            return docx_bytes

        old_comments = zin.read("word/comments.xml") if "word/comments.xml" in names else None
        max_id = -1
        if old_comments:
            cr = LET.fromstring(old_comments)
            for c in cr:
                if c.tag == f"{{{W_NS}}}comment":
                    try:
                        max_id = max(max_id, int(c.get(f"{{{W_NS}}}id") or -1))
                    except ValueError:
                        pass

        new_comments_root = LET.Element(f"{{{W_NS}}}comments", nsmap=_NSMAP)
        cid = max_id + 1

        for sec, msg in section_messages:
            p_idx = sec_to_idx.get(sec)
            if p_idx is None:
                continue
            paras = [c for c in body if c.tag == f"{{{W_NS}}}p"]
            if p_idx >= len(paras):
                continue
            p_el = paras[p_idx]

            c_el = LET.SubElement(new_comments_root, f"{{{W_NS}}}comment")
            c_el.set(f"{{{W_NS}}}id", str(cid))
            c_el.set(f"{{{W_NS}}}author", author_display)
            c_el.set(
                f"{{{W_NS}}}date",
                datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            )
            c_el.set(f"{{{W_NS}}}initials", "RV")
            cp = LET.SubElement(c_el, f"{{{W_NS}}}p")
            cr_el = LET.SubElement(cp, f"{{{W_NS}}}r")
            ct = LET.SubElement(cr_el, f"{{{W_NS}}}t")
            ct.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            ct.text = msg

            start = LET.Element(f"{{{W_NS}}}commentRangeStart")
            start.set(f"{{{W_NS}}}id", str(cid))
            end = LET.Element(f"{{{W_NS}}}commentRangeEnd")
            end.set(f"{{{W_NS}}}id", str(cid))
            ref_run = LET.Element(f"{{{W_NS}}}r")
            ref = LET.SubElement(ref_run, f"{{{W_NS}}}commentReference")
            ref.set(f"{{{W_NS}}}id", str(cid))

            p_el.insert(0, start)
            p_el.append(end)
            p_el.append(ref_run)
            cid += 1

        merged_comments = _merge_comments_xml(old_comments, new_comments_root)

        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if name in (rels_path, "word/document.xml", "word/comments.xml", "[Content_Types].xml"):
                    continue
                zout.writestr(item, zin.read(name))

            zout.writestr(rels_path, LET.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            zout.writestr("word/document.xml", LET.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            zout.writestr("word/comments.xml", merged_comments)

            zout.writestr("[Content_Types].xml", _patch_content_types(zin))

    return buf.getvalue()


def generate_title_page_files_docx(
    final_data_json: dict[str, Any],
    *,
    mode: str,
    file_name_clean: str,
    file_name_track: str | None = None,
    old_data_json: dict[str, Any] | None = None,
    output_dir: str = ".",
    comment_author: str = "格式修订",
) -> list[str]:
    """
    按 mode 生成 Clean（必选）；mode == \"revise_format\" 时额外生成痕迹版（批注）。

    返回已写入文件的绝对路径列表。
    """
    os.makedirs(output_dir, exist_ok=True)
    doc = build_title_page_document(final_data_json)
    clean_path = os.path.abspath(os.path.join(output_dir, file_name_clean))
    doc.save(clean_path)
    out = [clean_path]

    if (mode or "").strip() != "revise_format":
        return out

    track_name = file_name_track or (os.path.splitext(file_name_clean)[0] + "_痕迹版.docx")
    track_path = os.path.abspath(os.path.join(output_dir, track_name))

    msgs = collect_section_comments(final_data_json, old_data_json)
    num_aff = len(_collect_affiliations_final(final_data_json))

    with open(clean_path, "rb") as f:
        raw = f.read()
    tracked = inject_comments_into_docx_bytes(
        raw,
        msgs,
        num_affiliations=num_aff,
        author_display=comment_author,
    )
    with open(track_path, "wb") as f:
        f.write(tracked)
    out.append(track_path)
    return out


def generate_title_page_files(
    final_data_json: dict[str, Any],
    *,
    mode: str,
    old_data_json: dict[str, Any] | None = None,
    comment_author: str = "格式修订",
) -> Tuple[bytes, Optional[bytes]]:
    """
    按 mode 生成 Clean（必选）；mode == "revise_format" 时额外生成痕迹版（批注）。

    返回一个二元组 (clean_bytes, tracked_bytes_or_none)：
      - clean_bytes: 清稿的 docx 二进制数据（bytes）
      - tracked_bytes_or_none: 若 mode == "revise_format" 则为痕迹版的 docx bytes，否则为 None
    """
    # 生成清稿文档对象并保存到内存 BytesIO 获取 bytes
    doc = build_title_page_document(final_data_json)
    bio = io.BytesIO()
    doc.save(bio)
    clean_bytes = bio.getvalue()

    if (mode or "").strip() != "revise_format":
        return clean_bytes, None

    # 生成痕迹版所需的注释信息
    msgs = collect_section_comments(final_data_json, old_data_json)
    num_aff = len(_collect_affiliations_final(final_data_json))

    # 将清稿 bytes 注入批注，得到痕迹版 bytes
    tracked_bytes = inject_comments_into_docx_bytes(
        clean_bytes,
        msgs,
        num_affiliations=num_aff,
        author_display=comment_author,
    )
    return clean_bytes, tracked_bytes