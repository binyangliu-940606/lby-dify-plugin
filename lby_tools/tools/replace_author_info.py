from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from lxml import etree

import io
import zipfile
import requests
import re
import copy


class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        old_file_url = tool_parameters["old_file_url"]
        new_file_url = tool_parameters["new_file_url"]

        new_author_info = tool_parameters["new_author_info"]
        old_author_info = tool_parameters["old_author_info"]

        new_ethical_statement = tool_parameters["new_ethical_statement"]
        old_ethical_statement = tool_parameters["old_ethical_statement"]

        new_funding = tool_parameters["new_funding"]
        old_funding = tool_parameters["old_funding"]

        file_name = tool_parameters["file_name"]

         #获取字节流，文件名，mime_type
        out_bytes = apply_replace_or_insert(
            old_file_url=old_file_url,
            new_file_url=new_file_url,
            old_author_info=old_author_info,
            old_ethical_statement=old_ethical_statement,
            old_funding=old_funding,
            new_author_info=new_author_info,
            new_ethical_statement=new_ethical_statement,
            new_funding=new_funding,
        )

        #因为byte字节流传递出错，需要转换为base64，以供工具流中下一节点使用
        # b64_str = base64.b64encode(out_bytes).decode("utf-8")

        # processed_word = {
        #     "data": b64_str,
        #     "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        #     "name": "manuscript-PMID(Marked).docx",
        # }

        # yield self.create_json_message({
        #     "result": {
        #         "processed_word":processed_word,
        #     }
        # })
        yield self.create_blob_message(
            blob=out_bytes,          # 直接传入 bytes，不需要 base64 编码
            meta={
                "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",   # 告诉 Dify 这是什么类型的文件
                "filename": file_name,
            },
        )


# =============================================================================
#  WordprocessingML 命名空间与基础工具
# =============================================================================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(local: str) -> str:
    """生成带命名空间的标签名，例如 w:p => {..}p"""
    return f"{{{W_NS}}}{local}"


def qa(local: str) -> str:
    """生成带命名空间的属性名，例如 w:type => {..}type"""
    return f"{{{W_NS}}}{local}"


def download_bytes(url: str, timeout: int = 120) -> bytes:
    """下载文件字节流"""
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    return r.content


def read_document_xml(docx_bytes: bytes) -> tuple[zipfile.ZipFile, bytes]:
    """读取 docx 中的 word/document.xml"""
    zin = zipfile.ZipFile(io.BytesIO(docx_bytes), "r")
    xml = zin.read("word/document.xml")
    return zin, xml

def extract_author_block_from_reference(ref_data: dict) -> list:
    """
    从参考文件提取“作者信息板块”：
    规则：从文档开头开始取，直到 Ethical Statement 标题段落之前（不包含标题）。
    这样能覆盖作者名单、单位、#说明、通讯作者、邮箱电话等全部段落，并保留格式。
    """
    if not ref_data:
        return []

    all_paras = ref_data.get("all_paragraphs", [])
    if not all_paras:
        return []

    # 找到 "Ethical Statement" 标题在参考文件中的位置
    end_idx = None
    for i, p in enumerate(all_paras):
        if is_title_match(get_paragraph_text(p), ETHICAL_TITLE):
            end_idx = i
            break

    # 如果没找到 Ethical Statement，则认为作者块到文档某个合理位置（这里就取到文末）
    if end_idx is None:
        end_idx = len(all_paras)

    # 从开头取到 end_idx（不包含 end_idx）
    picked = []
    for p in all_paras[:end_idx]:
        # 过滤掉“纯空白且无任何run”的段落（但一般作者块内空行也是段落，需要保留的话可去掉此过滤）
        # 这里选择保留空段落（有些格式化空行在Word里也很重要），所以不做过滤
        picked.append(copy.deepcopy(p))

    return picked

def write_docx_with_new_document_xml(
    zin: zipfile.ZipFile,
    new_document_xml: bytes,
    extra_media_from: bytes | None = None,
) -> bytes:
    """
    将新的 document.xml 写回 docx，并可选拷贝参考文件的 media 资源（保险起见）
    """
    out_buf = io.BytesIO()
    zout = zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED)

    existing_names = set(zin.namelist())

    # 原 docx 的所有文件照抄，但 document.xml 替换为新的
    for item in zin.infolist():
        if item.filename == "word/document.xml":
            zout.writestr(item, new_document_xml)
        else:
            zout.writestr(item, zin.read(item.filename))

    # 额外拷贝参考文件的 media（通常作者/伦理/基金不涉及图片，但拷贝更安全）
    if extra_media_from:
        try:
            zin_new = zipfile.ZipFile(io.BytesIO(extra_media_from), "r")
            for item in zin_new.infolist():
                if item.filename.startswith("word/media/") and item.filename not in existing_names:
                    zout.writestr(item, zin_new.read(item.filename))
            zin_new.close()
        except Exception:
            pass

    zin.close()
    zout.close()
    return out_buf.getvalue()


# =============================================================================
#  XML 段落读取/插入/删除工具
# =============================================================================

def get_body(root):
    return root.find(q("body"))


def get_body_paragraphs(root) -> list:
    """获取 body 下的所有 w:p 段落（不含表格内部段落）"""
    body = get_body(root)
    if body is None:
        return []
    return list(body.findall(q("p")))


def get_paragraph_text(p) -> str:
    """获取段落所有 w:t 拼接后的纯文本"""
    texts = []
    for t in p.iter(q("t")):
        if t.text:
            texts.append(t.text)
    return "".join(texts).strip()


def normalize_text(s: str) -> str:
    """规整文本：压缩空白+小写，用于匹配"""
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def clone_paragraphs(paras: list) -> list:
    """深拷贝段落列表（保留样式/字号/颜色/加粗等）"""
    return [copy.deepcopy(p) for p in paras]


def create_empty_paragraph(template_para=None) -> etree._Element:
    """
    创建一个“空行段落”
    - 如果提供模板段落（template_para），则复制其 pPr，尽量保持段落格式一致
    """
    new_p = etree.Element(q("p"))
    if template_para is not None:
        ppr = template_para.find(q("pPr"))
        if ppr is not None:
            new_p.append(copy.deepcopy(ppr))
    return new_p


def insert_paragraphs_before(body, ref_para, new_paras: list):
    """在 ref_para 前插入段落"""
    children = list(body)
    idx = None
    for i, ch in enumerate(children):
        if ch is ref_para:
            idx = i
            break
    if idx is None:
        for p in new_paras:
            body.append(p)
    else:
        for j, p in enumerate(new_paras):
            body.insert(idx + j, p)


def insert_paragraphs_after(body, ref_para, new_paras: list):
    """在 ref_para 后插入段落"""
    children = list(body)
    idx = None
    for i, ch in enumerate(children):
        if ch is ref_para:
            idx = i
            break
    if idx is None:
        for p in new_paras:
            body.append(p)
    else:
        for j, p in enumerate(new_paras):
            body.insert(idx + 1 + j, p)


def remove_paragraphs_by_index_range(body, body_paras: list, start_idx: int, end_idx: int):
    """删除段落索引区间 [start_idx, end_idx]"""
    for i in range(start_idx, end_idx + 1):
        p = body_paras[i]
        parent = p.getparent()
        if parent is not None:
            parent.remove(p)


# =============================================================================
#  在旧文档里定位“旧内容段落范围”（用于替换时的删除）
# =============================================================================

def find_paragraph_range_containing_text(body_paras: list, search_text: str):
    """
    尝试在 body_paras 中找出包含 search_text 的连续段落范围
    支持 multi-line（search_text 有换行时）
    返回 (start_idx, end_idx) 或 None
    """
    if not (search_text or "").strip():
        return None

    search_lines = [normalize_text(x) for x in (search_text or "").split("\n") if x.strip()]
    para_norms = [normalize_text(get_paragraph_text(p)) for p in body_paras]

    if not search_lines:
        return None

    # 多行：连续段落逐行匹配
    if len(search_lines) > 1:
        # for i in range(len(para_norms)):
        #     if search_lines[0] and search_lines[0] in para_norms[i]:
        #         ok = True
        #         end = i
        #         for j in range(1, len(search_lines)):
        #             if i + j >= len(para_norms):
        #                 ok = False
        #                 break
        #             if search_lines[j] and search_lines[j] not in para_norms[i + j]:
        #                 ok = False
        #                 break
        #             end = i + j
        #         if ok:
        #             return (i, end)
        re_arr = []
        for i in range(len(search_lines)):
            for j in range(len(para_norms)):
                if search_lines[i] and search_lines[i] in para_norms[j]:
                    re_arr.append(j)
        if len(re_arr) > 1:
            return (sorted(re_arr)[0],sorted(re_arr)[len(re_arr)-1])

    # 单段落：全文包含
    full = normalize_text(search_text)
    for i, pn in enumerate(para_norms):
        if full in pn or (len(full) > 20 and pn and pn in full):
            return (i, i)

    return None


def delete_old_block_if_any(body, root, old_text: str) -> None:
    """替换规则第一步：如果 old_text 有值，就在旧文档中定位并删除"""
    if not (old_text or "").strip():
        return
    body_paras = get_body_paragraphs(root)
    rng = find_paragraph_range_containing_text(body_paras, old_text)
    if rng is None:
        return
    s, e = rng
    remove_paragraphs_by_index_range(body, body_paras, s, e)


# =============================================================================
#  标题定位：Ethical Statement / Funding / References
# =============================================================================

# 你给的参考文件标题是完全如下两行：
ETHICAL_TITLE = "Ethical Statement"
FUNDING_TITLE = "Funding"


def is_title_match(text: str, title: str) -> bool:
    """标题匹配：忽略大小写、连续空白差异"""
    return normalize_text(text) == normalize_text(title)


def find_title_paragraph_index(body_paras: list, title: str) -> int | None:
    """在旧文档中找标题段落索引"""
    for i, p in enumerate(body_paras):
        if is_title_match(get_paragraph_text(p), title):
            return i
    return None


def find_references_index(body_paras: list) -> int | None:
    """找 References 的段落索引（常见几种写法）"""
    patterns = [
        re.compile(r"^\s*references?\s*$", re.IGNORECASE),
        re.compile(r"^\s*bibliography\s*$", re.IGNORECASE),
        re.compile(r"^\s*literature\s+cited\s*$", re.IGNORECASE),
        re.compile(r"^\s*works?\s+cited\s*$", re.IGNORECASE),
    ]
    for i, p in enumerate(body_paras):
        txt = get_paragraph_text(p)
        for pat in patterns:
            if pat.match(txt or ""):
                return i
    return None


# =============================================================================
#  作者信息：插到第一页末尾下一行
# =============================================================================

def find_first_page_end_index(body_paras: list) -> int:
    """
    找到第一页最后一个“有文字”的段落索引
    - 优先通过分页符/分节符判断第一页范围
    - 找不到则取前30段中的最后非空段落
    """
    if not body_paras:
        return -1

    page_break_idx = None

    for i, p in enumerate(body_paras):
        # 分页符 <w:br w:type="page"/>
        for br in p.iter(q("br")):
            if br.get(qa("type")) == "page":
                page_break_idx = i
                break
        if page_break_idx is not None:
            break

        # 分节符 <w:sectPr>（很多文档用它分页）
        ppr = p.find(q("pPr"))
        if ppr is not None and ppr.find(q("sectPr")) is not None:
            page_break_idx = i
            break

    # 没找到分页标识：用前30个段落近似第一页
    if page_break_idx is None:
        limit = min(30, len(body_paras))
        last = 0
        for i in range(limit):
            if get_paragraph_text(body_paras[i]).strip():
                last = i
        return last

    # 找到分页标识：在 [0, page_break_idx] 内找最后非空段落
    last = 0
    for i in range(page_break_idx + 1):
        if get_paragraph_text(body_paras[i]).strip():
            last = i
    return last


def insert_author_to_first_page(body, root, author_paras: list) -> None:
    """
    新增规则：
      作者信息插入到文档第一页；如果第一页有数据，在原有数据下一行插入
    实现：
      找第一页末尾非空段落 -> 后插入 [空行] + 作者段落
    """
    body_paras = get_body_paragraphs(root)
    if not body_paras:
        for p in author_paras:
            body.append(copy.deepcopy(p))
        return

    last_idx = find_first_page_end_index(body_paras)
    ref_para = body_paras[last_idx] if last_idx >= 0 else body_paras[0]

    # 用第一页最后段落当模板做空行（尽量格式一致）
    insert_list = [create_empty_paragraph(ref_para)] + clone_paragraphs(author_paras)
    insert_paragraphs_after(body, ref_para, insert_list)


# =============================================================================
#  从参考文件中提取“带格式段落”（保证样式一致）
# =============================================================================

def parse_reference_doc(docx_bytes: bytes) -> dict:
    """解析参考doc，缓存所有段落（带格式）"""
    zin = zipfile.ZipFile(io.BytesIO(docx_bytes), "r")
    doc_xml = zin.read("word/document.xml")
    zin.close()
    root = etree.fromstring(doc_xml, parser=etree.XMLParser(recover=True))
    all_paras = get_body_paragraphs(root)
    return {
        "docx_bytes": docx_bytes,
        "all_paragraphs": [copy.deepcopy(p) for p in all_paras],
    }


def find_block_paragraphs_in_reference_by_exact_text(ref_data: dict, block_text: str) -> list:
    """
    由于你的需求明确：new_xxx 与参考文件内容“一样”
    因此可以通过 block_text 在参考文件中精确定位段落范围，然后直接拷贝 w:p（保留样式）
    """
    if not ref_data or not (block_text or "").strip():
        return []

    paras = ref_data.get("all_paragraphs", [])
    rng = find_paragraph_range_containing_text(paras, block_text)
    if rng is None:
        return []
    s, e = rng
    return clone_paragraphs(paras[s : e + 1])


def make_simple_paragraph(text: str) -> etree._Element:
    """后备：找不到参考样式时，用纯文本生成段落（无样式）"""
    p = etree.Element(q("p"))
    r = etree.Element(q("r"))
    t = etree.Element(q("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def get_formatted_or_fallback(ref_data: dict, block_text: str) -> list:
    """优先从参考文件取带格式的段落；失败才回退纯文本"""
    formatted = find_block_paragraphs_in_reference_by_exact_text(ref_data, block_text)
    if formatted:
        return formatted

    # 回退：按行拆段落
    lines = (block_text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out = []
    for ln in lines:
        out.append(make_simple_paragraph(ln) if ln.strip() else etree.Element(q("p")))
    return out


def find_title_para_in_reference(ref_data: dict, title: str):
    """从参考文件中找到标题段落（用于“旧文档没有标题时”，插入同样的标题格式）"""
    if not ref_data:
        return None
    for p in ref_data.get("all_paragraphs", []):
        if is_title_match(get_paragraph_text(p), title):
            return copy.deepcopy(p)
    return None


# =============================================================================
#  伦理 / 基金 插入规则实现
# =============================================================================

def insert_block_under_title_or_create_before_references(
    body,
    root,
    title: str,
    block_paras: list,
    title_para_from_ref=None,
):
    """
    新增规则：
      - 若旧文档存在标题（Ethical Statement / Funding）：插入到标题下一行
      - 若旧文档不存在标题：需要“加上标题”，并把 [空行 + 标题 + 内容 + 空行] 插入到 References 上方
        （若找不到 References，则插到文末 sectPr 前/文末）
    注意：
      - 标题样式：优先使用参考文件标题段落（title_para_from_ref）
      - 内容样式：block_paras 来自参考文件（带格式）
    """
    body_paras = get_body_paragraphs(root)

    # 1) 有标题：直接插在标题段落后面
    title_idx = find_title_paragraph_index(body_paras, title)
    if title_idx is not None:
        title_p = body_paras[title_idx]
        insert_paragraphs_after(body, title_p, clone_paragraphs(block_paras))
        return

    # 2) 没标题：构造标题段落
    title_p_new = copy.deepcopy(title_para_from_ref) if title_para_from_ref is not None else make_simple_paragraph(title)

    # 3) 在 References 前插入：前后各空一行
    body_paras = get_body_paragraphs(root)
    ref_idx = find_references_index(body_paras)

    # 用 references 上一段做空行模板（更接近原文档排版）
    template_para = (
        body_paras[ref_idx - 1] if (ref_idx is not None and ref_idx > 0) else (body_paras[-1] if body_paras else None)
    )
    empty1 = create_empty_paragraph(template_para)
    empty2 = create_empty_paragraph(template_para)

    insert_list = [empty1, title_p_new] + clone_paragraphs(block_paras) + [empty2]

    if ref_idx is not None:
        ref_para = body_paras[ref_idx]
        insert_paragraphs_before(body, ref_para, insert_list)
    else:
        # 如果没有 References：插到 sectPr 前（如果有），否则追加到末尾
        sect_pr = body.find(q("sectPr"))
        if sect_pr is not None:
            children = list(body)
            idx = None
            for i, ch in enumerate(children):
                if ch is sect_pr:
                    idx = i
                    break
            if idx is None:
                for p in insert_list:
                    body.append(p)
            else:
                for j, p in enumerate(insert_list):
                    body.insert(idx + j, p)
        else:
            for p in insert_list:
                body.append(p)

def build_paragraphs_from_text_lines(text: str) -> list:
    """
    将多行文本按行拆成多个段落。用于参考文件切块失败的兜底。
    注意：此兜底不含参考样式。
    """
    lines = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out = []
    for ln in lines:
        if ln.strip():
            out.append(make_simple_paragraph(ln))
        else:
            out.append(etree.Element(q("p")))
    return out

# =============================================================================
#  核心入口：根据规则对旧文档执行替换/新增
# =============================================================================

def apply_replace_or_insert(
    *,
    old_file_url: str,
    new_file_url: str,
    old_author_info: str,
    old_ethical_statement: str,
    old_funding: str,
    new_author_info: str,
    new_ethical_statement: str,
    new_funding: str,
) -> bytes:
    """
    参数说明（与你需求一致）：
      old_file_url: 被替换文件链接
      new_file_url: 参考文件链接（样式来源）
      old_author_info / old_ethical_statement / old_funding: 旧文档中需要被替换的文本（替换规则用）
      new_author_info / new_ethical_statement / new_funding: 新文本（与参考文件内容一致，用于抓取带格式段落）

    规则说明：
      - 替换规则：旧值有值 -> 先删旧块 -> 再执行新增规则
      - 新增规则：
          作者：第一页末尾下一行插入
          伦理：存在 Ethical Statement 标题则标题下一行插入；否则创建标题+内容，插到 References 上方（前后空一行）
          基金：同理 Funding
      - 格式：插入段落从参考文件复制，保留字号/颜色/粗细等
    """
    # 下载旧文档与参考文档
    old_docx = download_bytes(old_file_url)
    ref_docx = download_bytes(new_file_url)

    # 解析参考文档（用于提取带格式段落、标题段落）
    ref_data = parse_reference_doc(ref_docx)

    # 解析旧文档 XML
    zin, old_xml = read_document_xml(old_docx)
    root = etree.fromstring(old_xml, parser=etree.XMLParser(recover=True))
    body = get_body(root)
    if body is None:
        zin.close()
        return old_docx

    # ----------------------------
    # 作者信息：替换/新增
    # ----------------------------
    if (new_author_info or "").strip():
        # 替换规则：先删旧作者块
        delete_old_block_if_any(body, root, old_author_info)

        # 关键：作者块不要用全文匹配，直接按参考文件板块切出“作者信息”全部段落（保留格式）
        author_paras = extract_author_block_from_reference(ref_data)

        # 兜底：如果参考文件切块失败，再退回用文本构建（至少内容完整）
        if not author_paras:
            author_paras = build_paragraphs_from_text_lines(new_author_info)

        insert_author_to_first_page(body, root, author_paras)

    # ----------------------------
    # 伦理法则：替换/新增
    # ----------------------------
    if (new_ethical_statement or "").strip():
        delete_old_block_if_any(body, root, old_ethical_statement)

        ethical_paras = get_formatted_or_fallback(ref_data, new_ethical_statement)
        ethical_title_ref = find_title_para_in_reference(ref_data, ETHICAL_TITLE)

        insert_block_under_title_or_create_before_references(
            body=body,
            root=root,
            title=ETHICAL_TITLE,
            block_paras=ethical_paras,
            title_para_from_ref=ethical_title_ref,
        )

    # ----------------------------
    # 项目基金：替换/新增
    # ----------------------------
    if (new_funding or "").strip():
        delete_old_block_if_any(body, root, old_funding)

        funding_paras = get_formatted_or_fallback(ref_data, new_funding)
        funding_title_ref = find_title_para_in_reference(ref_data, FUNDING_TITLE)

        insert_block_under_title_or_create_before_references(
            body=body,
            root=root,
            title=FUNDING_TITLE,
            block_paras=funding_paras,
            title_para_from_ref=funding_title_ref,
        )

    # 写回 document.xml 并生成新的 docx
    new_document_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    out_docx = write_docx_with_new_document_xml(zin, new_document_xml, extra_media_from=ref_docx)
    return out_docx



