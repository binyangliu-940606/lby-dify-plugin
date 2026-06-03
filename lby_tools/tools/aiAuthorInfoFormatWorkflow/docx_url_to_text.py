import io
import zipfile
import requests
import json
from lxml import etree
from typing import Any


from collections.abc import Generator
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

# ---------------------------------------------------------------------------
# Tool 入口
# ---------------------------------------------------------------------------

class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        file_url = tool_parameters["file_url"]

        text = docx_url_to_text(file_url)

        yield self.create_text_message(text)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}

def _clean_revisions_and_comments(root):
    """
    对一个 word XML root 做如下清理：
    - 删除所有 <w:del> 和 <w:moveFrom>（删除或移动来源内容）
    - 将 <w:ins> 和 <w:moveTo> "unwrap"（保留其内部文本）
    - 删除注释标记：w:commentRangeStart, w:commentRangeEnd, w:commentReference
    - 删除所有 w:delText（作为保险）
    """
    # 删除全部 w:del、w:moveFrom、w:delText
    for tag in ("del", "moveFrom", "delText"):
        for el in root.xpath(".//w:" + tag, namespaces=NSMAP):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    # 展开 w:ins, w:moveTo （将子元素提升到父节点位置）
    for tag in ("ins", "moveTo"):
        for el in root.xpath(".//w:" + tag, namespaces=NSMAP):
            parent = el.getparent()
            if parent is None:
                continue
            index = parent.index(el)
            # 插入 el 的子元素到 parent 的该位置
            for child in list(el):
                parent.insert(index, child)
                index += 1
            # 删除原来的 el
            parent.remove(el)

    # 删除注释标记（这些本身不包含要保留的文本）
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for el in root.xpath(".//w:" + tag, namespaces=NSMAP):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

def _text_of_element(el):
    """
    按文档顺序遍历元素，提取文本：
    - w:t 的文本
    - w:br -> 换行
    - w:tab -> tab
    - w:cr -> 换行
    """
    parts = []
    for node in el.iter():
        # node.tag 可能包含 namespace
        tag = etree.QName(node.tag).localname if isinstance(node.tag, str) else None
        if tag == "t":
            if node.text:
                parts.append(node.text)
        elif tag == "br" or tag == "cr":
            parts.append("\n")
        elif tag == "tab":
            parts.append("\t")
    return "".join(parts)

def _extract_text_from_xml_bytes(xml_bytes):
    """
    返回这个 word XML（document/header/footer/footnotes 等）中的按段落组织的文本列表
    """
    parser = etree.XMLParser(remove_comments=True, recover=True)
    root = etree.fromstring(xml_bytes, parser=parser)
    # 清理修订/注释标记
    _clean_revisions_and_comments(root)

    paragraphs = []
    # 所有段落 w:p（包括表格单元里的 p）
    for p in root.xpath(".//w:p", namespaces=NSMAP):
        text = _text_of_element(p)
        # 去掉两端空白，但保留空段落为空字符串（便于换行）
        paragraphs.append(text)
    return paragraphs

def docx_url_to_text(url, timeout=30):
    """
    从 URL 下载 docx 并返回去掉修订和批注后的纯文本（字符串）。
    包含：主文档、所有 headers、footers、footnotes、endnotes 内容（按文件顺序拼接）。
    如果下载或解析失败会抛出异常。
    """
    # 下载
    resp = requests.get(url, stream=True, timeout=timeout)
    resp.raise_for_status()
    data = resp.content

    # 读取 zip
    zf = zipfile.ZipFile(io.BytesIO(data))

    # 需要处理的文件：document.xml + headers/footers + footnotes/endnotes
    text_parts = []

    # 主文档
    if "word/document.xml" in zf.namelist():
        xml_bytes = zf.read("word/document.xml")
        paras = _extract_text_from_xml_bytes(xml_bytes)
        text_parts.extend(paras)

    # headers 和 footers（按文件名排序以尽量保证顺序一致）
    for name in sorted(n for n in zf.namelist() if n.startswith("word/header") and n.endswith(".xml")):
        xml_bytes = zf.read(name)
        paras = _extract_text_from_xml_bytes(xml_bytes)
        if paras:
            text_parts.append("")  # 分隔
            text_parts.extend(paras)

    for name in sorted(n for n in zf.namelist() if n.startswith("word/footer") and n.endswith(".xml")):
        xml_bytes = zf.read(name)
        paras = _extract_text_from_xml_bytes(xml_bytes)
        if paras:
            text_parts.append("")  # 分隔
            text_parts.extend(paras)

    # footnotes / endnotes
    for special in ("word/footnotes.xml", "word/endnotes.xml"):
        if special in zf.namelist():
            xml_bytes = zf.read(special)
            paras = _extract_text_from_xml_bytes(xml_bytes)
            if paras:
                text_parts.append("")  # 分隔
                text_parts.extend(paras)

    # 组合段落为字符串，段落之间用换行符分隔
    # 去掉开头结尾多余空行
    # 保留用户文档中空段落（会成为空行）
    result = "\n".join(text_parts).strip("\n")
    return result

