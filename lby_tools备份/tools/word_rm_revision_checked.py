from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

import io
import zipfile
import requests
from lxml import etree
from urllib.parse import urlparse
import os

class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        url = tool_parameters["query"]
        color = tool_parameters["color"]
        file_name = tool_parameters["file_name"]

        #获取字节流，文件名，mime_type
        out_bytes, name, mime = process_docx_url_accept_deletions(url,color)

        #因为byte字节流传递出错，需要转换为base64，以供工具流中下一节点使用
        # b64_str = base64.b64encode(out_bytes).decode("utf-8")

        # processed_word = {
        #     "data": b64_str,
        #     "mime_type": mime,
        #     "name": name,
        # }

        # yield self.create_json_message({
        #     "result": {
        #         "processed_word":processed_word,
        #     }
        # })

        # ============== 关键改动在这里 ==============
        # 使用 create_blob_message 直接返回文件字节流
        yield self.create_blob_message(
            blob=out_bytes,          # 直接传入 bytes，不需要 base64 编码
            meta={
                "mime_type": mime,   # 告诉 Dify 这是什么类型的文件
                "filename": file_name,
            },
        )


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def q(local: str) -> str:   # element Clark notation
    return f"{{{W_NS}}}{local}"

def qa(local: str) -> str:  # attribute Clark notation (e.g. w:val)
    return f"{{{W_NS}}}{local}"

def remove_nodes(root, tags: list[str]):
    for t in tags:
        for el in list(root.iter(q(t))):
            p = el.getparent()
            if p is not None:
                p.remove(el)

def ensure_rpr(run_el):
    """确保 <w:r> 下存在 <w:rPr> 并返回它"""
    rpr = run_el.find(q("rPr"))
    if rpr is None:
        rpr = etree.Element(q("rPr"))
        run_el.insert(0, rpr)
    return rpr

def mark_inserted_text_style(root, color: str | None):
    """
    将所有修订“新增内容”(含：纯新增、替换修改后的新文字) 标记为：
    - color='yellow' -> 黄色高亮
    - color='red'    -> 红色字体
    """
    if not color:
        return

    mode = color.strip().lower()
    if mode not in ("yellow", "red"):
        raise ValueError("color must be 'yellow' or 'red'")

    # 关键点：无论是“新增修订”还是“修改修订(替换的新内容)”，新内容都在 <w:ins>
    for ins in root.iter(q("ins")):
        for run in ins.iter(q("r")):
            # 仅处理包含真实文本的 run，避免把图片、域代码等也染色
            if run.find(q("t")) is None:
                continue

            rpr = ensure_rpr(run)

            if mode == "yellow":
                hl = rpr.find(q("highlight"))
                if hl is None:
                    hl = etree.Element(q("highlight"))
                    rpr.append(hl)
                hl.set(qa("val"), "yellow")
            else:
                c = rpr.find(q("color"))
                if c is None:
                    c = etree.Element(q("color"))
                    rpr.append(c)
                c.set(qa("val"), "FF0000")

def accept_delete_revisions_in_xml(xml_bytes: bytes, color: str | None = None) -> bytes:
    parser = etree.XMLParser(recover=True)
    root = etree.fromstring(xml_bytes, parser=parser)

    # 1) 只接受“删除修订”：移除旧内容
    remove_nodes(root, ["del", "delText", "delInstrText"])

    # 2) 给所有 <w:ins>（含新增 + 替换修改后的新内容）打上高亮/红字
    mark_inserted_text_style(root, color=color)

    # 3) 清理修订属性变更节点（按你原逻辑保留）
    remove_nodes(root, ["pPrChange", "rPrChange", "tblPrChange", "trPrChange", "tcPrChange"])

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

def disable_track_revisions_in_settings(settings_xml: bytes) -> bytes:
    parser = etree.XMLParser(recover=True)
    root = etree.fromstring(settings_xml, parser=parser)

    for el in list(root.iter(q("trackRevisions"))):
        p = el.getparent()
        if p is not None:
            p.remove(el)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

def accept_delete_revisions_in_docx(docx_bytes: bytes, color: str | None = None) -> bytes:
    zin = zipfile.ZipFile(io.BytesIO(docx_bytes), "r")
    out_buf = io.BytesIO()
    zout = zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED)

    targets = {"word/document.xml", "word/footnotes.xml", "word/endnotes.xml"}
    for name in zin.namelist():
        if (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
            targets.add(name)

    for item in zin.infolist():
        data = zin.read(item.filename)

        if item.filename in targets:
            try:
                data = accept_delete_revisions_in_xml(data, color=color)
            except Exception:
                pass

        if item.filename == "word/settings.xml":
            try:
                data = disable_track_revisions_in_settings(data)
            except Exception:
                pass

        zout.writestr(item, data)

    zin.close()
    zout.close()
    return out_buf.getvalue()

def process_docx_url_accept_deletions(url: str, color: str | None = None):
    r = requests.get(url, timeout=120)
    r.raise_for_status()

    out_bytes = accept_delete_revisions_in_docx(r.content, color=color)

    path = urlparse(url).path
    name = os.path.basename(path) or "manuscript-PMID(Marked).docx"
    if not name.lower().endswith(".docx"):
        name = "manuscript-PMID(Marked).docx"

    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return out_bytes, name, mime
