from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from collections.abc import Generator
from typing import Any

import os
import re
import json
import requests
import tempfile
from docx import Document
from docx.oxml.ns import qn

class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        payload_json = json.loads(tool_parameters["payload_json"])

        # orig_paras = payload_json["orig_paras"]
        iter2_list = payload_json["iter2_list"]
        orgin_docx_url = payload_json["orgin_docx_url"]

        tmp_path = download_to_temp(orgin_docx_url)
        para_id_to_text = build_rewrite_map(iter2_list)

        out_bytes = apply_rewrites_to_docx(tmp_path, para_id_to_text)
        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        filename = "final.docx"

        yield self.create_blob_message(
                blob=out_bytes,          # 直接传入 bytes，不需要 base64 编码
                meta={
                    "mime_type": mime_type,   # 告诉 Dify 这是什么类型的文件
                    "filename": filename,
                },
            )

# XML 1.0 允许的字符范围，去掉不允许的控制字符/NULL等
_ILLEGAL_XML_RE = re.compile(
    r"[\x00-\x08\x0B\x0C\x0E-\x1F]"  # C0控制字符(除\t \n \r)
)

def sanitize_for_docx_xml(s: str) -> str:
    if s is None:
        return ""
    # 1) 去掉非法控制字符
    s = _ILLEGAL_XML_RE.sub("", s)
    # 2) 保险起见：去掉 Unicode surrogate（有些模型/流程可能产生）
    s = s.encode("utf-8", "surrogatepass").decode("utf-8", "ignore")
    return s


def build_rewrite_map(iter2_list):
    """
    iter2_list: list[dict], each has 'para_id' and 'rewritten_text'
    Return: dict {para_id: rewritten_text}
    """
    mp = {}
    for item in iter2_list or []:
        pid = item.get("para_id")
        txt = (item.get("rewritten_text") or "").strip()
        if pid and txt:
            mp[pid] = txt
    return mp

def reset_run_color_to_default(run):
    """
    将run的字体颜色恢复为默认（移除<w:color>设置）
    """
    r = run._element
    rPr = r.rPr
    if rPr is None:
        return

    color = rPr.find(qn('w:color'))
    if color is not None:
        rPr.remove(color)

def reset_paragraph_run_colors_to_default(paragraph):
    for run in paragraph.runs:
        # 只要存在颜色设置就移除；没有则跳过
        rPr = run._element.rPr
        if rPr is None:
            continue
        if rPr.find(qn('w:color')) is not None:
            reset_run_color_to_default(run)

def redistribute_text_to_docx_runs(new_text, runs):
    """
    runs: list[docx.text.run.Run]
    原 run 样式保留，仅替换 run.text
    """
    if not runs:
        return

    new_text = sanitize_for_docx_xml(new_text)

    old_lengths = [len(r.text or "") for r in runs]
    total_old = sum(old_lengths)

    if total_old == 0:
        runs[0].text = new_text
        for r in runs[1:]:
            r.text = ""
        return

    consumed = 0
    for i, (r, l) in enumerate(zip(runs, old_lengths)):
        if i == len(runs) - 1:
            part = new_text[consumed:]
        else:
            take = round(len(new_text) * (l / total_old))
            part = new_text[consumed:consumed + take]

        r.text = sanitize_for_docx_xml(part)
        consumed += len(part)

def apply_rewrites_to_docx(input_docx_path, para_id_to_text):
    """
    读取 input_docx_path -> 回填 -> 保存到临时docx -> 读取bytes(blob) -> 删除临时文件(输出+输入) -> 返回bytes
    """
    doc = Document(input_docx_path)

    for idx, p in enumerate(doc.paragraphs, start=1):
        pid = f"p_{idx:05d}"
        new_text = para_id_to_text.get(pid)
        if not new_text:
            continue

        redistribute_text_to_docx_runs(new_text, p.runs)
        reset_paragraph_run_colors_to_default(p)

    out_tmp_path = None
    try:
        fd, out_tmp_path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)  # 避免 Windows 文件占用

        doc.save(out_tmp_path)

        with open(out_tmp_path, "rb") as f:
            blob = f.read()

        return blob

    finally:
        # 删除输出临时文件
        if out_tmp_path and os.path.exists(out_tmp_path):
            try:
                os.remove(out_tmp_path)
            except Exception:
                pass

        # 删除输入文件（input_docx_path）
        if input_docx_path and os.path.exists(input_docx_path):
            try:
                os.remove(input_docx_path)
            except Exception:
                pass


def download_to_temp(url: str, suffix=".docx", headers=None, timeout=60) -> str:
    headers = headers or {}
    r = requests.get(url, stream=True, headers=headers, timeout=timeout)
    r.raise_for_status()
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    with open(path, "wb") as f:
        for chunk in r.iter_content(chunk_size=1024 * 1024):
            if chunk:
                f.write(chunk)
    return path
