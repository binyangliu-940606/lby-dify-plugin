

import json
import zipfile

from lxml import etree
from typing import Dict, Any, List
from collections.abc import Generator
from docx import Document
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        docx_path = tool_parameters["docx_path"]

        yield self.create_json_message({
            "return_data": extract_docx_paras_and_tables_ooxml(docx_path),
        })



W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def _node_text(node) -> str:
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))

def _extract_tbl_grid(tbl_node, max_rows: int, max_cols: int) -> List[List[str]]:
    grid_full = []
    for tr in tbl_node.findall("./w:tr", namespaces=NS):
        row = []
        for tc in tr.findall("./w:tc", namespaces=NS):
            row.append(_node_text(tc).strip())
        if any(c.strip() for c in row):
            grid_full.append(row)

    # 裁剪（确定性，控制 token）
    grid = []
    for r in grid_full[:max_rows]:
        rr = r[:max_cols] + [""] * max(0, max_cols - len(r))
        grid.append(rr[:max_cols])
    return grid

def _collect_paragraphs(root) -> List[Dict[str, Any]]:
    paras = []
    # 注意：这里的 p_index 是“该 part 内的段落序号”，不是全局 document 段落序号
    for i, p in enumerate(root.findall(".//w:p", namespaces=NS)):
        txt = _node_text(p).strip()
        if txt:
            paras.append({"p": i, "text": txt})
    return paras

def extract_docx_paras_and_tables_ooxml(orig_path: str, max_rows: int = 80, max_cols: int = 30) -> Dict[str, Any]:
    """
    增强版：从多个 part 抽取 table/paragraph，解决表格在 header/footer/footnotes 等位置导致抓不到的问题。
    输出：
      - orig_paragraphs: [{part, p, text}]
      - orig_tables: [{tbl_index, part, grid, n_rows_hint, n_cols_hint, preview}]
    """
    # 常见需要扫描的 parts（可按需扩展）
    target_parts_prefix = [
        "word/document.xml",
        "word/header",      # header1.xml...
        "word/footer",      # footer1.xml...
        "word/footnotes.xml",
        "word/endnotes.xml",
    ]

    orig_paragraphs: List[Dict[str, Any]] = []
    orig_tables: List[Dict[str, Any]] = []
    tbl_global_index = 0

    with zipfile.ZipFile(orig_path, "r") as z:
        names = z.namelist()

        def should_scan(name: str) -> bool:
            if not name.endswith(".xml"):
                return False
            for pfx in target_parts_prefix:
                if name == pfx or name.startswith(pfx):
                    return True
            return False

        for name in names:
            if not should_scan(name):
                continue

            try:
                xml = z.read(name)
                root = etree.fromstring(xml)

                # 段落：用于 LLM 抽 title/note
                paras = _collect_paragraphs(root)
                for p in paras:
                    p["part"] = name
                orig_paragraphs.extend(paras)

                # 表格：用于 LLM 选择 tbl_index
                tbls = root.findall(".//w:tbl", namespaces=NS)
                for t in tbls:
                    grid = _extract_tbl_grid(t, max_rows=max_rows, max_cols=max_cols)
                    # 预览：前3行前6列
                    preview_lines = []
                    for r in grid[:3]:
                        preview_lines.append(" | ".join([c for c in r[:6] if c is not None]))
                    preview = "\n".join(preview_lines).strip()

                    # 粗略规模提示（用裁剪后不准确，但足够让 LLM 判断）
                    n_rows_hint = len(grid)
                    n_cols_hint = max((len(r) for r in grid), default=0)

                    orig_tables.append({
                        "tbl_index": tbl_global_index,
                        "part": name,
                        "grid": grid,
                        "n_rows_hint": n_rows_hint,
                        "n_cols_hint": n_cols_hint,
                        "preview": preview
                    })
                    tbl_global_index += 1

            except Exception as e:
                # 某些 xml 可能不是标准 wordprocessingml（比如 settings），忽略即可
                continue

    return {"orig_paragraphs": orig_paragraphs, "orig_tables": orig_tables}