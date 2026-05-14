import re
import json
import zipfile

from typing import Dict, Any, List, Tuple, Optional
from lxml import etree

from collections.abc import Generator
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        payload_json = json.loads(tool_parameters["payload_json"])
        orig_path = payload_json["orig_path"]
        normalized_table_keys = payload_json["normalized_table_keys"]

        yield self.create_json_message({
            "return_data": extract_tables_from_orig_ooxml(orig_path, normalized_table_keys),
        })

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def _qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"

# 表题匹配：兼容 Table/Tab/表/Supplementary table/Supplemental table 等
TAB_TITLE_PAT = re.compile(
    r"""(?ix)^\s*
    (Supplementary\s+table|Supplemental\s+table|Table|Tab\.?|表)
    \s*
    (S?\s*\d+)
    \s*[:：.\-]?\s*
    (.*)$
    """
)

def _norm_table_key(num_str: str) -> str:
    """
    归一化 key：
    - S1 / s 1 -> Table S1
    - 1 -> Table 1
    """
    s = (num_str or "").replace(" ", "")
    if s.upper().startswith("S"):
        return f"Table S{s[1:]}"
    return f"Table {s}"

def _node_text(node) -> str:
    """抽取 w:p 或任意节点下所有 w:t 文本并拼接"""
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))

def _extract_tbl_grid(tbl_node) -> List[List[str]]:
    """
    从 w:tbl 节点提取二维表格文本。
    - 按 w:tr 行
    - 按 w:tc 单元格
    - 单元格文本：拼接其内部 w:t
    """
    grid: List[List[str]] = []
    for tr in tbl_node.findall(".//w:tr", namespaces=NS):
        row = []
        tcs = tr.findall("./w:tc", namespaces=NS)
        for tc in tcs:
            txt = _node_text(tc).strip()
            # docx 表格里同一行可能出现“重复 cell 对象引用”（合并单元格），这里保守直接写文本
            row.append(txt)
        grid.append(row)

    # 清理：去掉全空行（有些文档会带空行）
    def is_empty_row(r: List[str]) -> bool:
        return all((c or "").strip() == "" for c in r)

    grid = [r for r in grid if not is_empty_row(r)]
    return grid

def _is_blank_paragraph(p_node) -> bool:
    return _node_text(p_node).strip() == ""

def _is_table_title_paragraph(p_node) -> Optional[Tuple[str, str]]:
    """
    判断是否为表题段落。
    返回 (table_key, title_rest) 或 None
    """
    txt = _node_text(p_node).strip()
    m = TAB_TITLE_PAT.match(txt)
    if not m:
        return None
    num = m.group(2)
    title_rest = (m.group(3) or "").strip()
    return _norm_table_key(num), title_rest

def extract_tables_from_orig_ooxml(orig_path: str, normalized_table_keys: List[str]) -> Dict[str, Any]:
    """
    稳定版：基于 OOXML，按 document.xml 中 w:body 的真实顺序混合遍历 w:p 与 w:tbl。
    逻辑：
    1) 找到表题段落（w:p，匹配 TAB_TITLE_PAT）
    2) 从该表题段落之后向后找第一个 w:tbl 作为该表
    3) 表注：从该表格之后开始，收集连续段落文本直到：
       - 遇到空行累计>=2
       - 遇到下一张表题段落
       - 或遇到明显“章节标题”（这里先不做复杂标题识别，你可后续扩展）
    """
    want = set(normalized_table_keys or [])
    tables_out: Dict[str, Any] = {}
    missing: List[str] = []

    with zipfile.ZipFile(orig_path, "r") as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(".//w:body", namespaces=NS)
    if body is None:
        return {"tables_from_orig": {}, "tables_missing_in_orig": list(want)}

    children = list(body)  # w:p/w:tbl/...
    i = 0
    while i < len(children):
        node = children[i]
        if node.tag == _qn("p"):
            tit = _is_table_title_paragraph(node)
            if not tit:
                i += 1
                continue

            table_key, title_rest = tit
            # 只抽取属于 want 的表
            if table_key not in want or table_key in tables_out:
                i += 1
                continue

            # 向后找第一个 w:tbl
            j = i + 1
            tbl_node = None
            while j < len(children):
                if children[j].tag == _qn("tbl"):
                    tbl_node = children[j]
                    break
                # 如果中途又遇到下一个表题，说明该表没有表格
                if children[j].tag == _qn("p") and _is_table_title_paragraph(children[j]):
                    break
                j += 1

            grid = _extract_tbl_grid(tbl_node) if tbl_node is not None else []

            # 表注：从表格后面开始连续收集段落
            note_lines: List[str] = []
            k = (j + 1) if tbl_node is not None else (i + 1)
            empty_count = 0
            while k < len(children):
                n2 = children[k]
                if n2.tag == _qn("p"):
                    if _is_table_title_paragraph(n2):
                        break
                    t2 = _node_text(n2).strip()
                    if t2 == "":
                        empty_count += 1
                        if empty_count >= 2:
                            break
                    else:
                        empty_count = 0
                        note_lines.append(t2)
                elif n2.tag == _qn("tbl"):
                    # 遇到另一个表格，通常表示表注结束（或文档结构异常），这里停止
                    break
                k += 1

            note_text = "\n".join(note_lines).strip()

            tables_out[table_key] = {
                "title": title_rest,
                "note": note_text,
                "grid": grid,
                "source": {"from": "orig_ooxml", "title_index": i, "tbl_index": j if tbl_node is not None else None}
            }

            # i 跳到 k（提高效率）
            i = k
            continue

        i += 1

    for k in want:
        if k not in tables_out:
            missing.append(k)

    return {"tables_from_orig": tables_out, "tables_missing_in_orig": missing}