

import json

import re
from typing import Dict, Any, List, Tuple
from docx import Document
from collections.abc import Generator
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from .pubmed_lookup_build_pairs_and_comments import pubmed_lookup_build_pairs_and_comments


class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        docx_path = tool_parameters["docx_path"]

        ref_info = extract_reference_section_and_citations(docx_path),
        pmid_info = pubmed_lookup_build_pairs_and_comments(ref_info[0]["ref_citations"],ref_info[0]["ref_entries"],ref_info[0]["ref_start"])

        yield self.create_json_message({
            "ref_citations": ref_info[0]["ref_citations"],
            "ref_entries": ref_info[0]["ref_entries"],
            "pmid_normalize_pairs": pmid_info["pmid_normalize_pairs"],
            "pmid_comment_tasks": pmid_info["pmid_comment_tasks"],
        })

# 编号制引用： [12] (12) 【12】
CIT_PAT = re.compile(r"(\[\s*\d+\s*\]|\(\s*\d+\s*\)|【\s*\d+\s*】)")

# 参考文献区标题（常见）
REF_HEAD_PAT = re.compile(r"(?i)^\s*(references|reference|参考文献)\s*$")

# 参考文献条目起始（尽量兼容）
# 例如: [1] xxxx   1. xxxx   【1】xxxx
REF_ITEM_PAT = re.compile(r"^\s*(\[\s*(\d+)\s*\]|【\s*(\d+)\s*】|(\d+)\.)\s*(.+)$")

def _norm_cit_mark(mark: str) -> Tuple[str, str]:
    """
    将引用标记标准化，返回 (kind, index_str)
    kind: bracket/paren/cn_bracket
    """
    m = re.sub(r"\s+", "", mark)
    idx = re.sub(r"\D", "", m)
    if m.startswith("["):
        return "bracket", idx
    if m.startswith("("):
        return "paren", idx
    if m.startswith("【"):
        return "cn_bracket", idx
    return "unknown", idx

def extract_reference_section_and_citations(sub_docx_path: str) -> Dict[str, Any]:
    """
    1) 提取正文中所有编号制引用出现（用于替换）
    2) 定位 References 区并解析参考文献列表（index -> text）
    """
    doc = Document(sub_docx_path)
    paras = [p.text or "" for p in doc.paragraphs]

    # 1) 正文编号引用（全篇扫描）
    ref_citations: List[Dict[str, Any]] = []
    for p_i, text in enumerate(paras):
        for m in CIT_PAT.finditer(text):
            raw = m.group(1)
            kind, idx = _norm_cit_mark(raw)
            if idx:
                ref_citations.append({
                    "p": p_i,
                    "raw": raw,
                    "kind": kind,
                    "index": idx
                })

    # 2) 定位参考文献区起点
    ref_start = None
    for i, t in enumerate(paras):
        if REF_HEAD_PAT.match(t.strip()):
            ref_start = i
            break

    # 3) 解析参考文献条目
    ref_entries: Dict[str, str] = {}
    if ref_start is not None:
        # 从标题下一段开始
        cur_idx = None
        buf = []
        for i in range(ref_start + 1, len(paras)):
            line = paras[i].strip()
            if not line:
                continue

            m = REF_ITEM_PAT.match(line)
            if m:
                # flush previous
                if cur_idx is not None and buf:
                    ref_entries[str(cur_idx)] = " ".join(buf).strip()
                # new item
                cur_idx = m.group(2) or m.group(3) or m.group(4)
                content = (m.group(5) or "").strip()
                buf = [content] if content else []
            else:
                # continuation line
                if cur_idx is not None:
                    buf.append(line)

        # flush last
        if cur_idx is not None and buf:
            ref_entries[str(cur_idx)] = " ".join(buf).strip()

    return {
        "ref_citations": ref_citations,
        "ref_entries": ref_entries,
        "ref_start": ref_start
    }
    # return ref_citations, ref_entries