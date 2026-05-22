

import json
import requests
import re
from typing import Dict, Any, List, Tuple
from docx import Document
# from collections.abc import Generator
# from dify_plugin import Tool
# from dify_plugin.entities.tool import ToolInvokeMessage

# class LbyToolsTool(Tool):
#     def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
#         #获取参数如文件url
#         payload_json = json.loads(tool_parameters["payload_json"])
#         ref_citations = payload_json["ref_citations"]
#         ref_entries = payload_json["ref_entries"]

#         yield self.create_json_message({
#             "return_data": pubmed_lookup_build_pairs_and_comments(ref_citations,ref_entries),
#         })


CIT_PAT = re.compile(
    r"(\[\s*\d[\d\s,，、;；/.\-–—]*\s*\]|\u3010\s*\d[\d\s,，、;；/.\-–—]*\s*\u3011)"
)

_DOI_RE = re.compile(r"\b10\.\d{4,9}/[^\s\"<>]+", re.I)

CHINESE_CHAR_PAT = re.compile(r"[\u4e00-\u9fff]")

def _is_chinese_ref(text: str) -> bool:
    return bool(CHINESE_CHAR_PAT.search(text or ""))

# def _clean_query(text: str) -> str:
#     """
#     对参考文献文本做简单清洗，作为 PubMed 查询 query。
#     只做确定性去噪，不做模型推理。
#     """
#     t = (text or "").strip()
#     # 去掉多余空格
#     t = re.sub(r"\s+", " ", t)
#     # 去掉明显的页码/卷期等噪声（保守）
#     t = re.sub(r"\b\d{4}\b", "", t)  # 年份去掉（可选）
#     t = t.strip()
#     # query 太长会影响 esearch，截断
#     return t[:240]

def _clean_query(ref: str) -> str:
    """
    传入一条原始参考文献字符串，输出更适合 PubMed ESearch 的 query。
    规则（优先级）：
    1) 若存在合法 DOI：直接用 DOI（命中率最高）
    2) 否则：用“Title（引号包裹） + 第一作者姓氏”作为查询
    """
    if not ref:
        return ""

    s = ref.strip()

    # 1) 去掉开头编号 [37]
    # s = re.sub(r"^\s*\[\d+\]\s*", "", s)

    # 2) 标准化空白
    s = re.sub(r"\s+", " ", s)

    # 3) 抽 DOI（若存在）
    m = _DOI_RE.search(s)
    if m:
        doi = m.group(0).rstrip(").,;")
        return doi

    # 4) 无 DOI：抽第一作者姓氏（逗号前）
    first_author = ""
    m = re.match(r"\s*([^,]+),", s)
    if m:
        first_author = m.group(1).strip()

    # 5) 抽 Title：通常在最后一个作者分隔（", and X," 或 ", et al.,"）之后，
    #    到期刊名之前（期刊名前一般是 ". " 或 ", "）
    # 尽量从第一个“文章标题起点”切片：
    title_start = None
    for pat in (r",\s*et al\.,\s*", r",\s*and\s+[^,]+,\s*", r",\s*[^,]+,\s*"):
        mm = re.search(pat, s)
        if mm:
            title_start = mm.end()
            break

    tail = s[title_start:] if title_start is not None else s

    # 期刊名通常在标题后的第一个 ". " 之后开始
    # 例如："... osteoarthritis. Osteoarthritis Cartilage, 2015."
    title = tail
    mm = re.search(r"\.\s+[A-Z].*?,\s*\d{4}\b", tail)
    if mm:
        title = tail[: mm.start()].strip()

    # 清理 title 里可能的残留
    title = title.strip().strip(" .;,")

    # 6) 组装 query：标题 + 第一作者（尽量提高召回）
    if title and first_author:
        return f"\"{title}\" {first_author}[Author]"
    if title:
        return f"\"{title}\""
    if first_author:
        return f"{first_author}[Author]"
    return ""

def _pubmed_esearch_pmids(query: str, retmax: int = 3, timeout: int = 20) -> List[str]:
    """
    PubMed eutils esearch：返回 PMID 列表（最多 retmax）
    """
    if not query:
        return []
    
    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    
    params = {
        "db": "pubmed",
        "term": query,
        "retmode": "json",
        "retmax": str(retmax)
    }

    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    j = r.json()
    ids = j.get("esearchresult", {}).get("idlist", []) or []
    return [str(x) for x in ids if str(x).isdigit()]


def pubmed_lookup_build_pairs_and_comments(
    ref_citations: List[Dict[str, Any]],
    ref_entries: Dict[str, str],
    ref_start
) -> Dict[str, Any]:
    """
    输入：
      - ref_citations: 正文中出现的 [1]/(1)/【1】列表
      - ref_entries: 参考文献列表 index->text

    输出：
      - pmid_normalize_pairs: 用于 Node14 替换（kind="pmid"）
      - pmid_comment_tasks: 中文/查不到 PMID 的批注任务（anchor 用原始引用标记 raw）
    """
    pmid_pairs = []
    comment_tasks = []

    # 为避免重复查询：index -> 结果缓存
    cache: Dict[str, Dict[str, Any]] = {}
    
    def split_citation_numbers(cit: str) -> List[int]:
        """
        输入形如: [1/2/100], [1,2,3], [1-3], 【1，2；3-5/7】...
        输出数字列表（仅“切割”，不展开区间：1-3 -> [1,3]）
        """
        # 去掉外层括号 [] 或 【】
        inner = re.sub(r'^\s*(?:\[|【)\s*|\s*(?:\]|】)\s*$', '', cit)

        # 按任意非数字分隔符切割
        nums = [int(x) for x in re.split(r"\D+", inner) if x]
        return nums

    for cit in ref_citations or []:
        
        p_hint = cit.get("p")
        if ref_start:
            if p_hint>ref_start:
                continue

        for idx in split_citation_numbers(cit.get("raw")):

            # idx = cit.get("index")
            raw = cit.get("raw")  # 原始引用标记，如 [12]

            if not idx or not raw:
                continue

            if idx not in cache:
                ref_text = (ref_entries or {}).get(str(idx), "").strip()
                if not ref_text:
                    cache[idx] = {"status": "missing_entry"}
                elif _is_chinese_ref(ref_text):
                    cache[idx] = {"status": "chinese", "ref_text": ref_text}
                else:
                    # 英文：查 pubmed
                    query = _clean_query(ref_text)
                    try:
                        pmids = _pubmed_esearch_pmids(query, retmax=3)
                    except Exception as e:
                        pmids = []
                    if pmids:
                        cache[idx] = {"status": "ok", "pmids": pmids, "ref_text": ref_text}
                    else:
                        cache[idx] = {"status": "not_found", "ref_text": ref_text}

            info = cache[idx]
            st = info["status"]

            if st == "ok":
                pmids = info["pmids"]
                rep = "(" + "; ".join([f"PMID: {x}" for x in pmids]) + ")"
                # 替换对：把 raw（例如 [12]）替换成 rep
                pmid_pairs.append({
                    "kind": "pmid",
                    "find": raw,
                    "replace": rep,
                    "idx": idx,
                })
            elif st == "chinese":
                comment_tasks.append({
                    "anchor": raw,
                    "p_hint": p_hint,
                    "text": f"此引用为中文文献（编号 {idx}），未查询 PMID。"
                })
            elif st == "not_found":
                comment_tasks.append({
                    "anchor": raw,
                    "p_hint": p_hint,
                    "text": f"未在 PubMed 查询到 PMID（编号 {idx}）。"
                })
            elif st == "missing_entry":
                comment_tasks.append({
                    "anchor": raw,
                    "p_hint": p_hint,
                    "text": f"未在参考文献列表中找到该编号条目（编号 {idx}）。"
                })

    # 去重 pmid_pairs（同一个 raw 多次出现会导致重复替换开销）
    uniq = {}
    for p in pmid_pairs:
        key = (p["kind"], p["find"], p["replace"])
        uniq[key] = p
    pmid_pairs = list(uniq.values())

    return {
        "pmid_normalize_pairs": pmid_pairs,
        "pmid_comment_tasks": comment_tasks
    }
    # return pmid_pairs,comment_tasks