from typing import Dict, Any, List
from docx import Document

def find_first_occurrences(docx_path: str, keys: List[str]) -> Dict[str, Dict[str, int]]:
    """
    在 docx 中查找每个 key 第一次出现的位置：
    返回：{key: {"p": 段落索引, "start": 起始字符, "end": 结束字符}}
    """
    doc = Document(docx_path)
    res: Dict[str, Dict[str, int]] = {}

    # 为了效率：把 keys 按长度降序，避免短 key 误命中（可选）
    keys_sorted = sorted(set(keys), key=len, reverse=True)

    for p_i, p in enumerate(doc.paragraphs):
        text = p.text or ""
        if not text:
            continue
        for k in keys_sorted:
            if k in res:
                continue
            idx = text.find(k)
            if idx != -1:
                res[k] = {"p": p_i, "start": idx, "end": idx + len(k)}

        if len(res) == len(keys_sorted):
            break

    return res