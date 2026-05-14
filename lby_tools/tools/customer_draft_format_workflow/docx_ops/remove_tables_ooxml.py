import re
from lxml import etree
from .docx_zip_edit import edit_docx_zip

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def _qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"

TABLE_TITLE_PAT = re.compile(r"(?i)^\s*(Table|Supplementary\s+table|Supplemental\s+table|表)\b")

def _node_text(node):
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))

def remove_table_blocks(docx_path: str, out_path: str) -> dict:
    """
    删除次稿中所有表块（表题段落 + 表格 + 表注段落）
    使用 edit_docx_zip 重写 docx，避免 zip Duplicate name。
    """
    def editor(files: dict):
        doc_root = etree.fromstring(files["word/document.xml"])
        body = doc_root.find(".//w:body", namespaces=NS)
        children = list(body)

        to_remove = set()
        i = 0
        while i < len(children):
            node = children[i]
            if node.tag == _qn_w("p"):
                txt = _node_text(node).strip()
                if TABLE_TITLE_PAT.match(txt):
                    # 删除表题
                    to_remove.add(i)

                    # 删除紧随其后的表格
                    j = i + 1
                    while j < len(children) and children[j].tag == _qn_w("tbl"):
                        to_remove.add(j)
                        j += 1

                    # 删除表注：直到空行累计>=2 或遇到下一个表题
                    empty_count = 0
                    while j < len(children) and children[j].tag == _qn_w("p"):
                        t2 = _node_text(children[j]).strip()
                        if TABLE_TITLE_PAT.match(t2):
                            break
                        if t2 == "":
                            empty_count += 1
                            to_remove.add(j)
                            j += 1
                            if empty_count >= 2:
                                break
                        else:
                            empty_count = 0
                            to_remove.add(j)
                            j += 1

                    i = j
                    continue
            i += 1

        for idx in sorted(to_remove, reverse=True):
            body.remove(children[idx])

        files["word/document.xml"] = etree.tostring(
            doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )

    edit_docx_zip(docx_path, out_path, editor)
    return {"ok": True}