from lxml import etree
from datetime import datetime

from .docx_zip_edit import edit_docx_zip

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS_URI = "http://schemas.openxmlformats.org/package/2006/relationships"
NSMAP_W = {"w": W_NS}


def _qn_w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _ensure_comments_part(files: dict):
    """确保 word/comments.xml 存在"""
    if "word/comments.xml" not in files:
        root = etree.Element(_qn_w("comments"), nsmap=NSMAP_W)
        files["word/comments.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def _ensure_comments_relationship(files: dict):
    """确保 document.xml.rels 里存在 comments 关系"""
    rel_path = "word/_rels/document.xml.rels"
    if rel_path not in files:
        # 极少数 docx 可能缺 rels，这里简单创建
        root = etree.Element(f"{{{REL_NS_URI}}}Relationships", nsmap={None: REL_NS_URI})
        files[rel_path] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    rel_root = etree.fromstring(files[rel_path])

    ns = {"rel": REL_NS_URI}
    for r in rel_root.findall("rel:Relationship", namespaces=ns):
        if (r.get("Type") or "").endswith("/comments"):
            return

    # 生成新 rId
    rids = []
    for r in rel_root.findall("rel:Relationship", namespaces=ns):
        rid = r.get("Id", "")
        if rid.startswith("rId"):
            try:
                rids.append(int(rid[3:]))
            except:
                pass
    new_rid = f"rId{max(rids) + 1 if rids else 1}"

    new_rel = etree.SubElement(rel_root, f"{{{REL_NS_URI}}}Relationship")
    new_rel.set("Id", new_rid)
    new_rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
    new_rel.set("Target", "comments.xml")

    files[rel_path] = etree.tostring(rel_root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def _get_next_comment_id(comments_root) -> int:
    max_id = -1
    for c in comments_root.findall("w:comment", namespaces=NSMAP_W):
        cid = c.get(_qn_w("id"))
        try:
            max_id = max(max_id, int(cid))
        except:
            continue
    return max_id + 1


def create_comment_only(docx_path: str, out_path: str, comment_text: str, author: str = "dify") -> dict:
    """
    只创建 comment（写 comments.xml + rels），不修改 document.xml。
    采用重写 zip 的方式，避免 Duplicate name 警告。
    """
    def editor(files: dict):
        _ensure_comments_part(files)
        _ensure_comments_relationship(files)

        comments_root = etree.fromstring(files["word/comments.xml"])
        new_id = _get_next_comment_id(comments_root)

        c = etree.SubElement(comments_root, _qn_w("comment"))
        c.set(_qn_w("id"), str(new_id))
        c.set(_qn_w("author"), author)
        c.set(_qn_w("date"), datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))

        p = etree.SubElement(c, _qn_w("p"))
        r = etree.SubElement(p, _qn_w("r"))
        t = etree.SubElement(r, _qn_w("t"))
        t.text = comment_text

        files["word/comments.xml"] = etree.tostring(comments_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        # 把 comment_id 放到 files dict 里供外部读取不方便，所以返回值由外部再读一遍会很麻烦
        # 这里采用闭包变量
        editor.comment_id = new_id  # type: ignore

    editor.comment_id = None  # type: ignore
    edit_docx_zip(docx_path, out_path, editor)
    return {"ok": True, "comment_id": editor.comment_id}