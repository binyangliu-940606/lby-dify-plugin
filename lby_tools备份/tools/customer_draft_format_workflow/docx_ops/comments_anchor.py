from lxml import etree
from .docx_zip_edit import edit_docx_zip

# WordprocessingML（docx 里 word/document.xml）使用的命名空间
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def _qn_w(local: str) -> str:
    """生成带命名空间的 tag 名：例如 local='p' -> '{...}p'（lxml 写法）"""
    return f"{{{W_NS}}}{local}"

def _get_paragraph_nodes(doc_root):
    """
    获取文档 body 下所有段落节点 w:p。
    注意：这里只抓 body 直接子孙里的段落（一般正文足够），不含页眉页脚等部位。
    """
    return doc_root.findall(".//w:body/w:p", namespaces=NS)

def _paragraph_text_and_t_nodes(p_node):
    """
    读取一个段落节点中的所有 w:t 文本节点，并拼接成“段落纯文本” full。
    同时返回每个 w:t 在 full 中对应的字符区间 spans：
      spans: List[(t_node, start_pos, end_pos)]
    其中 start_pos/end_pos 是该 w:t 文本在段落拼接字符串 full 里的位置范围。
    """
    t_nodes = p_node.findall(".//w:t", namespaces=NS)
    spans = []
    cur = 0
    full = []
    for t in t_nodes:
        txt = t.text or ""
        full.append(txt)
        spans.append((t, cur, cur + len(txt)))
        cur += len(txt)
    return "".join(full), spans

def _split_t_node(t_node, cut_pos: int):
    """
    在某个 w:t 节点内部按 cut_pos 切成左右两段：
      原 t_node 保留左半段文本
      新建一个兄弟 w:r（尽量复制 rPr 样式），里面放右半段 w:t
    用途：为了让 comment 的 start/end 能精确落在 run 边界上。
    """
    txt = t_node.text or ""
    left = txt[:cut_pos]
    right = txt[cut_pos:]

    # t_node 的父链上找到所属的 w:r（run）
    r = t_node.getparent()
    while r is not None and r.tag != _qn_w("r"):
        r = r.getparent()
    if r is None:
        # 找不到 run（结构异常）则不处理
        return

    # 原 w:t 改为左半段
    t_node.text = left

    # 新建 run，用来承载右半段文本，并尽量复制原 run 的样式（w:rPr）
    new_r = etree.Element(_qn_w("r"))
    rPr = r.find("w:rPr", namespaces=NS)
    if rPr is not None:
        # 深拷贝 rPr（通过 tostring/fromstring）
        new_r.append(etree.fromstring(etree.tostring(rPr)))

    # 新 run 下创建 w:t，并写入右半段
    new_t = etree.SubElement(new_r, _qn_w("t"))
    new_t.text = right

    # 将新 run 插入到原 run 后面（同级相邻）
    r.addnext(new_r)

def add_comment_precise(docx_path: str, out_path: str, paragraph_index: int, start: int, end: int, comment_id: int) -> dict:
    """
    在指定段落 paragraph_index 的 [start, end) 字符范围内精确插入：
      - w:commentRangeStart
      - w:commentRangeEnd
      - w:commentReference
    说明：
    - 通过重写 docx(zip) 的方式修改 word/document.xml，避免 zip 内 Duplicate name 问题。
    - start/end 是相对于“段落拼接纯文本”的字符下标。
    """
    def editor(files: dict):
        """
        edit_docx_zip 的回调编辑器：
        - files 是 zip 内文件的 bytes 映射（如 files['word/document.xml']）
        - 在此直接改写 files['word/document.xml'] 即可
        """
        if "word/document.xml" not in files:
            raise RuntimeError("docx 缺少 word/document.xml")

        # 解析 document.xml
        doc_root = etree.fromstring(files["word/document.xml"])

        # 获取所有段落节点
        ps = _get_paragraph_nodes(doc_root)
        if paragraph_index < 0 or paragraph_index >= len(ps):
            # 段落索引越界：标记失败，并把 xml 写回（维持结构一致）
            editor.ok = False  # type: ignore
            editor.error = "paragraph_index 越界"
            files["word/document.xml"] = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            return

        # 目标段落
        p = ps[paragraph_index]

        # --------
        # 切分边界：先切 end 再切 start（避免先切 start 后，end 的相对位置发生漂移）
        # --------
        full, spans = _paragraph_text_and_t_nodes(p)

        # 对 start/end 做安全裁剪，避免越界
        if end > len(full):
            end2 = len(full)
        else:
            end2 = end
        if start < 0:
            start2 = 0
        else:
            start2 = start

        # start/end 不合法直接失败
        if end2 <= start2:
            editor.ok = False  # type: ignore
            editor.error = "start/end 非法"
            files["word/document.xml"] = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            return

        # 1) 找到 end2 落在某个 w:t 的中间位置，则把该 w:t 切开，使 end2 恰好成为边界
        for t, a, b in spans:
            if a < end2 < b:
                _split_t_node(t, end2 - a)
                break

        # 切完 end 后，重新计算 spans（因为结构变了）
        full, spans = _paragraph_text_and_t_nodes(p)

        # 2) 同理处理 start2
        for t, a, b in spans:
            if a < start2 < b:
                _split_t_node(t, start2 - a)
                break

        # --------
        # 再次定位 start/end 边界所属的 run（w:r）
        # 要求：start2 必须等于某个 w:t 的起点 a；end2 必须等于某个 w:t 的终点 b
        # 这样才能把 commentRangeStart/End 插在 run 边界上
        # --------
        full, spans = _paragraph_text_and_t_nodes(p)
        start_r = None
        end_r = None
        for t, a, b in spans:
            if a == start2:
                # 从 t 向上找到所属 run
                r = t.getparent()
                while r is not None and r.tag != _qn_w("r"):
                    r = r.getparent()
                start_r = r
            if b == end2:
                # 注意：这里用 b==end2 来找 end 的 run（range end 放在该 run 后面）
                r = t.getparent()
                while r is not None and r.tag != _qn_w("r"):
                    r = r.getparent()
                end_r = r

        # 如果无法精确落在边界上，说明切分失败或段落结构特殊
        if start_r is None or end_r is None:
            editor.ok = False  # type: ignore
            editor.error = "无法精确锚定边界"
            files["word/document.xml"] = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            return

        # --------
        # 插入 comment 的三个关键节点：
        # 1) commentRangeStart：插在 start_r 前
        # 2) commentRangeEnd：插在 end_r 后
        # 3) commentReference：紧跟在 commentRangeEnd 后（通常放在一个 w:r 里）
        # --------
        crs = etree.Element(_qn_w("commentRangeStart"))
        crs.set(_qn_w("id"), str(comment_id))
        start_r.addprevious(crs)

        cre = etree.Element(_qn_w("commentRangeEnd"))
        cre.set(_qn_w("id"), str(comment_id))
        end_r.addnext(cre)

        # reference run：用于在文档中显示批注的小标记（引引用）
        ref_r = etree.Element(_qn_w("r"))
        ref = etree.SubElement(ref_r, _qn_w("commentReference"))
        ref.set(_qn_w("id"), str(comment_id))
        cre.addnext(ref_r)

        # 写回修改后的 document.xml
        files["word/document.xml"] = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        editor.ok = True  # type: ignore

    # 默认状态：失败，直到 editor 内部成功设置 ok=True
    editor.ok = False  # type: ignore
    editor.error = ""  # type: ignore

    # 执行 zip 内文件编辑：读取 docx_path -> 调用 editor -> 输出到 out_path
    edit_docx_zip(docx_path, out_path, editor)
    return {"ok": editor.ok, "error": editor.error}