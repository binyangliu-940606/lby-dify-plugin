import os
import json
import re

from collections.abc import Generator
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from typing import Dict, Any, List, Union, Optional, Tuple
from .apply_changes import apply_all_changes
from .blue_highlight_references_v3 import blue_highlight_references_v3
from .zip_outputs_and_cleanup import zip_outputs_and_cleanup

class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        payload_json = json.loads(tool_parameters["payload_json"])

        sub_paragraphs = payload_json["sub_paragraphs"]
        normalize_pairs = payload_json["normalize_pairs"]
        normalized_fig_keys = payload_json["normalized_fig_keys"]
        normalized_table_keys = payload_json["normalized_table_keys"]
        fig_legends = payload_json["fig_legends"]
        fig_legend_fill = payload_json["fig_legend_fill"]
        tables_from_orig = payload_json["tables_from_orig"]
        tables_from_attach = payload_json["tables_from_attach"]
        table_text_fill = payload_json["table_text_fill"]
        figure_match_plan = payload_json["figure_match_plan"]
        pmid_normalize_pairs = payload_json["pmid_normalize_pairs"]
        pmid_comment_tasks = payload_json["pmid_comment_tasks"]
        normalized_additional_keys = payload_json["normalized_additional_keys"]
        order_id = payload_json["order_id"]
        orig_tables = payload_json["orig_tables"] 

        sub_docx_path = payload_json["sub_path"]
        work_dir = payload_json["work_dir"]


        datainfo1 = merge_assets_and_build_comment_tasks(
                                sub_paragraphs,
                                normalize_pairs,
                                normalized_fig_keys,
                                normalized_table_keys,
                                fig_legends,
                                fig_legend_fill,
                                tables_from_orig,
                                tables_from_attach,
                                table_text_fill,
                                figure_match_plan,
                                pmid_normalize_pairs,
                                pmid_comment_tasks,
                                orig_tables,
                            )
        fig_assets = datainfo1["fig_assets"]
        table_assets = datainfo1["table_assets"]
        comment_tasks = datainfo1["comment_tasks"]
        normalize_pairs_final = datainfo1["normalize_pairs_final"]
        figure_copy_plan = datainfo1["figure_copy_plan"]


        datainfo2 = apply_all_changes(
                            sub_docx_path,
                            work_dir,
                            normalize_pairs_final,
                            fig_assets,
                            table_assets,
                            comment_tasks,
                            figure_copy_plan,
                        )
        final_docx_path = datainfo2["final_docx_path"]
        table_docx_path = datainfo2["table_docx_path"]
        final_figure_dir = datainfo2["final_figure_dir"]
        change_spans = datainfo2["change_spans"]
        figure_insert_reports = datainfo2["figure_insert_reports"]
        

        docx_in = final_docx_path
        base_dir = os.path.dirname(docx_in)
        base_name = os.path.splitext(os.path.basename(docx_in))[0]
        docx_out = os.path.join(base_dir, f"{base_name}.blue.docx")

        datainfo3 = blue_highlight_references_v3(
                docx_in,
                docx_out,
                normalized_fig_keys,
                normalized_table_keys,
                normalized_additional_keys,
            )
        final_docx_blue_path = datainfo3["final_docx_blue_path"]

        zip_bytes = zip_outputs_and_cleanup(work_dir,final_docx_blue_path, final_figure_dir, table_docx_path,order_id)
        
        yield self.create_blob_message(
            blob=zip_bytes,          # 直接传入 bytes，不需要 base64 编码
            meta={
                "mime_type": "application/zip",
                "filename": f"实验方案修改({order_id})-AI已修改格式.zip",
            },
        )

def _figure_match_plan_to_dict(figure_match_plan: Union[Dict[str, Any], List[Any]]) -> Dict[str, Any]:
    """
    将 figure_match_plan 统一转成 dict：
      - 若本身是 dict：直接返回
      - 若是 list：兼容少量变体
    """
    if isinstance(figure_match_plan, dict):
        return figure_match_plan

    out = {}
    if isinstance(figure_match_plan, list):
        for it in figure_match_plan:
            if not isinstance(it, dict):
                continue
            # 结构1：{"figure_id":"Figure 1","path":"..."}
            # if "figure_id" in it and ("path" in it or "file_path" in it):
            #     fid = it.get("figure_id")
            #     path = it.get("path") or it.get("file_path")
            #     out[fid] = {"path": path, "reason": it.get("reason", "")}
            #     continue
            # 结构1：{'figure_id': 'Figure 2A', 'figure_info': {'path': 'C:\\Users\\ADMINI~1\\AppData\\Local\\Temp\\paper_work_5xvcegl9\\figure_raw\\Figures\\figure2\\figure2_600_画板 1.tif', 'reason': 'Matches Figure 2, prefers .tif format.'}}
            if "figure_id" in it and "figure_info" in it:
                fid = it.get("figure_id")
                finfo = it.get("figure_info")
                if "path" in finfo and "reason" in finfo:
                  path = finfo.get("path")
                  reason = finfo.get("reason", "")
                  out[fid] = {"path": path, "reason": reason}
                  continue
                
            # 结构2：{"Figure 1": {"path":"..."}}
            for k, v in it.items():
                if isinstance(v, dict) and ("path" in v or "file_path" in v):
                    out[k] = {"path": v.get("path") or v.get("file_path"), "reason": v.get("reason", "")}
    return out


def _find_first_paragraph_index(paras: List[Dict[str, Any]], needle: str) -> int:
    """在段落列表中找到 needle 第一次出现的段落索引 p（用于 p_hint），找不到返回 -1"""
    if not needle:
        return -1
    for it in paras or []:
        text = it.get("text") or ""
        if needle in text:
            return it.get("p", -1)
    return -1


def merge_assets_and_build_comment_tasks(
    sub_paragraphs: List[Dict[str, Any]],
    normalize_pairs: List[Dict[str, Any]],
    normalized_fig_keys: List[str],
    normalized_table_keys: List[str],
    fig_legends: Dict[str, Any],
    fig_legend_fill: Dict[str, Any],
    tables_from_orig: Dict[str, Any],
    tables_from_attach: Dict[str, Any],
    table_text_fill: Dict[str, Any],
    figure_match_plan,  # 允许 dict 或 list
    pmid_normalize_pairs: Optional[List[Dict[str, Any]]] = None,  # 新增：NodeF2 输出
    pmid_comment_tasks: Optional[List[Dict[str, Any]]] = None,    # 新增：NodeF2 输出
    orig_tables: List[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Node13：合并图/表资产，并构造 comment_tasks（缺失图片/缺失表/文献异常等）。

    关键点：
    - comment_tasks 不输出 start/end，只输出 anchor + p_hint + text
      （Node14 会在最终 docx 中重新定位 anchor 并精确插入 Word Comment）
    - normalize_pairs_final 会合并 PMID 的替换对（kind="pmid"）
    """

    # 将 figure_match_plan 转为 dict：{ "Figure 1": {"path": "...", "reason": "..."} }
    figure_match_map = _figure_match_plan_to_dict(figure_match_plan)

    fig_assets: Dict[str, Any] = {}
    table_assets: Dict[str, Any] = {}
    comment_tasks: List[Dict[str, Any]] = []

    def split_trailing_upper(s: str) -> Tuple[str, str, bool]:
        s = s.strip()
        m = re.fullmatch(r"(.*?)([A-Z])", s)
        if not m:
            return "", "", False
        return m.group(1), m.group(2), True

    # 1) 图：合并图题/图注 + 图片路径
    for fk in normalized_fig_keys or []:
        leg = (fig_legends or {}).get(fk, {}) or {}
        fill = (fig_legend_fill or {}).get(fk, {}) or {}
        # leg = (fig_legends or {}).get(strip_last_letter_if_endswith_letter(fk), {}) or {}
        # fill = (fig_legend_fill or {}).get(strip_last_letter_if_endswith_letter(fk), {}) or {}

        # title = (leg.get("title") or fill.get("title") or "").strip()
        # caption = (leg.get("caption") or fill.get("caption") or "").strip()
        title = (fill.get("title") or leg.get("title") or  "").strip()
        caption = (fill.get("caption") or leg.get("caption") or "").strip()

        #兼容此种情况
        # "fig_legends": {
        #     "Figure 2": {
        #       "title": "Characterization of the chemical structure of nanomaterials.",
        #       "caption": "(A) Particle size distribution of PAMAM; (B) Particle size distribution of PAMAM/siTLR4 complexes; (C) Particle size distribution of SeNPs; (D) Gel electrophoresis of PAMAM/siTLR4; (E) TEM image of PAMAM/siTLR4 complex; (F) Zeta potential of SeNPs, PAMAM, and PAMAM/siTLR4; (G) NMR spectra of HA and HA-CHO; (H) Infrared spectra of HA and HA-CHO; (I) Gelation time of hydrogel formed by mixing different concentrations of PAMAM (wt%) and 5% (wt%) HA-CHO in equal volumes; (J) Relationship between G', G'' and time for different hydrogel materials; (K) Frequency-dependent behavior of G' and G'' in hydrogel samples; (L) Mechanical stress-strain responses of hydrogels with different formulations.",
        #       "caption_by_part": {
        #         "A": "Particle size distribution of PAMAM;",
        #         "B": "Particle size distribution of PAMAM/siTLR4 complexes;",
        #         "C": "Particle size distribution of SeNPs;",
        #         "D": "Gel electrophoresis of PAMAM/siTLR4;",
        #         "E": "TEM image of PAMAM/siTLR4 complex;",
        #         "F": "Zeta potential of SeNPs, PAMAM, and PAMAM/siTLR4;",
        #         "G": "NMR spectra of HA and HA-CHO;",
        #         "H": "Infrared spectra of HA and HA-CHO;",
        #         "I": "Gelation time of hydrogel formed by mixing different concentrations of PAMAM (wt%) and 5% (wt%) HA-CHO in equal volumes;",
        #         "J": "Relationship between G', G'' and time for different hydrogel materials;",
        #         "K": "Frequency-dependent behavior of G' and G'' in hydrogel samples;",
        #         "L": "Mechanical stress-strain responses of hydrogels with different formulations."
        #       }
        first_key,sub_key,is_sun = split_trailing_upper(fk)
        if is_sun:
            if title == '':
                leg = (fig_legends or {}).get(first_key, {}) or {}
                title = (leg.get("title") or  "").strip()
            
            if caption == '':
                leg = (fig_legends or {}).get(first_key, {}) or {}
                if 'caption_by_part' in leg:
                    leg_caption_by_part = leg.get('caption_by_part')
                    if sub_key in leg_caption_by_part:
                        caption = (leg_caption_by_part.get(sub_key) or  "").strip()


        match = figure_match_map.get(fk)
        img_path = match.get("path") if isinstance(match, dict) else None

        missing_image = not (img_path and os.path.exists(img_path))

        # Figure 文件夹仍保持原始后缀（不转码），插入 docx 的 tif->png 由 Node14 处理
        ext = os.path.splitext(img_path)[1] if img_path else ".tif"
        dst_name = f"{fk}{ext}"

        fig_assets[fk] = {
            "title": title,
            "caption": caption,
            "image_src_path": img_path,
            "image_dst_name": dst_name,
            "already_exists_in_sub": False,
            "missing": missing_image or (title == "" and caption == "")
        }

        if missing_image:
            p_hint = _find_first_paragraph_index(sub_paragraphs, fk)
            if p_hint != -1:
                comment_tasks.append({
                    "anchor": fk,
                    "p_hint": p_hint,
                    "text": f"未在原稿与附件中找到对应图片文件：{fk}"
                })
            else:
                comment_tasks.append({
                    "anchor": fk,
                    "text": f"未在原稿与附件中找到对应图片文件：{fk}"
                })


    def clean_2d_list_by_header(data):
        """
        根据二维列表的表头（第一行）去除空白列
        规则：表头为空字符串的列，整列删除
        """
        if not data:
            return []

        # 1. 获取表头（第一行），标记**非空列的索引**
        header = data[0]
        keep_indexes = [i for i, val in enumerate(header) if val.strip() != '']

        # 2. 每一行都只保留这些索引的内容
        cleaned_data = []
        for row in data:
            cleaned_row = [row[i] for i in keep_indexes]
            cleaned_data.append(cleaned_row)

        return cleaned_data


    # 2) 表：原稿优先，其次附件，再补 title/note
    for tk in normalized_table_keys or []:
        base = (tables_from_orig or {}).get(tk) or (tables_from_attach or {}).get(tk) or {}
        fill = (table_text_fill or {}).get(tk) or {}

        title = (base.get("title") or fill.get("title") or "").strip()
        note = (base.get("note") or fill.get("note") or "").strip()
        grid = base.get("grid") or []

        if not grid and 'grid_ref' in base:
            if 'tbl_index' in base['grid_ref']:
                grid_ref = base['grid_ref']['tbl_index']
                for item_table in orig_tables:
                    if 'tbl_index' in item_table and item_table['tbl_index']== grid_ref:
                        grid = clean_2d_list_by_header(item_table['grid'])
                        break

        missing = (not grid)
        table_assets[tk] = {"title": title, "note": note, "grid": grid, "missing": missing}

        if missing:
            p_hint = _find_first_paragraph_index(sub_paragraphs, tk)
            if p_hint != -1:
                comment_tasks.append({
                    "anchor": tk,
                    "p_hint": p_hint,
                    "text": f"未在原稿与附件中找到对应表格内容：{tk}"
                })
            else:
                comment_tasks.append({
                    "anchor": tk,
                    "text": f"未在原稿与附件中找到对应表格内容：{tk}"
                })

    # 3) Figure 拷贝计划：仅拷贝找到的图片
    figure_copy_plan: List[Dict[str, Any]] = []
    for fk, a in fig_assets.items():
        src = a.get("image_src_path")
        if src and os.path.exists(src):
            figure_copy_plan.append({"src": src, "dst_name": a.get("image_dst_name")})

    # 4) 合并 PMID 替换对 + PMID 批注任务
    normalize_pairs_final = list(normalize_pairs or [])
    if pmid_normalize_pairs:
        normalize_pairs_final.extend([x for x in pmid_normalize_pairs if isinstance(x, dict)])

    if pmid_comment_tasks:
        for t in pmid_comment_tasks:
            if isinstance(t, dict) and t.get("anchor") and t.get("text"):
                comment_tasks.append(t)

    return {
        "fig_assets": fig_assets,
        "table_assets": table_assets,
        "comment_tasks": comment_tasks,
        "normalize_pairs_final": normalize_pairs_final,
        "figure_copy_plan": figure_copy_plan
    }