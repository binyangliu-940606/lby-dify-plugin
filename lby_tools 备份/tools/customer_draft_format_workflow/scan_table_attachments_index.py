import os
import json
import pandas as pd

from collections.abc import Generator
from typing import Dict, Any, List
from docx import Document

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage



class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        payload_json = json.loads(tool_parameters["payload_json"])
        table_raw_dir = payload_json["table_raw_dir"]
        table_files_meta = payload_json["table_files_meta"]

        yield self.create_json_message(
            scan_table_attachments_index(table_raw_dir, table_files_meta),
        )


def scan_table_attachments_index(table_raw_dir: str, table_files_meta: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    输出一个“轻量索引”给 LLM 节点做匹配计划：
    - docx: 提取前 N 段文本片段（避免太长）
    - xlsx: 列出 sheet 名
    - csv: 前 N 行预览
    """
    idx = {}
    for f in table_files_meta or []:
        path = f["path"]
        ext = (f.get("ext") or "").lower()

        item = {"path": path, "ext": ext}

        try:
            if ext == ".docx":
                doc = Document(path)
                preview = []
                for p in doc.paragraphs[:80]:
                    t = (p.text or "").strip()
                    if t:
                        preview.append(t[:200])
                item["docx_preview_paras"] = preview

            elif ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
                xl = pd.ExcelFile(path)
                item["sheets"] = xl.sheet_names[:50]

            elif ext == ".csv":
                df = pd.read_csv(path, nrows=30, encoding="utf-8", engine="python")
                item["csv_preview"] = df.head(20).to_dict(orient="split")

            else:
                # 其他类型先不索引
                item["note"] = "unsupported_preview"
        except Exception as e:
            item["error"] = str(e)

        idx[os.path.basename(path)] = item

    return {"table_attachment_index": idx}