

import json

from typing import Dict, Any, List
from collections.abc import Generator
from docx import Document
# from dify_plugin import Tool
# from dify_plugin.entities.tool import ToolInvokeMessage


# class LbyToolsTool(Tool):
#     def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
#         #获取参数如文件url
#         payload_json = json.loads(tool_parameters["payload_json"])
#         docx_path = payload_json["docx_path"]

#         yield self.create_json_message({
#             "return_data": extract_docx_plaintext_and_positions(docx_path),
#         })

def extract_docx_plaintext_and_positions(docx_path: str) -> Dict[str, Any]:
    """
    只做确定性抽取：段落索引 -> 段落文本
    注：run级位置不在这里做，后续替换/标蓝工具会处理 run 拆分。
    """
    doc = Document(docx_path)
    paras: List[Dict[str, Any]] = []
    for i, p in enumerate(doc.paragraphs):
        paras.append({"p": i, "text": p.text or ""})
    return {"sub_paragraphs": paras}