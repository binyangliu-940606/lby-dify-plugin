import json
import os
import io
import shutil
import zipfile

from typing import Dict, Any
# from collections.abc import Generator
# from dify_plugin import Tool
# from dify_plugin.entities.tool import ToolInvokeMessage


# class LbyToolsTool(Tool):
#     def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
#         #获取参数如文件url
#         payload_json = json.loads(tool_parameters["payload_json"])
#         work_dir = payload_json["work_dir"]
#         final_docx_blue_path = payload_json["final_docx_blue_path"]
#         final_figure_dir = payload_json["final_figure_dir"]
#         table_docx_path = payload_json["table_docx_path"]
#         order_id = payload_json["order_id"]

#         zip_bytes = zip_outputs_and_cleanup(work_dir,final_docx_blue_path, final_figure_dir, table_docx_path)
        
#         yield self.create_blob_message(
#             blob=zip_bytes,          # 直接传入 bytes，不需要 base64 编码
#             meta={
#                 "mime_type": "application/zip",
#                 "filename": f"实验方案修改({order_id})-AI已修改格式.zip",
#             },
#         )

def zip_outputs_and_cleanup(
    work_dir: str,
    final_docx_blue_path: str,
    final_figure_dir: str,
    table_docx_path: str,
    order_id:str
) -> Dict[str, Any]:
    """
    Node16：打包并清理
    输入：
      - work_dir：临时工作目录（最后会整目录删除）
      - final_docx_blue_path：最终（已标蓝）的次稿docx路径
      - final_figure_dir：Figure 文件夹路径
      - table_docx_path：Table.docx 路径
    输出：
      - zip_blob：zip 的二进制 bytes（用于返回 BLOB 流）
    """

    # 基本校验（缺文件就报错，便于定位问题）
    if not os.path.exists(final_docx_blue_path):
        raise FileNotFoundError(f"final_docx_blue_path 不存在: {final_docx_blue_path}")
    if not os.path.exists(table_docx_path):
        raise FileNotFoundError(f"table_docx_path 不存在: {table_docx_path}")
    if not os.path.isdir(final_figure_dir):
        # Figure 文件夹允许为空，但目录至少应存在
        os.makedirs(final_figure_dir, exist_ok=True)

    buf = io.BytesIO()

    # 打包结构：
    # /Final.docx（最终文件）
    # /Table.docx
    # /Figures/*
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # 最终 docx：统一命名 Final.docx（或你也可以用原文件名）
        z.write(final_docx_blue_path, arcname=f"实验方案修改({order_id})-AI已修改格式.docx")

        # Table.docx 固定命名
        z.write(table_docx_path, arcname="Tables.docx")

        # Figures 文件夹
        for base, _, files in os.walk(final_figure_dir):
            for fn in files:
                fp = os.path.join(base, fn)
                rel = os.path.relpath(fp, final_figure_dir)
                z.write(fp, arcname=os.path.join("Figures", rel))

    zip_bytes = buf.getvalue()

    # # 再写入本地
    # with open("mem_output.zip", "wb") as f:
    #     f.write(zip_bytes)

    # 第七步：清理所有产生的文档/临时文件，释放磁盘空间
    if work_dir and os.path.isdir(work_dir):
        shutil.rmtree(work_dir, ignore_errors=True)

    return zip_bytes
    # return {"zip_blob": zip_bytes}