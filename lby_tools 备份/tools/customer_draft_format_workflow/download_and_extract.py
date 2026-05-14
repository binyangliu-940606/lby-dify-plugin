import os
import zipfile
import requests
import json

from typing import Dict, Any, List
from collections.abc import Generator

# from dify_plugin import Tool
# from dify_plugin.entities.tool import ToolInvokeMessage



# class LbyToolsTool(Tool):
#     def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
#         #获取参数如文件url
#         payload_json = json.loads(tool_parameters["payload_json"])
#         work_dir = payload_json["work_dir"]
#         figure_zip_url = payload_json["figure_zip_url"]
#         table_zip_url = payload_json["table_zip_url"]

#         yield self.create_json_message({
#             "return_data": download_and_extract(figure_zip_url, table_zip_url, work_dir),
#         })

def _ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)

def _download(url: str, out_path: str, timeout: int = 120):
    with requests.get(url, stream=True, timeout=timeout) as r:
        r.raise_for_status()
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)

def _list_files(root: str) -> List[Dict[str, Any]]:
    metas = []
    for base, _, files in os.walk(root):
        for fn in files:
            fp = os.path.join(base, fn)
            metas.append({
                "name": fn,
                "path": fp,
                "ext": os.path.splitext(fn)[1].lower(),
                "size": os.path.getsize(fp),
            })
    return metas

def download_and_extract(figure_zip_url: str, table_zip_url: str, work_dir: str) -> Dict[str, Any]:
    figure_raw_dir = os.path.join(work_dir, "figure_raw")
    table_raw_dir = os.path.join(work_dir, "table_raw")
    _ensure_dir(figure_raw_dir)
    _ensure_dir(table_raw_dir)

    if figure_zip_url:
        fig_zip = os.path.join(work_dir, "figure.zip")
        _download(figure_zip_url, fig_zip)
        with zipfile.ZipFile(fig_zip, "r") as z:
            z.extractall(figure_raw_dir)

    if table_zip_url:
        tab_zip = os.path.join(work_dir, "table.zip")
        _download(table_zip_url, tab_zip)
        with zipfile.ZipFile(tab_zip, "r") as z:
            z.extractall(table_raw_dir)

    # return {
    #     "figure_raw_dir": figure_raw_dir,
    #     "table_raw_dir": table_raw_dir,
    #     "figure_files_meta": _list_files(figure_raw_dir),
    #     "table_files_meta": _list_files(table_raw_dir),
    # }
    return figure_raw_dir,table_raw_dir,_list_files(figure_raw_dir),_list_files(table_raw_dir)
