import os
import tempfile
import json
import requests

from typing import Dict, Any
from collections.abc import Generator
from .download_and_extract import download_and_extract
from .extract_docx_plaintext_and_positions import extract_docx_plaintext_and_positions

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage



class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数如文件url
        payload_json = json.loads(tool_parameters["payload_json"])
        orig_file_path = payload_json["orig_file_url"]
        sub_file_path = payload_json["new_docx_url"]
        figure_zip_url = payload_json["figure_zip_url"]
        table_zip_url = payload_json["table_zip_url"]

        work_dir,orig_path,sub_path = prepare_workspace(orig_file_path, sub_file_path)

        figure_raw_dir,table_raw_dir,figure_files_meta,table_files_meta = download_and_extract(figure_zip_url, table_zip_url, work_dir)

        sub_paragraphs = extract_docx_plaintext_and_positions(sub_path)

        orig_paragraphs = extract_docx_plaintext_and_positions(orig_path)

        yield self.create_json_message({
            "work_dir": work_dir,
            "orig_path":orig_path,
            "sub_path":sub_path,
            "figure_raw_dir":figure_raw_dir,
            "table_raw_dir":table_raw_dir,
            "figure_files_meta":figure_files_meta,
            "table_files_meta":table_files_meta,
            "sub_paragraphs":sub_paragraphs['sub_paragraphs'],
            "orig_paragraphs":orig_paragraphs['sub_paragraphs'],
        })



def _ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)


def _download_to_file(url: str, out_path: str, timeout: int = 180, retries: int = 3):
    """从可下载URL下载文件到 out_path（流式 + 重试）"""
    last_err = None
    for _ in range(retries):
        try:
            with requests.get(url, stream=True, timeout=timeout) as r:
                r.raise_for_status()
                with open(out_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=1024 * 1024):
                        if chunk:
                            f.write(chunk)
            return
        except Exception as e:
            last_err = e
    raise last_err


def _assert_docx_is_zip(file_path: str, name: str):
    """校验下载结果是否是 docx(zip)，文件头应为 PK"""
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        raise RuntimeError(f"{name} 下载后为空文件")

    with open(file_path, "rb") as f:
        head = f.read(4)

    if head[:2] != b"PK":
        with open(file_path, "rb") as f:
            sample = f.read(300)
        try:
            sample_text = sample.decode("utf-8", errors="replace")
        except Exception:
            sample_text = str(sample)
        raise RuntimeError(
            f"{name} 下载结果不是docx(zip)。文件头={head!r}，前300字节预览：{sample_text}"
        )


def prepare_workspace(orig_url: str, sub_url: str) -> Dict[str, Any]:
    """
    你的输入是 JSON 字符串，例如：
      {"new_docx_url":"http://.../xxx.docx?..."}
    本函数会解析 JSON，取 new_docx_url 下载到临时目录，输出 orig_path/sub_path。
    """

    work_dir = tempfile.mkdtemp(prefix="paper_work_")
    _ensure_dir(work_dir)

    orig_path = os.path.join(work_dir, "orig.docx")
    sub_path = os.path.join(work_dir, "sub.docx")

    _download_to_file(orig_url, orig_path)
    _download_to_file(sub_url, sub_path)

    _assert_docx_is_zip(orig_path, "原稿docx")
    _assert_docx_is_zip(sub_path, "次稿docx")

    # return {"work_dir": work_dir, "orig_path": orig_path, "sub_path": sub_path}
    return work_dir,orig_path,sub_path

