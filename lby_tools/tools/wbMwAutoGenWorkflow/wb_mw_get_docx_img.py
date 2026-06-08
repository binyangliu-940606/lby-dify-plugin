# -*- coding: utf-8 -*-
import io
import zipfile
import requests
import base64

from collections.abc import Generator
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from typing import Any

# 无需转换、可直接输出的格式
ALLOWED_IMAGE_EXTS = frozenset({"gif", "jpeg", "jpg", "webp", "png"})

# ---------------------------------------------------------------------------
# 图片处理
# ---------------------------------------------------------------------------


def _normalize_ext(ext: str) -> str:
    ext = ext.lower()
    return "jpeg" if ext == "jpg" else ext


def _mime_for_ext(ext: str) -> str:
    return f"image/{ext}"


def _convert_to_png(img_bytes: bytes) -> bytes:
    from PIL import Image

    with Image.open(io.BytesIO(img_bytes)) as img:
        if img.mode in ("RGBA", "RGB", "L", "P"):
            converted = img
        else:
            converted = img.convert("RGBA")
        out = io.BytesIO()
        converted.save(out, format="PNG")
        return out.getvalue()


def _prepare_image(
    img_bytes: bytes, ext: str, filename: str
) -> tuple[bytes, str, str]:
    """
    返回 (最终字节, 最终文件名, mime_type)
    不在允许列表中的格式会转为 PNG。
    """
    ext = _normalize_ext(ext)
    if ext in ALLOWED_IMAGE_EXTS:
        return img_bytes, filename, _mime_for_ext(ext)

    png_bytes = _convert_to_png(img_bytes)
    base_name = filename.rsplit(".", 1)[0]
    return png_bytes, f"{base_name}.png", "image/png"


# ---------------------------------------------------------------------------
# Tool 入口
# ---------------------------------------------------------------------------


class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        doc_url = tool_parameters["url"]

        # 1. 下载 docx 到内存
        resp = requests.get(doc_url, timeout=30)
        resp.raise_for_status()
        docx_bytes = resp.content

        # 2. docx 本质是 zip，图片在 word/media/ 下
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            for name in zf.namelist():
                if name.startswith("word/media/") and not name.endswith("/"):
                    ext = name.rsplit(".", 1)[-1].lower()
                    if ext in ("png", "jpg", "jpeg", "gif", "bmp", "webp", "emf", "wmf"):
                        img_bytes = zf.read(name)
                        raw_filename = name.split("/")[-1]

                        try:
                            final_bytes, final_filename, mime_type = _prepare_image(
                                img_bytes, ext, raw_filename
                            )
                        except Exception as e:
                            yield self.create_text_message(
                                f"图片 {raw_filename} 转换失败: {e}"
                            )
                            continue

                        yield self.create_blob_message(
                            blob=final_bytes,
                            meta={
                                "mime_type": mime_type,
                                "filename": final_filename,
                            },
                        )