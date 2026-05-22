import zipfile
from typing import Dict, Callable


def read_docx_zip(docx_path: str) -> Dict[str, bytes]:
    """读取 docx(zip) 内所有文件，返回 {name: bytes}"""
    files = {}
    with zipfile.ZipFile(docx_path, "r") as z:
        for info in z.infolist():
            files[info.filename] = z.read(info.filename)
    return files


def write_docx_zip(files: Dict[str, bytes], out_path: str):
    """把 {name: bytes} 写成一个新的 docx(zip)，不会产生重复 name"""
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)


def edit_docx_zip(docx_path: str, out_path: str, editor: Callable[[Dict[str, bytes]], None]):
    """
    读取 docx -> editor(files) 原地修改 dict -> 写回 out_path
    editor 中只需要修改 files["word/document.xml"] 等即可
    """
    files = read_docx_zip(docx_path)
    editor(files)
    write_docx_zip(files, out_path)