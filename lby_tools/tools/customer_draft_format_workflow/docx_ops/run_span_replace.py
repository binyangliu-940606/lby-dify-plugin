from dataclasses import dataclass
from typing import List, Tuple, Dict, Any, Optional
from docx.text.paragraph import Paragraph
from docx.shared import RGBColor

@dataclass
class RunSpan:
    """描述一次“命中文字串”在 paragraph 内的字符范围（相对 paragraph 拼接文本）"""
    start: int
    end: int  # 不含 end（Python slice 语义）

@dataclass
class ChangeSpan:
    """用于后续标蓝：记录替换后文本所在的段落与字符范围"""
    paragraph_index: int
    start: int
    end: int
    kind: str  # figure/table/pmid/additional 等

def _paragraph_full_text(p: Paragraph) -> str:
    return "".join(r.text for r in p.runs)

def _build_run_char_map(p: Paragraph) -> List[Tuple[int, int, int]]:
    """
    构建 run 的字符偏移映射：
    返回列表：[(run_index, run_start, run_end), ...]，run_end 不含
    """
    mapping = []
    cur = 0
    for i, r in enumerate(p.runs):
        t = r.text or ""
        mapping.append((i, cur, cur + len(t)))
        cur += len(t)
    return mapping

def _split_run_at(p: Paragraph, run_idx: int, split_pos_in_run: int):
    """
    把某个 run 按 split_pos_in_run 拆成两个 run，尽量保留原 run 样式
    - split_pos_in_run: 在该 run.text 内部的切分点
    """
    run = p.runs[run_idx]
    text = run.text or ""
    left = text[:split_pos_in_run]
    right = text[split_pos_in_run:]

    # 原 run 保留左边
    run.text = left

    # 在后面插入一个新 run 放右边，并复制样式
    new_run = p.add_run(right)
    # 复制常见样式（粗斜体下划线/字体/颜色等）
    new_run.bold = run.bold
    new_run.italic = run.italic
    new_run.underline = run.underline
    if run.font is not None:
        new_run.font.name = run.font.name
        new_run.font.size = run.font.size
        new_run.font.color.rgb = run.font.color.rgb if run.font.color else None

    # python-docx 的 add_run 是追加到段落末尾，不是插入到 run_idx 后
    # 为了真正插入，需要操作底层 xml，把 new_run._r 移动到 run._r 后面
    run._r.addnext(new_run._r)

def replace_text_preserve_runs(
    p: Paragraph,
    find_text: str,
    replace_text: str,
) -> List[RunSpan]:
    """
    在一个段落内做“保留格式”的替换：
    - 支持 find_text 跨多个 runs
    - 实现思路：
      1) 取 paragraph 全文 full_text
      2) 找到第一次出现的位置 start/end（可扩展为多次）
      3) 把命中范围的左右边界对齐到 run 边界（必要时拆 run）
      4) 将命中范围覆盖的 runs 文本合并替换为 replace_text（只改文本，不动样式）
    返回：替换后 replace_text 所在的 paragraph 字符范围（用于标蓝/批注锚定）
    """
    full = _paragraph_full_text(p)
    if not full or find_text not in full:
        return []

    spans = []
    # 这里先实现“替换所有出现次数”（更符合归一化需要）
    idx = 0
    while True:
        start = full.find(find_text, idx)
        if start == -1:
            break
        end = start + len(find_text)
        spans.append(RunSpan(start=start, end=end))
        idx = end

    # 逐个 span 替换：注意替换会改变全文长度，所以要做偏移修正
    offset = 0
    replaced_spans: List[RunSpan] = []

    for sp in spans:
        s = sp.start + offset
        e = sp.end + offset

        # 每次循环重建映射（因为 run 可能被拆/文本被改）
        mapping = _build_run_char_map(p)

        # 找到包含 s/e 的 run
        def locate(pos: int):
            for run_idx, a, b in mapping:
                if a <= pos < b:
                    return run_idx, pos - a, a, b
            # pos 可能刚好等于最后边界
            if mapping and pos == mapping[-1][2]:
                last = mapping[-1]
                return last[0], last[2] - last[1], last[1], last[2]
            return None

        L = locate(s)
        R = locate(e)
        if L is None or R is None:
            continue

        l_run, l_in, l_a, l_b = L
        r_run, r_in, r_a, r_b = R

        # 1) 先把右边界对齐：把 r_run 在 r_in 处拆开（如果不是边界）
        if r_in != 0 and r_in != (r_b - r_a):
            _split_run_at(p, r_run, r_in)

        # 2) 再把左边界对齐：把 l_run 在 l_in 处拆开（如果不是边界）
        # 注意：若左边 run 被拆，会影响右边 run 下标，因此先拆右再拆左通常更安全
        if l_in != 0 and l_in != (l_b - l_a):
            _split_run_at(p, l_run, l_in)
            # 左边拆完后，命中范围从“拆出来的右半 run”开始
            l_run = l_run + 1

        # 3) 重新获取映射，定位命中范围覆盖的 run 索引区间
        mapping2 = _build_run_char_map(p)
        # 重新定位 s/e（此时边界应落在 run 边界上）
        # 我们用字符区间判断 run 覆盖
        run_indices = []
        for run_idx, a, b in mapping2:
            if a >= s and b <= e and a != b:
                run_indices.append(run_idx)
        if not run_indices:
            # 命中可能覆盖到空 run 或边界特殊情况，保守跳过
            continue

        first_run = run_indices[0]
        last_run = run_indices[-1]

        # 4) 执行替换：把 first_run 文本设为 replace_text，其余覆盖 run 清空
        p.runs[first_run].text = replace_text
        for i in range(first_run + 1, last_run + 1):
            p.runs[i].text = ""

        # 5) 记录替换后 replace_text 的字符范围
        new_full = _paragraph_full_text(p)
        # 为稳妥：从 s 开始向后找 replace_text 的第一次出现
        new_start = new_full.find(replace_text, s)
        if new_start == -1:
            new_start = new_full.find(replace_text)
        new_end = new_start + len(replace_text) if new_start != -1 else new_start

        replaced_spans.append(RunSpan(new_start, new_end))

        # 6) 更新 offset：长度变化
        offset += len(replace_text) - len(find_text)
        full = new_full

    return replaced_spans