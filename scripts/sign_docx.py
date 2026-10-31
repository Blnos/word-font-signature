#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档字体签名工具（适配 WorkBuddy / Antigravity）

功能：
  1. fonts : 列出内置可选字体（龙藏体、平方长安）。
  2. scan  : 扫描 .docx，按文档顺序列出所有候选签字位置（含表格内），不修改文档，用于确认位置。
  3. sign  : 用所选手写字体把姓名渲染成透明签名图，自适应尺寸后插入指定位置。
  4. batch : 批量签署，支持文件夹整批（内容相同）或按 CSV 清单逐文件签署不同内容；字体只选一次，支持全签/仅首位置，失败隔离并汇总。

设计要点：
  - 直接在原文档上修改，最大程度保留原格式（不重建文档）。
  - 签名图 4 倍超采样并裁剪到文字边界，保证清晰、无多余留白。
  - Word 内只固定宽度，高度严格按图片真实宽高比等比计算，避免拉伸变形。

依赖：python-docx、Pillow（pip install python-docx pillow）
"""

import argparse
import csv
import datetime
import re
import sys
import tempfile
from pathlib import Path

# ----------------------------------------------------------------------------
# 依赖检查
# ----------------------------------------------------------------------------
try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph
except ImportError:
    print("缺少依赖 python-docx，请先执行：pip install python-docx")
    sys.exit(2)

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("缺少依赖 Pillow，请先执行：pip install pillow")
    sys.exit(2)


# 内置字体目录（相对本脚本：../assets/fonts/）
FONTS_DIR = Path(__file__).resolve().parent.parent / "assets" / "fonts"

# 可选字体注册表：id 即对话选项编号
FONT_REGISTRY = [
    {"id": "1", "name": "龙藏体", "desc": "笔画较粗、连笔潇洒，签名感强", "file": "LongCang-Regular.ttf"},
    {"id": "2", "name": "平方长安", "desc": "纤细清秀、字形端正", "file": "PingFangChangAnTi-2.ttf"},
]

# 编号/别名 -> 注册表下标
FONT_ALIASES = {
    "1": 0, "longcang": 0, "long": 0, "龙藏体": 0, "龙藏": 0,
    "2": 1, "changan": 1, "pingfang": 1, "平方长安": 1, "长安": 1, "平方": 1,
}


def resolve_font(value):
    """把 --font 的输入解析为 (字体路径, 显示名)。支持编号、别名或自定义路径。"""
    if not value:
        item = FONT_REGISTRY[0]
        return FONTS_DIR / item["file"], item["name"]
    raw = str(value).strip()
    idx = FONT_ALIASES.get(raw) or FONT_ALIASES.get(raw.lower())
    if idx is not None:
        item = FONT_REGISTRY[idx]
        return FONTS_DIR / item["file"], item["name"]
    return Path(value).expanduser(), raw


# ----------------------------------------------------------------------------
# 文档遍历：按文档顺序产出所有段落（含嵌套表格内段落）
# ----------------------------------------------------------------------------
def iter_paragraphs(container, location="正文"):
    """生成 (Paragraph, 位置描述)。container 为文档对象或单元格 Cell。"""
    if hasattr(container, "_tc"):  # 单元格
        root = container._tc
    else:  # 文档对象，正文容器
        root = container.element.body

    for child in root.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, container), location
        elif child.tag == qn("w:tbl"):
            table = Table(child, container)
            # 直接遍历底层 XML 行/列元素，避免 row.cells 代理对象重建与误去重；
            # 合并单元格在 XML 中表现为单个带 gridSpan 的 tc，天然不重复。
            for r_idx, tr in enumerate(table._tbl.tr_lst):
                for c_idx, tc in enumerate(tr.tc_lst):
                    cell = _Cell(tc, table)
                    loc = f"{location} › 表格(第{r_idx + 1}行,第{c_idx + 1}列)"
                    yield from iter_paragraphs(cell, loc)


def paragraph_has_image(paragraph):
    """判断段落是否已含图片（避免重复签署）。"""
    return bool(paragraph._p.findall(".//" + qn("w:drawing")))


# 精确签字词，避免把标题/正文中的“签署”误判为签字行
SIGNATURE_TERMS_CN = ("签字", "签名", "签章", "签署人", "会签", "签：", "落款")
SIGNATURE_TERMS_EN = ("signature", "sign here", "signed", "signatory")


def is_signature_paragraph(text):
    """判断段落文本是否为签字行。"""
    lowered = text.lower()
    if any(term in text for term in SIGNATURE_TERMS_CN):
        return True
    if any(term in lowered for term in SIGNATURE_TERMS_EN):
        return True
    return False


def collect_signature_positions(doc):
    """返回所有候选签字位置列表。"""
    positions = []
    for idx, (paragraph, location) in enumerate(iter_paragraphs(doc)):
        text = paragraph.text.strip()
        if text and is_signature_paragraph(text) and not paragraph_has_image(paragraph):
            positions.append({
                "seq": len(positions),
                "para_index": idx,
                "location": location,
                "text": text,
                "paragraph": paragraph,
            })
    return positions


# ----------------------------------------------------------------------------
# 签名图片生成
# ----------------------------------------------------------------------------
def render_signature_image(name, font_path, out_path):
    """4 倍超采样渲染紧贴文字边界的透明签名图，返回 (像素宽, 像素高)。"""
    supersample = 4
    font = ImageFont.truetype(str(font_path), 88 * supersample)

    probe = ImageDraw.Draw(Image.new("RGBA", (8, 8)))
    bbox = probe.textbbox((0, 0), name, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    pad = 14 * supersample
    canvas = Image.new("RGBA", (text_w + pad * 2, text_h + pad * 2), (255, 255, 255, 0))
    draw = ImageDraw.Draw(canvas)
    draw.text((pad - bbox[0], pad - bbox[1]), name, font=font, fill=(0, 0, 0, 255))
    canvas.save(str(out_path))
    return canvas.width, canvas.height


def auto_width_pt(name):
    """按姓名长度估算 Word 内目标宽度（pt），并限制在合理区间。"""
    width = 26 * len(name.strip())
    return max(58, min(width, 150))


# ----------------------------------------------------------------------------
# 插入
# ----------------------------------------------------------------------------
def insert_inline(paragraph, img_path, width_pt, height_pt):
    """在目标段落末尾内联追加签名（紧跟“签字：”标签）。"""
    run = paragraph.add_run()
    run.add_picture(str(img_path), width=Pt(width_pt), height=Pt(height_pt))


def insert_on_new_line(paragraph, img_path, width_pt, height_pt):
    """在目标段落下方新建一行放置签名。"""
    new_element = OxmlElement("w:p")
    paragraph._p.addnext(new_element)
    new_paragraph = Paragraph(new_element, paragraph._parent)
    new_paragraph.add_run().add_picture(
        str(img_path), width=Pt(width_pt), height=Pt(height_pt)
    )


# ----------------------------------------------------------------------------
# 可选：日期填写（保守策略，默认关闭）
# ----------------------------------------------------------------------------
def fill_date_fields(doc):
    """仅处理“日期：”后为空的段落，追加当天日期，返回处理数量。"""
    today = datetime.date.today()
    date_text = f"{today.year}年{today.month}月{today.day}日"
    count = 0
    for paragraph, _ in iter_paragraphs(doc):
        text = paragraph.text.strip()
        if re.match(r"^日期[:：]\s*[_＿\s]*$", text):
            paragraph.add_run(date_text)
            count += 1
    return count


# ----------------------------------------------------------------------------
# 核心：单文档签署（sign 与 batch 共用）
# ----------------------------------------------------------------------------
def build_signed_filename(in_path, name):
    """
    默认输出文件名：去掉原文件名末尾的“模板”，再追加“_签署内容已签署”。
    例：上市公司董秘书面确认模板.docx + 王佳蓓 → 上市公司董秘书面确认_王佳蓓已签署.docx
    """
    stem = Path(in_path).stem
    if stem.endswith("模板"):
        stem = stem[:-2]
    sig = str(name).strip()
    for ch in '\\/:*?"<>|':  # 去除文件系统非法字符
        sig = sig.replace(ch, "")
    sig = sig.replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")
    if not sig:
        sig = "签名"
    return f"{stem}_{sig}已签署.docx"


def append_fallback_signature(doc, name, font_path, width_pt=None, fill_date=False):
    """兜底：未识别到签字位置时，在文档末尾空一行，右对齐插入签名（右下方）。"""
    actual_width = width_pt if width_pt else auto_width_pt(name)

    with tempfile.TemporaryDirectory() as tmp_dir:
        img_path = Path(tmp_dir) / "signature.png"
        img_w, img_h = render_signature_image(name, font_path, img_path)
        height_pt = actual_width * img_h / img_w  # 严格等比，防变形

        # 先加一个空段落（与最后一行文字空一行），再加右对齐签名段落
        doc.add_paragraph()
        sign_paragraph = doc.add_paragraph()
        sign_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sign_paragraph.add_run().add_picture(
            str(img_path), width=Pt(actual_width), height=Pt(height_pt)
        )

        date_count = fill_date_fields(doc) if fill_date else 0

    return {
        "inserted": 1,
        "dates": date_count,
        "width": actual_width,
        "height": height_pt,
        "fallback": True,
    }


def sign_one_document(in_path, out_path, name, font_path, mode="all", index=0,
                      width_pt=None, new_line=False, fill_date=False):
    """
    处理单个文档并保存，返回结果 dict。
    mode：
      - "all"   ：文档内所有候选位置都签（batch 默认）
      - "first" ：只签第一个位置（batch --first-only）
      - "index" ：签第 index 个位置（sign 单文档模式）
    未识别到任何签字位置时，走末尾兜底（空一行 + 右对齐）。
    """
    doc = Document(str(in_path))
    positions = collect_signature_positions(doc)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)

    if not positions:
        result = append_fallback_signature(
            doc, name, font_path, width_pt=width_pt, fill_date=fill_date
        )
        doc.save(str(out_path))
        result["pictures"] = len(Document(str(out_path)).inline_shapes)
        return result

    if mode == "first":
        targets = [positions[0]]
    elif mode == "index":
        if index < 0 or index >= len(positions):
            raise IndexError(
                f"位置序号 {index} 超出范围（共 {len(positions)} 个，序号 0~{len(positions) - 1}）"
            )
        targets = [positions[index]]
    else:  # "all"
        targets = positions

    actual_width = width_pt if width_pt else auto_width_pt(name)

    with tempfile.TemporaryDirectory() as tmp_dir:
        img_path = Path(tmp_dir) / "signature.png"
        img_w, img_h = render_signature_image(name, font_path, img_path)
        height_pt = actual_width * img_h / img_w  # 严格等比，防变形

        insert_fn = insert_on_new_line if new_line else insert_inline
        for target in targets:
            insert_fn(target["paragraph"], img_path, actual_width, height_pt)

        date_count = fill_date_fields(doc) if fill_date else 0

        doc.save(str(out_path))

    verify_doc = Document(str(out_path))
    return {
        "inserted": len(targets),
        "pictures": len(verify_doc.inline_shapes),
        "dates": date_count,
        "width": actual_width,
        "height": height_pt,
        "fallback": False,
    }


def collect_docx_files(input_dir, recursive, skip_dirs=None):
    """收集文件夹内待签 .docx，跳过 Word 锁文件（~$ 开头）及指定输出目录。"""
    base = Path(input_dir)
    if not base.is_dir():
        raise NotADirectoryError(f"输入路径不是文件夹：{input_dir}")
    skip = {Path(d).resolve() for d in (skip_dirs or [])}
    pattern = "**/*.docx" if recursive else "*.docx"
    files = []
    for candidate in base.glob(pattern):
        if not (candidate.is_file() and not candidate.name.startswith("~$")):
            continue
        resolved = candidate.resolve()
        if any(skip_path == resolved or skip_path in resolved.parents for skip_path in skip):
            continue  # 递归时跳过自己的输出目录
        files.append(candidate)
    return sorted(files)


def resolve_manifest_file(raw_file, base_dirs):
    """把清单中的文件项解析为实际路径，支持绝对路径或相对基准目录。"""
    candidate = Path(raw_file).expanduser()
    if candidate.is_absolute():
        return candidate
    for base in base_dirs:
        resolved = base / candidate
        if resolved.exists():
            return resolved
    return base_dirs[0] / candidate  # 暂存，处理时再报不存在


def parse_manifest(manifest_path, input_dir=None):
    """解析 CSV 清单，返回 [(文件路径, 签署内容), ...]，保持顺序，重复文件后者覆盖。"""
    path = Path(manifest_path)
    if not path.is_file():
        raise FileNotFoundError(f"清单文件不存在：{manifest_path}")

    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        rows = [r for r in csv.reader(fh) if r and any(c.strip() for c in r)]
    if not rows:
        raise ValueError("清单内容为空")

    header = [c.strip() for c in rows[0]]
    has_header = (
        any(k in header[0] for k in ("文件", "路径", "file", "path"))
        or (len(header) > 1 and any(k in header[1] for k in ("内容", "签署", "签名", "content")))
    )
    data_rows = rows[1:] if has_header else rows

    base_dirs = [path.parent]
    if input_dir:
        base_dirs.append(Path(input_dir))

    ordered = {}
    for row in data_rows:
        if len(row) < 2:
            continue
        raw_file = row[0].strip()
        content = row[1].strip()
        if not raw_file or not content or raw_file.startswith("#"):
            continue
        file_path = resolve_manifest_file(raw_file, base_dirs)
        ordered[str(file_path)] = content
    return [(Path(f), c) for f, c in ordered.items()]


# ----------------------------------------------------------------------------
# 子命令
# ----------------------------------------------------------------------------
def cmd_scan(args):
    doc = Document(args.input)
    positions = collect_signature_positions(doc)
    print(f"文档：{args.input}")
    print(f"共识别到 {len(positions)} 个候选签字位置：\n")
    if not positions:
        print("未找到签字位置。可在 references/positioning.md 查看关键词与排查方法。")
        return 0
    for p in positions:
        preview = p["text"] if len(p["text"]) <= 40 else p["text"][:40] + "…"
        print(f"[{p['seq']}] 段落序{p['para_index']}｜{p['location']}")
        print(f"     文本：{preview}\n")
    return 0


def cmd_fonts(args):
    """列出内置可选字体。"""
    print("可用签名字体：")
    for item in FONT_REGISTRY:
        print(f"  [{item['id']}] {item['name']} — {item['desc']}")
    return 0


def cmd_sign(args):
    if not args.name:
        print("签名模式需要 --name 指定姓名。")
        return 2
    font_path, font_label = resolve_font(args.font)
    if not font_path.exists():
        print(f"字体文件不存在：{font_path}（可用 fonts 子命令查看内置字体）")
        return 2

    in_path = Path(args.input)
    out_path = Path(args.output) if args.output else in_path.with_name(
        build_signed_filename(in_path, args.name)
    )
    mode = "all" if args.all else "index"
    try:
        result = sign_one_document(
            in_path, out_path, args.name, font_path,
            mode=mode, index=args.index, width_pt=args.width,
            new_line=args.new_line, fill_date=args.fill_date,
        )
    except (ValueError, IndexError) as exc:
        print(str(exc))
        return 1

    print("签署完成")
    print(f"  输入文档：{in_path}")
    print(f"  签名姓名：{args.name}")
    print(f"  所用字体：{font_label}")
    print(f"  插入位置：{result['inserted']} 处（宽度 {result['width']:.1f}pt，"
          f"高度 {result['height']:.1f}pt，等比缩放）")
    if result.get("fallback"):
        print("  说明：未识别到签字位置，已在文末空一行并右对齐放置签名。")
    if result["dates"]:
        print(f"  已填日期：{result['dates']} 处")
    print(f"  输出文档：{out_path}")
    print(f"  验证：输出文档内嵌图片共 {result['pictures']} 个，文件可正常打开。")
    return 0


def cmd_batch(args):
    font_path, font_label = resolve_font(args.font)
    if not font_path.exists():
        print(f"字体文件不存在：{font_path}（可用 fonts 子命令查看内置字体）")
        return 2

    # 确定输出目录
    if args.output_dir:
        out_dir = Path(args.output_dir)
    elif args.manifest:
        out_dir = Path(args.manifest).parent / "批量签署结果"
    else:
        out_dir = Path(args.input_dir) / "批量签署结果"

    # 构建任务：[(文件, 内容), ...]
    if args.manifest:
        try:
            tasks = parse_manifest(args.manifest, args.input_dir)
        except (FileNotFoundError, ValueError) as exc:
            print(str(exc))
            return 2
    else:
        if not args.name:
            print("批量模式需要 --name 指定姓名，或使用 --manifest 清单。")
            return 2
        try:
            files = collect_docx_files(args.input_dir, args.recursive, skip_dirs=[out_dir])
        except NotADirectoryError as exc:
            print(str(exc))
            return 2
        tasks = [(file, args.name) for file in files]

    if not tasks:
        print("没有可处理的文档（已跳过锁文件与输出目录）。")
        return 1

    mode = "first" if args.first_only else "all"

    success, failed = [], []
    for file, content in tasks:
        out_path = out_dir / build_signed_filename(file, content)
        try:
            result = sign_one_document(
                file, out_path, content, font_path,
                mode=mode, width_pt=args.width,
                new_line=args.new_line, fill_date=args.fill_date,
            )
            success.append((file.name, content, result))
        except Exception as exc:  # 单文件失败隔离，不影响其余文档
            failed.append((file.name, content, str(exc)))

    print("批量签署完成")
    print(f"  所用字体：{font_label}")
    print(f"  共 {len(tasks)} 份：成功 {len(success)}，失败 {len(failed)}")
    print(f"  输出目录：{out_dir}")
    for name, content, result in success:
        tag = "文末兜底" if result.get("fallback") else "定位签署"
        print(f"    [成功][{tag}] {name}（内容“{content}”）：插入 {result['inserted']} 处，"
              f"内嵌图片 {result['pictures']} 个")
    for name, content, err in failed:
        print(f"    [失败] {name}（内容“{content}”）：{err}")
    return 0 if not failed else 1


# ----------------------------------------------------------------------------
# 入口
# ----------------------------------------------------------------------------
def build_parser():
    parser = argparse.ArgumentParser(description="Word 文档字体签名工具")
    sub = parser.add_subparsers(dest="command", required=True)

    p_scan = sub.add_parser("scan", help="扫描并列出候选签字位置，不修改文档")
    p_scan.add_argument("--input", "-i", required=True, help="待扫描的 .docx 路径")
    p_scan.set_defaults(func=cmd_scan)

    p_fonts = sub.add_parser("fonts", help="列出内置可选字体")
    p_fonts.set_defaults(func=cmd_fonts)

    p_sign = sub.add_parser("sign", help="生成字体签名并插入文档")
    p_sign.add_argument("--input", "-i", required=True, help="待签署的 .docx 路径")
    p_sign.add_argument("--name", "-n", required=True, help="用于生成签名的姓名")
    p_sign.add_argument("--output", "-o", help="输出路径（默认命名为 原文件名_签署人已签署.docx）")
    p_sign.add_argument("--index", type=int, default=0, help="插入第几个候选位置（默认 0）")
    p_sign.add_argument("--all", action="store_true", help="在所有候选位置插入同一签名")
    p_sign.add_argument("--width", type=float, help="签名宽度（pt），默认按姓名长度自适应")
    p_sign.add_argument("--new-line", action="store_true", help="在目标段落下方新建一行放置签名")
    p_sign.add_argument("--fill-date", action="store_true", help="同时填写空白“日期：”栏")
    p_sign.add_argument("--font", help="选择字体：编号/别名（1 龙藏体，2 平方长安）或自定义字体文件路径")
    p_sign.set_defaults(func=cmd_sign)

    p_batch = sub.add_parser("batch", help="批量签署：文件夹整批或按清单逐文件签署")
    p_batch.add_argument("--input-dir", "-d", help="待签署文档所在文件夹（清单模式可省略）")
    p_batch.add_argument("--name", "-n", help="签名内容（整批相同）；清单模式由清单提供")
    p_batch.add_argument("--manifest", "-m", help="CSV 清单路径（两列：文件名,签署内容）")
    p_batch.add_argument("--font", help="字体编号/别名或路径（整批只需选一次）")
    p_batch.add_argument("--output-dir", help="输出目录（默认建“批量签署结果”）")
    p_batch.add_argument("--first-only", action="store_true",
                         help="每份文档只签第一个位置（默认文档内所有位置都签）")
    p_batch.add_argument("--recursive", action="store_true", help="文件夹模式递归处理子文件夹")
    p_batch.add_argument("--width", type=float, help="签名宽度（pt）")
    p_batch.add_argument("--new-line", action="store_true", help="签名放到目标段落下一行")
    p_batch.add_argument("--fill-date", action="store_true", help="同时填写空白“日期：”栏")
    p_batch.set_defaults(func=cmd_batch)
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    try:
        return args.func(args)
    except Exception as exc:  # 顶层兜底，输出可读信息
        print(f"执行失败：{exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
