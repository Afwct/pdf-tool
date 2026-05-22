#!/usr/bin/env python3
"""PDF 工具箱 CLI（PyMuPDF 开源版）。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

_CORE = Path(__file__).resolve().parent
if str(_CORE) not in sys.path:
    sys.path.insert(0, str(_CORE))

import converters as C  # noqa: E402

IMAGE_EXT = ", ".join(sorted(C.IMAGE_EXT))

SUBCOMMANDS = frozenset({
    "pdf2word", "pdf2excel", "pdf2ppt", "pdf2image", "pdf2md", "pdf2json", "pdf2txt",
    "word2pdf", "excel2pdf", "ppt2pdf", "image2pdf", "encrypt", "decrypt",
})

# 仅传文件路径时，按扩展名自动选择子命令
EXT_TO_CMD: dict[str, str] = {}
for _ext in (".doc", ".docx", ".docm", ".dot", ".dotx"):
    EXT_TO_CMD[_ext] = "word2pdf"
for _ext in (".xls", ".xlsx", ".xlsm"):
    EXT_TO_CMD[_ext] = "excel2pdf"
for _ext in (".ppt", ".pptx", ".pptm"):
    EXT_TO_CMD[_ext] = "ppt2pdf"
for _ext in C.IMAGE_EXT:
    EXT_TO_CMD[_ext] = "image2pdf"
EXT_TO_CMD[".pdf"] = "pdf2txt"


def _path(s: str) -> Path:
    return Path(s).expanduser().resolve()


def cmd_pdf2word(a: argparse.Namespace) -> int:
    out = C.pdf_to_word(_path(a.input), _path(a.output) if a.output else None)
    print(f"已生成: {out}")
    return 0


def cmd_pdf2excel(a: argparse.Namespace) -> int:
    out = C.pdf_to_excel(_path(a.input), _path(a.output) if a.output else None)
    print(f"已生成: {out}")
    return 0


def cmd_pdf2ppt(a: argparse.Namespace) -> int:
    out = C.pdf_to_ppt(_path(a.input), _path(a.output) if a.output else None, dpi=a.dpi)
    print(f"已生成: {out}")
    return 0


def cmd_pdf2image(a: argparse.Namespace) -> int:
    paths = C.pdf_to_images(
        _path(a.input),
        _path(a.output) if a.output else None,
        fmt=a.format,
        dpi=a.dpi,
    )
    if not paths:
        print("错误: 未导出任何图片", file=sys.stderr)
        return 1
    print(f"已导出 {len(paths)} 张 -> {paths[0].parent}")
    return 0


def cmd_pdf2md(a: argparse.Namespace) -> int:
    out = C.pdf_to_markdown(_path(a.input), _path(a.output) if a.output else None)
    print(f"已生成: {out}")
    return 0


def cmd_pdf2json(a: argparse.Namespace) -> int:
    out = C.pdf_to_json(_path(a.input), _path(a.output) if a.output else None)
    print(f"已生成: {out}")
    return 0


def cmd_pdf2txt(a: argparse.Namespace) -> int:
    out = C.pdf_to_text(_path(a.input), _path(a.output) if a.output else None)
    print(f"已生成: {out}")
    return 0


def cmd_to_pdf(a: argparse.Namespace) -> int:
    out = C.file_to_pdf(_path(a.input), _path(a.output) if a.output else None)
    print(f"已生成: {out}")
    return 0


def cmd_encrypt(a: argparse.Namespace) -> int:
    out = C.encrypt_pdf(
        _path(a.input),
        _path(a.output) if a.output else None,
        user_password=a.password,
        owner_password=a.owner,
    )
    print(f"已加密: {out}")
    return 0


def cmd_decrypt(a: argparse.Namespace) -> int:
    out = C.decrypt_pdf(
        _path(a.input),
        _path(a.output) if a.output else None,
        password=a.password,
    )
    print(f"已解密: {out}")
    return 0


def _default_image_pdf_out(sources: list[Path]) -> Path:
    if len(sources) == 1:
        return sources[0].with_suffix(".pdf")
    return sources[0].parent / f"{sources[0].stem}_合并.pdf"


def _merge_path_args(parts: list[str]) -> list[Path]:
    """合并被空格拆开的 Windows 路径（批处理未加引号时）。"""
    out: list[Path] = []
    buf = ""
    for part in parts:
        if part.startswith("-"):
            continue
        trial = f"{buf} {part}".strip() if buf else part.strip().strip('"')
        p = Path(trial).expanduser()
        if p.is_file():
            out.append(p.resolve())
            buf = ""
        else:
            buf = trial
    if buf:
        p = Path(buf.strip().strip('"')).expanduser()
        if p.is_file():
            out.append(p.resolve())
    return out


def _paths_from_list_file(path: Path) -> list[Path]:
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    return _merge_path_args([ln.strip().strip('"') for ln in lines if ln.strip()])


def cmd_image2pdf(a: argparse.Namespace) -> int:
    if a.pick:
        sources = C.pick_image_files()
        if not sources:
            print("未选择图片，已取消。", file=sys.stderr)
            return 1
    elif a.list_file:
        lf = _path(a.list_file)
        if not lf.is_file():
            print(f"错误: 列表文件不存在 {lf}", file=sys.stderr)
            return 1
        sources = _paths_from_list_file(lf)
    else:
        sources = _merge_path_args(a.input)
        if not sources:
            print("错误: 请指定图片或使用 --pick", file=sys.stderr)
            return 1

    for src in sources:
        if not src.is_file():
            print(f"错误: 找不到文件 {src}", file=sys.stderr)
            return 1
        ext = src.suffix.lower()
        if ext not in C.IMAGE_EXT:
            print(f"错误: 不支持格式 {ext or '(无扩展名)'} -> {src}", file=sys.stderr)
            return 1

    out = _path(a.output) if a.output else _default_image_pdf_out(sources)
    print(f"共 {len(sources)} 张图片")
    out = C.images_to_pdf(sources, out)
    print(f"已生成: {out}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="PDF 工具箱")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add_io(sp: argparse.ArgumentParser) -> None:
        sp.add_argument("input", help="输入文件")
        sp.add_argument("-o", "--output", help="输出路径（可选）")

    for name, help_ in (
        ("pdf2word", "PDF -> Word"),
        ("pdf2excel", "PDF -> Excel"),
        ("pdf2md", "PDF -> Markdown"),
        ("pdf2json", "PDF -> JSON"),
        ("pdf2txt", "PDF -> 文本"),
    ):
        s = sub.add_parser(name, help=help_)
        add_io(s)
        s.set_defaults(func=globals()[f"cmd_{name}"])

    for name, help_ in (
        ("word2pdf", "Word -> PDF"),
        ("excel2pdf", "Excel -> PDF"),
        ("ppt2pdf", "PPT -> PDF"),
    ):
        s = sub.add_parser(name, help=help_)
        add_io(s)
        s.set_defaults(func=cmd_to_pdf)

    s = sub.add_parser("pdf2ppt", help="PDF -> PPT")
    add_io(s)
    s.add_argument("--dpi", type=int, default=150)
    s.set_defaults(func=cmd_pdf2ppt)

    s = sub.add_parser("pdf2image", help="PDF -> 图片")
    add_io(s)
    s.add_argument("-f", "--format", choices=("png", "jpeg", "jpg"), default="png")
    s.add_argument("--dpi", type=int, default=150)
    s.set_defaults(func=cmd_pdf2image)

    s = sub.add_parser("image2pdf", help="图片 -> PDF（可多图）")
    s.add_argument("input", nargs="*", default=[])
    s.add_argument("-o", "--output")
    s.add_argument("--pick", action="store_true")
    s.add_argument(
        "--list-file",
        help="文本文件，每行一个图片路径（多图拖放备用）",
    )
    s.set_defaults(func=cmd_image2pdf)

    s = sub.add_parser("encrypt", help="加密 PDF")
    add_io(s)
    s.add_argument("-p", "--password", required=True)
    s.add_argument("--owner")
    s.set_defaults(func=cmd_encrypt)

    s = sub.add_parser("decrypt", help="解密 PDF")
    add_io(s)
    s.add_argument("-p", "--password", default="")
    s.set_defaults(func=cmd_decrypt)

    return p


def _is_existing_file(s: str) -> bool:
    if s.startswith("-"):
        return False
    return Path(s).expanduser().is_file()


def preprocess_argv(argv: list[str]) -> list[str]:
    """支持: python pdf_tool.py 文件.docx  -> 自动 word2pdf"""
    if not argv:
        return argv
    if argv[0] in SUBCOMMANDS:
        return argv
    if _is_existing_file(argv[0]):
        ext = Path(argv[0]).suffix.lower()
        if ext in C.IMAGE_EXT and len(argv) >= 1:
            if all(
                _is_existing_file(x) and Path(x).suffix.lower() in C.IMAGE_EXT
                for x in argv
            ):
                return ["image2pdf", *argv]
        cmd = EXT_TO_CMD.get(ext)
        if cmd:
            print(f"[auto] {ext} -> {cmd}", file=sys.stderr)
            return [cmd, *argv]
    return argv


def main(argv: list[str] | None = None) -> int:
    raw = list(argv) if argv is not None else sys.argv[1:]
    argv = preprocess_argv(raw)
    try:
        args = build_parser().parse_args(argv)
    except SystemExit as exc:
        if raw and raw[0] not in SUBCOMMANDS:
            print(
                "\n用法示例: python pdf_tool.py word2pdf \"文件.docx\"\n"
                "或进入 word2pdf 文件夹，拖文件到 run.bat\n",
                file=sys.stderr,
            )
        raise exc
    try:
        return args.func(args)
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
