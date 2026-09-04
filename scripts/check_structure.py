# -*- coding: utf-8 -*-
"""结构自检：版本原子目录、md/docx 配对提醒。

用法：py -3.12 scripts/check_structure.py

规则：
  1. course/<课程名>/<版本>/ 下必须有 大纲/ 和 教案/ 两个目录（教案可为空，仅 .gitkeep）；
  2. 文件名采用课程代码原名（如 25JX31802-…），不以版本年为前缀；
  3. 大纲/ 中若存在 .docx 而定稿缺同名 .md（事实源）→ 提示为警告（允许 docx 先行）；
  4. md 若有但缺同名 .docx（可编译到学校版）→ 警告；
  5. 每个版本目录必须有 README.md 自查清单；
  6. 每个版本目录都能追溯到 version/<版本>.md。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

# Windows 中文路径/输出：固定 UTF-8，避免控制台代码页导致乱码
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

REPO_ROOT = Path(__file__).resolve().parent.parent
VERSIONS = {"2025", "2027"}


def _content_files(d: Path):
    """返回目录下真实内容文件（排除 .gitkeep / .gitignore 等占位）。"""
    return [p for p in d.iterdir() if p.is_file() and p.name not in (".gitkeep", ".gitignore")]


def main() -> int:
    errors: list[str] = []
    warnings: list[str] = []
    course_root = REPO_ROOT / "course"
    if not course_root.is_dir():
        print("[错误] 缺少 course/ 目录")
        return 1

    for course_dir in sorted(p for p in course_root.iterdir() if p.is_dir()):
        for ver_dir in sorted(p for p in course_dir.iterdir() if p.is_dir()):
            ver = ver_dir.name
            rel = ver_dir.relative_to(REPO_ROOT)
            if ver.startswith("."):
                continue

            if ver not in VERSIONS:
                errors.append(f"{rel}: 未识别的版本目录（应为 2025/2027）")
                continue

            # 6. 可追溯到 version/<ver>.md
            if not (REPO_ROOT / "version" / f"{ver}.md").is_file():
                errors.append(f"{rel}: 缺少 version/{ver}.md 场景说明")

            # 5. 版本目录必须有 README
            if not (ver_dir / "README.md").is_file():
                errors.append(f"{rel}: 缺少 README.md 自查清单")

            # 1. 必须有 大纲/ 与 教案/ 目录
            for kind in ("大纲", "教案"):
                kind_dir = ver_dir / kind
                if not kind_dir.is_dir():
                    errors.append(f"{rel}: 缺少 {kind}/ 目录")
                    continue

                files = _content_files(kind_dir)
                if not files:
                    # 教案可为空（尚未生成），不报错
                    if kind == "大纲":
                        warnings.append(f"{kind_dir.relative_to(REPO_ROOT)}: 尚未生成内容")
                    continue

                # 2. 命名不强制版本年前缀，此处不检查前缀

                # 3/4. md/docx 配对提醒
                md_stems = {p.stem for p in files if p.suffix == ".md"}
                docx_stems = {p.stem for p in files if p.suffix == ".docx"}
                for stem in sorted(md_stems - docx_stems):
                    warnings.append(
                        f"{kind_dir.relative_to(REPO_ROOT)}: {stem}.md 缺同名 .docx（编译产物待生成）"
                    )
                for stem in sorted(docx_stems - md_stems):
                    warnings.append(
                        f"{kind_dir.relative_to(REPO_ROOT)}: {stem}.docx 缺同名 .md（事实源待补）"
                    )

    # major/ 引用层校验：每个专业引用的课程目录必须真实存在
    major_root = REPO_ROOT / "major"
    if major_root.is_dir():
        for major_dir in sorted(p for p in major_root.iterdir() if p.is_dir()):
            for ref_file in sorted(major_dir.glob("*.md")):
                text = ref_file.read_text(encoding="utf-8", errors="replace")
                # 匹配 markdown 链接 [..](..../course/<课程名>/...) 中的课程名
                for m in re.findall(r"\]\(\.\./\.\./course/([^/)]+)/", text):
                    if not (REPO_ROOT / "course" / m).is_dir():
                        errors.append(f"{ref_file.relative_to(REPO_ROOT)}: 引用了不存在的课程 {m}")

    for w in warnings:
        print("[警告]", w)

    if errors:
        print("[结构自检失败]")
        for e in errors:
            print("  -", e)
        return 1

    print("[结构自检通过] 版本目录原子完整，major 引用有效。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
