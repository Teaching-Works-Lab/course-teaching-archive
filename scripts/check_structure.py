# -*- coding: utf-8 -*-
"""结构自检：版本目录原子性、md/docx 成对、版本前缀一致。

用法：py -3.12 scripts/check_structure.py

硬规则：
  1. course/<课程名>/<版本>/ 下必须有 大纲/ 和 教案/ 两个目录；
  2. 大纲/ 与 教案/ 内每个 .md 必须有同名 .docx（反之亦然）；
  3. 文件名版本前缀（2025-/2027-）必须与所在版本目录一致；
  4. 每个版本目录必须有 README.md 自查清单；
  5. 每个版本目录都能追溯到 version/<版本>.md。
"""
from __future__ import annotations

import sys
from pathlib import Path

# Windows 中文路径/输出：固定 UTF-8，避免控制台代码页导致乱码
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

REPO_ROOT = Path(__file__).resolve().parent.parent
VERSIONS = {"2025", "2027"}


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

            if ver not in VERSIONS:
                errors.append(f"{rel}: 未识别的版本目录（应为 2025/2027）")
                continue

            # 5. 能追溯到 version/<ver>.md
            if not (REPO_ROOT / "version" / f"{ver}.md").is_file():
                errors.append(f"{rel}: 缺少 version/{ver}.md 场景说明")

            # 4. 版本目录必须有 README.md
            if not (ver_dir / "README.md").is_file():
                errors.append(f"{rel}: 缺少 README.md 自查清单")

            # 1. 必须有 大纲/ 和 教案/
            for kind in ("大纲", "教案"):
                kind_dir = ver_dir / kind
                if not kind_dir.is_dir():
                    errors.append(f"{rel}: 缺少 {kind}/ 目录")
                    continue

                md_files = sorted(kind_dir.glob("*.md"))
                docx_files = sorted(kind_dir.glob("*.docx"))
                if not md_files and not docx_files:
                    warnings.append(f"{kind_dir.relative_to(REPO_ROOT)}: 尚未生成内容")
                    continue

                # 2. md/docx 成对
                md_stems = {p.stem for p in md_files}
                docx_stems = {p.stem for p in docx_files}
                for stem in sorted(md_stems - docx_stems):
                    errors.append(f"{kind_dir.relative_to(REPO_ROOT)}: {stem}.md 缺少同名 .docx")
                for stem in sorted(docx_stems - md_stems):
                    errors.append(f"{kind_dir.relative_to(REPO_ROOT)}: {stem}.docx 缺少同名 .md")

                # 3. 版本前缀一致
                for f in md_files + docx_files:
                    if not f.name.startswith(ver + "-"):
                        errors.append(f"{f.relative_to(REPO_ROOT)}: 文件名缺少 {ver}- 前缀")

    for w in warnings:
        print("[警告]", w)

    if errors:
        print("[结构自检失败]")
        for e in errors:
            print("  -", e)
        return 1

    print("[结构自检通过] 版本目录原子完整、md/docx 成对、版本前缀一致。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
