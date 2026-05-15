"""查询某周或所有周的材料状态。"""

import os
import sys
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _configure_output_streams():
    """Avoid UnicodeEncodeError on Windows GBK consoles."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


_configure_output_streams()

from officecli_helper import (
    DocxCache, detect_week_status, get_missing_items,
    get_inserted_items, MATERIAL_ORDER, DEFAULT_SEMESTER,
    DocumentLockedError, release_docx, wait_until_file_available
)
from config import resolve_docx_path


def check_status(docx_path, week_title=None):
    doc = DocxCache(docx_path)
    weeks = doc.find_all_weeks()

    if not weeks:
        return "总表中暂无任何周次的记录。"

    if week_title:
        if "学期" not in week_title:
            week_title = DEFAULT_SEMESTER + week_title
        keyword = week_title.replace(DEFAULT_SEMESTER, "").strip() or week_title[-20:]
        status = detect_week_status(docx_path, keyword)
        return format_status(week_title, get_inserted_items(status), get_missing_items(status))
    else:
        lines = []
        for idx, title in weeks:
            keyword = title.replace(DEFAULT_SEMESTER, "").strip() or title[-20:]
            status = detect_week_status(docx_path, keyword)
            lines.append(format_status(title, get_inserted_items(status), get_missing_items(status)))
            lines.append("")
        return "\n".join(lines).strip()


def format_status(title, inserted, missing):
    parts = [f"【{title}】"]
    items = []
    for item in MATERIAL_ORDER:
        items.append(f"[有] {item}" if item in inserted else f"[缺] {item}")
    parts.append("  ".join(items))
    report = [m for m in missing if m not in ("会议记录", "标题", "简要介绍")]
    if report:
        parts.append(f"还缺：{'、'.join(report)}")
    else:
        parts.append("材料齐全")
    return "\n".join(parts)


def main():
    parser = argparse.ArgumentParser(description="查询会议材料状态")
    parser.add_argument("--docx", default=None, help="总表路径（可选）")
    parser.add_argument("--week-title", default=None, help="周标题关键词")

    args = parser.parse_args()
    docx_path, config_message = resolve_docx_path(args.docx)
    if not docx_path:
        print(config_message)
        sys.exit(1)
    try:
        print(check_status(docx_path, args.week_title))
    except DocumentLockedError as e:
        print(str(e))
        sys.exit(1)
    finally:
        release_docx(docx_path)
        wait_until_file_available(docx_path)


if __name__ == "__main__":
    main()
