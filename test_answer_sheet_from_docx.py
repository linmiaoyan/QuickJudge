# -*- coding: utf-8 -*-
"""临时脚本：用指定 Word 测试答题卡解析与生成。"""
import os
import sys
import json

# 确保项目根在 path 中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

DOCX_PATH = r"d:\WeChat\xwechat_files\wxid_u4aslyzqaizt22_41c9\msg\file\2026-03\高三英语限时练（二）答案.docx"
OUTPUT_HTML = os.path.join(os.path.dirname(__file__), "data", "exports", "test_答题卡_限时练2.html")

def main():
    from utils.paper_parser import parse_paper_docx
    from utils.answer_sheet_generator import generate_answer_sheet_html

    if not os.path.isfile(DOCX_PATH):
        print("Word 文件不存在:", DOCX_PATH)
        return
    parsed = parse_paper_docx(DOCX_PATH)
    print("解析结果 keys:", list(parsed.keys()))
    if parsed.get("error"):
        print("解析错误:", parsed["error"])
        return
    print("选择题区块:", json.dumps(parsed.get("choice_sections", []), ensure_ascii=False, indent=2))
    print("选择题答案数量:", len(parsed.get("choice_answers", {})))
    print("主观题区块:", json.dumps(parsed.get("subjective_sections", []), ensure_ascii=False, indent=2)[:500])
    html = generate_answer_sheet_html(parsed, title="限时练2", show_answer_keys=True)
    os.makedirs(os.path.dirname(OUTPUT_HTML), exist_ok=True)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print("已生成答题卡 HTML:", OUTPUT_HTML)

if __name__ == "__main__":
    main()
