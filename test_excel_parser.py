"""
测试脚本：验证 Excel → HTML 解析输出

用法：
    python test_excel_parser.py "文件路径.xls"
    python test_excel_parser.py "文件路径.xlsx" --sheet "Sheet名称"

不指定 --sheet 时默认解析第一个 Sheet。
输出 HTML 文件可在浏览器中打开，直观检查解析结果。
"""

import sys
import os
from excel_to_params import ExcelParser


def main():
    if len(sys.argv) < 2:
        print("用法: python test_excel_parser.py <Excel文件路径> [--sheet <Sheet名称>]")
        print()
        print("示例:")
        print('  python test_excel_parser.py "2GHV012529_ELK-04 for STEP(AB)_Released (1).xls"')
        print('  python test_excel_parser.py "1HC0071250 GIS OFFER CATALOG BA (1).xlsm" --sheet "Sheet1"')
        return

    file_path = sys.argv[1]

    sheet_name = None
    if '--sheet' in sys.argv:
        idx = sys.argv.index('--sheet')
        if idx + 1 < len(sys.argv):
            sheet_name = sys.argv[idx + 1]

    print(f"{'=' * 60}")
    print(f"Excel 解析测试")
    print(f"{'=' * 60}")
    print(f"文件: {file_path}")

    parser = ExcelParser(file_path)

    # 列出所有 Sheet
    sheets = parser.get_sheet_names()
    print(f"\nSheet 列表 ({len(sheets)} 个):")
    for i, name in enumerate(sheets):
        print(f"  [{i}] {name}")

    if sheet_name is None:
        sheet_name = sheets[0]
        print(f"\n默认选择第一个 Sheet: {sheet_name}")
    else:
        if sheet_name not in sheets:
            print(f"\n错误: Sheet '{sheet_name}' 不存在")
            return
        print(f"\n指定 Sheet: {sheet_name}")

    # 统计信息
    print(f"\n解析中...")
    stats = parser.get_sheet_stats(sheet_name)
    print(f"  原始行数: {stats['total_rows']}")
    print(f"  有效行数: {stats['non_empty_rows']}")
    print(f"  有效列数: {stats['columns']}")
    print(f"  过滤空行: {stats['empty_rows_removed']}")

    # 完整 HTML
    html = parser.parse_sheet(sheet_name)
    print(f"  HTML 长度: {len(html)} 字符")

    # 分块测试
    chunks = parser.parse_sheet_to_chunks(sheet_name, rows_per_chunk=100)
    print(f"\n分块结果 (每块100行):")
    print(f"  总块数: {len(chunks)}")
    for i, chunk in enumerate(chunks):
        print(f"  块 {i + 1}: {len(chunk)} 字符")

    # 保存完整 HTML 到文件，可在浏览器中打开检查
    output_path = "test_excel_output.html"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Excel解析测试 - {os.path.basename(file_path)}</title>
    <style>
        body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
        h1 {{ color: #1f77b4; }}
        h3 {{ color: #2c3e50; margin-top: 30px; }}
        table {{ border-collapse: collapse; margin: 10px 0; background: white; font-size: 13px; }}
        td {{ border: 1px solid #ddd; padding: 6px 10px; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        .stats {{ background: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
    </style>
</head>
<body>
<h1>Excel 解析结果</h1>
<div class="stats">
    <p><b>文件:</b> {os.path.basename(file_path)}</p>
    <p><b>Sheet:</b> {sheet_name}</p>
    <p><b>有效行数:</b> {stats['non_empty_rows']}（原始 {stats['total_rows']} 行）</p>
    <p><b>有效列数:</b> {stats['columns']}</p>
    <p><b>HTML 长度:</b> {len(html)} 字符</p>
</div>
{html}
</body>
</html>""")

    print(f"\n{'=' * 60}")
    print(f"输出已保存到: {output_path}")
    print(f"请在浏览器中打开查看解析结果是否正确")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
