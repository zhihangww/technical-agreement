"""
测试脚本：验证 Excel → LLM 提取的完整流程

用法：
    python test_excel_extractor.py "文件路径.xls" --sheet "Sheet名" --chunks 1
    
参数：
    --sheet   指定 Sheet 名称（不指定则默认第一个）
    --chunks  最多处理几个分块（不指定则全部处理，调试时建议设为 1 省 token）
    --model   模型名称（默认 azure/gpt-4o）

示例（只处理第1个分块，快速验证）：
    python test_excel_extractor.py "2GHV012529_ELK-04 for STEP(AB)_Released (1).xls" --chunks 1
"""

import sys
import json
import os
from dotenv import load_dotenv

load_dotenv()

from excel_to_params import ExcelParamExtractor


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    file_path = sys.argv[1]
    sheet_name = None
    max_chunks = 0
    model = "azure/gpt-4o"

    # 解析命令行参数
    args = sys.argv[2:]
    i = 0
    while i < len(args):
        if args[i] == '--sheet' and i + 1 < len(args):
            sheet_name = args[i + 1]
            i += 2
        elif args[i] == '--chunks' and i + 1 < len(args):
            max_chunks = int(args[i + 1])
            i += 2
        elif args[i] == '--model' and i + 1 < len(args):
            model = args[i + 1]
            i += 2
        else:
            i += 1

    print(f"{'=' * 60}")
    print(f"Excel 参数提取测试")
    print(f"{'=' * 60}")
    print(f"文件: {file_path}")
    print(f"模型: {model}")
    if max_chunks > 0:
        print(f"限制: 仅处理前 {max_chunks} 个分块")

    # 加载文件
    extractor = ExcelParamExtractor(model=model)
    sheets = extractor.load_file(file_path)

    print(f"\nSheet 列表 ({len(sheets)} 个):")
    for idx, name in enumerate(sheets):
        print(f"  [{idx}] {name}")

    if sheet_name is None:
        sheet_name = sheets[0]
    print(f"\n处理 Sheet: {sheet_name}")

    # 执行提取
    print(f"\n{'=' * 60}")
    print(f"开始提取...")
    print(f"{'=' * 60}")

    result = extractor.extract(
        sheet_name=sheet_name,
        rows_per_chunk=100,
        max_chunks=max_chunks
    )

    # 打印结果
    print(f"\n{'=' * 60}")
    print(f"提取结果")
    print(f"{'=' * 60}")
    print(f"总参数数: {result['total_extracted']}")

    print(f"\n--- 中文参数名列表 ({len(result['chinese_names'])} 个) ---")
    for name in result['chinese_names']:
        print(f"  {name}")

    print(f"\n--- 英文参数名列表 ({len(result['english_names'])} 个) ---")
    for name in result['english_names']:
        print(f"  {name}")

    print(f"\n--- 规范库条目 ({len(result['spec_entries'])} 个) ---")
    for entry in result['spec_entries'][:20]:
        val_display = entry['value'] if entry['value'] else '(无规范值)'
        print(f"  {entry['name']}: {val_display}")
    if len(result['spec_entries']) > 20:
        print(f"  ... 还有 {len(result['spec_entries']) - 20} 个")

    # 保存完整结果到 JSON
    output_path = "test_excel_extract_result.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"\n{'=' * 60}")
    print(f"完整结果已保存到: {output_path}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
