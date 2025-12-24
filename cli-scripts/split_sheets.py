import argparse
import os
import pandas as pd
from openpyxl import Workbook


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)


def save_sheet(df, out_path, sheet_name):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))

    wb.save(out_path)


def split_excel_file(file_path, output_dir, name_mode):
    base_name = os.path.basename(file_path)
    print(f"👉 处理: {base_name}")

    try:
        sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
    except Exception as e:
        print(f"⚠️ 无法读取 {file_path} ，已跳过 -> {e}")
        return

    for sheet_name, df in sheets.items():
        sheet_folder = os.path.join(output_dir, sheet_name)
        ensure_dir(sheet_folder)

        if name_mode == "source":
            out_file_name = base_name
        else:
            safe_sheet_name = sheet_name.strip() or "sheet"
            out_file_name = f"{safe_sheet_name}.xlsx"

            i = 1
            while os.path.exists(os.path.join(sheet_folder, out_file_name)):
                out_file_name = f"{safe_sheet_name}_{i}.xlsx"
                i += 1

        out_path = os.path.join(sheet_folder, out_file_name)

        save_sheet(df, out_path, sheet_name)

        print(f"   ✔️ 生成: {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Excel Sheet 分离工具")

    parser.add_argument(
        "-s", "--source",
        required=True,
        help="源 Excel 文件 或 文件夹"
    )

    parser.add_argument(
        "-o", "--output",
        required=True,
        help="输出根目录（必须是文件夹）"
    )

    parser.add_argument(
        "--name-mode",
        choices=["source", "sheet"],
        default="source",
        help="生成文件命名模式: source=源文件名 sheet=sheet名(重复自动编号)"
    )

    args = parser.parse_args()

    ensure_dir(args.output)

    path = args.source

    if os.path.isfile(path):
        split_excel_file(path, args.output, args.name_mode)

    elif os.path.isdir(path):
        for f in os.listdir(path):
            if f.lower().endswith((".xls", ".xlsx")):
                split_excel_file(
                    os.path.join(path, f),
                    args.output,
                    args.name_mode
                )
    else:
        print("❌ -s 既不是文件也不是文件夹")


if __name__ == "__main__":
    main()
