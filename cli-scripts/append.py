import argparse
import os
import pandas as pd
from openpyxl import load_workbook


def read_sheet(file, sheet):
    try:
        return pd.read_excel(file, sheet_name=sheet, dtype=str)
    except Exception:
        return None


def append_to_target(target_file, target_sheet, df_append, mode):
    if not os.path.exists(target_file):
        with pd.ExcelWriter(target_file, engine="openpyxl") as writer:
            df_append.to_excel(writer, sheet_name=target_sheet, index=False)
        return

    try:
        existing = pd.read_excel(target_file, sheet_name=target_sheet, dtype=str)
        sheet_exists = True
    except Exception:
        existing = None
        sheet_exists = False

    if not sheet_exists:
        with pd.ExcelWriter(target_file, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
            df_append.to_excel(writer, sheet_name=target_sheet, index=False)
        return

    if mode == "no-header":
        combined = pd.concat([existing, df_append], ignore_index=True)

    elif mode == "header-intersection":
        common_cols = [c for c in df_append.columns if c in existing.columns]
        if not common_cols:
            print("⚠️ 没有共同表头，跳过该数据块")
            return
        combined = pd.concat(
            [existing[common_cols], df_append[common_cols]],
            ignore_index=True
        )

    elif mode == "header-union":
        all_cols = list(dict.fromkeys(list(existing.columns) + list(df_append.columns)))
        existing2 = existing.reindex(columns=all_cols)
        df2 = df_append.reindex(columns=all_cols)
        combined = pd.concat([existing2, df2], ignore_index=True)
    else:
        raise ValueError("未知模式")

    with pd.ExcelWriter(target_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        combined.to_excel(writer, sheet_name=target_sheet, index=False)


def process_file(file, src_sheet, target_file, target_sheet, mode):
    print(f"👉 处理: {file}")
    df = read_sheet(file, src_sheet)

    if df is None:
        print(f"⚠️ {file} 未找到 sheet: {src_sheet}，跳过")
        return

    append_to_target(target_file, target_sheet, df, mode)


def main():
    parser = argparse.ArgumentParser(description="Excel Sheet 追加工具")

    parser.add_argument(
        "-s", "--source",
        required=True,
        help="源 Excel 文件 或 文件夹"
    )

    parser.add_argument("-ss", "--src-sheet", required=True, help="源sheet")
    parser.add_argument("-t", "--target", required=True, help="目标Excel文件")
    parser.add_argument("-ts", "--target-sheet", required=True, help="目标sheet")

    parser.add_argument(
        "-m",
        "--mode",
        required=True,
        choices=["no-header", "header-intersection", "header-union"],
        help="追加模式",
    )

    args = parser.parse_args()

    path = args.source

    if os.path.isfile(path):
        process_file(path, args.src_sheet, args.target, args.target_sheet, args.mode)

    elif os.path.isdir(path):
        for f in os.listdir(path):
            if f.lower().endswith((".xls", ".xlsx")):
                process_file(
                    os.path.join(path, f),
                    args.src_sheet,
                    args.target,
                    args.target_sheet,
                    args.mode
                )
    else:
        print("❌ -s 既不是文件也不是文件夹")


if __name__ == "__main__":
    main()
