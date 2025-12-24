import os
import re
import shutil
import argparse
import logging

logging.basicConfig(level=logging.INFO, format="👉 %(message)s")

def load_ids(txt_file, pattern=None):
    ids = []
    regex = re.compile(pattern) if pattern else None

    with open(txt_file, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            if regex:
                m = regex.search(line)
                if m:
                    ids.append(m.group(1) if m.groups() else m.group(0))
                else:
                    logging.warning(f"⚠️ 未匹配到: {line}")
            else:
                ids.append(line)

    logging.info(f"🎯 共加载匹配关键字: {len(ids)}")
    return ids


def path_match(target, ids):
    return any(i in target for i in ids)


def ensure_dir(p):
    os.makedirs(p, exist_ok=True)


def copy_keep_structure(full_path, output_root, source_root):
    rel = os.path.relpath(full_path, source_root)
    dst = os.path.join(output_root, rel)
    ensure_dir(os.path.dirname(dst))
    shutil.copy2(full_path, dst)
    logging.info(f"📦 保持结构复制 -> {dst}")


def copy_flat(full_path, output_root):
    ensure_dir(output_root)
    filename = os.path.basename(full_path)
    dst = os.path.join(output_root, filename)

    base, ext = os.path.splitext(filename)
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_root, f"{base}_{i}{ext}")
        i += 1

    shutil.copy2(full_path, dst)
    logging.info(f"📦 扁平复制 -> {dst}")


def copy_depth(full_path, output_root, source_root, depth):
    rel = os.path.relpath(full_path, source_root)
    parts = rel.split(os.sep)

    keep = parts[:depth]
    name = "__".join(parts[depth:]) if len(parts) > depth else parts[-1]

    dst_dir = os.path.join(output_root, *keep)
    ensure_dir(dst_dir)
    dst = os.path.join(dst_dir, name)

    base, ext = os.path.splitext(name)
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(dst_dir, f"{base}_{i}{ext}")
        i += 1

    shutil.copy2(full_path, dst)
    logging.info(f"📦 depth复制 -> {dst}")


def extract(source_root, txt_file, output_root, scope, mode, depth, pattern):
    ids = load_ids(txt_file, pattern)

    source_root = os.path.abspath(source_root)
    output_root = os.path.abspath(output_root)

    ensure_dir(output_root)

    matched = False   # ← 新增

    for root, dirs, files in os.walk(source_root):
        # 排除 output 自己
        dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(root, d)) != output_root]

        rel_root = os.path.relpath(root, source_root)

        # 目录匹配
        if scope in ("dirs", "all"):
            folder_name = os.path.basename(root)
            check_target = folder_name if scope == "dirs" else rel_root

            if path_match(check_target, ids):
                matched = True   # ← 命中
                logging.info(f"📂 匹配到目录: {root}")
                for f in files:
                    full_path = os.path.join(root, f)

                    if mode == "keep":
                        copy_keep_structure(full_path, output_root, source_root)
                    elif mode == "depth":
                        copy_depth(full_path, output_root, source_root, depth)
                    else:
                        copy_flat(full_path, output_root)

        # 文件匹配
        if scope in ("files", "all"):
            for f in files:
                full_path = os.path.join(root, f)

                check_target = f if scope == "files" else os.path.relpath(full_path, source_root)

                if path_match(check_target, ids):
                    matched = True   # ← 命中
                    logging.info(f"📄 匹配到文件: {full_path}")

                    if mode == "keep":
                        copy_keep_structure(full_path, output_root, source_root)
                    elif mode == "depth":
                        copy_depth(full_path, output_root, source_root, depth)
                    else:
                        copy_flat(full_path, output_root)

    # 统一输出
    if not matched:
        logging.warning("没有找到任何匹配项，请检查 pattern / scope / txt 是否正确！")

def main():
    HELP = {
        "desc": "模糊提取文件工具：根据 txt 关键字/正则匹配源目录中的文件或文件夹并复制到目标目录",

        "source": "源目录路径（支持递归扫描多层子目录）",
        "txt": "关键字/规则文本文件，每一行一条数据",
        "output": "输出目录（会自动创建，并自动避免把自身再扫描进入）",

        "scope": "匹配范围：files=只匹配文件，dirs=只匹配目录，all=路径整体匹配（默认 all）",

        "pattern": "可选：对 txt 每一行应用的正则表达式，例如 '(\\d{14})' 用于提取准考证号；不提供时整行参与匹配",

        "mode": (
            "输出结构模式：\n"
            "  flat  = 默认，扁平化输出到同一目录\n"
            "  keep  = 保留原始层级结构\n"
            "  depth = 仅保留指定层级，其余用“路径折叠表示”"
        ),

        "depth": "当 --mode=depth 生效，指定保留的目录层级数（默认 2）"
    }

    p = argparse.ArgumentParser(description=HELP["desc"])
    p.add_argument("-s", "--source", required=True, help=HELP["source"])
    p.add_argument("-t", "--txt", required=True, help=HELP["txt"])
    p.add_argument("-o", "--output", required=True, help=HELP["output"])
    p.add_argument("--scope", choices=["files", "dirs", "all"], default="all",help=HELP["scope"])
    p.add_argument("--pattern", help=HELP["pattern"])
    p.add_argument("--mode", choices=["keep", "depth", "flat"], default="flat", help=HELP["mode"])
    p.add_argument("--depth", type=int, default=2, help=HELP["depth"])

    args = p.parse_args()

    extract(
        args.source,
        args.txt,
        args.output,
        args.scope,
        args.mode,
        args.depth,
        args.pattern
    )

if __name__ == "__main__":
    main()
