# excel-batch-tools

Batch Excel cleaning, merging and sheet splitting tools based on Python.

## 1. 安装

```
pip install -r requirements.txt
```

环境建议：

* Python 3.8+
* Windows / Mac / Linux

---

## 2. 当前功能

### clean（清洗工具）

* 支持单文件
* 支持文件夹批量处理
* 支持列缺失阈值删除
* 支持行缺失阈值清洗（可保护指定列不删除）
* 自动 dirty 记录输出（记录清洗掉的行列）
* 支持命令行参数
* 支持日志输出
* 未指定输出时自动生成 `clean_xxx`

---

### 2.1 clean 使用方法

#### 单文件模式

```sh
python clean.py -i test.xlsx -s 问卷 -o cleaned.xlsx
```

#### 批量模式

```sh
python clean.py \
 -i excel_folder \
 -s sheet_name \
 -o cleaned_folder \
 --dirty dirty_folder \
 --col-threshold 0.7 \
 --row-threshold 0.7 \
 --protect-cols A-F
```

示例

```sh
python clean.py -i ./data -s "sheet_name" -o ./clean_result --dirty ./dirty_log
```

---

## 3. append（数据追加工具）

用于将源 Excel 的 sheet 追加到目标 Excel 指定 sheet
支持：

* 单文件模式
* 文件夹批量模式
* 支持表头匹配模式：

  * `header-intersection` 仅保留公共列
  * `header-union` 自动按并集合并
* 支持无表头直接追加模式
* 自动跳过不存在指定 sheet 的源文件

---

### 3.1 append 使用方法

基本语法

```sh
python append_excel.py 
    -s 源文件或文件夹
    -ss 源sheet
    -t 目标文件
    -ts 目标sheet
    -m no-header | header-intersection | header-union
```

单文件

```sh
python append_excel.py -s a.xlsx -ss "source_sheet_name" -t result.xlsx -ts "target_sheet_name" -m header-union
```

批量

```sh
python append_excel.py -s ./data -ss "source_sheet_name" -t result.xlsx -ts "target_sheet_name" -m header-intersection
```

---

## 4. split_sheets（Sheet 分离工具）

功能说明：

* 读取 Excel 的所有 sheet
* 在输出目录下为每个 sheet 建立一个同名文件夹
* 按源文件分别导出
* 支持：

  * 文件模式
  * 文件夹批量模式
  * 文件名模式选择：

    * 保持源文件名
    * 使用 sheet 名（重复自动编号）

---

### 4.1 split_sheets 使用方法

单文件

```sh
python split_sheets.py -s data.xlsx -o ./out
```

批量文件夹

```sh
python split_sheets.py -s ./data -o ./split_out
```

使用 sheet 命名模式

```sh
python split_sheets.py -s ./data -o ./split_out --name-mode sheet
```

输出结构示例

```
/output
   /Sheet1
       fileA.xlsx
       fileB.xlsx
   /Sheet2
       fileA.xlsx
```


## 5. extract（文件迁移工具）
功能说明：
- 递归扫描源目录
- txt 每行用 自定义正则 模糊匹配（可指定提取规则）
- 复制到目标目录时可以指定下面三种规则
    - keep —— 保持原目录层级
    - depth —— 用“路径符号”折叠为有限层（可指定 --depth N）
    - flat —— 默认，全部扁平放到同一层
- 匹配文件 → 复制文件，匹配文件夹 → 递归复制内部结构，可通过 `--scope` 指定匹配对象：
    - `files`：只匹配文件名（默认）
    - `dirs`：只匹配文件夹
    - `all`：整个相对路径匹配

---

### 5.1 extract 使用方法
基础用法

```sh
python extract.py \
  -s ./source_root \
  -t ids.txt \
  -o ./output \
  --scope all \
  --pattern "(\\d{14})" \
  --mode depth \
  --depth 2
```

```
-s        源目录
-t        txt 文件
-o        输出目录
--scope   files | dirs | all
--pattern 正则（作用于每一行 txt）
--mode    keep | depth | flat(默认)
--depth   depth模式下保留多少层（默认 2）
```

匹配文件夹名

```sh
python extract.py -s ./data -t ids.txt -o ./output --scope dirs
```

对整个路径模糊匹配

```sh
python extract.py -s ./data -t ids.txt -o ./output --scope all
```
