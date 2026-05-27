# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

签到时数统计表转换工具。将钉钉/签到平台导出的 xlsx 签到详情，转换为学生会的标准零散时数表（姓名、班级、学号、活动名称、活动日期、课时数）。

## File roles

- **`checkin_hours_app.py`** — 独立 GUI 应用 (customtkinter, v2.0)，核心逻辑内联于此文件，**双击即可运行**。支持自动推断/手动填写两种模式，含明暗主题切换、自定义时数类型和数量。
- **`checkin_core.py`** — 核心处理逻辑模块，导出 `auto_detect(input_path)` 和 `process_checkin(input_path, output_path, activity_name, activity_date, hours, callback)`，可被 GUI/CLI 共同调用。
- **`process_checkin.py`** — 命令行批处理工具，支持 `--auto` 自动推断活动名称和日期，也支持手动指定参数。

## Architecture notes

**签到达人数据格式约定（硬编码在 `raw.iloc[10:]`）：**
- Row 0: 活动标题（col 1）
- Row 1: 活动时间（col 1，含 `YYYY-MM-DD` 格式日期）
- Row 0–9: 元数据/列头行
- Row 10+: 人员数据行
  - Col 0: 姓名, Col 2: 部门（含班级，格式 `xxx-班级`）, Col 4: 工号（学号）, Col 6: 签到时间/状态

**核心处理流程：**
1. 读取 xlsx（`header=None`）
2. 从 row 10 开始提取数据
3. 筛选已签到人员（col 6 非空且非 "未签到"）
4. `checkin_hours_app.py` 额外过滤 "已隐藏" 学号
5. 提取班级（从部门字段分割 `-` 取最后一段，去除括号内容）
6. 排序（v2.0 先班级后学号，旧版只按学号）
7. 用 openpyxl 写入格式化 xlsx（Arial 字体、居中、黑色细边框）

**两个 GUI 版本的核心差异：**
- `checkin_hours_app.py` (v2.0): 核心逻辑内联，支持明暗主题、自定义时数类型/数量、"已隐藏"过滤、按班级+学号排序
- `checkin_core.py` + 旧 GUI: 逻辑分离到 core 模块，仅预设时数选项，无自定义、无班级排序

## Running

```bash
# GUI (standalone, no dependencies beyond requirements)
python checkin_hours_app.py

# CLI batch processing
python process_checkin.py --input 签到.xlsx --output 输出.xlsx --auto
python process_checkin.py --input 签到.xlsx --output 输出.xlsx \
  --activity-name "3.11海亮校园宣讲" --activity-date "2026.03.11" \
  --hours "就业指导2课时"

# Use core module programmatically
python -c "from checkin_core import process_checkin; ..."
```
