# 签到时数统计表转换工具

将钉钉/签到平台导出的 xlsx 签到详情，一键转换为学生会标准的零散时数统计表。

## 功能

- **自动推断** 活动名称和日期，无需手动填写
- **手动模式** 应对特殊活动命名
- 自动过滤未签到和已隐藏人员
- 按班级+学号排序，输出格式化 xlsx（边框、居中、Arial 字体）
- 明暗主题切换
- 自定义时数类型和数量

## 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行 GUI

```bash
python checkin_hours_app.py
```

### 命令行批处理

```bash
# 自动推断活动信息
python process_checkin.py --input 签到详情.xlsx --output 输出.xlsx --auto

# 手动指定参数
python process_checkin.py --input 签到详情.xlsx --output 输出.xlsx \
  --activity-name "3.11海亮校园宣讲" --activity-date "2026.03.11" \
  --hours "就业指导2课时"
```

### 作为模块调用

```python
from checkin_core import auto_detect, process_checkin

# 自动推断活动信息
info = auto_detect("签到详情.xlsx")

# 执行转换
result = process_checkin(
    input_path="签到详情.xlsx",
    output_path="输出.xlsx",
    activity_name=info["activity_name"],
    activity_date=info["activity_date"],
    hours="就业指导2课时",
)
print(f"处理完成，共 {result.count} 条记录")
```

## 输入格式

程序读取签到平台导出的 xlsx 文件，预期格式：
- 第 1-10 行为元数据/列头
- 第 11 行起为人员签到数据，包含姓名、部门（班级）、工号（学号）、签到状态

## 依赖

- Python 3.9+
- pandas ≥ 2.0
- openpyxl ≥ 3.1
- customtkinter ≥ 5.2
