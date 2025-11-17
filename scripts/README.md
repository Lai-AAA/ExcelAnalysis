# Python辅助脚本使用说明

## 环境要求

- Python 3.8+
- 依赖库：pandas, openpyxl

## 安装依赖

```bash
pip install -r requirements.txt
```

## 脚本说明

### 1. data_parser.py - 数据解析与考勤验证

**功能**：批量解析多个Excel数据源，自动完成考勤验证和黑名单生成

**使用方法**：
```bash
python data_parser.py --input ./data --config config.json --output ./output
```

**参数说明**：
- `--input`: 输入文件路径或文件夹路径（支持批量处理）
- `--config`: 配置文件路径（JSON格式，可选，默认使用内置配置）
- `--output`: 输出目录（默认：./output）

**输出文件**：
- `parsed_data.xlsx`: 解析后的有效数据（已考勤）
- `blacklist.xlsx`: 未考勤人员黑名单
- `report.txt`: 统计报告

### 2. data_filter.py - 多维度数据筛选

**功能**：按指定维度筛选数据，支持复杂条件组合

**使用方法**：
```bash
python data_filter.py --input ./output/parsed_data.xlsx --config filter_config.json --output ./output
```

**参数说明**：
- `--input`: 输入文件路径（parsed_data.xlsx）
- `--config`: 筛选条件配置文件路径（JSON格式，必需）
- `--output`: 输出目录（默认：./output）

**配置文件示例**（filter_config.json）：
```json
{
  "semester": "2024-2025学年第X学期",
  "activityTypes": ["活动", "讲座"],
  "activityName": "XXX活动",
  "scoreType": "学业分",
  "minScore": 1,
  "maxScore": 3,
  "classes": ["班级1", "班级2"],
  "includeOtherCollege": true,
  "sortByClass": true,
  "useRegex": false
}
```

**输出文件**：
- `filtered_data.xlsx`: 筛选后的候选加分名单
- `filter_log.txt`: 筛选日志

### 3. list_generator.py - 加分名单生成

**功能**：按模板生成标准加分名单Excel，严格遵循格式规范

**使用方法**：

活动/讲座类：
```bash
python list_generator.py --input ./output/filtered_data.xlsx --template activity --semester "2024-2025学年第X学期" --activity "XXX活动" --scoreType "学业分" --score 1 --output ./output
```

比赛类：
```bash
python list_generator.py --input ./output/filtered_data.xlsx --template competition --semester "2024-2025学年第X学期" --activity "XXX比赛" --scoreType "学业分" --output ./output
```

**参数说明**：
- `--input`: 输入文件路径（filtered_data.xlsx）
- `--template`: 模板类型（activity 或 competition）
- `--semester`: 学年学期
- `--activity`: 活动名称
- `--scoreType`: 加分类型（学业分 或 品德分）
- `--score`: 加分分数（仅活动/讲座类需要）
- `--output`: 输出目录（默认：./output）

**输出文件**：
- 标准格式的加分名单Excel文件（按模板命名规则命名）

## 完整工作流程示例

```bash
# 1. 解析数据并验证考勤
python data_parser.py --input ./data --output ./output

# 2. 筛选数据
python data_filter.py --input ./output/parsed_data.xlsx --config filter_config.json --output ./output

# 3. 生成加分名单
python list_generator.py --input ./output/filtered_data.xlsx --template activity --semester "2024-2025学年第X学期" --activity "XXX活动" --scoreType "学业分" --score 1 --output ./output
```

## 注意事项

1. 输入Excel文件必须包含"数据源"工作表
2. 数据源必须包含15个核心字段（见需求文档）
3. 配置文件使用UTF-8编码
4. 输出文件与前端导出文件完全兼容，可交叉使用

