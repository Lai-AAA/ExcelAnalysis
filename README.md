
# Vue加分名单生成系统

## 项目简介

Vue加分名单生成系统是一个纯前端实现的Excel数据处理工具，用于工学院学生加分统计。系统支持上传Excel数据源，通过考勤验证、数据筛选，按标准模板导出活动/讲座类、比赛类加分名单，同时生成未考勤人员黑名单。

## 核心特性

- ✅ 纯前端实现，无后端依赖
- ✅ 支持多Excel数据源上传和批量处理
- ✅ 自动考勤验证和黑名单生成
- ✅ 多维度数据筛选
- ✅ 标准模板格式导出（活动/讲座类、比赛类）
- ✅ Python脚本辅助处理大数据量（≥3000行）

## 技术栈

- **前端框架**: Vue 3 (Vite)
- **UI组件库**: Element Plus
- **Excel处理**: SheetJS (xlsx)
- **样式方案**: Scss
- **Python脚本**: Python 3.8+, Pandas, OpenPyXL

## 快速开始

### 环境要求

- Node.js 16+
- npm 或 yarn

### 安装依赖

```bash
npm install
```

### 开发运行

```bash
npm run dev
```

访问 http://localhost:3000

### 构建生产版本

```bash
npm run build
```

构建产物在 `dist` 目录，可直接部署到静态服务器。

## 使用流程

### 1. 上传文件

- 支持 `.xlsx` 格式
- 单个文件不超过10MB
- 最多同时上传3个文件
- 文件必须包含"数据源"工作表
- 数据源必须包含15个核心字段（见下方字段说明）

### 2. 数据解析

系统自动：
- 验证必需字段
- 去重处理（按"班级+姓名+活动名称"）
- 格式转换和有效性校验
- 显示统计信息

### 3. 考勤验证

- 配置考勤关键词（默认：已考勤、到场、参与）
- 自动区分已考勤（可加分）和未考勤（黑名单）人员
- 支持导出黑名单

### 4. 数据筛选

支持多维度筛选：
- 学年学期
- 活动类型（活动、讲座、比赛）
- 活动名称（模糊搜索）
- 加分类型（学业分、品德分）
- 加分分数范围
- 行政班级（多选）
- 是否包含其他学院人员
- 排序选项

### 5. 导出名单

- 选择模板类型（活动/讲座类 或 比赛类）
- 填写活动信息
- 预览数据
- 导出标准格式Excel文件

## 数据源格式要求

### 必需字段（15个）

1. 序号
2. 学年学期
3. 是否为工学院举办
4. 行政班级
5. 学号
6. 姓名
7. 活动类型
8. 活动名称
9. 加分类型
10. 加分分数
11. 奖项
12. 部门
13. 负责人
14. 联系电话
15. 备注（用于考勤验证）

### Excel文件结构

- 工作表名称：**数据源**（必须）
- 第一行为字段名
- 第二行开始为数据

## Python脚本使用

针对大数据量场景（≥3000行），可使用Python脚本提升处理效率。

### 安装Python依赖

```bash
cd scripts
pip install -r requirements.txt
```

### 脚本说明

1. **data_parser.py** - 数据解析与考勤验证
2. **data_filter.py** - 多维度数据筛选
3. **list_generator.py** - 加分名单生成

详细使用说明请参考 [scripts/README.md](./scripts/README.md)

## 项目结构

```
ExcelDemo/
├── src/                    # 源代码目录
│   ├── views/             # 页面组件
│   │   ├── Home.vue       # 首页（文件上传）
│   │   ├── Parse.vue      # 解析页（数据解析与考勤验证）
│   │   ├── Filter.vue     # 筛选页（数据筛选）
│   │   └── Export.vue     # 导出页（名单导出）
│   ├── utils/             # 工具函数
│   │   ├── excel.js       # Excel处理
│   │   ├── templates.js   # 模板配置
│   │   ├── filter.js      # 数据筛选
│   │   └── storage.js     # 本地存储
│   ├── router/            # 路由配置
│   ├── styles/            # 样式文件
│   ├── App.vue            # 根组件
│   └── main.js            # 入口文件
├── scripts/               # Python脚本
│   ├── data_parser.py     # 数据解析脚本
│   ├── data_filter.py     # 数据筛选脚本
│   ├── list_generator.py  # 名单生成脚本
│   ├── requirements.txt   # Python依赖
│   └── README.md          # 脚本使用说明
├── index.html             # HTML模板
├── vite.config.js         # Vite配置
├── package.json           # 项目配置
└── README.md              # 项目说明
```

## 浏览器兼容性

- Chrome 90+
- Edge 90+
- Firefox 88+

## 性能要求

- 前端解析1000-3000行数据耗时≤3秒
- Python脚本解析≥3000行数据耗时≤1秒

## 常见问题

### 1. 文件上传失败

- 检查文件格式是否为 `.xlsx`
- 检查文件大小是否超过10MB
- 检查文件是否包含"数据源"工作表

### 2. 字段缺失错误

- 检查数据源是否包含所有15个必需字段
- 检查字段名称是否完全匹配（注意空格）

### 3. 导出格式不正确

- 确保使用Microsoft Excel 2016+或WPS打开
- 检查模板类型是否选择正确

### 4. Python脚本运行失败

- 检查Python版本是否为3.8+
- 检查是否安装了所有依赖：`pip install -r requirements.txt`
- 检查输入文件路径是否正确

## 维护说明

- 数据源字段变更：修改 `src/utils/excel.js` 中的 `REQUIRED_FIELDS`
- 考勤规则变更：修改考勤验证逻辑或配置文件
- 模板扩展：在 `src/utils/templates.js` 中添加新模板配置

## 许可证

MIT License

## 联系方式

如有问题或建议，请联系项目维护者。


# ExcelAnalysis
excel文件分析

