# 项目背景

## 项目目的

Excel Merge Tool 是一个基于 Python 的工具，用于将订单 Excel/CSV 文件与支付流水文件进行自动匹配，填充"支付手续费"列。支持三种入口模式：交互式命令行、CLI 参数模式、Flask HTTP API。另提供销售报表账期标注和月度报表生成功能。

## 技术栈

- Python 3.7+
- pandas>=1.3.0（on_bad_lines 参数需要 >=1.3）
- openpyxl>=3.0.0（读写 .xlsx）
- xlrd>=2.0.0（读取 .xls）
- Flask>=2.0.0（HTTP API 服务）
- werkzeug>=2.0.0（Flask 依赖）
- pathlib（文件路径操作）
- re（正则表达式，P-number 提取）
- zipfile（Excel 引擎检测）

## 项目结构

```
./
├── utils.py                # 核心业务逻辑（~930 行）：匹配、读写、报表
├── cli.py                  # CLI 入口（argparse）：order_file payment_file [-o] [--month] [--output-dir]
├── excel_merge.py          # 交互式入口：从 ExcelForHandel/ 选择文件
├── excel_merge_api.py      # Flask API：/merge, /merge/json, /download/<file>, /health
├── setup.py                # 包配置：console_scripts excel-merge & excel-merge-cli
├── requirements.txt        # 运行时依赖
├── ExcelForHandel/         # 输入数据目录
├── documents/              # ARCHITECTURE.md, TECHNICAL_DOCS.md, USAGE_EXAMPLES.md
├── openspec/               # OpenSpec 配置（config.yaml, project.md）
└── test_*.py               # 4 个手动测试脚本（无断言）
```

## 项目约定

### 代码风格

- 遵循 PEP 8，4 空格缩进
- 列名使用业务需求中指定的中文（订单号、商户订单号、支付手续费等）
- 金额列名包含全角括号：`支出金额（-元）`、`收入金额（+元）`
- 函数参数和返回值使用类型提示（`Optional[str]`, `pd.DataFrame`, `Path`）
- `verbose: bool = False` 模式用于可选调试输出
- `snake_case` 命名函数和变量
- import 顺序：标准库 → 第三方 → 本地（绝对导入）

### 架构模式

- 核心逻辑集中在 utils.py（~930 行），三个入口文件（cli.py, excel_merge.py, excel_merge_api.py）均调用 utils 函数
- 依赖关系：
  ```
  cli.py ──────────┐
  excel_merge.py ──┤──→ utils.py
  excel_merge_api.py┘
  ```
- 文件读取支持多种编码，回退链：gbk → utf-8 → gb2312 → latin-1 → utf-8-sig
- CSV 分隔符重试：`,` → `;` → `\t` → `sep=None` 自动检测
- Excel 引擎检测：.xlsx 通过 zipfile 探测选择 openpyxl 或 xlrd；.xls 始终用 xlrd

### 匹配算法（优先级顺序）

1. **精确匹配**：订单号前 20 字符 ↔ 商户订单号前 20 字符（列名通过 "商户"+"订单" 子串匹配）
2. **P-number 匹配**：从外部订单号和商品名称中提取 `r"P\d+"`（区分大小写）进行比较
3. **连字符匹配**：外部订单号各部分 ↔ 商品名称最后一个 "-" 后的部分
4. **业务类型校验**（所有匹配必须通过）：
   - 正单（订单金额 > 0）：支付记录须为"收费"或"服务费"
   - 退单（订单金额 < 0）：支付记录须为"退费"或"退款"
5. **金额赋值**：
   - 正单 → 支出金额（-元）（预期为负值）
   - 退单 → 收入金额（+元）（预期为正值）
   - 零金额 → 支付手续费 = 0.0，跳过匹配

### 销售报表工作流

两阶段处理（通过 CLI `--month YYYYMM` 触发）：
1. **阶段一**：执行订单-支付匹配 + 标注销售报表账期
   - `add_sales_report_period()`：标注"全退"（重复订单号金额之和为 0）和"已取消"（订单状态含"取消"且金额为 0）
2. **阶段二**：筛选未标注行并生成月度报表
   - `filter_unmarked_and_generate_report()`：按出行日期筛选 1 年窗口内的数据，输出 report_YYYYMM.xlsx

### 文件处理

- **默认原地修改**：覆盖原始订单文件（CLI 用 `-o` 可重定向输出）
- **CSV 注释**：以 `#` 开头的行自动跳过，第一个非注释行作为表头
- **CSV 写入**：使用 `utf-8-sig` 编码
- **文件查找**：先搜索当前目录，再搜索 `ExcelForHandel/` 子目录
- **订单号保护**：强制转为字符串类型，防止 Excel 数字转换

### Flask API

- 端点：GET `/`（上传页面）, GET `/health`, POST `/merge`（返回文件）, POST `/merge/json`（返回 JSON + 下载链接）, GET `/download/<filename>`
- 上传文件保存到 `uploads/`，处理结果保存到 `results/`
- 支持格式：.xlsx, .xls, .csv（最大 16MB）

### 测试策略

- 当前为 4 个手动测试脚本（test_*.py），仅打印输出，无断言
- 不是真正的 pytest 测试套件
- 无 CI/CD 配置，无代码覆盖率测量
- 验证脚本（verify_result.py, verify_original.py）用于手动验证输出

### Git 工作流

- 新功能开发使用特性分支
- 主分支用于稳定发布
- 提交信息应描述对匹配逻辑、文件处理或错误修复所做的更改

## 领域上下文

- 工具处理旅游行业的订单数据与支付流水匹配
- 启动交互模式后列出 ExcelForHandel 目录下所有文件供选择
- 订单号截取前 20 字符是支付服务商格式的硬编码业务规则
- P-number 格式：P 后跟数字（如 P2507021103060001），用于外部订单号与商品名称的关联
- 业务类型区分：收费/服务费（正常扣款）vs 退费/退款（退还手续费）
- 全退判定：相同订单号出现多次且金额之和为 0
- 已取消判定：订单状态包含"取消"且订单金额为 0

## 重要约束

- 必须处理多种编码的文件（UTF-8、GBK、GB2312、Latin-1、UTF-8-sig）
- 订单号必须保留为字符串，防止 Excel 数字转换
- 必须正确处理 NaN 值
- 默认原地修改原始文件（不创建新文件）
- 支持 Excel（.xlsx、.xls）和 CSV 文件格式
- P-number 正则 `r"P\d+"` 区分大小写，不匹配小写 p
- 金额列名使用全角括号，必须精确匹配

## 外部依赖

- **pandas**：Excel/CSV 数据处理，DataFrame 操作
- **openpyxl**：读写 .xlsx 文件
- **xlrd**：读取 .xls 文件
- **Flask**：HTTP API 服务框架
- **werkzeug**：Flask 的 WSGI 工具库（安全文件名处理等）
- **标准库**：pathlib（路径）, re（正则）, zipfile（引擎检测）, datetime, logging
