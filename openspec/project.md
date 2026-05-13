# Project Context

## Purpose

Excel Merge Tool 是一个基于 Python 的工具，用于将订单 Excel/CSV 文件与支付流水文件进行自动匹配，填充“支付手续费”列。支持三种入口模式：交互式命令行、CLI 参数模式、Flask HTTP API。另提供销售报表账期标注和月度报表生成功能。

主要使用场景：
- 旅游行业订单数据与第三方支付流水的对账。
- AI Agent / 自动化脚本的非交互式批处理（通过 `--json --quiet`）。
- 内部团队通过 Web 界面上传文件进行临时合并。

## Tech Stack

- **语言**：Python 3.7+
- **数据处理**：pandas>=1.3.0（`on_bad_lines` 参数需要 >=1.3）
- **Excel 引擎**：openpyxl>=3.0.0（.xlsx 读写）、xlrd>=2.0.0（.xls 读取）
- **HTTP 服务**：Flask>=2.0.0、werkzeug>=2.0.0
- **标准库**：pathlib、re（P-number 提取）、zipfile（Excel 引擎检测）、datetime、logging、argparse
- **打包**：setuptools（`setup.py` 提供 `excel-merge` / `excel-merge-cli` console_scripts）

## Project Conventions

### Code Style

- 遵循 PEP 8，**行长度 ≤ 79 字符**，4 空格缩进，`snake_case` 命名函数和变量。
- import 顺序：标准库 → 第三方 → 本地（绝对导入）。
- 多行表达式：优先使用圆括号包裹，避免使用 `\` 行 continuation。
- 函数签名使用类型提示（`Optional[str]`, `pd.DataFrame`, `Path`）。
- `verbose: bool = False` 模式用于可选调试输出。
- 业务列名使用中文（订单号、商户订单号、支付手续费等）。
- 金额列名包含全角括号：`支出金额（-元）`、`收入金额（+元）`，必须精确匹配。
- NaN 值必须使用 `pd.isna()` / `pd.notna()` 检查，**禁止与字符串 `"nan"` 直接比较**。

### 架构模式

- **单核心模块**：`utils.py` 集中所有业务逻辑（~930 行，匹配/读写/报表）。三个入口文件均调用 utils 函数：

  ```
  cli.py ──────────┐
  excel_merge.py ──┼──→ utils.py
  excel_merge_api.py┘
  ```

  核心函数表：

  | Function | Line | Purpose |
  |----------|------|---------|
  | `extract_p_number()` | 15 | Regex `r"P\d+"` 提取 |
  | `process_excel_files()` | 189 | 主匹配循环（exact→P-number→hyphen） |
  | `add_sales_report_period()` | 572 | 标注全退/已取消 |
  | `filter_unmarked_and_generate_report()` | 743 | 阶段二筛选（内存操作，无文件输出） |
  | `process_sales_report_workflow()` | 887 | 端到端销售报表工作流 |

- **入口职责分离**：
  - `cli.py`（argparse + JSON envelope，`main_cli()`，console script `excel-merge-cli`）
  - `excel_merge.py`（交互式 + 非交互式，`main()`，console script `excel-merge`）
  - `excel_merge_api.py`（Flask 端点）
- **文件查找**：先搜索 cwd，再搜索 `ExcelForHandel/` 目录。
- **CLI/API 输出契约**：
  - JSON 信封：`{ok: bool, data: {...}, error: str|null}`
  - 语义化退出码：0=成功，1=一般错误，2=使用错误（含传递已移除的 flags），3=文件未找到，4=处理错误
  - stdout 输数据，stderr 输日志
- **编码回退链**：`gbk → utf-8 → gb2312 → latin-1 → utf-8-sig`。
- **CSV 分隔符重试**：`"," → ";" → "\t" → sep=None`。
- **Excel 引擎选择**：.xlsx 用 zipfile 探测 → openpyxl（成功）/ xlrd（BadZipFile）；.xls 始终用 xlrd。
- **匹配算法优先级**：
  1. 订单号前 20 字符 ↔ 商户订单号（精确匹配）
  2. P-number 匹配：`r"P\d+"` 从 `外部订单号` ↔ `商品名称` 提取后比较，**区分大小写**（不匹配小写 `p`）
  3. 连字符匹配：`外部订单号` 各部分 ↔ `商品名称` 最后 `-` 后的段
  4. 业务类型校验：正单（金额>0）↔ 收费/服务费；退单（金额<0）↔ 退费/退款
  5. 金额赋值：正单→`支出金额（-元）`；退单→`收入金额（+元）`；零金额→`支付手续费=0.0`
- **销售报表两阶段**：阶段一匹配 + 标注（全退/已取消）；阶段二筛选未标注 + 1 年出行日期窗口 + 在原 DataFrame 中标注“销售报表YYYYMM”，由调用方就地写回订单文件（CLI 不再生成独立的 `report_YYYYMM.xlsx`；HTTP API 内部仍可落盘以提供下载）。

### 测试策略

-- **测试框架**：pytest（`pip install -r requirements-dev.txt`），测试目录 `tests/`。
-- **现状**：仓库根的 `test_*.py` 是手动验证脚本（仅 `print`，无断言），**不是真正的 pytest 测试套件**；无 CI/CD、无覆盖率工具。
-- **手动验证**：`verify_result.py`、`verify_original.py` 用于人工对比输出与原始数据。
-- **新功能验证要求**：
  - 涉及匹配逻辑的变更必须覆盖正单、退单、零金额三种情况。
  - 涉及文件读取的变更必须同时验证 CSV 和 Excel 路径。
  - CLI 行为变更需运行 `python cli.py order.xlsx payment.xlsx` 与 `--json --quiet` 两种模式。

### Git 工作流

- 新功能/bugfix 在特性分支开发（如 `feature/excel-merger-new-feature`），主分支保持稳定。
- 提交信息描述对匹配逻辑、文件处理、CLI 输出契约或错误修复的具体改动。
- 不在主分支直接提交未经验证的破坏性变更；涉及 CLI 输出格式或退出码的修改需要同步更新 `openspec/specs/cli-output/spec.md`。

## Domain Context

- **业务领域**：旅游行业的订单数据与第三方支付流水对账。
- **20 字符订单号截取**：支付服务商的硬编码格式（订单号前 20 字符 = 商户订单号前 20 字符）。
- **P-number 格式**：`P` 后跟数字（如 `P2507021103060001`），用于外部订单号与商品名称的关联。**区分大小写**，正则 `r"P\d+"` 不匹配小写 `p`。
- **业务类型分类**：
  - 收费 / 服务费 = 正常扣款（与正单匹配）
  - 退费 / 退款 = 退还手续费（与退单匹配）
- **销售报表标注**：
  - **全退**：相同订单号出现多次且金额之和为 0
  - **已取消**：订单状态字段包含"取消" **且** 订单金额为 0

## Important Constraints

- 必须处理多种编码的输入文件（GBK 优先，覆盖 UTF-8、GB2312、Latin-1、UTF-8-sig）。
- 订单号列必须保留为字符串类型，防止 Excel 数字转换或科学计数法。
- NaN 值必须使用 `pd.isna()` / `pd.notna()` 检查，禁止与字符串 `"nan"` 直接比较。
- 默认行为是**就地修改原始订单文件**（CLI 已移除 `-o`/`--output`/`--output-dir`，无法重定向）；任何破坏该默认行为的变更必须在 spec 中明确标注。
- 支持 Excel（.xlsx、.xls）和 CSV 文件格式，禁止删除任一格式支持。
- 金额列名使用全角括号 `（-元）` / `（+元）`，禁止替换为半角。
- CLI 退出码 0/1/2/3/4 是对外契约；新增退出码必须更新 `cli-output` spec。
- JSON 信封 `{ok, data, error}` 是对外契约；新增字段允许，删除字段视为破坏性变更。

## External Dependencies

- **pandas** — Excel/CSV 数据处理与 DataFrame 操作（核心运行时依赖）。
- **openpyxl** — 读写 .xlsx 文件。
- **xlrd** — 读取 .xls 文件（仅老格式，不支持 .xlsx 写入）。
- **Flask** — HTTP API 服务框架（`excel_merge_api.py`）。
- **werkzeug** — Flask 的 WSGI 工具库（`secure_filename` 等）。
- **标准库**：pathlib、re、zipfile、datetime、logging、argparse、json、sys。
