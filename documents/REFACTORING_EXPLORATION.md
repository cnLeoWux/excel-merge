# 现有架构重构探索建议

> 本文档记录一次架构探索结论，用于后续 OpenSpec proposal、设计讨论或渐进式重构规划。本文只描述方向，不代表已实施。

## 结论

现有架构最值得优先重构的地方，不是立刻拆分 `utils.py`，而是先收敛：

1. 行为契约
2. 三个入口的编排逻辑
3. 核心业务模块边界

更推荐的顺序是：

```text
契约先行
   │
   ▼
OpenSpec / 文档 / 测试对齐
   │
   ▼
抽统一 workflow/service 层
   │
   ▼
拆分 utils.py
   │
   ▼
纯化匹配引擎
   │
   ▼
再考虑性能优化和项目结构整理
```

当前最大风险不是“文件太大”，而是同一业务能力被 CLI、交互入口、Flask API 以不同方式编排、输出、写回和处理错误。

## 当前架构概览

```text
                    ┌──────────────┐
                    │    cli.py    │
                    │ 命令行入口   │
                    └──────┬───────┘
                           │
┌────────────────┐         │         ┌────────────────────┐
│ excel_merge.py │─────────┼────────▶│      utils.py       │
│ 交互式入口     │         │         │ 核心业务大单体       │
└────────────────┘         │         │                    │
                           │         │ - 文件读取           │
┌────────────────────┐     │         │ - 编码判断           │
│ excel_merge_api.py │─────┘         │ - 匹配算法           │
│ Flask API 入口     │               │ - 写回文件           │
└────────────────────┘               │ - 销售报表账期       │
                                     │ - 日期解析           │
                                     │ - 工作流编排         │
                                     └────────────────────┘
```

`utils.py` 是明显的核心重力井，但入口层也承担了过多业务编排职责。

## 重构优先级

### P0：先修正契约漂移

优先明确 OpenSpec、文档、测试、实现之间的不一致。需要先回答：

- CLI 到底使用 `--month YYYYMM`，还是位置参数？
- 是否正式支持 `--match-only` / `--mark-only`？
- 无 month 时 CLI 是否允许交互输入？
- `/merge/json` 返回 `ok/data/error`，还是 `success`？
- API 月报是否生成独立 report 文件，而 CLI 不生成？
- `.xls` 到底是否支持原地写回？
- `process_excel_files()` 是否只负责匹配手续费，还是允许顺手标记销售报表账期？

如果不先定清楚，后续拆模块会把当前不一致扩散到更多文件中。

### P1：抽统一 workflow/service 层

这是最高 ROI 的结构性重构。

目标是让三个入口变薄：

```text
              ┌──────────────┐
              │    cli.py    │
              └──────┬───────┘
                     │
┌────────────────┐   │   ┌────────────────────┐
│ excel_merge.py │───┼──▶│  workflow/service  │
└────────────────┘   │   └─────────┬──────────┘
                     │             │
┌────────────────────┐             ▼
│ excel_merge_api.py │      ┌──────────────┐
└────────────────────┘      │ core modules │
                            └──────────────┘
```

入口层只负责：

- 参数解析
- CLI 文本 / JSON 输出
- HTTP 请求 / 响应
- 交互式文件选择
- exit code 或 HTTP status 映射

统一 workflow/service 层负责：

- `run_merge(order_file, payment_file)`
- `run_sales_report(order_file, payment_file, month)`
- 统计结果构建
- 错误模型归一化
- 写回策略协调
- 面向入口层的稳定 Result 对象

收益：

- CLI、交互模式、API 行为更一致
- JSON 输出和统计字段更统一
- 错误处理集中
- 后续测试更容易
- `utils.py` 拆分方向更清楚

### P2：再拆分 `utils.py`

`utils.py` 确实值得拆，但应在 workflow 边界明确后进行。建议第一轮只机械搬迁，不改变行为。

候选拆分：

```text
file_io.py
  read_file_with_appropriate_method()
  write_result_file()
  find_file_path()

matching.py
  extract_p_number()
  match_orders_by_p_number()
  match_payment_fees()
  process_excel_files()  # 可作为兼容 wrapper

sales_report.py
  add_sales_report_period()
  filter_unmarked_and_generate_report()
  process_sales_report_workflow()

date_utils.py
  parse_date()
  get_year_month()

backup.py
  auto_backup()

workflow.py
  面向 CLI / API / 交互入口的统一编排

utils.py
  暂时保留兼容导出，降低迁移风险
```

重点原则：结构变化和行为变化不要混在同一步里。

### P3：纯化匹配引擎

当前核心匹配函数同时承担文件 IO、列处理、匹配、手续费赋值和部分销售报表职责。更理想的方向是让匹配算法变成纯 DataFrame 逻辑：

```text
read_file(order_file)       read_file(payment_file)
        │                           │
        ▼                           ▼
    order_df                    payment_df
        │                           │
        └────────────┬──────────────┘
                     ▼
          match_payment_fees()
                     │
                     ▼
              result_df
```

目标形态：

```python
def match_payment_fees(order_df, payment_df, verbose=False):
    ...
```

再保留文件路径 wrapper：

```python
def process_excel_files(order_file, payment_file, verbose=False):
    order_df = read_file_with_appropriate_method(order_file)
    payment_df = read_file_with_appropriate_method(payment_file)
    return match_payment_fees(order_df, payment_df, verbose)
```

这样核心业务测试可以脱离文件系统，边界情况也更容易覆盖。

### P4：统一文件写回和持久化策略

需要明确：

- 是否自动备份
- 是否原子写
- 写入失败是否保证原文件不损坏
- CSV 编码策略
- `.xls` 是否支持写回
- API 下载 MIME 类型如何匹配真实文件格式
- 找不到文件时 `find_file_path()` 返回 `None` 还是原始路径

这些属于产品契约，应先进入 OpenSpec，再进入实现。

### P5：最后再做性能优化和项目结构整理

匹配逻辑可能存在 O(n×m) 性能问题，但不建议最先优化。

原因是匹配规则较业务化：

```text
精确 20 字符匹配
      │
      ▼
P-number 匹配
      │
      ▼
连字符 fallback
      │
      ▼
业务类型校验
  正金额 => 收费 / 服务费
  负金额 => 退费 / 退款
```

贸然向量化或索引化容易改变边界行为。更稳妥路线：

```text
characterization tests
        │
        ▼
纯化匹配函数
        │
        ▼
锁住行为
        │
        ▼
优化性能
```

`src/` 包结构、`pyproject.toml`、清理根目录脚本等工程卫生工作也有价值，但应排在行为边界稳定之后。

## 不建议优先做的事情

### 不建议一开始大拆 `utils.py`

直接拆会让当前不一致扩散。应先明确 workflow 边界，再拆职责模块。

### 不建议先做性能优化

性能优化可能影响核心匹配规则，除非已有明确大文件性能痛点，否则应先补测试和纯化逻辑。

### 不建议先重写 Flask API

API 的问题可以通过统一 workflow/service 层自然缓解。先重写 API 可能继续复制现有业务混乱。

### 不建议先包化项目结构

包结构整洁不等于行为一致。应先解决入口编排和契约问题。

## 推荐渐进式路线

```text
阶段 0：探索 / 决策
  明确 CLI、API、销售报表、文件写回的真实契约

阶段 1：OpenSpec 对齐
  修正 cli-input / cli-output / http-api / file-io / sales-report spec

阶段 2：补 characterization tests
  锁住当前关键行为，尤其是边界案例

阶段 3：抽 workflow/service 层
  让 CLI、交互入口、API 都调用同一套编排

阶段 4：拆 utils.py
  机械搬迁，不改行为

阶段 5：纯化匹配引擎
  文件 IO 和 DataFrame 匹配逻辑分离

阶段 6：性能优化、包结构整理、API 清理
```

## 最值得单独立项的 change

如果只选一个重构 change，建议是：

> 新增统一 workflow/service 层，并让 CLI、交互入口、Flask API 逐步变成薄入口。

该 change 的核心问题是明确：

- 谁负责业务编排？
- 谁负责输出格式？
- 谁负责文件写回？
- 谁负责错误模型？
- 谁负责统计结果？/

这些边界清楚后，`utils.py` 的拆分会更自然、更低风险。

## 适合后续创建的 OpenSpec 方向

可以考虑拆成多个较小 change：

1. `align-cli-api-contracts`
   - 对齐 CLI/API 输出、参数、错误格式、销售报表行为。

2. `introduce-workflow-service-layer`
   - 引入统一 workflow/service 层，入口层变薄。

3. `split-utils-by-responsibility`
   - 按 file_io、matching、sales_report、date_utils、backup 拆分 `utils.py`。

4. `pure-dataframe-matching-engine`
   - 将匹配引擎纯化为 DataFrame in / DataFrame out。

5. `define-safe-persistence-policy`
   - 明确备份、原子写、`.xls`、CSV 编码、API 下载格式。

## 后续探索问题

- CLI 是否应彻底禁止自动化场景下的交互输入？
- API 和 CLI 是否必须共享完全相同的 JSON envelope？
- API 生成 report 文件而 CLI 不生成，是否是刻意差异？
- `process_excel_files()` 当前是否应该被视为“历史兼容 API”？
- 销售报表账期标记是否应从普通匹配流程中移出？
- `.xls` 原地写回是否值得继续支持？
- 是否需要引入明确的 `Result` / `Error` 数据结构？
