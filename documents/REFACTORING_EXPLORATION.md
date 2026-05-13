# 现有架构重构探索建议

> 本文档记录架构探索结论，并已按当前实现状态更新：`workflow_service.py` 已落地为共享 service 层；后续重构重点转为在保持行为不变的前提下拆分 `utils.py`。

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
已抽出 workflow_service.py
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

当前最大风险已经从“入口层重复编排”转为“核心业务仍集中在 `utils.py`，拆分时可能改变历史行为”。后续所有结构调整都应先用 OpenSpec 和测试锁定契约。

## 当前架构概览

```text
                    ┌──────────────┐
                    │    cli.py    │
                    │ CLI adapter  │
                    └──────┬───────┘
                           │
┌────────────────┐         │         ┌────────────────────┐
│ excel_merge.py │─────────┼────────▶│ workflow_service.py │
│ 交互式入口     │         │         │ 应用编排/统计/错误   │
└────────────────┘         │         └─────────┬──────────┘
                           │                   │
┌────────────────────┐     │                   ▼
│ excel_merge_api.py │─────┘         ┌────────────────────┐
│ Flask API 入口     │               │      utils.py       │
└────────────────────┘               │ 核心业务大单体       │
                                     │ - 文件读取           │
                                     │ - 编码判断           │
                                     │ - 匹配算法           │
                                     │ - 写回文件           │
                                     │ - 销售报表账期       │
                                     │ - 日期解析           │
                                     └────────────────────┘
```

`workflow_service.py` 已减轻入口层编排压力；`utils.py` 仍是核心重力井。

## 重构优先级

### P0：先修正契约漂移

优先明确 OpenSpec、文档、测试、实现之间的不一致。当前已定契约：

- `cli.py` 使用位置参数 `target_month` 触发完整工作流。
- `--match-only` / `--mark-only` 是正式 reduced workflow，当前仍要求提供 `target_month`。
- `/merge/json` 使用 API shape：`success`、`download_url`、`statistics`、`files`；CLI JSON 使用 `ok/data/error`。
- API 月报模式生成独立可下载 report 文件；CLI 月报工作流不生成独立报表文件，只就地写回订单文件。
- `.xls` 读取受支持；写回是否可用取决于 pandas/writer 环境，失败应向上抛错。
- `process_excel_files()` 当前仍会刷新 `销售报表账期`，这是历史兼容副作用。

如果不先定清楚，后续拆模块会把当前不一致扩散到更多文件中。

### P1：维护并收敛 workflow/service 层

该层已经实现，后续重点是保持边界清晰。

当前目标是让三个入口继续保持变薄：

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

- `run_match_only(order_file, payment_file)`
- `run_mark_only(order_file)`
- `run_sales_report(order_file, payment_file, month)`
- `prepare_api_merge(...)`
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
  parse_date()
  get_year_month()
  filter_unmarked_and_generate_report()
  process_sales_report_workflow()

utils.py
  暂时保留兼容导出，降低迁移风险
```

当前 OpenSpec change `refactor-core-workflow-boundaries` 采用的目标模块正是 `file_io.py`、`matching.py`、`sales_report.py`，并保留 `utils.py` 作为兼容 facade。

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

已定事实：CLI 已移除另存为参数，主输出路径为原订单文件；`cli.py` 当前会在调用 service 前创建备份。API 仍生成 `results/` 下的可下载文件。

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

阶段 3：维护 workflow/service 层（已落地）
  让 CLI、交互入口、API 都调用同一套编排，并防止 adapter 输出格式下沉

阶段 4：拆 utils.py
  机械搬迁，不改行为

阶段 5：纯化匹配引擎
  文件 IO 和 DataFrame 匹配逻辑分离

阶段 6：性能优化、包结构整理、API 清理
```

## 最值得单独立项的 change

如果只选下一个重构 change，建议是：

> `refactor-core-workflow-boundaries`：在 `workflow_service.py` 已存在的前提下，先补行为锁定测试，再按职责拆分 `utils.py`。

该 change 的核心问题是明确：

- 谁负责业务编排？
- 谁负责输出格式？
- 谁负责文件写回？
- 谁负责错误模型？
- 谁负责统计结果？/

这些边界目前已初步清楚，下一步重点是防止拆分 `utils.py` 时改变匹配优先级、CSV fallback 顺序或 CLI/API 输出契约。

## 适合后续创建的 OpenSpec 方向

可以考虑拆成多个较小 change：

1. `refactor-core-workflow-boundaries`
   - 当前活跃方向：行为锁定测试 + `file_io.py` / `matching.py` / `sales_report.py` 拆分 + `utils.py` facade。

2. `split-utils-by-responsibility`
   - 若需要更小粒度，可从当前 change 中拆出纯迁移任务。

3. `pure-dataframe-matching-engine`
   - 将匹配引擎纯化为 DataFrame in / DataFrame out。

4. `define-safe-persistence-policy`
   - 明确备份、原子写、`.xls`、CSV 编码、API 下载格式。

## 后续探索问题

- CLI 是否应彻底禁止自动化场景下的交互输入？
- API 和 CLI 当前不共享完全相同的 JSON envelope；是否长期保持差异仍可后续讨论。
- API 生成 report 文件而 CLI 不生成，当前应视为刻意契约差异。
- `process_excel_files()` 当前应该被视为“历史兼容 API”。
- 销售报表账期标记是否应从普通匹配流程中移出？
- `.xls` 原地写回是否值得继续支持？
- 是否需要引入明确的 `Result` / `Error` 数据结构？
