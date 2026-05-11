## MODIFIED Requirements

### Requirement: cli.py positional argument mode

`cli.py` MUST expose the current command-line interface using two required positional file arguments and an optional positional `target_month`. The standard workflow is the full sales-report workflow, so callers SHOULD provide or obtain `target_month` unless the user explicitly requests a reduced mode such as `--match-only`.

#### Scenario: Files without target month enter month acquisition
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json --quiet`
- **THEN** `order.xlsx` 被解析为订单文件
- **AND** `payment.xlsx` 被解析为支付流水文件
- **AND** `target_month` 为 `None`
- **AND** CLI 尝试通过交互式提示获取 `target_month`，而不是默认执行基础匹配

#### Scenario: Sales report positional month
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --json --quiet`
- **THEN** `target_month` 被解析为 `202602`
- **AND** CLI 执行完整销售报表工作流

#### Scenario: Missing required positional files
- **WHEN** 用户执行 `python cli.py --json`
- **THEN** argparse 报告缺少必需的位置参数
- **AND** 进程退出码为 2

### Requirement: Agent/Skill month acquisition for default full workflow

AI Agent integrations and skills MUST treat the full sales-report workflow as the default workflow. If the user provides two files without an explicit month, the Agent or Skill MUST try to infer `target_month` before running the CLI, and MUST ask the user when inference is not reliable.

#### Scenario: Infer month from filename or context
- **WHEN** 用户提供订单文件和支付流水文件
- **AND** 文件名或聊天上下文包含可识别的月份信息（如 `202603`、`2026年3月`、`3月份`、`上个月`）
- **THEN** Agent/Skill 将月份归一化为 `YYYYMM`
- **AND** 使用该 `target_month` 调用完整工作流

#### Scenario: Ask user when month is missing
- **WHEN** 用户提供订单文件和支付流水文件
- **AND** Agent/Skill 无法从文件名或上下文可靠识别月份
- **THEN** Agent/Skill MUST ask the user which month to process before invoking the CLI
- **AND** MUST NOT silently fall back to `--match-only` or basic matching

#### Scenario: Explicit reduced workflow request
- **WHEN** 用户明确表示只需要匹配手续费、不要销售报表或不要账期标注
- **THEN** Agent/Skill MAY run `--match-only`
- **AND** 该行为 MUST be treated as an explicit reduced workflow, not the default
