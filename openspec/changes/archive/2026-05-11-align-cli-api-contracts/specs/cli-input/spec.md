## 修改后的 Requirements

### Requirement: `cli.py` 位置参数模式

`cli.py` MUST 通过两个必需的位置文件参数和一个可选的位置参数 `target_month` 暴露当前命令行接口。标准工作流是完整销售报表工作流，因此调用方 SHOULD 提供或获取 `target_month`，除非用户明确请求如 `--match-only` 之类的缩减模式。

#### Scenario: 无 target_month 时进入月份获取
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx --json --quiet`
- **THEN** `order.xlsx` 被解析为订单文件
- **AND** `payment.xlsx` 被解析为支付流水文件
- **AND** `target_month` 为 `None`
- **AND** CLI 尝试通过交互式提示获取 `target_month`，而不是默认执行基础匹配

#### Scenario: 销售报表位置月份
- **WHEN** 用户执行 `python cli.py order.xlsx payment.xlsx 202602 --json --quiet`
- **THEN** `target_month` 被解析为 `202602`
- **AND** CLI 执行完整销售报表工作流

#### Scenario: 缺少必需的位置文件
- **WHEN** 用户执行 `python cli.py --json`
- **THEN** argparse 报告缺少必需的位置参数
- **AND** 进程退出码为 2

### Requirement: 默认完整工作流的 Agent/Skill 月份获取

AI Agent 集成和 skills MUST 将完整销售报表工作流视为默认工作流。如果用户提供两个文件但没有明确月份，Agent 或 Skill MUST 在运行 CLI 前尝试推断 `target_month`，并且在推断不可靠时 MUST 询问用户。

#### Scenario: 从文件名或上下文推断月份
- **WHEN** 用户提供订单文件和支付流水文件
- **AND** 文件名或聊天上下文包含可识别的月份信息（如 `202603`、`2026年3月`、`3月份`、`上个月`）
- **THEN** Agent/Skill 将月份归一化为 `YYYYMM`
- **AND** 使用该 `target_month` 调用完整工作流

#### Scenario: 月份缺失时询问用户
- **WHEN** 用户提供订单文件和支付流水文件
- **AND** Agent/Skill 无法从文件名或上下文可靠识别月份
- **THEN** Agent/Skill MUST ask the user which month to process before invoking the CLI
- **AND** MUST NOT silently fall back to `--match-only` or basic matching

#### Scenario: 显式缩减工作流请求
- **WHEN** 用户明确表示只需要匹配手续费、不要销售报表或不要账期标注
- **THEN** Agent/Skill MAY run `--match-only`
- **AND** 该行为 MUST be treated as an explicit reduced workflow, not the default
