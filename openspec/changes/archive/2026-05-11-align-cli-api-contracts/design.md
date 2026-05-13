## 背景

项目目前有三个面向用户的入口，共享同一套 Excel 合并与销售报表逻辑：

- `cli.py` for scripted/Agent usage.
- `excel_merge.py` for interactive local usage.
- `excel_merge_api.py` for HTTP upload/download usage.

核心行为集中在 `utils.py`，但各入口对输入、输出、错误与持久化的处理不同。最近的规格对齐已记录当前行为，其中一个重要产品规则是：默认的 Agent/Skill 工作流是完整销售报表工作流，因此需要通过用户明确回答或可靠的文件名/上下文推断获取 `target_month`。

本次变更只聚焦契约对齐。它应当先让对外行为变得明确，再进行后续 workflow/service 层重构。

## 目标 / 非目标

**目标：**

- 让默认的 Agent/Skill 自动化路径明确：获取 `target_month` 并运行完整工作流，除非用户明确要求仅匹配的缩减流程。
- 通过现有 `ok/data/error` 信封让 CLI JSON 对自动化保持可预测。
- 确定并文档化 API JSON 契约，使 `/merge/json` 不再与 CLI 文档和测试冲突。
- 让测试与面向 Agent 的文档与已选契约保持一致。
- 保留 CLI 的就地输出语义：CLI 将订单文件原地写回，不创建 `report_*.xlsx` 文件。
- 保留 API 下载语义：HTTP 调用方通过 `results/` 和 `/download/<filename>` 获取结果/报表文件。

**非目标：**

- 本次变更不引入 workflow/service 层。
- 本次变更不拆分 `utils.py`，也不对 `process_excel_files()` 做纯函数化。
- 本次变更不修改核心匹配算法或 P-number/hyphen 优先级。
- 不引入新的运行时依赖。
- 不移除现有 HTTP 端点。

## 决策

### 决策 1：将完整工作流视为默认自动化意图

对于 Agent/Skill 用法，提供两份文件意味着“运行完整的匹配与销售报表工作流”，除非用户明确表示只需要匹配。由于完整工作流需要月份，Agent/Skill 必须在调用 CLI 前获取 `target_month`。

理由：

- 这与飞书式使用场景的业务预期一致。
- 它避免在缺少月份时静默地少处理文件。
- 它让 `--match-only` 保持为有意的缩减工作流，而不是意外回退。

备选方案：

- 将两文件 CLI 调用默认设为仅匹配。否决，因为产品预期默认应完整处理。
- 即使文件名上下文已经明显，也始终询问月份。否决，因为可靠的文件名/上下文推断能保持流程高效。

### 决策 2：本次变更保持 CLI 位置参数调用

当前 `cli.py` 契约使用 `order_file payment_file [target_month]`。本次变更应围绕该行为对齐文档和测试，而不是立即引入 `--month`。

理由：

- 它将本次契约对齐变更的实现范围降到最低。
- Skill 已经成功使用位置参数 `target_month`。
- 后续 CLI 体验改进可以在需要时把 `--month` 作为兼容别名加入。

备选方案：

- 现在就加入 `--month` 并废弃位置参数 `target_month`。暂缓，因为本次变更已经覆盖 CLI、API、文档和测试。

### 决策 3：保留 CLI JSON 信封，并允许按工作流区分统计项

CLI JSON 应继续使用：

```json
{ "ok": true, "data": { ... }, "error": null }
```

完整工作流统计 MAY 在 `total_rows`、`matched_rows` 与 `match_rate` 之外包含 `marked_rows`。缩减模式 MAY 暴露与所选模式相符的统计项。

理由：

- 该信封对 Agent 友好，并且已在 CLI helper 中实现。
- `marked_rows` 对完整销售报表工作流的可观测性很有帮助。
- 强行让所有模式使用完全相同的统计对象，要么丢失有用数据，要么需要无意义字段。

备选方案：

- 所有模式都使用严格相同的统计字段。否决，因为 `--mark-only` 没有有意义的 match rate。

### 决策 4：目前保留 API 专属响应形状，但使其明确

就本次变更而言，`/merge/json` 应继续保持 API 专属形状，使用 `success`、`session_id`、`download_url`、`statistics` 和 `files`。规格和测试应明确说明，这与 CLI JSON 是有意不同的。

理由：

- HTTP 调用方需要下载 URL 和文件标识。
- 这可以避免破坏已经依赖 `success` 和 `download_url` 的现有客户端。
- 通过文档化差异，仍然可以消除歧义。

备选方案：

- 将 API JSON 改成 CLI 的 `ok/data/error` 信封。暂缓，因为那会是破坏性 API 变更。之后可通过版本化端点或兼容层引入。

### 决策 5：保持 CLI 与 API 持久化语义分离

CLI 和交互模式会将更新后的订单数据写回订单文件，不创建独立报表文件。API 模式 MAY 将可下载的合并/报表文件持久化到 `results/`，因为 HTTP 客户端需要可下载产物。

理由：

- CLI 用法按契约是文件本地且就地写回的。
- HTTP 用法以请求/响应为导向，需要服务器端结果路径。
- 核心销售报表工作流本身仍不写报表文件；API 持久化仍是适配层职责。

备选方案：

- 强制 API 镜像 CLI 的就地行为。否决，因为上传文件存放在服务器管理的临时路径中，调用方需要下载能力。

### 决策 6：尽可能修复 API 校验不对称

除非兼容性原因阻止，`/merge/json` 应与 `/merge` 使用相同的文件扩展名校验。错误响应应保持 API 形状。

理由：

- 两个端点接受相同文件类型，应对不支持文件尽早失败。
- 这能减少两个 HTTP 路由之间的分歧行为。

备选方案：

- 保留 `/merge/json` 更弱的校验。否决，因为这会保留一个本可避免的契约差距。

## Risks / Trade-offs

- [Risk] Existing docs or tests may still assume `--month` instead of positional `target_month`. → Mitigation: update Agent docs, usage docs, and tests in the same change; leave a future change for adding `--month` alias.
- [Risk] API and CLI continue to use different JSON envelopes. → Mitigation: document the difference explicitly and keep API-specific response fields stable; consider a future versioned API envelope if needed.
- [Risk] Full workflow default may block Agent execution when month cannot be inferred. → Mitigation: Skill must ask the user instead of falling back to matching-only; this is intentional to avoid incomplete processing.
- [Risk] Updating tests may expose current implementation quirks such as `cli.py` interactive EOF cancellation. → Mitigation: cover current behavior where it remains part of contract, and isolate future behavior changes into separate changes.
- [Risk] API report files and CLI no-report-file semantics can look inconsistent. → Mitigation: specify that API persistence is adapter-level behavior while the core workflow remains in-memory.

## Migration Plan

1. Update delta specs for `cli-input`, `cli-output`, `http-api`, `agent-documentation`, `automated-testing`, and `sales-report`.
2. Update Agent-facing docs and Skill docs to describe full-workflow default and month acquisition.
3. Align tests with the selected CLI/API contracts.
4. Apply implementation changes only where required by the aligned contract, especially API file validation and documentation/test expectations.
5. Run `openspec validate --all --strict` and the pytest suite.

Rollback strategy: revert this change’s spec/doc/test/implementation updates together. Since no database migration or new dependency is expected, rollback is a git revert.

## Open Questions

- Should a future change add `--month YYYYMM` as an alias while keeping positional `target_month` for compatibility?
- Should a future API version expose the CLI-style `ok/data/error` envelope while preserving the current `/merge/json` shape?
- Should `cli.py` continue interactive prompting in JSON mode, or should Agent/Skill month acquisition fully replace that path in a later CLI cleanup?
