# ADR-0005: AI 协作工程纪律

- Status: Accepted
- Date: 2026-05-01
- Related: ADR-0002, ADR-0004

## Context

本仓库的代码主要通过 AI 协作（Claude Code）开发。LLM 协作有几个公认的物理约束和失效模式：

- **智能区退化**：上下文超过约 100k token 后注意力分布退化，AI 推理质量明显下滑
- **遗忘性**：每次新会话从零开始；"压缩历史"会沉积偏见与漂移
- **静默误对齐**：缺乏对齐机制时 AI 容易产出"看起来对但方向错"的代码
- **文档腐烂污染**：过期 PRD、handover、TODO 让 AI 基于过时信息推理
- **同 context 自审失效**：写过代码的 AI 来 review 自己的代码会自我合理化

ADR-0002 定义了"什么算合格输出"，ADR-0004 定义了"new family 的设计约束"。但都没规定"AI 协作下应该怎么写"。本 ADR 把若干已被实战验证有效的 AI 协作纪律固化为团队约束。

## Decision

### 1. TDD 作为新 family / 新 public API 的默认流程

工作顺序必须是：

1. 先写 round-trip 契约测试（依据 ADR-0004 闭环要求）
2. 跑测试，确认它失败
3. 实现 parser / generator，让测试由红转绿
4. 重构

这与 ADR-0004 的 round-trip 约束天然合拍：测试就是闭环 spec，AI 不能在实现里"作弊"绕开闭环。

bug 修复同样适用：先写一个能复现 bug 的测试，再修。

### 2. 重大变更前写 PRD

**触发条件**：

- 新 family
- breaking API change
- 新 preset 系列
- metadata schema 演进
- 跨多个模块的重构

不触发：bug 修复、依赖升级、文档调整、单 family 内部小改。

**PRD 内容**（放在 `docs/prds/<id>-<slug>.md`）：

- 问题陈述
- 解决方案概要
- 影响的 ADR 与现有 family
- 测试策略（必须能映射到 L1/L2/L3）
- 目标 golden reference 或验收例

**生命周期**：`in-progress → completed → archived (docs/prds/archive/) 或删除`。

### 3. 反对文档腐烂

- 完成的 PRD 立即移到 `docs/prds/archive/` 或删除
- handover 类文档（如 `2026-04-19-range-snapshot-handover.md`）必须在 14 天内：合并入 ADR、合并入 PRD、或归档删除
- TODO / FIXME 由季度 audit 清理：要么做、要么开 issue、要么删
- 不写 "removed for now" / "kept for compatibility" 这类注释，直接删代码
- README、ROADMAP、ADR 与代码不一致即视为 bug，必须立即修复

### 4. Code review 必须新开 context

- 写 family 的对话不能同时 review 它
- PR 提交后开新会话做 review
- Review 上下文只包含：PR diff + 相关 ADR + 相关 golden 断言 + 直接相关的现有测试
- 推荐 review 用更高 capability 模型（如 Opus），写代码可用更快模型

### 5. 智能区切片：限制单任务的工作集

不要把全量引擎喂给单个 AI 任务。新建任务时按 family 切片：

| 任务类型 | 推荐工作集 |
|---|---|
| 改一个 family | 该 family 文件 + 其测试 + 相关 oxml/* + 涉及的 ADR |
| 加新 family | 一个模板 family + parser.py + 新 PRD + ADR-0004 |
| 改 parser | parser.py + 当前 family 的 metadata 写入端 + 测试 |
| 改 visual diff | L3 框架 + 1 份 golden + diff 工具 |

跨 family 改动拆成独立子任务串行执行，不要在单 context 里做。当 AI 频繁需要"再看一下另一个文件"时，是切片粒度太粗的信号。

## Consequences

正面：

- TDD + ADR-0004 双保险，新 family 闭环不破
- PRD 强制对齐，减少静默误对齐
- 文档腐烂规则让 AI 始终读到当前 truth
- review 隔离能发现 self-blind spot
- 智能区切片让 AI 推理质量稳定

负面：

- PRD 写作有成本（用触发条件严格限定）
- 工程纪律需要团队共识，违反时 review 必须阻拦
- 切片粒度需要持续校准

## Alternatives Considered

**完全自由**：靠 PR review 兜底，不立规矩。
否决：实战中已观察到 AI 在大 context、无 PRD 下产出"方向性错误"代码，到 review 阶段才发现修复成本高。

**全自动 Ralph 循环 + Docker sandbox**（Matt Pocock 完整框架）：
否决：本项目是 library + 视觉验收，自动化收益有限；引入 sandbox 基础设施成本高。保留为 future option，等 v1.0 后视情况引入。

**每次代码变更都写 PRD**：
否决：bug 修复、依赖升级、文档调整不需要 PRD。触发条件被严格限定，避免形式主义。
