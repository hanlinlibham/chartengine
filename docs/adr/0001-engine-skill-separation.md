# ADR-0001: pptchartengine 与 skill 的分离

- Status: Accepted
- Date: 2026-05-01

## Context

pptchartengine 当前状态：

- 9,642 行 Python 代码，18 个模块，19 个 semantic family
- 已是独立 git 仓库（`hanlinlibham/pptchartengine`）
- 与 `pptfi`（报告级 orchestrator）的边界清晰：引擎管图表内核与图表级语义层，pptfi 管 connectors / job spec / 报告编排
- 公开 API 表面 60+ 符号，已经事实上是一个库

目标用户场景是金融分析师生成专业 PPT 报告，存在两类使用入口：

1. Python 工程师 / `pptfi` 等上层项目：直接 `import pptchartengine` 调用
2. Claude / AI agent：通过 skill 工作流读取代码模板、生成调用代码并执行

把 9.6k 行直接塞进 skill 仓库会让 skill 严重臃肿，且让 `pptfi` 失去可复用的引擎依赖。

## Decision

采用 **独立 pip 包 + 薄 skill** 架构：

- `pptchartengine` 作为独立 pip 包发布到 PyPI
- 单独写一个 finance-chart skill，仅承担：
  - 工作流文档（金融报告高频场景的代码模板）
  - 输入数据 schema 说明
  - 调用 cookbook 与排错指南
- `pptfi` 通过 pip 依赖该包，不再 vendor 复制
- skill 依赖该 pip 包并 pin 版本，skill repo 不持有引擎源码

## Consequences

正面：

- 引擎 API 在外部使用压力下被持续打磨
- skill 体量小、易导航，Claude context 友好
- `pptfi` 与未来其他金融工具可复用引擎，不重复发明
- 引擎可独立 versioning、CHANGELOG、breaking-change policy

负面：

- 双 release cadence 需要协调（skill 与 pip 包）
- 破坏性改动需要 deprecation cycle
- 初始发布、CI、Trusted Publishing 配置投入
- skill 需在用户没装 pip 包时给出清晰报错和安装提示

## Alternatives Considered

**A：纯 skill 内置**——所有引擎代码塞进 skill 仓库。
否决：9.6k 行远超合理 skill 规模；`pptfi` 无法复用；外部使用压力消失。

**B：vendor 复制**——pptfi 和 skill 各自 vendor 一份引擎。
否决：双向同步成本高；分支漂移风险；违反单一信息源原则。

**C：monorepo（engine + skill + pptfi 同仓）**——
否决：引擎已经有独立 git 仓库且 remote 公开；金融场景特定的 pptfi 不适合与通用引擎绑定 release 节奏。
