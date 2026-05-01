# pptchartengine Roadmap

## North Star

> 让分析师在 30 分钟内，借助 Claude + 本引擎产出一份达到卖方研究院水准的金融 PPT 报告。

可被验证的"达到水准"= [ADR-0003](docs/adr/0003-golden-reference-reports.md) 的 golden reference reports 全部 visual diff 在阈值内。

## 交付哲学

采用**垂直切片**而非水平分层堆叠（参见 [ADR-0005](docs/adr/0005-ai-collaboration-discipline.md)）。

不做：

```
全量 L2 → 全量 L3 → 全量 skill → 一次性 v1.0.0
```

要做：

```
golden #1 端到端打通 → v0.1.0 →
golden #2 端到端打通 → v0.2.0 →
golden #3 端到端打通 → v0.3.0 →
剩余 family 补齐 → v1.0.0
```

每个 slice 都端到端：选定一份 golden → 它涉及的 family 各自补 L2 + L3 visual diff + skill cookbook + dogfood → 闸门全绿 → 发版（仅声明支持当前 slice 范围，参见 [ADR-0002](docs/adr/0002-vertical-first-quality-gates.md)）。

## 衡量指标

- **首版交付提前**：v0.1.0 在 4-6 周内发布（vs 水平模式 12-14 周）
- **L3 通过率**：tag 发版时声明支持的 golden 100% 绿
- **视觉 diff**：每页与 baseline 像素差 < 阈值（初期 5%，后续按字体差异收紧）
- **Skill dogfood 命中率**：Claude 用 skill 一次性产出可执行且语义正确代码 ≥ 80%
- **PPT 编辑性**：reference.pptx 在 PowerPoint 打开后图表可双击编辑、无报错对话框

## Milestones

### M1 — 引擎清理 & Slice 准备（约 2 周）

**目标**：把当前未提交工作落地，建立 slice 1 起跑线。

- [ ] 把 untracked 的 3 个文件（`plot_area.py`、`range_snapshot.py`、`semantic_family.py`）commit
- [ ] 整理已 modified 文件的 commit 划分
- [ ] LICENSE 文件加入（建议 Apache-2.0）
- [ ] 选定 3 份 golden reference reports 并按优先级排序
- [ ] 为 slice 1 选定的 golden 准备脱敏输入数据
- [ ] 建立 `docs/prds/`、`goldens/` 目录骨架
- [ ] `2026-04-19-range-snapshot-handover.md` 等历史 handover 文档归档或合并入 ADR（参见 ADR-0005 反文档腐烂）

退出条件：`git status` clean、3 份 golden 选定、slice 1 的 golden 输入数据已脱敏入库。

### Slice 1 — 第一份金标准端到端打通（+2 ~ +6 周，目标 v0.1.0）

**目标**：选定的 golden #1 通过完整闸门，发布首个 pip 版本。

假设 golden #1 涉及 ~3-4 个 family（待 M1 选定后填入）。本 slice 内对**这些 family** 完成：

L2 层：

- [ ] 这些 family 各补 1 happy path + 1-2 edge case 测试
- [ ] L2 通用 assertion helper 落地（颜色 / 字号 / legend / 轴范围）

L3 层：

- [ ] PPT → PNG 渲染 pipeline（`soffice --headless`）
- [ ] 视觉 diff 工具集成（pixelmatch 或 imagehash + 阈值）
- [ ] golden #1 的端到端 fixture：输入 → 引擎 → pptx → png → diff
- [ ] baseline 更新流程文档化

Skill + Dogfood：

- [ ] 写 `ppt-finance-chart-skill` 骨架
- [ ] 这些 family 的 cookbook 入口
- [ ] dogfood 测试（Claude 用 skill → 产出 → diff golden #1）

发版准备：

- [ ] `pyproject.toml` metadata 补全（authors / license / urls / classifiers）
- [ ] CHANGELOG.md 启动，明确声明本版本支持范围
- [ ] GitHub Actions Trusted Publishing 配置
- [ ] CI 闸门接好（L1 + 涉及 family 的 L2 + golden #1 的 L3 + skill dogfood）
- [ ] TestPyPI 验证一轮
- [ ] v0.1.0 tag 发版到正式 PyPI（README 明确声明仅支持 slice 1 涉及的 family）
- [ ] 国内镜像（清华 / 阿里）同步验证

退出条件：v0.1.0 在国内镜像可装、golden #1 在 CI nightly 自动比对绿、README 准确反映支持范围。

### Slice 2 — 第二份金标准（+6 ~ +9 周，目标 v0.2.0）

**目标**：复用 slice 1 的脚手架，扩展支持 family。

- [ ] golden #2 涉及的新 family（去重后）补 L2
- [ ] golden #2 端到端 fixture 接入
- [ ] skill cookbook 增补对应场景
- [ ] dogfood 扩展
- [ ] v0.2.0 发版，CHANGELOG 与 README 更新声明范围
- [ ] 复盘 slice 1 的反馈，回写 ADR / 修订工程纪律

退出条件：v0.2.0 发版、L2 + L3 + dogfood 在新增范围内全绿。

### Slice 3 — 第三份金标准（+9 ~ +11 周，目标 v0.3.0）

**目标**：覆盖第三类金融报告场景。

- [ ] golden #3 涉及 family L2
- [ ] golden #3 端到端 fixture
- [ ] skill cookbook 增补
- [ ] dogfood 扩展
- [ ] v0.3.0 发版

退出条件：3 份 golden 在 CI nightly 全部稳定通过 ≥ 2 周。

### M-final — 剩余 family 补齐（+11 ~ +16 周，目标 v1.0.0）

**目标**：剩余 semantic family 按 [ADR-0004](docs/adr/0004-round-trip-metadata-principle.md) round-trip 约束逐个补齐。

- [ ] 没被 3 份 golden 覆盖的 family 各补 happy path L2
- [ ] PowerPoint 真实打开半自动验收（macOS osascript，季度跑）
- [ ] API 表面 review，可疑 export 收窄到 internal `_*`
- [ ] 文档完整性 review（README、ADR、ROADMAP 一致性）
- [ ] v1.0.0 发版（API stable 承诺）

退出条件：所有 public family 都满足 ADR-0002 闸门、API 稳定承诺成立。

### 维护期（v1.0.0 后）

- 用户反馈循环（issue triage、API 易用性反馈）
- 季度 baseline review
- 新 family 按真实需求增补，遵守 ADR-0004 round-trip 约束 + ADR-0005 工程纪律
- 视情况开放 conda-forge

## 不在路线图内

- candlestick / OHLC（按 README 当前安全范围排除）
- sankey / mekko / tornado
- 把引擎扩展为通用图表库（违反 [ADR-0001](docs/adr/0001-engine-skill-separation.md) 垂直定位）
- 把报告编排能力下沉到引擎（违反 [ADR-0001](docs/adr/0001-engine-skill-separation.md) 与 pptfi 的边界）

## 主要风险

| 风险 | 缓解 |
|-----|------|
| Golden baseline 主观性 | 季度 PM 评审 + 多人交叉确认 |
| 字体跨平台差异导致视觉 diff 假阳性 | 阈值容忍 + 锁定 CI 渲染环境（Docker 镜像） |
| pip 与 skill 双 release 漂移 | skill pin 引擎版本，CI 检查兼容性矩阵 |
| 早期 slice 暴露架构假设错误 | 这正是垂直切片的目的——v0.1.0 时纠正比 v1.0.0 时纠正便宜 |
| 19 个 family 测试体量爆炸 | 每个 slice 只覆盖该 slice 需要的 family；剩余在 M-final 用 parametric fixture 复用断言 |
| 真实金融数据合规 | 脱敏后版本入 golden，保留行业代码隐去公司名 |
