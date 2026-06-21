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
performance_attribution → v0.1.0 →
factor_style_analysis  → v0.2.0 →
valuation_snapshot     → v0.3.0 →
剩余 family 补齐        → v1.0.0
```

每个 slice 都端到端：选定一份 golden → 它涉及的 family 各自补 L2 + L3 visual diff + skill cookbook + dogfood → 闸门全绿 → 发版（仅声明支持当前 slice 范围，参见 [ADR-0002](docs/adr/0002-vertical-first-quality-gates.md)）。

## 衡量指标

- **首版交付提前**：v0.1.0 在 4-6 周内发布（vs 水平模式 12-14 周）
- **L3 通过率**：tag 发版时声明支持的 golden 100% 绿
- **视觉 diff**：每页与 baseline 像素差 < 阈值（初期 5%，后续按字体差异收紧）
- **Skill dogfood 命中率**：Claude 用 skill 一次性产出可执行且语义正确代码 ≥ 80%
- **PPT 编辑性**：reference.pptx 在 PowerPoint 打开后图表可双击编辑、无报错对话框

## 2026-06-01 规划复盘

当前项目已经更清楚地收敛成 **PowerPoint chart asset kernel**。这带来一个路线图修正：

- PyPI 首发不应被完整报告级 golden、L3 visual diff、skill dogfood 阻塞。
- 首发应先证明这个包作为独立 SDK 是可信的：安装、metadata、MIT license、README 定位、核心 create/parse/inspect/replace 契约、wheel/sdist 安装烟测。
- golden / skill / dogfood 仍然重要，但它们证明的是“上层报告生成工作流”的质量，不应和内核包能否发布混在同一个门槛里。

因此后续拆成两条线：

1. **Package release track**：把 `pptchartengine` 作为 MIT Python 包发布出去，声明 alpha 状态和明确支持范围。
2. **Report quality track**：继续用 `pptfi` / skill / golden reports 验证端到端金融报告效果。

近期优先级：

1. `P0` 发布前最后闸门：TestPyPI → 正式 PyPI → 安装烟测 → 国内镜像同步验证。
2. `P0` GitHub CI：跑 `pytest`、`build`、`twine check`，并保留 release-readiness 测试。
3. `P1` API 收口：复核 `__all__` 中 support 类导出，避免 0.x 早期暴露过多半内部符号。
4. `P1` 核心 family 文档：为 combo / waterfall / scatter / bubble / range_snapshot 各补最小 cookbook。
5. `P2` L3 visual diff：继续推进，但作为“声明 golden 支持”的门槛，而不是 PyPI 首发门槛。

## Milestones

### M1 — 引擎清理 & Slice 准备（约 2 周，至 ~2026-05-15）

**目标**：把当前未提交工作落地，建立 slice 1 起跑线。

- [x] 把 untracked 的 3 个文件 commit（工程师 `7bbe1c7` 完成）
- [x] 整理已 modified 文件的 commit（同上）
- [x] LICENSE 文件加入（MIT）
- [x] 选定 3 份 golden 并按分析类型命名、按优先级排序（[ADR-0003](docs/adr/0003-golden-reference-reports.md)）
- [x] 建立 `docs/prds/` 目录骨架（含 [PRD-0001](docs/prds/0001-range-snapshot-visual-polish.md) range_snapshot 视觉精修）
- [x] `2026-04-19-range-snapshot-handover.md` 归档到 `docs/prds/archive/`，关键判断转入 PRD-0001
- [ ] 在 pptfi 仓库建立 `goldens/` 目录骨架
- [ ] 为 slice 1 准备 `performance_attribution` 脱敏输入数据

退出条件：`git status` clean、3 份 golden 选定（done）、slice 1 输入数据已脱敏入库。

### Slice 1 — performance_attribution（+2 ~ +6 周，目标 v0.1.0）

**Golden**：业绩归因报告（基金 / 组合收益来源拆解）

**涉及 family**：`performance_compare` (combo) + `attribution_decomposition` (waterfall) + `regime_table_panel`

**为什么作为首版**：所有涉及的 family 都是纯 native chart，视觉质量结构性最强，是稳健的 v0.1.0。

L2 层：

- [ ] 这些 family 各补 1 happy path + 1-2 edge case 测试
- [ ] L2 通用 assertion helper 落地（颜色 / 字号 / legend / 轴范围）

L3 层：

- [ ] PPT → PNG 渲染 pipeline（`soffice --headless`）
- [ ] 视觉 diff 工具集成（pixelmatch 或 imagehash + 阈值）
- [ ] golden #1 端到端 fixture：输入 → 引擎 → pptx → png → diff
- [ ] baseline 更新流程文档化

Skill + Dogfood：

- [ ] 写 `ppt-finance-chart-skill` 骨架
- [ ] 这些 family 的 cookbook 入口
- [ ] dogfood 测试（Claude 用 skill → 产出 → diff golden #1）

发版准备：

- [x] `pyproject.toml` metadata 补全（authors / license / urls / classifiers）
- [x] CHANGELOG.md 启动，明确声明本版本支持范围
- [x] README 重写为开源入口，并补充相似项目差异说明
- [x] release-readiness tests 覆盖 license / pyproject / README / MANIFEST
- [x] GitHub Actions CI / Trusted Publishing workflow 文件
- [x] GitHub `testpypi` / `pypi` environments 创建
- [ ] TestPyPI / PyPI 侧 pending trusted publisher 配置
- [ ] CI 闸门接好（pytest + build + twine check + release-readiness tests）
- [ ] TestPyPI 验证一轮
- [ ] v0.1.0 tag 发版到正式 PyPI（README 明确声明仅支持 slice 1 涉及的 family）
- [ ] 国内镜像（清华 / 阿里）同步验证
- [ ] golden #1 的 L3 + skill dogfood 作为发版后 quality milestone 接入

退出条件：v0.1.0 在 PyPI / 国内镜像可装、README / CHANGELOG 准确反映 alpha 支持范围、核心契约测试与发布构建在 CI 绿。

### Slice 2 — factor_style_analysis（+6 ~ +9 周，目标 v0.2.0）

**Golden**：因子风格分析报告（组合在风格 / 因子上的暴露）

**涉及 family**（去重后新增）：`style_box` (scatter) + `factor_exposure` (combo) + `style_allocation` + `score_overlay`

**为什么 slice 2**：引入 scatter family 端到端验证；combo 已经在 slice 1 跑通，只需扩展场景。

- [ ] 上述 family 各补 L2
- [ ] golden #2 端到端 fixture 接入
- [ ] skill cookbook 增补对应场景
- [ ] dogfood 扩展
- [ ] v0.2.0 发版，CHANGELOG 与 README 更新声明范围
- [ ] 复盘 slice 1 反馈，回写 ADR / 修订工程纪律

退出条件：v0.2.0 发版、L2 + L3 + dogfood 在新增范围内全绿。

### Slice 3 — valuation_snapshot（+9 ~ +12 周，目标 v0.3.0）

**Golden**：估值快照报告（市场 / 行业当前估值在历史区间位置）

**涉及 family**：`range_snapshot` + `score_overlay`（复用）+ `heatmap_matrix`

**前置依赖**：[PRD-0001](docs/prds/0001-range-snapshot-visual-polish.md) 视觉精修必须在本 slice 内完成。range_snapshot 当前是 native chart + custom overlay 复合结构，视觉质量低于 combo，需调参精修。

- [ ] PRD-0001 range_snapshot 视觉精修：plot area / bar gap / tick label / current label / axis break 调参
- [ ] heatmap_matrix family L2
- [ ] golden #3 端到端 fixture
- [ ] skill cookbook 增补
- [ ] dogfood 扩展
- [ ] v0.3.0 发版

退出条件：3 份 golden 在 CI nightly 全部稳定通过 ≥ 2 周。

### M-final — 剩余 family 补齐（+12 ~ +16 周，目标 v1.0.0）

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
| range_snapshot 视觉精修未完成阻塞 v0.3.0 | PRD-0001 显式跟踪；如必要可推迟 v0.3.0 |
| 19 个 family 测试体量爆炸 | 每个 slice 只覆盖该 slice 需要的 family；剩余在 M-final 用 parametric fixture 复用断言 |
| 真实金融数据合规 | 脱敏后版本入 golden，按分析类型命名隐去具体主题 |
