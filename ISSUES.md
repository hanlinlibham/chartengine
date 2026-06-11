# pptchartengine Issues

## Open

- `PCE-003` Split future chart families out of combo core.
  Goal: keep `combo` stable while adding `stacked`, `waterfall`, and `scatter`.
  Status: phase 1 started; `stacked` round-trip support added.

## Done

- `PCE-009` Structured-output alignment: `chart_spec_schema()` JSON Schema
  (draft-07) for constrained decoding; enums lock chart/axis/legend/grouping
  at the decoding layer while additionalProperties stays open for the
  tolerant alias layer. Pairs with validate_spec() for semantics.

- `PCE-008` Professional polish round vs market-guide benchmark deck.
  Result: title+subtitle system (bold near-black title + gray unit line, all chart
  kinds, left-aligned paragraphs); range chart refined (40%-slot bars, average dash
  dynamically wider than bar, white-bordered current diamond, 0"x" axis format,
  sort-by-current, subtitle); annotation out-of-range clipping (correctness bug);
  vband (recession shading); dark zero-baseline when data crosses zero; nice-unit
  ladder gains 1.5/3 steps (GTM 3%-step axes, tighter headroom); `sort` for single-
  series rankings; "times"/"倍" number-format alias. Gap assessment:
  workspace/260611/JPM差距评估.md (~85% GTM chart-pattern coverage; remaining gaps
  are mini-tables/page chrome → pptfi, arrow legends + trendlines → engine backlog).

- `PCE-007` Capability-lowering pack for weak LLM callers.
  Result: `chart_spec_examples()` few-shot gallery (12 scenarios, keyword lookup);
  annotation auto-math (`average` computes mean from a series, `band` accepts
  quantiles — the engine does arithmetic, not the model); auto dual-axis in
  inference mode when two columns differ >50x in magnitude. Assessment doc:
  workspace/260611/模型能力评估.md.

- `PCE-006` GTM pattern library (general market-guide deck patterns).
  Result: contribution preset (stacked parts + orange total line, palette skips
  orange), new `create_range_chart` family (range bar + average dash + current
  diamond, embedded metadata), horizontal bar orientation, native value labels
  (dLbls), category highlight (dPt), forecast hatching (pattFill) + divider,
  annotation layer (average/reference hline with inline label, target band,
  vline, last-point dot + date/value callout) on pinned plot area. New `gtm`
  color scheme. All spec-addressable; 15 new tests.

- `PCE-005` Report-grade visual polish pass.
  Result: new `polish.py` runs after every build — smart nice-range axis scaling
  (4/5/6-tick joint optimization, dual axes share tick count so gridlines align),
  gapWidth 80/50, sans title 13pt bold left-aligned, gray 9pt axis text, value-axis
  lines removed. Waterfall rebuilt on pinned plot area (manualLayout) + explicit
  axis range so slide-level value labels / connectors align exactly (connectors now
  default ON, y-axis hidden, muted sage/brick/navy palette, thousands separators,
  collision-aware label placement). Scatter/bubble get nice-range zoom, bottom
  tick labels, bubble alpha 72%. Fixed strftime cascade bug (`yyyy/mm` → literal
  `%mm` on axis) and daily-data duplicate month labels (granularity-aware format).
  Schemes: new `advisory`; `default` reordered dark-first. Opt out via `polish=False`
  / spec `"polish": false`.

- `PCE-004` Declarative spec layer for low-capability LLM callers.
  Result: `render_chart()/validate_spec()` accept a single JSON-friendly dict with
  alias normalization (CN/EN), smart inference (categories/series/date-axis),
  collected did-you-mean errors, custom palettes, and per-series style overrides.
  Reference doc in `SPEC.md`; `chart_spec_reference()` returns a prompt-ready cheat sheet.

- `PCE-001` Embed chart metadata into the workbook and use it during parsing.
  Result: parser now recovers `categories_col` and original series keys more reliably.

- `PCE-002` Add round-trip tests against generated `.pptx` files.
  Result: generated charts now have round-trip tests for semantic fields and stacked grouping.
