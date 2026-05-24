# pptchartengine Issues

Living log of open issues, known limitations, and historical breakages worth
remembering. Order: **Open** (P0/P1/P2) → **Known limits** → **Done**.

ADR-0005 §3 mandates this file stays current: when an issue ships, move it
to Done; when a workaround becomes permanent, promote it to Known limits.

---

## Open

### P1 — quality & completeness

- `PCE-007` **Round-trip safety net** — matrix harness covers ADR-0006 §3
  first-batch 7 chart_types (combo / scatter / bubble / plain_bar /
  plain_line / area / pie) × 3 contracts = 21 test items. Future expansion:
  semantic_family integration, more native chart variants.
  Status: **v1 done** (2026-05-24, all green).

### P2 — backlog

- `PCE-003` Split future chart families out of combo core.
  Goal: keep `combo` stable while adding `stacked`, `waterfall`, and `scatter`.
  Status: phase 1 started; `stacked` round-trip support added.

---

## Known limits (Won't fix in current iteration)

- `PCE-LIM-001` **`replace_pptx_chart_data` first batch: 7 chart types only.**
  Per ADR-0006 §3: `bar / line / combo / area / pie / scatter / bubble`.
  Out of scope this iteration: linked external workbook, pivot chart,
  SmartArt, plugin charts, charts without embedded workbook, custom overlay
  families needing plot area rebuild. All such inputs return
  `ReplaceResult(status="failed", error_code="unsupported_chart_type" | ...)`.

- `PCE-LIM-002` **pandas 3.0 + numpy 2.x not supported.**
  `pyproject.toml` pins `pandas>=2.2,<3` and `numpy>=1.26,<2`. Background:
  on 2026-05-24 first-install in a fresh conda env auto-resolved to
  pandas 3.0.3 + numpy 2.4.6, which broke 6 round-trip contract tests
  (`categories_col` read back as `nan`; numpy ABI mismatch with
  existing `python-pptx`). Upper-bound pin restores the green baseline
  (35→42→49 tests over the day). To relax: re-run full `pytest tests/`
  and verify all contract tests stay green before unpinning.

- `PCE-LIM-003` **Metadata persistence: layer 1 only (consolidated 2026-05-24).**
  ADR-0007 §3 sketches a layered strategy (workbook hidden sheet → semantic
  anchor → custom XML part → shape alt text). **Layer 1 is now consolidated**
  in `src/pptchartengine/metadata.py` with `ChartMetadataV1` + `write_chart_metadata`
  public (PCE-009 done). Layer 2 (semantic_anchor) is implemented separately
  in `semantic_anchor.py` for shape-composition families. Layer 3 (custom XML
  part) and layer 4 (shape alt text) require a separate PRD before
  implementation, including PowerPoint / Keynote / Google Slides
  compatibility verification and round-trip migration plan.

---

## Done

- `PCE-009` **Metadata consolidation: layer 1 single-source-of-truth.** (2026-05-24)
  Created `src/pptchartengine/metadata.py` with authoritative
  `METADATA_SHEET_NAME` + `METADATA_SCHEMA_VERSION` constants (previously
  duplicated in `api.py` and `parser.py` — silent-drift hazard). New public
  API: `ChartMetadataV1` dataclass + `write_chart_metadata` function;
  legacy `_write_embedded_metadata` preserved as backward-compat in
  `metadata.py`. `api.py` / `parser.py` now re-import from `metadata.py`;
  `scatter.py` / `semantic_family.py` unchanged thanks to re-export. ADR-0007
  §1 metadata lifecycle public API count rose from 0 to 4. 81 tests all green.

- `PCE-008` **Checked-in external .pptx fixture file.** (2026-05-24)
  New `tests/fixtures/external_sample.pptx` (54KB) + generator
  `_generate_external_sample.py`. Persistent multi-slide multi-chart fixture
  (bar / line / scatter / pie + named shape + non-chart text box) covering
  scenarios inline fixtures miss (cross-process artifact, byte-level drift
  detection, explicit_name extraction). 7 new tests in
  `tests/test_external_fixture_file.py` all green.

- `PCE-006` **ADR-0007 §1 public API audit completed.** (2026-05-24)
  All 123 public names in `__init__.py` classified across 6 buckets:
  create (26) / parse (23) / metadata (0 → now 4 after PCE-009) /
  inspect (3) / replace (3) / support (68). Doc:
  `docs/public_api_classification.md`. No "uncategorisable" name remains —
  ADR-0007 §1 工程约束 satisfied.

- `PCE-005` **`replace_pptx_chart_data` kernel implemented.** (2026-05-24)
  ADR-0006 §2 first-batch (bar/line/combo/area/pie/scatter/bubble). Result:
  in-place chart data replacement preserving shape identity, position, size,
  chart type, theme, and editability. 14 contract tests pass covering
  ADR-0006 §5 invariants 1, 3, 4, 5, 7, 8 (2 and 6 satisfied indirectly via
  python-pptx). New error codes added beyond ADR-0006 §3:
  `categories_required_for_category_chart`, `x_values_required_for_scatter`,
  `xy_length_mismatch`, `size_values_required_for_bubble`,
  `bubble_length_mismatch`.

- `PCE-004` **`inspect_pptx_charts` kernel implemented.** (2026-05-24)
  ADR-0006 §1 chart inventory: returns ordered list of `ChartInventoryItem`
  with stable `ChartSelector` (slide_index + shape_id + chart_part +
  optional explicit_name). 6 contract tests pass (empty / no-chart slides /
  single combo / multi-slide ordering / selector stability / read-only).

- `PCE-002` Add round-trip tests against generated `.pptx` files. (historic)
  Result: generated charts now have round-trip tests for semantic fields
  and stacked grouping.

- `PCE-001` Embed chart metadata into the workbook and use it during parsing. (historic)
  Result: parser recovers `categories_col` and original series keys more reliably.
