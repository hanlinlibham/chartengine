# Changelog

All notable changes to `ablechart` will be documented in this file.

## 0.2.0 - Unreleased

Added (issue #9, Gap A — report-grade "last mile", first slice):

- **Axis titles**: `axis_title` field on `CategoryAxisConfig` and
  `ValueAxisConfig`, written as native `c:catAx`/`c:valAx > c:title` on the
  category axis and both value axes (primary + secondary). Round-trips through
  save/reopen; off by default.
- **Legend-entry control**: hide a single series from the legend via
  `show_in_legend=False` on a `series_config` entry (or `legend: false` on a
  spec series), emitting a `c:legendEntry delete=1` while the series stays on
  the plot. (Manual XY legend placement is still not exposed.)
- **Tick-label rotation**: `tick_label_rotation` (degrees) field on
  `CategoryAxisConfig` and `ValueAxisConfig`, written as `a:bodyPr@rot`. Also
  fixes `polish.set_axis_text` silently dropping the rotation when it rebuilds
  the axis text element (an instance of the polish-vs-config clobbering noted in
  #9) — rotation is now preserved across the polish pass. Off by default.
- Document the remaining styling-granularity gaps (tick-label rotation, log
  scale, display units, legend-entry control) in README `## Current Limits`.
- Rename the debug env var to `ABLECHART_DEBUG_STDOUT` (old
  `PPTCHARTENGINE_DEBUG_STDOUT` kept as an alias).

Fixed (issue #14 — range chart legend wired into spec):

- The spec/`render_chart` range branch now passes `range_name` /
  `average_name` / `current_name` and supports `legend: "none"` (also
  off/hide/无) to hide the legend. Non-valuation range charts (allocation
  bands, spread ranges) can rename or drop the legend from the job layer
  instead of post-editing chart XML. `create_range_chart` gains a `show_legend`
  flag; default legend name behavior unchanged.

Fixed (issue #13 — number-format correctness):

- `_normalize_number_format` no longer passes an unrecognized token straight
  through as an Excel `formatCode`. A bare word like `"number"` used to land in
  `numFmt` and render as garbage on the axis; it is now rejected with a
  did-you-mean (via the spec error list, or a warning) and falls back to no
  explicit format. Real formatCodes (`0.0%`, `#,##0`, `yyyy-mm-dd`, `"¥"0`) and
  aliases still work; added `general`/`auto` aliases. New public
  `supported_number_format_tokens()` lists the accepted aliases.

Fixed (issue #9, Gap B — robustness):

- **Value-axis resolution by `axId`** instead of `axPos`. `oxml.axes.resolve_value_axes`
  identifies the primary/secondary value axis by the axId the plot group
  references (document order), so `number_format` / scale land on the right
  axis even on horizontal bar combos where the value axes sit at `axPos='b'`/`'t'`
  (the old `axPos in ('l','b')` / `=='r'` lookup could miss the secondary
  entirely).

Still open from issue #9: legend-entry control, log scale, native display
units, and the LibreOffice category-title double-render in secondary-axis
combos (needs the 4-axis combo restructure).

## 0.1.1 - 2026-06-21

Security hardening (no exploitable issue in 0.1.0; defense-in-depth):

- Relax `lxml` pin to `>=5,<7` so installs can resolve to the patched 6.1.0+
  (PYSEC-2026-87 ships a safe `resolve_entities` default). ablechart never calls
  lxml parsers directly; XML reads go through python-pptx, which already rejects
  DOCTYPE / external entities. Verified against lxml 6.1.0.
- Pin all GitHub Actions to full commit SHAs (with version comments) in the CI
  and publish workflows, removing mutable-tag supply-chain risk on the
  OIDC-publishing pipeline.

## 0.1.0 - 2026-06-21

Initial public package candidate.

Supported scope:

- editable native PowerPoint combo charts
- waterfall, scatter, bubble, and range snapshot chart families
- semantic metadata round-trip for generated charts
- chart inventory and template-safe data replacement for the first-batch native chart types
- experimental semantic component families for single-slide financial report panels

Packaging:

- MIT license
- Python `>=3.10`
- PyPI metadata and source distribution manifest prepared
- Release-readiness tests for license, package metadata, README positioning, and sdist manifest
- GitHub Actions CI plus TestPyPI / PyPI Trusted Publishing workflow prepared
