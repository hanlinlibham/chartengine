# Changelog

All notable changes to `ablechart` will be documented in this file.

## 0.2.0 - Unreleased

Added (issue #9, Gap A — report-grade "last mile", first slice):

- **Axis titles**: `axis_title` field on `CategoryAxisConfig` and
  `ValueAxisConfig`, written as native `c:catAx`/`c:valAx > c:title` on the
  category axis and both value axes (primary + secondary). Round-trips through
  save/reopen; off by default.
- Document the remaining styling-granularity gaps (tick-label rotation, log
  scale, display units, legend-entry control) in README `## Current Limits`.
- Rename the debug env var to `ABLECHART_DEBUG_STDOUT` (old
  `PPTCHARTENGINE_DEBUG_STDOUT` kept as an alias).

Fixed (issue #9, Gap B — robustness):

- **Value-axis resolution by `axId`** instead of `axPos`. `oxml.axes.resolve_value_axes`
  identifies the primary/secondary value axis by the axId the plot group
  references (document order), so `number_format` / scale land on the right
  axis even on horizontal bar combos where the value axes sit at `axPos='b'`/`'t'`
  (the old `axPos in ('l','b')` / `=='r'` lookup could miss the secondary
  entirely).

Still open from issue #9: tick-label rotation, legend-entry control, and the
LibreOffice category-title double-render in secondary-axis combos (needs the
4-axis combo restructure; Gap B's axId fix does **not** address it).

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
