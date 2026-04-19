# pptchartengine Issues

## Open

- `PCE-003` Split future chart families out of combo core.
  Goal: keep `combo` stable while adding `stacked`, `waterfall`, and `scatter`.
  Status: phase 1 started; `stacked` round-trip support added.

## Done

- `PCE-001` Embed chart metadata into the workbook and use it during parsing.
  Result: parser now recovers `categories_col` and original series keys more reliably.

- `PCE-002` Add round-trip tests against generated `.pptx` files.
  Result: generated charts now have round-trip tests for semantic fields and stacked grouping.
