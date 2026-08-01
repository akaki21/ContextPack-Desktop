# AI handoff guide

Use generated Markdown for the first reading pass and originals only for verification or editing.

## PDF prompt

```text
Read the Markdown first and build a content map. Use the source PDF and selected page images only to verify important or uncertain details. Flag possible OCR/conversion errors. Goal: [goal]. Scope: [pages/topics]. Desired output: [format].
```

## Excel prompt

```text
Use workbook-info.md to identify relevant sheets and ranges. Check formulas.md for calculation logic and values.md for results. Verify reported issues in the original workbook. Goal: [goal]. Scope: [sheets/ranges/period]. Desired output: [format].
```

For large files, naming exact pages, sheets, or ranges saves the most time and tokens.
