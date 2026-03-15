# StreamScheme

Fast, typed, streaming read and write of tabular data in xlsx format. Nothing else.

## Is StreamScheme for you?

| What you need | StreamScheme |
|---|---|
| Read tabular xlsx data, row by row | Yes |
| Write tabular xlsx data, row by row | Yes |
| Typed cells (text, numbers, dates, booleans) | Yes |
| Low memory, streaming — no full-file buffering | Yes |
| Roundtrip: write then read back identically | Yes |
| Cell formatting, fonts, colors | No |
| Merged cells | No |
| Formulas | No |
| Charts or images | No |
| Multiple sheets | No |
| Column widths, row heights | No |
| Headers, footers, print settings | No |
| Password protection | No |

StreamScheme does one thing: move typed tabular data in and out of xlsx as fast as possible. If you need presentation, use a full Excel library.

---

## Acknowledgments

StreamScheme's date format detection code is adapted from
[MiniExcel](https://github.com/MiniExcelFinancial/MiniExcel) (Apache 2.0),
which credits [ExcelNumberFormat](https://github.com/andersnm/ExcelNumberFormat)
(MIT) by andersnm.

[SpreadCheetah](https://github.com/sveinungf/spreadcheetah) (MIT) served as
inspiration and the primary performance comparison baseline.

---

## Status

Concept - work in progress.
