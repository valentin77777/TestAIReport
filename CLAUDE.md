# Project notes for Claude

## 1C code style
- Write all 1C code and metadata in **English syntax**: `&AtClient`, `Procedure … EndProcedure`,
  `New SpreadsheetDocument`, English names for objects, forms, templates, areas, parameters, commands.

## Spreadsheet templates (mxl-compile)
- Column widths in Template.xml are NOT characters. The whole A4 landscape page is ~780 units,
  A4 portrait ~540. Default width 10 is very narrow.
- Reference widths from a template the user adjusted by hand in 1C (TestReport):
  - No. column: 61
  - Client: 56
  - Contract: 74
  - Title row height: 75
- Rule of thumb: short columns (No., dates) at least ~40–60, text columns (names, contracts) ~60–120.
  Or use `"page": "A4-landscape"` with `"Nx"` proportions so widths fill the page.
- Buttons on forms: set `<Type>UsualButton</Type>` explicitly (1C writes it on save).
- form-edit generates handler stubs in Russian syntax (`&НаСервере`, `Процедура`), and companion
  element names with Russian suffixes (`КонтекстноеМеню`, `РасширеннаяПодсказка`) — rewrite them in English.

## HTML widgets on forms
- HTML field on Windows thin client may use the Internet Explorer 11 engine: write ES5 JavaScript,
  flexbox with -ms- prefixes, inline SVG, `<meta http-equiv="X-UA-Compatible" content="IE=edge">`,
  no external CDNs. Keep the HTML in a Text template and fill placeholders from code.
