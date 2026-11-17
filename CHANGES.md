# What changed in this update (and where to edit it yourself)

| # | Change | File / place to edit |
|---|---|---|
| Crash | "Label 'Mfg Date' not found": labels split by Word are now matched across runs | coa/docx_builder.py `_fill_header` |
| 1 | Text no longer touches cell lines (padding + centred) | docx_builder.py `CELL_PAD`, `_pad_and_centre` |
| 2 | FG / RM / WS auto-detected from product name + batch | coa/loader.py `detect_category` |
| 3 | Sample Qty per category (FG 10g, RM 10g, WS 1ml) | coa/resolver.py `CATEGORY_PRESETS` |
| 4 | Packaging default "Original Packaging" (type to override) | resolver.py `DEFAULT_PACKAGING` |
| 5 | cfu/g or cfu/ml from pack size; manual "Result unit" box | resolver.py `pack_unit`, page 3 "Result unit" |
| 6 | "Is it RM & WS sample? Yes" tick = edits affect only that COA | static/js/customize.js (`O`, `store`, `showSample`) |
| 7 | Remark line 2 = "Above test done by as per In-house STP." | docx_builder.py `REMARK_LINE_2`; templates/_preview.html |
| 8 | New colours, spinners, home + download pages | static/css/app.css, base.html, upload.html, download.html |
| 9 | Brand = home; step links clickable, state kept | base.html, app.py `nav_state`, `prefs.json` |
| 10 | Review: Sr. No., Product Name, Date of Analysis/Release, search highlight | select.html, static/js/select.js |
