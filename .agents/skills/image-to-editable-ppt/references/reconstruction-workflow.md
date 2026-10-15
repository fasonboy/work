# Reconstruction workflow

## 1. Analyze the reference

Create a compact inventory before coding:

- canvas size and aspect ratio;
- major regions and exact reading order;
- repeated card or panel geometry;
- background, border, accent, and text colors;
- font families, sizes, weights, alignment, and wrapping;
- raster images versus vector-like icons;
- connectors, arrows, charts, tables, and captions;
- elements that should form logical PowerPoint groups.

Use the original image only as a visual reference. Do not include it in the final slide as a full-page image, hidden layer, or tracing overlay.

## 2. Choose reconstruction primitives

| Source element | PowerPoint implementation |
| --- | --- |
| Title, label, paragraph, caption | Editable textbox |
| Card, panel, header band, border | Native PowerPoint shape |
| Simple line | Native PowerPoint line |
| Semantic arrow between elements | Native PowerPoint connector with a target-end arrowhead |
| Icon or line illustration | Embedded SVG plus embedded PNG fallback |
| Photo, texture, complex raster region | Embedded high-resolution PNG |
| Actual row-and-column table | Native PowerPoint table created with `slide.tables.add(...)` |
| Decorative grid without tabular semantics | Native shapes when a table object would be misleading |
| Complex scientific illustration | Separate generated/located image asset, never the whole slide |

Create node shapes first, then connect their shape objects with `slide.shapes.connect(...)`. Use straight or elbow connectors to match the reference. In the current Artifact Tool exporter, `tail` is the target end, so a typical directed connector uses `head: { type: "none" }` and `tail: { type: "triangle", width: "sm", length: "sm" }`. Never assemble a semantic arrow from several independent line shapes or from a line plus a separate triangle.

Create every visible data table as one native table object. Put headers and body values in table cells, then style cell ranges, row heights, column widths, borders, fills, margins, and alignment through the table API. Do not redraw a data table with independent rectangles, rules, and text boxes merely to match its appearance; that prevents users from applying PowerPoint table operations.

## 3. Adapt aspect ratio

When the reference and requested slide ratios differ:

1. Preserve every required element.
2. Preserve relative ordering and major alignments.
3. Rebalance gutters and vertical spacing.
4. Apply only modest nonuniform compression to modular layouts.
5. Explain any unavoidable adaptation in progress commentary, not on the slide.

## 4. Author for editability

- Give important shapes stable names.
- Keep text separate from decorative shapes when independent editing is useful.
- Use one textbox per semantic text block.
- Avoid flattening charts, labels, or arrows into one image.
- Avoid group-prohibiting locks such as `a:spLocks noGrp="1"`.
- Use logical native groups only when the user benefits from moving a module as one object; the group must remain ungroupable.

## 5. Author typography for normal PowerPoint editing

Artifact exports may duplicate bold, font size, and font-face settings at both paragraph-default and run levels. Paragraph defaults can override later PowerPoint edits, causing bold to reappear or font-family/font-size changes to seem ineffective.

Normalize as follows:

- remove `b` and `sz` from `a:defRPr` and `a:endParaRPr`;
- remove `a:latin`, `a:ea`, and `a:cs` font-face children from `a:defRPr` and `a:endParaRPr`;
- remove run-level `b="0"` as redundant;
- retain run-level `b="1"` only for text that should initially appear bold;
- retain run-level `sz` and run-level `a:latin`/`a:ea`/`a:cs` so the initial appearance remains faithful;
- do not use `a:normAutofit` or `a:spAutoFit` unless automatic fitting was explicitly requested.

This keeps the initial formatting on visible runs while allowing users to change bold, font family, and font size normally.

## 6. Embed media portably

All image relationships must point inside the PPTX package. Reject any relationship using `TargetMode="External"`.

Modern Office stores an SVG and a PNG fallback for the same picture. Some exporters create a tiny transparent placeholder PNG. Older PowerPoint or WPS may ignore the SVG and display that placeholder, making icons disappear on another computer.

For each `ppt/media/*.svg` with a same-basename `.png` fallback:

1. rasterize the SVG at high density;
2. resize to a 1024 × 1024 transparent canvas with contain fit;
3. overwrite only the fallback PNG inside the copied PPTX package;
4. retain the SVG relationship and vector payload.

## 7. Validate

Verify all of the following:

- ZIP integrity passes;
- external media relationship count is zero;
- every SVG has a nontrivial PNG fallback;
- `noGrp="1"` count is zero;
- paragraph-default bold count is zero;
- paragraph-default font-size count is zero;
- paragraph-default font-face count is zero;
- visible text retains run-level font size and font face where those were authored;
- automatic-fit tag count is zero unless explicitly allowed;
- every expected semantic arrow is a `p:cxnSp` native connector;
- every directed connector has a target-end arrowhead, and no arrow is a chain of disconnected line shapes;
- every expected data table is an `a:tbl` native PowerPoint table;
- each native table has the expected row and column structure, and no data table is a grid of unrelated shape objects;
- all slides render successfully;
- overflow test passes;
- rendered slide matches the reference at full size;
- no full-slide source image exists in the delivered deck.
