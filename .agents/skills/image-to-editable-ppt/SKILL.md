---
name: image-to-editable-ppt
description: Convert one or more reference images, screenshots, diagrams, posters, infographics, or scanned slide images into visually faithful, fully editable, portable PowerPoint (.pptx) slides. Use when the user asks to turn PNG/JPG screenshots into editable PPT, recreate a slide from an image, preserve the source layout and visual hierarchy, separate every element, redraw icons as SVG, embed images so the deck works on other computers, or fix grouping, bold-toggle, font-family, font-size, arrow, native-table, missing-image, SVG fallback, portability, and high-resolution export problems in an image-derived deck.
---

# Image to Editable PPT

Rebuild reference images as native, editable PowerPoint slides without placing the full reference image as a slide background. Produce a self-contained deck that remains clear and functional when transferred to another computer.

## Required companion workflow

Use the installed `presentations:Presentations` skill for every PPTX deliverable. Read its `SKILL.md`, required style guidance, Artifact Tool documentation, and QA requirements before implementation. Use `@oai/artifact-tool` from a JavaScript ES module; do not use `python-pptx`.

Read [references/reconstruction-workflow.md](references/reconstruction-workflow.md) before planning or authoring.

## Core workflow

1. Inspect every reference image at original resolution. Record pixel dimensions, aspect ratio, palette, typography, layout regions, reading order, repeated components, icons, images, charts, arrows, and captions.
2. Resolve the slide ratio explicitly. If the reference ratio differs from the requested ratio, preserve all content and relationships while adapting spacing; never silently crop important content.
3. Create a layer map. Treat each text block, icon, image, line, arrow, border, background surface, and diagram node as an independent object.
4. Rebuild the slide with editable text boxes and native shapes. Recreate every actual row-and-column table with `slide.tables.add(...)`, not a grid assembled from rectangles and lines. Create every semantic arrow as a native PowerPoint connector with an arrowhead on its target end; never simulate an arrow with disconnected line segments or a separate triangle. Use SVG for icons and line art. Use embedded high-resolution PNG for photographic or raster content. Do not use the full source image as a background, overlay, mask, or hidden tracing layer in the delivered deck.
5. Match the reference's hierarchy, alignment, colors, corner radii, line weights, whitespace, and text wrapping. Preserve source font sizes when a reference is supplied.
6. Export with Artifact Tool to a new PPTX. Preserve the source artifact when editing an existing deck.
7. Run `scripts/normalize_portable_pptx.mjs` on the exported PPTX. This removes grouping locks, neutralizes paragraph-default bold/font-family/font-size inheritance while preserving run formatting, rejects automatic text shrink by default, creates real embedded PNG fallbacks for SVG assets, validates native connectors and native tables when requested, and rejects external media links.
8. Render every final slide, inspect it at full size, run the presentation overflow test, and compare it with the reference. Fix clipping, wrapping, overlaps, broken arrows, blurry assets, and missing fallbacks before delivery.

## Non-negotiable output properties

- Keep all visible text editable.
- Keep each meaningful visual element independently selectable.
- Allow elements to be grouped and ungrouped in PowerPoint.
- Store bold only on runs that are initially bold; keep paragraph defaults neutral so users can cancel bold normally.
- Store font family and font size on text runs, not paragraph defaults. Users must be able to select text and change both properties normally in PowerPoint.
- Do not enable `normAutofit` or `spAutoFit` unless the user explicitly requests automatic fitting; automatic shrink can make manual font-size changes appear ineffective.
- Build relationship arrows with native PowerPoint connectors (`slide.shapes.connect`). In the current Artifact Tool exporter, put the target arrowhead in `tail`, for example `head: { type: "none" }` and `tail: { type: "triangle" }`.
- Do not construct arrows from multiple straight-line shapes, a line plus an independent triangle, or rasterized arrow artwork.
- Build actual tabular data as native PowerPoint tables. Users must be able to select the table and use PowerPoint's Table Design/Layout commands to add rows or columns, resize cells, and edit borders or fills.
- Do not imitate a table by arranging independent rectangles, line shapes, and text boxes. Decorative grids that are not tabular data are exempt.
- Embed all media inside the PPTX. Do not emit `TargetMode="External"` image relationships.
- For every embedded SVG produced by Artifact Tool, include a non-empty transparent PNG fallback at high resolution for older PowerPoint and WPS versions.
- Keep SVG icons as vector assets when requested; PNG fallbacks supplement rather than replace them.
- Never deliver a deck that depends on a temporary, hidden, cache, or source-image folder.
- Verify the final PPTX archive, rendered slide, and overflow results.

## Portable normalization

Run after the Artifact Tool export:

```bash
node scripts/normalize_portable_pptx.mjs \
  --input /absolute/path/raw.pptx \
  --output /absolute/path/final-portable.pptx \
  --sharp-module /absolute/path/to/node_modules/sharp \
  --min-native-connectors EXPECTED_ARROW_COUNT \
  --min-target-arrows EXPECTED_ARROW_COUNT \
  --min-native-tables EXPECTED_TABLE_COUNT
```

Obtain the bundled Node module path from the workspace dependency loader when available. The script accepts either the `sharp` package directory or its entry module. It modifies all slide XML files, never the source PPTX.

## Delivery

Return only the final PPTX unless the user requests previews or working files. State that media is embedded, SVG fallbacks are present, grouping is enabled, bold/font-family/font-size remain editable, arrows are native connectors, data tables are native tables, and visual/overflow QA passed only when each claim was actually verified.
