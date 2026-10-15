import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const run = process.env.PPT_RUN_DIR;
const project = process.env.PPT_PROJECT_DIR;
const modules = process.env.RUNTIME_NODE_MODULES;
const skillDir = process.env.PRESENTATIONS_SKILL_DIR;
assert(run && project && modules && skillDir, 'Run bash scripts/check-ppt-environment.sh');
const requireRuntime = createRequire(path.join(modules, '__ppt_check__.cjs'));
const artifactEntry = requireRuntime.resolve('@oai/artifact-tool');
const sharpEntry = requireRuntime.resolve('sharp');
const { Presentation, PresentationFile, FileBlob } = await import(pathToFileURL(artifactEntry).href);
const sharp = requireRuntime('sharp');
const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir, 'container_tools/artifact_tool_utils.mjs')).href);
const artifactRoot = path.dirname(path.dirname(artifactEntry));
for (const doc of ['API_QUICK_START.md', 'api/API_DOCS.md']) {
  await fs.access(path.join(artifactRoot, 'docs/presentations', doc));
}
const artifactPackage = JSON.parse(await fs.readFile(path.join(artifactRoot, 'package.json'), 'utf8'));
const output = path.join(run, 'output');
await fs.mkdir(output, { recursive: true });

const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64"><circle cx="32" cy="32" r="22" fill="#2563eb"/></svg>';
const png = await sharp(Buffer.from(svg)).resize(1024, 1024).png().toBuffer();
const metadata = await sharp(png).metadata();
assert(metadata.width === 1024 && metadata.height === 1024 && metadata.hasAlpha);
await fs.writeFile(path.join(run, 'sharp-test.png'), png);
const payload = 'zip/unzip functional roundtrip\n';
await fs.writeFile(path.join(run, 'zip-test.txt'), payload);
execFileSync('zip', ['-q', 'zip-test.zip', 'zip-test.txt'], { cwd: run });
execFileSync('unzip', ['-t', path.join(run, 'zip-test.zip')]);
assert.equal(execFileSync('unzip', ['-p', path.join(run, 'zip-test.zip'), 'zip-test.txt'], { encoding: 'utf8' }), payload);
console.log('PASS: sharp SVG rasterization and zip/unzip roundtrip');

execFileSync(process.env.RUNTIME_NODE, [path.join(skillDir, 'container_tools/mark_artifact_operation_started.mjs'), '--operation-kind', 'create', '--expected-output-count', '1', '--output-format', 'pptx'], { stdio: 'inherit' });
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = deck.slides.add();
slide.background.fill = '#ffffff';
function textShape(name, text, x, y, width, height, fill = 'none', fontSize = 26, bold = false) {
  const shape = slide.shapes.add({ name, geometry: fill === 'none' ? 'textbox' : 'rect', position: { left: x, top: y, width, height }, fill, line: { fill: 'none', width: 0 } });
  shape.text = text;
  shape.text.style = { typeface: 'Liberation Sans', fontSize, bold, color: '#0f172a', autoFit: 'none', verticalAlignment: 'middle', insets: 10 };
  return shape;
}
textShape('editable-title', 'Editable PPT environment test', 64, 48, 1040, 70, 'none', 38, true);
const left = textShape('source', 'Reference image', 64, 190, 330, 90, '#dbeafe');
const right = textShape('target', 'Editable objects', 650, 190, 330, 90, '#dcfce7');
slide.shapes.connect(left, right, { kind: 'straight', fromSide: 'right', toSide: 'left', line: { fill: '#2563eb', width: 3 }, head: { type: 'none' }, tail: { type: 'triangle', width: 'med', length: 'med' } });
const table = slide.tables.add({ rows: 3, columns: 2, left: 64, top: 370, width: 916, height: 180, values: [['Object', 'Implementation'], ['Text', 'Editable textbox'], ['Data', 'Native PowerPoint table']] });
for (let r = 0; r < 3; r++) {
  table.rows[r].height = 60;
  for (let c = 0; c < 2; c++) {
    const cell = table.getCell(r, c);
    cell.fill = r === 0 ? '#dbeafe' : '#f8fafc';
    cell.text.style = { typeface: 'Liberation Sans', fontSize: 24, bold: r === 0, autoFit: 'none', color: '#0f172a' };
  }
}
table.borders.assign({ fill: '#cbd5e1', width: 1, style: 'solid' });
slide.images.add({ svg, alt: 'Embedded SVG icon with PNG fallback', position: { left: 1080, top: 190, width: 96, height: 96 } });
const raw = path.join(run, 'raw.pptx');
const portable = path.join(run, 'portable.pptx');
await (await PresentationFile.exportPptx(deck)).save(raw);
const normalization = execFileSync(process.env.RUNTIME_NODE, [path.join(project, '.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'), '--input', raw, '--output', portable, '--sharp-module', sharpEntry, '--min-native-connectors', '1', '--min-target-arrows', '1', '--min-native-tables', '1'], { encoding: 'utf8' });
await fs.writeFile(path.join(run, 'normalization.json'), normalization);
const normalized = JSON.parse(normalization);
assert.equal(normalized.slides, 1);
assert(normalized.svgCount > 0 && normalized.pngFallbackCount === normalized.svgCount, 'Every SVG must have a generated PNG fallback');
console.log('PASS: PPTX export and upstream portability normalizer');
const finalPath = path.join(output, 'smoke.pptx');
await finalizePresentation({ workspaceDir: run, candidatePath: portable, finalPath, pythonExecutable: process.env.RUNTIME_PYTHON, integrityValidatorPath: path.join(skillDir, 'container_tools/inspect_presentation_package_integrity.py'), layoutValidatorPath: path.join(skillDir, 'container_tools/inspect_presentation_layout_geometry.py'), layoutArgs: ['--expected-slide-size-emu', '12192000,6858000', '--validate-bullet-geometry', '--validate-heading-fit', '--require-native-table-slide', '1'], explicitTotalSlideCount: 1, requiredNativeTableOwnerSlides: [1], fontPolicy: { basis: 'design', families: ['Liberation Sans'] }, verifyArtifactToolImport: true, receiptPath: path.join(run, 'validation.json') });
const imported = await PresentationFile.importPptx(await FileBlob.load(finalPath));
assert.equal(imported.slides.items.length, 1);
const preview = await imported.export({ slide: imported.slides.items[0], format: 'png', scale: 1 });
const rendered = Buffer.from(await preview.arrayBuffer());
const previewMeta = await sharp(rendered).metadata();
assert(previewMeta.width >= 1280 && previewMeta.height >= 720);
await fs.writeFile(path.join(output, 'slide-1.png'), rendered);
await fs.writeFile(path.join(run, 'dependencies.json'), JSON.stringify({ tested_at: new Date().toISOString(), node: process.version, presentations_skill: skillDir, artifact_tool: artifactPackage.version, sharp: sharp.versions.sharp, libvips: sharp.versions.vips, normalization: normalized, render: { width: previewMeta.width, height: previewMeta.height } }, null, 2));
console.log('PASS: official finalizer, PPTX re-import and final slide PNG rendering');
