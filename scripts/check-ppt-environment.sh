#!/usr/bin/env bash
set -euo pipefail
PPT_PROJECT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PPT_PROJECT_DIR"
source scripts/ppt-runtime.sh
mkdir -p .local/ppt-checks
PPT_RUN_DIR="$(mktemp -d "$PPT_PROJECT_DIR/.local/ppt-checks/run-XXXXXX")"
export PPT_PROJECT_DIR PPT_RUN_DIR
"$RUNTIME_PYTHON" - <<'PY'
from pathlib import Path
import hashlib, json, os, re
root = Path(os.environ['PPT_PROJECT_DIR'])
meta = json.loads((root / 'docs/image-to-editable-ppt-source.json').read_text())
skill = root / meta['installed_directory']
actual = {str(p.relative_to(skill)) for p in skill.rglob('*') if p.is_file()}
assert actual == set(meta['files_sha256']), 'Skill file list differs from upstream manifest'
for name, digest in meta['files_sha256'].items():
    assert hashlib.sha256((skill / name).read_bytes()).hexdigest() == digest, f'Skill checksum mismatch: {name}'
for doc in skill.rglob('*.md'):
    text = doc.read_text()
    for link in re.findall(r'\]\(([^)]+)\)', text):
        if '://' not in link and not link.startswith('#'):
            assert (doc.parent / link.split('#')[0]).is_file(), f'Missing reference: {link}'
text = (skill / 'SKILL.md').read_text()
assert text.startswith('---\nname: image-to-editable-ppt\n')
assert 'Required companion workflow' in text
print('PASS: all 4 upstream skill files, SHA256 checksums, frontmatter and local references')
PY
"$RUNTIME_NODE" --check .agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs
"$RUNTIME_NODE" scripts/ppt-smoke-test.mjs
"$RUNTIME_PYTHON" scripts/verify-pptx.py "$PPT_RUN_DIR/output/smoke.pptx" > "$PPT_RUN_DIR/package-check.json"
"$RUNTIME_PYTHON" "$PRESENTATIONS_SKILL_DIR/container_tools/slides_test.py" "$PPT_RUN_DIR/output/smoke.pptx" > "$PPT_RUN_DIR/overflow.log" 2>&1
# This upstream runner may return exit 0 even for a detected overflow.
"$RUNTIME_PYTHON" - <<'PY'
from pathlib import Path
import json, os
run = Path(os.environ['PPT_RUN_DIR'])
log = (run / 'overflow.log').read_text()
print(log, end='')
assert 'Test passed. No overflow detected.' in log and 'ERROR:' not in log, 'Overflow check failed; inspect overflow.log'
report = json.loads((run / 'dependencies.json').read_text())
report['package_checks'] = json.loads((run / 'package-check.json').read_text())
report['overflow'] = 'passed'
report['status'] = 'passed'
report['scope'] = 'Synthetic one-slide generation, not image recognition or arbitrary-image reconstruction'
(run / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
print('PASS: functional dependency test; report:', run / 'report.json')
PY
