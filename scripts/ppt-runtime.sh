#!/usr/bin/env bash
# Source this file from a project script. Use only the cloud-supplied runtime.
set -euo pipefail
export RUNTIME_NODE="${CODEX_PRIMARY_RUNTIME_NODE:?Cloud Node.js runtime is missing}"
export RUNTIME_NODE_MODULES="${CODEX_PRIMARY_RUNTIME_NODE_MODULES:?Cloud Node modules are missing}"
export RUNTIME_PYTHON="${CODEX_PRIMARY_RUNTIME_PYTHON:?Cloud Python runtime is missing}"
export RUNTIME_BIN_DIR="${CODEX_PRIMARY_RUNTIME:?Cloud runtime is missing}/dependencies/bin/override"
for ppt_path in "$RUNTIME_NODE" "$RUNTIME_NODE_MODULES" "$RUNTIME_PYTHON" "$RUNTIME_BIN_DIR"; do
  [[ "$ppt_path" = /* && -e "$ppt_path" ]] || { echo "Missing absolute runtime path: $ppt_path" >&2; return 1; }
done
[[ -x "$RUNTIME_NODE" && -x "$RUNTIME_PYTHON" && -d "$RUNTIME_NODE_MODULES" && -d "$RUNTIME_BIN_DIR" ]]
export PRESENTATIONS_SKILL_DIR="${PRESENTATIONS_SKILL_DIR:-/opt/codex/skills/builtins/presentations}"
for ppt_file in SKILL.md style_guidelines.md references/implementation.md references/finalization.md container_tools/artifact_tool_utils.mjs container_tools/mark_artifact_operation_started.mjs container_tools/slides_test.py; do
  [[ -f "$PRESENTATIONS_SKILL_DIR/$ppt_file" ]] || { echo "Missing presentations:Presentations file: $ppt_file" >&2; return 1; }
done
for ppt_command in zip unzip; do
  command -v "$ppt_command" >/dev/null || { echo "Missing dependency: $ppt_command" >&2; return 1; }
done
export CODEX_NODE_MODULES="$RUNTIME_NODE_MODULES"
export PATH="$RUNTIME_BIN_DIR:$PATH"
