# Project instructions

This project vendors the image-to-editable-ppt skill in `.agents/skills/image-to-editable-ppt`.
For image-to-PPT work, read that skill and its reconstruction reference, then the installed `presentations:Presentations` companion skill and its required guidance.

- Use the existing isolated checkout. Do not create a Git worktree unless the user explicitly requests one.
- Run `bash scripts/check-ppt-environment.sh` from this project to validate dependencies and functional PPTX generation. It creates ignored outputs under `.local/ppt-checks/`.
- Source `scripts/ppt-runtime.sh` for the platform-supplied runtime. Do not replace the private Artifact Tool or fall back to another authoring library when a core dependency is unavailable. Report the actual blocker.
- Use Artifact Tool from JavaScript ES modules. Keep reference-image text editable, use native connectors and native tables, embed media, run the upstream normalizer, then validate and render the final PPTX.
- Inspect every rendered deliverable and perform its overflow checks. A synthetic smoke test proves the toolchain, not reconstruction fidelity for a user's images or native PowerPoint/WPS interaction.
- The four vendored skill files are unmodified upstream copies. Their source revision and SHA256 checksums are in `docs/image-to-editable-ppt-source.json`. Preserve this provenance when intentionally upgrading them.
- Keep temporary PPTX files, previews and runtime-specific links in ignored `.local/` directories. Never save credentials in project files.
