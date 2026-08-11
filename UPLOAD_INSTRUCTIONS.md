# Pulse pre-Codex repository patch

This folder mirrors only the repository paths that should be added or replaced before starting Codex Phase 1.

## Replace existing files

- `README.md`
- `CHANGELOG.md`
- `ROADMAP.md`
- `tests/README.md`
- `office-scripts/README.md`
- `releases/README.md`

## Add new files

- `AGENTS.md`
- `docs/BUILD_0_3_0_SPEC.md`
- `CODEX_TASK_0_3_0_PHASE1.md`

## Important

Do **not** delete other repository files or folders.

Do **not** upload this wrapper file (`UPLOAD_INSTRUCTIONS.md`) into the Pulse repository unless you want to keep it; it is only a checklist for the manual upload.

After these files are committed, the repository should state:

- validated checkpoint: **0.2.0-QA**
- active development: **0.3.0 — Mapping + Reporting Groups**
- 0.3.0 is **not yet released**
- mapping is hierarchical and Lovable-style: map at the highest safe level, inherit to descendants, and allow more-specific overrides.

Recommended next development branch:

`build/0.3.0-mapping-phase1`

Then give Codex the task in `CODEX_TASK_0_3_0_PHASE1.md`.
