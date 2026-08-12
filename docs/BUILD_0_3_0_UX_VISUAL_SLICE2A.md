# Build 0.3.0 UX — Visual Slice 2A

## Purpose

Visual Slice 2A is a presentation-only pass over the six accepted primary Pulse
sheets. It establishes a restrained shared visual system, repairs Reports, and
makes Performance explanation readable without changing the accepted Phase 2C
calculation or interaction architecture.

The implementation is owned by
`office-scripts/Build_0_3_0_UX_Visual_Slice2A.ts`. Run it after Phase 2C and UX
Information Architecture Slice 1.

## Authoritative physical baseline

The implementation was inspected against
`Pulse_0.3.0_Phase2C_16of16_QAUX.xlsx.xlsx`, which was verified to contain:

- 47 worksheets;
- exactly six visible primary sheets in the accepted order;
- valid Overview links to the other five primary sheets;
- `tblPerformanceInteractionQA` with `QA-0302C-01` through
  `QA-0302C-16`;
- 16/16 `PASS`.

## Mandatory preflight

The script fails before visual mutation unless it finds:

- the exact 47-sheet checkpoint contract;
- the six visible sheets ordered as Overview, Performance, Reports, Imports,
  Mapping, Settings;
- all 41 supporting sheets in normal-hidden state;
- the five accepted Overview links;
- the accepted Reports-to-Performance formulas;
- the Phase 2C selection and QA tables;
- all 16 Phase 2C QA checks as `PASS`;
- a formula-free Performance Explain surface at its expected dynamic location.

## Exact presentation mutations

### Shared primary-sheet treatment

- Gridlines are disabled only on Overview, Performance, Reports, Imports,
  Mapping, and Settings.
- Title rows use navy `#172033`, white text, and restrained 18-point type.
- Subtitle rows use light blue `#EAF2FF`, explicit wrapping, and explicit row
  heights for Excel for the web.
- Each primary sheet is saved at A1; Overview remains the final active sheet.

### Performance

- Existing controls, selections, result card, matrix, Total, Grand Total, and
  sorting remain in their existing cells.
- Input selectors and Include columns use pale yellow `#FFF4D6`.
- Selection/status outputs use light blue `#EAF2FF`.
- Blank spacer rows are reduced without moving functional ranges.
- A concise note distinguishes Detail Reporting Group from matrix RPG
  selection.
- The former eight-row technical Explain block becomes four operational rows:
  Metric, Definition, Total, and Comparison.
- Explanation values span presentation-only merged cells B:H and use explicit
  row heights; the remaining former explanation rows are blank and compact.

### Reports

- Existing formulas remain linked to Performance exactly as before.
- Label and value columns receive practical Office Scripts widths.
- Current/comparison context and status rows wrap with explicit heights.
- Result rows receive right alignment while preserving the already accepted
  workbook-local share, pp, and NOK number formats.
- Section hierarchy and informational fills use the shared Pulse palette.
- No new formulas, matrix, report engine, or print system is introduced.

### Overview, Imports, Mapping, and Settings

- Overview keeps its existing values and navigation; only common title,
  subtitle, gridline, alignment, and clipping treatment is applied.
- Imports keeps its complete table and schema; only common title, subtitle,
  header wrapping, row height, and freeze-pane treatment is applied.
- Mapping keeps its action form, IDs, hierarchy tables, validations, and
  behavior; input cells are visually distinguished and existing instruction
  areas receive explicit wrapping/heights.
- Settings keeps all SettingID/Value semantics; widths, wrapped notes, and
  editable Value highlighting are normalized.

## Frozen architecture and postflight protection

The script does not write formulas, add/delete/resize tables, rename objects,
change sheet visibility/order, modify validations, add protection, or touch
hidden engineering sheets.

Before and after formatting, it fingerprints:

- every Performance formula;
- every Reports formula;
- values/formulas in Imports, Mapping, Settings, both Performance selection
  tables, and the Phase 2C QA table;
- table names, ranges, and headers for those protected tables.

It also verifies worksheet/table counts, IA visibility/order, Overview links,
gridline state, Reports clipping safeguards, and 16/16 Phase 2C QA.

## Rerun and rollback

The script is deterministic and rerunnable. It resets only its owned formats,
the approved Detail note, and the compact Explain copy. The four explanation
merges are explicitly unmerged before being recreated, preventing merge
accumulation. It adds no conditional formats or workbook objects on rerun.

Normal development remains in the canonical `Pulse_Current.xlsx`. Rollback
requires an intentionally created checkpoint/release copy; do not search for or
substitute an ad-hoc workbook. For a reproducible rebuild, run the accepted
Phase 2C script, IA Slice 1, then Visual Slice 2A. There is intentionally no
destructive reverse-migration script.

## Live validation

Use `tests/build-0.3.0-ux-visual-slice2a-checklist.md`. Static tests protect the
architecture, but Excel-for-web inspection remains required for final wrapping,
column widths, dropdown behavior, and visible number rendering.
