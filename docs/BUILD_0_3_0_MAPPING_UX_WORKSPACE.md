# Pulse 0.3.0 — Mapping UX Workspace

This slice restores the useful hierarchy-browsing behavior of the approved
Lovable concept without imitating a web accordion. It remains a façade over
the established stable IDs, Mapping Rules, Effective Mapping resolver and
Weekly Mapping Attention projection.

## Visible workspace

`Mapping` presents:

1. weekly Mapping health and Performance-classification freshness;
2. a `Browse by` control for `Source Category` or read-only `Reporting Group` validation;
3. the relevant business-name selector and a compact selected-context summary;
4. one shared member workspace near the top of the sheet;
5. an intentional bulk-action area in Source Category mode; and
6. secondary 93-category and nine-group overview tables below the working area.

The stable-ID catalogs and the authoritative Main/Subcategory/Product audit
tables remain available on hidden `_Mapping_Audit`. Their table names and data
contracts are unchanged; they no longer create a technical tail beneath the
normal Mapping workspace.

The Main Category overview is navigation, not a bulk-selection table. Mapping
an entire Main Category requires choosing `Entire shown category` in the
separate `Apply to` control.

The shared member workspace shows business names, current Reporting Group,
`Inherited`, `Custom`, `Unmapped`, `Excluded`, `Identity Pending`, `Conflict`
or `Inactive Target`, plus historical Facts/Sales and attention. Stable IDs,
rule lineage and Historical Quantity remain in hidden columns and retained
engineering tables.

The health area is a compact set of Mapped, Unmapped, Identity Pending,
Conflict and Inactive Target cards showing exact Product, Fact and Sales
impact. It contains no cache IDs or fingerprints.

## Reporting Group validation

`Browse by = Reporting Group` exposes a business-name selector for the nine
active Reporting Groups. The shared member workspace becomes read-only and shows Product, source
Main Category, Subcategory, Sales Account, inherited/custom state and 85-week
Facts/Sales. It is a formula-driven façade over a hidden one-row-per-Product
catalog on `_Mapping_Audit`; ProductID and ReportingGroupID remain backstage.

Only Products whose current accepted resolver status is `Mapped` appear.
Unmapped, explicit exclusions, Identity Pending, Conflict and Inactive Target
remain visible in Mapping health and Source Category attention but are not
members of an active Reporting Group. The nine-row overview reconciles unique
Product membership to the mapped weekly population without inferring from POS
names or categories.

Reporting Group membership is intentionally read-only in this slice. Mapping
changes still use the established Source Category selection and atomic bulk
action. Attempting to submit an action while browsing by Reporting Group fails
before any rule write.

## Category and member interaction

`Show category` selects one Main Category working set. `View` filters that set
to `All`, `Unmapped`, `Custom`, `Identity Pending` or `Excluded`. These controls
recalculate through same-sheet Excel formulas; no Office Script rerun is
required merely to browse.

The visible member range is one fixed 400-row Excel Table over hidden,
deterministically sorted catalogs on `_Mapping_Audit`. Source Category mode
uses its accepted 150-row selectable bound and Reporting Group mode uses the
same physical area with a 400-row read-only bound. This avoids editable dynamic
spill ranges, keeps the two browse modes visually coherent, and supports the
current largest category (119 hierarchy members) and Reporting Group (352
Products) with explicit capacity checks.

The member `Select` dropdown offers `Yes` and `No`; blank remains a valid
unselected value for preserved workbook state. `No` makes individual
deselection explicit without changing stable-ID selection or batch semantics.

After any member is selected, Browse by, Category and View validation sources
collapse to their current values. This prevents normal Excel-for-web dropdown
use from rebinding selected rows to different stable IDs. Apply or clear the
selection before changing browse mode, category or view.

Successful dropdown wiring remains backstage and does not display a QA/debug
readiness banner on the user-facing Mapping workspace.

The upper Mapping workspace uses a compact fixed geometry at 100% zoom. Control
values, short states and impact measures are centered; identity and explanatory
text stays left-aligned. The Action and Assign-to controls retain enough width
for every approved option without consuming the free-text space reserved for
Notes. Mapping health cards keep each state adjacent to its Product, Fact and
Sales values.

## Atomic bulk actions

The visible actions are:

- `Assign Reporting Group` — creates a new rule or replaces an existing custom
  rule at each selected stable node;
- `Leave Unmapped` — creates the established Product-only explicit exclusion;
- `Remove custom mapping` — deactivates the explicit rule at each selected node
  so normal parent inheritance resumes.

The merged Action and Assign-to controls use same-sheet range-backed dropdown
sources so Excel for the web can resolve them after the final Mapping geometry
is applied. Action choices recalculate from the current stable-ID selection:
assignment is offered for a safe batch, `Leave Unmapped` only for selected
Products, and `Remove custom mapping` only when every requested node has an
explicit active rule (or the intentionally selected whole category does).
Assign-to exposes the nine active Reporting Group business names only when
assignment is the selected action. Reporting Group browse mode and Identity
Pending members remain read-only.

The user can apply an action to selected members or intentionally to the entire
shown Main Category. A broad parent change never deletes or overwrites lower
Subcategory/Product overrides.

Before any write, the script validates the complete batch. It rejects:

- Identity Pending, Conflict or Inactive Target selections;
- duplicate stable nodes;
- a selected Subcategory together with one of its Products;
- non-Product `Leave Unmapped` requests;
- `Remove custom mapping` where any selected node has no custom rule; and
- overlapping active rules at any selected node.

New rule rows are appended together and prior rules are deactivated together.
If either operation fails, appended rows are removed and the original Status
column is restored. There is no partial accepted batch.

## Authority and freshness

`tblWeeklyMappingAttention` supplies the complete 1,237-Product universe and
85-week impact. The existing Phase 1 resolver is applied to those stable
Product identities using current Mapping Rules and Reporting Groups; the
workspace does not introduce another resolver or mapping authority.

Any requested mapping mutation rebuilds Effective Mapping and surfaces
`Performance refresh required`. The Active weekly cache is not rewritten in
this slice. Identity Pending remains distinct, visible and review-only.

## Non-goals

This slice does not resolve pending identities, rebuild weekly history, alter
Product hierarchy authority, redesign Mapping Rules, change Performance,
Reports or Imports, or introduce another registry, resolver or metric engine.
