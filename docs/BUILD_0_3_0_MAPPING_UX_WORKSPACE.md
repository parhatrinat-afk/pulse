# Pulse 0.3.0 — Mapping UX Workspace

This slice restores the useful hierarchy-browsing behavior of the approved
Lovable concept without imitating a web accordion. It remains a façade over
the established stable IDs, Mapping Rules, Effective Mapping resolver and
Weekly Mapping Attention projection.

## Visible workspace

`Mapping` presents:

1. weekly Mapping health and Performance-classification freshness;
2. `Show category` and `View` controls;
3. an intentional bulk-action area;
4. a read-only Main Category overview with current category rule, hierarchy
   size, 85-week impact and attention;
5. one bounded member workspace containing the shown category's Subcategories
   and Products.

The stable-ID catalogs and the authoritative Main/Subcategory/Product audit
tables remain available on hidden `_Mapping_Audit`. Their table names and data
contracts are unchanged; they no longer create a technical tail beneath the
normal Mapping workspace.

The Main Category overview is navigation, not a bulk-selection table. Mapping
an entire Main Category requires choosing `Entire shown category` in the
separate `Apply to` control.

The member workspace shows business names, current Reporting Group,
`Inherited`, `Custom`, `Unmapped`, `Excluded`, `Identity Pending`, `Conflict`
or `Inactive Target`, plus historical Facts/Sales and attention. Stable IDs,
rule lineage and Historical Quantity remain in hidden columns and retained
engineering tables.

## Category and member interaction

`Show category` selects one Main Category working set. `View` filters that set
to `All`, `Unmapped`, `Custom`, `Identity Pending` or `Excluded`. These controls
recalculate through same-sheet Excel formulas; no Office Script rerun is
required merely to browse.

The visible member range is a fixed 150-row Excel Table over a hidden,
deterministically sorted member catalog on `_Mapping_Audit`. Only its `Select`
cells are editable.
This avoids an editable dynamic-spill range while supporting the current
largest category (119 members) with a bounded capacity check.

After any member is selected, the Category and View validation sources collapse
to their current values. This prevents normal Excel-for-web dropdown use from
rebinding selected rows to different stable IDs. Apply or clear the selection
before changing category or view.

## Atomic bulk actions

The visible actions are:

- `Assign Reporting Group` — creates a new rule or replaces an existing custom
  rule at each selected stable node;
- `Leave Unmapped` — creates the established Product-only explicit exclusion;
- `Remove custom mapping` — deactivates the explicit rule at each selected node
  so normal parent inheritance resumes.

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
