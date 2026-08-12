# Pulse Build 0.3.0 — Mapping + Reporting Groups

Status: implementation specification
Starting checkpoint: `Pulse_Build_0_2_0_QA.xlsx.xlsx`

## 1. Goal

Build 0.3.0 introduces the semantic classification layer needed to stop exposing raw POS taxonomy directly to Performance.

Target flow:

`Raw POS data -> Mapping -> Reporting Groups -> Metrics/KPIs -> Performance`

More completely within the existing architecture:

`Source data -> Adapter/Staging -> Validation -> Published Facts -> Mapping -> Reporting Groups -> Metric engine -> Performance/Reports`

The build must preserve the validated 0.2.0 behavior around source traceability, independent current/comparison scopes, channel neutrality, and deterministic metrics.

## 2. Reconciled starting state

The validated 0.2.0-QA workbook contains 36 sheets. Relevant current structures include:

- `Reporting Categories` — ~130 POS-derived source-default categories, not yet a true business reporting taxonomy.
- `Source Classifications` — source Main/Sub combinations with a default reporting-category ID.
- `Products` — stable product records with `EffectiveReportingCategoryID`.
- `Remap Assistant` — current exception-entry surface.
- `Remap Rules` — override table; current design states Product rules win over source-subcategory rules.
- `Effective Categories` — read-only mapping resolution output.
- `_Metric_Calc` — selector/helper surface used by Performance.
- `Performance` — current Category Sales Share vertical slice.
- `Reports` — consumes the Performance result.

The 0.2.0 QA patch fixed readability and replaced fragile literal dropdowns with range-backed selectors. Those fixes must not regress.

## 3. Architectural correction in 0.3.0

The existing term `Reporting Category` is overloaded. In 0.2.0 it is mostly a generated mirror of POS/source subcategories. That source-derived structure is not the same thing as the business-defined analysis taxonomy required by Pulse.

Build 0.3.0 therefore separates:

### Source structure

Examples:

- Source Main Category
- Source Subcategory
- Product
- Sales Account
- Channel

These remain factual/source dimensions.

### Business semantic structure

`Reporting Group`

A Reporting Group is Pulse/business-owned and exists specifically to support analysis and reporting independently of how the POS happens to categorize an item.

Reporting Groups are not KPIs.

## 4. Reporting Group model

Create an authoritative Reporting Groups table/sheet.

Recommended columns:

| Field | Purpose |
|---|---|
| ReportingGroupID | Stable machine ID, e.g. `RPG-0001` |
| ReportingGroupName | Human-readable business name |
| DomainID | Commercial Performance initially |
| Active | Yes/No; controls downstream availability |
| SortOrder | Presentation ordering |
| Description | Plain-language scope |
| Notes | Human-owned context/configuration |

Initial configurable groups:

1. Add-ons
2. Non-Alcohol
3. Spirits/Cocktails
4. Coffee & Tea
5. Beer & Cider
6. Desserts
7. Wine & Sake
8. Starters
9. Mains

These are seed configuration. The system must not assume this list is permanent.

### IDs and renaming

Mappings must reference `ReportingGroupID`, not display text. Renaming `ReportingGroupName` must not break existing mappings.

### Active/inactive behavior

- Active groups are available to Performance/Reports.
- Inactive groups remain stored and may retain mappings/history.
- Deactivation does not rewrite facts or delete rules.
- If a selected/mapped group becomes inactive, QA should expose the condition rather than silently remap it.

## 5. Mapping model

### Generic hierarchy model

The Mapping experience follows the validated Lovable-style concept: users browse the actual hierarchy supplied by the source/POS, understand the descendants beneath a node, map at the highest safe level, and add more-specific exceptions only where needed.

The architecture is hierarchy-based. It must not be permanently defined as exactly three fixed levels.

The currently observed Katria hierarchy is:

`Source Main Category -> Source Subcategory -> Product`

Build 0.3.0 may therefore author rules at Main Category, Subcategory, or Product level for Katria. These are the currently supported rule targets, not the permanent definition of Pulse's mapping architecture. Future adapters may expose additional or differently named hierarchy levels without requiring a redesign of Reporting Groups, mapping semantics, or metrics.

### Inheritance and deterministic precedence

The general resolution rule is:

**The most-specific applicable explicit mapping wins; otherwise inherit the nearest mapped ancestor; otherwise remain Unmapped.**

For the current Katria hierarchy, that resolves as:

1. Active Product mapping/override
2. Active Source Subcategory mapping
3. Active Source Main Category mapping
4. Unmapped / needs review

A higher-level mapping exists only when a human explicitly authors it. Descendants inherit that mapping automatically unless a more-specific explicit mapping applies. Inheritance is computed state, not a copied rule per descendant.

Examples:

- Map source main category `Drinks` to a Reporting Group only if the user intentionally decides the whole branch is safe.
- Map subcategory `Soft Drinks` to `Non-Alcohol` even if its main category is broader.
- Override an individual mocktail product to `Non-Alcohol` when it lives inside a campaign/welcome-drinks subcategory.
- Override `Drink Mix` products to `Spirits/Cocktails` even if the source structure makes them appear non-alcoholic.

A product override may intentionally differ from its inherited parent mapping. It must survive later remapping of any ancestor. Clearing or deactivating the override restores resolution from the nearest valid mapped ancestor.

The administration and audit surfaces must distinguish:

- explicit mapping authored on the current node;
- inherited mapping and the ancestor rule it came from;
- a more-specific descendant exception;
- Unmapped state;
- conflict or inactive-target state.

### Explicit human control

No automatic fuzzy classification belongs in the deterministic resolver.

The product may later suggest mappings, but a suggestion must be accepted by a human before it becomes an authoritative rule. Text similarity, names, or sales context must never silently create an effective mapping.

### Effective dates and conflicts

Retain effective-date capability already present in Remap Rules.

For any date-aware resolution:

- an active rule applies only inside its effective interval;
- blank end date means open-ended;
- two overlapping active rules for the same source node/scope at the same hierarchy level are a validation conflict;
- conflicting rules must not be silently resolved by row order;
- precedence across different hierarchy levels remains most-specific applicable explicit rule first.

## 6. Recommended workbook migration

Do not destroy existing source data or mapping structures.

### Add

- `Reporting Groups` — business-owned configuration table.
- `Mapping` (or evolve `Remap Assistant`) — primary administration surface for browsing hierarchy and assigning Reporting Groups.
- `Effective Mapping` (or evolve `Effective Categories`) — read-only resolution/audit surface.

### Preserve / migrate

- Preserve `Source Classifications` as source hierarchy.
- Preserve `Products` as product register.
- Preserve `Remap Rules` data/audit capability, but evolve target field from Reporting Category to Reporting Group.
- Preserve raw/fact/import sheets unchanged except for additive fields only when necessary for lineage.
- Preserve existing pages/functions unless an explicit migration is required.

### Legacy `Reporting Categories`

Do not delete it in 0.3.0.

Recommended handling:

- retain it for compatibility/traceability during the migration;
- clearly mark it as legacy/source-default classification rather than the new business taxonomy;
- stop using it as the authoritative Performance selector once Reporting Groups are live;
- avoid creating new business logic that depends on its POS-derived IDs.

A later validated build can decide whether to hide, rename, or retire the legacy sheet.

## 7. Mapping user experience

The Mapping surface should make three things easy:

1. Understand the POS hierarchy.
2. See the underlying sales/products affected by a mapping.
3. Assign or override a Reporting Group with minimal clicks.

### Suggested layout

A practical Excel-first design can use a browse/filter panel plus a mapping table rather than trying to mimic a full tree-control UI.

Recommended browse fields:

- Source System
- Source Main Category
- Source Subcategory
- Product
- Sales Account
- Current Effective Reporting Group
- Resolution Source (Product / Subcategory / Main / Unmapped)
- Sales NOK in selected inspection dataset (informational)
- Quantity in selected inspection dataset (informational)
- Mapping status

The sales columns are for understanding impact; they must not become part of rule identity.

### Mapping actions

The user should be able to:

- assign a main category to a Reporting Group;
- assign a subcategory to a Reporting Group;
- override one product;
- clear/deactivate a mapping rule without deleting source facts;
- see how many products inherit from the selected rule;
- identify product-level exceptions below a parent mapping;
- filter to Unmapped / Overrides / Conflicts / Inactive-target issues.

### Safety cue

When mapping at Source Main Category level, show the breadth of the change (e.g. affected subcategories/products) so high-level mapping is deliberate.

Do not block high-level mapping merely because it is broad.

## 8. Effective mapping output

Create one deterministic resolved row per product for the current rule state, with at least:

- ProductID
- SourceProductName
- SourceMainCategory
- SourceSubCategory
- SourceClassificationID
- MainCategoryRuleID / target
- SubcategoryRuleID / target
- ProductRuleID / target
- EffectiveReportingGroupID
- EffectiveReportingGroupName
- ResolutionSource
- ResolutionStatus

Suggested `ResolutionStatus` values:

- Mapped
- Unmapped
- Conflict
- Inactive Target

This surface is read-only output and should be suitable for QA/reconciliation.

## 9. Metric migration

The 0.2.0 formula is currently:

`Category Sales Share = category sales amount / total sales amount inside selected scope`

In 0.3.0 the semantic definition becomes:

`Reporting Group Sales Share = sales amount of facts/products resolving to the selected Reporting Group / total sales amount inside selected scope`

The denominator remains the selected scope total; it is not restricted to mapped items unless a future KPI explicitly defines otherwise.

Unmapped items therefore remain in the denominator for Sales Share. This prevents mapping incompleteness from artificially inflating mapped-group share.

The metric engine should identify group membership by stable `ReportingGroupID`.

### Future metric compatibility

The mapping layer should support later metrics without redesign:

- Quantity Share
- Sales NOK
- Quantity
- Attach Rate
- percentage-point / percentage change vs comparison
- NOK impact at the current selected-scope sales base
- trends and leaderboards

Do not implement all future metrics in 0.3.0 unless needed for validation. The priority is a correct reusable classification layer.

## 10. Performance behavior

Once 0.3.0 mapping is live:

- Performance must select from active Reporting Groups, not the ~130 POS-derived categories.
- Current/comparison dataset independence remains.
- Channel continues to be a scope/property dimension where exposed; it is not embedded in the metric definition.
- Performance and Reports consume the same centralized metric result.
- Do not remove the existing detailed restaurant view solely for simplification.

The immediate user-visible win is that Performance becomes business-readable while the raw POS taxonomy remains available deeper in the administration layer.

## 11. QA and acceptance criteria

Build 0.3.0 is not complete until all of the following are validated.

### Structural

- Reporting Groups table exists with unique stable IDs.
- Initial nine groups exist and are editable configuration.
- Active/inactive is enforced in downstream selector availability.
- Source Classifications and Products remain intact.
- Raw source worksheets and `_Sales_Facts` are not rewritten by mapping.

### Resolution

- Main-category rule correctly maps descendants without more-specific rules.
- Subcategory rule overrides Main-category rule.
- Product rule overrides both parent levels.
- Product override persists when a parent mapping later changes.
- Removing/deactivating a product override restores the next valid inherited mapping.
- Unmapped products remain explicitly Unmapped.
- Same-level overlapping active conflicts are surfaced and do not silently resolve.
- Inactive Reporting Group targets are surfaced.

### Reconciliation

For a selected dataset/scope:

`Mapped sales + Unmapped sales = total source/fact sales`

and

`Mapped quantity + Unmapped quantity = total source/fact quantity`

within normal numeric precision.

Changing a mapping may redistribute sales/quantity between Reporting Groups but must not change total sales or total quantity.

### Metric

- Reporting Group Sales Share matches an independently calculated test case.
- The denominator includes the full selected scope, including currently unmapped facts.
- Performance and Reports display the same result for the same selected scope/group.
- Current and comparison remain independently selectable.

### Regression

- 0.2.0-QA readability fixes remain.
- Performance/Reports do not return to tiny columns.
- Dropdowns remain authoritative range/table-backed, not giant literal lists.
- Existing import/publication structures continue to work.
- No existing pages are removed merely because 0.3.0 makes them less central.

## 12. Build implementation strategy

Use staged implementation rather than one monolithic script.

### Phase 1 — model + migration

- add Reporting Groups table;
- seed nine configurable groups;
- evolve rule schema to target ReportingGroupID and support Main/Sub/Product scopes;
- build deterministic Effective Mapping resolver;
- add QA outputs/tests;
- do not yet rewrite all Performance formulas.

### Phase 2A — metric contract, bridge, and reconciliation

- define deterministic Reporting Group metric membership and denominator rules;
- materialize one derived analysis row per immutable sales fact joined by
  ProductID to current Effective Mapping;
- apply current mapping state to historical facts for analysis without fact-date
  versioning;
- reject stale Effective Mapping before generating metric output;
- store MappingAsOfDate and a deterministic semantic mapping fingerprint;
- reconcile fact count, Sales NOK, and Quantity across Mapped, Unmapped,
  Conflict, and Inactive Target states;
- support only explicit human-authored legacy CAT/RPG equivalence definitions
  and visible side-by-side variance;
- preserve `_Metric_Calc`, Performance, Reports, KPI Registry, and KPI-0001 on
  the validated legacy path.

### Phase 2B — central calculation and minimal presentation cutover

- change active metric membership from legacy ReportingCategoryID to
  EffectiveReportingGroupID;
- update selector helper lists to active Reporting Groups;
- migrate Category Sales Share presentation to Reporting Group Sales Share;
- keep current/comparison behavior and optional channel scope;
- do not redesign Performance.

### Phase 2C — interactive Sales Performance model

- preserve `tblMetricRPGResults` as the authoritative KPI-0001 result layer;
- reuse its additive Restaurant-scope numerator and denominator components for
  arbitrary selected eligible-restaurant subsets;
- add stable-ID Yes/No restaurant and active Reporting Group selection, where
  Include state is authoritative and All versus Custom is derived;
- keep the single detail Reporting Group selector independent of matrix
  multi-selection;
- calculate share/PP Grand Total from summed numerators and denominators, and
  NOK Impact as aggregated current numerator minus aggregated comparison share
  × aggregated current denominator;
- add a Total column over the currently selected Reporting Groups. Current
  Sales NOK Total is the selected current numerator; denominator-derived modes
  use each scope denominator once;
- sort only the visible restaurant presentation by Total or a displayed RPG,
  using the full-precision numeric helper and never including Grand Total;
- keep current and comparison ImportIDs independently selectable;
- require Excel recalculation only for normal selection/display changes; and
- do not materialize restaurant combinations, add another metric engine, add
  new KPIs/channel UI, or redesign Performance.

### Phase 3 — Mapping UX + release QA

- make the hierarchy browse/mapping surface practical;
- show affected products and sales/quantity context;
- add mapping-status filters/summary;
- complete regression/reconciliation tests;
- generate the 0.3.0-QA workbook checkpoint.

## 13. Documentation changes required with 0.3.0

Update together with implementation:

- `README.md` current checkpoint/milestone status;
- `CHANGELOG.md`;
- `ROADMAP.md`;
- `docs/architecture.md` — explicitly name Mapping -> Reporting Groups;
- `docs/data-model.md` — replace overloaded Reporting Category concept with Reporting Group and mapping-rule scopes;
- `docs/kpi-definitions.md` — migrate Category Sales Share terminology when metric integration lands;
- build-specific QA checklist/expected fixture;
- Office Script README/release notes.

Do not erase historical 0.2.0 documentation; distinguish historical build definitions from current architecture.
