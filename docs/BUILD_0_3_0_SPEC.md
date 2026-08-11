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

### Supported mapping scopes in 0.3.0

1. Source Main Category
2. Source Subcategory
3. Product

This matches the hierarchy already present in the source data and gives the user a practical high-level mapping path with product exceptions.

### Deterministic precedence

`Product override > Source Subcategory mapping > Source Main Category mapping > Unmapped`

Examples:

- Map source main category `Drinks` to a Reporting Group only if the user intentionally decides the whole branch is safe.
- Map subcategory `Soft Drinks` to `Non-Alcohol` even if its main category is broader.
- Override an individual mocktail product to `Non-Alcohol` when it lives inside a campaign/welcome-drinks subcategory.
- Override `Drink Mix` products to `Spirits/Cocktails` even if the source structure makes them appear non-alcoholic.

### Explicit human control

No automatic fuzzy classification belongs in the deterministic resolver.

The product may later suggest mappings, but a suggestion must be accepted by a human before it becomes an authoritative rule.

### Effective dates

Retain effective-date capability already present in Remap Rules.

For any date-aware resolution:

- an active rule applies only inside its effective interval;
- blank end date means open-ended;
- two overlapping active rules for the same scope object at the same precedence are a validation conflict;
- conflicting rules must not be silently resolved by row order.

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
- estimated NOK impact
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

### Phase 2 — metric integration

- change metric membership from legacy ReportingCategoryID to EffectiveReportingGroupID;
- update selector helper lists to active Reporting Groups;
- migrate Category Sales Share presentation to Reporting Group Sales Share;
- keep current/comparison behavior.

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
