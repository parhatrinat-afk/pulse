# Data Model

## Core concepts

### Import
A bounded dataset entering Pulse with identity, period metadata, validation state, and publication state.

### Fact
A standardized business observation derived from an accepted import. Facts should retain enough lineage to identify their source import and business entity.

### Restaurant
A stable organizational entity. The neutral base system must not depend on a specific restaurant chain's names.

### Product
A stable product identity independent of changing source-system labels or categories.

### Reporting Category (legacy 0.2.0 bridge)
The validated 0.2.0 workbook uses POS-derived default categories for its first
metric path. This structure is retained for compatibility during 0.3.0.

### Reporting Group
A business-owned semantic classification with a stable ID, editable display
name, active status, ordering, and human context. Mapping rules reference the ID.

### Source Classification
The category or classification supplied by the originating system.

### Mapping Rule
A durable explicit rule targeting a generic source hierarchy scope/node and a
Reporting Group ID. `RuleAction=Map` is the legacy/default behavior. The narrow
`RuleAction=Exclude` extension is valid only at Product scope, has no Reporting
Group target, and explicitly resolves that Product to Unmapped. Current
supported scopes are SourceMainCategory, SourceSubCategory, and Product.
Effective intervals are inclusive; overlapping active rules on the same
node/scope are conflicts regardless of action.

### Effective Mapping
One computed row per product showing applicable ancestor/product rules, the
winning rule, Reporting Group, explicit versus inherited state, and status.
Resolution is most-specific explicit mapping, otherwise nearest mapped ancestor,
otherwise Unmapped. An active Product exclusion wins at Product specificity and
records `ResolutionState=Explicit exclusion` plus the winning rule ID. When the
exclusion is deactivated, normal ancestor inheritance resumes.

### Metric Reporting Group Fact
A derived analytical row joining one immutable Sales Fact to its Product's
current Effective Mapping. It preserves fact lineage and source measures while
adding effective Reporting Group, resolver state, MappingAsOfDate, mapping
fingerprint, and metric refresh timestamp. It is regenerated analysis state and
must never replace or rewrite the Sales Fact.

Build 0.3.0 deliberately applies the current mapping state to historical facts
for analysis. Historical fact-date mapping versions are not defined.

### Legacy CAT/RPG Equivalence Definition
A human-authored ID-to-ID declaration used only for migration comparison. One
definition may contain several legacy ReportingCategoryIDs but must target one
ReportingGroupID. Display-name similarity is never authoritative equivalence.

### Remap Assistant
A legacy 0.2.0 review surface retained during the migration. The Phase 1
hierarchy workflow lives on `Mapping` and persists to `Mapping Rules`.

### KPI Definition
Central metadata describing a metric, its status, domain, formatting, supported views, and presentation behavior.

### Metric Result
A deterministic calculated value derived from the validated Reporting Group
fact bridge for a defined KPI, dataset, Reporting Group, and organizational
scope. Build 0.3.0 Phase 2B stores KPI-0001 results in
`tblMetricRPGResults`. The result records numerator, full-scope denominator,
metric value, mapping date/fingerprint, and calculation timestamp.

Company results also record `RestaurantScopeFingerprint`, derived from the
sorted RestaurantIDs where `Status=Active` and `ReportingEnabled=Yes`. This
configuration defines the shared default Performance company scope without
mutating or removing facts outside the scope. It is reusable scope metadata,
not part of KPI-0001's metric definition.

Company and restaurant results share one contract. Current/comparison is a
presentation choice between independently materialized ImportIDs rather than a
property of a result row. ChannelScope is explicitly All channels for the
validated Phase 2B presentation.

After the weekly Performance cutover, `tblMetricRPGResults` remains the intact
fixed-import regression/rollback layer. Active Performance does not delete,
rewrite, or parse this table.

### Performance Interaction Selection

Phase 2C stores user-facing Include state by stable RestaurantID and
ReportingGroupID. It is presentation configuration, not a fact, Mapping Rule,
metric result, or ReportingEnabled replacement. The selected scope is exactly
the rows marked Yes. When every row is Yes the derived scope is All; otherwise
it is a subset. Selection state is authoritative; All versus Custom is derived,
not user-selected. Existing choices survive refresh by ID, while a newly
eligible ID defaults to No after the catalog already exists.

The interaction layer preserves one additive component contract. The original
fixed-import path derives those components from Phase 2B Restaurant results;
the active weekly path derives them from weekly RPG/scope cache sums. It does
not add rows to `tblMetricRPGResults`; Company rows remain fixed-import controls
for regression and rollback.

The Phase 2C Total and sorting surfaces are derived formula state only. Total
stores selected-RPG current/comparison numerators, one denominator for each
selected period scope, and a numeric selected-display value. Sorting stores a
numeric key and a presentation-order RestaurantID list. Neither surface changes
the canonical component row keyed by RestaurantID or becomes a metric-result
table.

### Weekly Analytical Cache

The Build 0.3.0 weekly cache is versioned derived state, not a replacement
for source reports, published facts, Mapping, or retained Phase 2B results. Its
scope table has one row per CacheVersion, SourcePeriodKey, and RestaurantID and
stores the source denominator plus Mapped, Unmapped, Identity Pending, Conflict,
and Inactive Target additive components. Its RPG table adds ReportingGroupID and
stores only Mapped fact count, Sales NOK, and Quantity. The denominator is not
duplicated per Reporting Group.

Identity Pending facts remain in source scope and outside RPG numerators.
Performance eligibility is shared scope metadata, not fact suppression. A
version manifest records source/preflight/mapping-content/group/scope
fingerprints. MappingAsOfDate and the existing Phase 2A mapping fingerprint are
retained separately as audit metadata. Weekly cache staleness is driven by the
date-neutral MappingContentFingerprint, so a date-only advance does not
invalidate unchanged mapping content. Candidate validation and activation are
separate states. Only one validated `Active` / `Active` version may supply the
weekly Performance path.

### Weekly Performance Period Selection

Current and Compare each store ISO Year, From week, and To week as user-facing
selection state. Generated summaries expose labels such as `2026 W01–W32` and
never expose SourcePeriodKey or CacheVersion. Each requested week must exist in
the active version's period manifest before the affected side is available.

For each selected restaurant and Reporting Group, Performance sums weekly
`MappedSalesNOK` from the RPG cache and sums weekly `SourceSalesNOK` once from
the scope cache. The existing Phase 2C numeric components then calculate share,
PP Change, Current Sales NOK, and NOK Impact. The visible matrix remains an
isolated presentation facade. `tblMetricRPGResults` remains available for
rollback but is not an input to these weekly component formulas.

## Import safety

Pulse should support correction/recovery when a wrong period, partial report, duplicate dataset, or other human error is finalized. Publication is a controlled state, not irreversible destruction of prior data.
