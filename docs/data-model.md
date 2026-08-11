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
Reporting Group ID. Current supported scopes are SourceMainCategory,
SourceSubCategory, and Product. Effective intervals are inclusive; overlapping
active rules on the same node/scope are conflicts.

### Effective Mapping
One computed row per product showing applicable ancestor/product rules, the
winning rule, Reporting Group, explicit versus inherited state, and status.
Resolution is most-specific explicit mapping, otherwise nearest mapped ancestor,
otherwise Unmapped.

### Remap Assistant
A legacy 0.2.0 review surface retained during the migration. The Phase 1
hierarchy workflow lives on `Mapping` and persists to `Mapping Rules`.

### KPI Definition
Central metadata describing a metric, its status, domain, formatting, supported views, and presentation behavior.

### Metric Result
A deterministic calculated value derived from published facts for a defined scope, period, and KPI.

## Import safety

Pulse should support correction/recovery when a wrong period, partial report, duplicate dataset, or other human error is finalized. Publication is a controlled state, not irreversible destruction of prior data.
