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

### Reporting Category
A Pulse-owned business classification used for analysis. It does not have to match the POS/source category.

### Source Classification
The category or classification supplied by the originating system.

### Remap Rule
A durable mapping rule translating a recognized source/product condition into the intended Pulse reporting classification.

### Remap Assistant
A review surface for items that cannot safely be classified using existing rules.

### KPI Definition
Central metadata describing a metric, its status, domain, formatting, supported views, and presentation behavior.

### Metric Result
A deterministic calculated value derived from published facts for a defined scope, period, and KPI.

## Import safety

Pulse should support correction/recovery when a wrong period, partial report, duplicate dataset, or other human error is finalized. Publication is a controlled state, not irreversible destruction of prior data.
