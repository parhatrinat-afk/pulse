# Architecture

## Purpose

Pulse separates raw/source data, standardized facts, business classification, deterministic metrics, presentation, and human context.

## Conceptual flow

1. **Source data** — uploaded or API-provided operational data.
2. **Adapter / staging** — source-specific translation into Pulse's neutral structure.
3. **Validation** — period, coverage, duplicates, structure, and import safeguards.
4. **Published facts** — trusted standardized observations.
5. **Classification** — products and source categories mapped to stable reporting concepts.
6. **Metric engine** — deterministic calculations from published facts.
7. **KPI Registry** — defines which metrics are active and how they may be presented.
8. **Performance** — interactive exploration of trusted KPI results.
9. **Reports** — meeting-ready consumption of the same metric results.
10. **Human context** — acknowledgement/explanation without rewriting the underlying facts.

## Layers

### User layer
- Overview
- Performance
- Reports
- Imports
- Settings

### Administration layer
Includes KPI configuration, organization, classification/mapping, exclusions, coverage expectations, and publication controls.

### Engine layer
Includes raw data, staging, facts, audit records, environment/build metadata, and future metric result stores.

## Domain isolation

Pulse may eventually contain Commercial Performance, Labour, Guest Experience, and other domains. A domain is only included in a final product/view when enabled for that context. Missing or parked domains must not implicitly penalize a restaurant.
