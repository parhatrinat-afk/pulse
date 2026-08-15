# Comparison Model

## Principle

Current and comparison are two independent analysis scopes.

Pulse does not hard-code 2025, prior year, prior week, or any other dataset as the universal baseline.

## Model

A comparison consists of:

### Current
- dataset or period
- channel
- organizational scope
- KPI/Reporting Group

### Compare with
- dataset or period
- channel
- organizational scope
- same KPI/Reporting Group

The KPI definition remains constant. Only the scope changes.

## Performance restaurant scope contract

Unless an approved KPI definition explicitly says otherwise, organizational
scope uses the shared Performance restaurant set where `Status=Active` and
`ReportingEnabled=Yes`. This scope is independent of metric identity and must
be applied consistently to current and comparison results.

Phase 2C lets the user interactively choose any subset inside that eligible set.
The same selected RestaurantIDs apply to both the current and comparison
components. After the weekly cutover, Current and Compare independently select
ISO Year, From week, and To week from the single fresh active weekly cache.
Changing restaurant/RPG selection or either complete week range is Excel
recalculation, not a metric refresh.
The Performance Total column aggregates the currently selected Reporting
Groups. PP Change and NOK Impact are recalculated from aggregate current and
comparison components; they are not sums of displayed row percentages or
impacts. Restaurant sorting is presentation-only and uses the same numeric
selected-display helpers, so changing either dataset may change rank without
changing component identity.

## Guardrail behavior

Pulse may identify factual differences such as:

- different channel
- different period length
- the same complete period selected on both sides
- different organizational scope

These are informational warnings and do not block calculation. Identical
periods intentionally return zero PP Change and NOK Impact.

A requested range that is reversed, unavailable, or missing any required week
is not merely an unusual comparison: the affected side is invalid/incomplete
and its calculation is blocked rather than silently using a partial period.

## Why

Pulse informs. Humans decide.

An unusual comparison may be intentional and valuable. Pulse should expose the evidence needed to understand the comparison without taking ownership of the business conclusion.
