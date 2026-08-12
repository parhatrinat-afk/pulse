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

## Guardrail behavior

Pulse may identify factual differences such as:

- different channel
- different period length
- same dataset selected on both sides
- different organizational scope

These are informational warnings.

They do not block the calculation.

## Why

Pulse informs. Humans decide.

An unusual comparison may be intentional and valuable. Pulse should expose the evidence needed to understand the comparison without taking ownership of the business conclusion.
