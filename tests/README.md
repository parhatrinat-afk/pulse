# tests

This directory contains deterministic validation and regression assets for Pulse.

Current assets include Build 0.2.0 validation material such as the build checklist and expected output data. Future build work should extend these tests rather than treating this directory as a placeholder.

For Build 0.3.0, tests should cover at minimum:

- Reporting Group registry integrity.
- Hierarchical mapping resolution.
- More-specific override precedence.
- Inherited mapping behavior.
- Unmapped handling.
- Active/inactive Reporting Group behavior.
- Reconciliation of mapped aggregates to immutable source sales facts.
- Regression protection for the validated 0.2.0 metric behavior where applicable.

Tests should be deterministic, explainable, and tied to a reproducible build or workbook checkpoint.
