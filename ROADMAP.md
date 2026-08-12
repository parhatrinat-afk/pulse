# Pulse Roadmap

The roadmap is intentionally conservative. New ideas are preserved, but they do not automatically enter the active build.

## Validated foundation — Build 0.2.0-QA

The first deterministic reporting vertical slice has been built and QA-checked:

`Published sales facts → deterministic Category Sales Share → Performance → Reports`

The validated checkpoint supports independent current/comparison dataset selection and does not define In-house as part of the KPI itself.

## Active build — 0.3.0 Mapping + Reporting Groups

Build the semantic layer that sits between source-system structure and business reporting:

`Raw POS data → Source hierarchy → Hierarchical Mapping → Reporting Groups → Metrics / KPIs → Performance`

### Primary goals

- Provide an easy way to browse and understand the POS hierarchy and its underlying products/sales.
- Create business-defined Reporting Groups independent of raw POS categories.
- Map a clean source branch once when safe.
- Inherit mappings to descendants.
- Allow more-specific mappings to override inherited mappings for mixed categories or exceptional products.
- Preserve raw source data and source classifications.
- Make mapping deterministic and explainable.
- Allow Reporting Groups to be activated/deactivated for Performance availability.
- Keep the architecture adaptable to future source systems and hierarchy depths.
- Move Performance toward business Reporting Groups instead of exposing roughly 130 raw POS classifications.

### Build sequencing

1. **Phase 1 — semantic model and hierarchical mapping foundation**
   - Implemented on the development branch: Reporting Group registry,
     hierarchy-aware rules/resolution, browse/action surface, explicit versus
     inherited visibility, and mapping QA.
   - Passed live Excel-for-web migration and behavioral QA on 2026-08-11.

2. **Phase 2A — Reporting Group metric contract and bridge**
   - Implemented and live-Excel validated: immutable fact-to-Effective-Mapping bridge,
     current-state mapping fingerprint, stale-state rejection, explicit mapping
     coverage, Reporting Group aggregation, and human-configured legacy CAT/RPG
     side-by-side comparison.
   - The accepted Lovable migration produces 15,935 Mapped and 2,151 Unmapped
     facts with zero Conflict/Inactive Target rows.

3. **Phase 2B — metric and minimal Performance cutover**
   - Implemented and live-Excel validated on 2026-08-12.
   - Central KPI-0001 results consume the validated Phase 2A Reporting Group
     bridge and are materialized in `_Metric_Calc`.
   - Independent current/comparison datasets, all-channel behavior, company and
     restaurant scopes, and the existing Performance design are preserved.
   - Performance and Reports consume the same centralized metric result.
   - The shared active/ReportingEnabled restaurant scope passed a 16 → 15 → 16
     restaurant round trip with deterministic cardinality and scope fingerprint
     restoration.

4. **Phase 3 — Mapping/Performance UX and release QA**
   - Refine Mapping and reporting usability using real uploaded reports.
   - Complete release reconciliation, regression, and checkpoint QA.
   - Simplify existing pages only when real-world validation supports it.

## Following sales-domain builds

After the Reporting Group foundation is proven:

- Quantity Share.
- Add-ons Attach Rate.
- Percentage-point Change.
- Sales NOK and Quantity views.
- Estimated NOK Impact where the methodology is sufficiently defensible.
- Trend and leaderboard views where appropriate.

## Parked until the sales foundation is proven

- Labour domain: labour cost, worked hours, sick leave, restaurant and department splits.
- Reviews / Google Business Performance adapter.
- Forecasting.
- Cross-domain presentation.
- Configurable Restaurant of the Month scoring.
- AI-assisted observations.
- Automated presentation/export features beyond the validated reporting foundation.

## Real-world validation

Experience changes should be informed by actual use in meetings. Existing pages and controls should not be removed merely for visual simplicity; hide or reconsider first, then remove only when usage evidence supports it.
