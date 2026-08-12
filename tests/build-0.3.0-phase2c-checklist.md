# Build 0.3.0 Phase 2C live Excel checklist

Accepted 2026-08-13 in the canonical
`OneDrive/Pulse/Development/Pulse_Current.xlsx`: 16/16 Interaction QA PASS.
This checklist remains the reproducible live-validation procedure; the accepted
evidence is recorded in `docs/BUILD_0_3_0_PHASE2C_MIGRATION.md`.

- [ ] Run the accepted migration, Phase 1, Phase 2A, and Phase 2B prerequisites.
- [ ] Confirm all Phase 2B Metric Results QA checks are PASS before Phase 2C.
- [ ] Run `Build_0_3_0_Phase2C.ts` successfully in Excel for the web.
- [ ] Confirm interactive dropdown validation is ready 8/8.
- [ ] Confirm all sixteen Phase 2C Interaction QA checks are PASS.
- [ ] Confirm QA-0302C-09 specifically reports PASS and lists exactly PP Change,
      Current Share, Comparison Share, Current Sales NOK, and NOK Impact.
- [ ] Confirm every eligible Restaurant and active RPG defaults to Include=Yes.
      Performance should report all selected, show Total immediately after
      Restaurant, show all nine active RPGs, and show Grand Total. Confirm
      sorting defaults to Total / Highest first.
- [ ] Confirm all-selected Add-ons remains W31 NOK 131,487 / 1.47%, comparison
      1.09%, and approximately +0.38 pp.
- [ ] Confirm an arbitrary three-restaurant Grand Total is summed numerator ÷
      summed denominator, not an average of percentages.
- [ ] Confirm a one-restaurant Grand Total equals its Restaurant row.
- [ ] Confirm deselecting and reselecting Swift updates/restores immediately
      without an Office Script.
- [ ] Confirm Mains and every other RPG can be independently removed/restored by
      changing only its Include value.
- [ ] With exactly one RPG selected, confirm Total equals that RPG in all five
      Display modes.
- [ ] With multiple RPGs selected, confirm Total includes exactly those RPG
      numerators and uses each dataset/scope denominator once.
- [ ] With All RPGs selected, confirm Total share remains below 100% when
      accepted non-RPG sales remain in the denominator.
- [ ] With zero RPGs selected, confirm the established blank matrix behavior.
- [ ] Confirm the separate detail RPG selector works independently of matrix
      multi-selection.
- [ ] In the Norwegian workbook, confirm PP Change shows two comma-decimal
      places (for example `+0,38 pp`, `-0,18 pp`, `0,00 pp`) rather than
      `+000 pp`/`-002 pp`.
- [ ] Confirm Current Share and Comparison Share show two comma-decimal places
      (for example `1,47%`, `5,45%`, `68,50%`) rather than
      `001%`/`004%`/`080%`.
- [ ] Confirm Current Sales NOK retains whole-NOK local thousands grouping and
      the ` NOK` suffix.
- [ ] Confirm NOK Impact equals current selected-scope RPG sales minus comparison
      RPG share × current selected-scope total sales, uses local thousands
      grouping, and displays `+` or `-` when nonzero and `0 NOK` at zero.
- [ ] For All/All Add-ons with W31 2026 current and 2025 comparison, confirm NOK
      Impact is approximately `+34 019 NOK` using the workbook's actual local
      group separator.
- [ ] Swap current/comparison datasets and confirm Add-ons NOK Impact becomes
      approximately `-1 585 928 NOK`; select the same dataset twice and confirm
      `0 NOK`.
- [ ] Confirm a period-length mismatch retains the existing warning but does not
      block NOK Impact.
- [ ] Using a multi-restaurant Custom selection, independently sum the selected
      current/comparison numerator and denominator components and confirm Grand
      Total NOK Impact uses the aggregate comparison share and aggregate current
      sales base rather than SUM of restaurant impacts.
- [ ] Confirm the visible matrix cells are presentation text while the bounded
      selected-display helper and all six component matrices remain numeric.
- [ ] Confirm changing Display mode does not change any underlying current or
      comparison numerator, denominator, or share helper value.
- [ ] Confirm Current Sales NOK Total equals the selected current numerator and
      does not become unavailable solely because its denominator is zero.
- [ ] Confirm PP Change and NOK Impact Total/Grand Total are recomputed from
      aggregate components and never summed from visible RPG/restaurant values.
- [ ] Confirm Sort by Total and each displayed RPG uses the underlying numeric
      value, never the visible text facade.
- [ ] Confirm Highest/Lowest are exact reversals for available values; unavailable
      values remain last and Grand Total never moves.
- [ ] Confirm exact numeric ties resolve deterministically by RestaurantID.
- [ ] Find PP Change and NOK Impact values that display at the same rounded value
      but differ at full precision; confirm sorting follows full precision.
- [ ] Confirm only restaurants marked Include=Yes are sorted.
- [ ] Change Display and then either dataset; confirm restaurant order
      recalculates immediately without an Office Script.
- [ ] Hide the RPG retained in Sort by. Confirm Sort status visibly reports the
      Total fallback, the selector value is preserved, and the RPG sort resumes
      after the RPG is restored.
- [ ] Record representative canonical RestaurantID component/helper values,
      change Sort by and Order, then confirm those values and their canonical
      rows remain identical; only Performance presentation order may change.
- [ ] Confirm denominator-dependent modes display `—` for a zero denominator
      while canonical share remains 0; NOK Impact displays `—` if either the
      current or comparison denominator is zero.
- [ ] Confirm current/comparison swap and same-dataset behavior.
- [ ] Confirm zero-restaurant and zero-RPG selections fail safely.
- [ ] Confirm the visible restaurant/RPG summaries naturally show all selected
      or `n of total selected`, with no separate All/Custom controls.
- [ ] Confirm Sort status, comparison/scope status, and the matrix explanation
      are fully visible with no clipping or overlap.
- [ ] Confirm Reports agrees with Performance and exposes the same restaurant scope.
- [ ] Confirm ReportingEnabled=No restaurants are not exposed after refresh.
- [ ] Confirm facts, Phase 2A bridge, Phase 2B results, mapping, KPI, legacy CAT,
      Performance-scope eligibility, and reconciliation totals remain unchanged.
- [ ] Confirm normal selection changes require Excel recalculation only.
- [ ] Confirm Excel reports no unsupported Map/Set iterator compatibility error.
- [ ] Confirm Excel reports no material read-method-in-loop performance warning.
