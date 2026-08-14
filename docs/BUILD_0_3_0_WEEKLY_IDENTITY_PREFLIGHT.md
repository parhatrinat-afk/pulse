# Build 0.3.0 — Weekly Identity Preflight

## Scope

This bounded slice turns accepted weekly parser rows into a deterministic,
read-only identity proposal before a compact weekly cache is designed. It does
not mutate `Pulse_Current.xlsx`, publish facts, activate a cache, create mapping
rules, alter Performance, supersede the accepted baseline/W31 imports, or
implement OneDrive/Power Automate intake.

The accepted source remains the 84-week `Sales per Item` corpus under the
human-owned scope contract `SCOPE-030-WEEKLY-SALES-PER-ITEM`. Period and source
identity remain independent of filenames and folders.

## Identity contract

The preflight consumes parsed rows plus a read-only snapshot of the accepted
Pulse catalogs and mapping state.

| Entity | Exact stable key |
|---|---|
| Restaurant | `SourceSystemID + exact SourceRestaurantName` |
| Product | `SourceSystemID + exact Item + exact Sales Account` |
| Source classification | `SourceSystemID + exact Main Category + exact Sub Category` |

Exact source strings are preserved. No case folding, punctuation removal,
trimming, fuzzy matching, filename inference, or folder inference participates
in identity.

Resolution is:

1. one exact current catalog match → reuse its existing stable Pulse ID;
2. no exact current match → propose the next stable ID in deterministic exact-key
   order; or
3. more than one current match, or an unresolved multi-path Product hierarchy →
   `Identity Pending` with source impact for review.

The preflight never changes an existing ProductID, RestaurantID, or
SourceClassificationID. Candidate allocation begins after the current maximum
ID and is reproducible under source-row reordering. Rerunning against a catalog
that already contains accepted candidates reuses them rather than proposing
duplicates.

Identity recognition is not a mapping decision. A new ProductID is Mapped only
when the accepted hierarchy and active mapping rules already resolve it. It
otherwise remains Unmapped. The existing hierarchical resolver remains the one
mapping engine.

## Hierarchy authority

For an existing ProductID, the accepted Pulse Product catalog remains mapping
authority. Weekly Main/Sub paths are lineage and QA evidence only; input row
order never overwrites the Product's current SourceClassificationID.

A new ProductKey observed under exactly one classification receives that exact
classification candidate. A new ProductKey observed under multiple paths still
receives one deterministic ProductID candidate, but its hierarchy is
`Identity Pending` until reviewed.

The nine accepted authority-divergence checks compare only paths already present
in the accepted Pulse classification catalog. Newly observed classification
candidates remain candidate lineage; they do not become alternate mapping
authority merely by appearing in weekly source rows.

## Frozen catalog evidence

`tests/fixtures/build-0.3.0-weekly-identity-catalog.json` is a read-only test
snapshot taken from the accepted canonical workbook state:

- SourceSystemID: `SRC-TEST-SALES`;
- 16 Restaurants;
- 1,041 Products;
- 145 Source Classifications;
- 9 Reporting Groups;
- 133 Mapping Rules;
- Catalog as-of date: 2026-08-12; and
- Mapping fingerprint: `MAP-342029f71a922b47`.

The snapshot is deterministic test evidence, not a second live catalog. Before
any later workbook mutation, the future implementation must reject or refresh
this evidence if the canonical workbook catalogs/mapping fingerprint differ.

## Complete 84-week result

| Measure | Result |
|---|---:|
| Reports | 84 |
| Source facts | 245,632 |
| Source Sales NOK | 484,728,367.25 |
| Source Quantity | 2,469,988.09 |
| Known Restaurants reused | 16 |
| Known ProductKeys reused | 688 |
| Known classifications reused | 142 |
| New Restaurant candidates | 2 (`RST-0017`–`RST-0018`) |
| New Product candidates | 193 (`PRD-001042`–`PRD-001234`) |
| New classification candidates | 19 (`SCL-00146`–`SCL-00164`) |
| Stable-identity facts | 245,512 |
| Identity Pending facts | 120 |
| Duplicate proposed IDs / keys | 0 / 0 |
| Reconciliation | PASS |

The two new Restaurant candidates are intentionally separate:

| RestaurantID | Exact source name | ReportingEnabled | Facts | Sales NOK | Quantity |
|---|---|---:|---:|---:|---:|
| RST-0017 | Test Department (Not for User) | No | 19 | 4,133 | 18 |
| RST-0018 | Test Department (Not for Users) | No | 964 | 344,559 | 1,699.48 |

No existing Restaurant configuration is changed.

The 19 exact classification candidates are:

| ID | Main Category | Sub Category |
|---|---|---|
| SCL-00146 | 17. Mai Set Meny | 17. Mai Set Meny |
| SCL-00147 | 17. May Alkoholfritt | 17. May Alkoholfritt |
| SCL-00148 | 17. May Cocktails | 17. May Cocktails |
| SCL-00149 | 17. May Food's | 17. May Combo's |
| SCL-00150 | 17. May Food's | 17. May Dessert |
| SCL-00151 | 17. May Food's | 17. May Hot Food |
| SCL-00152 | 17. May Food's | 17. May Starters |
| SCL-00153 | 17. May Wines | 17. May Wines |
| SCL-00154 | 17. May Øl & Cider | 17. May Øl & Cider |
| SCL-00155 | GiftCard | GiftCard |
| SCL-00156 | New Classic Cocktails | New Classic Cocktails |
| SCL-00157 | Signature Cocktails | Signature Cocktails |
| SCL-00158 | Special deal KJ | Beer |
| SCL-00159 | Special deal KJ | Drinks |
| SCL-00160 | Special deal KJ | Wine |
| SCL-00161 | Summer Drinks | Summer Drinks |
| SCL-00162 | Summer Menu | Summer Menu |
| SCL-00163 | White Wine | White Wine |
| SCL-00164 | Wine & Sake | Wine & Sake |

## Identity Pending review

These six new exact ProductKeys receive stable candidate ProductIDs, but their
canonical hierarchy path requires a human choice. Until then their 120 facts
remain explicit `Identity Pending`; none disappears or is silently mapped.

| ProductID | Exact Item / Sales Account | Observed paths | Facts | Sales NOK | Quantity |
|---|---|---|---:|---:|---:|
| PRD-001102 | Fuelbox Teams - English / 3900 - Salg Effekter 25% | Bag › Bag; Merch › Fuelbox | 2 | 1,299 | 1 |
| PRD-001125 | Hot Food Combo / 3000 - Salg Mat 25% | Overtime Food B2B › Overtime Food B2B; Solheimsviken Special › Solheimsviken Special | 3 | 597 | 3 |
| PRD-001138 | Korean Fried Chicken / 3000 - Salg Mat 25% | Snacks › Snacks; Woks › Woks | 59 | 111,052 | 856 |
| PRD-001187 | Staff Coca Cola / 3040 - Salg Brus 25% | Staff Drinks Swift › Staff Drinks Swift; Staff Drinks › Staff Softdrinks | 5 | 123 | 5 |
| PRD-001211 | Sumo Reusable Bag / 3000 - Salg Mat 15% | Doggy Bag › Doggy Bag; Packaging › Packaging | 45 | 810 | 81 |
| PRD-001220 | Tempura Poke Bowl / 3000 - Salg Mat 25% | Overtime Food B2B › Overtime Food B2B; Solheimsviken Special › Solheimsviken Special | 6 | 995 | 5 |

## Existing Product hierarchy review

These are review evidence, not automatic changes. Current Pulse hierarchy and
mapping remain authoritative unless a human explicitly decides otherwise.

| ProductID | Product | Current outcome | Alternate observed outcome | Impact of alternate path |
|---|---|---|---|---:|
| **PRD-000689** | **Red Curry** | **RPG-0001 Add-ons** via Sauces › Sauces | **RPG-0009 Mains** via Mains › Hot Food | **4 facts / NOK 2,233 / Qty 7** |
| PRD-000027 | Aperol Spritz | RPG-0003 Spirits/Cocktails | Unmapped via Campaign Drinks* | 4 / 714 / 6 |
| PRD-000104 | Big Family Combo | RPG-0009 Mains | Unmapped via Vinterferie tilbud | 17 / 133,089 / 110 |
| PRD-000233 | Crispy Cantonese Chicken Poké | RPG-0009 Mains | Unmapped via Overtime Food B2B | 1 / 199 / 1 |
| PRD-000296 | Dirty Tempura | RPG-0009 Mains | Unmapped via 17. May Food's › 17. May Sushi | 11 / 139,073 / 484 |
| PRD-000365 | Family Combo | RPG-0009 Mains | Unmapped via Vinterferie tilbud | 19 / 246,421 / 328.67 |
| PRD-000449 | Gyoza Dumplings | RPG-0009 Mains | Unmapped via Start Line* | 2 / 627 / 3 |
| PRD-000596 | NYC Orange Chicken | RPG-0009 Mains | Unmapped via Vinterferie tilbud | 19 / 154,134 / 508.9 |
| PRD-000870 | Stet ris med kylling | RPG-0009 Mains | Unmapped via Overtime Food B2B | 3 / 1,592 / 8 |

## Mapping-state evidence

The current accepted mapping state applied to stable weekly identities produces:

| State | Facts | Sales NOK | Quantity |
|---|---:|---:|---:|
| Mapped | 226,485 | 474,160,541.60 | 2,340,106.61 |
| Unmapped | 19,027 | 10,452,949.65 | 128,930.48 |
| Conflict | 0 | 0 | 0 |
| Inactive Target | 0 | 0 | 0 |
| Identity Pending | 120 | 114,876.00 | 951.00 |

All five states reconcile exactly to 245,632 facts, NOK 484,728,367.25, and
Quantity 2,469,988.09. Unmapped and Identity Pending remain visible; neither is
removed from denominators or source totals.

Deterministic fingerprints are:

- source corpus `WSC-349b8bfd096ace2e`;
- accepted catalog `IDC-56484a052b6a4cab`;
- candidates `IDN-4803e1706372df44`;
- review `IDR-5eb0d5643dc88499`; and
- complete preflight `IDP-4cd1159238339096`.

## Running the read-only preflight

The command requires exact caller-supplied corpus and catalog paths:

```text
node src/imports/audit-weekly-identity-preflight.mjs <exact-read-only-corpus-path> \
  --catalog tests/fixtures/build-0.3.0-weekly-identity-catalog.json \
  --expected tests/expected-build-0.3.0-weekly-identity.json
```

It reads source files and emits JSON only. It never searches for a corpus,
writes source reports, or opens/mutates the canonical workbook.

## Live Excel boundary

No live Excel run is required for this slice because it adds no Office Script
and no workbook mutation. Before a later candidate-catalog/cache mutation, use
the connected canonical `Pulse_Current.xlsx` to confirm that its catalog maxima,
mapping rules, and mapping fingerprint still match the frozen snapshot. A
difference must cause a refresh/review, not silent use of stale identity evidence.
