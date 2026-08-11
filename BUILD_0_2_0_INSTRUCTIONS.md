# Pulse Build 0.2.0 implementation package

This package changes the **implementation method**, not the Pulse product direction.

The current workbook checkpoint remains **Pulse Build 0.1.1**. Instead of ChatGPT repeatedly rewriting the growing `.xlsx` file, the repository now owns a reproducible Office Script that applies the next build inside Excel.

## Build 0.2.0

The first implemented KPI is **Category Sales Share**:

`Category Sales Share = category sales amount / total sales amount inside the selected scope`

The KPI is **not restricted to In-house**.

The user independently selects:

- current dataset,
- current channel,
- comparison dataset,
- comparison channel,
- reporting category.

Available channels are read from the published facts. `All channels` is also available.

A dataset is not permanently defined as a baseline. It becomes the comparison baseline only when the user selects it under **Compare with**.

Pulse may show that current and comparison scopes differ. It does not block intentional comparisons.

## How to generate Pulse Build 0.2.0

1. Open your current `excel/Pulse_Build_0_1_1.xlsx` in Microsoft Excel for the web.
2. Open **Automate**.
3. Create a new Office Script.
4. Replace the default code with `office-scripts/Build_0_2_0.ts`.
5. Name it `Pulse - Build 0.2.0`.
6. Run it once.
7. Review **Performance** and **Reports**.
8. Save a copy as `Pulse_Build_0_2_0.xlsx`.
9. Put that workbook in `excel/`.
10. Commit the script, documentation, tests, and generated workbook to GitHub.

Do not overwrite 0.1.1. It remains the rollback checkpoint.

## Expected current-data sanity checks

With the data currently stored in Build 0.1.1, selecting `Add-ons` gives approximately:

| Dataset | Channel | Category Sales Share |
|---|---|---:|
| Week 31, 2026 | In-house | 0.1686% |
| 2025 full year | In-house | 0.1169% |
| Week 31, 2026 | Takeaway | 0.0877% |
| 2025 full year | Takeaway | 0.1542% |

These values are only regression checks for the current test data. They are not hard-coded into Pulse.
