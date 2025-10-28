# Excel/VBA implementation of the DQ/materiality workflow

This folder contains a VBA module (`DQMateriality.bas`) that operationalises the
compliance surveillance materiality scoring logic using only Excel-native
functionality. Import the module into a macro-enabled workbook (`.xlsm`) and
create the tables/named ranges listed below to mirror the Python application's
behaviour.

## Workbook layout

| Sheet  | Table / named range | Purpose |
| ------ | ------------------- | ------- |
| `Incidents` | `IncidentsRaw` | Raw incident feed imported from the surveillance CSV. Required columns: `Serial Number`, `Date`, `Source System`, `Asset Class`, `Transaction Type`, `Total no. of Records Received`, `Count of Unrequired Records`, `Total Count of Required  Records`, `Count of Failed Records`, `% of Records Impacted`, `Failed KDE Name`, `DMOP Data  Quality Filter That Failed`, `Scenarios Impacted`, `Potential missing alerts per scenario`, plus any additional descriptive fields provided by the analysts. |
| `Incidents` | `IncidentsExpanded` | Populated by the macro with columns `Serial_Number`, `Source_System`, `Asset_Class`, `Incident_Date`, `Failed_Records`, `Pct_Records_Impacted`, `Scenario_Name`, `Model_Family`, `Missing_Alerts`. Leave headers in place; the macro resizes the table when it runs. |
| `History` | `HistoryRaw` | Surveillance history keyed by system and scenario. Columns: `Source_System`, `Scenario_Name`, `Period_Start`, `Period_End`, `Records_Observed`, `Alerts_Investigated`, `Materiality_Positive`. |
| `Output` | `OutputResults` | Final metrics table. Provide headers matching the macro output (`Serial_Number`, `Source_System`, `Asset_Class`, `Incident_Date`, `Scenario_Name`, `Model_Family`, `Severity`, `Failed_Records`, `Pct_Records_Impacted`, `History_AlertRate`, `Missing_Alerts`, `Likelihood_Band`, `DQ_Final_Risk`, `Materiality_Ratio`, `Materiality_Score`, `Reserved_1`, `Reserved_2`, `Reserved_3`, `Reserved_4`, `Reserved_5`, `Run_Timestamp`, `Run_User`, `Workbook_Version`, `Notes`). |
| `Audit` | `AuditLog` | Tracks every macro run with `Run_Timestamp`, `Run_User`, `Row_Count`, `Digest`. |
| `Config` | `SeverityThresholds` | Ordered table of severity bands. Columns: `MinPct`, `Severity` (e.g., 0 → Low, 25 → Medium, 50 → High). |
| `Config` | `LikelihoodThresholds` | Ordered table of likelihood bands with columns `MinImpact`, `Band`. |
| `Config` | `DQMatrix` | DQ risk matrix. Column 1 = `Severity`, remaining columns named for likelihood bands containing the final risk labels. |
| `Config` | `ScenarioModelFamilies` | Map each normalised scenario to a model family. Columns: `Scenario_Name`, `Model_Family`. |
| `Config` | `MaterialityRatios` | Long-term materiality ratios per asset class. Column 1 contains the asset class label (use `Default`/blank for the fallback row) and subsequent columns hold ratios keyed by the DQ risk labels (e.g. `High`, `Medium`, `Low`). |
| `Config` | `MaterialCategories` | Defines the SWAT output categories to display in the summary and whether each counts as STOR-related. Columns: `Category` (or `Category_Key`), optional `Display_Label`, and `Is_STOR_Related`. |
| `Config` | `Config_LookbackDays` | Named cell containing the history lookback window in days. |
| `Config` | `Config_RunUser` | Named cell holding the analyst's name/ID. |
| `Config` | `Config_WorkbookVersion` | Named cell for the workbook version string. |
| `Metrics` | `MaterialOutputsRaw` | Optional SWAT material-output counts by asset class. Columns: `Asset_Class`, `Category` (matching the config table), optional `Category_Label`, and `Count`. |
| `Metrics` | `STSAlertsRaw` | Optional STS alert counts by asset class and escalation. Columns: `Asset_Class`, `Escalation_Level`, `Alert_Count`. |
| `Metrics` | `AlertSummary` | Destination table for the cross-asset summary produced by `RefreshAlertSummary`. The macro overwrites all data rows and will resize the table to match the asset classes present. |

The macro relies on Excel's built-in `VBScript.RegExp` and `Scripting.Dictionary`
objects, both of which are present on Windows without additional installation.

## Running the macro

1. Import `DQMateriality.bas` into the workbook (VBA Editor → File → Import File...).
2. Populate the configuration sheet and ensure each table uses Excel's
   *Format as Table* so that the `ListObject` names match the list above.
3. Load new incident/history data into the raw tables.
4. Run the `RunDQMateriality` macro. It performs the following steps:
   - Expands the incidents table into one row per impacted scenario, pairing the
     analyst-supplied materiality counts with the scenario list while
     consolidating duplicates.
   - Uses the `ScenarioModelFamilies` table to enrich each scenario with its
     model family before writing to `IncidentsExpanded`.
   - Aggregates history for the configured lookback window per source-system and
     scenario and computes the baseline alert rate.
   - Calculates the severity (from `% of Records Impacted`), likelihood (from the
     potential missed alerts), resolves the DQ materiality risk, and applies the
     asset-class ratio from `MaterialityRatios` to derive the materiality score.
   - Writes the results into `OutputResults` and appends an audit entry with a
     SHA-256 digest of the written rows (falls back to a checksum if SHA-256 is
     unavailable on the machine).
   - If the optional metrics tables are present, refreshes the cross-asset
     summary in `AlertSummary`.

The workbook contains no external dependencies; all numeric routines fall back
on VBA implementations if the corresponding Excel worksheet functions are not
available (e.g., `BETA.INV`).

## Additional helpers

- `ExpandIncidents` can be run independently to refresh only the expanded
  incident table.
- The audit digest function attempts to use the Windows cryptography provider.
  On locked-down machines it degrades gracefully to a lightweight checksum while
  still recording an audit trail.
- `RefreshAlertSummary` can be executed on demand to rebuild the summary without
  rerunning the full pipeline. The macro produces three sections:
  - **Material Outputs (SWAT Data)** – one row per configured category plus
    totals for all material outputs and the subset flagged as STOR-related.
  - **STS Closure Bucket – Escalation Level** – aggregation of the alert table
    by escalation level, including a total row.
  - **Ratios** – the material-output and STOR ratios expressed as percentages of
    total alerts per asset class. Ratios use the totals from the preceding
    sections, so blank or zero alert counts yield zero ratios.
