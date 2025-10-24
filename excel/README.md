# Excel/VBA implementation of the DQ/STOR workflow

This folder contains a VBA module (`DQSTOR.bas`) that re-implements the Python
`dqstor` pipeline using only Excel-native functionality. Import the module into
a macro-enabled workbook (`.xlsm`) and create the tables/named ranges listed
below to mirror the Python application's behaviour.

## Workbook layout

| Sheet  | Table / named range | Purpose |
| ------ | ------------------- | ------- |
| `Incidents` | `IncidentsRaw` | Raw weekly incident feed with columns `Incident_ID`, `Incident_Date`, `Model_Scope`, `Records_Impacted`, `Pct_Volume_Impacted`, `Alert_Impacted` (additional descriptive columns are ignored). |
| `Incidents` | `IncidentsExpanded` | Output table populated by the macro with columns `Incident_ID`, `Model_Scope`, `Incident_Date`, `Records_Impacted`, `Pct_Volume_Impacted`, `Alert_Impact`. Leave headers in place; the macro resizes the table when it runs. |
| `History` | `HistoryRaw` | Surveillance history. Columns: `Model_Scope`, `Period_Start`, `Period_End`, `Records_Observed`, `Alerts_Investigated`, `STORs_Filed`. |
| `Output` | `OutputResults` | Final metrics table. Provide headers matching the Python output (`Incident_ID`, `Model_Scope`, `Incident_Date`, `Severity`, `Records_Impacted`, `Baseline_AlertRate`, `Missed_Alerts`, `Likelihood_Band`, `DQ_Final_Risk`, `Jeffreys_alpha`, `Jeffreys_beta`, `STOR_Rate_Mean`, `STOR_Rate_95UCB`, `Expected_Missed_STORs_Mean`, `Expected_Missed_STORs_95UCB`, `P_AtLeast_One_Missed_STOR_95UCB`, `Run_Timestamp`, `Run_User`, `Workbook_Version`, `Notes`). |
| `Audit` | `AuditLog` | Tracks every macro run with `Run_Timestamp`, `Run_User`, `Row_Count`, `Digest`. |
| `Config` | `SeverityThresholds` | Ordered table of severity bands. Columns: `MinPct`, `Severity` (e.g., 0 → Low, 25 → Medium, 50 → High). |
| `Config` | `LikelihoodThresholds` | Ordered table of likelihood bands with columns `MinImpact`, `Band`. |
| `Config` | `DQMatrix` | DQ risk matrix. Column 1 = `Severity`, remaining columns named for likelihood bands containing the final risk labels. |
| `Config` | `Config_LookbackDays` | Named cell containing the history lookback window in days. |
| `Config` | `Config_RunUser` | Named cell holding the analyst's name/ID. |
| `Config` | `Config_WorkbookVersion` | Named cell for the workbook version string. |

The macro relies on Excel's built-in `VBScript.RegExp` and `Scripting.Dictionary`
objects, both of which are present on Windows without additional installation.

## Running the macro

1. Import `DQSTOR.bas` into the workbook (VBA Editor → File → Import File...).
2. Populate the configuration sheet and ensure each table uses Excel's
   *Format as Table* so that the `ListObject` names match the list above.
3. Load new incident/history data into the raw tables.
4. Run the `RunDQSTOR` macro. It performs the following steps:
   - Expands the incidents table into one row per alert impact using the
     `ParseAlertImpacts` helper.
   - Aggregates history for the configured lookback window and computes
     baseline alert/STOR rates.
   - Calculates missed alerts, DQ classifications, Jeffreys posterior metrics,
     expected missed STORs, and the Poisson probability of missing at least one
     STOR.
   - Writes the results into `OutputResults` and appends an audit entry with a
     SHA-256 digest of the written rows (falls back to a checksum if SHA-256 is
     unavailable on the machine).

The workbook contains no external dependencies; all numeric routines fall back
on VBA implementations if the corresponding Excel worksheet functions are not
available (e.g., `BETA.INV`).

## Additional helpers

- `ExpandIncidents` can be run independently to refresh only the expanded
  incident table.
- The audit digest function attempts to use the Windows cryptography provider.
  On locked-down machines it degrades gracefully to a lightweight checksum while
  still recording an audit trail.

