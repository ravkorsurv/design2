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
| `Incidents` | `IncidentsExpanded` | Populated by the macro with columns `Serial_Number`, `Source_System`, `Incident_Date`, `Failed_Records`, `Pct_Records_Impacted`, `Scenario_Name`, `Missing_Alerts`. Leave headers in place; the macro resizes the table when it runs. |
| `History` | `HistoryRaw` | Surveillance history keyed by system and scenario. Columns: `Source_System`, `Scenario_Name`, `Period_Start`, `Period_End`, `Records_Observed`, `Alerts_Investigated`, `Materiality_Positive`. |
| `Output` | `OutputResults` | Final metrics table. Provide headers matching the macro output (`Serial_Number`, `Source_System`, `Incident_Date`, `Scenario_Name`, `Severity`, `Failed_Records`, `Pct_Records_Impacted`, `History_AlertRate`, `Missing_Alerts`, `Likelihood_Band`, `DQ_Final_Risk`, `Jeffreys_alpha`, `Jeffreys_beta`, `Materiality_Rate_Mean`, `Materiality_Rate_95UCB`, `Expected_Materiality_Mean`, `Expected_Materiality_95UCB`, `P_AtLeast_One_Material_Event_95UCB`, `Run_Timestamp`, `Run_User`, `Workbook_Version`, `Notes`). |
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

1. Import `DQMateriality.bas` into the workbook (VBA Editor → File → Import File...).
2. Populate the configuration sheet and ensure each table uses Excel's
   *Format as Table* so that the `ListObject` names match the list above.
3. Load new incident/history data into the raw tables.
4. Run the `RunDQMateriality` macro. It performs the following steps:
   - Expands the incidents table into one row per impacted scenario, pairing the
     analyst-supplied materiality counts with the scenario list while
     consolidating duplicates.
   - Aggregates history for the configured lookback window per source-system and
     scenario and computes the baseline alert rate and Jeffreys prior parameters.
   - Calculates the severity (from `% of Records Impacted`), likelihood (from the
     potential missed alerts), the DQ materiality risk, and the Poisson
     probability of observing at least one material event at the 95% upper
     confidence bound.
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
