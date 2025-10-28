Attribute VB_Name = "DQMateriality"
Option Explicit

' VBA workflow for assessing surveillance data quality/materiality using the
' incident feed provided by Compliance Surveillance.  The module assumes the
' workbook contains the following tables/named ranges:
'   - Sheet "Incidents" with tables:
'       * IncidentsRaw: raw incident feed including the scenario/materiality text
'       * IncidentsExpanded: destination table populated by ExpandIncidents
'   - Sheet "History" with table HistoryRaw storing historical surveillance metrics
'   - Sheet "Output" with table OutputResults for the computed risk scores
'   - Sheet "Audit" with table AuditLog to track each run
'   - Sheet "Config" with helper tables:
'       * SeverityThresholds (columns: MinPct, Severity, optional Description)
'       * LikelihoodThresholds (columns: MinImpact, Band)
'       * DQMatrix (first column Severity, subsequent columns named for bands)
'   - Named ranges on Config sheet:
'       * Config_LookbackDays (number of days for history window)
'       * Config_RunUser (current analyst/user string)
'       * Config_WorkbookVersion (text identifier)
'
' The workbook only relies on Excel/VBA features (no external add-ins).  All
' calculations are orchestrated via the RunDQMateriality macro below.

Private Const SHEET_INCIDENTS As String = "Incidents"
Private Const SHEET_HISTORY As String = "History"
Private Const SHEET_OUTPUT As String = "Output"
Private Const SHEET_AUDIT As String = "Audit"
Private Const SHEET_CONFIG As String = "Config"
Private Const SHEET_METRICS As String = "Metrics"

Private Const TABLE_INCIDENTS_RAW As String = "IncidentsRaw"
Private Const TABLE_INCIDENTS_EXPANDED As String = "IncidentsExpanded"
Private Const TABLE_HISTORY_RAW As String = "HistoryRaw"
Private Const TABLE_OUTPUT As String = "OutputResults"
Private Const TABLE_AUDIT As String = "AuditLog"
Private Const TABLE_SEVERITY As String = "SeverityThresholds"
Private Const TABLE_LIKELIHOOD As String = "LikelihoodThresholds"
Private Const TABLE_DQMATRIX As String = "DQMatrix"
Private Const TABLE_MATERIALITY_RATIOS As String = "MaterialityRatios"
Private Const TABLE_SCENARIO_FAMILY As String = "ScenarioModelFamilies"

Private Type IncidentColumnIndexes
    SourceSystem As Long
    AssetClass As Long
    ScenarioName As Long
    ModelFamily As Long
    FailedRecords As Long
    PercentImpacted As Long
    MissingAlerts As Long
    IncidentDate As Long
    SerialNumber As Long
End Type

Private Function BuildIncidentColumnIndexes(ByVal headerMap As Object) As IncidentColumnIndexes
    Dim indexes As IncidentColumnIndexes

    With indexes
        .SourceSystem = headerMap("Source_System")
        .AssetClass = headerMap("Asset_Class")
        .ScenarioName = headerMap("Scenario_Name")
        .ModelFamily = headerMap("Model_Family")
        .FailedRecords = headerMap("Failed_Records")
        .PercentImpacted = headerMap("Pct_Records_Impacted")
        .MissingAlerts = headerMap("Missing_Alerts")
        .IncidentDate = headerMap("Incident_Date")
        .SerialNumber = headerMap("Serial_Number")
    End With

    BuildIncidentColumnIndexes = indexes
End Function
Private Const TABLE_MATERIAL_CATEGORIES As String = "MaterialCategories"
Private Const TABLE_MATERIAL_OUTPUTS As String = "MaterialOutputsRaw"
Private Const TABLE_STS_ALERTS As String = "STSAlertsRaw"
Private Const TABLE_ALERT_SUMMARY As String = "AlertSummary"

Public Sub RunDQMateriality()
    On Error GoTo HandleError

    Dim runTimestamp As Date
    runTimestamp = Now

    ExpandIncidents

    WriteHistoryFromIncidents

    Dim rollup As Object
    Set rollup = BuildHistoryRollup

    Dim outputRows As Variant
    outputRows = ComputeOutputRows(rollup, runTimestamp)

    WriteOutput outputRows
    AppendAuditEntry outputRows, runTimestamp

    RefreshAlertSummary False

    MsgBox "DQ/materiality calculations complete", vbInformation
    Exit Sub

HandleError:
    Dim errorMessage As String
    errorMessage = Err.Description

    On Error Resume Next
    AppendAuditEntry Empty, runTimestamp, errorMessage

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim historyTable As ListObject
    Set historyTable = wb.Worksheets(SHEET_HISTORY).ListObjects(TABLE_HISTORY_RAW)
    ClearTable historyTable

    Dim outputTable As ListObject
    Set outputTable = wb.Worksheets(SHEET_OUTPUT).ListObjects(TABLE_OUTPUT)
    ClearTable outputTable
    On Error GoTo 0

    MsgBox "DQ/materiality calculations failed. " & _
           "See AuditLog entry for " & Format$(runTimestamp, "yyyy-mm-dd hh:nn:ss") & "." & vbCrLf & _
           "Details: " & errorMessage, vbCritical
End Sub

Public Sub ExpandIncidents()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim srcTable As ListObject
    Set srcTable = wb.Worksheets(SHEET_INCIDENTS).ListObjects(TABLE_INCIDENTS_RAW)

    Dim dstTable As ListObject
    Set dstTable = wb.Worksheets(SHEET_INCIDENTS).ListObjects(TABLE_INCIDENTS_EXPANDED)

    ClearTable dstTable

    If srcTable.ListRows.Count = 0 Then Exit Sub

    Dim data As Variant
    data = srcTable.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(srcTable)

    Dim assetIndex As Long
    assetIndex = ResolveHeaderIndex(headerMap, "Asset Class", "Asset_Class")

    Dim familyMap As Object
    Set familyMap = LoadScenarioFamilies()

    Dim outRows As Collection
    Set outRows = New Collection

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim incidentId As String
        incidentId = NzString(data(r, headerMap("Serial Number")))

        Dim incidentDate As Date
        incidentDate = NzDate(data(r, headerMap("Date")))

        Dim sourceSystem As String
        sourceSystem = NzString(data(r, headerMap("Source System")))

        Dim assetClass As String
        If assetIndex > 0 Then
            assetClass = NzString(data(r, assetIndex))
        Else
            assetClass = ""
        End If

        Dim failedRecords As Double
        failedRecords = NzDouble(data(r, headerMap("Count of Failed Records")))

        Dim percentImpacted As Double
        percentImpacted = NzDouble(data(r, headerMap("% of Records Impacted")))

        Dim scenariosText As String
        scenariosText = NzString(data(r, headerMap("Scenarios Impacted")))

        Dim potentialText As String
        potentialText = NzString(data(r, headerMap("Potential missing alerts per scenario")))

        Dim scenarioImpacts As Collection
        Set scenarioImpacts = ParseScenarioMateriality(scenariosText, potentialText, sourceSystem)

        Dim impactItem As Variant
        For Each impactItem In scenarioImpacts
            Dim rowValues(1 To 9) As Variant
            rowValues(1) = incidentId
            rowValues(2) = sourceSystem
            rowValues(3) = assetClass
            rowValues(4) = incidentDate
            rowValues(5) = failedRecords
            rowValues(6) = percentImpacted
            rowValues(7) = impactItem("Scenario")
            rowValues(8) = LookupModelFamily(impactItem("Scenario"), familyMap)
            rowValues(9) = impactItem("MissingAlerts")
            outRows.Add rowValues
        Next impactItem
    Next r

    If outRows.Count = 0 Then Exit Sub

    Dim outputData() As Variant
    ReDim outputData(1 To outRows.Count, 1 To 9)

    Dim i As Long, c As Long
    For i = 1 To outRows.Count
        Dim values() As Variant
        values = outRows(i)
        For c = 1 To 9
            outputData(i, c) = values(c)
        Next c
    Next i

    dstTable.Resize dstTable.Range.Resize(RowSize:=outRows.Count + 1)
    dstTable.DataBodyRange.Value = outputData
End Sub

Private Function LoadScenarioFamilies() As Object
    Dim dict As Object
    Set dict = NewDictionary()

    Dim tbl As ListObject
    Set tbl = GetTableIfExists(SHEET_CONFIG, TABLE_SCENARIO_FAMILY)
    If tbl Is Nothing Then
        Set LoadScenarioFamilies = dict
        Exit Function
    End If

    If tbl.ListRows.Count = 0 Then
        Set LoadScenarioFamilies = dict
        Exit Function
    End If

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(tbl)

    Dim scenarioIndex As Long
    scenarioIndex = ResolveHeaderIndex(headerMap, "Scenario_Name", "Scenario")
    If scenarioIndex = 0 Then
        Set LoadScenarioFamilies = dict
        Exit Function
    End If

    Dim familyIndex As Long
    familyIndex = ResolveHeaderIndex(headerMap, "Model_Family", "Family", "Model Family")
    If familyIndex = 0 Then
        Set LoadScenarioFamilies = dict
        Exit Function
    End If

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim scenarioName As String
        scenarioName = NormalizeScenarioName(data(r, scenarioIndex))
        If scenarioName <> "" Then
            Dim familyName As String
            familyName = NzString(data(r, familyIndex))
            If familyName <> "" Then
                dict(scenarioName) = familyName
            End If
        End If
    Next r

    Set LoadScenarioFamilies = dict
End Function

Private Function LookupModelFamily(ByVal scenarioName As String, ByVal familyMap As Object) As String
    Dim normalized As String
    normalized = NormalizeScenarioName(scenarioName)
    If normalized = "" Then
        LookupModelFamily = "Unspecified"
        Exit Function
    End If

    If Not familyMap Is Nothing Then
        If familyMap.Exists(normalized) Then
            LookupModelFamily = CStr(familyMap(normalized))
            Exit Function
        End If
    End If

    LookupModelFamily = normalized
End Function

Private Function ParseScenarioMateriality(ByVal scenariosText As String, _
                                          ByVal potentialText As String, _
                                          ByVal fallbackScenario As String) As Collection
    Dim impacts As New Collection

    Dim scenarioSet As Object
    Set scenarioSet = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    scenarioSet.CompareMode = vbTextCompare
    On Error GoTo 0

    Dim potentialDict As Object
    Set potentialDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    potentialDict.CompareMode = vbTextCompare
    On Error GoTo 0

    Dim rx As Object
    Set rx = GetScenarioRegex()

    Dim matches As Object
    Set matches = rx.Execute(potentialText)

    Dim i As Long
    For i = 0 To matches.Count - 1
        Dim match As Object
        Set match = matches(i)

        Dim scenarioName As String
        scenarioName = NormalizeScenarioName(match.SubMatches(0))

        Dim materialityValue As Double
        materialityValue = NzDouble(match.SubMatches(1))

        If potentialDict.Exists(scenarioName) Then
            potentialDict(scenarioName) = potentialDict(scenarioName) + materialityValue
        Else
            potentialDict.Add scenarioName, materialityValue
        End If

        scenarioSet(scenarioName) = True
    Next i

    Dim rawScenarios As Variant
    rawScenarios = SplitScenarios(scenariosText)

    If Not IsEmpty(rawScenarios) Then
        For i = LBound(rawScenarios) To UBound(rawScenarios)
            Dim candidate As String
            candidate = NormalizeScenarioName(rawScenarios(i))
            If candidate <> "" Then
                scenarioSet(candidate) = True
            End If
        Next i
    End If

    If scenarioSet.Count = 0 Then
        Dim defaultScenario As String
        defaultScenario = NormalizeScenarioName(fallbackScenario)
        If defaultScenario = "" Then defaultScenario = "Unspecified"
        scenarioSet(defaultScenario) = True
    End If

    Dim scenarioKey As Variant
    For Each scenarioKey In scenarioSet.Keys
        Dim entry As Object
        Set entry = CreateObject("Scripting.Dictionary")
        entry.Add "Scenario", scenarioKey
        If potentialDict.Exists(scenarioKey) Then
            entry.Add "MissingAlerts", potentialDict(scenarioKey)
        Else
            entry.Add "MissingAlerts", 0#
        End If
        impacts.Add entry
    Next scenarioKey

    Set ParseScenarioMateriality = impacts
End Function

Private Function GetScenarioRegex() As Object
    Static scenarioRegex As Object

    If scenarioRegex Is Nothing Then
        Set scenarioRegex = CreateObject("VBScript.RegExp")
    End If

    scenarioRegex.Global = True
    scenarioRegex.Pattern = "([^\(]+?)\s*\(([^\)]+)\)"

    Set GetScenarioRegex = scenarioRegex
End Function

Private Function BuildHistoryRollup() As Object
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim tbl As ListObject
    Set tbl = wb.Worksheets(SHEET_INCIDENTS).ListObjects(TABLE_INCIDENTS_EXPANDED)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    dict.CompareMode = vbTextCompare
    On Error GoTo 0

    If tbl.ListRows.Count = 0 Then
        Set BuildHistoryRollup = dict
        Exit Function
    End If

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(tbl)

    Dim indexes As IncidentColumnIndexes
    indexes = BuildIncidentColumnIndexes(headerMap)

    Dim lookbackDays As Long
    lookbackDays = CLng(GetNamedRange("Config_LookbackDays"))

    Dim windowStart As Date
    windowStart = Date - lookbackDays

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim incidentDate As Date
        incidentDate = NzDate(data(r, indexes.IncidentDate))
        If incidentDate < windowStart Then GoTo ContinueRow

        Dim sourceSystem As String
        sourceSystem = NzString(data(r, indexes.SourceSystem))

        Dim scenarioName As String
        scenarioName = NzString(data(r, indexes.ScenarioName))

        Dim modelFamily As String
        modelFamily = NzString(data(r, indexes.ModelFamily))

        Dim key As String
        key = BuildHistoryKey(sourceSystem, scenarioName, modelFamily)

        Dim bucket As Object
        If dict.Exists(key) Then
            Set bucket = dict(key)
        Else
            Set bucket = CreateHistoryBucket()
        End If

        Dim failedRecords As Double
        failedRecords = NzDouble(data(r, indexes.FailedRecords))
        bucket("TotalFailedRecords") = bucket("TotalFailedRecords") + failedRecords

        Dim percentImpacted As Double
        percentImpacted = NzDouble(data(r, indexes.PercentImpacted))

        Dim missingAlerts As Double
        missingAlerts = NzDouble(data(r, indexes.MissingAlerts))
        bucket("TotalMissingAlerts") = bucket("TotalMissingAlerts") + missingAlerts

        If (missingAlerts > 0#) Or (percentImpacted > 0#) Then
            bucket("PositiveIncidents") = bucket("PositiveIncidents") + 1
        End If

        bucket("IncidentCount") = bucket("IncidentCount") + 1

        Set dict(key) = bucket
ContinueRow:
    Next r

    Set BuildHistoryRollup = dict
End Function

Private Function BuildOutputRow(ByVal rowIndex As Long, ByVal data As Variant, ByVal indexes As IncidentColumnIndexes, ByVal rollup As Object, ByVal severityTable As Variant, ByVal likelihoodTable As Variant, ByVal dqMatrix As Object, ByVal materialityRatios As Object, ByVal runTimestamp As Date, ByVal runUser As String, ByVal workbookVersion As String) As Variant
    Dim sourceSystem As String
    sourceSystem = NzString(data(rowIndex, indexes.SourceSystem))

    Dim assetClass As String
    assetClass = NzString(data(rowIndex, indexes.AssetClass))

    Dim scenarioName As String
    scenarioName = NzString(data(rowIndex, indexes.ScenarioName))

    Dim modelFamily As String
    modelFamily = NzString(data(rowIndex, indexes.ModelFamily))

    Dim failedRecords As Double
    failedRecords = NzDouble(data(rowIndex, indexes.FailedRecords))

    Dim percentImpacted As Double
    percentImpacted = NzDouble(data(rowIndex, indexes.PercentImpacted))

    Dim missingAlerts As Double
    missingAlerts = NzDouble(data(rowIndex, indexes.MissingAlerts))

    Dim incidentDate As Date
    incidentDate = NzDate(data(rowIndex, indexes.IncidentDate))

    Dim incidentId As String
    incidentId = NzString(data(rowIndex, indexes.SerialNumber))

    Dim historyKey As String
    historyKey = BuildHistoryKey(sourceSystem, scenarioName, modelFamily)

    Dim bucket As Object
    If rollup.Exists(historyKey) Then
        Set bucket = rollup(historyKey)
    Else
        Set bucket = CreateHistoryBucket()
    End If

    Dim totalFailedHistory As Double
    totalFailedHistory = bucket("TotalFailedRecords")

    Dim totalMissingHistory As Double
    totalMissingHistory = bucket("TotalMissingAlerts")

    Dim historyAlertRate As Double
    If totalFailedHistory = 0 Then
        historyAlertRate = 0
    Else
        historyAlertRate = totalMissingHistory / totalFailedHistory
    End If

    Dim missedAlerts As Double
    missedAlerts = missingAlerts

    Dim likelihoodBand As String
    likelihoodBand = DetermineLikelihood(missedAlerts, likelihoodTable)

    Dim severity As String
    severity = DetermineSeverity(percentImpacted, severityTable)

    Dim dqFinal As String
    dqFinal = ResolveDQFinal(severity, likelihoodBand, dqMatrix)

    Dim materialityRatio As Double
    Dim ratioFound As Boolean
    materialityRatio = ResolveMaterialityRatio(assetClass, dqFinal, materialityRatios, ratioFound)

    Dim materialityScore As Double
    materialityScore = missedAlerts * materialityRatio

    Dim positiveIncidents As Double
    If bucket.Exists("PositiveIncidents") Then
        positiveIncidents = bucket("PositiveIncidents")
    Else
        positiveIncidents = 0#
    End If

    Dim negativeIncidents As Double
    negativeIncidents = bucket("IncidentCount") - positiveIncidents
    If negativeIncidents < 0# Then negativeIncidents = 0#

    Dim jeffreysAlpha As Double
    jeffreysAlpha = positiveIncidents + 0.5

    Dim jeffreysBeta As Double
    jeffreysBeta = negativeIncidents + 0.5

    Dim storRateMean As Double
    If (jeffreysAlpha + jeffreysBeta) = 0# Then
        storRateMean = 0#
    Else
        storRateMean = jeffreysAlpha / (jeffreysAlpha + jeffreysBeta)
    End If

    Dim storRateUpper95 As Double
    storRateUpper95 = ComputeBetaInverse(0.95, jeffreysAlpha, jeffreysBeta)

    Dim jeffreysMateriality As Double
    jeffreysMateriality = materialityScore * storRateUpper95

    Dim noteText As String
    If bucket("IncidentCount") = 0 Then
        noteText = "No lookback history available for " & sourceSystem & " / " & scenarioName
    Else
        noteText = ""
    End If

    If Not ratioFound Then
        If Len(noteText) > 0 Then noteText = noteText & vbCrLf
        noteText = noteText & "No materiality ratio configured for asset class '" & IIf(Len(assetClass) = 0, "Unspecified", assetClass) & "' and risk '" & dqFinal & "'"
    End If

    Dim result(1 To 24) As Variant
    result(1) = incidentId
    result(2) = sourceSystem
    result(3) = assetClass
    result(4) = incidentDate
    result(5) = scenarioName
    result(6) = modelFamily
    result(7) = severity
    result(8) = failedRecords
    result(9) = percentImpacted
    result(10) = historyAlertRate
    result(11) = missedAlerts
    result(12) = likelihoodBand
    result(13) = dqFinal
    result(14) = materialityRatio
    result(15) = materialityScore
    result(16) = jeffreysAlpha
    result(17) = jeffreysBeta
    result(18) = storRateMean
    result(19) = storRateUpper95
    result(20) = jeffreysMateriality
    result(21) = runTimestamp
    result(22) = runUser
    result(23) = workbookVersion
    result(24) = noteText

    BuildOutputRow = result
End Function

Private Function ComputeBetaInverse(ByVal probability As Double, ByVal alpha As Double, ByVal beta As Double) As Double
    If probability <= 0# Then
        ComputeBetaInverse = 0#
        Exit Function
    ElseIf probability >= 1# Then
        ComputeBetaInverse = 1#
        Exit Function
    End If

    Dim resultValue As Double
    On Error Resume Next
    resultValue = Application.WorksheetFunction.Beta_Inv(probability, alpha, beta)
    If Err.Number = 0 Then
        ComputeBetaInverse = resultValue
        Exit Function
    End If

    Err.Clear
    resultValue = Application.WorksheetFunction.BetaInv(probability, alpha, beta)
    If Err.Number = 0 Then
        ComputeBetaInverse = resultValue
        Exit Function
    End If
    On Error GoTo 0

    ComputeBetaInverse = BetaInverseFallback(probability, alpha, beta)
End Function

Private Function BetaInverseFallback(ByVal probability As Double, ByVal alpha As Double, ByVal beta As Double) As Double
    Dim lower As Double
    Dim upper As Double
    lower = 0#
    upper = 1#

    Dim mid As Double
    Dim iter As Long
    For iter = 1 To 60
        mid = (lower + upper) / 2#
        Dim cdf As Double
        cdf = RegularizedIncompleteBeta(mid, alpha, beta)
        If cdf > probability Then
            upper = mid
        Else
            lower = mid
        End If
    Next iter

    BetaInverseFallback = (lower + upper) / 2#
End Function

Private Function RegularizedIncompleteBeta(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
    If x <= 0# Then
        RegularizedIncompleteBeta = 0#
        Exit Function
    ElseIf x >= 1# Then
        RegularizedIncompleteBeta = 1#
        Exit Function
    End If

    Dim bt As Double
    bt = Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b) + a * Log(x) + b * Log(1# - x))

    Dim resultValue As Double
    If x < (a + 1#) / (a + b + 2#) Then
        resultValue = bt * BetaContinuedFraction(x, a, b) / a
    Else
        resultValue = 1# - bt * BetaContinuedFraction(1# - x, b, a) / b
    End If

    If resultValue < 0# Then
        resultValue = 0#
    ElseIf resultValue > 1# Then
        resultValue = 1#
    End If

    RegularizedIncompleteBeta = resultValue
End Function

Private Function BetaContinuedFraction(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
    Const MAX_ITER As Long = 200
    Const EPS As Double = 3E-7
    Const FPMIN As Double = 1E-30

    Dim qab As Double
    Dim qap As Double
    Dim qam As Double
    qab = a + b
    qap = a + 1#
    qam = a - 1#

    Dim c As Double
    Dim d As Double
    Dim h As Double

    c = 1#
    d = 1# - (qab * x / qap)
    If Abs(d) < FPMIN Then d = FPMIN
    d = 1# / d
    h = d

    Dim m As Long
    For m = 1 To MAX_ITER
        Dim m2 As Long
        m2 = 2 * m

        Dim aa As Double
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1# + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1# / d
        h = h * d * c

        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1# + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1# / d
        Dim delta As Double
        delta = d * c
        h = h * delta

        If Abs(delta - 1#) < EPS Then Exit For
    Next m

    BetaContinuedFraction = h
End Function

Private Function LogGamma(ByVal z As Double) As Double
    Const g As Double = 7
    Const PI_CONST As Double = 3.14159265358979#
    Dim p As Variant
    p = Array(0.99999999999980993#, 676.5203681218851#, -1259.1392167224028#, _
              771.32342877765313#, -176.61502916214059#, 12.507343278686905#, _
              -0.13857109526572012#, 9.9843695780195716E-6, 1.5056327351493116E-7)

    If z < 0.5 Then
        LogGamma = Log(PI_CONST) - Log(Sin(PI_CONST * z)) - LogGamma(1# - z)
        Exit Function
    End If

    z = z - 1#
    Dim x As Double
    x = p(0)

    Dim i As Long
    For i = 1 To UBound(p)
        x = x + p(i) / (z + i)
    Next i

    Dim t As Double
    t = z + g + 0.5

    LogGamma = 0.5 * Log(2# * PI_CONST) + (z + 0.5) * Log(t) - t + Log(x)
End Function

Private Function ComputeOutputRows(ByVal rollup As Object, ByVal runTimestamp As Date) As Variant
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim tbl As ListObject
    Set tbl = wb.Worksheets(SHEET_INCIDENTS).ListObjects(TABLE_INCIDENTS_EXPANDED)

    Dim rowCount As Long
    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then
        ComputeOutputRows = VBA.Array()
        Exit Function
    End If

    Dim outputRows() As Variant
    ReDim outputRows(1 To rowCount, 1 To 24)

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(tbl)

    Dim indexes As IncidentColumnIndexes
    indexes = BuildIncidentColumnIndexes(headerMap)

    Dim severityTable As Variant
    severityTable = LoadTableData(SHEET_CONFIG, TABLE_SEVERITY)

    Dim likelihoodTable As Variant
    likelihoodTable = LoadTableData(SHEET_CONFIG, TABLE_LIKELIHOOD)

    Dim dqMatrix As Object
    Set dqMatrix = LoadDQMatrix()

    Dim materialityRatios As Object
    Set materialityRatios = LoadMaterialityRatios()

    Dim workbookVersion As String
    workbookVersion = CStr(GetNamedRange("Config_WorkbookVersion"))

    Dim runUser As String
    runUser = CStr(GetNamedRange("Config_RunUser"))

    Dim rowValues As Variant
    Dim outCol As Long
    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1)
        rowValues = BuildOutputRow(rowIndex, data, indexes, rollup, severityTable, likelihoodTable, dqMatrix, materialityRatios, runTimestamp, runUser, workbookVersion)

        For outCol = 1 To 24
            outputRows(rowIndex, outCol) = rowValues(outCol)
        Next outCol
    Next rowIndex

    ComputeOutputRows = outputRows
End Function

Public Sub RefreshAlertSummary(Optional ByVal showConfirmation As Boolean = True)
    Dim summaryTable As ListObject
    Set summaryTable = GetTableIfExists(SHEET_METRICS, TABLE_ALERT_SUMMARY)
    If summaryTable Is Nothing Then
        If showConfirmation Then
            MsgBox "Alert summary table not configured; skipping refresh.", vbInformation
        End If
        Exit Sub
    End If

    Dim summaryData As Variant
    Dim assetClasses As Collection
    Dim assetLabels As Object
    Call BuildAlertSummary(summaryData, assetClasses, assetLabels)

    WriteSummaryTable summaryTable, summaryData, assetClasses, assetLabels

    If showConfirmation Then
        MsgBox "Alert summary refreshed", vbInformation
    End If
End Sub

Private Sub BuildAlertSummary(ByRef summaryData As Variant, ByRef assetClasses As Collection, ByRef assetLabels As Object)
    Dim categories As Collection
    Dim storFlags As Object
    Call LoadMaterialCategories(categories, storFlags)

    Dim categoriesSeen As Object
    Set categoriesSeen = NewDictionary()

    Dim materialCounts As Object
    Set materialCounts = NewDictionary()

    Dim materialTotals As Object
    Set materialTotals = NewDictionary()

    Dim storTotals As Object
    Set storTotals = NewDictionary()

    Set assetLabels = NewDictionary()

    Call LoadMaterialOutputs(materialCounts, materialTotals, storTotals, categoriesSeen, storFlags, assetLabels)

    Dim alertCounts As Object
    Set alertCounts = NewDictionary()

    Dim alertTotals As Object
    Set alertTotals = NewDictionary()

    Dim levelInfo As Object
    Set levelInfo = NewDictionary()

    Call LoadSTSAlerts(alertCounts, alertTotals, levelInfo, assetLabels)

    Set assetClasses = BuildAssetClassList(materialTotals, alertTotals, assetLabels)

    If assetClasses Is Nothing Or assetClasses.Count = 0 Then
        summaryData = VBA.Array()
        Exit Sub
    End If

    Dim rowDefs As Collection
    Set rowDefs = New Collection

    Dim catItem As Variant
    For Each catItem In categories
        Dim rowDef As Object
        Set rowDef = CreateRowDef("Material Outputs (SWAT Data)", catItem("Label"), catItem("Key"), "CATEGORY", CBool(catItem("IsStor")))
        rowDefs.Add rowDef
        If categoriesSeen.Exists(catItem("Key")) Then categoriesSeen.Remove catItem("Key")
    Next catItem

    If categoriesSeen.Count > 0 Then
        Dim extraKeys As Variant
        extraKeys = categoriesSeen.Keys
        If IsArray(extraKeys) Then
            Call SortVariantArray(extraKeys)
            Dim i As Long
            For i = LBound(extraKeys) To UBound(extraKeys)
                Dim key As String
                key = extraKeys(i)
                Dim label As String
                label = categoriesSeen(key)
                Dim isStor As Boolean
                isStor = DetermineStorFlag(key, label, storFlags)
                rowDefs.Add CreateRowDef("Material Outputs (SWAT Data)", label, key, "CATEGORY", isStor)
            Next i
        End If
    End If

    rowDefs.Add CreateRowDef("Material Outputs (SWAT Data)", "Total Material Output Count", "TOTAL_MATERIAL", "TOTAL_MATERIAL")
    rowDefs.Add CreateRowDef("Material Outputs (SWAT Data)", "Total STOR-related Count", "TOTAL_STOR", "TOTAL_STOR")

    Dim levelKeys As Variant
    levelKeys = SortLevelKeys(levelInfo)
    If IsArray(levelKeys) Then
        Dim levelIndex As Long
        For levelIndex = LBound(levelKeys) To UBound(levelKeys)
            Dim levelKey As String
            levelKey = levelKeys(levelIndex)
            Dim info As Object
            Set info = levelInfo(levelKey)
            rowDefs.Add CreateRowDef("STS Closure Bucket - Escalation Level", CStr(info("Label")), levelKey, "ALERT_LEVEL")
        Next levelIndex
    End If

    rowDefs.Add CreateRowDef("STS Closure Bucket - Escalation Level", "Total Alerts", "TOTAL_ALERTS", "TOTAL_ALERTS")
    rowDefs.Add CreateRowDef("Ratios", "Total Material Output Ratio (% per alert)", "RATIO_MATERIAL", "RATIO_MATERIAL")
    rowDefs.Add CreateRowDef("Ratios", "STOR Ratio (% per alert)*", "RATIO_STOR", "RATIO_STOR")

    Dim rowCount As Long
    rowCount = rowDefs.Count
    If rowCount = 0 Then
        summaryData = VBA.Array()
        Exit Sub
    End If

    Dim columnCount As Long
    columnCount = 2 + assetClasses.Count

    ReDim summaryData(1 To rowCount, 1 To columnCount)

    Dim rowPosition As Long
    rowPosition = 0

    Dim colIndex As Long
    Dim rowDef As Variant
    For Each rowDef In rowDefs
        rowPosition = rowPosition + 1
        summaryData(rowPosition, 1) = rowDef("Section")
        summaryData(rowPosition, 2) = rowDef("Label")
        For colIndex = 1 To assetClasses.Count
            Dim assetKey As String
            assetKey = assetClasses(colIndex)
            Dim value As Double
            Select Case rowDef("Type")
                Case "CATEGORY"
                    value = DictGet(materialCounts, CompoundKey(assetKey, rowDef("Key")))
                Case "TOTAL_MATERIAL"
                    value = DictGet(materialTotals, assetKey)
                Case "TOTAL_STOR"
                    value = DictGet(storTotals, assetKey)
                Case "ALERT_LEVEL"
                    value = DictGet(alertCounts, CompoundKey(assetKey, rowDef("Key")))
                Case "TOTAL_ALERTS"
                    value = DictGet(alertTotals, assetKey)
                Case "RATIO_MATERIAL"
                    value = ComputeRatio(materialTotals, alertTotals, assetKey)
                Case "RATIO_STOR"
                    value = ComputeRatio(storTotals, alertTotals, assetKey)
                Case Else
                    value = 0
            End Select
            summaryData(rowPosition, 2 + colIndex) = value
        Next colIndex
    Next rowDef
End Sub

Private Function ComputeRatio(ByVal numeratorDict As Object, ByVal denominatorDict As Object, ByVal key As String) As Double
    Dim denominator As Double
    denominator = DictGet(denominatorDict, key)
    If denominator = 0 Then
        ComputeRatio = 0
    Else
        ComputeRatio = (DictGet(numeratorDict, key) / denominator) * 100#
    End If
End Function

Private Function CreateRowDef(ByVal section As String, ByVal label As String, ByVal key As String, ByVal rowType As String, Optional ByVal isStor As Boolean = False) As Object
    Dim entry As Object
    Set entry = NewDictionary()
    entry("Section") = section
    entry("Label") = label
    entry("Key") = key
    entry("Type") = rowType
    entry("IsStor") = isStor
    Set CreateRowDef = entry
End Function

Private Sub LoadMaterialCategories(ByRef categories As Collection, ByRef storFlags As Object)
    Set categories = New Collection
    Set storFlags = NewDictionary()

    Dim tbl As ListObject
    Set tbl = GetTableIfExists(SHEET_CONFIG, TABLE_MATERIAL_CATEGORIES)
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(tbl)

    Dim keyIndex As Long
    keyIndex = ResolveHeaderIndex(headerMap, "Category_Key", "Category", "Outcome_Type", "Outcome Type")
    If keyIndex = 0 Then Exit Sub

    Dim labelIndex As Long
    labelIndex = ResolveHeaderIndex(headerMap, "Display_Label", "Label", "Category_Label", "Category Label")

    Dim storIndex As Long
    storIndex = ResolveHeaderIndex(headerMap, "Is_STOR_Related", "IsStorRelated", "STOR_Flag", "STOR Flag", "Is_STOR")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim rawKey As String
        rawKey = NzString(data(r, keyIndex))
        If rawKey = "" Then GoTo ContinueRow

        Dim key As String
        key = NormalizeScenarioName(rawKey)
        If key = "" Then GoTo ContinueRow

        Dim label As String
        If labelIndex > 0 Then
            label = NzString(data(r, labelIndex))
            If label = "" Then label = rawKey
        Else
            label = rawKey
        End If

        Dim isStor As Boolean
        If storIndex > 0 Then
            isStor = ToBoolean(data(r, storIndex))
        Else
            isStor = (InStr(1, label, "stor", vbTextCompare) > 0)
        End If

        Dim entry As Object
        Set entry = NewDictionary()
        entry("Key") = key
        entry("Label") = label
        entry("IsStor") = isStor
        categories.Add entry
        storFlags(key) = isStor
ContinueRow:
    Next r
End Sub

Private Sub LoadMaterialOutputs(ByVal materialCounts As Object, ByVal materialTotals As Object, ByVal storTotals As Object, ByVal categoriesSeen As Object, ByVal storFlags As Object, ByVal assetLabels As Object)
    Dim tbl As ListObject
    Set tbl = GetTableIfExists(SHEET_METRICS, TABLE_MATERIAL_OUTPUTS)
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(tbl)

    Dim assetIndex As Long
    assetIndex = ResolveHeaderIndex(headerMap, "Asset_Class", "Asset Class")
    Dim categoryIndex As Long
    categoryIndex = ResolveHeaderIndex(headerMap, "Category_Key", "Category", "Outcome_Type", "Outcome Type", "Output_Type", "Output Type")
    Dim labelIndex As Long
    labelIndex = ResolveHeaderIndex(headerMap, "Category_Label", "Category Label", "Display_Label", "Display Label")
    Dim countIndex As Long
    countIndex = ResolveHeaderIndex(headerMap, "Count", "Output_Count", "Output Count", "Material_Count", "Material Count")

    If assetIndex = 0 Or categoryIndex = 0 Or countIndex = 0 Then Exit Sub

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim assetLabel As String
        assetLabel = NzString(data(r, assetIndex))
        If assetLabel = "" Then assetLabel = "Unspecified"
        Dim assetKey As String
        assetKey = NormalizeScenarioName(assetLabel)
        If assetKey = "" Then assetKey = "Unspecified"
        If Not assetLabels.Exists(assetKey) Then
            assetLabels(assetKey) = assetLabel
        End If

        Dim rawCategory As String
        rawCategory = NzString(data(r, categoryIndex))
        If rawCategory = "" Then GoTo ContinueRow

        Dim categoryKey As String
        categoryKey = NormalizeScenarioName(rawCategory)
        If categoryKey = "" Then GoTo ContinueRow

        Dim categoryLabel As String
        If labelIndex > 0 Then
            categoryLabel = NzString(data(r, labelIndex))
            If categoryLabel = "" Then categoryLabel = rawCategory
        Else
            categoryLabel = rawCategory
        End If

        If Not categoriesSeen.Exists(categoryKey) Then
            categoriesSeen(categoryKey) = categoryLabel
        End If

        Dim countValue As Double
        countValue = NzDouble(data(r, countIndex))
        If countValue = 0 Then GoTo ContinueRow

        DictAdd materialCounts, CompoundKey(assetKey, categoryKey), countValue
        DictAdd materialTotals, assetKey, countValue
        If DetermineStorFlag(categoryKey, categoryLabel, storFlags) Then
            DictAdd storTotals, assetKey, countValue
        End If
ContinueRow:
    Next r
End Sub

Private Sub LoadSTSAlerts(ByVal alertCounts As Object, ByVal alertTotals As Object, ByVal levelInfo As Object, ByVal assetLabels As Object)
    Dim tbl As ListObject
    Set tbl = GetTableIfExists(SHEET_METRICS, TABLE_STS_ALERTS)
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(tbl)

    Dim assetIndex As Long
    assetIndex = ResolveHeaderIndex(headerMap, "Asset_Class", "Asset Class")
    Dim levelIndex As Long
    levelIndex = ResolveHeaderIndex(headerMap, "Escalation_Level", "Escalation Level", "Level")
    Dim countIndex As Long
    countIndex = ResolveHeaderIndex(headerMap, "Alert_Count", "Alerts", "Count")

    If assetIndex = 0 Or levelIndex = 0 Or countIndex = 0 Then Exit Sub

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim assetLabel As String
        assetLabel = NzString(data(r, assetIndex))
        If assetLabel = "" Then assetLabel = "Unspecified"
        Dim assetKey As String
        assetKey = NormalizeScenarioName(assetLabel)
        If assetKey = "" Then assetKey = "Unspecified"
        If Not assetLabels.Exists(assetKey) Then
            assetLabels(assetKey) = assetLabel
        End If

        Dim levelLabel As String
        levelLabel = NzString(data(r, levelIndex))
        If levelLabel = "" Then levelLabel = "Unspecified"
        Dim levelKey As String
        levelKey = NormalizeScenarioName(levelLabel)
        If levelKey = "" Then levelKey = "Unspecified"

        Dim countValue As Double
        countValue = NzDouble(data(r, countIndex))
        If countValue <> 0 Then
            DictAdd alertCounts, CompoundKey(assetKey, levelKey), countValue
            DictAdd alertTotals, assetKey, countValue
        End If

        If Not levelInfo.Exists(levelKey) Then
            Dim info As Object
            Set info = NewDictionary()
            info("Label") = levelLabel
            info("Rank") = DetermineEscalationRank(levelKey)
            levelInfo(levelKey) = info
        End If
    Next r
End Sub

Private Function DetermineStorFlag(ByVal categoryKey As String, ByVal categoryLabel As String, ByVal storFlags As Object) As Boolean
    If Not storFlags Is Nothing Then
        If storFlags.Exists(categoryKey) Then
            DetermineStorFlag = CBool(storFlags(categoryKey))
            Exit Function
        End If
    End If
    DetermineStorFlag = (InStr(1, categoryLabel, "stor", vbTextCompare) > 0)
End Function

Private Function SortLevelKeys(ByVal levelInfo As Object) As Variant
    Dim keys As Variant
    keys = levelInfo.Keys
    If Not IsArray(keys) Then
        SortLevelKeys = keys
        Exit Function
    End If

    Dim i As Long, j As Long
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CompareLevels(keys(i), keys(j), levelInfo) > 0 Then
                Dim tmp As Variant
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i

    SortLevelKeys = keys
End Function

Private Function CompareLevels(ByVal keyA As String, ByVal keyB As String, ByVal levelInfo As Object) As Long
    Dim rankA As Long
    Dim rankB As Long
    rankA = LevelRank(levelInfo, keyA)
    rankB = LevelRank(levelInfo, keyB)
    If rankA <> rankB Then
        CompareLevels = Sgn(rankB - rankA)
        Exit Function
    End If

    Dim labelA As String
    Dim labelB As String
    labelA = CStr(levelInfo(keyA)("Label"))
    labelB = CStr(levelInfo(keyB)("Label"))
    CompareLevels = StrComp(labelA, labelB, vbTextCompare)
End Function

Private Function LevelRank(ByVal levelInfo As Object, ByVal key As String) As Long
    If levelInfo.Exists(key) Then
        Dim info As Object
        Set info = levelInfo(key)
        If info.Exists("Rank") Then
            LevelRank = CLng(info("Rank"))
            Exit Function
        End If
    End If
    LevelRank = 0
End Function

Private Function DetermineEscalationRank(ByVal levelKey As String) As Long
    Dim i As Long
    For i = 1 To Len(levelKey)
        Dim ch As String
        ch = Mid$(levelKey, i, 1)
        If ch >= "0" And ch <= "9" Then
            DetermineEscalationRank = CLng(Mid$(levelKey, i))
            Exit Function
        End If
    Next i
    DetermineEscalationRank = 0
End Function

Private Sub SortVariantArray(ByRef values As Variant)
    Dim i As Long, j As Long
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If StrComp(CStr(values(i)), CStr(values(j)), vbTextCompare) > 0 Then
                Dim tmp As Variant
                tmp = values(i)
                values(i) = values(j)
                values(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function BuildAssetClassList(ByVal materialTotals As Object, ByVal alertTotals As Object, ByVal assetLabels As Object) As Collection
    Dim combined As Object
    Set combined = NewDictionary()

    Dim key As Variant
    If Not assetLabels Is Nothing Then
        For Each key In assetLabels.Keys
            combined(key) = assetLabels(key)
        Next key
    End If

    For Each key In materialTotals.Keys
        If Not combined.Exists(key) Then
            If assetLabels.Exists(key) Then
                combined(key) = assetLabels(key)
            Else
                combined(key) = key
            End If
        End If
    Next key

    For Each key In alertTotals.Keys
        If Not combined.Exists(key) Then
            If assetLabels.Exists(key) Then
                combined(key) = assetLabels(key)
            Else
                combined(key) = key
            End If
        End If
    Next key

    Dim keys As Variant
    keys = combined.Keys
    If Not IsArray(keys) Then
        Set BuildAssetClassList = New Collection
        Exit Function
    End If

    Dim i As Long, j As Long
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(CStr(combined(keys(i))), CStr(combined(keys(j))), vbTextCompare) > 0 Then
                Dim tmp As Variant
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i

    Dim result As New Collection
    For i = LBound(keys) To UBound(keys)
        result.Add keys(i)
        If Not assetLabels.Exists(keys(i)) Then
            assetLabels(keys(i)) = combined(keys(i))
        End If
    Next i

    Set BuildAssetClassList = result
End Function

Private Function NewDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    dict.CompareMode = vbTextCompare
    On Error GoTo 0
    Set NewDictionary = dict
End Function

Private Sub DictAdd(ByVal dict As Object, ByVal key As String, ByVal value As Double)
    If dict.Exists(key) Then
        dict(key) = dict(key) + value
    Else
        dict.Add key, value
    End If
End Sub

Private Function DictGet(ByVal dict As Object, ByVal key As String) As Double
    If dict.Exists(key) Then
        DictGet = CDbl(dict(key))
    Else
        DictGet = 0#
    End If
End Function

Private Function CompoundKey(ByVal leftKey As String, ByVal rightKey As String) As String
    CompoundKey = NormalizeScenarioName(leftKey) & "|" & NormalizeScenarioName(rightKey)
End Function

Private Sub WriteSummaryTable(ByVal summaryTable As ListObject, ByVal summaryData As Variant, ByVal assetClasses As Collection, ByVal assetLabels As Object)
    Dim columnCount As Long
    If assetClasses Is Nothing Then
        columnCount = 2
    Else
        columnCount = 2 + assetClasses.Count
    End If

    Dim rowCount As Long
    rowCount = ArrayRowCount(summaryData)

    Dim totalRows As Long
    totalRows = rowCount + 1
    If totalRows < 2 Then totalRows = 2
    If columnCount < 2 Then columnCount = 2

    summaryTable.Resize summaryTable.Range.Resize(RowSize:=totalRows, ColumnSize:=columnCount)

    summaryTable.HeaderRowRange.Cells(1, 1).Value = "Section"
    summaryTable.HeaderRowRange.Cells(1, 2).Value = "Metric"

    Dim colIndex As Long
    If Not assetClasses Is Nothing Then
        For colIndex = 1 To assetClasses.Count
            Dim assetKey As String
            assetKey = assetClasses(colIndex)
            Dim headerLabel As String
            If Not assetLabels Is Nothing Then
                If assetLabels.Exists(assetKey) Then
                    headerLabel = CStr(assetLabels(assetKey))
                Else
                    headerLabel = assetKey
                End If
            Else
                headerLabel = assetKey
            End If
            summaryTable.HeaderRowRange.Cells(1, 2 + colIndex).Value = headerLabel
        Next colIndex
    End If

    Dim targetRange As Range
    Set targetRange = summaryTable.DataBodyRange
    If rowCount = 0 Then
        If Not targetRange Is Nothing Then
            targetRange.ClearContents
        End If
        Exit Sub
    End If

    targetRange.Value = summaryData
End Sub

Private Function GetTableIfExists(ByVal sheetName As String, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetTableIfExists = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If Err.Number <> 0 Then
        Set GetTableIfExists = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function ResolveHeaderIndex(ByVal headerMap As Object, ParamArray candidates() As Variant) As Long
    Dim candidate As Variant
    For Each candidate In candidates
        If headerMap.Exists(CStr(candidate)) Then
            ResolveHeaderIndex = headerMap(CStr(candidate))
            Exit Function
        End If
    Next candidate
    ResolveHeaderIndex = 0
End Function

Private Function ToBoolean(ByVal value As Variant) As Boolean
    If IsMissing(value) Or IsNull(value) Then
        ToBoolean = False
    ElseIf VarType(value) = vbBoolean Then
        ToBoolean = value
    ElseIf IsNumeric(value) Then
        ToBoolean = (CDbl(value) <> 0)
    Else
        Dim text As String
        text = LCase$(Trim$(CStr(value)))
        Select Case text
            Case "true", "yes", "y", "t", "1"
                ToBoolean = True
            Case Else
                ToBoolean = False
        End Select
    End If
End Function

Private Sub WriteOutput(ByVal rows As Variant)
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim tbl As ListObject
    Set tbl = wb.Worksheets(SHEET_OUTPUT).ListObjects(TABLE_OUTPUT)

    ClearTable tbl

    Dim rowCount As Long
    rowCount = ArrayRowCount(rows)
    If rowCount = 0 Then Exit Sub

    tbl.Resize tbl.Range.Resize(RowSize:=rowCount + 1)
    tbl.DataBodyRange.Value = rows
End Sub

Private Sub WriteHistoryFromIncidents()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim srcTable As ListObject
    Set srcTable = wb.Worksheets(SHEET_INCIDENTS).ListObjects(TABLE_INCIDENTS_EXPANDED)

    Dim historyTable As ListObject
    Set historyTable = wb.Worksheets(SHEET_HISTORY).ListObjects(TABLE_HISTORY_RAW)

    ClearTable historyTable
    EnsureHistoryTableLayout historyTable

    If srcTable.ListRows.Count = 0 Then
        UpdateHistoryNamedRanges historyTable
        Exit Sub
    End If

    Dim data As Variant
    data = srcTable.DataBodyRange.Value

    Dim headerMap As Object
    Set headerMap = BuildHeaderIndex(srcTable)

    Dim indexes As IncidentColumnIndexes
    indexes = BuildIncidentColumnIndexes(headerMap)

    Dim lookbackDays As Long
    lookbackDays = CLng(GetNamedRange("Config_LookbackDays"))

    Dim windowStart As Date
    windowStart = Date - lookbackDays

    Dim rows As Collection
    Set rows = New Collection

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim incidentDate As Date
        incidentDate = NzDate(data(r, indexes.IncidentDate))
        If incidentDate < windowStart Then GoTo ContinueRow

        Dim rowValues(1 To 9) As Variant
        rowValues(1) = incidentDate
        rowValues(2) = NzString(data(r, indexes.SourceSystem))
        rowValues(3) = NzString(data(r, indexes.AssetClass))
        rowValues(4) = NzString(data(r, indexes.ScenarioName))
        rowValues(5) = NzString(data(r, indexes.ModelFamily))
        rowValues(6) = NzDouble(data(r, indexes.FailedRecords))
        rowValues(7) = NzDouble(data(r, indexes.PercentImpacted))
        rowValues(8) = NzDouble(data(r, indexes.MissingAlerts))
        rowValues(9) = NzString(data(r, indexes.SerialNumber))
        rows.Add rowValues
ContinueRow:
    Next r

    If rows.Count = 0 Then
        UpdateHistoryNamedRanges historyTable
        Exit Sub
    End If

    Dim rowCount As Long
    rowCount = rows.Count

    Dim outputData() As Variant
    ReDim outputData(1 To rowCount, 1 To 9)

    Dim i As Long, c As Long
    For i = 1 To rowCount
        Dim values() As Variant
        values = rows(i)
        For c = 1 To 9
            outputData(i, c) = values(c)
        Next c
    Next i

    Dim headers As Variant
    headers = HistoryTableHeaders()

    historyTable.Resize historyTable.Range.Resize(RowSize:=rowCount + 1, ColumnSize:=UBound(headers) - LBound(headers) + 1)
    historyTable.DataBodyRange.Value = outputData

    UpdateHistoryNamedRanges historyTable
End Sub

Private Sub AppendAuditEntry(ByVal rows As Variant, ByVal runTimestamp As Date, Optional ByVal errorMessage As String = "")
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim tbl As ListObject
    Set tbl = wb.Worksheets(SHEET_AUDIT).ListObjects(TABLE_AUDIT)

    Dim runUser As String
    runUser = CStr(GetNamedRange("Config_RunUser"))

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add

    Dim target As Range
    Set target = newRow.Range
    target.Cells(1, 1).Value = runTimestamp
    target.Cells(1, 2).Value = runUser
    If LenB(errorMessage) = 0 Then
        target.Cells(1, 3).Value = ArrayRowCount(rows)
        target.Cells(1, 4).Value = ComputeRowsDigest(rows)
    Else
        target.Cells(1, 3).Value = ""
        target.Cells(1, 4).Value = "ERROR: " & errorMessage
    End If
End Sub

Private Function ComputeRowsDigest(ByVal rows As Variant) As String
    Dim rowCount As Long
    rowCount = ArrayRowCount(rows)
    If rowCount = 0 Then
        ComputeRowsDigest = ""
        Exit Function
    End If

    Dim buffer As String
    Dim r As Long, c As Long
    For r = LBound(rows, 1) To UBound(rows, 1)
        For c = LBound(rows, 2) To UBound(rows, 2)
            buffer = buffer & CStr(rows(r, c)) & "|"
        Next c
        buffer = buffer & vbLf
    Next r
    ComputeRowsDigest = Sha256Hex(buffer)
End Function

Private Function Sha256Hex(ByVal value As String) As String
    On Error GoTo fallback

    Dim sha As Object
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")

    Dim bytes() As Byte
    bytes = StrConv(value, vbFromUnicode)

    Dim hash() As Byte
    hash = sha.ComputeHash_2(bytes)

    Sha256Hex = BytesToHex(hash)
    Exit Function

fallback:
    Sha256Hex = SimpleChecksum(value)
End Function

Private Function BytesToHex(ByRef bytes() As Byte) As String
    Dim i As Long
    Dim chars() As String
    ReDim chars(0 To UBound(bytes))
    For i = 0 To UBound(bytes)
        chars(i) = Right$("0" & Hex$(bytes(i)), 2)
    Next i
    BytesToHex = Join(chars, "")
End Function

Private Function SimpleChecksum(ByVal value As String) As String
    Dim crc As Long
    crc = 0
    Dim i As Long
    For i = 1 To Len(value)
        crc = ((crc + Asc(Mid$(value, i, 1))) And &HFFFFFFFF)
        crc = ((crc * 31) And &HFFFFFFFF)
    Next i
    SimpleChecksum = Right$("00000000" & Hex$(crc), 8)
End Function

Private Function ArrayRowCount(ByVal rows As Variant) As Long
    On Error GoTo emptyArray
    If Not IsArray(rows) Then GoTo emptyArray
    ArrayRowCount = UBound(rows, 1) - LBound(rows, 1) + 1
    Exit Function
emptyArray:
    ArrayRowCount = 0
End Function

Private Function DetermineSeverity(ByVal pct As Double, ByVal tableData As Variant) As String
    Dim result As String
    result = "Medium"

    If ArrayRowCount(tableData) = 0 Then
        NotifyMissingConfig TABLE_SEVERITY
        DetermineSeverity = result
        Exit Function
    End If

    Dim i As Long
    For i = LBound(tableData, 1) To UBound(tableData, 1)
        Dim threshold As Double
        threshold = NzDouble(tableData(i, 1))
        If pct >= threshold Then
            result = CStr(tableData(i, 2))
        End If
    Next i

    DetermineSeverity = result
End Function

Private Function DetermineLikelihood(ByVal alertImpact As Double, ByVal tableData As Variant) As String
    Dim result As String
    result = "Low"

    If ArrayRowCount(tableData) = 0 Then
        NotifyMissingConfig TABLE_LIKELIHOOD
        DetermineLikelihood = result
        Exit Function
    End If

    Dim i As Long
    For i = LBound(tableData, 1) To UBound(tableData, 1)
        Dim threshold As Double
        threshold = NzDouble(tableData(i, 1))
        If alertImpact >= threshold Then
            result = CStr(tableData(i, 2))
        End If
    Next i

    DetermineLikelihood = result
End Function

Private Sub NotifyMissingConfig(ByVal tableName As String)
    Dim tracker As Object
    Set tracker = MissingConfigAlertTracker()

    If tracker.Exists(tableName) Then Exit Sub

    tracker(tableName) = True
    MsgBox "Configuration table '" & tableName & "' has no rows. Default thresholds will be used.", vbExclamation
End Sub

Private Function MissingConfigAlertTracker() As Object
    Static tracker As Object
    If tracker Is Nothing Then
        Set tracker = CreateObject("Scripting.Dictionary")
    End If
    Set MissingConfigAlertTracker = tracker
End Function

Private Function ResolveDQFinal(ByVal severity As String, ByVal likelihood As String, ByVal dqMatrix As Object) As String
    If dqMatrix.Exists(severity) Then
        Dim rowDict As Object
        Set rowDict = dqMatrix(severity)
        If rowDict.Exists(likelihood) Then
            ResolveDQFinal = rowDict(likelihood)
            Exit Function
        End If
    End If
    ResolveDQFinal = "Medium"
End Function

Private Function BuildHeaderIndex(ByVal table As ListObject) As Object
    Dim dict As Object
    Set dict = NewDictionary()

    Dim headers As Variant
    headers = table.HeaderRowRange.Value

    Dim c As Long
    For c = 1 To UBound(headers, 2)
        dict(CStr(headers(1, c))) = c
    Next c

    Set BuildHeaderIndex = dict
End Function

Private Function LoadTableData(ByVal sheetName As String, ByVal tableName As String) As Variant
    Dim tbl As ListObject
    Set tbl = GetTableIfExists(sheetName, tableName)

    If tbl Is Nothing Then
        NotifyMissingConfig tableName
        LoadTableData = VBA.Array()
        Exit Function
    End If

    If tbl.ListRows.Count = 0 Then
        NotifyMissingConfig tableName
        LoadTableData = VBA.Array()
        Exit Function
    End If

    LoadTableData = tbl.DataBodyRange.Value
End Function

Private Function LoadDQMatrix() As Object
    Dim tbl As ListObject
    Set tbl = GetTableIfExists(SHEET_CONFIG, TABLE_DQMATRIX)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If tbl Is Nothing Then
        NotifyMissingConfig TABLE_DQMATRIX
        Set LoadDQMatrix = dict
        Exit Function
    End If

    Dim headers As Variant
    headers = tbl.HeaderRowRange.Value

    Dim data As Variant
    If tbl.ListRows.Count = 0 Then
        NotifyMissingConfig TABLE_DQMATRIX
        Set LoadDQMatrix = dict
        Exit Function
    End If
    data = tbl.DataBodyRange.Value

    Dim r As Long, c As Long
    For r = 1 To UBound(data, 1)
        Dim severity As String
        severity = CStr(data(r, 1))

        Dim rowDict As Object
        Set rowDict = CreateObject("Scripting.Dictionary")

        For c = 2 To UBound(headers, 2)
            rowDict(CStr(headers(1, c))) = data(r, c)
        Next c

        dict(severity) = rowDict
    Next r

    Set LoadDQMatrix = dict
End Function

Private Function LoadMaterialityRatios() As Object
    Dim dict As Object
    Set dict = NewDictionary()

    Dim tbl As ListObject
    Set tbl = GetTableIfExists(SHEET_CONFIG, TABLE_MATERIALITY_RATIOS)
    If tbl Is Nothing Then
        NotifyMissingConfig TABLE_MATERIALITY_RATIOS
        Set LoadMaterialityRatios = dict
        Exit Function
    End If

    If tbl.ListRows.Count = 0 Then
        NotifyMissingConfig TABLE_MATERIALITY_RATIOS
        Set LoadMaterialityRatios = dict
        Exit Function
    End If

    Dim headers As Variant
    headers = tbl.HeaderRowRange.Value

    Dim data As Variant
    data = tbl.DataBodyRange.Value

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim assetKey As String
        assetKey = NormalizeScenarioName(NzString(data(r, 1)))
        If assetKey = "" Then assetKey = "DEFAULT"

        Dim rowDict As Object
        Set rowDict = NewDictionary()

        Dim c As Long
        For c = 2 To UBound(headers, 2)
            Dim headerName As String
            headerName = NormalizeScenarioName(CStr(headers(1, c)))
            If headerName <> "" Then
                rowDict(headerName) = NzDouble(data(r, c))
            End If
        Next c

        dict(assetKey) = rowDict
    Next r

    Set LoadMaterialityRatios = dict
End Function

Private Function ResolveMaterialityRatio(ByVal assetClass As String, ByVal dqFinal As String, ByVal ratios As Object, ByRef ratioFound As Boolean) As Double
    ratioFound = False

    If ratios Is Nothing Then
        ResolveMaterialityRatio = 0#
        Exit Function
    End If

    Dim assetKey As String
    assetKey = NormalizeScenarioName(assetClass)
    If assetKey = "" Then assetKey = "DEFAULT"

    Dim riskKey As String
    riskKey = NormalizeScenarioName(dqFinal)

    Dim ratioValue As Double
    If TryResolveMaterialityRatio(assetKey, riskKey, ratios, ratioValue) Then
        ratioFound = True
        ResolveMaterialityRatio = ratioValue
        Exit Function
    End If

    If assetKey <> "DEFAULT" Then
        If TryResolveMaterialityRatio("DEFAULT", riskKey, ratios, ratioValue) Then
            ratioFound = True
            ResolveMaterialityRatio = ratioValue
            Exit Function
        End If
    End If

    If TryResolveMaterialityRatio(assetKey, "DEFAULT", ratios, ratioValue) Then
        ratioFound = True
        ResolveMaterialityRatio = ratioValue
        Exit Function
    End If

    If assetKey <> "DEFAULT" Then
        If TryResolveMaterialityRatio("DEFAULT", "DEFAULT", ratios, ratioValue) Then
            ratioFound = True
            ResolveMaterialityRatio = ratioValue
            Exit Function
        End If
    End If

    ResolveMaterialityRatio = 0#
End Function

Private Function TryResolveMaterialityRatio(ByVal assetKey As String, ByVal riskKey As String, ByVal ratios As Object, ByRef ratioValue As Double) As Boolean
    If ratios Is Nothing Then Exit Function
    If Not ratios.Exists(assetKey) Then Exit Function

    Dim rowDict As Object
    Set rowDict = ratios(assetKey)
    If rowDict Is Nothing Then Exit Function

    Dim searchKeys As Collection
    Set searchKeys = New Collection

    On Error Resume Next
    If Len(riskKey) > 0 Then
        searchKeys.Add riskKey, "K" & riskKey
    Else
        searchKeys.Add "", "K"
    End If
    If Err.Number <> 0 Then Err.Clear

    Dim fallbackOptions As Variant
    fallbackOptions = VBA.Array("Default", "DEFAULT", "All", "*", "")

    Dim i As Long
    For i = LBound(fallbackOptions) To UBound(fallbackOptions)
        Dim candidate As String
        candidate = NormalizeScenarioName(CStr(fallbackOptions(i)))
        searchKeys.Add candidate, "K" & candidate
        If Err.Number <> 0 Then Err.Clear
    Next i
    On Error GoTo 0

    Dim item As Variant
    For Each item In searchKeys
        Dim candidateKey As String
        candidateKey = CStr(item)
        If rowDict.Exists(candidateKey) Then
            ratioValue = NzDouble(rowDict(candidateKey))
            TryResolveMaterialityRatio = True
            Exit Function
        End If
    Next item

    TryResolveMaterialityRatio = False
End Function

Private Function BuildHistoryKey(ByVal sourceSystem As String, ByVal scenarioName As String, Optional ByVal modelFamily As String = "") As String
    Dim key As String
    key = NormalizeScenarioName(sourceSystem) & "|" & NormalizeScenarioName(scenarioName)
    If Len(modelFamily) > 0 Then
        key = key & "|" & NormalizeScenarioName(modelFamily)
    End If
    BuildHistoryKey = key
End Function

Private Function GetNamedRange(ByVal name As String) As Variant
    GetNamedRange = ThisWorkbook.Names(name).RefersToRange.Value
End Function

Private Function CreateHistoryBucket() As Object
    Dim bucket As Object
    Set bucket = CreateObject("Scripting.Dictionary")
    bucket.CompareMode = vbTextCompare
    bucket("IncidentCount") = 0#
    bucket("PositiveIncidents") = 0#
    bucket("TotalFailedRecords") = 0#
    bucket("TotalMissingAlerts") = 0#
    Set CreateHistoryBucket = bucket
End Function

Private Sub EnsureHistoryTableLayout(ByVal tbl As ListObject)
    Dim headers As Variant
    headers = HistoryTableHeaders()

    Dim columnCount As Long
    columnCount = UBound(headers) - LBound(headers) + 1

    Dim targetRange As Range
    Set targetRange = tbl.Range.Resize(RowSize:=tbl.Range.Rows.Count, ColumnSize:=columnCount)
    tbl.Resize targetRange

    Dim headerRange As Range
    Set headerRange = tbl.HeaderRowRange

    Dim i As Long
    For i = 1 To columnCount
        headerRange.Cells(1, i).Value = headers(i - 1)
    Next i
End Sub

Private Function HistoryTableHeaders() As Variant
    HistoryTableHeaders = VBA.Array("Incident_Date", "Source_System", "Asset_Class", "Scenario_Name", "Model_Family", "Failed_Records", "Pct_Records_Impacted", "Missing_Alerts", "Serial_Number")
End Function

Private Sub UpdateHistoryNamedRanges(ByVal tbl As ListObject)
    AssignHistoryNamedRange "HistoryIncidentDates", tbl, "Incident_Date"
    AssignHistoryNamedRange "HistoryFailedRecords", tbl, "Failed_Records"
    AssignHistoryNamedRange "HistoryPercentImpacted", tbl, "Pct_Records_Impacted"
    AssignHistoryNamedRange "HistoryMissingAlerts", tbl, "Missing_Alerts"
End Sub

Private Sub AssignHistoryNamedRange(ByVal name As String, ByVal tbl As ListObject, ByVal columnName As String)
    Dim listColumn As ListColumn
    On Error Resume Next
    Set listColumn = tbl.ListColumns(columnName)
    On Error GoTo 0

    If listColumn Is Nothing Then Exit Sub

    Dim columnRange As Range
    On Error Resume Next
    Set columnRange = listColumn.DataBodyRange
    On Error GoTo 0

    If columnRange Is Nothing Then
        Set columnRange = listColumn.Range.Offset(1, 0).Resize(1)
    End If

    If columnRange Is Nothing Then Exit Sub

    On Error Resume Next
    ThisWorkbook.Names(name).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=name, RefersTo:=columnRange
End Sub

Private Function SplitScenarios(ByVal value As String) As Variant
    Dim trimmed As String
    trimmed = Trim$(value)

    If trimmed = "" Then
        SplitScenarios = Empty
        Exit Function
    End If

    Dim parts() As String
    parts = Split(trimmed, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = NormalizeScenarioName(parts(i))
    Next i

    SplitScenarios = parts
End Function

Private Function NormalizeScenarioName(ByVal value As String) As String
    Dim text As String
    text = CStr(value)
    text = Replace(text, vbCrLf, " ")
    text = Replace(text, vbTab, " ")
    text = Trim$(text)

    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop

    NormalizeScenarioName = text
End Function

Private Sub ClearTable(ByVal table As ListObject)
    On Error Resume Next
    If table.DataBodyRange Is Nothing Then
        ' nothing to clear
    Else
        table.DataBodyRange.ClearContents
        table.Resize table.Range.Resize(RowSize:=1)
    End If
    On Error GoTo 0
End Sub

Private Function NzString(ByVal value As Variant) As String
    If IsError(value) Then
        NzString = ""
    ElseIf IsMissing(value) Or IsNull(value) Then
        NzString = ""
    ElseIf VarType(value) = vbDate Then
        NzString = Format$(value, "yyyy-mm-dd")
    Else
        NzString = Trim$(CStr(value))
    End If
End Function

Private Function NzDouble(ByVal value As Variant) As Double
    If IsError(value) Or IsNull(value) Or value = "" Then
        NzDouble = 0
    Else
        NzDouble = CDbl(value)
    End If
End Function

Private Function NzDate(ByVal value As Variant) As Date
    If IsDate(value) Then
        NzDate = CDate(value)
    Else
        NzDate = 0
    End If
End Function

