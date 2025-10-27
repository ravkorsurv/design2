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
Private Const TABLE_SCENARIO_FAMILY As String = "ScenarioModelFamilies"
Private Const TABLE_MATERIAL_CATEGORIES As String = "MaterialCategories"
Private Const TABLE_MATERIAL_OUTPUTS As String = "MaterialOutputsRaw"
Private Const TABLE_STS_ALERTS As String = "STSAlertsRaw"
Private Const TABLE_ALERT_SUMMARY As String = "AlertSummary"

Public Sub RunDQMateriality()
    Dim runTimestamp As Date
    runTimestamp = Now

    ExpandIncidents

    Dim rollup As Object
    Set rollup = BuildHistoryRollup

    Dim outputRows As Variant
    outputRows = ComputeOutputRows(rollup, runTimestamp)

    WriteOutput outputRows
    AppendAuditEntry outputRows, runTimestamp

    RefreshAlertSummary False

    MsgBox "DQ/materiality calculations complete", vbInformation
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
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "([^\(]+?)\s*\(([^\)]+)\)"

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

Private Function BuildHistoryRollup() As Object
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim tbl As ListObject
    Set tbl = wb.Worksheets(SHEET_HISTORY).ListObjects(TABLE_HISTORY_RAW)

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

    Dim lookbackDays As Long
    lookbackDays = CLng(GetNamedRange("Config_LookbackDays"))

    Dim windowStart As Date
    windowStart = Date - lookbackDays

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim periodEnd As Date
        periodEnd = NzDate(data(r, headerMap("Period_End")))
        If periodEnd < windowStart Then GoTo ContinueRow

        Dim sourceSystem As String
        sourceSystem = NzString(data(r, headerMap("Source_System")))

        Dim scenarioName As String
        scenarioName = NzString(data(r, headerMap("Scenario_Name")))

        Dim key As String
        key = BuildHistoryKey(sourceSystem, scenarioName)

        Dim bucket As Variant
        If dict.Exists(key) Then
            bucket = dict(key)
        Else
            bucket = CreateHistoryBucket()
        End If

        bucket(0) = bucket(0) + NzDouble(data(r, headerMap("Records_Observed")))
        bucket(1) = bucket(1) + NzDouble(data(r, headerMap("Alerts_Investigated")))
        bucket(2) = bucket(2) + NzDouble(data(r, headerMap("Materiality_Positive")))

        dict(key) = bucket
ContinueRow:
    Next r

    Set BuildHistoryRollup = dict
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

    Dim severityTable As Variant
    severityTable = LoadTableData(SHEET_CONFIG, TABLE_SEVERITY)

    Dim likelihoodTable As Variant
    likelihoodTable = LoadTableData(SHEET_CONFIG, TABLE_LIKELIHOOD)

    Dim dqMatrix As Object
    Set dqMatrix = LoadDQMatrix()

    Dim workbookVersion As String
    workbookVersion = CStr(GetNamedRange("Config_WorkbookVersion"))

    Dim runUser As String
    runUser = CStr(GetNamedRange("Config_RunUser"))

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1)
        Dim sourceSystem As String
        sourceSystem = NzString(data(rowIndex, headerMap("Source_System")))

        Dim assetClass As String
        assetClass = NzString(data(rowIndex, headerMap("Asset_Class")))

        Dim scenarioName As String
        scenarioName = NzString(data(rowIndex, headerMap("Scenario_Name")))

        Dim modelFamily As String
        modelFamily = NzString(data(rowIndex, headerMap("Model_Family")))

        Dim failedRecords As Double
        failedRecords = NzDouble(data(rowIndex, headerMap("Failed_Records")))

        Dim percentImpacted As Double
        percentImpacted = NzDouble(data(rowIndex, headerMap("Pct_Records_Impacted")))

        Dim missingAlerts As Double
        missingAlerts = NzDouble(data(rowIndex, headerMap("Missing_Alerts")))

        Dim incidentDate As Date
        incidentDate = NzDate(data(rowIndex, headerMap("Incident_Date")))

        Dim incidentId As String
        incidentId = NzString(data(rowIndex, headerMap("Serial_Number")))

        Dim historyKey As String
        historyKey = BuildHistoryKey(sourceSystem, scenarioName)

        Dim bucket As Variant
        If rollup.Exists(historyKey) Then
            bucket = rollup(historyKey)
        Else
            bucket = CreateHistoryBucket()
        End If

        Dim historyAlertRate As Double
        If bucket(0) = 0 Then
            historyAlertRate = 0
        Else
            historyAlertRate = bucket(1) / bucket(0)
        End If

        Dim missedAlerts As Double
        missedAlerts = missingAlerts

        Dim likelihoodBand As String
        likelihoodBand = DetermineLikelihood(missedAlerts, likelihoodTable)

        Dim severity As String
        severity = DetermineSeverity(percentImpacted, severityTable)

        Dim dqFinal As String
        dqFinal = ResolveDQFinal(severity, likelihoodBand, dqMatrix)

        Dim alpha As Double
        alpha = bucket(2) + 0.5

        Dim beta As Double
        beta = (bucket(1) - bucket(2)) + 0.5

        Dim materialityMean As Double
        materialityMean = alpha / (alpha + beta)

        Dim materiality95 As Double
        materiality95 = BetaInverse(alpha, beta, 0.95)

        Dim expectedMean As Double
        expectedMean = missedAlerts * materialityMean

        Dim expected95 As Double
        expected95 = missedAlerts * materiality95

        Dim pAtLeastOne As Double
        pAtLeastOne = 1 - Exp(-expected95)

        Dim noteText As String
        If bucket(0) = 0 And bucket(1) = 0 Then
            noteText = "No lookback history available for " & sourceSystem & " / " & scenarioName
        Else
            noteText = ""
        End If

        Dim outCol As Long
        outCol = 1

        outputRows(rowIndex, outCol) = incidentId: outCol = outCol + 1
        outputRows(rowIndex, outCol) = sourceSystem: outCol = outCol + 1
        outputRows(rowIndex, outCol) = assetClass: outCol = outCol + 1
        outputRows(rowIndex, outCol) = incidentDate: outCol = outCol + 1
        outputRows(rowIndex, outCol) = scenarioName: outCol = outCol + 1
        outputRows(rowIndex, outCol) = modelFamily: outCol = outCol + 1
        outputRows(rowIndex, outCol) = severity: outCol = outCol + 1
        outputRows(rowIndex, outCol) = failedRecords: outCol = outCol + 1
        outputRows(rowIndex, outCol) = percentImpacted: outCol = outCol + 1
        outputRows(rowIndex, outCol) = historyAlertRate: outCol = outCol + 1
        outputRows(rowIndex, outCol) = missedAlerts: outCol = outCol + 1
        outputRows(rowIndex, outCol) = likelihoodBand: outCol = outCol + 1
        outputRows(rowIndex, outCol) = dqFinal: outCol = outCol + 1
        outputRows(rowIndex, outCol) = alpha: outCol = outCol + 1
        outputRows(rowIndex, outCol) = beta: outCol = outCol + 1
        outputRows(rowIndex, outCol) = materialityMean: outCol = outCol + 1
        outputRows(rowIndex, outCol) = materiality95: outCol = outCol + 1
        outputRows(rowIndex, outCol) = expectedMean: outCol = outCol + 1
        outputRows(rowIndex, outCol) = expected95: outCol = outCol + 1
        outputRows(rowIndex, outCol) = pAtLeastOne: outCol = outCol + 1
        outputRows(rowIndex, outCol) = runTimestamp: outCol = outCol + 1
        outputRows(rowIndex, outCol) = runUser: outCol = outCol + 1
        outputRows(rowIndex, outCol) = workbookVersion: outCol = outCol + 1
        outputRows(rowIndex, outCol) = noteText
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

Private Sub AppendAuditEntry(ByVal rows As Variant, ByVal runTimestamp As Date)
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim tbl As ListObject
    Set tbl = wb.Worksheets(SHEET_AUDIT).ListObjects(TABLE_AUDIT)

    Dim digest As String
    digest = ComputeRowsDigest(rows)

    Dim runUser As String
    runUser = CStr(GetNamedRange("Config_RunUser"))

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add

    Dim target As Range
    Set target = newRow.Range
    target.Cells(1, 1).Value = runTimestamp
    target.Cells(1, 2).Value = runUser
    target.Cells(1, 3).Value = ArrayRowCount(rows)
    target.Cells(1, 4).Value = digest
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

Private Function BetaInverse(ByVal alpha As Double, ByVal beta As Double, ByVal quantile As Double) As Double
    On Error GoTo tryLegacy
    BetaInverse = Application.WorksheetFunction.Beta_Inv(quantile, alpha, beta)
    Exit Function
tryLegacy:
    On Error GoTo numeric
    BetaInverse = Application.WorksheetFunction.BetaInv(quantile, alpha, beta)
    Exit Function
numeric:
    BetaInverse = BetaInverseNumeric(alpha, beta, quantile)
End Function

Private Function BetaInverseNumeric(ByVal alpha As Double, ByVal beta As Double, ByVal quantile As Double) As Double
    Const MAX_ITER As Long = 80
    Const EPS As Double = 1E-8

    Dim lo As Double, hi As Double, mid As Double
    lo = 0
    hi = 1

    Dim iter As Long
    For iter = 1 To MAX_ITER
        mid = (lo + hi) / 2
        Dim cdf As Double
        cdf = RegularizedBeta(alpha, beta, mid)
        If Abs(cdf - quantile) < EPS Then
            BetaInverseNumeric = mid
            Exit Function
        End If
        If cdf < quantile Then
            lo = mid
        Else
            hi = mid
        End If
    Next iter

    BetaInverseNumeric = (lo + hi) / 2
End Function

Private Function RegularizedBeta(ByVal alpha As Double, ByVal beta As Double, ByVal x As Double) As Double
    On Error GoTo legacyNew
    RegularizedBeta = Application.WorksheetFunction.Beta_Dist(x, alpha, beta, True)
    Exit Function
legacyNew:
    On Error GoTo legacyShort
    RegularizedBeta = Application.WorksheetFunction.BetaDist(x, alpha, beta, True)
    Exit Function
legacyShort:
    RegularizedBeta = Application.WorksheetFunction.BetaDist(x, alpha, beta)
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
    Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)

    If tbl.ListRows.Count = 0 Then
        Dim emptyData(1 To 1, 1 To tbl.ListColumns.Count) As Variant
        LoadTableData = emptyData
    Else
        LoadTableData = tbl.DataBodyRange.Value
    End If
End Function

Private Function LoadDQMatrix() As Object
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets(SHEET_CONFIG).ListObjects(TABLE_DQMATRIX)

    Dim headers As Variant
    headers = tbl.HeaderRowRange.Value

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim data As Variant
    If tbl.ListRows.Count = 0 Then
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

Private Function BuildHistoryKey(ByVal sourceSystem As String, ByVal scenarioName As String) As String
    BuildHistoryKey = NormalizeScenarioName(sourceSystem) & "|" & NormalizeScenarioName(scenarioName)
End Function

Private Function GetNamedRange(ByVal name As String) As Variant
    GetNamedRange = ThisWorkbook.Names(name).RefersToRange.Value
End Function

Private Function CreateHistoryBucket() As Variant
    Dim bucket(0 To 2) As Double
    bucket(0) = 0
    bucket(1) = 0
    bucket(2) = 0
    CreateHistoryBucket = bucket
End Function

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

