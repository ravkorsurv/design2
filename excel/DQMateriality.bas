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

Private Const TABLE_INCIDENTS_RAW As String = "IncidentsRaw"
Private Const TABLE_INCIDENTS_EXPANDED As String = "IncidentsExpanded"
Private Const TABLE_HISTORY_RAW As String = "HistoryRaw"
Private Const TABLE_OUTPUT As String = "OutputResults"
Private Const TABLE_AUDIT As String = "AuditLog"
Private Const TABLE_SEVERITY As String = "SeverityThresholds"
Private Const TABLE_LIKELIHOOD As String = "LikelihoodThresholds"
Private Const TABLE_DQMATRIX As String = "DQMatrix"

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
            Dim rowValues(1 To 7) As Variant
            rowValues(1) = incidentId
            rowValues(2) = sourceSystem
            rowValues(3) = incidentDate
            rowValues(4) = failedRecords
            rowValues(5) = percentImpacted
            rowValues(6) = impactItem("Scenario")
            rowValues(7) = impactItem("MissingAlerts")
            outRows.Add rowValues
        Next impactItem
    Next r

    If outRows.Count = 0 Then Exit Sub

    Dim outputData() As Variant
    ReDim outputData(1 To outRows.Count, 1 To 7)

    Dim i As Long, c As Long
    For i = 1 To outRows.Count
        Dim values() As Variant
        values = outRows(i)
        For c = 1 To 7
            outputData(i, c) = values(c)
        Next c
    Next i

    dstTable.Resize dstTable.Range.Resize(RowSize:=outRows.Count + 1)
    dstTable.DataBodyRange.Value = outputData
End Sub

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
    ReDim outputRows(1 To rowCount, 1 To 22)

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

        Dim scenarioName As String
        scenarioName = NzString(data(rowIndex, headerMap("Scenario_Name")))

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
        outputRows(rowIndex, outCol) = incidentDate: outCol = outCol + 1
        outputRows(rowIndex, outCol) = scenarioName: outCol = outCol + 1
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
    Set dict = CreateObject("Scripting.Dictionary")

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

