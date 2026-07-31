<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic" script:moduleType="normal">Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Private Const SH_DASH As String = "Dashboard"
Private Const SH_SETTINGS As String = "Settings"
Private Const SH_COSTS As String = "Costs"
Private Const SH_RESULTS As String = "Results"
Private Const SH_SCENARIO As String = "Scenario"
Private Const SH_HISTORY As String = "Price_History"
Private Const SH_DAILY As String = "Daily_Data"
Private Const SH_FIN As String = "RAW_Finance"
Private Const SH_FUNNEL As String = "RAW_Funnel"
Private Const SH_STOCK As String = "RAW_Stock"
Private Const SH_STOCKDET As String = "RAW_Stock_Detail"
Private Const SH_LOG As String = "Log"

Private gModelWb As Workbook
Private gCostImportSummary As String

Private Sub BindModelWorkbook()
    If gModelWb Is Nothing Then
        Set gModelWb = ActiveWorkbook
    End If
    If gModelWb Is Nothing Then
        Err.Raise 91, "BindModelWorkbook", "The model workbook is not active."
    End If
End Sub

Public Function SyntaxSmoke() As Long
    SyntaxSmoke = 1
End Function

Public Sub TestVBA()
    On Error GoTo EH
    Set gModelWb = ActiveWorkbook
    BindModelWorkbook
    MsgBox "VBA and workbook access work correctly (Excel 2019 mode, v10, hierarchical elasticity by subject)." & vbCrLf & _
           "Workbook: " & gModelWb.Name & vbCrLf & _
           "Dashboard: " & gModelWb.Worksheets(SH_DASH).Name, _
           vbInformation, "WB Price Optimizer"
    Exit Sub
EH:
    MsgBox "Workbook access error " & Err.Number & ": " & Err.Description, _
           vbCritical, "WB Price Optimizer"
End Sub

Public Sub InitializeModel()
    On Error GoTo EH
    Set gModelWb = ActiveWorkbook
    BindModelWorkbook
    gModelWb.Worksheets(SH_DASH).Activate
    MsgBox "Model is ready. Select four files: finance, funnel, stock history and cost master.", vbInformation, "WB Price Optimizer"
    Exit Sub
EH:
    MsgBox "Initialization error " & Err.Number & ":" & vbCrLf & Err.Description, vbCritical, "WB Price Optimizer"
End Sub

Public Sub SetupButtons()
    On Error Resume Next
    Set gModelWb = ActiveWorkbook
    BindModelWorkbook
    gModelWb.Worksheets(SH_DASH).Activate
    On Error GoTo 0
End Sub

Public Sub SelectReports()
    Dim fFin As String, fFunnel As String, fStock As String, fCosts As String
    Dim stage As String
    Dim errNum As Long, errDesc As String, errSource As String, errLine As Long
    On Error GoTo EH
    Set gModelWb = ActiveWorkbook
    BindModelWorkbook

100 stage = "selecting the detailed financial report"
110 fFin = PickExcelFile("1 of 4: Select the detailed financial report")
120 If Len(fFin) = 0 Then Exit Sub

130 stage = "selecting the sales funnel report"
140 fFunnel = PickExcelFile("2 of 4: Select the sales funnel report")
150 If Len(fFunnel) = 0 Then Exit Sub

160 stage = "selecting the stock history report"
170 fStock = PickExcelFile("3 of 4: Select the stock history report")
180 If Len(fStock) = 0 Then Exit Sub

190 stage = "selecting the cost master file"
200 fCosts = PickExcelFile("4 of 4: Select the cost file (article and cost)")
210 If Len(fCosts) = 0 Then Exit Sub

220 stage = "starting report import"
230 ImportAllReports fFin, fFunnel, fStock, fCosts
240 Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl
    MsgBox "Error while " & stage & "." & vbCrLf & _
           "Error " & errNum & ": " & errDesc & vbCrLf & _
           "Source: " & errSource & vbCrLf & _
           "Line: " & errLine, vbCritical, "WB Price Optimizer"
End Sub

Private Function PickExcelFile(ByVal titleText As String) As String
    Dim selectedFile As Variant
    selectedFile = Application.GetOpenFilename( _
        FileFilter:="Excel files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:=titleText, _
        MultiSelect:=False)

    If VarType(selectedFile) = vbBoolean Then
        If selectedFile = False Then Exit Function
    End If

    PickExcelFile = CStr(selectedFile)
End Function

Public Sub RefreshFromFolder()
    Dim folder As String
    Dim fFin As String, fFunnel As String, fStock As String, fCosts As String
    Dim defaultFolder As String
    On Error GoTo EH
    Set gModelWb = ActiveWorkbook
    BindModelWorkbook

    defaultFolder = Txt(gModelWb.Worksheets(SH_SETTINGS).Range("B25").Value2)
    folder = InputBox("Enter the folder containing four files: finance, funnel, stock history and cost master:", _
                      "WB Price Optimizer", defaultFolder)
    If Len(folder) = 0 Then Exit Sub
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator

    DetectReportFiles folder, fFin, fFunnel, fStock, fCosts
    If Len(fFin) = 0 Or Len(fFunnel) = 0 Or Len(fStock) = 0 Or Len(fCosts) = 0 Then
        MsgBox "Four files were not identified." & vbCrLf & vbCrLf & _
               "Financial: " & IIf(Len(fFin) > 0, fFin, "not found") & vbCrLf & _
               "Funnel: " & IIf(Len(fFunnel) > 0, fFunnel, "not found") & vbCrLf & _
               "Stock: " & IIf(Len(fStock) > 0, fStock, "not found") & vbCrLf & _
               "Costs: " & IIf(Len(fCosts) > 0, fCosts, "not found") & vbCrLf & vbCrLf & _
               "Use SELECT 4 FILES for manual selection.", vbExclamation, "WB Price Optimizer"
        Exit Sub
    End If
    ImportAllReports fFin, fFunnel, fStock, fCosts
    Exit Sub
EH:
    MsgBox "Folder refresh error " & Err.Number & ":" & vbCrLf & Err.Description, vbCritical, "WB Price Optimizer"
End Sub

Private Sub DetectReportFiles(ByVal folder As String, ByRef fFin As String, ByRef fFunnel As String, _
                              ByRef fStock As String, ByRef fCosts As String)
    Dim f As String, fullName As String
    Dim wb As Workbook
    Dim sheetCount As Long, cols1 As Long, cols3 As Long, cols4 As Long, rows1 As Long
    Dim dt As Date, dtFin As Date, dtFunnel As Date, dtStock As Date, dtCosts As Date

    f = Dir$(folder & "*.xls*")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then
            fullName = folder & f
            If StrComp(fullName, gModelWb.fullName, vbTextCompare) <> 0 Then
                Set wb = Nothing
                On Error Resume Next
                Set wb = Workbooks.Open(Filename:=fullName, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
                On Error GoTo EH
                If Not wb Is Nothing Then
                    sheetCount = wb.Worksheets.count
                    cols1 = 0: cols3 = 0: cols4 = 0: rows1 = 0
                    If sheetCount >= 1 Then
                        cols1 = wb.Worksheets(1).UsedRange.Columns.count
                        rows1 = wb.Worksheets(1).UsedRange.Rows.count
                    End If
                    If sheetCount >= 3 Then cols3 = wb.Worksheets(3).UsedRange.Columns.count
                    If sheetCount >= 4 Then cols4 = wb.Worksheets(4).UsedRange.Columns.count
                    dt = FileDateTime(fullName)
                    If sheetCount = 1 And cols1 >= 70 Then
                        If Len(fFin) = 0 Or dt > dtFin Then fFin = fullName: dtFin = dt
                    ElseIf sheetCount >= 4 And cols3 >= 60 And cols4 >= 20 Then
                        If Len(fStock) = 0 Or dt > dtStock Then fStock = fullName: dtStock = dt
                    ElseIf sheetCount >= 3 And cols3 >= 15 And cols3 <= 40 Then
                        If Len(fFunnel) = 0 Or dt > dtFunnel Then fFunnel = fullName: dtFunnel = dt
                    ElseIf sheetCount = 1 And cols1 >= 2 And cols1 <= 10 And rows1 >= 2 Then
                        If Len(fCosts) = 0 Or dt > dtCosts Then fCosts = fullName: dtCosts = dt
                    End If
                    wb.Close SaveChanges:=False
                    Set wb = Nothing
                End If
            End If
        End If
        f = Dir$()
    Loop
    Exit Sub
EH:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, "DetectReportFiles", Err.Description
End Sub

Private Sub ImportAllReports(ByVal fFin As String, ByVal fFunnel As String, ByVal fStock As String, ByVal fCosts As String)
    Dim oldCalc As XlCalculation
    Dim stage As String
    Dim errNum As Long, errDesc As String, errSource As String, errLine As Long
    On Error GoTo EH
    BindModelWorkbook
    gCostImportSummary = ""

1000 oldCalc = Application.Calculation
1010 Application.ScreenUpdating = False
1020 Application.EnableEvents = False
1030 Application.DisplayAlerts = False
1040 Application.Calculation = xlCalculationManual

1100 stage = "importing the financial report"
1110 ImportSheetValues fFin, 1, SH_FIN
1120 stage = "importing the funnel report"
1130 ImportSheetValues fFunnel, 3, SH_FUNNEL
1140 stage = "importing the stock report"
1150 ImportSheetValues fStock, 3, SH_STOCK
1160 stage = "importing stock details"
1170 ImportSheetValues fStock, 4, SH_STOCKDET
1180 stage = "importing article costs"
1190 ImportCostsFile fCosts

1200 stage = "saving import settings"
1210 With gModelWb.Worksheets(SH_SETTINGS)
1220     .Range("B24").Value = Now
1230     .Range("B25").Value = Left$(fFin, InStrRev(fFin, Application.PathSeparator))
1240 End With

1300 Application.DisplayAlerts = True
1310 Application.EnableEvents = True
1320 Application.ScreenUpdating = True
1330 Application.Calculation = oldCalc

1400 stage = "calculating all articles"
1410 RecalculateAll
1420 Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Report processing error while " & stage & "." & vbCrLf & _
           "Error " & errNum & ": " & errDesc & vbCrLf & _
           "Source: " & errSource & vbCrLf & _
           "Line: " & errLine, vbCritical, "WB Price Optimizer"
End Sub

Private Sub ImportSheetValues(ByVal filePath As String, ByVal sourceIndex As Long, ByVal destinationSheet As String)
    Dim wb As Workbook, src As Worksheet, dst As Worksheet
    Dim arr As Variant, ur As Range
    Dim errNum As Long, errDesc As String, errSource As String, errLine As Long
    On Error GoTo EH
    BindModelWorkbook

2000 If Len(Dir$(filePath)) = 0 Then Err.Raise 53, "ImportSheetValues", "File not found: " & filePath
2010 Set dst = gModelWb.Worksheets(destinationSheet)
2020 dst.Cells.Clear
2030 Set wb = Application.Workbooks.Open( _
        Filename:=filePath, _
        ReadOnly:=True, _
        UpdateLinks:=False, _
        AddToMru:=False, _
        IgnoreReadOnlyRecommended:=True)
2040 If sourceIndex < 1 Or sourceIndex > wb.Worksheets.count Then
2050     Err.Raise vbObjectError + 100, "ImportSheetValues", _
             "Required sheet index " & sourceIndex & " not found in " & filePath
2060 End If
2070 Set src = wb.Worksheets(sourceIndex)
2080 Set ur = src.UsedRange
2090 arr = ur.Value2
2100 If ur.Cells.CountLarge = 1 Then
2110     dst.Range("A1").Value2 = arr
2120 Else
2130     dst.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value2 = arr
2140 End If
2150 dst.Rows(1).Font.Bold = True
2160 wb.Close SaveChanges:=False
2170 Set wb = Nothing
2180 Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    On Error GoTo 0
    Err.Raise errNum, "ImportSheetValues line " & errLine & " / " & errSource, errDesc
End Sub

Private Sub ImportCostsFile(ByVal filePath As String)
    Dim wb As Workbook, src As Worksheet, dst As Worksheet
    Dim lastRow As Long, r As Long, cr As Long, nextRow As Long
    Dim art As String, unitCost As Double, warehouseDelivery As Double
    Dim loaded As Long, added As Long, duplicates As Long, skipped As Long
    Dim errNum As Long, errDesc As String, errSource As String, errLine As Long
    On Error GoTo EH
    BindModelWorkbook

3000 If Len(Dir$(filePath)) = 0 Then Err.Raise 53, "ImportCostsFile", "File not found: " & filePath
3010 Set dst = gModelWb.Worksheets(SH_COSTS)
3020 Set wb = Application.Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False, _
        AddToMru:=False, IgnoreReadOnlyRecommended:=True)
3030 If wb.Worksheets.count < 1 Then Err.Raise vbObjectError + 201, "ImportCostsFile", "The cost file has no worksheets."
3040 Set src = wb.Worksheets(1)
3050 lastRow = LastUsedRow(src)
3060 If lastRow < 2 Then Err.Raise vbObjectError + 202, "ImportCostsFile", "The cost file has no data rows."
3070 nextRow = LastUsedRow(dst) + 1
3080 If nextRow < 2 Then nextRow = 2

3090 For r = 2 To lastRow
3100     art = Txt(src.Cells(r, 1).Value2)
3110     unitCost = SafeVal(src.Cells(r, 2).Value2)
3120     warehouseDelivery = SafeVal(src.Cells(r, 3).Value2)
3130     If warehouseDelivery < 0 Then warehouseDelivery = 0
3140     If Len(art) = 0 Or unitCost <= 0 Then
3150         skipped = skipped + 1
3160     ElseIf CostArticleSeenBefore(src, r, art) Then
3170         duplicates = duplicates + 1
3180     Else
3190         cr = FindCostRow(art)
3200         If cr = 0 Then
3210             cr = nextRow
3220             nextRow = nextRow + 1
3230             dst.Cells(cr, 1).NumberFormat = "@"
3240             dst.Cells(cr, 1).Value2 = art
3250             dst.Cells(cr, 16).Value2 = "YES"
3260             added = added + 1
3270         End If
3280         dst.Cells(cr, 4).ClearContents
3290         dst.Cells(cr, 5).Value2 = unitCost
3300         dst.Range(dst.Cells(cr, 6), dst.Cells(cr, 15)).ClearContents
3310         dst.Cells(cr, 8).Value2 = warehouseDelivery
3320         dst.Cells(cr, 10).Value2 = warehouseDelivery
3330         dst.Cells(cr, 16).Value2 = "YES"
3340         dst.Cells(cr, 17).Value2 = "Cost and warehouse delivery imported from " & Dir$(filePath) & _
                " on " & Format$(Now, "dd.mm.yyyy hh:nn")
3350         dst.Range(dst.Cells(cr, 1), dst.Cells(cr, 17)).Interior.Pattern = xlNone
3360         loaded = loaded + 1
3370     End If
3380 Next r

3390 wb.Close SaveChanges:=False
3400 Set wb = Nothing
3410 gCostImportSummary = "Costs loaded: " & loaded & "; new cost rows: " & added & _
        "; duplicates skipped: " & duplicates & "; invalid rows skipped: " & skipped
3420 WriteLog "Cost import", gCostImportSummary
3430 Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    On Error GoTo 0
    Err.Raise errNum, "ImportCostsFile line " & errLine & " / " & errSource, errDesc
End Sub

Private Function CostArticleSeenBefore(ByVal src As Worksheet, ByVal currentRow As Long, ByVal art As String) As Boolean
    Dim r As Long
    For r = 2 To currentRow - 1
        If StrComp(Txt(src.Cells(r, 1).Value2), art, vbTextCompare) = 0 Then
            CostArticleSeenBefore = True
            Exit Function
        End If
    Next r
End Function

Public Sub RecalculateAll()
    Dim oldCalc As XlCalculation
    Dim wsRes As Worksheet, wsHist As Worksheet, wsDaily As Worksheet
    Dim lastResultRow As Long, lastDailyRow As Long
    Dim stage As String
    Dim calcErrNum As Long, calcErrDesc As String, calcErrSource As String, calcErrLine As Long
    On Error GoTo EH
    Set gModelWb = ActiveWorkbook
    BindModelWorkbook

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsRes = gModelWb.Worksheets(SH_RESULTS)
    Set wsHist = gModelWb.Worksheets(SH_HISTORY)
    Set wsDaily = gModelWb.Worksheets(SH_DAILY)

    stage = "clearing previous calculations"
    wsRes.Rows("2:" & wsRes.Rows.count).ClearContents
    wsHist.Rows("2:" & wsHist.Rows.count).ClearContents
    wsDaily.Rows("2:" & wsDaily.Rows.count).ClearContents

    stage = "reading costs"
    AddArticlesFromCosts wsRes, lastResultRow
    stage = "reading the financial report"
    AggregateFinance wsRes, lastResultRow
    stage = "reading the funnel report"
    AppendFunnelDaily wsRes, lastResultRow, wsDaily, lastDailyRow
    stage = "reading stock history"
    AppendStockDaily wsRes, lastResultRow, wsDaily, lastDailyRow
    stage = "reading stock details"
    AggregateStockDetail wsRes, lastResultRow
    stage = "adding missing cost rows"
    EnsureCostRowsNoDictionary wsRes, lastResultRow

    stage = "sorting result articles (Excel 2019 mode)"
    SortResultsByArticle wsRes, lastResultRow
    stage = "sorting daily data (Excel 2019 mode)"
    SortDailyByArticleDate wsDaily, lastDailyRow
    stage = "consolidating daily data (Excel 2019 mode)"
    lastDailyRow = ConsolidateDailyRows(wsDaily, lastDailyRow)

    stage = "calculating raw article metrics"
    CalculateResultRows wsRes, lastResultRow, wsHist, wsDaily, lastDailyRow

    stage = "calculating hierarchical elasticity by subject and repricing"
    ApplyHierarchicalElasticityAndReprice wsRes, lastResultRow

    stage = "sorting final results"
    SortResultsByUplift wsRes, lastResultRow
    stage = "updating dashboard"
    UpdateDashboard lastResultRow
    stage = "building scenario"
    BuildScenario
    gModelWb.Worksheets(SH_SETTINGS).Range("B24").Value = Now
    WriteLog "Calculation", "Articles processed: " & Application.Max(0, lastResultRow - 1)

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Len(gCostImportSummary) > 0 Then
        MsgBox "Calculation completed. Articles processed: " & Application.Max(0, lastResultRow - 1) & vbCrLf & _
               gCostImportSummary, vbInformation, "WB Price Optimizer"
        gCostImportSummary = ""
    Else
        MsgBox "Calculation completed. Articles processed: " & Application.Max(0, lastResultRow - 1), vbInformation, "WB Price Optimizer"
    End If
    Exit Sub
EH:
    calcErrNum = Err.Number
    calcErrDesc = Err.Description
    calcErrSource = Err.Source
    calcErrLine = Erl
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Calculation error while " & stage & "." & vbCrLf & _
           "Error " & calcErrNum & ": " & calcErrDesc & vbCrLf & _
           "Source: " & calcErrSource & vbCrLf & _
           "Line: " & calcErrLine, vbCritical, "WB Price Optimizer"
End Sub

Private Sub AddArticlesFromCosts(ByVal wsRes As Worksheet, ByRef lastResultRow As Long)
    'Results are seeded from WB reports. The cost master may contain more articles than the selected period.
    lastResultRow = 1
End Sub

Private Function EnsureResultArticle(ByVal wsRes As Worksheet, ByRef lastResultRow As Long, _
                                     ByVal art As String, ByVal wbArt As String, ByVal itemName As String) As Long
    Dim rr As Long
    rr = FindResultRow(wsRes, lastResultRow, art)
    If rr = 0 Then
        lastResultRow = Application.Max(2, lastResultRow + 1)
        rr = lastResultRow
        wsRes.Cells(rr, 1).NumberFormat = "@"
        wsRes.Cells(rr, 1).Value = art
    End If
    If Len(wbArt) > 0 Then wsRes.Cells(rr, 2).Value = wbArt
    If Len(itemName) > 0 Then wsRes.Cells(rr, 3).Value = itemName
    EnsureResultArticle = rr
End Function

Private Function FindResultRow(ByVal wsRes As Worksheet, ByVal lastResultRow As Long, ByVal art As String) As Long
    Dim r As Long
    If lastResultRow < 2 Then Exit Function
    For r = 2 To lastResultRow
        If StrComp(Txt(wsRes.Cells(r, 1).Value2), art, vbTextCompare) = 0 Then
            FindResultRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub AggregateFinance(ByVal wsRes As Worksheet, ByRef lastResultRow As Long)
    Dim ws As Worksheet, r As Long, lastRow As Long, rr As Long
    Dim art As String, qty As Double, priceVal As Double, deliveries As Double
    Dim saleDate As Date, lastSaleDate As Double
    Const cWB As Long = 4, cArt As Long = 6, cName As Long = 7
    Const cSaleDate As Long = 13, cQty As Long = 14, cPrice As Long = 15, cCommission As Long = 24
    Const cDelCount As Long = 35, cRetCount As Long = 36, cDelivery As Long = 37
    Const cStorage As Long = 60, cHold As Long = 61, cAccept As Long = 62, cMode As Long = 75
    Set ws = gModelWb.Worksheets(SH_FIN)
    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then Exit Sub
    For r = 2 To lastRow
        art = Txt(ws.Cells(r, cArt).Value2)
        If Len(art) > 0 Then
            rr = EnsureResultArticle(wsRes, lastResultRow, art, Txt(ws.Cells(r, cWB).Value2), Txt(ws.Cells(r, cName).Value2))
            If Len(Txt(ws.Cells(r, cMode).Value2)) > 0 Then wsRes.Cells(rr, 47).Value2 = Txt(ws.Cells(r, cMode).Value2)
            qty = SafeVal(ws.Cells(r, cQty).Value2)
            priceVal = SafeVal(ws.Cells(r, cPrice).Value2)
            If qty > 0 And priceVal > 0 Then
                wsRes.Cells(rr, 37).Value2 = SafeVal(wsRes.Cells(rr, 37).Value2) + qty
                wsRes.Cells(rr, 38).Value2 = SafeVal(wsRes.Cells(rr, 38).Value2) + priceVal * qty
                wsRes.Cells(rr, 39).Value2 = SafeVal(wsRes.Cells(rr, 39).Value2) + SafeVal(ws.Cells(r, cCommission).Value2) * qty
                saleDate = ParseDateValue(ws.Cells(r, cSaleDate).Value)
                lastSaleDate = SafeVal(wsRes.Cells(rr, 48).Value2)
                If saleDate <> 0 Then
                    If CDbl(saleDate) >= lastSaleDate Then
                        wsRes.Cells(rr, 48).Value2 = CDbl(saleDate)
                        wsRes.Cells(rr, 49).Value2 = priceVal
                    End If
                ElseIf SafeVal(wsRes.Cells(rr, 49).Value2) <= 0 Then
                    wsRes.Cells(rr, 49).Value2 = priceVal
                End If
            End If
            wsRes.Cells(rr, 40).Value2 = SafeVal(wsRes.Cells(rr, 40).Value2) + Abs(SafeCellVal(ws, r, cDelivery))
            deliveries = Abs(SafeCellVal(ws, r, cDelCount)) + Abs(SafeCellVal(ws, r, cRetCount))
            wsRes.Cells(rr, 41).Value2 = SafeVal(wsRes.Cells(rr, 41).Value2) + deliveries
            wsRes.Cells(rr, 42).Value2 = SafeVal(wsRes.Cells(rr, 42).Value2) + Abs(SafeCellVal(ws, r, cStorage))
            wsRes.Cells(rr, 43).Value2 = SafeVal(wsRes.Cells(rr, 43).Value2) + Abs(SafeCellVal(ws, r, cAccept))
            wsRes.Cells(rr, 44).Value2 = SafeVal(wsRes.Cells(rr, 44).Value2) + Abs(SafeCellVal(ws, r, cHold))
        End If
    Next r
End Sub

Private Sub AppendFunnelDaily(ByVal wsRes As Worksheet, ByRef lastResultRow As Long, _
                              ByVal wsDaily As Worksheet, ByRef lastDailyRow As Long)
    Dim ws As Worksheet, r As Long, lastRow As Long, rr As Long
    Dim art As String, subjectName As String, dt As Date
    Dim orders As Double, buys As Double, price As Double
    Const cArt As Long = 1, cWB As Long = 2, cName As Long = 3, cSubject As Long = 4
    Const colDate As Long = 10, cOrders As Long = 14, cBuy As Long = 15, cAmount As Long = 20

    Set ws = gModelWb.Worksheets(SH_FUNNEL)
    lastRow = LastUsedRow(ws)
    lastDailyRow = 1
    If lastRow < 3 Then Exit Sub

    For r = 3 To lastRow
        art = Txt(ws.Cells(r, cArt).Value2)
        dt = ParseDateValue(ws.Cells(r, colDate).Value)

        If Len(art) > 0 And dt <> 0 Then
            rr = EnsureResultArticle(wsRes, lastResultRow, art, _
                                     Txt(ws.Cells(r, cWB).Value2), _
                                     Txt(ws.Cells(r, cName).Value2))

            subjectName = Txt(ws.Cells(r, cSubject).Value2)
            If Len(subjectName) > 0 Then wsRes.Cells(rr, 51).Value2 = subjectName

            orders = SafeVal(ws.Cells(r, cOrders).Value2)
            buys = SafeVal(ws.Cells(r, cBuy).Value2)
            price = 0
            If orders > 0 Then price = SafeVal(ws.Cells(r, cAmount).Value2) / orders

            lastDailyRow = lastDailyRow + 1
            wsDaily.Cells(lastDailyRow, 1).Value = art
            wsDaily.Cells(lastDailyRow, 2).Value = wsRes.Cells(rr, 3).Value
            wsDaily.Cells(lastDailyRow, 3).Value = dt
            wsDaily.Cells(lastDailyRow, 4).Value = orders
            wsDaily.Cells(lastDailyRow, 5).Value = buys
            If orders > 0 Then wsDaily.Cells(lastDailyRow, 6).Value = buys / orders
            wsDaily.Cells(lastDailyRow, 7).Value = price
            wsDaily.Cells(lastDailyRow, 8).Value = 0
            wsDaily.Cells(lastDailyRow, 9).Value = "NO"
        End If
    Next r
End Sub
Private Sub AppendStockDaily(ByVal wsRes As Worksheet, ByRef lastResultRow As Long, _
                             ByVal wsDaily As Worksheet, ByRef lastDailyRow As Long)
    Dim ws As Worksheet, r As Long, c As Long, lastRow As Long, lastCol As Long, rr As Long
    Dim art As String, subjectName As String, dt As Date, stock As Double
    Const cArt As Long = 1, cName As Long = 2, cWB As Long = 3, cSubject As Long = 4

    Set ws = gModelWb.Worksheets(SH_STOCK)
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)
    If lastRow < 3 Then Exit Sub

    For r = 3 To lastRow
        art = Txt(ws.Cells(r, cArt).Value2)
        If Len(art) > 0 Then
            rr = EnsureResultArticle(wsRes, lastResultRow, art, _
                                     Txt(ws.Cells(r, cWB).Value2), _
                                     Txt(ws.Cells(r, cName).Value2))

            subjectName = Txt(ws.Cells(r, cSubject).Value2)
            If Len(Txt(wsRes.Cells(rr, 51).Value2)) = 0 And Len(subjectName) > 0 Then
                wsRes.Cells(rr, 51).Value2 = subjectName
            End If

            For c = 8 To lastCol
                dt = ParseDateValue(ws.Cells(2, c).Value)
                If dt <> 0 Then
                    stock = SafeVal(ws.Cells(r, c).Value2)
                    lastDailyRow = lastDailyRow + 1
                    wsDaily.Cells(lastDailyRow, 1).Value = art
                    wsDaily.Cells(lastDailyRow, 2).Value = wsRes.Cells(rr, 3).Value
                    wsDaily.Cells(lastDailyRow, 3).Value = dt
                    wsDaily.Cells(lastDailyRow, 4).Value = 0
                    wsDaily.Cells(lastDailyRow, 5).Value = 0
                    wsDaily.Cells(lastDailyRow, 7).Value = 0
                    wsDaily.Cells(lastDailyRow, 8).Value = stock
                    wsDaily.Cells(lastDailyRow, 9).Value = "YES"
                End If
            Next c
        End If
    Next r
End Sub
Private Sub AggregateStockDetail(ByVal wsRes As Worksheet, ByRef lastResultRow As Long)
    Dim ws As Worksheet, r As Long, lastRow As Long, rr As Long
    Dim art As String, subjectName As String
    Const cArt As Long = 1, cName As Long = 2, cWB As Long = 3, cSubject As Long = 4
    Const cStock As Long = 20, cPrice As Long = 27

    Set ws = gModelWb.Worksheets(SH_STOCKDET)
    lastRow = LastUsedRow(ws)
    If lastRow < 3 Then Exit Sub

    For r = 3 To lastRow
        art = Txt(ws.Cells(r, cArt).Value2)
        If Len(art) > 0 Then
            rr = EnsureResultArticle(wsRes, lastResultRow, art, _
                                     Txt(ws.Cells(r, cWB).Value2), _
                                     Txt(ws.Cells(r, cName).Value2))

            subjectName = Txt(ws.Cells(r, cSubject).Value2)
            If Len(Txt(wsRes.Cells(rr, 51).Value2)) = 0 And Len(subjectName) > 0 Then
                wsRes.Cells(rr, 51).Value2 = subjectName
            End If

            If SafeVal(ws.Cells(r, cPrice).Value2) > 0 Then
                wsRes.Cells(rr, 45).Value = SafeVal(ws.Cells(r, cPrice).Value2)
            End If
            wsRes.Cells(rr, 46).Value = SafeVal(wsRes.Cells(rr, 46).Value2) + _
                                        SafeVal(ws.Cells(r, cStock).Value2)
        End If
    Next r
End Sub
Private Sub EnsureCostRowsNoDictionary(ByVal wsRes As Worksheet, ByVal lastResultRow As Long)
    Dim ws As Worksheet, rr As Long, cr As Long, nextRow As Long
    Dim art As String
    Set ws = gModelWb.Worksheets(SH_COSTS)
    nextRow = LastUsedRow(ws) + 1
    If nextRow < 2 Then nextRow = 2
    For rr = 2 To lastResultRow
        art = Txt(wsRes.Cells(rr, 1).Value2)
        cr = FindCostRow(art)
        If cr = 0 Then
            ws.Cells(nextRow, 1).NumberFormat = "@"
            ws.Cells(nextRow, 1).Value2 = art
            ws.Cells(nextRow, 2).Value2 = wsRes.Cells(rr, 2).Value2
            ws.Cells(nextRow, 3).Value2 = wsRes.Cells(rr, 3).Value2
            ws.Cells(nextRow, 5).Value2 = 0
            ws.Cells(nextRow, 8).Value2 = 0
            ws.Cells(nextRow, 10).Value2 = 0
            ws.Cells(nextRow, 16).Value2 = "YES"
            ws.Cells(nextRow, 17).Value2 = "No cost in the cost master"
            ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 17)).Interior.Color = RGB(255, 249, 196)
            nextRow = nextRow + 1
        Else
            ws.Cells(cr, 2).Value2 = wsRes.Cells(rr, 2).Value2
            ws.Cells(cr, 3).Value2 = wsRes.Cells(rr, 3).Value2
        End If
    Next rr
End Sub

Private Function FindCostRow(ByVal art As String) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = gModelWb.Worksheets(SH_COSTS)
    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then Exit Function
    For r = 2 To lastRow
        If StrComp(Txt(ws.Cells(r, 1).Value2), art, vbTextCompare) = 0 Then
            FindCostRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub SortResultsByArticle(ByVal ws As Worksheet, ByVal lastRow As Long)
    On Error GoTo EH
    If lastRow < 3 Then Exit Sub

5510 ws.Range("A1:BI" & lastRow).Sort _
        Key1:=ws.Range("A2"), _
        Order1:=xlAscending, _
        Header:=xlYes
5520 Exit Sub
EH:
    Err.Raise Err.Number, "SortResultsByArticle line " & Erl, Err.Description
End Sub

Private Sub SortDailyByArticleDate(ByVal ws As Worksheet, ByVal lastRow As Long)
    On Error GoTo EH
    If lastRow < 3 Then Exit Sub

5610 ws.Range("A1:I" & lastRow).Sort _
        Key1:=ws.Range("A2"), _
        Order1:=xlAscending, _
        Key2:=ws.Range("C2"), _
        Order2:=xlAscending, _
        Header:=xlYes
5620 Exit Sub
EH:
    Err.Raise Err.Number, "SortDailyByArticleDate line " & Erl, Err.Description
End Sub

Private Function ConsolidateDailyRows(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim readRow As Long, writeRow As Long, groupRow As Long
    Dim art As String, currentArt As String, itemName As String
    Dim dateKey As Double, currentDateKey As Double
    Dim orders As Double, buys As Double
    Dim priceSum As Double, priceWeight As Double
    Dim stock As Double, known As Boolean
    Dim rowOrders As Double, rowPrice As Double
    Dim lastOutputRow As Long

    On Error GoTo EH
    If lastRow < 2 Then
        ConsolidateDailyRows = 1
        Exit Function
    End If

5810 readRow = 2
5820 writeRow = 2

5830 Do While readRow <= lastRow
5840     art = Txt(ws.Cells(readRow, 1).Value2)
5850     itemName = Txt(ws.Cells(readRow, 2).Value2)
5860     dateKey = SafeVal(ws.Cells(readRow, 3).Value2)
5870     orders = 0
5880     buys = 0
5890     priceSum = 0
5900     priceWeight = 0
5910     stock = 0
5920     known = False
5930     groupRow = readRow

5940     Do While groupRow <= lastRow
5950         currentArt = Txt(ws.Cells(groupRow, 1).Value2)
5960         currentDateKey = SafeVal(ws.Cells(groupRow, 3).Value2)
5970         If StrComp(currentArt, art, vbTextCompare) <> 0 Then Exit Do
5980         If Abs(currentDateKey - dateKey) > 0.0000001 Then Exit Do

5990         rowOrders = SafeVal(ws.Cells(groupRow, 4).Value2)
6000         rowPrice = SafeVal(ws.Cells(groupRow, 7).Value2)
6010         orders = orders + rowOrders
6020         buys = buys + SafeVal(ws.Cells(groupRow, 5).Value2)

6030         If rowPrice > 0 Then
6040             If rowOrders > 0 Then
6050                 priceSum = priceSum + rowPrice * rowOrders
6060                 priceWeight = priceWeight + rowOrders
6070             ElseIf priceWeight = 0 Then
6080                 priceSum = rowPrice
6090                 priceWeight = 1
6100             End If
6110         End If

6120         If UCase$(Txt(ws.Cells(groupRow, 9).Value2)) = "YES" Then
6130             known = True
6140             stock = stock + SafeVal(ws.Cells(groupRow, 8).Value2)
6150         End If
6160         groupRow = groupRow + 1
6170     Loop

6180     ws.Cells(writeRow, 1).Value2 = art
6190     ws.Cells(writeRow, 2).Value2 = itemName
6200     ws.Cells(writeRow, 3).Value2 = dateKey
6210     ws.Cells(writeRow, 4).Value2 = orders
6220     ws.Cells(writeRow, 5).Value2 = buys
6230     If orders > 0 Then
6240         ws.Cells(writeRow, 6).Value2 = buys / orders
6250     Else
6260         ws.Cells(writeRow, 6).ClearContents
6270     End If
6280     If priceWeight > 0 Then
6290         ws.Cells(writeRow, 7).Value2 = priceSum / priceWeight
6300     Else
6310         ws.Cells(writeRow, 7).Value2 = 0
6320     End If
6330     ws.Cells(writeRow, 8).Value2 = stock
6340     If known Then
6350         ws.Cells(writeRow, 9).Value2 = "YES"
6360     Else
6370         ws.Cells(writeRow, 9).Value2 = "NO"
6380     End If

6390     writeRow = writeRow + 1
6400     readRow = groupRow
6410 Loop

6420 lastOutputRow = writeRow - 1
6430 If lastOutputRow < lastRow Then
6440     ws.Range(ws.Cells(lastOutputRow + 1, 1), ws.Cells(lastRow, 9)).ClearContents
6450 End If
6460 ConsolidateDailyRows = lastOutputRow
6470 Exit Function
EH:
    Err.Raise Err.Number, "ConsolidateDailyRows line " & Erl, Err.Description
End Function

Private Sub CalculateResultRows(ByVal wsRes As Worksheet, ByVal lastResultRow As Long, _
                                ByVal wsHist As Worksheet, ByVal wsDaily As Worksheet, ByVal lastDailyRow As Long)
    Dim rr As Long, dStart As Long, dEnd As Long, pointer As Long, histRow As Long
    Dim art As String
    Dim rowErrNum As Long, rowErrDesc As String

    pointer = 2
    histRow = 2

    For rr = 2 To lastResultRow
        On Error GoTo RowError

        art = Txt(wsRes.Cells(rr, 1).Value2)

        Do While pointer <= lastDailyRow And _
                 StrComp(Txt(wsDaily.Cells(pointer, 1).Value2), art, vbTextCompare) < 0
            pointer = pointer + 1
        Loop

        dStart = pointer

        Do While pointer <= lastDailyRow And _
                 StrComp(Txt(wsDaily.Cells(pointer, 1).Value2), art, vbTextCompare) = 0
            pointer = pointer + 1
        Loop

        dEnd = pointer - 1

        CalculateOneResultRow wsRes, rr, wsHist, histRow, wsDaily, dStart, dEnd

ContinueRow:
        On Error GoTo 0
    Next rr

    Exit Sub

RowError:
    rowErrNum = Err.Number
    rowErrDesc = Err.Description

    On Error Resume Next
    wsRes.Cells(rr, 35).Value2 = "??????"
    wsRes.Cells(rr, 36).Value2 = "Calculation skipped: " & rowErrNum & " - " & rowErrDesc
    WriteLog "Article calculation error", "Article " & art & ": " & rowErrNum & " - " & rowErrDesc
    Err.Clear
    On Error GoTo 0

    Resume ContinueRow
End Sub

Private Sub CalculateOneResultRow(ByVal wsRes As Worksheet, ByVal rr As Long, _
                                  ByVal wsHist As Worksheet, ByRef histRow As Long, _
                                  ByVal wsDaily As Worksheet, ByVal dStart As Long, ByVal dEnd As Long)
    Dim s As Worksheet, cws As Worksheet, cr As Long, mode As String, active As String, art As String
    Dim cost As Double, comm As Double, tax As Double, logistics As Double, warehouseDelivery As Double
    Dim currentPrice As Double, actualComm As Double, actualLog As Double, reportPrice As Double
    Dim actualStorage As Double, actualAccept As Double, actualHold As Double, reportServices As Double
    Dim latestFunnelPrice As Double, latestFinancePrice As Double, latestPriceDate As Date
    Dim buyRate As Double, avgOrders As Double
    Dim dataDays As Long, priceGroups As Long, elasticity As Double
    Dim rawElasticity As Double, rawReliable As Boolean, minPriceSpread As Double
    Dim stockAvailability As Double, lastStock As Double
    Dim r As Long, dt As Date, latestDate As Date, orders As Double, buys As Double, price As Double, stock As Double, known As Boolean
    Dim totalOrders As Double, totalBuys As Double, recentOrders As Double
    Dim x As Double, y As Double, sx As Double, sy As Double, sxx As Double, sxy As Double, nReg As Long
    Dim stepPrice As Double, bucket As Double, bi As Long, bucketCount As Long, rowCapacity As Long
    Dim bPrice() As Double, bDays() As Double, bOrders() As Double, bBuys() As Double, bPriceSum() As Double
    Dim stockKnown As Long, stockPositive As Long, arrFBS As Variant, arrFBW As Variant
    Dim fixedFBS As Double, fixedFBW As Double, breakFBS As Double, breakFBW As Double
    Dim selectedBest As Double, selectedCurrentProfit As Double, selectedProfit As Double, uplift As Double
    Dim confidence As String, status As String, reportMode As String, denom As Double
    Dim maxHistPrice As Double, minHistPrice As Double
    Dim finSales As Double, finPriceSum As Double, finCommSum As Double, finLog As Double
    Dim finStorage As Double, finAccept As Double, finHold As Double
    Dim errNum As Long, errDesc As String, errLine As Long
    On Error GoTo EH

7000 Set s = gModelWb.Worksheets(SH_SETTINGS)
7010 Set cws = gModelWb.Worksheets(SH_COSTS)
7020 art = Txt(wsRes.Cells(rr, 1).Value2)
7030 cr = FindCostRow(art)
7040 If cr = 0 Then Exit Sub
7050 active = UCase$(Txt(cws.Cells(cr, 16).Value2))
7060 If active = "NO" Then Exit Sub

7070 cost = SafeVal(cws.Cells(cr, 5).Value2)
7080 warehouseDelivery = SafeVal(cws.Cells(cr, 8).Value2)
7090 If warehouseDelivery < 0 Then warehouseDelivery = 0
7100 tax = SafeVal(s.Range("B3").Value2)
7110 stepPrice = SafeVal(s.Range("B13").Value2)
7120 If stepPrice <= 0 Then stepPrice = 50

7130 finSales = SafeVal(wsRes.Cells(rr, 37).Value2)
7140 finPriceSum = SafeVal(wsRes.Cells(rr, 38).Value2)
7150 finCommSum = SafeVal(wsRes.Cells(rr, 39).Value2)
7160 finLog = SafeVal(wsRes.Cells(rr, 40).Value2)
7170 finStorage = SafeVal(wsRes.Cells(rr, 42).Value2)
7180 finAccept = SafeVal(wsRes.Cells(rr, 43).Value2)
7190 finHold = SafeVal(wsRes.Cells(rr, 44).Value2)
7200 latestFinancePrice = SafeVal(wsRes.Cells(rr, 49).Value2)

7210 If finSales > 0 Then
7220 actualComm = Abs(finCommSum / finSales) / 100
7230 actualLog = Abs(finLog) / finSales
7240 actualStorage = Abs(finStorage) / finSales
7250 actualAccept = Abs(finAccept) / finSales
7260 actualHold = Abs(finHold) / finSales
7270 End If
7280 comm = actualComm
7290 logistics = actualLog
7300 reportServices = actualStorage + actualAccept + actualHold

7310 reportPrice = SafeVal(wsRes.Cells(rr, 45).Value2)
7320 reportMode = UCase$(Txt(wsRes.Cells(rr, 47).Value2))
7330 If InStr(1, reportMode, "FBW", vbTextCompare) > 0 Then
7340 mode = "FBW"
7350 ElseIf InStr(1, reportMode, "FBS", vbTextCompare) > 0 Then
7360 mode = "FBS"
7370 ElseIf reportPrice > 0 Then
7380 mode = "FBW"
7390 Else
7400 mode = ""
7410 End If
7420 cws.Cells(cr, 4).Value2 = mode

7430 If dEnd >= dStart And dStart >= 2 Then
7440 rowCapacity = dEnd - dStart + 1
7450 If rowCapacity < 1 Then rowCapacity = 1
7460 If rowCapacity > 100000 Then rowCapacity = 100000
7470 ReDim bPrice(1 To rowCapacity)
7480 ReDim bDays(1 To rowCapacity)
7490 ReDim bOrders(1 To rowCapacity)
7500 ReDim bBuys(1 To rowCapacity)
7510 ReDim bPriceSum(1 To rowCapacity)
7520 For r = dStart To dEnd
7530 dt = ParseDateValue(wsDaily.Cells(r, 3).Value)
7540 orders = SafeVal(wsDaily.Cells(r, 4).Value2)
7550 buys = SafeVal(wsDaily.Cells(r, 5).Value2)
7560 price = SafeVal(wsDaily.Cells(r, 7).Value2)
7570 stock = SafeVal(wsDaily.Cells(r, 8).Value2)
7580 known = UCase$(Txt(wsDaily.Cells(r, 9).Value2)) = "YES"
7590 If dt > latestDate Then latestDate = dt
7600 If price > 0 And dt >= latestPriceDate Then
7610 latestPriceDate = dt
7620 latestFunnelPrice = price
7630 End If
7640 If orders <> 0 Or buys <> 0 Or price <> 0 Then
7650 dataDays = dataDays + 1
7660 totalOrders = totalOrders + orders
7670 totalBuys = totalBuys + buys
7680 End If
7690 If known Then
7700 stockKnown = stockKnown + 1
7710 If stock > 0 Then stockPositive = stockPositive + 1
7720 lastStock = stock
7730 End If
7740 If price > 0 And price < 100000000# Then
7750 If minHistPrice = 0 Or price < minHistPrice Then minHistPrice = price
7760 If price > maxHistPrice Then maxHistPrice = price
7770 bucket = Int(price / stepPrice + 0.5) * stepPrice
7780 bi = FindBucketIndex(bPrice, bucketCount, bucket)
7790 If bi = 0 And bucketCount < rowCapacity Then
7800 bucketCount = bucketCount + 1
7810 bi = bucketCount
7820 bPrice(bi) = bucket
7830 End If
7840 If bi > 0 Then
7850 bDays(bi) = bDays(bi) + 1
7860 bOrders(bi) = bOrders(bi) + orders
7870 bBuys(bi) = bBuys(bi) + buys
7880 bPriceSum(bi) = bPriceSum(bi) + price
7890 End If
7900 If orders > 0 And (Not known Or stock > 0) Then
7910 x = Log(price)
7920 If orders < 0.5 Then y = Log(0.5) Else y = Log(orders)
7930 sx = sx + x
7940 sy = sy + y
7950 sxx = sxx + x * x
7960 sxy = sxy + x * y
7970 nReg = nReg + 1
7980 End If
7990 End If
8000 Next r
8010 If latestDate <> 0 Then
8020 For r = dStart To dEnd
8030 dt = ParseDateValue(wsDaily.Cells(r, 3).Value)
8040 If dt >= latestDate - 29 And dt <= latestDate Then recentOrders = recentOrders + SafeVal(wsDaily.Cells(r, 4).Value2)
8050 Next r
8060 avgOrders = recentOrders / 30
8070 End If
8080 End If

8090 If reportPrice > 0 Then
8100 currentPrice = reportPrice
8110 ElseIf latestFunnelPrice > 0 Then
8120 currentPrice = latestFunnelPrice
8130 ElseIf latestFinancePrice > 0 Then
8140 currentPrice = latestFinancePrice
8150 ElseIf finSales > 0 Then
8160 currentPrice = finPriceSum / finSales
8170 End If

8180 If avgOrders <= 0 And dataDays > 0 Then avgOrders = totalOrders / dataDays
8190 If totalOrders > 0 Then buyRate = totalBuys / totalOrders Else buyRate = 0
8200 If buyRate < 0 Then buyRate = 0
8210 If buyRate > 1 Then buyRate = 1
8220 If stockKnown > 0 Then stockAvailability = stockPositive / stockKnown
8230 If SafeVal(wsRes.Cells(rr, 46).Value2) <> 0 Then lastStock = SafeVal(wsRes.Cells(rr, 46).Value2)
8240 priceGroups = bucketCount

8250 elasticity = SafeVal(s.Range("B17").Value2)
8260 rawElasticity = 0
8270 rawReliable = False
8280 minPriceSpread = SafeVal(s.Range("B30").Value2)
8290 denom = nReg * sxx - sx * sx
8300 If nReg >= SafeVal(s.Range("B20").Value2) And _
         priceGroups >= SafeVal(s.Range("B21").Value2) And _
         Abs(denom) > 0.0000001 And _
         minHistPrice > 0 And _
         maxHistPrice >= minHistPrice * (1 + minPriceSpread) Then
8310 rawElasticity = (nReg * sxy - sx * sy) / denom
8320 If rawElasticity < 0 Then rawReliable = True
8330 End If
8340 If rawReliable Then
8350 elasticity = ClampDouble(rawElasticity, SafeVal(s.Range("B18").Value2), SafeVal(s.Range("B19").Value2))
8360 Else
8370 elasticity = SafeVal(s.Range("B17").Value2)
8380 End If

8390 fixedFBS = cost + logistics + warehouseDelivery + reportServices
8400 fixedFBW = fixedFBS
8410 denom = 1 - comm - tax
8420 If denom > 0 Then
8430 breakFBS = fixedFBS / denom
8440 breakFBW = fixedFBW / denom
8450 End If

8460 status = "OK"
8470 If cost <= 0 Then status = "No cost in cost master"
8480 If currentPrice <= 0 Then status = "No current price in WB reports"
8490 If actualComm <= 0 Then status = "No commission in financial report"
8500 If actualLog <= 0 Then status = "No logistics in financial report"
8510 If avgOrders <= 0 Or buyRate <= 0 Then status = "Insufficient demand data"
8520 If cost > 0 And currentPrice > 0 And actualComm > 0 And actualLog > 0 And avgOrders > 0 And buyRate > 0 Then
8530 arrFBS = BestPrice(currentPrice, breakFBS, fixedFBS, comm, tax, 0, buyRate, avgOrders, elasticity, maxHistPrice)
8540 arrFBW = BestPrice(currentPrice, breakFBW, fixedFBW, comm, tax, 0, buyRate, avgOrders, elasticity, maxHistPrice)
8550 Else
8560 arrFBS = Array(0#, 0#, 0#, 0#)
8570 arrFBW = Array(0#, 0#, 0#, 0#)
8580 End If

8590 If mode = "FBW" Then
8600 selectedBest = arrFBW(0)
8610 selectedCurrentProfit = arrFBW(3)
8620 selectedProfit = arrFBW(2)
8630 Else
8640 selectedBest = arrFBS(0)
8650 selectedCurrentProfit = arrFBS(3)
8660 selectedProfit = arrFBS(2)
8670 End If
8680 uplift = selectedProfit - selectedCurrentProfit
8690 confidence = ConfidenceLabel(dataDays, priceGroups, stockAvailability, minHistPrice, maxHistPrice, cost)

8700 With wsRes
8710 .Cells(rr, 4).Value2 = mode
8720 .Cells(rr, 5).Value2 = currentPrice
8730 .Cells(rr, 6).Value2 = cost
8740 .Cells(rr, 7).Value2 = comm
8750 .Cells(rr, 8).Value2 = tax
8760 .Cells(rr, 9).Value2 = logistics
8770 .Cells(rr, 10).Value2 = buyRate
8780 .Cells(rr, 11).Value2 = avgOrders
8790 .Cells(rr, 12).Value2 = dataDays
8800 .Cells(rr, 13).Value2 = priceGroups
8810 .Cells(rr, 14).Value2 = elasticity
8820 .Cells(rr, 15).Value2 = stockAvailability
8830 .Cells(rr, 16).Value2 = lastStock
8840 .Cells(rr, 17).Value2 = breakFBS
8850 .Cells(rr, 18).Value2 = arrFBS(0)
8860 .Cells(rr, 19).Value2 = arrFBS(1)
8870 .Cells(rr, 20).Value2 = arrFBS(3)
8880 .Cells(rr, 21).Value2 = arrFBS(2)
8890 .Cells(rr, 22).Value2 = breakFBW
8900 .Cells(rr, 23).Value2 = arrFBW(0)
8910 .Cells(rr, 24).Value2 = arrFBW(1)
8920 .Cells(rr, 25).Value2 = arrFBW(3)
8930 .Cells(rr, 26).Value2 = arrFBW(2)
8940 .Cells(rr, 27).Value2 = selectedBest
8950 .Cells(rr, 28).Value2 = selectedCurrentProfit
8960 .Cells(rr, 29).Value2 = selectedProfit
8970 .Cells(rr, 30).Value2 = uplift
8980 .Cells(rr, 31).Value2 = actualComm
8990 .Cells(rr, 32).Value2 = actualLog
9000 .Cells(rr, 33).Value2 = currentPrice
9010 .Cells(rr, 34).Value2 = reportMode
9020 .Cells(rr, 35).Value2 = confidence
9030 .Cells(rr, 36).Value2 = status
9040 .Cells(rr, 50).Value2 = reportServices
         .Cells(rr, 52).Value2 = rawElasticity
         .Cells(rr, 55).Value2 = elasticity
         .Cells(rr, 58).Value2 = nReg
         .Cells(rr, 59).Value2 = 0
         .Cells(rr, 60).Value2 = maxHistPrice
         If rawReliable Then
             .Cells(rr, 61).Value2 = "??"
         Else
             .Cells(rr, 61).Value2 = "???"
         End If
9050 End With

9060 For bi = 1 To bucketCount
9070 wsHist.Cells(histRow, 1).Value2 = art
9080 wsHist.Cells(histRow, 2).Value2 = wsRes.Cells(rr, 3).Value2
9090 wsHist.Cells(histRow, 3).Value2 = bPrice(bi)
9100 wsHist.Cells(histRow, 4).Value2 = bDays(bi)
9110 wsHist.Cells(histRow, 5).Value2 = bOrders(bi)
9120 If bDays(bi) > 0 Then
9130 wsHist.Cells(histRow, 6).Value2 = bOrders(bi) / bDays(bi)
9140 wsHist.Cells(histRow, 7).Value2 = bBuys(bi) / bDays(bi)
9150 wsHist.Cells(histRow, 8).Value2 = bPriceSum(bi) / bDays(bi)
9160 End If
9170 histRow = histRow + 1
9180 Next bi
9190 Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl
    Err.Raise errNum, "CalculateOneResultRow article " & art & " line " & errLine, errDesc
End Sub

Private Sub ApplyHierarchicalElasticityAndReprice(ByVal wsRes As Worksheet, ByVal lastResultRow As Long)
    Dim s As Worksheet
    Dim rr As Long, excludeRow As Long
    Dim subjectName As String, sourceText As String
    Dim rawElasticity As Double, subjectElasticity As Double
    Dim portfolioElasticity As Double, baseElasticity As Double
    Dim appliedElasticity As Double, individualWeight As Double
    Dim minElasticity As Double, maxElasticity As Double, fallbackElasticity As Double
    Dim shrinkageK As Double
    Dim minSubjectCount As Long, minPortfolioCount As Long
    Dim subjectCount As Long, portfolioCount As Long
    Dim displayedSubjectCount As Long, displayedPortfolioCount As Long
    Dim rawReliable As Boolean, subjectAvailable As Boolean, portfolioAvailable As Boolean
    Dim useHierarchy As Boolean
    Dim errNum As Long, errDesc As String, errLine As Long

    On Error GoTo EH
    Set s = gModelWb.Worksheets(SH_SETTINGS)

    With wsRes
        .Range("AZ:BC").NumberFormat = "0.00"
        .Range("BE:BF").NumberFormat = "0"
        .Range("BG:BG").NumberFormat = "0.0%"
        .Range("BH:BH").NumberFormat = "#,##0 ""?"""
        .Columns("AY:AY").ColumnWidth = 25
        .Columns("AZ:BC").ColumnWidth = 20
        .Columns("BD:BD").ColumnWidth = 34
        .Columns("BE:BF").ColumnWidth = 18
        .Columns("BG:BG").ColumnWidth = 18
        .Columns("BH:BH").ColumnWidth = 22
        .Columns("BI:BI").ColumnWidth = 20
    End With

    fallbackElasticity = SafeVal(s.Range("B17").Value2)
    minElasticity = SafeVal(s.Range("B18").Value2)
    maxElasticity = SafeVal(s.Range("B19").Value2)
    minSubjectCount = CLng(SafeVal(s.Range("B27").Value2))
    minPortfolioCount = CLng(SafeVal(s.Range("B28").Value2))
    shrinkageK = SafeVal(s.Range("B29").Value2)
    useHierarchy = (SafeVal(s.Range("B31").Value2) <> 0)

    If minSubjectCount < 1 Then minSubjectCount = 3
    If minPortfolioCount < 1 Then minPortfolioCount = 5
    If shrinkageK < 0 Then shrinkageK = 0
    If fallbackElasticity >= 0 Then fallbackElasticity = -1.5
    If minElasticity >= maxElasticity Then
        minElasticity = -4
        maxElasticity = -0.2
    End If

    For rr = 2 To lastResultRow
        rawReliable = (UCase$(Txt(wsRes.Cells(rr, 61).Value2)) = "??")
        If rawReliable Then
            excludeRow = rr
        Else
            excludeRow = 0
        End If

        subjectName = Txt(wsRes.Cells(rr, 51).Value2)

        subjectCount = 0
        subjectElasticity = AverageReliableElasticity( _
            wsRes, lastResultRow, subjectName, True, excludeRow, _
            minElasticity, maxElasticity, subjectCount)
        displayedSubjectCount = subjectCount
        If rawReliable Then displayedSubjectCount = displayedSubjectCount + 1
        subjectAvailable = (Len(subjectName) > 0 And subjectCount > 0 And _
                            displayedSubjectCount >= minSubjectCount)

        portfolioCount = 0
        portfolioElasticity = AverageReliableElasticity( _
            wsRes, lastResultRow, "", False, excludeRow, _
            minElasticity, maxElasticity, portfolioCount)
        displayedPortfolioCount = portfolioCount
        If rawReliable Then displayedPortfolioCount = displayedPortfolioCount + 1
        portfolioAvailable = (portfolioCount > 0 And _
                              displayedPortfolioCount >= minPortfolioCount)

        If subjectAvailable Then
            baseElasticity = subjectElasticity
        ElseIf portfolioAvailable Then
            baseElasticity = portfolioElasticity
        Else
            baseElasticity = fallbackElasticity
        End If

        If Not useHierarchy Then
            If rawReliable Then
                appliedElasticity = ClampDouble(SafeVal(wsRes.Cells(rr, 52).Value2), minElasticity, maxElasticity)
                sourceText = "?????????????? ??????"
                individualWeight = 1
            Else
                appliedElasticity = fallbackElasticity
                sourceText = "????????? ???????? ??????"
                individualWeight = 0
            End If
        ElseIf rawReliable Then
            rawElasticity = ClampDouble(SafeVal(wsRes.Cells(rr, 52).Value2), minElasticity, maxElasticity)
            If shrinkageK = 0 Then
                individualWeight = 1
            Else
                individualWeight = SafeVal(wsRes.Cells(rr, 58).Value2) / _
                                   (SafeVal(wsRes.Cells(rr, 58).Value2) + shrinkageK)
            End If
            individualWeight = ClampDouble(individualWeight, 0, 1)
            appliedElasticity = individualWeight * rawElasticity + _
                                (1 - individualWeight) * baseElasticity

            If subjectAvailable Then
                sourceText = "??????? + ??????? ?? ????????"
            ElseIf portfolioAvailable Then
                sourceText = "??????? + ??????? ?? ????????????"
            Else
                sourceText = "??????? + ????????? ???????? ??????"
            End If
        Else
            individualWeight = 0
            appliedElasticity = baseElasticity
            If subjectAvailable Then
                sourceText = "??????? ?? ????????"
            ElseIf portfolioAvailable Then
                sourceText = "??????? ?? ????????????"
            Else
                sourceText = "????????? ???????? ??????"
            End If
        End If

        appliedElasticity = ClampDouble(appliedElasticity, minElasticity, maxElasticity)

        If subjectAvailable Then
            wsRes.Cells(rr, 53).Value2 = subjectElasticity
        Else
            wsRes.Cells(rr, 53).ClearContents
        End If
        If portfolioAvailable Then
            wsRes.Cells(rr, 54).Value2 = portfolioElasticity
        Else
            wsRes.Cells(rr, 54).ClearContents
        End If
        wsRes.Cells(rr, 55).Value2 = appliedElasticity
        wsRes.Cells(rr, 56).Value2 = sourceText
        wsRes.Cells(rr, 57).Value2 = displayedSubjectCount
        wsRes.Cells(rr, 59).Value2 = individualWeight
        wsRes.Cells(rr, 14).Value2 = appliedElasticity

        If InStr(1, sourceText, "?????????", vbTextCompare) > 0 Then
            wsRes.Cells(rr, 35).Value2 = "??????"
        ElseIf Not rawReliable Then
            wsRes.Cells(rr, 35).Value2 = "???????"
        End If

        RepriceResultRow wsRes, rr, appliedElasticity
    Next rr

    Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl
    Err.Raise errNum, "ApplyHierarchicalElasticityAndReprice line " & errLine, errDesc
End Sub

Private Function AverageReliableElasticity(ByVal wsRes As Worksheet, ByVal lastResultRow As Long, _
                                           ByVal subjectFilter As String, ByVal filterBySubject As Boolean, _
                                           ByVal excludeRow As Long, ByVal minElasticity As Double, _
                                           ByVal maxElasticity As Double, ByRef reliableCount As Long) As Double
    Dim r As Long
    Dim e As Double, sumElasticity As Double
    Dim rowSubject As String

    reliableCount = 0
    sumElasticity = 0

    For r = 2 To lastResultRow
        If r <> excludeRow Then
            If UCase$(Txt(wsRes.Cells(r, 61).Value2)) = "??" Then
                If filterBySubject Then
                    rowSubject = Txt(wsRes.Cells(r, 51).Value2)
                    If StrComp(rowSubject, subjectFilter, vbTextCompare) <> 0 Then GoTo ContinueAverage
                End If

                e = SafeVal(wsRes.Cells(r, 52).Value2)
                If e < 0 Then
                    e = ClampDouble(e, minElasticity, maxElasticity)
                    sumElasticity = sumElasticity + e
                    reliableCount = reliableCount + 1
                End If
            End If
        End If
ContinueAverage:
    Next r

    If reliableCount > 0 Then
        AverageReliableElasticity = sumElasticity / reliableCount
    Else
        AverageReliableElasticity = 0
    End If
End Function

Private Sub RepriceResultRow(ByVal wsRes As Worksheet, ByVal rr As Long, ByVal elasticity As Double)
    Dim currentPrice As Double, cost As Double, comm As Double, tax As Double
    Dim logistics As Double, buyRate As Double, avgOrders As Double
    Dim breakFBS As Double, breakFBW As Double, fixedFBS As Double, fixedFBW As Double
    Dim maxHistPrice As Double, denom As Double
    Dim arrFBS As Variant, arrFBW As Variant
    Dim mode As String
    Dim selectedBest As Double, selectedCurrentProfit As Double
    Dim selectedProfit As Double, uplift As Double

    currentPrice = SafeVal(wsRes.Cells(rr, 5).Value2)
    cost = SafeVal(wsRes.Cells(rr, 6).Value2)
    comm = SafeVal(wsRes.Cells(rr, 7).Value2)
    tax = SafeVal(wsRes.Cells(rr, 8).Value2)
    logistics = SafeVal(wsRes.Cells(rr, 9).Value2)
    buyRate = SafeVal(wsRes.Cells(rr, 10).Value2)
    avgOrders = SafeVal(wsRes.Cells(rr, 11).Value2)
    breakFBS = SafeVal(wsRes.Cells(rr, 17).Value2)
    breakFBW = SafeVal(wsRes.Cells(rr, 22).Value2)
    maxHistPrice = SafeVal(wsRes.Cells(rr, 60).Value2)
    mode = UCase$(Txt(wsRes.Cells(rr, 4).Value2))

    denom = 1 - comm - tax
    If denom > 0 Then
        fixedFBS = breakFBS * denom
        fixedFBW = breakFBW * denom
    End If

    If cost > 0 And currentPrice > 0 And comm > 0 And logistics > 0 And _
       avgOrders > 0 And buyRate > 0 Then
        arrFBS = BestPrice(currentPrice, breakFBS, fixedFBS, comm, tax, 0, _
                           buyRate, avgOrders, elasticity, maxHistPrice)
        arrFBW = BestPrice(currentPrice, breakFBW, fixedFBW, comm, tax, 0, _
                           buyRate, avgOrders, elasticity, maxHistPrice)
    Else
        arrFBS = Array(0#, 0#, 0#, 0#)
        arrFBW = Array(0#, 0#, 0#, 0#)
    End If

    wsRes.Cells(rr, 18).Value2 = arrFBS(0)
    wsRes.Cells(rr, 19).Value2 = arrFBS(1)
    wsRes.Cells(rr, 20).Value2 = arrFBS(3)
    wsRes.Cells(rr, 21).Value2 = arrFBS(2)
    wsRes.Cells(rr, 23).Value2 = arrFBW(0)
    wsRes.Cells(rr, 24).Value2 = arrFBW(1)
    wsRes.Cells(rr, 25).Value2 = arrFBW(3)
    wsRes.Cells(rr, 26).Value2 = arrFBW(2)

    If mode = "FBW" Then
        selectedBest = arrFBW(0)
        selectedCurrentProfit = arrFBW(3)
        selectedProfit = arrFBW(2)
    Else
        selectedBest = arrFBS(0)
        selectedCurrentProfit = arrFBS(3)
        selectedProfit = arrFBS(2)
    End If

    uplift = selectedProfit - selectedCurrentProfit
    wsRes.Cells(rr, 27).Value2 = selectedBest
    wsRes.Cells(rr, 28).Value2 = selectedCurrentProfit
    wsRes.Cells(rr, 29).Value2 = selectedProfit
    wsRes.Cells(rr, 30).Value2 = uplift
End Sub

Private Function FindBucketIndex(ByRef prices() As Double, ByVal count As Long, ByVal target As Double) As Long
    Dim i As Long
    For i = 1 To count
        If Abs(prices(i) - target) < 0.0001 Then FindBucketIndex = i: Exit Function
    Next i
End Function

Private Function BestPrice(ByVal currentPrice As Double, ByVal breakEven As Double, ByVal fixedCost As Double, _
                           ByVal comm As Double, ByVal tax As Double, ByVal adRate As Double, _
                           ByVal buyRate As Double, ByVal avgOrders As Double, ByVal elasticity As Double, _
                           ByVal maxHistPrice As Double) As Variant
    Dim s As Worksheet
    Dim stepP As Double, downPct As Double, upPct As Double, daysForecast As Double
    Dim lowP As Double, highP As Double, p As Double, ratio As Double
    Dim orders As Double, buys As Double, demandFactor As Double
    Dim profitUnit As Double, monthly As Double
    Dim bestP As Double, bestMonthly As Double, bestUnit As Double
    Dim currentUnit As Double, currentMonthly As Double
    Dim iterationCount As Long
    Const MAX_PRICE As Double = 100000000#
    Const MAX_ITERATIONS As Long = 10000

    If currentPrice <= 0 Or avgOrders <= 0 Or buyRate <= 0 Then
        BestPrice = Array(0#, 0#, 0#, 0#)
        Exit Function
    End If

    Set s = gModelWb.Worksheets(SH_SETTINGS)

    stepP = SafeVal(s.Range("B13").Value2)
    If stepP <= 0 Then stepP = 50

    downPct = SafeVal(s.Range("B14").Value2)
    upPct = SafeVal(s.Range("B15").Value2)
    daysForecast = SafeVal(s.Range("B16").Value2)
    If daysForecast <= 0 Then daysForecast = 30

    currentPrice = ClampDouble(currentPrice, 0.01, MAX_PRICE)
    breakEven = ClampDouble(breakEven, 0, MAX_PRICE)
    maxHistPrice = ClampDouble(maxHistPrice, 0, MAX_PRICE)
    fixedCost = ClampDouble(fixedCost, -MAX_PRICE, MAX_PRICE)
    avgOrders = ClampDouble(avgOrders, 0, 1000000#)
    buyRate = ClampDouble(buyRate, 0, 1)
    elasticity = ClampDouble(elasticity, -10, 0)

    lowP = Application.Max(breakEven + stepP, currentPrice * (1 - downPct))
    highP = Application.Max(currentPrice * (1 + upPct), maxHistPrice * 1.1)

    lowP = ClampDouble(lowP, stepP, MAX_PRICE)
    highP = ClampDouble(highP, lowP, MAX_PRICE)

    p = Int((lowP + stepP - 0.0001) / stepP) * stepP
    If p < stepP Then p = stepP

    Do While p <= highP + 0.0001 And iterationCount < MAX_ITERATIONS
        iterationCount = iterationCount + 1

        ratio = p / currentPrice
        demandFactor = SafeDemandFactor(ratio, elasticity)

        orders = avgOrders * demandFactor
        orders = Application.Min(avgOrders * 2.5, Application.Max(0, orders))

        buys = orders * buyRate
        profitUnit = p * (1 - comm - tax - adRate) - fixedCost
        monthly = SafeMultiply4(profitUnit, buys, daysForecast, 1#)

        If bestP = 0 Or monthly > bestMonthly Then
            bestP = p
            bestMonthly = monthly
            bestUnit = profitUnit
        End If

        p = p + stepP
    Loop

    currentUnit = currentPrice * (1 - comm - tax - adRate) - fixedCost
    currentMonthly = SafeMultiply4(currentUnit, avgOrders, buyRate, daysForecast)

    BestPrice = Array(bestP, bestUnit, bestMonthly, currentMonthly)
End Function

Private Function SafeDemandFactor(ByVal ratio As Double, ByVal elasticity As Double) As Double
    Dim exponentValue As Double

    If ratio <= 0 Then
        SafeDemandFactor = 0
        Exit Function
    End If

    exponentValue = elasticity * Log(ratio)

    If exponentValue > 50 Then
        exponentValue = 50
    ElseIf exponentValue < -50 Then
        exponentValue = -50
    End If

    SafeDemandFactor = Exp(exponentValue)

    If SafeDemandFactor < 0 Then SafeDemandFactor = 0
    If SafeDemandFactor > 2.5 Then SafeDemandFactor = 2.5
End Function

Private Function SafeMultiply4(ByVal a As Double, ByVal b As Double, _
                               ByVal c As Double, ByVal d As Double) As Double
    Dim resultValue As Double

    On Error GoTo TooLarge
    resultValue = a * b
    resultValue = resultValue * c
    resultValue = resultValue * d

    If resultValue > 1E+100 Then resultValue = 1E+100
    If resultValue < -1E+100 Then resultValue = -1E+100

    SafeMultiply4 = resultValue
    Exit Function

TooLarge:
    If (a < 0) Xor (b < 0) Xor (c < 0) Xor (d < 0) Then
        SafeMultiply4 = -1E+100
    Else
        SafeMultiply4 = 1E+100
    End If
End Function

Private Function ClampDouble(ByVal valueToClamp As Double, ByVal minValue As Double, _
                             ByVal maxValue As Double) As Double
    If valueToClamp < minValue Then
        ClampDouble = minValue
    ElseIf valueToClamp > maxValue Then
        ClampDouble = maxValue
    Else
        ClampDouble = valueToClamp
    End If
End Function

Public Sub BuildScenario()
    Dim art As String, wsR As Worksheet, wsS As Worksheet, r As Long
    Dim currentPrice As Double, comm As Double, tax As Double, logistics As Double
    Dim buyRate As Double, avgOrders As Double, elasticity As Double
    Dim mode As String, cr As Long, cws As Worksheet, settings As Worksheet
    Dim cost As Double, reportServices As Double, warehouseDelivery As Double, fixedCost As Double
    Dim stepP As Double, lowP As Double, highP As Double
    Dim p As Double, orders As Double, buys As Double, profitUnit As Double, monthly As Double, rowOut As Long
    On Error GoTo EH
    Set gModelWb = ActiveWorkbook
    BindModelWorkbook
    Set wsR = gModelWb.Worksheets(SH_RESULTS)
    Set wsS = gModelWb.Worksheets(SH_SCENARIO)
    Set cws = gModelWb.Worksheets(SH_COSTS)
    Set settings = gModelWb.Worksheets(SH_SETTINGS)
    art = Txt(settings.Range("B26").Value2)
    If Len(art) = 0 Then art = Txt(wsS.Range("B2").Value2)
    If Len(art) = 0 Then Exit Sub
    r = FindResultRow(wsR, LastUsedRow(wsR), art)
    If r = 0 Then Exit Sub
    currentPrice = SafeVal(wsR.Cells(r, 5).Value2)
    If currentPrice <= 0 Then Exit Sub
    comm = SafeVal(wsR.Cells(r, 7).Value2)
    tax = SafeVal(wsR.Cells(r, 8).Value2)
    logistics = SafeVal(wsR.Cells(r, 9).Value2)
    buyRate = SafeVal(wsR.Cells(r, 10).Value2)
    avgOrders = SafeVal(wsR.Cells(r, 11).Value2)
    elasticity = SafeVal(wsR.Cells(r, 14).Value2)
    mode = UCase$(Txt(wsR.Cells(r, 4).Value2))
    reportServices = SafeVal(wsR.Cells(r, 50).Value2)
    cr = FindCostRow(art)
    If cr = 0 Then Exit Sub
    cost = SafeVal(cws.Cells(cr, 5).Value2)
    warehouseDelivery = SafeVal(cws.Cells(cr, 8).Value2)
    fixedCost = cost + logistics + warehouseDelivery + reportServices
    wsS.Rows("5:" & wsS.Rows.count).ClearContents
    wsS.Range("B2").Value2 = art
    wsS.Range("D2").Value2 = wsR.Cells(r, 3).Value2
    wsS.Range("F2").Value2 = mode
    stepP = SafeVal(settings.Range("B13").Value2)
    If stepP <= 0 Then stepP = 50
    lowP = currentPrice * (1 - SafeVal(settings.Range("B14").Value2))
    If lowP < 0 Then lowP = 0
    highP = currentPrice * (1 + SafeVal(settings.Range("B15").Value2))
    p = Int((lowP + stepP - 0.0001) / stepP) * stepP
    rowOut = 5
    Do While p <= highP + 0.0001 And rowOut < 10000
        orders = avgOrders * SafeDemandFactor(p / currentPrice, elasticity)
        If orders < 0 Then orders = 0
        If orders > avgOrders * 2.5 Then orders = avgOrders * 2.5
        buys = orders * buyRate
        profitUnit = p * (1 - comm - tax) - fixedCost
        monthly = SafeMultiply4(profitUnit, buys, SafeVal(settings.Range("B16").Value2), 1#)
        wsS.Cells(rowOut, 1).Value2 = p
        wsS.Cells(rowOut, 2).Value2 = orders
        wsS.Cells(rowOut, 3).Value2 = buys
        wsS.Cells(rowOut, 4).Value2 = profitUnit
        wsS.Cells(rowOut, 5).Value2 = monthly
        If Abs(p - currentPrice) < stepP / 2 Then wsS.Cells(rowOut, 6).Value2 = "Current"
        rowOut = rowOut + 1
        p = p + stepP
    Loop
    UpdateScenarioChart rowOut - 1
    Exit Sub
EH:
    WriteLog "Scenario error", Err.Number & " - " & Err.Description
End Sub

Private Sub SortResultsByUplift(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < 3 Then Exit Sub
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("AD2:AD" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:BI" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub UpdateDashboard(ByVal lastResultRow As Long)
    Dim ws As Worksheet
    Dim wsR As Worksheet
    Dim i As Long
    Dim srcRow As Long
    Dim articleCount As Long
    Dim lastDashboardRow As Long
    Dim outputLastRow As Long

    Set ws = gModelWb.Worksheets(SH_DASH)
    Set wsR = gModelWb.Worksheets(SH_RESULTS)

    articleCount = Application.Max(0, lastResultRow - 1)

    ws.Range("B4").Value = articleCount
    ws.Range("E4").Value = Application.CountIf( _
        wsR.Range("AJ2:AJ" & Application.Max(2, lastResultRow)), "*cost*")
    ws.Range("H4").Value = Application.Sum( _
        wsR.Range("AB2:AB" & Application.Max(2, lastResultRow)))
    ws.Range("K4").Value = Application.Sum( _
        wsR.Range("AC2:AC" & Application.Max(2, lastResultRow)))
    ws.Range("N4").Value = ws.Range("K4").Value - ws.Range("H4").Value
    ws.Range("B6").Value = gModelWb.Worksheets(SH_SETTINGS).Range("B24").Value

    lastDashboardRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastDashboardRow < 10 Then lastDashboardRow = 10
    ws.Range("A10:G" & lastDashboardRow).ClearContents

    If articleCount > 0 Then
        outputLastRow = 9 + articleCount
        ws.Range("A10:G10").Copy
        ws.Range("A10:G" & outputLastRow).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        For i = 0 To articleCount - 1
            srcRow = i + 2
            ws.Cells(10 + i, 1).Value = wsR.Cells(srcRow, 1).Value
            ws.Cells(10 + i, 2).Value = wsR.Cells(srcRow, 3).Value
            ws.Cells(10 + i, 3).Value = wsR.Cells(srcRow, 4).Value
            ws.Cells(10 + i, 4).Value = wsR.Cells(srcRow, 5).Value
            ws.Cells(10 + i, 5).Value = wsR.Cells(srcRow, 27).Value
            ws.Cells(10 + i, 6).Value = wsR.Cells(srcRow, 30).Value
            ws.Cells(10 + i, 7).Value = wsR.Cells(srcRow, 36).Value
        Next i

        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.Range("A9:G" & outputLastRow).AutoFilter
    Else
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.Range("A9:G10").AutoFilter
    End If

    UpdateDashboardCharts 9 + articleCount
End Sub
Private Sub UpdateDashboardCharts(ByVal lastTopRow As Long)
    'Charts are intentionally disabled in the compatibility build.
    'All calculations and tables remain available.
End Sub

Private Sub UpdateScenarioChart(ByVal lastRow As Long)
    'Charts are intentionally disabled in the compatibility build.
End Sub

Private Function ConfidenceLabel(ByVal dataDays As Long, ByVal groups As Long, ByVal stockAvail As Double, _
                                 ByVal minPrice As Double, ByVal maxPrice As Double, ByVal cost As Double) As String
    Dim score As Long
    If cost <= 0 Then ConfidenceLabel = "??? ???????": Exit Function
    If dataDays >= 30 Then score = score + 1
    If groups >= 3 Then score = score + 1
    If stockAvail = 0 Or stockAvail >= 0.6 Then score = score + 1
    If minPrice > 0 And maxPrice / minPrice >= 1.05 Then score = score + 1
    If score >= 4 Then
        ConfidenceLabel = "???????"
    ElseIf score >= 2 Then
        ConfidenceLabel = "???????"
    Else
        ConfidenceLabel = "??????"
    End If
End Function


Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then LastUsedRow = 1 Else LastUsedRow = f.Row
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then LastUsedCol = 1 Else LastUsedCol = f.Column
End Function

Private Function SafeVal(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or Len(Txt(v)) = 0 Then
        SafeVal = 0
    ElseIf IsNumeric(v) Then
        SafeVal = CDbl(v)
    Else
        SafeVal = 0
    End If
End Function

Private Function SafeCellVal(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long) As Double
    If colNum <= 0 Then SafeCellVal = 0 Else SafeCellVal = SafeVal(ws.Cells(rowNum, colNum).Value2)
End Function

Private Function NzOrDefault(ByVal v As Variant, ByVal defaultValue As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or Len(Txt(v)) = 0 Then NzOrDefault = SafeVal(defaultValue) Else NzOrDefault = SafeVal(v)
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Txt = "" Else Txt = Trim$(CStr(v))
End Function

Private Function ParseDateValue(ByVal v As Variant) As Date
    On Error GoTo BadDate
    If IsDate(v) Then ParseDateValue = CDate(v): Exit Function
    If IsNumeric(v) And CDbl(v) > 30000 Then ParseDateValue = CDate(CDbl(v)): Exit Function
BadDate:
    ParseDateValue = 0
End Function

Private Sub WriteLog(ByVal actionText As String, ByVal details As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = gModelWb.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If Len(Txt(ws.Cells(1, 1).Value2)) = 0 Then
        ws.Cells(1, 1).Value2 = "Date/time"
        ws.Cells(1, 2).Value2 = "User"
        ws.Cells(1, 3).Value2 = "Action"
        ws.Cells(1, 4).Value2 = "Details"
        ws.Rows(1).Font.Bold = True
    End If
    r = Application.Max(2, LastUsedRow(ws) + 1)
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = Environ$("Username")
    ws.Cells(r, 3).Value = actionText
    ws.Cells(r, 4).Value = details
End Sub
