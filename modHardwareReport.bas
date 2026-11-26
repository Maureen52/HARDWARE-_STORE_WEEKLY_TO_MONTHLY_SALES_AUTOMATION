Attribute VB_Name = "modHardwareReport"
Option Explicit
'=====================================================================
' HARDWARE STORE - WEEKLY -> MONTHLY SALES AUTOMATION
'
' Main macro : RunMonthlyReport   (does everything below in one click)
'   Step 1 Consolidate  Week1..WeekN sheets  -> Monthly_Raw
'   Step 2 Clean        Monthly_Raw          -> Monthly_Clean (table tblSales),
'                                               Rejected_Rows, Cleaning_Log
'   Step 3 Validate     consistency checks   -> Data_Checks
'   Step 4 Pivots       8 pivot tables       -> Pivots
'   Step 5 Dashboard    KPI cards + 6 pivot charts + 4 slicers -> Dashboard
'
' Other macros : SetupControlPanel (creates the Control sheet + buttons)
'                SaveMonthlyArchive, StartNewMonth, ResetSlicers
'
' REQUIRES
'   - Sheets named Week1, Week2, ... with the 10 standard headers in A1:J1
'   - A 'Lookups' sheet (product master + payment / category aliases)
'   - Windows desktop Excel (Scripting.Dictionary, slicers = Excel 2013+)
'
' BUSINESS RULES / ASSUMPTIONS (change them here if the store works differently)
'   - Text dates are read as day/month/year (Kenyan format).
'   - Two rows with the same date, receipt, product, quantity and price are
'     treated as an accidental duplicate and the second one is removed.
'   - Rows with a missing/zero quantity, missing receipt, missing product or
'     unreadable date are REJECTED (listed in Rejected_Rows with the reason).
'   - Negative quantities are treated as possible returns: they are NOT netted
'     off revenue, but their value is reported in Cleaning_Log.
'   - Missing prices are filled from the standard price in Lookups.
'   - A price that is ~10x or ~1/10 of the standard price is treated as a
'     decimal-shift typo and corrected; every correction is listed in Data_Checks.
'   - Other prices outside the flag range are KEPT but flagged for review.
'   - Rows dated outside the reporting month are KEPT but flagged.
'   - Receipts are not counted distinctly (slicer-driven pivots cannot do this),
'     so the dashboard reports line items and average line value.
'
' Written as a portfolio project; sample data is synthetic.
'=====================================================================

' ---- settings you can change ----
Private Const FLAG_PRICE_LOW As Double = 0.5    ' flag prices below 50% of standard
Private Const FLAG_PRICE_HIGH As Double = 2     ' flag prices above 200% of standard
Private Const SHIFT_LOW As Double = 8           ' price/standard between 8 and 12 -> divide by 10
Private Const SHIFT_HIGH As Double = 12         ' price/standard between 1/12 and 1/8 -> multiply by 10
Private Const REPORT_AUTHOR As String = ""       ' type your full name here; if empty, Excel's user name is used

Private Const WEEK_PATTERN As String = "WEEK#*"
Private Const SH_RAW As String = "Monthly_Raw"
Private Const SH_CLEAN As String = "Monthly_Clean"
Private Const SH_REJ As String = "Rejected_Rows"
Private Const SH_LOG As String = "Cleaning_Log"
Private Const SH_CHECK As String = "Data_Checks"
Private Const SH_PIV As String = "Pivots"
Private Const SH_CALC As String = "Calc"
Private Const SH_DASH As String = "Dashboard"
Private Const SH_LOOK As String = "Lookups"
Private Const SH_CTRL As String = "Control"
Private Const TBL As String = "tblSales"
Private Const MAX_ISSUES As Long = 300

' lookup dictionaries (loaded from the Lookups sheet)
Private dProdName As Object, dProdCat As Object, dProdUnit As Object, dProdPrice As Object
Private dPay As Object, dCat As Object

' run counters
Private cRead As Long, cBlank As Long, cRej As Long, cKept As Long
Private cBadDate As Long, cNoReceipt As Long, cNoProduct As Long, cBadQty As Long, cBadPrice As Long, cDup As Long
Private cTextDate As Long, cTextNum As Long, cProdFix As Long, cCatFix As Long, cPayFix As Long
Private cPriceFill As Long, cUnitFill As Long, cCustFill As Long, cSalesFill As Long, cUnmatched As Long
Private cPriceFix As Long, cRejNoQty As Long, gOutOfMonth As Long
Private cRejValue As Double, cRetValue As Double
Private gFix As Collection

Private gStep As String
Private gSlicerStage As String, gSlicerNote As String
Private pvCat As PivotTable, pvProd As PivotTable, pvPay As PivotTable, pvSp As PivotTable
Private pvCust As PivotTable, pvWk As PivotTable, pvDay As PivotTable, pvKpi As PivotTable
Private gMonthStart As Date
Private gCheckRow As Long, gIssueRow As Long, gIssueCount As Long

'=====================================================================
' MAIN ENTRY POINT
'=====================================================================
Public Sub RunMonthlyReport()
    Dim t As Double
    t = Timer
    On Error GoTo Fail
    SpeedUp True
    ResetCounters
    ClearOutputs
    Step1_Consolidate
    Step2_Clean
    Step3_Validate
    Step4_Pivots
    Step5_Dashboard
    ResetSlicers
    SetDocProperties
    SpeedUp False
    ThisWorkbook.Worksheets(SH_DASH).Activate
    MsgBox "Monthly report for " & Format$(gMonthStart, "mmmm yyyy") & " is ready." & vbCrLf & vbCrLf & _
           "Rows read: " & cRead & vbCrLf & _
           "Clean rows kept: " & cKept & vbCrLf & _
           "Rows removed (blank/invalid/duplicate): " & (cBlank + cRej) & vbCrLf & _
           "Price typos auto-corrected: " & cPriceFix & vbCrLf & vbCrLf & _
           "See Data_Checks for anything that needs review." & vbCrLf & _
           "Time taken: " & Format$(Timer - t, "0.0") & " seconds", vbInformation, "Hardware Report"
    Exit Sub
Fail:
    Dim msg As String
    msg = "Stopped during step " & gStep & vbCrLf & vbCrLf & Err.Description
    SpeedUp False
    MsgBox msg, vbCritical, "Hardware Report - error"
End Sub

'=====================================================================
' STEP 1 - CONSOLIDATE WEEKLY SHEETS
'=====================================================================
Private Sub Step1_Consolidate()
    Dim wsOut As Worksheet, ws As Worksheet
    Dim hdr As Variant
    Dim i As Long, n As Long, lastR As Long, outR As Long, nWeeks As Long

    gStep = "1 - Consolidate weekly sheets"
    hdr = Array("Date", "Receipt_No", "Customer", "Category", "Product", "Unit", "Qty", "Unit_Price", "Payment_Method", "Salesperson")

    Set wsOut = NewSheet(SH_RAW)
    wsOut.Range("A1:J1").Value = hdr
    wsOut.Range("K1").Value = "Source_Week"
    outR = 2

    For Each ws In ThisWorkbook.Worksheets
        If UCase$(ws.Name) Like WEEK_PATTERN Then
            For i = 0 To 9
                If LCase$(Trim$(CStr(ws.Cells(1, i + 1).Value))) <> LCase$(CStr(hdr(i))) Then
                    Err.Raise vbObjectError + 520, , "Sheet '" & ws.Name & "' has the wrong header in column " & (i + 1) & _
                        ". Expected '" & hdr(i) & "'."
                End If
            Next i
            lastR = LastRowAny(ws)
            If lastR >= 2 Then
                n = lastR - 1
                ' Copy (not .Value) so text dates/numbers stay exactly as typed
                ws.Range("A2:J" & lastR).Copy wsOut.Range("A" & outR)
                wsOut.Range("K" & outR).Resize(n, 1).Value = ws.Name
                outR = outR + n
            End If
            nWeeks = nWeeks + 1
        End If
    Next ws
    Application.CutCopyMode = False

    If nWeeks = 0 Then Err.Raise vbObjectError + 521, , "No sheets named Week1, Week2, ... were found."
    cRead = outR - 2
    If cRead < 1 Then Err.Raise vbObjectError + 522, , "The weekly sheets contain no data rows."
    StyleHeader wsOut.Range("A1:K1")
    wsOut.Columns("A:K").AutoFit
End Sub

'=====================================================================
' STEP 2 - CLEAN
'=====================================================================
Private Sub Step2_Clean()
    Dim wsRaw As Worksheet, wsC As Worksheet, wsR As Worksheet, wsL As Worksheet
    Dim lo As ListObject
    Dim raw As Variant, outA() As Variant, rejA() As Variant
    Dim nRaw As Long, i As Long, j As Long, k As Long, kr As Long
    Dim reason As String, v As Variant
    Dim dt As Date, okDate As Boolean, okN As Boolean, wasText As Boolean
    Dim qty As Double, price As Double
    Dim receipt As String, prod As String, pk As String, cat As String, rawCat As String
    Dim pay As String, rawPay As String, cust As String, sp As String, unit As String, sKey As String
    Dim seen As Object
    Dim ratio As Double, rq As Double, rp As Double, wt2 As Boolean, fixNote As String

    gStep = "2 - Clean data"
    LoadLookups

    Set wsRaw = ThisWorkbook.Worksheets(SH_RAW)
    nRaw = cRead
    raw = wsRaw.Range("A2:K" & nRaw + 1).Value
    ReDim outA(1 To nRaw, 1 To 11)
    ReDim rejA(1 To nRaw, 1 To 12)
    Set seen = NewDict()

    For i = 1 To nRaw
        reason = ""
        fixNote = ""

        ' ---- 1. completely blank rows are dropped silently (counted)
        If RowBlank(raw, i, 10) Then
            cBlank = cBlank + 1
            GoTo NextRow
        End If

        ' ---- 2. date (real dates kept; text dates parsed as dd/mm/yyyy)
        v = raw(i, 1)
        okDate = False
        If IsBlank(v) Then
            reason = "Missing date": cBadDate = cBadDate + 1
        Else
            If VarType(v) = vbString Then
                okDate = ParseDate(CStr(v), dt)
                If okDate Then cTextDate = cTextDate + 1
            ElseIf VarType(v) = vbDate Then
                dt = v: okDate = True
            ElseIf IsNumeric(v) Then
                If CDbl(v) > 36526 Then dt = CDate(CDbl(v)): okDate = True
            End If
            If Not okDate Then reason = "Unreadable date": cBadDate = cBadDate + 1
        End If

        ' ---- 3. receipt + product
        receipt = UCase$(CleanText(raw(i, 2)))
        prod = CleanText(raw(i, 5))
        pk = KeyOf(prod)
        If reason = "" And receipt = "" Then reason = "Missing receipt number": cNoReceipt = cNoReceipt + 1
        If reason = "" And prod = "" Then reason = "Missing product": cNoProduct = cNoProduct + 1

        ' ---- 4. quantity
        If reason = "" Then
            If Not ToNumber(raw(i, 7), qty, wasText) Then
                reason = "Missing or non-numeric quantity": cBadQty = cBadQty + 1
            ElseIf qty <= 0 Then
                reason = "Zero or negative quantity": cBadQty = cBadQty + 1
            ElseIf wasText Then
                cTextNum = cTextNum + 1
            End If
        End If

        ' ---- 5. price (blank/zero -> standard price from Lookups if product known)
        If reason = "" Then
            okN = ToNumber(raw(i, 8), price, wasText)
            If okN And price > 0 Then
                If wasText Then cTextNum = cTextNum + 1
            ElseIf dProdPrice.Exists(pk) Then
                price = dProdPrice(pk): cPriceFill = cPriceFill + 1
            Else
                reason = "Missing price and product not in Lookups": cBadPrice = cBadPrice + 1
            End If
        End If

        ' ---- 5b. decimal-shift price typos (x10 or /10 of the standard price) are corrected
        If reason = "" Then
            If dProdPrice.Exists(pk) Then
                If dProdPrice(pk) > 0 Then
                    ratio = price / dProdPrice(pk)
                    If ratio >= SHIFT_LOW And ratio <= SHIFT_HIGH Then
                        fixNote = "Charged " & Format$(price, "#,##0.00") & _
                               ", corrected to " & Format$(Round(price / 10, 2), "#,##0.00") & " (standard " & Format$(dProdPrice(pk), "#,##0.00") & ")"
                        price = Round(price / 10, 2)
                    ElseIf ratio >= 1 / SHIFT_HIGH And ratio <= 1 / SHIFT_LOW Then
                        fixNote = "Charged " & Format$(price, "#,##0.00") & _
                               ", corrected to " & Format$(Round(price * 10, 2), "#,##0.00") & " (standard " & Format$(dProdPrice(pk), "#,##0.00") & ")"
                        price = Round(price * 10, 2)
                    End If
                End If
            End If
        End If

        ' ---- 6. exact duplicates
        If reason = "" Then
            sKey = CStr(CLng(dt)) & "|" & receipt & "|" & pk & "|" & CStr(qty) & "|" & CStr(price)
            If seen.Exists(sKey) Then
                reason = "Duplicate row": cDup = cDup + 1
            Else
                seen.Add sKey, 1
            End If
        End If

        ' ---- rejected rows are logged, not lost
        If reason <> "" Then
            kr = kr + 1
            For j = 1 To 11
                rejA(kr, j) = raw(i, j)
            Next j
            rejA(kr, 12) = reason
            ' keep text dates as text so Excel does not re-interpret them month-first
            If VarType(raw(i, 1)) = vbString Then rejA(kr, 1) = "'" & CStr(raw(i, 1))
            ' estimate the revenue that is NOT in the report (duplicates are not lost sales)
            If reason <> "Duplicate row" Then
                If ToNumber(raw(i, 7), rq, wt2) Then
                    If Not ToNumber(raw(i, 8), rp, wt2) Then rp = 0
                    If rp <= 0 Then
                        If dProdPrice.Exists(pk) Then rp = dProdPrice(pk) Else rp = 0
                    End If
                    If rq > 0 Then
                        cRejValue = cRejValue + rq * rp
                    ElseIf rq < 0 Then
                        cRetValue = cRetValue + rq * rp
                    End If
                Else
                    cRejNoQty = cRejNoQty + 1
                End If
            End If
            GoTo NextRow
        End If

        ' ---- 7. standardise text fields
        rawCat = CleanText(raw(i, 4))
        If dProdName.Exists(pk) Then
            If dProdName(pk) <> prod Then cProdFix = cProdFix + 1
            prod = dProdName(pk)
            cat = dProdCat(pk)
            unit = dProdUnit(pk)
            If CleanText(raw(i, 6)) = "" Then cUnitFill = cUnitFill + 1
        Else
            cUnmatched = cUnmatched + 1
            If dCat.Exists(KeyOf(rawCat)) Then
                cat = dCat(KeyOf(rawCat))
            Else
                cat = "Uncategorised"
            End If
            unit = CleanText(raw(i, 6))
            If unit = "" Then unit = "unit": cUnitFill = cUnitFill + 1
        End If
        If rawCat <> cat Then cCatFix = cCatFix + 1

        rawPay = CleanText(raw(i, 9))
        If dPay.Exists(KeyOf(rawPay)) Then
            pay = dPay(KeyOf(rawPay))
        Else
            pay = "Other"
        End If
        If rawPay <> pay Then cPayFix = cPayFix + 1

        cust = CleanText(raw(i, 3))
        If cust = "" Then
            cust = "Walk-in": cCustFill = cCustFill + 1
        Else
            cust = StrConv(cust, vbProperCase)
            If KeyOf(cust) = "walkin" Then cust = "Walk-in"
        End If

        sp = CleanText(raw(i, 10))
        If sp = "" Then
            sp = "Unassigned": cSalesFill = cSalesFill + 1
        Else
            sp = StrConv(sp, vbProperCase)
        End If

        ' ---- 8. store clean row
        If fixNote <> "" Then LogFix dt, receipt, prod, fixNote: cPriceFix = cPriceFix + 1
        k = k + 1
        outA(k, 1) = CDbl(Int(dt))
        outA(k, 2) = receipt
        outA(k, 3) = cust
        outA(k, 4) = cat
        outA(k, 5) = prod
        outA(k, 6) = unit
        outA(k, 7) = qty
        outA(k, 8) = price
        outA(k, 9) = pay
        outA(k, 10) = sp
        outA(k, 11) = CStr(raw(i, 11))
NextRow:
    Next i

    cKept = k
    cRej = kr
    If k = 0 Then Err.Raise vbObjectError + 530, , "No valid rows remain after cleaning. Check Rejected_Rows."

    ' ---------------- Monthly_Clean (Excel table) ----------------
    Set wsC = NewSheet(SH_CLEAN)
    wsC.Range("A1:L1").Value = Array("Date", "Receipt_No", "Customer", "Category", "Product", "Unit", "Qty", "Unit_Price", _
                                     "Payment_Method", "Salesperson", "Source_Week", "Revenue")
    wsC.Range("A2").Resize(k, 11).Value = outA
    Set lo = wsC.ListObjects.Add(xlSrcRange, wsC.Range("A1").Resize(k + 1, 12), , xlYes)
    lo.Name = TBL
    lo.TableStyle = "TableStyleMedium2"
    lo.ListColumns("Revenue").DataBodyRange.FormulaR1C1 = "=RC7*RC8"   ' Qty (col G) x Unit_Price (col H)
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns("Date").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    wsC.Columns("A").NumberFormat = "dd-mmm-yyyy"
    wsC.Columns("H").NumberFormat = "#,##0.00"
    wsC.Columns("L").NumberFormat = "#,##0.00"
    wsC.Columns("A:L").AutoFit
    FreezeTop wsC

    ' ---------------- Rejected_Rows ----------------
    Set wsR = NewSheet(SH_REJ)
    wsR.Range("A1:L1").Value = Array("Date", "Receipt_No", "Customer", "Category", "Product", "Unit", "Qty", "Unit_Price", _
                                     "Payment_Method", "Salesperson", "Source_Week", "Reason_Removed")
    StyleHeader wsR.Range("A1:L1")
    If kr > 0 Then
        wsR.Range("A2").Resize(kr, 12).Value = rejA
        wsR.Columns("A").NumberFormat = "dd-mmm-yyyy"
    Else
        wsR.Range("A2").Value = "No rows were rejected."
    End If
    wsR.Columns("A:L").AutoFit
    FreezeTop wsR

    ' ---------------- Cleaning_Log ----------------
    Set wsL = NewSheet(SH_LOG)
    wsL.Range("A1").Value = "CLEANING LOG"
    wsL.Range("A1").Font.Bold = True: wsL.Range("A1").Font.Size = 14
    wsL.Range("A3:B3").Value = Array("Row reconciliation", "Rows")
    StyleHeader wsL.Range("A3:B3")
    LogLine wsL, 4, "Rows read from weekly sheets", cRead
    LogLine wsL, 5, "Blank rows removed", cBlank
    LogLine wsL, 6, "Rows rejected (see Rejected_Rows)", cRej
    LogLine wsL, 7, "Clean rows kept (Monthly_Clean)", cKept
    wsL.Range("A8").Value = "Reconciliation (read - blank - rejected = kept)"
    wsL.Range("B8").Formula = "=IF(B4-B5-B6=B7,""PASS"",""FAIL"")"
    wsL.Range("A8:B8").Font.Bold = True

    wsL.Range("A10:B10").Value = Array("Why rows were rejected", "Rows")
    StyleHeader wsL.Range("A10:B10")
    LogLine wsL, 11, "Missing / unreadable date", cBadDate
    LogLine wsL, 12, "Missing receipt number", cNoReceipt
    LogLine wsL, 13, "Missing product", cNoProduct
    LogLine wsL, 14, "Missing / zero / negative quantity", cBadQty
    LogLine wsL, 15, "Missing price and product not in Lookups", cBadPrice
    LogLine wsL, 16, "Duplicate rows", cDup

    wsL.Range("A18:B18").Value = Array("Fixes applied to kept rows", "Cells / rows")
    StyleHeader wsL.Range("A18:B18")
    LogLine wsL, 19, "Text dates converted to real dates", cTextDate
    LogLine wsL, 20, "Numbers stored as text converted", cTextNum
    LogLine wsL, 21, "Product names standardised (case / spacing)", cProdFix
    LogLine wsL, 22, "Categories standardised or corrected from product master", cCatFix
    LogLine wsL, 23, "Payment methods standardised", cPayFix
    LogLine wsL, 24, "Missing prices filled from standard price", cPriceFill
    LogLine wsL, 25, "Missing units filled", cUnitFill
    LogLine wsL, 26, "Blank customers set to Walk-in", cCustFill
    LogLine wsL, 27, "Blank salespeople set to Unassigned", cSalesFill
    LogLine wsL, 28, "Rows with a product not in Lookups", cUnmatched
    LogLine wsL, 29, "Prices corrected for x10 or /10 typos (see Data_Checks)", cPriceFix

    wsL.Range("A31:B31").Value = Array("Revenue NOT included in the report (estimate)", "KES / rows")
    StyleHeader wsL.Range("A31:B31")
    LogLine wsL, 32, "Rejected sales with readable quantity and price (duplicates excluded)", cRejValue
    LogLine wsL, 33, "Negative quantities - possible returns (value, not netted off revenue)", cRetValue
    LogLine wsL, 34, "Rejected rows with unreadable quantity (value unknown, rows)", cRejNoQty
    wsL.Range("B32:B33").NumberFormat = "#,##0"
    wsL.Columns("A").ColumnWidth = 72
    wsL.Columns("B").ColumnWidth = 14
End Sub

'=====================================================================
' STEP 3 - CONSISTENCY CHECKS
'=====================================================================
Private Sub Step3_Validate()
    Dim ws As Worksheet, lo As ListObject
    Dim d As Variant, n As Long, i As Long
    Dim dMonth As Object, dRec As Object, dRecCust As Object, dFlag As Object, dUnm As Object, dWmin As Object, dWmax As Object
    Dim conflict As Boolean
    Dim key As Variant, mk As String, bestKey As String, best As Long
    Dim sig As String, pk As String, wk As String, rec As String, prodTxt As String
    Dim dtv As Double, ratio As Double
    Dim nOut As Long, nRec As Long, nPrice As Long, nUnm As Long, nWeek As Long, nBad As Long
    Dim nFix As Long, itm As Variant

    gStep = "3 - Consistency checks"
    Set lo = ThisWorkbook.Worksheets(SH_CLEAN).ListObjects(TBL)
    d = lo.DataBodyRange.Value
    n = UBound(d, 1)

    Set ws = NewSheet(SH_CHECK)
    ws.Range("A1").Value = "DATA CONSISTENCY CHECKS"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Reporting month"
    ws.Range("A3:D3").Value = Array("Check", "Result", "Count", "What it means")
    StyleHeader ws.Range("A3:D3")
    ws.Range("A17").Value = "FLAGGED ITEMS (first " & MAX_ISSUES & ")"
    ws.Range("A17").Font.Bold = True
    ws.Range("A18:E18").Value = Array("Check", "Date", "Receipt", "Product", "Detail")
    StyleHeader ws.Range("A18:E18")
    gCheckRow = 4: gIssueRow = 19: gIssueCount = 0

    Set dMonth = NewDict(): Set dRec = NewDict(): Set dRecCust = NewDict(): Set dFlag = NewDict()
    Set dUnm = NewDict(): Set dWmin = NewDict(): Set dWmax = NewDict()

    ' dominant month = the month holding most rows
    For i = 1 To n
        mk = Format$(CDate(d(i, 1)), "yyyymm")
        dMonth(mk) = dMonth(mk) + 1
    Next i
    For Each key In dMonth.Keys
        If dMonth(key) > best Then best = dMonth(key): bestKey = CStr(key)
    Next key
    gMonthStart = DateSerial(CLng(Left$(bestKey, 4)), CLng(Mid$(bestKey, 5, 2)), 1)
    ws.Range("B2").Value = CDbl(gMonthStart)
    ws.Range("B2").NumberFormat = "mmmm yyyy"
    ws.Range("B2").HorizontalAlignment = xlLeft

    For i = 1 To n
        dtv = CDbl(d(i, 1))
        rec = CStr(d(i, 2))
        prodTxt = CStr(d(i, 5))
        pk = KeyOf(prodTxt)
        wk = CStr(d(i, 11))

        ' a) date outside the reporting month
        If Format$(CDate(dtv), "yyyymm") <> bestKey Then
            nOut = nOut + 1
            AddIssue ws, "Date outside reporting month", dtv, rec, prodTxt, "Found in " & wk & " - check for a typing error in the date"
        End If

        ' b) same receipt number with different date / payment / (named) customer
        conflict = False
        sig = CStr(dtv) & "|" & CStr(d(i, 9))
        If dRec.Exists(rec) Then
            If dRec(rec) <> sig Then conflict = True
        Else
            dRec.Add rec, sig
        End If
        If CStr(d(i, 3)) <> "Walk-in" Then
            If dRecCust.Exists(rec) Then
                If dRecCust(rec) <> CStr(d(i, 3)) Then conflict = True
            Else
                dRecCust.Add rec, CStr(d(i, 3))
            End If
        End If
        If conflict Then
            If Not dFlag.Exists(rec) Then
                dFlag.Add rec, 1
                nRec = nRec + 1
                AddIssue ws, "Receipt with conflicting details", dtv, rec, prodTxt, "Lines of this receipt differ in date, payment method or customer"
            End If
        End If

        ' c) price far from the standard price
        If dProdPrice.Exists(pk) Then
            If dProdPrice(pk) > 0 Then
                ratio = CDbl(d(i, 8)) / dProdPrice(pk)
                If ratio < FLAG_PRICE_LOW Or ratio > FLAG_PRICE_HIGH Then
                    nPrice = nPrice + 1
                    AddIssue ws, "Price far from standard", dtv, rec, prodTxt, _
                        "Charged " & Format$(d(i, 8), "#,##0") & " vs standard " & Format$(dProdPrice(pk), "#,##0")
                End If
            End If
        End If

        ' d) product missing from the Lookups master
        If Not dProdName.Exists(pk) Then
            If Not dUnm.Exists(pk) Then
                dUnm.Add pk, 1
                nUnm = nUnm + 1
                AddIssue ws, "Product not in Lookups", dtv, rec, prodTxt, "Category set to '" & CStr(d(i, 4)) & "'. Add it to Lookups"
            End If
        End If

        ' e) zero / negative values that slipped through
        If CDbl(d(i, 7)) <= 0 Or CDbl(d(i, 8)) <= 0 Then nBad = nBad + 1

        ' f) date span of each weekly sheet
        If dWmin.Exists(wk) Then
            If dtv < dWmin(wk) Then dWmin(wk) = dtv
            If dtv > dWmax(wk) Then dWmax(wk) = dtv
        Else
            dWmin.Add wk, dtv
            dWmax.Add wk, dtv
        End If
    Next i

    For Each key In dWmin.Keys
        If dWmax(key) - dWmin(key) > 6 Then
            nWeek = nWeek + 1
            AddIssue ws, "Week sheet spans more than 7 days", dWmin(key), "", "", CStr(key) & " runs from " & _
                Format$(CDate(dWmin(key)), "dd-mmm") & " to " & Format$(CDate(dWmax(key)), "dd-mmm")
        End If
    Next key

    nFix = 0
    For Each itm In gFix
        nFix = nFix + 1
        AddIssue ws, "Price auto-corrected", CDbl(itm(0)), CStr(itm(1)), CStr(itm(2)), CStr(itm(3))
    Next itm
    gOutOfMonth = nOut

    AddCheck ws, "Row reconciliation (read = blank + rejected + kept)", Abs(cRead - cBlank - cRej - cKept), "Difference between rows read and rows accounted for"
    AddCheck ws, "Dates outside reporting month", nOut, "Probably a typing error in the date"
    AddCheck ws, "Weekly sheets spanning more than 7 days", nWeek, "A week sheet contains dates from another week"
    AddCheck ws, "Receipts with conflicting details", nRec, "Same receipt number but different date, customer or payment"
    AddCheck ws, "Prices far from standard (<50% or >200%)", nPrice, "Kept in the totals - confirm against the receipt"
    AddCheck ws, "Prices auto-corrected (x10 or /10 typos)", nFix, "Corrected to the standard scale - confirm they were typos"
    AddCheck ws, "Products not in Lookups", nUnm, "Add them to the Lookups sheet so they are categorised correctly"
    AddCheck ws, "Zero / negative quantity or price in clean data", nBad, "Should always be 0"
    If gIssueCount > MAX_ISSUES Then ws.Range("A" & gIssueRow).Value = "... " & (gIssueCount - MAX_ISSUES) & " more flagged items not listed."

    ws.Columns("A").ColumnWidth = 52
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 70
End Sub

'=====================================================================
' STEP 4 - PIVOT TABLES (all share ONE pivot cache so slicers can drive them all)
'=====================================================================
Private Sub Step4_Pivots()
    Dim ws As Worksheet, pc As PivotCache, pf As PivotField, lo As ListObject
    Dim pivTotal As Double, dataTotal As Double

    gStep = "4 - Pivot tables"
    Set lo = ThisWorkbook.Worksheets(SH_CLEAN).ListObjects(TBL)
    Set ws = NewSheet(SH_PIV)
    ws.Range("A1").Value = "PIVOT TABLES (source: " & TBL & " on " & SH_CLEAN & ") - they feed the Dashboard charts and cards"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 12

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=TBL)

    Set pvCat = BuildPivot(pc, ws, "A4", "pvtCategory", "Category", "", True, "Revenue by category")
    Set pvProd = BuildPivot(pc, ws, "E4", "pvtProduct", "Product", "", True, "Top 10 products by revenue")
    On Error Resume Next
    pvProd.PivotFields("Product").PivotFilters.Add Type:=xlTopCount, DataField:=pvProd.DataFields(1), Value1:=10
    On Error GoTo 0
    Set pvPay = BuildPivot(pc, ws, "I4", "pvtPayment", "Payment_Method", "", True, "Revenue by payment method")
    Set pvSp = BuildPivot(pc, ws, "M4", "pvtSalesperson", "Salesperson", "", True, "Revenue by salesperson")
    Set pvCust = BuildPivot(pc, ws, "Q4", "pvtCustomer", "Customer", "", True, "Top 10 customers by revenue")
    On Error Resume Next
    pvCust.PivotFields("Customer").PivotFilters.Add Type:=xlTopCount, DataField:=pvCust.DataFields(1), Value1:=10
    On Error GoTo 0
    Set pvWk = BuildPivot(pc, ws, "U4", "pvtWeekCategory", "Source_Week", "Category", False, "Revenue by week and category")
    Set pvDay = BuildPivot(pc, ws, "AL4", "pvtDaily", "Date", "", False, "Daily revenue")
    On Error Resume Next
    pvDay.PivotFields("Date").NumberFormat = "dd-mmm"
    On Error GoTo 0

    ' KPI pivot: totals only (no row field) - the dashboard cards read it with GETPIVOTDATA
    ws.Range("AI2").Value = "KPI totals (feeds dashboard cards)"
    ws.Range("AI2").Font.Bold = True
    Set pvKpi = pc.CreatePivotTable(TableDestination:=ws.Range("AI4"), TableName:="pvtKPI")
    Set pf = pvKpi.AddDataField(pvKpi.PivotFields("Revenue"), "Total Revenue (KES)", xlSum)
    pf.NumberFormat = "#,##0"
    Set pf = pvKpi.AddDataField(pvKpi.PivotFields("Receipt_No"), "Line Items", xlCount)
    pf.NumberFormat = "#,##0"
    On Error Resume Next
    pvKpi.TableStyle2 = "PivotStyleMedium2"
    pvKpi.HasAutoFormat = False
    pvKpi.TableRange2.Columns.AutoFit
    On Error GoTo 0

    ' tie-out: pivot grand total must equal the table total
    On Error Resume Next
    pivTotal = pvCat.GetPivotData("Total Revenue (KES)").Value
    On Error GoTo 0
    dataTotal = Application.WorksheetFunction.Sum(lo.ListColumns("Revenue").DataBodyRange)
    AddCheck ThisWorkbook.Worksheets(SH_CHECK), "Pivot grand total = clean data total", _
        IIf(Abs(pivTotal - dataTotal) < 0.5, 0, 1), "Pivot " & Format$(pivTotal, "#,##0") & " vs data " & Format$(dataTotal, "#,##0")
End Sub

Private Function BuildPivot(pc As PivotCache, ws As Worksheet, ByVal anchor As String, ByVal ptName As String, _
                            ByVal rowFld As String, ByVal colFld As String, ByVal sortDesc As Boolean, _
                            ByVal title As String) As PivotTable
    Dim pt As PivotTable, pf As PivotField

    ws.Range(anchor).Offset(-2, 0).Value = title
    ws.Range(anchor).Offset(-2, 0).Font.Bold = True
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range(anchor), TableName:=ptName)
    pt.PivotFields(rowFld).Orientation = xlRowField
    If colFld <> "" Then pt.PivotFields(colFld).Orientation = xlColumnField

    Set pf = pt.AddDataField(pt.PivotFields("Revenue"), "Total Revenue (KES)", xlSum)
    pf.NumberFormat = "#,##0"
    If sortDesc Then pt.PivotFields(rowFld).AutoSort xlDescending, "Total Revenue (KES)"

    On Error Resume Next
    pt.RowAxisLayout xlTabularRow
    pt.TableStyle2 = "PivotStyleMedium2"
    pt.HasAutoFormat = False
    pt.TableRange2.Columns.AutoFit
    On Error GoTo 0
    Set BuildPivot = pt
End Function

'=====================================================================
' STEP 5 - DASHBOARD (pivot charts + KPI cards + slicers)
'=====================================================================
Private Sub Step5_Dashboard()
    Dim wsD As Worksheet, monthStart As Date, co As ChartObject
    Dim slicerErr As String

    gStep = "5 - Dashboard"
    monthStart = ThisWorkbook.Worksheets(SH_CHECK).Range("B2").Value

    Set wsD = NewSheet(SH_DASH)
    wsD.Activate
    ActiveWindow.DisplayGridlines = False
    wsD.Columns("A").ColumnWidth = 2
    wsD.Columns("B:M").ColumnWidth = 14
    wsD.Columns("N").ColumnWidth = 2
    wsD.Columns("O:Q").ColumnWidth = 14
    wsD.Range("A1:R2").Interior.Color = RGB(31, 56, 100)
    wsD.Range("B1").Value = "HARDWARE STORE - MONTHLY SALES DASHBOARD"
    wsD.Range("B1").Font.Size = 18: wsD.Range("B1").Font.Bold = True: wsD.Range("B1").Font.Color = vbWhite
    wsD.Range("B2").Value = Format$(monthStart, "mmmm yyyy") & "   |   generated " & Format$(Now, "dd-mmm-yyyy hh:nn") & _
                            "   |   click the slicers on the right to filter every chart and card"
    wsD.Range("B2").Font.Color = RGB(200, 210, 230)
    wsD.Rows(1).RowHeight = 30

    KpiCard wsD, "B4:C4", "B5:C6", "TOTAL REVENUE (KES)", "=IFERROR(GETPIVOTDATA(""Total Revenue (KES)"",Pivots!$AI$4),0)", "#,##0"
    KpiCard wsD, "D4:E4", "D5:E6", "LINE ITEMS SOLD", "=IFERROR(GETPIVOTDATA(""Line Items"",Pivots!$AI$4),0)", "#,##0"
    KpiCard wsD, "F4:G4", "F5:G6", "AVG LINE VALUE (KES)", "=IFERROR(B5/D5,0)", "#,##0"
    KpiCard wsD, "H4:I4", "H5:I6", "TOP CATEGORY", "=IFERROR(Pivots!$A$5,""-"")", "General"
    KpiCard wsD, "J4:K4", "J5:K6", "TOP PRODUCT", "=IFERROR(Pivots!$E$5,""-"")", "General"
    KpiCard wsD, "L4:M4", "L5:M6", "TOP SALESPERSON", "=IFERROR(Pivots!$M$5,""-"")", "General"
    wsD.Range("H5,J5,L5").Font.Size = 13
    wsD.Range("H5,J5,L5").WrapText = True

    PivotChart wsD, pvCat, xlColumnClustered, "Revenue by category (KES)", "B8:G22", False
    PivotChart wsD, pvDay, xlLine, "Daily revenue (KES)", "H8:M22", False
    Set co = PivotChart(wsD, pvProd, xlBarClustered, "Top 10 products (KES)", "B24:G38", False)
    On Error Resume Next
    co.Chart.Axes(xlCategory).ReversePlotOrder = True
    On Error GoTo 0
    PivotChart wsD, pvPay, xlDoughnut, "Payment methods", "H24:M38", True
    PivotChart wsD, pvWk, xlColumnStacked, "Weekly revenue by category (KES)", "B40:G54", True
    PivotChart wsD, pvSp, xlColumnClustered, "Revenue by salesperson (KES)", "H40:M54", False

    ' ---- slicers: one cache per field, connected to every pivot
    gSlicerStage = "": gSlicerNote = ""
    On Error Resume Next
    AddSlicers wsD, Array(pvCat, pvProd, pvPay, pvSp, pvCust, pvWk, pvDay, pvKpi)
    If Err.Number <> 0 Then slicerErr = "Failed while " & gSlicerStage & ": " & Err.Description
    On Error GoTo 0
    If slicerErr <> "" Then
        wsD.Range("O4").Value = "Slicers could not be completed."
        wsD.Range("O5").Value = slicerErr
    ElseIf gSlicerNote <> "" Then
        wsD.Range("O4").Value = "Some pivots were not linked to a slicer:"
        wsD.Range("O5").Value = gSlicerNote
    End If

    wsD.Range("B56").Value = "Source: Monthly_Clean (tblSales) via the Pivots sheet. Review Data_Checks for flagged items before sharing this report."
    wsD.Range("B56").Font.Italic = True: wsD.Range("B56").Font.Size = 9
    wsD.Range("B57").Value = "Not in the totals (whole month, unaffected by slicers): rejected sales worth about KES " & _
                             Format$(cRejValue, "#,##0") & " and possible returns of KES " & Format$(-cRetValue, "#,##0") & " - see Cleaning_Log."
    wsD.Range("B58").Value = "Rows dated outside the month and kept: " & gOutOfMonth & "   |   price typos auto-corrected: " & cPriceFix & _
                             "   |   details in Data_Checks."
    wsD.Range("B57:B58").Font.Italic = True: wsD.Range("B57:B58").Font.Size = 9
    wsD.Move Before:=ThisWorkbook.Worksheets(1)
    wsD.Range("A1").Select
End Sub

Private Sub AddSlicers(wsD As Worksheet, pts As Variant)
    Dim flds As Variant, caps As Variant, hts As Variant
    Dim i As Long, j As Long, t As Double, l As Double
    Dim sc As SlicerCache, sl As Slicer
    Dim ptMain As PivotTable, ptOther As PivotTable

    flds = Array("Source_Week", "Payment_Method", "Salesperson", "Category")
    caps = Array("Week", "Payment method", "Salesperson", "Category")
    hts = Array(105, 130, 135, 240)
    l = wsD.Range("O4").Left
    t = wsD.Range("O4").Top
    Set ptMain = pts(0)
    gSlicerNote = ""

    For i = 0 To 3
        gSlicerStage = "creating the slicer cache for " & CStr(flds(i))
        Set sc = ThisWorkbook.SlicerCaches.Add2(ptMain, CStr(flds(i)), "slc_" & CStr(flds(i)))

        For j = 1 To UBound(pts)
            Set ptOther = pts(j)
            gSlicerStage = "connecting the " & CStr(flds(i)) & " slicer to pivot " & ptOther.Name
            On Error Resume Next
            sc.PivotTables.AddPivotTable ptOther
            If Err.Number <> 0 Then
                gSlicerNote = gSlicerNote & CStr(flds(i)) & "->" & ptOther.Name & " (" & Err.Description & "); "
                Err.Clear
            End If
            On Error GoTo 0
        Next j

        gSlicerStage = "placing the " & CStr(flds(i)) & " slicer on the dashboard"
        Set sl = sc.Slicers.Add(wsD, , "sl_" & CStr(flds(i)), CStr(caps(i)), t, l, 185, CDbl(hts(i)))
        t = t + hts(i) + 8
    Next i
End Sub

'=====================================================================
' OTHER MACROS
'=====================================================================
Public Sub SaveMonthlyArchive()
    Dim p As String, fn As String, m As Date
    If Not SheetExists(SH_CHECK) Then
        MsgBox "Run RunMonthlyReport first.", vbExclamation: Exit Sub
    End If
    If ThisWorkbook.Path = "" Then
        MsgBox "Save this workbook first.", vbExclamation: Exit Sub
    End If
    ResetSlicers
    m = ThisWorkbook.Worksheets(SH_CHECK).Range("B2").Value
    p = ThisWorkbook.Path & Application.PathSeparator & "Archive"
    If Dir(p, vbDirectory) = "" Then MkDir p
    fn = p & Application.PathSeparator & "Hardware_Report_" & Format$(m, "yyyy-mm") & ".xlsm"
    ThisWorkbook.SaveCopyAs fn
    MsgBox "Archive saved:" & vbCrLf & fn, vbInformation
End Sub

Public Sub StartNewMonth()
    Dim ws As Worksheet, lastR As Long
    If MsgBox("This clears the data rows on every WeekN sheet (headers stay)." & vbCrLf & _
              "Have you saved the monthly archive?", vbYesNo + vbExclamation, "Start new month") <> vbYes Then Exit Sub
    For Each ws In ThisWorkbook.Worksheets
        If UCase$(ws.Name) Like WEEK_PATTERN Then
            lastR = LastRowAny(ws)
            If lastR >= 2 Then ws.Range("A2:J" & lastR).ClearContents
        End If
    Next ws
    MsgBox "Week sheets cleared. Paste the new month's weekly data, then run RunMonthlyReport.", vbInformation
End Sub

Public Sub SetupControlPanel()
    Dim ws As Worksheet, lines As Variant, i As Long

    Application.DisplayAlerts = False
    If SheetExists(SH_CTRL) Then ThisWorkbook.Worksheets(SH_CTRL).Delete
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = SH_CTRL
    ws.Cells.Font.Name = "Arial"
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 125

    lines = Array("HARDWARE STORE - MONTHLY SALES AUTOMATION", "", _
        "HOW TO USE", _
        "1. Paste each week's sales into Week1, Week2, Week3, Week4 (headers in row 1, data from row 2, same 10 columns).", _
        "2. Update the Lookups sheet only when you add a product, a payment alias or a category alias.", _
        "3. Click 'Run Monthly Report'. It consolidates, cleans, checks, builds the pivots and the dashboard.", _
        "4. Read Data_Checks and Cleaning_Log before sharing the dashboard.", _
        "5. Click 'Save Month Archive', then 'Start New Month' to clear the week sheets.", "", _
        "KEY RULES (full list at the top of the VBA module)", _
        "Duplicates removed | invalid rows moved to Rejected_Rows | x10 price typos corrected | returns not netted off | out-of-month dates flagged.", _
        "Sample data in Week1-Week4 is synthetic.")
    For i = 0 To UBound(lines)
        ws.Cells(i + 1, 2).Value = lines(i)
    Next i
    ws.Range("B1").Font.Bold = True: ws.Range("B1").Font.Size = 14
    ws.Range("B3").Font.Bold = True: ws.Range("B10").Font.Bold = True

    AddButton ws, "B14", "Run Monthly Report", "RunMonthlyReport"
    AddButton ws, "B17", "Save Month Archive", "SaveMonthlyArchive"
    AddButton ws, "B20", "Start New Month", "StartNewMonth"
End Sub

' Clears every slicer filter so the saved file always opens showing the full month
Public Sub ResetSlicers()
    Dim sc As SlicerCache
    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    On Error GoTo 0
End Sub

Private Sub SetDocProperties()
    On Error Resume Next
    With ThisWorkbook
        .BuiltinDocumentProperties("Title") = "Hardware Store - Weekly to Monthly Sales Automation"
        .BuiltinDocumentProperties("Subject") = "Excel VBA: data cleaning, consistency checks, pivots, slicers and dashboard"
        If Len(REPORT_AUTHOR) > 0 Then
            .BuiltinDocumentProperties("Author") = REPORT_AUTHOR
        ElseIf Len(Application.UserName) > 0 Then
            .BuiltinDocumentProperties("Author") = Application.UserName
        End If
    End With
    On Error GoTo 0
End Sub

Private Sub LogFix(ByVal dt As Date, ByVal receipt As String, ByVal prodName As String, ByVal detail As String)
    gFix.Add Array(CDbl(Int(dt)), receipt, prodName, detail)
End Sub

Private Sub AddButton(ws As Worksheet, ByVal addr As String, ByVal cap As String, ByVal macroName As String)
    Dim r As Range
    Set r = ws.Range(addr)
    With ws.Buttons.Add(r.Left, r.Top, 200, 34)
        .OnAction = macroName
        .Caption = cap
        .Font.Bold = True
        .Font.Size = 11
    End With
End Sub

'=====================================================================
' HELPERS
'=====================================================================
Private Sub SpeedUp(ByVal turnOn As Boolean)
    Application.ScreenUpdating = Not turnOn
    Application.DisplayAlerts = Not turnOn
    Application.EnableEvents = Not turnOn
    If Not turnOn Then Application.CutCopyMode = False
End Sub

Private Sub ResetCounters()
    cRead = 0: cBlank = 0: cRej = 0: cKept = 0
    cBadDate = 0: cNoReceipt = 0: cNoProduct = 0: cBadQty = 0: cBadPrice = 0: cDup = 0
    cTextDate = 0: cTextNum = 0: cProdFix = 0: cCatFix = 0: cPayFix = 0
    cPriceFill = 0: cUnitFill = 0: cCustFill = 0: cSalesFill = 0: cUnmatched = 0
    cPriceFix = 0: cRejNoQty = 0: gOutOfMonth = 0: cRejValue = 0: cRetValue = 0
    Set gFix = New Collection
End Sub

Private Sub ClearOutputs()
    Dim nm As Variant, q As Long
    For q = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        ThisWorkbook.SlicerCaches(q).Delete
    Next q
    For Each nm In Array(SH_DASH, SH_CALC, SH_PIV, SH_CHECK, SH_LOG, SH_REJ, SH_CLEAN, SH_RAW)
        If SheetExists(CStr(nm)) Then ThisWorkbook.Worksheets(CStr(nm)).Delete
    Next nm
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next ws
End Function

Private Function NewSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    ws.Cells.Font.Name = "Arial"
    Set NewSheet = ws
End Function

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Sub LoadLookups()
    Dim ws As Worksheet, r As Long, k As String
    If Not SheetExists(SH_LOOK) Then Err.Raise vbObjectError + 513, , "The '" & SH_LOOK & "' sheet is missing."
    Set ws = ThisWorkbook.Worksheets(SH_LOOK)
    Set dProdName = NewDict(): Set dProdCat = NewDict(): Set dProdUnit = NewDict(): Set dProdPrice = NewDict()
    Set dPay = NewDict(): Set dCat = NewDict()

    r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        k = KeyOf(CStr(ws.Cells(r, 1).Value))
        dProdName(k) = Trim$(CStr(ws.Cells(r, 1).Value))
        dProdCat(k) = Trim$(CStr(ws.Cells(r, 2).Value))
        dProdUnit(k) = Trim$(CStr(ws.Cells(r, 3).Value))
        If IsNumeric(ws.Cells(r, 4).Value) Then dProdPrice(k) = CDbl(ws.Cells(r, 4).Value) Else dProdPrice(k) = 0
        dCat(KeyOf(CStr(ws.Cells(r, 2).Value))) = Trim$(CStr(ws.Cells(r, 2).Value))
        r = r + 1
    Loop
    r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 6).Value))) > 0
        dPay(KeyOf(CStr(ws.Cells(r, 6).Value))) = Trim$(CStr(ws.Cells(r, 7).Value))
        dPay(KeyOf(CStr(ws.Cells(r, 7).Value))) = Trim$(CStr(ws.Cells(r, 7).Value))
        r = r + 1
    Loop
    r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 9).Value))) > 0
        dCat(KeyOf(CStr(ws.Cells(r, 9).Value))) = Trim$(CStr(ws.Cells(r, 10).Value))
        dCat(KeyOf(CStr(ws.Cells(r, 10).Value))) = Trim$(CStr(ws.Cells(r, 10).Value))
        r = r + 1
    Loop
End Sub

' lower-case text with everything except letters and digits removed
Private Function KeyOf(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    KeyOf = out
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then IsBlank = True: Exit Function
    If IsEmpty(v) Or IsNull(v) Then IsBlank = True: Exit Function
    If VarType(v) = vbString Then IsBlank = (Len(Trim$(Replace(CStr(v), Chr$(160), " "))) = 0)
End Function

Private Function CleanText(ByVal v As Variant) As String
    Dim s As String
    If IsBlank(v) Then Exit Function
    s = CStr(v)
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = s
End Function

Private Function RowBlank(arr As Variant, ByVal r As Long, ByVal nCols As Long) As Boolean
    Dim j As Long
    For j = 1 To nCols
        If Not IsBlank(arr(r, j)) Then RowBlank = False: Exit Function
    Next j
    RowBlank = True
End Function

Private Function ToNumber(ByVal v As Variant, ByRef result As Double, ByRef wasText As Boolean) As Boolean
    Dim s As String
    wasText = False
    If IsBlank(v) Then Exit Function
    If VarType(v) = vbString Then
        s = UCase$(Trim$(CStr(v)))
        s = Replace(s, Chr$(160), "")
        s = Replace(s, "KES", ""): s = Replace(s, "KSH", ""): s = Replace(s, "KSHS", "")
        s = Replace(s, ",", ""): s = Replace(s, " ", "")
        If Len(s) = 0 Then Exit Function
        If Not IsNumeric(s) Then Exit Function
        result = Val(s)
        wasText = True
        ToNumber = True
    ElseIf IsNumeric(v) Then
        result = CDbl(v)
        ToNumber = True
    End If
End Function

' Parses dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy, yyyy-mm-dd and "05 Oct 2026"
Private Function ParseDate(ByVal s As String, ByRef result As Date) As Boolean
    Dim p() As String, dd As Long, mm As Long, yy As Long, pos As Long
    s = Trim$(s)
    s = Replace(Replace(Replace(s, "-", "/"), ".", "/"), " ", "/")
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(2))) Then Exit Function
    If IsNumeric(p(1)) Then
        mm = CLng(p(1))
    Else
        pos = InStr("janfebmaraprmayjunjulaugsepoctnovdec", LCase$(Left$(p(1), 3)))
        If pos = 0 Or Len(p(1)) < 3 Then Exit Function
        mm = (pos + 2) \ 3
    End If
    If Len(p(0)) = 4 Then
        yy = CLng(p(0)): dd = CLng(p(2))
    Else
        dd = CLng(p(0)): yy = CLng(p(2))
        If yy < 100 Then yy = yy + 2000
    End If
    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Or yy < 2000 Or yy > 2100 Then Exit Function
    result = DateSerial(yy, mm, dd)
    If Day(result) <> dd Then Exit Function
    ParseDate = True
End Function

' Sum valCol by keyCol, return keys/vals sorted high -> low
Private Sub Aggregate(data As Variant, ByVal keyCol As Long, ByVal valCol As Long, keys() As String, vals() As Double, ByRef n As Long)
    Dim dict As Object, i As Long, j As Long, k As String, kv As Variant, iv As Variant
    Dim tk As String, tv As Double
    Set dict = NewDict()
    For i = 1 To UBound(data, 1)
        k = CStr(data(i, keyCol))
        dict(k) = dict(k) + CDbl(data(i, valCol))
    Next i
    n = dict.Count
    ReDim keys(1 To n)
    ReDim vals(1 To n)
    kv = dict.Keys
    iv = dict.Items
    For i = 1 To n
        keys(i) = CStr(kv(i - 1))
        vals(i) = CDbl(iv(i - 1))
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If vals(j) > vals(i) Then
                tv = vals(i): vals(i) = vals(j): vals(j) = tv
                tk = keys(i): keys(i) = keys(j): keys(j) = tk
            End If
        Next j
    Next i
End Sub

Private Function LastRowAny(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastRowAny = 0 Else LastRowAny = f.Row
End Function

Private Sub StyleHeader(r As Range)
    With r
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 56, 100)
    End With
End Sub

Private Sub FreezeTop(ws As Worksheet)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

Private Sub LogLine(ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal value As Double)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = value
End Sub

Private Sub AddCheck(ws As Worksheet, ByVal nm As String, ByVal bad As Long, ByVal note As String)
    ws.Cells(gCheckRow, 1).Value = nm
    ws.Cells(gCheckRow, 2).Value = IIf(bad = 0, "PASS", "REVIEW")
    ws.Cells(gCheckRow, 3).Value = bad
    ws.Cells(gCheckRow, 4).Value = note
    ws.Cells(gCheckRow, 2).Interior.Color = IIf(bad = 0, RGB(198, 239, 206), RGB(255, 235, 156))
    gCheckRow = gCheckRow + 1
End Sub

Private Sub AddIssue(ws As Worksheet, ByVal chk As String, ByVal dtv As Double, ByVal rec As String, _
                     ByVal prod As String, ByVal detail As String)
    gIssueCount = gIssueCount + 1
    If gIssueCount > MAX_ISSUES Then Exit Sub
    ws.Cells(gIssueRow, 1).Value = chk
    ws.Cells(gIssueRow, 2).Value = dtv
    ws.Cells(gIssueRow, 2).NumberFormat = "dd-mmm-yyyy"
    ws.Cells(gIssueRow, 3).Value = rec
    ws.Cells(gIssueRow, 4).Value = prod
    ws.Cells(gIssueRow, 5).Value = detail
    gIssueRow = gIssueRow + 1
End Sub

Private Sub KpiCard(ws As Worksheet, ByVal lblAddr As String, ByVal valAddr As String, ByVal lbl As String, _
                    ByVal f As String, ByVal fmt As String)
    With ws.Range(lblAddr)
        .Merge
        .Value = lbl
        .Font.Size = 9: .Font.Bold = True: .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(221, 235, 247)
    End With
    With ws.Range(valAddr)
        .Merge
        .Formula = f
        .NumberFormat = fmt
        .Font.Size = 18: .Font.Bold = True: .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(242, 247, 252)
    End With
    ws.Range(lblAddr).BorderAround xlContinuous, xlThin, , RGB(166, 190, 220)
    ws.Range(valAddr).BorderAround xlContinuous, xlThin, , RGB(166, 190, 220)
End Sub

Private Function PivotChart(ws As Worksheet, pt As PivotTable, ByVal ctype As Long, ByVal ttl As String, _
                            ByVal placeAddr As String, ByVal showLegend As Boolean) As ChartObject
    Dim co As ChartObject, pr As Range
    Set pr = ws.Range(placeAddr)
    Set co = ws.ChartObjects.Add(pr.Left + 2, pr.Top + 2, pr.Width - 4, pr.Height - 4)
    With co.Chart
        .SetSourceData pt.TableRange1
        .ChartType = ctype
        .HasTitle = True
        .ChartTitle.Text = ttl
        .ChartTitle.Font.Size = 12
        .HasLegend = showLegend
        If showLegend Then .Legend.Position = xlLegendPositionBottom
        On Error Resume Next
        .ShowAllFieldButtons = False
        If ctype = xlDoughnut Then
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.ShowPercentage = True
            .SeriesCollection(1).DataLabels.ShowValue = False
        Else
            .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        End If
        If ctype = xlLine Then .Axes(xlCategory).TickLabels.NumberFormat = "dd-mmm"
        On Error GoTo 0
    End With
    Set PivotChart = co
End Function
