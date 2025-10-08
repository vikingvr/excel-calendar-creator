Option Explicit

'==============================
' FLEXIBLE CALENDAR GENERATOR
' - CY (Calendar Year Jan–Dec) or FY (Fiscal Oct–Sep, label=end year)
' - Equal-height months (6 weeks), 2-row spacer between blocks
' - Custom-date fill & font from "Custom Dates"
' - Legend from "Custom Dates"!G (labels), legend labels forced to BLACK font
' - Version stamp at V48/V49: "YYYYMMDD-HHMM"
' - 3 spacer columns between month blocks (same width as day cells)
'==============================

' ---- Layout constants (shared) ----
Private Const BASE_ROW As Long = 4
Private Const BASE_COL As Long = 2
Private Const DAY_COLS As Long = 7
Private Const GAP_COLS As Long = 3
Private Const ACROSS As Long = 3

' title(1) + headers(1) + weeks(6) + notes(1) + spacer(2) = 11 rows
Private Const WEEK_ROWS As Long = 6
Private Const BLOCK_H As Long = 11

' Column widths
Private Const DAY_COL_WIDTH As Double = 2.2
Private Const GAP_COL_WIDTH As Double = 2.2

' Horizontal block width (7 day cols + 3 spacer cols)
Private Const BLOCK_W As Long = DAY_COLS + GAP_COLS

' Legend placement start row
Private Const LEGEND_START_ROW As Long = 47

Public Sub MakeCalendar()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Build a Fiscal calendar (Oct–Sep)?" & vbCrLf & _
                  "Yes = Fiscal  |  No = Annual  |  Cancel = Abort", _
                  vbYesNoCancel + vbQuestion, "Calendar Type")
    If resp = vbCancel Then Exit Sub

    Dim IsFiscal As Boolean: IsFiscal = (resp = vbYes)

    Dim yearLabel As Long, s As String
    If IsFiscal Then
        s = InputBox("Enter the FISCAL YEAR (END year). Example: 2026" & vbCrLf & _
                     "FY2026 = Oct 2025 – Sep 2026", "Fiscal Year", Year(Date) + 1)
    Else
        s = InputBox("Enter the ANNUAL YEAR (Jan–Dec). Example: 2026", "Annual Year", Year(Date))
    End If
    If Not IsNumeric(s) Or Len(s) = 0 Then Exit Sub
    yearLabel = CLng(s)

    Dim wkResp As VbMsgBoxResult
    wkResp = MsgBox("Start weeks on MONDAY?" & vbCrLf & "(No = Sunday)", _
                    vbYesNo + vbQuestion, "Week Start")
    Dim WeekStart As VbDayOfWeek
    WeekStart = IIf(wkResp = vbYes, vbMonday, vbSunday)

    Dim SheetName As String
    If IsFiscal Then
        SheetName = "FY" & yearLabel
        GenerateCalendar yearLabel, 10, WeekStart, True, SheetName
    Else
        SheetName = "CY" & yearLabel
        GenerateCalendar yearLabel, 1, WeekStart, False, SheetName
    End If

    Dim wsCal As Worksheet, wsCD As Worksheet
    Set wsCal = Worksheets(SheetName)
    On Error Resume Next
    Set wsCD = Worksheets("Custom Dates")
    On Error GoTo 0

    If Not wsCD Is Nothing Then
        ApplyCustomDates wsCal, wsCD, yearLabel, IIf(IsFiscal, 10, 1), IsFiscal, WeekStart
        BuildLegend wsCal, wsCD
    Else
        ClearLegendArea wsCal
    End If

    WriteVersionStamp wsCal
    wsCal.Range("A1").Select
End Sub

'==== Core calendar builder ====
Private Sub GenerateCalendar(ByVal LabelYear As Long, _
                             ByVal StartMonth As Long, _
                             ByVal WeekStart As VbDayOfWeek, _
                             ByVal IsFiscal As Boolean, _
                             ByVal SheetName As String)

    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = Worksheets.Add(After:=Sheets(Sheets.Count))
    ws.name = SheetName

    ws.Cells.Font.name = "Calibri"
    ws.Cells.Font.Size = 10

    Dim title As String
    If IsFiscal Then
        title = "Fiscal Year " & LabelYear & "  (Oct " & (LabelYear - 1) & " – Sep " & LabelYear & ")"
    Else
        title = "Calendar Year " & LabelYear
    End If
    With ws.Range("A1")
        .Value = title
        .Font.Bold = True
        .Font.Size = 20
    End With

    Dim i As Long
    For i = 0 To 11
        Dim m As Long, y As Long
        MapMonthYear LabelYear, StartMonth, i, IsFiscal, m, y

        Dim r As Long, c As Long, row0 As Long, col0 As Long
        r = i \ ACROSS
        c = i Mod ACROSS
        row0 = BASE_ROW + r * BLOCK_H
        col0 = BASE_COL + c * BLOCK_W

        Dim j As Long
        For j = 0 To DAY_COLS - 1
            ws.Columns(col0 + j).ColumnWidth = DAY_COL_WIDTH
        Next j
        For j = 0 To GAP_COLS - 1
            ws.Columns(col0 + DAY_COLS + j).ColumnWidth = GAP_COL_WIDTH
        Next j

        DrawMonth ws, row0, col0, m, y, MonthName(m), WeekStart
    Next i
End Sub

' Map i-th month to actual calendar month/year
Private Sub MapMonthYear(ByVal LabelYear As Long, _
                         ByVal StartMonth As Long, _
                         ByVal offset As Long, _
                         ByVal IsFiscal As Boolean, _
                         ByRef outMonth As Long, _
                         ByRef outYear As Long)

    Dim m As Long
    m = ((StartMonth - 1 + offset) Mod 12) + 1

    Dim y As Long
    If Not IsFiscal And StartMonth = 1 Then
        y = LabelYear
    Else
        If m >= StartMonth Then
            y = LabelYear - 1
        Else
            y = LabelYear
        End If
    End If

    outMonth = m
    outYear = y
End Sub

'==== Draw a single month (equal height) ====
Private Sub DrawMonth(ws As Worksheet, topRow As Long, leftCol As Long, _
                      ByVal monthNum As Long, ByVal yearNum As Long, _
                      ByVal titleText As String, ByVal WeekStart As VbDayOfWeek)

    With ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow, leftCol + (DAY_COLS - 1)))
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
    End With

    Dim headers As Variant
    If WeekStart = vbMonday Then
        headers = Array("M", "T", "W", "T", "F", "S", "S")
    Else
        headers = Array("S", "M", "T", "W", "T", "F", "S")
    End If

    Dim i As Long
    For i = 0 To DAY_COLS - 1
        With ws.Cells(topRow + 1, leftCol + i)
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    Next i

    Dim firstDay As Date: firstDay = DateSerial(yearNum, monthNum, 1)
    Dim daysInMonth As Long: daysInMonth = Day(DateSerial(yearNum, monthNum + 1, 0))
    Dim dow As Long: dow = VBA.Weekday(firstDay, IIf(WeekStart = vbMonday, vbMonday, vbSunday))
    Dim startCol As Long: startCol = dow - 1

    Dim r As Long, c As Long, cell As Range
    For r = 0 To WEEK_ROWS - 1
        For c = 0 To DAY_COLS - 1
            Set cell = ws.Cells(topRow + 2 + r, leftCol + c)
            Dim posIndex As Long: posIndex = r * 7 + c
            Dim dayNum As Long: dayNum = posIndex - startCol + 1

            If dayNum < 1 Or dayNum > daysInMonth Then
                cell.Value = ""
                cell.Borders.Weight = xlHairline
                cell.Interior.ColorIndex = xlNone
                cell.Font.Color = RGB(0, 0, 0)
            Else
                cell.Value = dayNum
                cell.HorizontalAlignment = xlRight

                Dim wknd As Boolean
                If WeekStart = vbMonday Then
                    wknd = (c = 5 Or c = 6)
                Else
                    wknd = (c = 0 Or c = 6)
                End If
                If wknd Then cell.Interior.Color = RGB(230, 230, 230)

                cell.Borders.Weight = xlHairline
            End If
        Next c
    Next r

    Dim lastGridRow As Long: lastGridRow = topRow + 1 + WEEK_ROWS
    With ws.Range(ws.Cells(topRow, leftCol), ws.Cells(lastGridRow, leftCol + (DAY_COLS - 1)))
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    With ws.Range(ws.Cells(lastGridRow + 1, leftCol), ws.Cells(lastGridRow + 1, leftCol + (DAY_COLS - 1)))
        .Merge
        .Value = ""
        .Borders(xlEdgeTop).Weight = xlThin
    End With
End Sub

'===========================
' CUSTOM DATES INTEGRATION
'===========================
Private Sub ApplyCustomDates(wsCalendar As Worksheet, wsCD As Worksheet, _
                             ByVal LabelYear As Long, ByVal StartMonth As Long, _
                             ByVal IsFiscal As Boolean, ByVal WeekStart As VbDayOfWeek)

    Dim lastRow As Long
    lastRow = LastUsedRow(wsCD, "A")
    If lastRow < 1 Then Exit Sub

    Dim mapColors As Object: Set mapColors = CreateObject("Scripting.Dictionary")
    BuildClassificationColorMap wsCD, mapColors

    Dim i As Long
    For i = 1 To lastRow
        Dim dt As Variant, label As String
        dt = wsCD.Cells(i, "A").Value
        label = Trim$(CStr(wsCD.Cells(i, "B").Value))

        If IsDate(dt) And Len(label) > 0 Then
            Dim d As Date: d = CDate(dt)
            If Not DateInThisCalendar(d, LabelYear, StartMonth, IsFiscal) Then GoTo NextRow

            Dim m As Long, y As Long
            m = Month(d): y = Year(d)

            Dim index As Long: index = FiscalOffsetForMonthYear(LabelYear, StartMonth, IsFiscal, m, y)
            If index < 0 Or index > 11 Then GoTo NextRow

            Dim rBlock As Long, cBlock As Long, row0 As Long, col0 As Long
            rBlock = index \ ACROSS
            cBlock = index Mod ACROSS
            row0 = BASE_ROW + rBlock * BLOCK_H
            col0 = BASE_COL + cBlock * BLOCK_W

            Dim firstDay As Date: firstDay = DateSerial(y, m, 1)
            Dim dow As Long: dow = VBA.Weekday(firstDay, IIf(WeekStart = vbMonday, vbMonday, vbSunday))
            Dim startCol As Long: startCol = dow - 1

            Dim dayNum As Long: dayNum = Day(d)
            Dim posIndex As Long: posIndex = startCol + (dayNum - 1)
            Dim rr As Long: rr = posIndex \ 7
            Dim cc As Long: cc = posIndex Mod 7

            Dim target As Range
            Set target = wsCalendar.Cells(row0 + 2 + rr, col0 + cc)

            If mapColors.Exists(LCase$(label)) Then
                Dim sty As Variant
                sty = mapColors(LCase$(label))     ' [0] fill, [1] font
                target.Interior.Color = sty(0)
                target.Font.Color = sty(1)
            Else
                target.Interior.Color = RGB(255, 245, 200)
                target.Font.Color = vbBlack
            End If
        End If
NextRow:
    Next i
End Sub

Private Sub BuildClassificationColorMap(wsCD As Worksheet, dict As Object)
    Dim lastG As Long: lastG = LastUsedRow(wsCD, "G")
    Dim r As Long
    For r = 1 To lastG
        Dim name As String: name = Trim$(CStr(wsCD.Cells(r, "G").Value))
        If Len(name) > 0 Then
            Dim fillClr As Long: fillClr = wsCD.Cells(r, "G").Interior.Color
            Dim fontClr As Long: fontClr = wsCD.Cells(r, "G").Font.Color
            dict(LCase$(name)) = Array(fillClr, fontClr)
        End If
    Next r
End Sub

'===========================
' LEGEND + VERSION STAMP
'===========================
Private Sub BuildLegend(wsCal As Worksheet, wsCD As Worksheet)
    Dim lastG As Long: lastG = LastUsedRow(wsCD, "G")
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim labels As Collection: Set labels = New Collection
    Dim fills As Collection: Set fills = New Collection

    Dim r As Long, nm As String
    For r = 1 To lastG
        nm = Trim$(CStr(wsCD.Cells(r, "G").Value))
        If Len(nm) > 0 Then
            If Not seen.Exists(LCase$(nm)) Then
                seen(LCase$(nm)) = True
                labels.Add nm
                fills.Add wsCD.Cells(r, "G").Interior.Color
                If labels.Count >= 7 Then Exit For
            End If
        End If
    Next r

    ClearLegendArea wsCal

    Dim starts(1 To 3) As Long
    starts(1) = 2    ' B
    starts(2) = 12   ' L
    starts(3) = 22   ' V

    Dim i As Long
    For i = 1 To labels.Count
        Dim groupIdx As Long: groupIdx = ((i - 1) \ 3) + 1
        Dim rowOff As Long: rowOff = (i - 1) Mod 3
        If groupIdx > 3 Then Exit For

        Dim baseCol As Long: baseCol = starts(groupIdx)
        Dim rr As Long: rr = LEGEND_START_ROW + rowOff

        With wsCal.Cells(rr, baseCol)
            .Value = ""
            .Interior.Color = fills(i)
            .Borders.LineStyle = xlContinuous
        End With
        With wsCal.Cells(rr, baseCol + 1)
            .Value = "="
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With wsCal.Cells(rr, baseCol + 2)
            .Value = labels(i)
            .Font.Color = vbBlack      ' <-- force legend text to black
        End With
    Next i
End Sub

Private Sub ClearLegendArea(wsCal As Worksheet)
    wsCal.Range("B47:D49").Clear
    wsCal.Range("L47:N49").Clear
    wsCal.Range("V47:X47").Clear
End Sub

Private Sub WriteVersionStamp(wsCal As Worksheet)
    wsCal.Range("V48").Value = "Version CAO:"
    wsCal.Range("V49").Value = Format(Now, "yyyymmdd-hhmm")
End Sub

'===========================
' UTILITIES
'===========================
Private Function LastUsedRow(ws As Worksheet, colLetter As String) As Long
    On Error Resume Next
    LastUsedRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
    On Error GoTo 0
End Function

Private Function DateInThisCalendar(ByVal d As Date, ByVal LabelYear As Long, _
                                    ByVal StartMonth As Long, ByVal IsFiscal As Boolean) As Boolean
    Dim startDate As Date, endDate As Date
    If IsFiscal Then
        startDate = DateSerial(LabelYear - 1, 10, 1)
        endDate = DateSerial(LabelYear, 9, 30)
    Else
        startDate = DateSerial(LabelYear, 1, 1)
        endDate = DateSerial(LabelYear, 12, 31)
    End If
    DateInThisCalendar = (d >= startDate And d <= endDate)
End Function

Private Function FiscalOffsetForMonthYear(ByVal LabelYear As Long, ByVal StartMonth As Long, _
                                          ByVal IsFiscal As Boolean, ByVal m As Long, ByVal y As Long) As Long
    Dim i As Long, mm As Long, yy As Long
    For i = 0 To 11
        MapMonthYear LabelYear, StartMonth, i, IsFiscal, mm, yy
        If mm = m And yy = y Then
            FiscalOffsetForMonthYear = i
            Exit Function
        End If
    Next i
    FiscalOffsetForMonthYear = -1
End Function


