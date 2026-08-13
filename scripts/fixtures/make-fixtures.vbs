' ─────────────────────────────────────────────────────────────────────
' make-fixtures.vbs — regenerate test/fixtures/ with real Microsoft Excel
'
' Issue #464: every binary input under test/ is built byte-by-byte by the
' test that reads it, so a writer bug the reader mirrors is invisible.
' The XLS and XLSB readers are the sharp end of that — they exist only to
' consume other tools' output and had never seen any. These fixtures are
' the other tool.
'
' Run from a Windows path, never from \\wsl.localhost\ (Excel refuses to
' SaveAs there):
'
'   cscript.exe //Nologo C:\hucre-fixtures\make-fixtures.vbs C:\hucre-fixtures
'
' then copy the output into test/fixtures/.
'
' Late-bound on purpose. Early binding (`New-Object -ComObject
' Excel.Application` from PowerShell) fails with TYPE_E_ELEMENTNOTFOUND on
' the interop cast on at least one otherwise-healthy Excel 16.0 install;
' VBScript is IDispatch-only and never performs that cast.
'
' On the author's name: `Application.UserName` does NOT keep it out of the
' file — Excel 16 stamps docProps/core.xml and SummaryInformation from the
' signed-in Office identity, not from that property. Verified: setting
' UserName still produced <dc:creator>a real name</dc:creator>.
' `Workbook.RemovePersonalInformation = True` before SaveAs is what
' actually blanks creator/lastModifiedBy in xlsx, xlsb *and* the BIFF
' SummaryInformation stream. It is set on every workbook below, and
' test/fixtures/PROVENANCE.md records the check.
'
' SaveAs format codes: 51 = .xlsx, 56 = .xls (BIFF8), 50 = .xlsb.
' ─────────────────────────────────────────────────────────────────────
Option Explicit

Const xlOpenXMLWorkbook = 51
Const xlExcel8 = 56
Const xlExcel12 = 50

Dim outDir, onlyOne
If WScript.Arguments.Count < 1 Then
  WScript.Echo "usage: cscript //Nologo make-fixtures.vbs <output-directory> [one-file-name]"
  WScript.Quit 2
End If
outDir = WScript.Arguments(0)

' Optional: regenerate a single fixture. Excel stamps a fresh timestamp
' and revision id on every save, so regenerating the whole set to add one
' file rewrites all eleven and buries the new one in the diff.
onlyOne = ""
If WScript.Arguments.Count > 1 Then onlyOne = LCase(WScript.Arguments(1))

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

Dim xl
Set xl = CreateObject("Excel.Application")
xl.DisplayAlerts = False
xl.Visible = False
xl.ScreenUpdating = False

Dim priorSheets
priorSheets = xl.SheetsInNewWorkbook
xl.SheetsInNewWorkbook = 1

' NOTE: the Windows default printer must be reachable before running this.
'
' Every PageSetup write goes through the active printer driver. With an
' offline default printer (a switched-off WSD network printer, say) Excel
' raises 1004 on every one of them *and* pops a modal "waiting for
' printer connection" dialog that DisplayAlerts does not suppress — so
' excel-pagesetup.xlsx comes out with nothing but its print area and the
' failure is easy to miss.
'
' Setting `Application.ActivePrinter` from here does not help: it fails
' with the same 1004. The default printer has to be changed at the OS
' level *before* Excel starts, e.g. from PowerShell:
'
'   (Get-WmiObject Win32_Printer -Filter "Name='Microsoft Print to PDF'").SetDefaultPrinter()
'
' The MakePageSetup failures are reported at the end of the run, so a
' silent bad fixture is not possible — but the fix is out here, not in
' the script.

' Everything from here on runs with errors trapped, so that a failure
' leaves no orphan EXCEL.EXE behind. Failures are collected and echoed.
Dim problems
problems = ""

On Error Resume Next

If Want("excel-basic.xlsx") Then MakeBasic outDir & "\excel-basic.xlsx", xlOpenXMLWorkbook
If Want("excel-basic.xls") Then MakeBasic outDir & "\excel-basic.xls", xlExcel8
If Want("excel-basic.xlsb") Then MakeBasic outDir & "\excel-basic.xlsb", xlExcel12
If Want("excel-strings.xlsx") Then MakeStrings outDir & "\excel-strings.xlsx", xlOpenXMLWorkbook
If Want("excel-strings.xlsb") Then MakeStrings outDir & "\excel-strings.xlsb", xlExcel12
If Want("excel-styled.xlsx") Then MakeStyled outDir & "\excel-styled.xlsx", xlOpenXMLWorkbook
If Want("excel-layout.xlsx") Then MakeLayout outDir & "\excel-layout.xlsx", xlOpenXMLWorkbook
If Want("excel-pagesetup.xlsx") Then MakePageSetup outDir & "\excel-pagesetup.xlsx", xlOpenXMLWorkbook
If Want("excel-styleonly.xlsx") Then MakeStyleOnly outDir & "\excel-styleonly.xlsx", xlOpenXMLWorkbook
If Want("excel-dates.xls") Then MakeDates outDir & "\excel-dates.xls", xlExcel8
If Want("excel-empty.xlsx") Then MakeEmpty outDir & "\excel-empty.xlsx", xlOpenXMLWorkbook
If Want("excel-chartsheet.xlsx") Then MakeChartsheet outDir & "\excel-chartsheet.xlsx", xlOpenXMLWorkbook
If Want("excel-sparse.xlsx") Then MakeSparse outDir & "\excel-sparse.xlsx", xlOpenXMLWorkbook

On Error Goto 0

xl.SheetsInNewWorkbook = priorSheets
' ActivePrinter is deliberately *not* restored: it is per-Excel-instance,
' this instance is about to quit, and assigning the offline printer back
' is exactly what pops the modal dialog we just worked around.
xl.Quit

If problems <> "" Then
  WScript.Echo "PROBLEMS:" & problems
  WScript.Quit 1
End If
If onlyOne = "" Then
  WScript.Echo "OK - wrote the full fixture set to " & outDir
Else
  WScript.Echo "OK - wrote " & onlyOne & " to " & outDir
End If

' ── helpers ─────────────────────────────────────────────────────────

Function Want(name)
  Want = (onlyOne = "") Or (onlyOne = LCase(name))
End Function

Sub Note(where)
  If Err.Number <> 0 Then
    problems = problems & vbCrLf & "  " & where & ": " & Err.Number & " " & Err.Description
    Err.Clear
  End If
End Sub

Function NewBook(sheetName)
  Dim wb
  Set wb = xl.Workbooks.Add
  Do While wb.Worksheets.Count > 1
    wb.Worksheets(wb.Worksheets.Count).Delete
  Loop
  wb.Worksheets(1).Name = sheetName
  ' Blank creator / lastModifiedBy in every container format. Must be set
  ' before SaveAs; see the header comment.
  wb.RemovePersonalInformation = True
  Set NewBook = wb
End Function

Sub Finish(wb, path, fmt)
  On Error Resume Next
  If fso.FileExists(path) Then fso.DeleteFile path
  wb.SaveAs path, fmt
  Note "SaveAs " & path
  wb.Close False
End Sub

' ── fixtures ────────────────────────────────────────────────────────

' The one sheet that exists in all three container formats, so that a
' difference between the XLSX, XLS and XLSB readers shows up as a
' difference in the golden model rather than being invisible.
Sub MakeBasic(path, fmt)
  On Error Resume Next
  Dim wb, ws
  Set wb = NewBook("Data")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "Name"
  ws.Range("B1").Value = "Qty"
  ws.Range("C1").Value = "Date"
  ws.Range("D1").Value = "Active"
  ws.Range("E1").Value = "Total"

  ws.Range("A2").Value = "Widget"
  ws.Range("B2").Value = 12
  ws.Range("C2").Value = DateSerial(2024, 3, 17)
  ws.Range("D2").Value = True
  ws.Range("E2").Formula = "=B2*2"

  ws.Range("A3").Value = "Gadget"
  ws.Range("B3").Value = -3.5
  ws.Range("C3").Value = DateSerial(1999, 12, 31)
  ws.Range("D3").Value = False
  ws.Range("E3").Formula = "=B3*2"

  ' A leap day, and a formula whose cached result is not an integer.
  ws.Range("A4").Value = "Doohickey"
  ws.Range("B4").Value = 0
  ws.Range("C4").Value = DateSerial(2024, 2, 29)
  ws.Range("D4").Value = True
  ws.Range("E4").Formula = "=SUM(B2:B4)"

  ' An error value with a cached result, and a string-valued formula —
  ' both are separate BIFF/BrtCell record types from a numeric one.
  ws.Range("A5").Value = "Broken"
  ws.Range("B5").Formula = "=1/0"
  ws.Range("C5").Formula = "=""x"" & ""y"""

  ws.Range("C2:C4").NumberFormat = "yyyy-mm-dd"
  Note "MakeBasic " & path
  Finish wb, path, fmt
End Sub

' String-table variety. Excel writes a shared string table and puts the
' xml:space="preserve" cases through it too — the #441 shape.
Sub MakeStrings(path, fmt)
  On Error Resume Next
  Dim wb, ws
  Set wb = NewBook("Strings")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "  leading"
  ws.Range("A2").Value = "trailing  "
  ws.Range("A3").Value = " both "
  ' A duplicate of A1, so the shared string table has a reused entry.
  ws.Range("A4").Value = "  leading"
  ws.Range("A5").Value = "plain"
  ' Non-ASCII across several scripts and a non-BMP codepoint (U+1F600),
  ' which is a surrogate pair in the UTF-16 the BIFF/XLSB tables use.
  '
  ' Built from ChrW rather than written literally on purpose: cscript
  ' reads a .vbs as the system ANSI codepage unless it is UTF-16, so a
  ' UTF-8 source file gets its non-ASCII literals mangled *before* Excel
  ' sees them. That produced a fixture full of mojibake that hucre then
  ' read back perfectly faithfully — a corpus bug wearing a reader bug's
  ' clothes. Keeping this file 7-bit clean makes the encoding a non-issue.
  ws.Range("A6").Value = "na" & ChrW(&HEF) & "ve " & _
    ChrW(&HFC) & "n" & ChrW(&HEF) & "code " & _
    ChrW(&H65E5) & ChrW(&H672C) & ChrW(&H8A9E) & " " & _
    ChrW(&H3A9) & ChrW(&H3BC) & ChrW(&H3AD) & ChrW(&H3B3) & ChrW(&H3B1) & " " & _
    ChrW(&HD83D) & ChrW(&HDE00)
  ' XML metacharacters, which have to survive escaping in xlsx/xlsb.
  ws.Range("A7").Value = "a & b < c > d ""quoted"" 'apos'"
  ' An embedded newline.
  ws.Range("A8").Value = "line one" & Chr(10) & "line two"
  ' A tab and a non-breaking space.
  ws.Range("A9").Value = "tab" & Chr(9) & "sep" & ChrW(160) & "nbsp"
  ws.Range("A10").Value = "  "

  ' A rich string: one run bold, one not. In BIFF/XLSB this is an
  ' RSTRING / rich shared-string entry rather than a plain one.
  ws.Range("B1").Value = "halfbold"
  ws.Range("B1").Characters(1, 4).Font.Bold = True

  Note "MakeStrings " & path
  Finish wb, path, fmt
End Sub

Sub MakeStyled(path, fmt)
  On Error Resume Next
  Dim wb, ws
  Set wb = NewBook("Styled")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "bold"
  ws.Range("A1").Font.Bold = True
  ws.Range("B1").Value = "italic"
  ws.Range("B1").Font.Italic = True
  ws.Range("C1").Value = "underline"
  ws.Range("C1").Font.Underline = 2   ' xlUnderlineStyleSingle
  ws.Range("D1").Value = "courier 14"
  ws.Range("D1").Font.Name = "Courier New"
  ws.Range("D1").Font.Size = 14
  ws.Range("E1").Value = "red"
  ws.Range("E1").Font.Color = RGB(255, 0, 0)

  ws.Range("A2").Value = "yellow fill"
  ws.Range("A2").Interior.Color = RGB(255, 255, 0)
  ws.Range("B2").Value = "bordered"
  ws.Range("B2").Borders.LineStyle = 1     ' xlContinuous
  ws.Range("B2").Borders.Weight = 2        ' xlThin

  ws.Range("C2").Value = 0.125
  ws.Range("C2").NumberFormat = "0.00%"
  ws.Range("D2").Value = 1234.5
  ws.Range("D2").NumberFormat = "#,##0.00"
  ws.Range("E2").Value = DateSerial(2024, 3, 17)
  ws.Range("E2").NumberFormat = "yyyy-mm-dd"

  ws.Range("A3").Value = "centred"
  ws.Range("A3").HorizontalAlignment = -4108   ' xlCenter
  ws.Range("A3").VerticalAlignment = -4108
  ws.Range("B3").Value = "wrapped text that is long enough to wrap"
  ws.Range("B3").WrapText = True

  ' A format applied to a whole column with nothing in it — the column
  ' carries the style, no cell does.
  ws.Columns("G").NumberFormat = "0.000"
  ws.Columns("G").ColumnWidth = 14

  Note "MakeStyled " & path
  Finish wb, path, fmt
End Sub

Sub MakeLayout(path, fmt)
  On Error Resume Next
  Dim wb, ws
  Set wb = NewBook("Layout")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "merged across"
  ws.Range("A1:C1").Merge
  ws.Range("D1").Value = "merged down"
  ws.Range("D1:D3").Merge

  ws.Range("A2").Value = "Item"
  ws.Range("B2").Value = "Value"
  ws.Range("A3").Value = "one"
  ws.Range("B3").Value = 5
  ws.Range("A4").Value = "two"
  ws.Range("B4").Value = 15
  ws.Range("A5").Value = "three"
  ws.Range("B5").Value = 25

  ' Conditional format: value > 10 gets a red fill.
  Dim fc
  Set fc = ws.Range("B3:B5").FormatConditions.Add(1, 5, "=10")  ' xlCellValue, xlGreater
  fc.Interior.Color = RGB(255, 199, 206)
  Note "MakeLayout conditional format"

  ws.Activate
  xl.ActiveWindow.SplitRow = 2
  xl.ActiveWindow.SplitColumn = 1
  xl.ActiveWindow.FreezePanes = True
  Note "MakeLayout freeze panes"

  ws.Rows(6).Hidden = True
  ws.Columns("F").Hidden = True

  Note "MakeLayout " & path
  Finish wb, path, fmt
End Sub

Sub MakePageSetup(path, fmt)
  On Error Resume Next
  Dim wb, ws
  Set wb = NewBook("Printed")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "Header row"
  ws.Range("A2").Value = "body"
  ws.Range("B1").Value = "Second"
  ws.Range("B2").Value = 42

  ' PageSetup writes go through the printer driver and are slow; batching
  ' them under PrintCommunication = False is Microsoft's own advice.
  xl.PrintCommunication = False
  ws.PageSetup.PaperSize = 8            ' xlPaperA3
  ws.PageSetup.Orientation = 2          ' xlLandscape
  ws.PageSetup.PrintTitleRows = "$1:$1"
  ws.PageSetup.PrintTitleColumns = "$A:$A"
  ws.PageSetup.LeftMargin = 36
  xl.PrintCommunication = True
  Note "MakePageSetup page setup"

  ' Zoom has to be off *and* committed to the driver before FitToPages*
  ' will take; batching them together fails with 1004.
  ws.PageSetup.Zoom = False
  ws.PageSetup.FitToPagesWide = 1
  ws.PageSetup.FitToPagesTall = 2
  Note "MakePageSetup fit to pages"

  ws.PageSetup.PrintArea = "$A$1:$B$2"
  Note "MakePageSetup print area"

  Note "MakePageSetup " & path
  Finish wb, path, fmt
End Sub

' #492: whole rows carrying a format with no values out to the right.
' The blow-up shape is a reader that materialises a cell per formatted
' column, so the sheet has to be formatted well past its used range.
Sub MakeStyleOnly(path, fmt)
  On Error Resume Next
  Dim wb, ws
  Set wb = NewBook("StyleOnly")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "a"
  ws.Range("B1").Value = "b"
  ws.Range("C1").Value = "c"
  ws.Range("A2").Value = 1
  ws.Range("B2").Value = 2
  ws.Range("C2").Value = 3

  ' Entire rows formatted — the row records carry a style, and past
  ' column C there is nothing else to carry it.
  ws.Rows("1:1").Font.Bold = True
  ws.Rows("2:4").Interior.Color = RGB(221, 235, 247)
  ws.Rows("2:4").NumberFormat = "0.00"

  Note "MakeStyleOnly " & path
  Finish wb, path, fmt
End Sub

' Date and number formats in BIFF8. The CJK codes here are *custom*
' FORMAT records (ifmt >= 164) with locale prefixes, not the CJK built-in
' ids 27-36 / 50-58 that #444 is about: a non-CJK Windows Excel will not
' emit those ids, which was checked by dumping the FORMAT records of a
' probe file. See PROVENANCE.md.
Sub MakeDates(path, fmt)
  On Error Resume Next
  Dim wb, ws, d
  Set wb = NewBook("Dates")
  Set ws = wb.Worksheets(1)
  d = DateSerial(2024, 3, 17)

  ws.Range("A1").Value = "iso"
  ws.Range("B1").Value = d
  ws.Range("B1").NumberFormat = "yyyy-mm-dd"

  ws.Range("A2").Value = "builtin 14"
  ws.Range("B2").Value = d
  ws.Range("B2").NumberFormat = "m/d/yyyy"

  ws.Range("A3").Value = "builtin 15"
  ws.Range("B3").Value = d
  ws.Range("B3").NumberFormat = "d-mmm-yy"

  ws.Range("A4").Value = "japanese"
  ws.Range("B4").Value = d
  ws.Range("B4").NumberFormat = "[$-411]yyyy""" & ChrW(&H5E74) & """m""" & ChrW(&H6708) & """d""" & ChrW(&H65E5) & """"

  ws.Range("A5").Value = "japanese era"
  ws.Range("B5").Value = d
  ws.Range("B5").NumberFormat = "[$-404]e/m/d"

  ws.Range("A6").Value = "korean"
  ws.Range("B6").Value = d
  ws.Range("B6").NumberFormat = "[$-412]yyyy""" & ChrW(&HB144) & """ mm""" & ChrW(&HC6D4) & """ dd""" & ChrW(&HC77C) & """"

  ws.Range("A7").Value = "time"
  ws.Range("B7").Value = TimeSerial(13, 45, 30)
  ws.Range("B7").NumberFormat = "hh:mm:ss"

  ws.Range("A8").Value = "datetime"
  ws.Range("B8").Value = DateSerial(2024, 3, 17) + TimeSerial(13, 45, 30)
  ws.Range("B8").NumberFormat = "yyyy-mm-dd hh:mm"

  ws.Range("A9").Value = "thousands"
  ws.Range("B9").Value = 1234567.891
  ws.Range("B9").NumberFormat = "#,##0.00"

  ws.Range("A10").Value = "percent"
  ws.Range("B10").Value = 0.125
  ws.Range("B10").NumberFormat = "0.00%"

  ws.Range("A11").Value = "text format"
  ws.Range("B11").NumberFormat = "@"
  ws.Range("B11").Value = "0042"

  Note "MakeDates " & path
  Finish wb, path, fmt
End Sub

Sub MakeEmpty(path, fmt)
  On Error Resume Next
  Dim wb
  Set wb = NewBook("Empty")
  Note "MakeEmpty " & path
  Finish wb, path, fmt
End Sub

' #499 — a chart on its own tab, not a ChartObject floating on a sheet.
'
' xl/workbook.xml's <sheets> lists every sheet whatever its kind; the
' relationship type is what says which is which, and a chart sheet's is
' .../chartsheet pointing at xl/chartsheets/. A reader that builds its
' sheet map from `worksheet` relationships alone finds no part for that
' rId. hucre threw `Invalid XLSX: missing worksheet file for sheet` and
' refused the entire workbook, ordinary worksheets included.
'
' Found by running hucre over a corpus of real instrument-exported
' workbooks: 52 of 538 failed on exactly this, which is why a chart
' sheet is worth a fixture of its own.
Sub MakeChartsheet(path, fmt)
  On Error Resume Next
  Dim wb, ws, ch
  Set wb = NewBook("Data")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "x"
  ws.Range("B1").Value = "y"
  ws.Range("A2").Value = 1
  ws.Range("B2").Value = 10
  ws.Range("A3").Value = 2
  ws.Range("B3").Value = 20
  ws.Range("A4").Value = 3
  ws.Range("B4").Value = 15

  ' Charts.Add makes a chart SHEET. ChartObjects.Add would make an
  ' embedded chart, which is a different thing and already covered.
  Set ch = wb.Charts.Add
  ch.Name = "Diagram"
  ch.ChartType = 4          ' xlLine
  ch.SetSourceData ws.Range("A1:B4")
  Note "MakeChartsheet chart sheet"

  ' The sheet order matters: the chart sheet sits before the worksheet,
  ' so a reader that gives up on the first unresolvable sheet never
  ' reaches the data.
  ch.Move wb.Worksheets(1)
  Note "MakeChartsheet order"

  Note "MakeChartsheet " & path
  Finish wb, path, fmt
End Sub

' #501 — a sheet that is sparse rather than large.
'
' `Sheet.rows` is a dense rectangle, so what a reader must allocate is
' the bounding box, not the cell count. A handful of values placed far
' to the right makes that box enormous while the file stays tiny: here
' about thirty values describe a 2,000 x 15,312 box, 30.6 million slots,
' past the 20,000,000 MAX_TOTAL_CELLS bound. The workbook is refused.
'
' The real file this came from was worse and less contrived: 76,277
' values across 507 distinct columns reaching column 15,311 on a
' 19,959-row sheet — 305,612,208 slots at a fill factor of 0.03%. Excel
' opens it without complaint. Neither `range`, `maxRows` nor raising
' `maxTotalCells` gives a usable way to read it.
Sub MakeSparse(path, fmt)
  On Error Resume Next
  Dim wb, ws, i
  Set wb = NewBook("Sparse")
  Set ws = wb.Worksheets(1)

  ws.Range("A1").Value = "left edge"
  ws.Range("B1").Value = 1
  ws.Range("C1").Value = 2

  ' A few islands of data far to the right. Column 15312 is VPX, which
  ' is well inside Excel's 16,384 — nothing here is out of spec.
  For i = 1 To 8
    ws.Cells(i, 5000).Value = "island A " & i
    ws.Cells(i, 10000).Value = i * 10
    ws.Cells(i, 15312).Value = "right edge " & i
  Next

  ' Push the used range down so rows x cols crosses the bound. One cell
  ' does it; the file stays a few KB because only real cells are stored.
  ws.Cells(2000, 15312).Value = "bottom right"

  Note "MakeSparse " & path
  Finish wb, path, fmt
End Sub
