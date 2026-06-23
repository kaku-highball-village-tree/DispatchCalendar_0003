Attribute VB_Name = "VBA_FillDates"
Option Explicit

' UserForm で対象年月を選択し、新規シートに選択月の日付と連番を入力します。
Public Sub SelectMonthAndFillDatesFromB1ToAF1()
    Dim objMonthSelectionForm As VBA_FillDatesMonthSelectionForm
    Dim dtSelectedDate As Date
    Dim objTargetWorkbook As Workbook
    Dim objTargetWorksheet As Worksheet
    Dim strWorksheetName As String
    Dim strOutputFilePath As String

    Set objMonthSelectionForm = New VBA_FillDatesMonthSelectionForm
    objMonthSelectionForm.Show vbModal

    If objMonthSelectionForm.IsAccepted Then
        dtSelectedDate = DateSerial(objMonthSelectionForm.SelectedYear, objMonthSelectionForm.SelectedMonth, 1)
        strWorksheetName = Format$(dtSelectedDate, "yyyy年mm月")
        Set objTargetWorkbook = ActiveWorkbook

        If WorksheetExists(strWorksheetName, objTargetWorkbook) Then
            MsgBox "シート「" & strWorksheetName & "」は既に存在します。", vbExclamation, "対象年月シート作成"
        ElseIf Len(objTargetWorkbook.Path) = 0 Then
            MsgBox "元の .xlsm ファイルを保存してから実行してください。", vbExclamation, "月別ファイル保存"
        Else
            strOutputFilePath = BuildMonthlyWorkbookPath(objTargetWorkbook, strWorksheetName)

            If Len(Dir$(strOutputFilePath)) > 0 Then
                strOutputFilePath = BuildTimestampedMonthlyWorkbookPath(objTargetWorkbook, strWorksheetName)
            End If

            If Len(Dir$(strOutputFilePath)) > 0 Then
                MsgBox "ファイル「" & strOutputFilePath & "」は既に存在します。", vbExclamation, "月別ファイル保存"
            Else
                Set objTargetWorksheet = objTargetWorkbook.Worksheets.Add(After:=objTargetWorkbook.Worksheets(objTargetWorkbook.Worksheets.Count))
                objTargetWorksheet.Name = strWorksheetName
                FillMonthDatesAndSequence objTargetWorksheet, dtSelectedDate
                SaveWorksheetAsXlsx objTargetWorksheet, strOutputFilePath
            End If
        End If
    End If

    Unload objMonthSelectionForm
End Sub

' B1 に開始日が入っている場合、B1:AF1 の完全に空のセルだけに同じ月の連続日付を入力します。
' 例: B1 = 2026/7/1 のとき、C1:AF1 に 2026/7/2 ～ 2026/7/31 を入力します。
' 例: B1 = 2026/9/1 のとき、C1:AE1 に 2026/9/2 ～ 2026/9/30 を入力し、AF1 は空白のままにします。
Public Sub FillBlankDatesFromB1ToAF1()
    FillBlankDatesInRow ActiveSheet.Range("B1"), ActiveSheet.Range("AF1")
End Sub

' startCell から endCell まで、startCell の日付を起点に連続日付で完全に空のセルを埋めます。
' 既に値が入っているセル、数式が入っているセル、数式の結果が空文字のセルは上書きしません。
' startCell と同じ年月の日付だけを入力し、翌月以降の日付は入力しません。
Public Sub FillBlankDatesInRow(ByVal startCell As Range, ByVal endCell As Range)
    Dim dtStartDate As Date
    Dim dtTargetDate As Date
    Dim rngTargetCell As Range
    Dim lDayOffset As Long

    If startCell.Worksheet.Name <> endCell.Worksheet.Name Then
        Err.Raise vbObjectError + 1000, "FillBlankDatesInRow", "startCell と endCell は同じワークシート上に指定してください。"
    End If

    If startCell.Row <> endCell.Row Then
        Err.Raise vbObjectError + 1001, "FillBlankDatesInRow", "startCell と endCell は同じ行に指定してください。"
    End If

    If endCell.Column < startCell.Column Then
        Err.Raise vbObjectError + 1002, "FillBlankDatesInRow", "endCell は startCell より右側のセルを指定してください。"
    End If

    If Not IsDate(startCell.Value) Then
        Err.Raise vbObjectError + 1003, "FillBlankDatesInRow", "startCell には日付を入力してください。"
    End If

    dtStartDate = CDate(startCell.Value)

    For Each rngTargetCell In startCell.Worksheet.Range(startCell, endCell).Cells
        lDayOffset = rngTargetCell.Column - startCell.Column
        dtTargetDate = DateAdd("d", lDayOffset, dtStartDate)

        If Year(dtTargetDate) = Year(dtStartDate) And Month(dtTargetDate) = Month(dtStartDate) Then
            If IsEmpty(rngTargetCell.Value) Then
                rngTargetCell.Value = dtTargetDate
                rngTargetCell.NumberFormatLocal = "yyyy/m/d"
            End If
        End If
    Next rngTargetCell
End Sub

Private Sub FillMonthDatesAndSequence(ByVal targetWorksheet As Worksheet, ByVal firstDateOfMonth As Date)
    Dim dtLastDateOfMonth As Date
    Dim lDayCount As Long
    Dim lDayIndex As Long

    dtLastDateOfMonth = DateSerial(Year(firstDateOfMonth), Month(firstDateOfMonth) + 1, 0)
    lDayCount = Day(dtLastDateOfMonth)

    For lDayIndex = 1 To lDayCount
        targetWorksheet.Cells(1, lDayIndex + 1).Value = DateSerial(Year(firstDateOfMonth), Month(firstDateOfMonth), lDayIndex)
        targetWorksheet.Cells(1, lDayIndex + 1).NumberFormatLocal = "yyyy""年""mm""月""dd""日""(aaa)"
        targetWorksheet.Cells(2 + ((lDayIndex - 1) * 3), 1).Value = lDayIndex
    Next lDayIndex
End Sub

Private Function WorksheetExists(ByVal worksheetName As String, ByVal targetWorkbook As Workbook) As Boolean
    Dim objWorksheet As Worksheet

    For Each objWorksheet In targetWorkbook.Worksheets
        If objWorksheet.Name = worksheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next objWorksheet

    WorksheetExists = False
End Function


Private Function BuildMonthlyWorkbookPath(ByVal targetWorkbook As Workbook, ByVal worksheetName As String) As String
    BuildMonthlyWorkbookPath = targetWorkbook.Path & Application.PathSeparator & "配車カレンダー_" & worksheetName & ".xlsx"
End Function


Private Function BuildTimestampedMonthlyWorkbookPath(ByVal targetWorkbook As Workbook, ByVal worksheetName As String) As String
    BuildTimestampedMonthlyWorkbookPath = targetWorkbook.Path & Application.PathSeparator & "配車カレンダー_" & worksheetName & "_" & Format$(Now, "yyyy_mm_dd_hh_nn_ss") & ".xlsx"
End Function

Private Sub SaveWorksheetAsXlsx(ByVal sourceWorksheet As Worksheet, ByVal outputFilePath As String)
    Dim objOutputWorkbook As Workbook

    sourceWorksheet.Copy
    Set objOutputWorkbook = ActiveWorkbook
    objOutputWorkbook.SaveAs Filename:=outputFilePath, FileFormat:=xlOpenXMLWorkbook
    objOutputWorkbook.Close SaveChanges:=False
End Sub
