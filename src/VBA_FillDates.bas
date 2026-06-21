Attribute VB_Name = "VBA_FillDates"
Option Explicit

' UserForm で対象年月を選択し、B1 に選択年月の1日を入力して B1:AF1 の日付を埋めます。
Public Sub SelectMonthAndFillDatesFromB1ToAF1()
    Dim objMonthSelectionForm As VBA_FillDatesMonthSelectionForm
    Dim dtSelectedDate As Date

    Set objMonthSelectionForm = New VBA_FillDatesMonthSelectionForm
    objMonthSelectionForm.Show vbModal

    If objMonthSelectionForm.IsAccepted Then
        dtSelectedDate = DateSerial(objMonthSelectionForm.SelectedYear, objMonthSelectionForm.SelectedMonth, 1)
        ActiveSheet.Range("B1").Value = dtSelectedDate
        ActiveSheet.Range("B1").NumberFormatLocal = "yyyy/m/d"
        FillBlankDatesFromB1ToAF1
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
