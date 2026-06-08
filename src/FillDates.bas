Attribute VB_Name = "FillDates"
Option Explicit

' B1 に開始日が入っている場合、B1:AE1 の空白セルだけに連続日付を入力します。
' 例: B1 = 2026/6/1 のとき、C1:AE1 に 2026/6/2 ～ 2026/6/30 を入力します。
Public Sub FillBlankDatesFromB1ToAE1()
    FillBlankDatesInRow ActiveSheet.Range("B1"), ActiveSheet.Range("AE1")
End Sub

' startCell から endCell まで、startCell の日付を起点に連続日付で空白セルを埋めます。
' 既に値が入っているセルは上書きしません。
Public Sub FillBlankDatesInRow(ByVal startCell As Range, ByVal endCell As Range)
    Dim dtStartDate As Date
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
        If Len(rngTargetCell.Value) = 0 Then
            lDayOffset = rngTargetCell.Column - startCell.Column
            rngTargetCell.Value = DateAdd("d", lDayOffset, dtStartDate)
            rngTargetCell.NumberFormatLocal = "yyyy/m/d"
        End If
    Next rngTargetCell
End Sub
