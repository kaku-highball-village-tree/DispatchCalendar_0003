Option Explicit

Private m_iSelectedYear As Long
Private m_iSelectedMonth As Long
Private m_bAccepted As Boolean

Public Property Get SelectedYear() As Long
    SelectedYear = m_iSelectedYear
End Property

Public Property Get SelectedMonth() As Long
    SelectedMonth = m_iSelectedMonth
End Property

Public Property Get IsAccepted() As Boolean
    IsAccepted = m_bAccepted
End Property

Private Sub UserForm_Initialize()
    m_iSelectedYear = Year(Date)
    m_iSelectedMonth = Month(Date)
    m_bAccepted = False
    UpdateMonthText
End Sub

Private Sub cmdPreviousMonth_Click()
    If m_iSelectedMonth = 1 Then
        m_iSelectedYear = m_iSelectedYear - 1
        m_iSelectedMonth = 12
    Else
        m_iSelectedMonth = m_iSelectedMonth - 1
    End If
    UpdateMonthText
End Sub

Private Sub cmdNextMonth_Click()
    If m_iSelectedMonth = 12 Then
        m_iSelectedYear = m_iSelectedYear + 1
        m_iSelectedMonth = 1
    Else
        m_iSelectedMonth = m_iSelectedMonth + 1
    End If
    UpdateMonthText
End Sub

Private Sub cmdOK_Click()
    m_bAccepted = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    m_bAccepted = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    m_bAccepted = False
End Sub

Private Sub UpdateMonthText()
    lblSelectedMonth.Caption = CStr(m_iSelectedYear) & "年" & Format$(m_iSelectedMonth, "00") & "月"
End Sub
