Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const GOLDEN_RATIO As Double = 1.618
Private Const MIN_FONT_SIZE As Double = 12
Private Const MAX_FONT_SIZE As Double = 24

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
    ResizeFormToDesktopRatio
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


Private Sub ResizeFormToDesktopRatio()
    Dim dFormWidth As Double
    Dim dFormHeight As Double
    Dim dBaseFontSize As Double

    dFormWidth = PixelsToPoints(GetSystemMetrics(SM_CXSCREEN)) / 2
    dFormHeight = dFormWidth / GOLDEN_RATIO
    dBaseFontSize = ClampDouble(dFormWidth / 45, MIN_FONT_SIZE, MAX_FONT_SIZE)

    Me.Width = dFormWidth
    Me.Height = dFormHeight

    ApplyResponsiveFonts dBaseFontSize
    ArrangeResponsiveControls dFormWidth, dFormHeight
End Sub

Private Sub ApplyResponsiveFonts(ByVal baseFontSize As Double)
    lblMessage.Font.Size = baseFontSize
    lblSelectedMonth.Font.Size = baseFontSize * 1.2
    cmdPreviousMonth.Font.Size = baseFontSize
    cmdNextMonth.Font.Size = baseFontSize
    cmdOK.Font.Size = baseFontSize
    cmdCancel.Font.Size = baseFontSize
End Sub

Private Sub ArrangeResponsiveControls(ByVal formWidth As Double, ByVal formHeight As Double)
    Dim dMargin As Double
    Dim dMessageHeight As Double
    Dim dNavigationTop As Double
    Dim dNavigationHeight As Double
    Dim dNavigationLeft As Double
    Dim dMonthWidth As Double
    Dim dButtonWidth As Double
    Dim dButtonHeight As Double
    Dim dSpacing As Double
    Dim dActionTop As Double
    Dim dActionButtonWidth As Double
    Dim dActionButtonHeight As Double
    Dim dActionLeft As Double

    dMargin = formWidth * 0.08
    dMessageHeight = formHeight * 0.12
    dNavigationTop = formHeight * 0.38
    dNavigationHeight = formHeight * 0.13
    dButtonWidth = formWidth * 0.14
    dMonthWidth = formWidth * 0.32
    dSpacing = formWidth * 0.04
    dButtonHeight = dNavigationHeight
    dNavigationLeft = (formWidth - ((dButtonWidth * 2) + dMonthWidth + (dSpacing * 2))) / 2

    lblMessage.Left = dMargin
    lblMessage.Top = formHeight * 0.18
    lblMessage.Width = formWidth - (dMargin * 2)
    lblMessage.Height = dMessageHeight

    cmdPreviousMonth.Left = dNavigationLeft
    cmdPreviousMonth.Top = dNavigationTop
    cmdPreviousMonth.Width = dButtonWidth
    cmdPreviousMonth.Height = dButtonHeight

    lblSelectedMonth.Left = cmdPreviousMonth.Left + cmdPreviousMonth.Width + dSpacing
    lblSelectedMonth.Top = dNavigationTop + (dNavigationHeight * 0.15)
    lblSelectedMonth.Width = dMonthWidth
    lblSelectedMonth.Height = dNavigationHeight * 0.7

    cmdNextMonth.Left = lblSelectedMonth.Left + lblSelectedMonth.Width + dSpacing
    cmdNextMonth.Top = dNavigationTop
    cmdNextMonth.Width = dButtonWidth
    cmdNextMonth.Height = dButtonHeight

    dActionButtonWidth = formWidth * 0.24
    dActionButtonHeight = formHeight * 0.16
    dActionTop = formHeight * 0.68
    dActionLeft = (formWidth - ((dActionButtonWidth * 2) + dSpacing)) / 2

    cmdOK.Left = dActionLeft
    cmdOK.Top = dActionTop
    cmdOK.Width = dActionButtonWidth
    cmdOK.Height = dActionButtonHeight

    cmdCancel.Left = cmdOK.Left + cmdOK.Width + dSpacing
    cmdCancel.Top = dActionTop
    cmdCancel.Width = dActionButtonWidth
    cmdCancel.Height = dActionButtonHeight
End Sub

Private Function PixelsToPoints(ByVal pixelValue As Long) As Double
    PixelsToPoints = CDbl(pixelValue) * 72 / GetScreenDpiX()
End Function

Private Function GetScreenDpiX() As Long
#If VBA7 Then
    Dim hDesktopDC As LongPtr
#Else
    Dim hDesktopDC As Long
#End If

    hDesktopDC = GetDC(0)
    If hDesktopDC <> 0 Then
        GetScreenDpiX = GetDeviceCaps(hDesktopDC, LOGPIXELSX)
        ReleaseDC 0, hDesktopDC
    End If

    If GetScreenDpiX <= 0 Then
        GetScreenDpiX = 96
    End If
End Function

Private Function ClampDouble(ByVal value As Double, ByVal minimumValue As Double, ByVal maximumValue As Double) As Double
    If value < minimumValue Then
        ClampDouble = minimumValue
    ElseIf value > maximumValue Then
        ClampDouble = maximumValue
    Else
        ClampDouble = value
    End If
End Function
