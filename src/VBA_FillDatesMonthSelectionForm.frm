VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBA_FillDatesMonthSelectionForm 
   Caption         =   "Ώ۔NI"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   StartUpPosition =   1  'I[i[ tH[̒
   Begin MSForms.CommandButton cmdCancel 
      Caption         =   "LZ"
      Height          =   360
      Left            =   2280
      TabIndex        =   5
      Top             =   1560
      Width           =   1200
   End
   Begin MSForms.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   1200
   End
   Begin MSForms.CommandButton cmdNextMonth 
      Caption         =   ""
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Top             =   900
      Width           =   720
   End
   Begin MSForms.Label lblSelectedMonth 
      Alignment       =   2  '
      Caption         =   "0000N00"
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1440
   End
   Begin MSForms.CommandButton cmdPreviousMonth 
      Caption         =   ""
      Height          =   360
      Left            =   600
      TabIndex        =   1
      Top             =   900
      Width           =   720
   End
   Begin MSForms.Label lblMessage 
      Alignment       =   2  '
      Caption         =   "Ώ۔NIĂ"
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3360
   End
End
Attribute VB_Name = "VBA_FillDatesMonthSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    lblSelectedMonth.Caption = CStr(m_iSelectedYear) & "N" & Format$(m_iSelectedMonth, "00") & ""
End Sub
