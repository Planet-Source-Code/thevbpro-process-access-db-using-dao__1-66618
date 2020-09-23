VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Database Maintenance - Main Menu"
   ClientHeight    =   6135
   ClientLeft      =   4035
   ClientTop       =   2040
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6555
   Begin VB.CommandButton cmdMainMenuOpt 
      Caption         =   "&Help"
      Height          =   612
      Index           =   4
      Left            =   1873
      TabIndex        =   4
      Top             =   3970
      Width           =   2772
   End
   Begin VB.CommandButton cmdMainMenuOpt 
      Caption         =   "E&xit"
      Height          =   612
      Index           =   5
      Left            =   1873
      TabIndex        =   5
      Top             =   4780
      Width           =   2772
   End
   Begin VB.CommandButton cmdMainMenuOpt 
      Caption         =   "&Report Menu"
      Height          =   612
      Index           =   3
      Left            =   1909
      TabIndex        =   3
      Top             =   3163
      Width           =   2772
   End
   Begin VB.CommandButton cmdMainMenuOpt 
      Caption         =   "&Job Maintenance"
      Height          =   612
      Index           =   2
      Left            =   1909
      TabIndex        =   2
      Top             =   2356
      Width           =   2772
   End
   Begin VB.CommandButton cmdMainMenuOpt 
      Caption         =   "&Department Maintenance"
      Height          =   612
      Index           =   1
      Left            =   1909
      TabIndex        =   1
      Top             =   1549
      Width           =   2772
   End
   Begin VB.CommandButton cmdMainMenuOpt 
      Caption         =   "&Employee Maintenance"
      Height          =   612
      Index           =   0
      Left            =   1909
      TabIndex        =   0
      Top             =   742
      Width           =   2772
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------------------
    CenterForm Me
End Sub

'------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------------------

    If KeyCode = vbKeyF1 Then
        ShowHelpForm
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdMainMenuOpt_Click(Index As Integer)
'------------------------------------------------------------------------
    Select Case Index
        Case 0
            frmEmpMaint.Show vbModal
        Case 1
            frmDeptMaint.Show vbModal
        Case 2
            frmJobMaint.Show vbModal
        Case 3
            frmReportMenu.Show vbModal
        Case 4
            ShowHelpForm
        Case 5
            End
    End Select

End Sub

'------------------------------------------------------------------------
Private Sub ShowHelpForm()
'------------------------------------------------------------------------

    gintHelpFileNbr = 1
    frmHelp.Show vbModal

End Sub
