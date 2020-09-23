VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmReportMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Menu"
   ClientHeight    =   3945
   ClientLeft      =   5235
   ClientTop       =   3795
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4545
   Begin Crystal.CrystalReport rptAnnSalExp 
      Left            =   1800
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   2880
      TabIndex        =   7
      Top             =   2700
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   675
      Left            =   120
      TabIndex        =   6
      Top             =   2700
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   4275
      Begin VB.OptionButton optDestination 
         Caption         =   "Printer"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optDestination 
         Caption         =   "Window"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4275
      Begin VB.OptionButton optReport 
         Caption         =   "Annual Salary Expenses by &Job"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3315
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Annual Salary Expenses by &Department"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmReportMenu"
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
Private Sub cmdOK_Click()
'------------------------------------------------------------------------

    On Error GoTo cmdOK_Click_Error
    
    Dim strReportName           As String
    Dim intReportDestination    As Integer
    
    If optReport(0).Value = True Then
        strReportName = "SALDEPT.RPT"
    Else
        strReportName = "SALJOB.RPT"
    End If
    
    If optDestination(0).Value = True Then
        intReportDestination = crptToWindow
    Else
        intReportDestination = crptToPrinter
    End If
    
    With rptAnnSalExp
        .ReportFileName = GetAppPath() & strReportName
        .DataFiles(0) = GetAppPath() & "EMPLOYEE.MDB"
        .Destination = intReportDestination
        .Action = 1                ' 1 = "Run the Report"
    End With
    
    Exit Sub

cmdOK_Click_Error:

    MsgBox "The following error has occurred:" & vbNewLine _
         & Err.Number & " - " & Err.Description, _
           vbCritical, _
           "cmdOK_Click"

End Sub

'------------------------------------------------------------------------
Private Sub cmdExit_Click()
'------------------------------------------------------------------------
    Unload Me
End Sub
