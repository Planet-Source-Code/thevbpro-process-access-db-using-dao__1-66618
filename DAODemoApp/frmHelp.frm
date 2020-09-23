VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Database Maintenance - Help"
   ClientHeight    =   5850
   ClientLeft      =   2130
   ClientTop       =   1005
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7440
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   5040
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   300
      TabIndex        =   1
      Top             =   5040
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox rtbHelp 
      Height          =   4635
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8176
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmHelp.frx":0000
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------------------

    Dim strHelpFileName As String
    
    CenterForm Me
    
    strHelpFileName = GetAppPath & "EDMHELP" & gintHelpFileNbr & ".DOC"
    
    rtbHelp.LoadFile strHelpFileName, rtfRTF

End Sub

'------------------------------------------------------------------------
Private Sub cmdOK_Click()
'------------------------------------------------------------------------

    Unload Me
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdAbout_Click()
'------------------------------------------------------------------------

    MsgBox "DAO (Data Access Objects) Demo" & vbNewLine _
         & "Employee Database Maintenance" & vbNewLine _
         & "Copyright " & Chr$(169) & " 2000-2005 thevbprogrammer.com", _
            vbInformation, _
            "About"

End Sub

