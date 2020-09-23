VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   1860
   ClientTop       =   1365
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSplash 
      Interval        =   3000
      Left            =   7200
      Top             =   5760
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000-2005 thevbprogrammer.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   7812
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Database Maintenance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   7812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DAO (Data Access Objects) Demo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7812
   End
End
Attribute VB_Name = "frmSplash"
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
Private Sub tmrSplash_Timer()
'------------------------------------------------------------------------
    tmrSplash.Enabled = False
    frmMainMenu.Show
    Unload Me
End Sub
