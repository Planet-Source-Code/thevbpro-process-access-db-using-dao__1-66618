VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmEmpMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Database Maintenance - Employee Maintenance"
   ClientHeight    =   6135
   ClientLeft      =   3540
   ClientTop       =   1275
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9240
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   435
      Left            =   7980
      TabIndex        =   42
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   435
      Left            =   7980
      TabIndex        =   40
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo "
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   435
      Left            =   7980
      TabIndex        =   41
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Frame fraSearchOuter 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   3720
      TabIndex        =   29
      Top             =   3180
      Width           =   5355
      Begin VB.Frame fraSearchInner 
         BorderStyle     =   0  'None
         Height          =   1932
         Left            =   180
         TabIndex        =   28
         Top             =   240
         Width           =   5115
         Begin VB.TextBox txtCriteria 
            Height          =   345
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   4935
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find First"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find Next"
            Height          =   375
            Index           =   2
            Left            =   2676
            TabIndex        =   38
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find Prev"
            Height          =   375
            Index           =   1
            Left            =   1404
            TabIndex        =   37
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find Last"
            Height          =   375
            Index           =   3
            Left            =   3960
            TabIndex        =   39
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cboField 
            Height          =   315
            ItemData        =   "frmEmpMaint.frx":0000
            Left            =   120
            List            =   "frmEmpMaint.frx":001C
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   285
            Width           =   3375
         End
         Begin VB.ComboBox cboRelOp 
            Height          =   315
            ItemData        =   "frmEmpMaint.frx":0076
            Left            =   3600
            List            =   "frmEmpMaint.frx":008F
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "F&ield:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   60
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "&Comparison:"
            Height          =   255
            Left            =   3600
            TabIndex        =   32
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "&Value:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   1920
      TabIndex        =   24
      Top             =   3180
      Width           =   1695
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Record"
         Height          =   375
         Left            =   180
         TabIndex        =   27
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Record"
         Height          =   375
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update Record"
         Height          =   375
         Left            =   180
         TabIndex        =   26
         Top             =   990
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   240
      TabIndex        =   19
      Top             =   3180
      Width           =   1575
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Prev Record"
         Height          =   375
         Left            =   180
         TabIndex        =   21
         Top             =   760
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next Record"
         Height          =   375
         Left            =   180
         TabIndex        =   22
         Top             =   1220
         Width           =   1215
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Last Record"
         Height          =   375
         Left            =   180
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&First Record"
         Height          =   375
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Return to Main Menu"
      Height          =   375
      Left            =   6960
      TabIndex        =   44
      Top             =   5580
      Width           =   2115
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4740
      TabIndex        =   43
      Top             =   5580
      Width           =   2115
   End
   Begin VB.Frame fraEmpInfoOuter 
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   7635
      Begin VB.Frame fraEmpInfoInner 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2535
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   7455
         Begin VB.TextBox txtSchedHrs 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   5580
            MaxLength       =   6
            TabIndex        =   18
            Top             =   2100
            Width           =   1815
         End
         Begin VB.TextBox txtEmpLast 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3960
            MaxLength       =   35
            TabIndex        =   7
            Top             =   480
            Width           =   3435
         End
         Begin VB.TextBox txtEmpFirst 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   960
            MaxLength       =   35
            TabIndex        =   5
            Top             =   480
            Width           =   2775
         End
         Begin VB.ComboBox cboHrlyRate 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2760
            TabIndex        =   16
            Top             =   2100
            Width           =   1815
         End
         Begin VB.ComboBox cboJob 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3960
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1260
            Width           =   3435
         End
         Begin VB.ComboBox cboDept 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   180
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1260
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker dtpHireDate 
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   2100
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   -2147483639
            CustomFormat    =   "M/d/yyyy"
            Format          =   22740995
            CurrentDate     =   36622
            MaxDate         =   73050
            MinDate         =   32874
         End
         Begin VB.Label lblHireDate 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   240
            TabIndex        =   14
            Top             =   2100
            Width           =   1515
         End
         Begin VB.Label lblEmpNbr 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   180
            TabIndex        =   3
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Na&me:"
            Height          =   255
            Left            =   3960
            TabIndex        =   6
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Sched. W&kly Hours:"
            Height          =   255
            Left            =   5580
            TabIndex        =   17
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Emp #:"
            Height          =   255
            Left            =   180
            TabIndex        =   2
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fir&st Name:"
            Height          =   255
            Left            =   960
            TabIndex        =   4
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Depa&rtment:"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Job:"
            Height          =   255
            Left            =   4020
            TabIndex        =   10
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Hir&e Date:"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Hourl&y Rate:"
            Height          =   255
            Left            =   2760
            TabIndex        =   15
            Top             =   1800
            Width           =   1515
         End
      End
   End
End
Attribute VB_Name = "frmEmpMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************
'************************************************************************
'**                                                                    **
'**               F O R M - L E V E L   V A R I A B L E S              **
'**                                                                    **
'************************************************************************
'************************************************************************

Private mobjEmpRst              As Recordset
Private mblnOKToExit            As Boolean
Private mvntBookMark            As Variant
Private mstrAction              As String

Private intCurrTabIndex         As Integer
Private mblnValidationError     As Boolean
Private mblnActivated           As Boolean
Private mblnChangeMade          As Boolean
    
'************************************************************************
'************************************************************************
'**                                                                    **
'**    E X E C U T A B L E   C O D E   B E G I N S   H E R E . . .     **
'**                                                                    **
'************************************************************************
'************************************************************************

'************************************************************************
'*                                                                      *
'*                    FORM Event Procedures                             *
'*                                                                      *
'************************************************************************

'------------------------------------------------------------------------
Private Sub Form_Activate()
'------------------------------------------------------------------------

    If mblnActivated Then Exit Sub Else mblnActivated = True
    
    CenterForm Me
    OpenEmpDatabase
    
    If gobjEmpDB.TableDefs("DeptMast").RecordCount = 0 Then
        MsgBox "There are no records in the DeptMast table. " _
             & "At least one record must be present in the DeptMast " _
             & "table in order for Employee maintenance to take place. ", _
             vbExclamation, "No DeptMast Records"
        Unload Me
        Exit Sub
    End If
    
    If gobjEmpDB.TableDefs("JobMast").RecordCount = 0 Then
        MsgBox "There are no records in the JobMast table. " _
             & "At least one record must be present in the JobMast " _
             & "table in order for Employee maintenance to take place. ", _
             vbExclamation, "No JobMast Records"
        Unload Me
        Exit Sub
    End If
    
    Set mobjEmpRst = gobjEmpDB.OpenRecordset("EmpMast", dbOpenDynaset)
    
    LoadDeptCombo
    LoadJobCombo
    
    cboField.ListIndex = 0
    cboRelOp.ListIndex = 0
    
    mblnOKToExit = True
    
    cmdFirst_Click

End Sub

'------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------------------

    If KeyCode = vbKeyF1 Then
        cmdHelp_Click
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------------------------------------

    Dim intResponse As Integer

    If Not mblnOKToExit Then
        MsgBox "You must complete or cancel the current action " _
             & "before you can exit", vbInformation, "Cannot Exit"
        Cancel = 1
        Exit Sub
    End If

    mobjEmpRst.Close
    Set mobjEmpRst = Nothing
    
    CloseEmpDatabase

End Sub

'************************************************************************
'*                           EMPLOYEE FIELDS                            *
'*                TextPlus and Comob Box Event Procedures               *
'************************************************************************

'------------------------------------------------------------------------
Private Sub txtEmpFirst_GotFocus()
'------------------------------------------------------------------------
    SelectTextBoxText txtEmpFirst
End Sub

'------------------------------------------------------------------------
Private Sub txtEmpFirst_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------

    If KeyAscii < 32 Then Exit Sub
    
    KeyAscii = ValidKey(ConvertUpper(KeyAscii), gstrUPPER_ALPHA_PLUS)
    
End Sub

'------------------------------------------------------------------------
Private Sub txtEmpFirst_Change()
'------------------------------------------------------------------------
    mblnChangeMade = True
End Sub

'------------------------------------------------------------------------
Private Sub txtEmpLast_GotFocus()
'------------------------------------------------------------------------
    SelectTextBoxText txtEmpLast
    intCurrTabIndex = txtEmpLast.TabIndex
    ValidateAllFields
End Sub

'------------------------------------------------------------------------
Private Sub txtEmpLast_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------

    If KeyAscii < 32 Then Exit Sub
    
    KeyAscii = ValidKey(ConvertUpper(KeyAscii), gstrUPPER_ALPHA_PLUS)
    
End Sub

'------------------------------------------------------------------------
Private Sub txtEmpLast_Change()
'------------------------------------------------------------------------
    mblnChangeMade = True
End Sub

'------------------------------------------------------------------------
Private Sub cboDept_GotFocus()
'------------------------------------------------------------------------
    intCurrTabIndex = cboDept.TabIndex
    ValidateAllFields
End Sub

'------------------------------------------------------------------------
Private Sub cboDept_Click()
'------------------------------------------------------------------------
    mblnChangeMade = True
End Sub

'------------------------------------------------------------------------
Private Sub cboJob_GotFocus()
'------------------------------------------------------------------------
    intCurrTabIndex = cboJob.TabIndex
    ValidateAllFields
End Sub

'------------------------------------------------------------------------
Private Sub cboJob_Click()
'------------------------------------------------------------------------

    Dim objTempRst      As Recordset
    
    Set objTempRst = gobjEmpDB.OpenRecordset _
        ("SELECT MinRate, AvgRate, MaxRate FROM JobMast " _
       & "WHERE JobNbr = " & cboJob.ItemData(cboJob.ListIndex))
    
    'Note: The first record (and only record in this case) is
    'always current when a recordset is open - therefore, it is
    'not necessary to do "objTempRst.MoveFirst"
    
    'Load the Hourly Rate combo box with the min, avg, and max rates
    'for the selected job, and pre-select the avg rate ...
    With cboHrlyRate
        .Clear
        .AddItem Format$(objTempRst!MinRate, "Fixed")
        .AddItem Format$(objTempRst!AvgRate, "Fixed")
        .AddItem Format$(objTempRst!MaxRate, "Fixed")
        .ListIndex = 1
    End With
    
    Set objTempRst = Nothing
    
    mblnChangeMade = True

End Sub

'------------------------------------------------------------------------
Private Sub dtpHireDate_GotFocus()
'------------------------------------------------------------------------
    intCurrTabIndex = dtpHireDate.TabIndex
    ValidateAllFields
End Sub

'------------------------------------------------------------------------
Private Sub dtpHireDate_Change()
'------------------------------------------------------------------------
    mblnChangeMade = True
End Sub

'------------------------------------------------------------------------
Private Sub cboHrlyRate_GotFocus()
'------------------------------------------------------------------------
    intCurrTabIndex = cboHrlyRate.TabIndex
    ValidateAllFields
End Sub

'------------------------------------------------------------------------
Private Sub cboHrlyRate_Change()
'------------------------------------------------------------------------
    mblnChangeMade = True
End Sub

'------------------------------------------------------------------------
Private Sub cboHrlyRate_Click()
'------------------------------------------------------------------------
    mblnChangeMade = True
End Sub

'------------------------------------------------------------------------
Private Sub cboHrlyRate_LostFocus()
'------------------------------------------------------------------------
    cboHrlyRate.Text = Format$(cboHrlyRate.Text, "Fixed")
End Sub

'------------------------------------------------------------------------
Private Sub txtSchedHrs_GotFocus()
'------------------------------------------------------------------------
    SelectTextBoxText txtSchedHrs
    intCurrTabIndex = txtSchedHrs.TabIndex
    ValidateAllFields
End Sub

'------------------------------------------------------------------------
Private Sub txtSchedHrs_Change()
'------------------------------------------------------------------------
    mblnChangeMade = True
End Sub

'------------------------------------------------------------------------
Private Sub txtSchedHrs_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------

    If KeyAscii < 32 Then Exit Sub
    
    KeyAscii = ValidKey(KeyAscii, gstrNUMERIC_DIGITS & ".")
    
    ' if text already has a decimal point, do not allow another ...
    If Chr$(KeyAscii) = "." And InStr(txtSchedHrs.Text, ".") > 0 Then
        KeyAscii = 0
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtSchedHrs_Validate(Cancel As Boolean)
'------------------------------------------------------------------------
    intCurrTabIndex = -1
    ValidateAllFields
    If mblnValidationError Then Cancel = True
End Sub

'------------------------------------------------------------------------
Private Sub txtSchedHrs_LostFocus()
'------------------------------------------------------------------------
    txtSchedHrs.Text = Format$(txtSchedHrs.Text, "Fixed")
End Sub

'------------------------------------------------------------------------
Private Sub txtCriteria_GotFocus()
'------------------------------------------------------------------------
    SelectTextBoxText txtCriteria
End Sub

'************************************************************************
'*                                                                      *
'*                          COMMAND BUTTON                              *
'*                         Event Procedures                             *
'*                                                                      *
'************************************************************************

'------------------------------------------------------------------------
Private Sub cmdFirst_Click()
'------------------------------------------------------------------------

    If mobjEmpRst.RecordCount = 0 Then Exit Sub

    mobjEmpRst.MoveFirst
    DisplayEmpRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdNext_Click()
'------------------------------------------------------------------------
    
    If mobjEmpRst.RecordCount = 0 Then Exit Sub

    With mobjEmpRst
        .MoveNext
        If .EOF Then
            Beep
            .MoveLast
        End If
    End With
    
    DisplayEmpRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdPrev_Click()
'------------------------------------------------------------------------

    If mobjEmpRst.RecordCount = 0 Then Exit Sub

    With mobjEmpRst
        .MovePrevious
        If .BOF Then
            Beep
            .MoveFirst
        End If
    End With
    
    DisplayEmpRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdLast_Click()
'------------------------------------------------------------------------

    If mobjEmpRst.RecordCount = 0 Then Exit Sub

    mobjEmpRst.MoveLast
    DisplayEmpRecord

End Sub


'------------------------------------------------------------------------
Private Sub cmdAdd_Click()
'------------------------------------------------------------------------

    ClearTheForm
    ResetFormControls True, vbWhite
    mblnChangeMade = False
    
    If mobjEmpRst.RecordCount > 0 Then
        mvntBookMark = mobjEmpRst.Bookmark
    End If

    mobjEmpRst.AddNew
    'display the Access(JET)-generated autonumber ...
    lblEmpNbr.Caption = mobjEmpRst!EmpNbr
    
    mstrAction = "ADD"
    txtEmpFirst.SetFocus
    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub cmdUpdate_Click()
'------------------------------------------------------------------------
    
    If mobjEmpRst.RecordCount = 0 Then
        MsgBox "There are no records currently on file to update.", _
               vbInformation, "Update Record"
        Exit Sub
    End If

    ResetFormControls True, vbWhite
    mblnChangeMade = False
    
    mvntBookMark = mobjEmpRst.Bookmark
    mobjEmpRst.Edit
    
    mstrAction = "UPDATE"
    txtEmpFirst.SetFocus
    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub cmdDelete_Click()
'------------------------------------------------------------------------

    If mobjEmpRst.RecordCount = 0 Then
        MsgBox "There are no records currently on file to delete.", _
               vbInformation, "Delete Record"
        Exit Sub
    End If

    If MsgBox("Are you sure you want to delete this record?", _
              vbQuestion + vbYesNo + vbDefaultButton2, _
              "Delete Record") = vbNo Then
        Exit Sub
    End If

    mobjEmpRst.Delete
    
    If mobjEmpRst.RecordCount = 0 Then
        ClearTheForm
    Else
        cmdNext_Click
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdHelp_Click()
'------------------------------------------------------------------------

    gintHelpFileNbr = 2
    frmHelp.Show vbModal
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdExit_Click()
'------------------------------------------------------------------------

    Unload Me
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdSave_Click()
'------------------------------------------------------------------------
    
    intCurrTabIndex = -1
    ValidateAllFields
    If mblnValidationError Then Exit Sub
    
    With mobjEmpRst
        !EmpFirst = txtEmpFirst.Text
        !EmpLast = txtEmpLast.Text
        !DeptNbr = cboDept.ItemData(cboDept.ListIndex)
        !JobNbr = cboJob.ItemData(cboJob.ListIndex)
        !HireDate = dtpHireDate.Value
        !HrlyRate = Val(cboHrlyRate.Text)
        !SchedHrs = Val(txtSchedHrs.Text)
        .Update
        .Bookmark = .LastModified
    End With
    
    ResetFormControls False, vbButtonFace
    mblnOKToExit = True
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdUndo_Click()
'------------------------------------------------------------------------

    If Not mblnChangeMade Then Exit Sub

    If MsgBox("Do you want to abandon your changes to this record?", _
                  vbQuestion + vbYesNo, "Undo") = vbNo Then
        Exit Sub
    End If

    If mstrAction = "ADD" Then
        ClearTheForm
    Else
        DisplayEmpRecord
    End If
    
    mblnChangeMade = False
    
    txtEmpFirst.SetFocus

End Sub

'------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------
        
    If mblnChangeMade Then
        If MsgBox("Do you want to abandon your changes to this record?", _
                      vbQuestion + vbYesNo, "Undo") = vbNo Then
            Exit Sub
        End If
    End If
    
    If mobjEmpRst.RecordCount = 0 Then
        ClearTheForm
    Else
        mobjEmpRst.Bookmark = mvntBookMark
        DisplayEmpRecord
    End If
    
    ResetFormControls False, vbButtonFace
    mblnOKToExit = True

End Sub

'************************************************************************
'*                                                                      *
'*                      "SEARCH" FRAME CONTROLS                         *
'*                         Event Procedures                             *
'*                                                                      *
'************************************************************************

'------------------------------------------------------------------------
Private Sub cboRelOp_Click()
'------------------------------------------------------------------------

    If cboRelOp.Text = "Like" Then
        If cboField.Text = "First Name" Or cboField.Text = "Last Name" Then
            ' it's OK
        Else
            MsgBox "Comparison operator 'Like' may only be used with the " _
                 & "fields 'First Name' or 'Last Name'.", vbInformation, _
                 "Invalid Comparison Operator"
            cboRelOp.SetFocus
        End If
    End If

End Sub

'------------------------------------------------------------------------
Private Sub cmdFind_Click(Index As Integer)
'------------------------------------------------------------------------

    Dim strFindString   As String
    
    ' perform this validation before moving on ...
    If cboField.Text = "Hire Date" Then
        If IsDate(txtCriteria.Text) Then
            txtCriteria.Text _
                = Format$(CDate(txtCriteria.Text), "m/d/yyyy")
        Else
            MsgBox "Criteria for 'Hire Date' is not valid.", _
                   vbExclamation, "Invalid Criteria"
            txtCriteria.SetFocus
            Exit Sub
        End If
    End If
    
    'save current rec's bookmark in case of NoMatch ...
    mvntBookMark = mobjEmpRst.Bookmark
    
    'start building the criteria string for the Find method with the field
    'name of the desired database field, based on the user's cboField selection ...
    Select Case cboField.Text
        Case "Emp #":           strFindString = "EmpNbr"
        Case "First Name":      strFindString = "EmpFirst"
        Case "Last Name":       strFindString = "EmpLast"
        Case "Dept #":          strFindString = "DeptNbr"
        Case "Job #":           strFindString = "JobNbr"
        Case "Hire Date":       strFindString = "HireDate"
        Case "Hourly Rate":     strFindString = "HrlyRate"
        Case "Sched. Wkly Hrs": strFindString = "SchedHrs"
    End Select
    
    'append the selected relational operator to the find string ...
    strFindString = strFindString & " " & cboRelOp.Text & " "
    
    'finally, append the value to search for to the find string ...
    If cboField.Text = "First Name" _
    Or cboField.Text = "Last Name" Then
        strFindString = strFindString _
                      & Chr$(34) & txtCriteria.Text & Chr$(34)
    ElseIf cboField.Text = "Hire Date" Then
        strFindString = strFindString _
                      & "#" & txtCriteria.Text & "#"
    Else
        strFindString = strFindString & Val(txtCriteria.Text)
    End If
    
    ' call the appropriate Find method, depending upon which
    ' button the user clicked ...
    Select Case Index
        Case 0: mobjEmpRst.FindFirst strFindString
        Case 1: mobjEmpRst.FindPrevious strFindString
        Case 2: mobjEmpRst.FindNext strFindString
        Case 3: mobjEmpRst.FindLast strFindString
    End Select
    
    ' deal with the match results ...
    If mobjEmpRst.NoMatch Then
        MsgBox "No (other) records matched your search criteria.", _
               vbInformation, "Not Found"
        mobjEmpRst.Bookmark = mvntBookMark
    Else
        ' the found record is now the current record ...
        DisplayEmpRecord
    End If
    
End Sub


'************************************************************************
'*                                                                      *
'*                        PROGRAMMER-DEFINED                            *
'*                 (Non-Event) Procedures & Functions                   *
'*                                                                      *
'************************************************************************

'------------------------------------------------------------------------
Private Sub LoadDeptCombo()
'------------------------------------------------------------------------

    Dim objTempRst  As Recordset
    
    Set objTempRst = gobjEmpDB.OpenRecordset("DeptMast", dbOpenTable)
    With objTempRst
        .MoveFirst
        Do Until .EOF
            cboDept.AddItem !DeptName & " (" & !DeptNbr & ")"
            cboDept.ItemData(cboDept.NewIndex) = !DeptNbr
            .MoveNext
        Loop
        .Close
    End With

    Set objTempRst = Nothing

End Sub

'------------------------------------------------------------------------
Private Sub LoadJobCombo()
'------------------------------------------------------------------------

    Dim objTempRst  As Recordset
    
    Set objTempRst = gobjEmpDB.OpenRecordset("JobMast", dbOpenTable)
    With objTempRst
        .MoveFirst
        Do Until .EOF
            cboJob.AddItem !JobTitle & " (" & !JobNbr & ")"
            cboJob.ItemData(cboJob.NewIndex) = !JobNbr
            .MoveNext
        Loop
        .Close
    End With

    Set objTempRst = Nothing

End Sub

'------------------------------------------------------------------------
Private Sub DisplayEmpRecord()
'------------------------------------------------------------------------

    Dim intX    As Integer

    With mobjEmpRst
        lblEmpNbr = !EmpNbr
        txtEmpFirst.Text = !EmpFirst
        txtEmpLast.Text = !EmpLast
        For intX = 0 To cboDept.ListCount - 1
            If !DeptNbr = cboDept.ItemData(intX) Then
                cboDept.ListIndex = intX
                Exit For
            End If
        Next
        For intX = 0 To cboJob.ListCount - 1
            If !JobNbr = cboJob.ItemData(intX) Then
                cboJob.ListIndex = intX         ' will invoke cboJob_Click event
                Exit For
            End If
        Next
        lblHireDate = Format$(!HireDate, "m/d/yyyy")
        dtpHireDate.Value = !HireDate
        cboHrlyRate.Text = Format$(!HrlyRate, "#0.00")
        txtSchedHrs.Text = Format$(!SchedHrs, "#0.00")
    End With
   
End Sub

'------------------------------------------------------------------------
Private Sub ResetFormControls(pblnEnabledValue As Boolean, lngColor As Long)
'------------------------------------------------------------------------

    Dim intX As Integer

    fraEmpInfoInner.Enabled = pblnEnabledValue
    
    txtEmpFirst.BackColor = lngColor
    txtEmpLast.BackColor = lngColor
    cboDept.BackColor = lngColor
    cboJob.BackColor = lngColor
    
    If pblnEnabledValue = True Then
        dtpHireDate.Value = CDate(lblHireDate)
    Else
        lblHireDate = Format$(dtpHireDate.Value, "m/d/yyyy")
    End If
    
    dtpHireDate.Visible = pblnEnabledValue
    lblHireDate.Visible = Not pblnEnabledValue
    
    cboHrlyRate.BackColor = lngColor
    txtSchedHrs.BackColor = lngColor

    cmdSave.Enabled = pblnEnabledValue
    cmdUndo.Enabled = pblnEnabledValue
    cmdCancel.Enabled = pblnEnabledValue

    cmdFirst.Enabled = Not pblnEnabledValue
    cmdNext.Enabled = Not pblnEnabledValue
    cmdPrev.Enabled = Not pblnEnabledValue
    cmdLast.Enabled = Not pblnEnabledValue
    cmdAdd.Enabled = Not pblnEnabledValue
    cmdUpdate.Enabled = Not pblnEnabledValue
    cmdDelete.Enabled = Not pblnEnabledValue
    cmdExit.Enabled = Not pblnEnabledValue
    
    fraSearchInner.Enabled = Not pblnEnabledValue
    cboField.BackColor = IIf(lngColor = vbWhite, vbButtonFace, vbWhite)
    cboRelOp.BackColor = IIf(lngColor = vbWhite, vbButtonFace, vbWhite)
    txtCriteria.BackColor = IIf(lngColor = vbWhite, vbButtonFace, vbWhite)
    
    For intX = 0 To 3
        cmdFind(intX).Enabled = Not pblnEnabledValue
    Next

    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub ClearTheForm()
'------------------------------------------------------------------------
    
    txtEmpFirst.Text = ""
    txtEmpLast.Text = ""
    cboDept.ListIndex = 0                   'default to first Dept in the list
    cboJob.ListIndex = 0                    'default to first Job in the list
    lblHireDate = Format$(Date, "m/d/yyyy") 'default to today's date
    cboHrlyRate.ListIndex = 1               'default to the average rate
    txtSchedHrs.Text = "40.00"              'default to 40 hrs per week
    
End Sub

'------------------------------------------------------------------------
Private Sub ValidateAllFields()
'------------------------------------------------------------------------

    mblnValidationError = False
    
    '*** First Name

    If txtEmpFirst.Text = "" Then
        MsgBox "First Name must not be blank", _
               vbExclamation, "First Name"
        mblnValidationError = True
        Beep
        txtEmpFirst.SetFocus
    End If
    
    If intCurrTabIndex = txtEmpLast.TabIndex Or mblnValidationError Then
        Exit Sub
    End If
    
    '*** Last Name
    If txtEmpLast.Text = "" Then
        MsgBox "Last Name must not be blank", _
               vbExclamation, "Last Name"
        mblnValidationError = True
        Beep
        txtEmpLast.SetFocus
    End If
    
    If intCurrTabIndex = cboDept.TabIndex Or mblnValidationError Then
        Exit Sub
    End If
    
    '*** Department
    '   (no validation logic needed)
    
    If intCurrTabIndex = cboJob.TabIndex Or mblnValidationError Then
        Exit Sub
    End If
       
    '*** Job
    '   (no validation logic needed)
    
    If intCurrTabIndex = dtpHireDate.TabIndex Or mblnValidationError Then
        Exit Sub
    End If

    '*** Hire Date
    '   (no validation logic needed)
    
    If intCurrTabIndex = cboHrlyRate.TabIndex Or mblnValidationError Then
        Exit Sub
    End If

    '*** Hourly Rate
    
    If cboHrlyRate.Text = "" Then
        MsgBox "Hourly Rate must be entered.", _
               vbExclamation, "Hourly Rate"
        mblnValidationError = True
        Beep
        cboHrlyRate.SetFocus
    ElseIf Not IsNumeric(cboHrlyRate.Text) Then
        MsgBox "Hourly Rate must be numeric.", _
               vbExclamation, "Hourly Rate"
        mblnValidationError = True
        Beep
        cboHrlyRate.SetFocus
    ElseIf Val(cboHrlyRate.Text) <= 0 Then
        MsgBox "Hourly Rate must be greater than zero.", _
               vbExclamation, "Hourly Rate"
        mblnValidationError = True
        Beep
        cboHrlyRate.SetFocus
    End If
    
    If intCurrTabIndex = txtSchedHrs.TabIndex Or mblnValidationError Then
        Exit Sub
    End If

    '*** Scheduled Hours

    If txtSchedHrs.Text = "" Then
        MsgBox "Hours must be entered.", _
               vbExclamation, "Hours"
        mblnValidationError = True
        Beep
        txtSchedHrs.SetFocus
    ElseIf Val(txtSchedHrs.Text) <= 0 Then
        MsgBox "Hours must be greater than zero.", _
               vbExclamation, "Hours"
        mblnValidationError = True
        Beep
        txtSchedHrs.SetFocus
    End If

End Sub
