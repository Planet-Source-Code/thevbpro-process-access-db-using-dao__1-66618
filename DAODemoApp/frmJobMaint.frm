VERSION 5.00
Begin VB.Form frmJobMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Database Maintenance - Job Maintenance"
   ClientHeight    =   4515
   ClientLeft      =   2595
   ClientTop       =   3105
   ClientWidth     =   7860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7860
   Begin VB.CommandButton cmdGoToJobTitle 
      Caption         =   "&Go to Job Title ..."
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Top             =   3720
      Width           =   1635
   End
   Begin VB.CommandButton cmdGoToJobNbr 
      Caption         =   "&Go to Job # ..."
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   3240
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo "
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Return"
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Prev Record"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   1120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next Record"
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   1880
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last Record"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First Record"
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Record"
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   3240
      Width           =   1395
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Record"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   1395
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Record"
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Frame fraJobData 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.TextBox txtJobField 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   3
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtJobField 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   2
         Left            =   2790
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtJobField 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtJobField 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   1260
         MaxLength       =   35
         TabIndex        =   3
         Top             =   1020
         Width           =   4275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ma&ximum:"
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "A&verage:"
         Height          =   255
         Left            =   2790
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Minimum:"
         Height          =   255
         Left            =   1260
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Job #:"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   420
         Width           =   675
      End
      Begin VB.Label lblJobNbr 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Job Title:"
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rates:"
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmJobMaint"
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

Private mobjJobRst                      As Recordset
Private mvntBookMark                    As Variant
Private mstrAction                      As String

Private mblnOKToExit                    As Boolean
Private mblnChangeMade                  As Boolean
Private mblnValidationError             As Boolean

Private mintCurrTabIndex                As Integer

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
Private Sub Form_Load()
'------------------------------------------------------------------------

    CenterForm Me
    
    OpenEmpDatabase
    
    Set mobjJobRst = gobjEmpDB.OpenRecordset("JobMast", dbOpenTable)
        
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

    CloseEmpDatabase

End Sub

'************************************************************************
'*                           JOB FIELDS                                 *
'************************************************************************

'------------------------------------------------------------------------
Private Sub txtJobField_GotFocus(Index As Integer)
'------------------------------------------------------------------------

    SelectTextBoxText txtJobField(Index)

    If Index > 0 Then
        mintCurrTabIndex = txtJobField(Index).TabIndex
        ValidateAllFields
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtJobField_KeyPress(Index As Integer, KeyAscii As Integer)
'------------------------------------------------------------------------

    If KeyAscii < 32 Then Exit Sub
    
    If Index > 0 Then
        ' rate field - allow only digits and decimal point
        KeyAscii = ValidKey(KeyAscii, gstrNUMERIC_DIGITS & ".")
        ' if text already has a decimal point, do not allow another ...
        If Chr$(KeyAscii) = "." And InStr(txtJobField(Index).Text, ".") > 0 Then
            KeyAscii = 0
        End If
    Else
        ' job description - force uppercase
        KeyAscii = ConvertUpper(KeyAscii)
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtJobField_Change(Index As Integer)
'------------------------------------------------------------------------

    mblnChangeMade = True
    
    If Index < 3 Then
        TabToNextTextBox txtJobField(Index), txtJobField(Index + 1)
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtJobField_LostFocus(Index As Integer)
'------------------------------------------------------------------------

    If Index > 0 Then
        txtJobField(Index).Text = Format$(txtJobField(Index).Text, "Fixed")
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtJobField_Validate(Index As Integer, Cancel As Boolean)
'------------------------------------------------------------------------

    ' this event is only being used for the last field on the form ...
    If Index = 3 Then
        mintCurrTabIndex = -1
        ValidateAllFields
        If mblnValidationError Then
            Cancel = True
        End If
    End If

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

    If mobjJobRst.RecordCount = 0 Then Exit Sub

    mobjJobRst.MoveFirst
    DisplayJobRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdNext_Click()
'------------------------------------------------------------------------
    
    If mobjJobRst.RecordCount = 0 Then Exit Sub

    With mobjJobRst
        .MoveNext
        If .EOF Then
            Beep
            .MoveLast
        End If
    End With
    
    DisplayJobRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdPrev_Click()
'------------------------------------------------------------------------

    If mobjJobRst.RecordCount = 0 Then Exit Sub

    With mobjJobRst
        .MovePrevious
        If .BOF Then
            Beep
            .MoveFirst
        End If
    End With
    
    DisplayJobRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdLast_Click()
'------------------------------------------------------------------------

    If mobjJobRst.RecordCount = 0 Then Exit Sub

    mobjJobRst.MoveLast
    DisplayJobRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdAdd_Click()
'------------------------------------------------------------------------

    ClearTheForm
    ResetFormControls True, vbWhite
    mblnChangeMade = False
    
    If mobjJobRst.RecordCount > 0 Then
        mvntBookMark = mobjJobRst.Bookmark
    End If

    mobjJobRst.AddNew
    'display the Access(JET)-generated autonumber ...
    lblJobNbr.Caption = mobjJobRst!JobNbr
    
    mstrAction = "ADD"
    txtJobField(0).SetFocus
    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub cmdUpdate_Click()
'------------------------------------------------------------------------
    
    If mobjJobRst.RecordCount = 0 Then
        MsgBox "There are no records currently on file to update.", _
               vbInformation, "Update Record"
        Exit Sub
    End If

    ResetFormControls True, vbWhite
    mblnChangeMade = False
    
    mvntBookMark = mobjJobRst.Bookmark
    mobjJobRst.Edit
    
    mstrAction = "UPDATE"
    txtJobField(0).SetFocus
    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub cmdDelete_Click()
'------------------------------------------------------------------------

    Dim objTempRst  As Recordset
    Dim intEmpCount As Integer
    
    If mobjJobRst.RecordCount = 0 Then
        MsgBox "There are no records currently on file to delete.", _
               vbInformation, "Delete Record"
        Exit Sub
    End If

    If MsgBox("Are you sure you want to delete this record?", _
              vbQuestion + vbYesNo + vbDefaultButton2, _
              "Delete Record") = vbNo Then
        Exit Sub
    End If

    ' check for referential integrity violation ...
    Set objTempRst = gobjEmpDB.OpenRecordset _
                     ("SELECT COUNT(*) AS EmpCount FROM EmpMast " _
                      & "WHERE JobNbr = " & mobjJobRst!JobNbr)
    intEmpCount = objTempRst!EmpCount
    objTempRst.Close
    Set objTempRst = Nothing
                  
    If intEmpCount > 0 Then
        MsgBox "This job record cannot be deleted because " _
             & "it is in use by one or more employees.", _
             vbExclamation, "Job Is In Use"
        Exit Sub
    End If

    mobjJobRst.Delete
    
    If mobjJobRst.RecordCount = 0 Then
        ClearTheForm
    Else
        cmdNext_Click
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdGoToJobNbr_Click()
'------------------------------------------------------------------------

    Dim strReqJobNbr       As String
    Dim lngReqJobNbr       As Long
    
    If mobjJobRst.Index = "idxJobName" Then
        If MsgBox("This search will cause the record browsing " _
                & "sequence to change to job number sequence. " _
                & "Is that OK?", vbYesNo + vbQuestion, _
                "Browse Sequence") = vbNo Then
            Exit Sub
        End If
    End If
    
    strReqJobNbr = InputBox _
        ("Type in the Job # that you are looking for. ", _
         "Go To Job # ...")
    
    If strReqJobNbr = "" Then
        ' user clicked the Cancel button on the input box
        ' or did not enter anything
        Exit Sub
    End If
    
    lngReqJobNbr = Val(strReqJobNbr)
    
    mvntBookMark = mobjJobRst.Bookmark
    
    mobjJobRst.Index = "idxJobNbrPK"
    mobjJobRst.Seek "=", lngReqJobNbr
    
    If mobjJobRst.NoMatch Then
        MsgBox "Job # " & lngReqJobNbr & " could not be found.", _
               vbExclamation, "Job # Not Found"
        mobjJobRst.Bookmark = mvntBookMark
    Else
        DisplayJobRecord
    End If

End Sub

'------------------------------------------------------------------------
Private Sub cmdGoToJobTitle_Click()
'------------------------------------------------------------------------

    Dim strReqJobTitle       As String
    
    If mobjJobRst.Index = "idxJobNbrPK" Then
        If MsgBox("This search will cause the record browsing " _
                & "sequence to change to job title sequence. " _
                & "Is that OK?", vbYesNo + vbQuestion, _
                "Browse Sequence") = vbNo Then
            Exit Sub
        End If
    End If
    
    strReqJobTitle = UCase$(InputBox _
        ("Type in the first several letters of the Job title that you are looking for. ", _
         "Go To Job # ..."))
    
    If strReqJobTitle = "" Then
        ' user clicked the Cancel button on the input box
        ' or did not enter anything
        Exit Sub
    End If
    
    mvntBookMark = mobjJobRst.Bookmark
    
    mobjJobRst.Index = "idxJobtitle"
    mobjJobRst.Seek ">=", strReqJobTitle
    
    If mobjJobRst.NoMatch Then
        MsgBox "Job Title beginning '" & strReqJobTitle & "' could not be found.", _
               vbExclamation, "Job Not Found"
        mobjJobRst.Bookmark = mvntBookMark
    Else
        DisplayJobRecord
    End If

End Sub

'------------------------------------------------------------------------
Private Sub cmdSave_Click()
'------------------------------------------------------------------------
    
    mintCurrTabIndex = -1
    ValidateAllFields
    If mblnValidationError Then Exit Sub
    
    With mobjJobRst
        !JobTitle = txtJobField(0).Text
        !MinRate = Val(txtJobField(1).Text)
        !AvgRate = Val(txtJobField(2).Text)
        !MaxRate = Val(txtJobField(3).Text)
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
        lblJobNbr.Caption = mobjJobRst!JobNbr
    Else
        DisplayJobRecord
    End If

    mblnChangeMade = False
    
    txtJobField(0).SetFocus

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
        
    If mobjJobRst.RecordCount = 0 Then
        ClearTheForm
    Else
        mobjJobRst.Bookmark = mvntBookMark
        DisplayJobRecord
    End If
    
    ResetFormControls False, vbButtonFace
    mblnOKToExit = True

End Sub

'------------------------------------------------------------------------
Private Sub cmdHelp_Click()
'------------------------------------------------------------------------

    gintHelpFileNbr = 4
    frmHelp.Show vbModal
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdExit_Click()
'------------------------------------------------------------------------

    Unload Me
    
End Sub


'************************************************************************
'*                                                                      *
'*                        PROGRAMMER-DEFINED                            *
'*                 (Non-Event) Procedures & Functions                   *
'*                                                                      *
'************************************************************************

'------------------------------------------------------------------------
Private Sub DisplayJobRecord()
'------------------------------------------------------------------------

    With mobjJobRst
        lblJobNbr.Caption = !JobNbr
        txtJobField(0).Text = !JobTitle
        txtJobField(1).Text = Format$(!MinRate, "Fixed")
        txtJobField(2).Text = Format$(!AvgRate, "Fixed")
        txtJobField(3).Text = Format$(!MaxRate, "Fixed")
    End With
    
End Sub

'------------------------------------------------------------------------
Private Sub ResetFormControls(blnEnabledValue As Boolean, lngColor As Long)
'------------------------------------------------------------------------

    Dim intX As Integer

    fraJobData.Enabled = blnEnabledValue
    
    For intX = 0 To 3
        txtJobField(intX).BackColor = lngColor
    Next
    
    cmdSave.Enabled = blnEnabledValue
    cmdUndo.Enabled = blnEnabledValue
    cmdCancel.Enabled = blnEnabledValue

    cmdGoToJobNbr.Enabled = Not blnEnabledValue
    cmdGoToJobTitle.Enabled = Not blnEnabledValue
    cmdFirst.Enabled = Not blnEnabledValue
    cmdNext.Enabled = Not blnEnabledValue
    cmdPrev.Enabled = Not blnEnabledValue
    cmdLast.Enabled = Not blnEnabledValue
    cmdAdd.Enabled = Not blnEnabledValue
    cmdUpdate.Enabled = Not blnEnabledValue
    cmdDelete.Enabled = Not blnEnabledValue
    cmdExit.Enabled = Not blnEnabledValue
    
    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub ClearTheForm()
'------------------------------------------------------------------------
    
    Dim intX    As Integer
    
    lblJobNbr = ""
    For intX = 0 To 3
        txtJobField(intX).Text = ""
    Next
    
End Sub

'------------------------------------------------------------------------
Private Sub ValidateAllFields()
'------------------------------------------------------------------------

    Dim intX    As Integer
        
    mblnValidationError = False
    
    For intX = 0 To 3
        If Not JobFieldIsValid(intX) Then
            mblnValidationError = True
            Beep
            txtJobField(intX).SetFocus
        End If
        If intX < 3 Then
            If mintCurrTabIndex = txtJobField(intX + 1).TabIndex _
            Or mblnValidationError Then
                Exit For
            End If
        End If
    Next
    
End Sub
    
'------------------------------------------------------------------------
Private Function JobFieldIsValid(intFieldIndex As Integer) As Boolean
'------------------------------------------------------------------------

    Dim strMBMsg        As String
    Dim strMBTitle      As String
    Dim blnItsValid     As Boolean

    blnItsValid = True
    
    Select Case intFieldIndex
        Case 0
            '*** Job Title
            If txtJobField(0).Text = "" Then
                strMBMsg = "Job Title must not be blank"
                strMBTitle = "Job Title"
                blnItsValid = False
            End If
        Case 1
            '*** Minimum Rate
            If Val(txtJobField(1).Text) <= 0 Then
                strMBMsg = "Minimum Rate must be greater than zero."
                strMBTitle = "Minimum Rate"
                blnItsValid = False
            End If
        Case 2
            '*** Average Rate
            If Val(txtJobField(2).Text) <= 0 Then
                strMBMsg = "Average Rate must be greater than zero."
                strMBTitle = "Average Rate"
                blnItsValid = False
            ElseIf Val(txtJobField(2).Text) < Val(txtJobField(1).Text) Then
                strMBMsg _
                    = "Average Rate must be greater than or equal to the Minimum Rate."
                strMBTitle = "Average Rate"
                blnItsValid = False
            End If
        Case 3
            '*** Maximum Rate
            If Val(txtJobField(3).Text) <= 0 Then
                strMBMsg = "Maximum Rate must be greater than zero."
                strMBTitle = "Maximum Rate"
                blnItsValid = False
            ElseIf Val(txtJobField(3).Text) < Val(txtJobField(2).Text) Then
                strMBMsg _
                    = "Maximum Rate must be greater than or equal to the Average Rate."
                strMBTitle = "Maxiumum Rate"
                blnItsValid = False
            End If
    End Select
    
    If blnItsValid Then
        JobFieldIsValid = True
    Else
        JobFieldIsValid = False
        MsgBox strMBMsg, vbExclamation, strMBTitle
    End If
    
End Function
