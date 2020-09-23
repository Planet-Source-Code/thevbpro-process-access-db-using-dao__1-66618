VERSION 5.00
Begin VB.Form frmDeptMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Database Maintenance - Department Maintenance"
   ClientHeight    =   4230
   ClientLeft      =   2460
   ClientTop       =   2790
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7830
   Begin VB.CommandButton cmdGoToDeptName 
      Caption         =   "&Go to Dept Name ..."
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   3540
      Width           =   1635
   End
   Begin VB.CommandButton cmdGoToDeptNbr 
      Caption         =   "&Go to Dept # ..."
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   3060
      Width           =   1635
   End
   Begin VB.Frame fraDeptData 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   6075
      Begin VB.TextBox txtDeptField 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   2
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1740
         Width           =   4155
      End
      Begin VB.TextBox txtDeptField 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1020
         Width           =   4155
      End
      Begin VB.TextBox txtDeptField 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   1260
         MaxLength       =   4
         TabIndex        =   2
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Lo&cation:"
         Height          =   315
         Left            =   300
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dept Na&me:"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dept #:"
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Record"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   3060
      Width           =   1395
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Record"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3060
      Width           =   1395
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Record"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   3060
      Width           =   1395
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First Record"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6420
      TabIndex        =   7
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last Record"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6420
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next Record"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6420
      TabIndex        =   9
      Top             =   1695
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Prev Record"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6420
      TabIndex        =   8
      Top             =   1005
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6420
      TabIndex        =   19
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Return"
      Height          =   375
      Left            =   6420
      TabIndex        =   20
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo "
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   3540
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3540
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   3540
      Width           =   1395
   End
End
Attribute VB_Name = "frmDeptMaint"
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

Private mobjDeptRst                     As Recordset
Private mvntBookMark                    As Variant
Private mstrAction                      As String
Private mblnOKToExit                    As Boolean
Private mblnValidationError             As Boolean
Private mblnChangeMade                  As Boolean
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
    
    Set mobjDeptRst = gobjEmpDB.OpenRecordset("DeptMast", dbOpenTable)
    mobjDeptRst.Index = "idxDeptNbrPK"
    
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

    mobjDeptRst.Close
    Set mobjDeptRst = Nothing

    CloseEmpDatabase

End Sub


'************************************************************************
'*                           DEPT FIELDS                                *
'************************************************************************

'------------------------------------------------------------------------
Private Sub txtDeptField_GotFocus(Index As Integer)
'------------------------------------------------------------------------

    SelectTextBoxText txtDeptField(Index)
    
    If Index > 0 Then
        mintCurrTabIndex = txtDeptField(Index).TabIndex
        ValidateAllFields
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtDeptField_KeyPress(Index As Integer, KeyAscii As Integer)
'------------------------------------------------------------------------

    If KeyAscii < 32 Then Exit Sub
    
    If Index = 0 Then
        ' dept number - allow only digits
        KeyAscii = ValidKey(KeyAscii, gstrNUMERIC_DIGITS)
    Else
        ' dept name or location - force uppercase
        KeyAscii = ConvertUpper(KeyAscii)
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtDeptField_Change(Index As Integer)
'------------------------------------------------------------------------

    mblnChangeMade = True
    
    If Index < 2 Then
        TabToNextTextBox txtDeptField(Index), txtDeptField(Index + 1)
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub txtDeptField_Validate(Index As Integer, Cancel As Boolean)
'------------------------------------------------------------------------

    ' this event is only being used for the last field on the form ...
    If Index = 2 Then
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

    If mobjDeptRst.RecordCount = 0 Then Exit Sub

    mobjDeptRst.MoveFirst
    DisplayDeptRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdNext_Click()
'------------------------------------------------------------------------
    
    If mobjDeptRst.RecordCount = 0 Then Exit Sub

    With mobjDeptRst
        .MoveNext
        If .EOF Then
            Beep
            .MoveLast
        End If
    End With
    
    DisplayDeptRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdPrev_Click()
'------------------------------------------------------------------------

    If mobjDeptRst.RecordCount = 0 Then Exit Sub

    With mobjDeptRst
        .MovePrevious
        If .BOF Then
            Beep
            .MoveFirst
        End If
    End With
    
    DisplayDeptRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdLast_Click()
'------------------------------------------------------------------------

    If mobjDeptRst.RecordCount = 0 Then Exit Sub

    mobjDeptRst.MoveLast
    DisplayDeptRecord

End Sub

'------------------------------------------------------------------------
Private Sub cmdAdd_Click()
'------------------------------------------------------------------------

    ClearTheForm
    
    mstrAction = "ADD"
    ResetFormControls True, vbWhite
    mblnChangeMade = False
    
    If mobjDeptRst.RecordCount > 0 Then
        mvntBookMark = mobjDeptRst.Bookmark
    End If

    mobjDeptRst.AddNew
    
    txtDeptField(0).SetFocus
    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub cmdUpdate_Click()
'------------------------------------------------------------------------
    
    If mobjDeptRst.RecordCount = 0 Then
        MsgBox "There are no records currently on file to update.", _
               vbInformation, "Update Record"
        Exit Sub
    End If
    
    mstrAction = "UPDATE"
    ResetFormControls True, vbWhite
    mblnChangeMade = False
    
    mvntBookMark = mobjDeptRst.Bookmark
    mobjDeptRst.Edit
    
    txtDeptField(1).SetFocus
    mblnOKToExit = False

End Sub

'------------------------------------------------------------------------
Private Sub cmdDelete_Click()
'------------------------------------------------------------------------

    Dim objTempRst  As Recordset
    Dim intEmpCount As Integer
    
    If mobjDeptRst.RecordCount = 0 Then
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
                  & "WHERE DeptNbr = " & mobjDeptRst!DeptNbr)
    intEmpCount = objTempRst!EmpCount
    objTempRst.Close
    Set objTempRst = Nothing
                  
    If intEmpCount > 0 Then
        MsgBox "This department record cannot be deleted because " _
             & "it is in use by one or more employees.", _
               vbExclamation, _
               "Department Is In Use"
        Exit Sub
    End If

    mobjDeptRst.Delete
    
    If mobjDeptRst.RecordCount = 0 Then
        ClearTheForm
    Else
        cmdNext_Click
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdSave_Click()
'------------------------------------------------------------------------
    
    mintCurrTabIndex = -1
    ValidateAllFields
    If mblnValidationError Then Exit Sub
    
    With mobjDeptRst
        If mstrAction = "ADD" Then
            !DeptNbr = txtDeptField(0).Text
        End If
        !DeptName = txtDeptField(1).Text
        !Location = txtDeptField(2).Text
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
        txtDeptField(0).SetFocus
    Else
        DisplayDeptRecord
        txtDeptField(1).SetFocus
    End If

    mblnChangeMade = False

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
    
    If mobjDeptRst.RecordCount = 0 Then
        ClearTheForm
    Else
        mobjDeptRst.Bookmark = mvntBookMark
        DisplayDeptRecord
    End If
    
    ResetFormControls False, vbButtonFace
    mblnOKToExit = True

End Sub

'------------------------------------------------------------------------
Private Sub cmdGoToDeptNbr_Click()
'------------------------------------------------------------------------

    Dim strReqDeptNbr       As String
    Dim lngReqDeptNbr       As Long
    
    If mobjDeptRst.Index = "idxDeptName" Then
        If MsgBox("This search will cause the record browsing " _
                & "sequence to change to department number sequence. " _
                & "Is that OK?", vbYesNo + vbQuestion, _
                "Browse Sequence") = vbNo Then
            Exit Sub
        End If
    End If
    
    strReqDeptNbr = InputBox _
        ("Type in the Department # that you are looking for. ", _
         "Go To Dept # ...")
    
    If strReqDeptNbr = "" Then
        ' user clicked the Cancel button on the input box
        ' or did not enter anything
        Exit Sub
    End If
    
    lngReqDeptNbr = Val(strReqDeptNbr)
    
    mvntBookMark = mobjDeptRst.Bookmark
    
    mobjDeptRst.Index = "idxDeptNbrPK"
    mobjDeptRst.Seek "=", lngReqDeptNbr
    
    If mobjDeptRst.NoMatch Then
        MsgBox "Dept # " & lngReqDeptNbr & " could not be found.", _
               vbExclamation, "Dept # Not Found"
        mobjDeptRst.Bookmark = mvntBookMark
    Else
        DisplayDeptRecord
    End If

End Sub

'------------------------------------------------------------------------
Private Sub cmdGoToDeptName_Click()
'------------------------------------------------------------------------

    Dim strReqDeptName       As String
    
    If mobjDeptRst.Index = "idxDeptNbrPK" Then
        If MsgBox("This search will cause the record browsing " _
                & "sequence to change to department name sequence. " _
                & "Is that OK?", vbYesNo + vbQuestion, _
                "Browse Sequence") = vbNo Then
            Exit Sub
        End If
    End If
    
    strReqDeptName = UCase$(InputBox _
        ("Type in the first several letters of the Department Name that you are looking for. ", _
         "Go To Dept # ..."))
    
    If strReqDeptName = "" Then
        ' user clicked the Cancel button on the input box
        ' or did not enter anything
        Exit Sub
    End If
    
    mvntBookMark = mobjDeptRst.Bookmark
    
    mobjDeptRst.Index = "idxDeptName"
    mobjDeptRst.Seek ">=", strReqDeptName
    
    If mobjDeptRst.NoMatch Then
        MsgBox "Dept Name beginning '" & strReqDeptName & "' could not be found.", _
               vbExclamation, "Dept Not Found"
        mobjDeptRst.Bookmark = mvntBookMark
    Else
        DisplayDeptRecord
    End If

End Sub

'------------------------------------------------------------------------
Private Sub cmdHelp_Click()
'------------------------------------------------------------------------

    gintHelpFileNbr = 3
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
Private Sub DisplayDeptRecord()
'------------------------------------------------------------------------

    Dim intX    As Integer

    With mobjDeptRst
        txtDeptField(0).Text = !DeptNbr
        txtDeptField(1).Text = !DeptName
        txtDeptField(2).Text = !Location
    End With
    
End Sub

'------------------------------------------------------------------------
Private Sub ResetFormControls(blnEnabledValue As Boolean, lngColor As Long)
'------------------------------------------------------------------------

    Dim intX As Integer

    fraDeptData.Enabled = blnEnabledValue
    
    For intX = 0 To 2
        txtDeptField(intX).BackColor = lngColor
    Next
    
    If mstrAction = "UPDATE" Then
        txtDeptField(0).Enabled = Not blnEnabledValue
    End If
    
    cmdSave.Enabled = blnEnabledValue
    cmdUndo.Enabled = blnEnabledValue
    cmdCancel.Enabled = blnEnabledValue

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
    
    For intX = 0 To 2
        txtDeptField(intX).Text = ""
    Next
    
End Sub

'------------------------------------------------------------------------
Private Sub ValidateAllFields()
'------------------------------------------------------------------------

    Dim intX    As Integer
        
    mblnValidationError = False
    
    For intX = 0 To 2
        If Not DeptFieldIsValid(intX) Then
            mblnValidationError = True
            Beep
            txtDeptField(intX).SetFocus
        End If
        If intX < 2 Then
            If mintCurrTabIndex = txtDeptField(intX + 1).TabIndex _
            Or mblnValidationError Then
                Exit For
            End If
        End If
    Next
    
End Sub
    
'------------------------------------------------------------------------
Private Function DeptFieldIsValid(intFieldIndex As Integer) As Boolean
'------------------------------------------------------------------------

    Dim strMBMsg        As String
    Dim strMBTitle      As String
    Dim blnItsValid     As Boolean

    blnItsValid = True
    
    Select Case intFieldIndex
        Case 0
            '*** Department Number
            If mstrAction = "ADD" Then
                ' validation checks for the department number are only
                ' applicable when adding, not updating a record ...
                If txtDeptField(0).Text = "" Then
                    strMBMsg = "Department Number must be entered"
                    strMBTitle = "Department Number"
                    blnItsValid = False
                ElseIf DeptExists(txtDeptField(0).Text) Then
                    strMBMsg = "Department '" & txtDeptField(0).Text _
                         & "' already exists."
                    strMBTitle = "Department Already Exists"
                    blnItsValid = False
                End If
            End If
        Case 1
            '*** Department Name
            If txtDeptField(1).Text = "" Then
                strMBMsg = "Department Name must not be blank"
                strMBTitle = "Department Name"
                blnItsValid = False
            End If
        Case Else
            '*** Location
            If txtDeptField(2).Text = "" Then
                strMBMsg = "Location must be entered"
                strMBTitle = "Location"
                blnItsValid = False
            End If
    End Select
    
    If blnItsValid Then
        DeptFieldIsValid = True
    Else
        DeptFieldIsValid = False
        MsgBox strMBMsg, vbExclamation, strMBTitle
    End If
    
End Function

'------------------------------------------------------------------------
Private Function DeptExists(strDeptNbr As String) As Boolean
'------------------------------------------------------------------------

    Dim objTempRst   As Recordset
    Dim intDeptCount As Integer
    
    Set objTempRst = gobjEmpDB.OpenRecordset _
                 ("SELECT COUNT(*) AS DeptCount FROM DeptMast " _
                  & "WHERE DeptNbr = " & strDeptNbr)
    intDeptCount = objTempRst!DeptCount
    objTempRst.Close
    Set objTempRst = Nothing
                  
    DeptExists = IIf(intDeptCount = 0, False, True)
    
End Function
