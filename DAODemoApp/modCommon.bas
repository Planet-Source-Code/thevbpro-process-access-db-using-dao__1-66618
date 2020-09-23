Attribute VB_Name = "modCommon"
Option Explicit

Public gobjEmpDB                    As Database
Public gintHelpFileNbr              As Integer
Public Const gstrNUMERIC_DIGITS     As String = "0123456789"
Public Const gstrUPPER_ALPHA_PLUS   As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ,'-"

'------------------------------------------------------------------------
Public Sub OpenEmpDatabase()
'------------------------------------------------------------------------

    Set gobjEmpDB = OpenDatabase(GetAppPath() & "EMPLOYEE.MDB")

End Sub

'------------------------------------------------------------------------
Public Sub CloseEmpDatabase()
'------------------------------------------------------------------------

    gobjEmpDB.Close
    Set gobjEmpDB = Nothing

End Sub

'------------------------------------------------------------------------
Public Sub CenterForm(pobjForm As Form)
'------------------------------------------------------------------------

    With pobjForm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With

End Sub

'------------------------------------------------------------------------
Public Function GetAppPath() As String
'------------------------------------------------------------------------

    GetAppPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")

End Function

'------------------------------------------------------------------------
Public Function ValidKey(pintKeyValue As Integer, _
                         pstrSearchString As String) As Integer
'------------------------------------------------------------------------

'  Common function to filter out keyboard characters passed to this
'  function from KeyPress events.
'
'  Typical call:
'      KeyAscii = ValidKey(KeyAscii, gstrNUMERIC_DIGITS)
'

    If pintKeyValue < 32 _
    Or InStr(pstrSearchString, Chr$(pintKeyValue)) > 0 Then
        'Do nothing - i.e., accept the control character or any key
        '             in the search string passed to this function ...
    Else
        'cancel (do not accept) any other key ...
        pintKeyValue = 0
    End If

    ValidKey = pintKeyValue

End Function

'------------------------------------------------------------------------
Public Function ConvertUpper(pintKeyValue As Integer) As Integer
'------------------------------------------------------------------------

'  Common function to force alphabetic keyboard characters to uppercase
'  when called from the KeyPress event.

'  Typical call:
'      KeyAscii = ConvertUpper(KeyAscii)
'

    If Chr$(pintKeyValue) >= "a" And Chr$(pintKeyValue) <= "z" Then
        pintKeyValue = pintKeyValue - 32
    End If

    ConvertUpper = pintKeyValue

End Function

'-----------------------------------------------------------------------------
Public Sub SelectTextBoxText(pobjTextbox As TextBox)
'-----------------------------------------------------------------------------

    With pobjTextbox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'-----------------------------------------------------------------------------
Public Sub TabToNextTextBox(pobjTextBox1 As TextBox, pobjTextBox2 As TextBox)
'-----------------------------------------------------------------------------

    If pobjTextBox2.Enabled = False Then Exit Sub
    
    If Len(pobjTextBox1.Text) = pobjTextBox1.MaxLength Then
        pobjTextBox2.SetFocus
    End If

End Sub


