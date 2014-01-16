Attribute VB_Name = "CommonHelper"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Public Function openFile(FilePath As String, ownerHwnd As Long) As Boolean
     Dim dummy As Long
     
               'open the file using the default Editor or viewer.
     dummy = ShellExecute(ownerHwnd, "Open", FilePath & Chr$(0), Chr$(0), _
                                          Left$(FilePath, InStr(FilePath, "\")), 0)
     openFile = dummy
     
End Function

Public Function extractStringValue(value As Object) As String
  If (Not IsNull(value)) Then
    extractStringValue = value
  Else
    extractStringValue = ""
  End If
End Function
Public Function extractDateValue(value As Object) As String
  If (Not IsNull(value)) Then
    extractDateValue = Format(value, Constants.DEFAULT_FORMAT)
  Else
    extractDateValue = ""
  End If
End Function
Public Function hasValidValue(value As String) As Boolean
   Dim isValid As Boolean
   isValid = True
   If (Not IsNull(value)) Then
   
     If (IsNumeric(value)) Then
       isValid = Val(value) > 0
     Else
       isValid = Trim(value) <> vbNullString
     End If
   End If
   hasValidValue = isValid
End Function
Public Sub sendWarning(txtBox As TextBox, errMsg As String)
  MsgBox errMsg, vbCritical
  txtBox.BackColor = vbRed
  txtBox.ForeColor = vbWhite
  txtBox.SetFocus
End Sub
Public Sub sendComboBoxWarning(cmbBox As ComboBox, errMsg As String)
  MsgBox errMsg, vbCritical
  cmbBox.BackColor = vbRed
  cmbBox.ForeColor = vbWhite
End Sub
Public Sub toDefaultSkin(txtBox As TextBox)
  txtBox.BackColor = vbWhite
  txtBox.ForeColor = vbBlack
End Sub
Public Sub toComboBoxDefaultSkin(cmbBox As ComboBox)
  cmbBox.BackColor = vbWhite
  cmbBox.ForeColor = vbBlack
End Sub
Public Function getFileName(flname As String) As String

    Dim posn As Integer, i As Integer
    Dim fName As String

    posn = 0
    For i = 1 To Len(flname)
        If (Mid(flname, i, 1) = "\") Then posn = i
    Next i

    fName = Right(flname, Len(flname) - posn)

    getFileName = fName
    
End Function
Public Function getImgPath() As String
  getImgPath = App.Path & "\" & Constants.IMG_FOLDER
End Function

Public Function getTemplatesPath() As String
  getTemplatesPath = App.Path & "\" & Constants.TEMPLATE_FOLDER
End Function
Public Function getTempPath() As String
  getTempPath = App.Path & "\" & Constants.TEMP_FOLDER
End Function

