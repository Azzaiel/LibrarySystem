Attribute VB_Name = "CommonHelper"
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
