Attribute VB_Name = "MainModule"
Option Explicit
Private rs As ADODB.Recordset
Private Const ADMIN_USERNAME As String = "admin"
Sub main()
   
   Set rs = UserSession.getUserByUserName(ADMIN_USERNAME)
   
   If (rs.RecordCount > 0) Then
     rs!Password = stringToMD5(ADMIN_USERNAME)
     'rs!FORCE_CHANGE = "T"
     rs.Update
     MsgBox "Admin password was reset to default", vbInformation
   Else
    rs.AddNew
    rs!username = ADMIN_USERNAME
    rs!role = "Admin"
    'rs!FORCE_CHANGE = "T"
    rs!Password = stringToMD5(ADMIN_USERNAME)
    rs.Update
    MsgBox "There was no Admin account found.... System has created a defualt Admin user", vbInformation
   End If
   
   Call DbInstance.closeRecordSet(rs)
   
End Sub

Public Function stringToMD5(strPassword As String) As String
   Dim bytBlock() As Byte
   Dim Hash As New MD5Hash
   bytBlock = StrConv(strPassword, vbFromUnicode)
   stringToMD5 = Hash.HashBytes(bytBlock)
End Function
