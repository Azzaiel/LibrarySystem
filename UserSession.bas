Attribute VB_Name = "UserSession"
Option Explicit
Private userName As String
Public Function getLoginUser() As String

  If (userName <> vbNullString) Then
   getLoginUser = userName
  Else
    getLoginUser = "System"
  End If
  
End Function






