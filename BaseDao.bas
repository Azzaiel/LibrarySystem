Attribute VB_Name = "BaseDao"
Option Explicit
Public Function getStringValue(rs As ADODB.Recordset, index As Integer) As String

  If (Not IsEmpty(rs.Fields(index)) And Not IsNull(rs.Fields(index))) Then
    getStringValue = rs.Fields(index)
  Else
    getStringValue = ""
  End If
  
End Function
