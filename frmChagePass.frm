VERSION 5.00
Begin VB.Form frmChagePass 
   Caption         =   "Forget Pasword"
   ClientHeight    =   2730
   ClientLeft      =   4260
   ClientTop       =   2910
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   5250
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmbClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtConfirmPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtNewPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtCurrentPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   0
         ToolTipText     =   "2w321312321"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Confirm Password"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "New Password"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "Current Password"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmChagePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private Sub cmbClose_Click()
  Unload Me
End Sub
Private Function isFormValid() As Boolean
  If (Not CommonHelper.hasValidValue(txtCurrentPass.Text)) Then
     Call CommonHelper.sendWarning(txtCurrentPass, "Please Enter your Current password")
     isFormValid = False
     Exit Function
  ElseIf (Not CommonHelper.hasValidValue(txtNewPass.Text)) Then
     Call CommonHelper.sendWarning(txtNewPass, "Please Enter your New password")
     isFormValid = False
     Exit Function
  ElseIf (Not CommonHelper.hasValidValue(txtConfirmPass.Text)) Then
     Call CommonHelper.sendWarning(txtConfirmPass, "Please Enter your New password")
     isFormValid = False
     Exit Function
  ElseIf (txtNewPass.Text <> txtConfirmPass.Text) Then
     MsgBox "New password and Confirm password does not match", vbCritical
     isFormValid = False
     Exit Function
  End If
  isFormValid = True
End Function
Private Sub resetFromSkin()
 Call CommonHelper.toDefaultSkin(txtCurrentPass)
 Call CommonHelper.toDefaultSkin(txtNewPass)
 Call CommonHelper.toDefaultSkin(txtConfirmPass)
End Sub
Private Sub cmdSubmit_Click()
  Call resetFromSkin
  If (isFormValid) Then
    Set rs = UserSession.getUserByUserName(UserSession.getLoginUser)
    If (rs.RecordCount = 1) Then
      Dim bytBlock() As Byte
      Dim Hash As New MD5Hash
      bytBlock = StrConv(txtCurrentPass, vbFromUnicode)
      If (UCase(rs!Password) = Hash.HashBytes(bytBlock)) Then
        bytBlock = StrConv(txtNewPass, vbFromUnicode)
        rs!Password = Hash.HashBytes(bytBlock)
        rs!FORCE_CHANGE = "F"
        rs.Update
        UserSession.forceChange = "F"
        MsgBox "Password was successfuly updated", vbInformation
        Unload Me
      Else
        MsgBox "Current Password is incorrect"
      End If
    Else
      MsgBox "System error, Please contact your Administrator"
    End If
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If (UserSession.forceChange = "T") Then
    Dim response As String
    response = MsgBox("Your are not allowed use the system unless you change your password. System will close if you contineu", vbOKCancel, "Question")
    If (response = vbOK) Then
       frmMain.frmControl.Visible = False
        Me.Hide
       frmlogin.Show vbModal
       Exit Sub
    Else
      Cancel = 1
    End If
  End If
End Sub

Private Sub txtConfirmPass_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSubmit_Click
  End If
End Sub

Private Sub txtCurrentPass_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSubmit_Click
  End If
End Sub

Private Sub txtNewPass_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSubmit_Click
  End If
End Sub
