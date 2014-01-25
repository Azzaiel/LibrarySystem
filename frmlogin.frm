VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Login"
   ClientHeight    =   1935
   ClientLeft      =   6675
   ClientTop       =   2715
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4605
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
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
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1440
      MaxLength       =   16
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "2w321312321"
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Username"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblName 
      BackColor       =   &H0080FF80&
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private attempCount As Integer
Private Sub cmbClose_Click()
  End
End Sub

Private Sub cmdSubmit_Click()
  If (Not CommonHelper.hasValidValue(txtUsername.Text)) Then
    MsgBox "Please enter a Username", vbCritical
    Exit Sub
  ElseIf (Not CommonHelper.hasValidValue(txtPassword.Text)) Then
    MsgBox "Please enter a Password", vbCritical
    Exit Sub
  End If
  
  Set rs = UserSession.getUserByUserName(txtUsername)
  
  If (rs.RecordCount > 0) Then
      Dim bytBlock() As Byte
      Dim Hash As New MD5Hash
      bytBlock = StrConv(txtPassword, vbFromUnicode)
      If (UCase(rs!Password) = Hash.HashBytes(bytBlock)) Then
        UserSession.username = rs!username
        UserSession.role = rs!role
        UserSession.forceChange = CommonHelper.extractStringValue(rs!FORCE_CHANGE)
        frmMain.frmControl.Visible = True
        If (rs!role = "Admin") Then
          frmMain.mnTransaction.Visible = True
          frmMain.mnUsers.Visible = True
          frmMain.mnLookups.Visible = True
          frmMain.dbBackum.Visible = True
        Else
          frmMain.mnTransaction.Visible = False
          frmMain.mnUsers.Visible = False
          frmMain.dbBackum.Visible = False
          frmMain.mnLookups.Visible = False
        End If
        txtUsername = ""
        txtPassword = ""
        txtUsername.SetFocus
        frmMain.lblIUser.Caption = "You are currently login as: " & UserSession.getLoginUser
        attempCount = 3
        Unload Me
        If (UserSession.forceChange = "T") Then
          frmChagePass.Show vbModal
        End If
        Exit Sub
      End If
  End If
  
  MsgBox "Username and Password does not match", vbCritical
  
  attempCount = attempCount - 1
  
  If (attempCount = 0) Then
    MsgBox "3 login attemps was reached, System is now on lockdown", vbCritical
    frmAccountLock.Show vbModal
  End If
  
End Sub

Private Sub Form_Load()
  attempCount = 3
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSubmit_Click
  End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSubmit_Click
  End If
End Sub
