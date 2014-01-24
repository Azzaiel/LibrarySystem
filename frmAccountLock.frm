VERSION 5.00
Begin VB.Form frmAccountLock 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1410
   ClientLeft      =   8655
   ClientTop       =   3255
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
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
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmbNewRec 
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
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblName 
      BackColor       =   &H0080FF80&
      Caption         =   "Enter Master Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmAccountLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private Const ADMIN_USERNAME As String = "admin"
Private Sub cmbNewRec_Click()
   If (txtPass.Text = "qwerty123") Then
      Call restAdminPass
      frmlogin.txtUserName = ADMIN_USERNAME
      frmlogin.txtPassword = ADMIN_USERNAME
      Unload Me
      frmlogin.cmdSubmit.value = True
   Else
     MsgBox "Incorrect masterkey"
   End If
End Sub
Sub restAdminPass()
   
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
Private Sub Command1_Click()
  End
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
    Call cmbNewRec_Click
  End If
End Sub
