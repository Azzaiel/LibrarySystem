VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   6675
   ClientTop       =   2715
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   7140
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

