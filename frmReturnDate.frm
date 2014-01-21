VERSION 5.00
Begin VB.Form frmReturnDate 
   Caption         =   "Select Return Date"
   ClientHeight    =   1575
   ClientLeft      =   7455
   ClientTop       =   3285
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3765
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblDueDate 
      BackColor       =   &H8000000E&
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
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "Due Date"
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
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmReturnDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdOk_Click()
  frmMain.selectedReturnDate = CDate(lblDueDate)
  Unload Me
End Sub

Private Sub Form_Load()
  lblDueDate = Format(DateAdd("ww", 1, Now), "mm/dd/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call frmMain.reloadBookStats
End Sub
