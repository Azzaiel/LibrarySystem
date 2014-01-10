VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpReturnDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   107544577
      CurrentDate     =   41649
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "Return Date"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
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
  frmMain.selectedReturnDate = dtpReturnDate
  Unload Me
End Sub

Private Sub Form_Load()
  dtpReturnDate = DateAdd("ww", 1, Now)
End Sub
