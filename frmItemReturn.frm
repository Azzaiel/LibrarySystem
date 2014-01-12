VERSION 5.00
Begin VB.Form frmItemReturn 
   ClientHeight    =   4935
   ClientLeft      =   5565
   ClientTop       =   2145
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7230
   Begin VB.TextBox txtCategory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtItemType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student info"
      Height          =   1095
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   6735
      Begin VB.Label Label16 
         BackColor       =   &H0080FF80&
         Caption         =   "Date Borrowed"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label txtBorrowedDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label txtRemainingDays 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5040
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label txtDueDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1320
         TabIndex        =   22
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Remaning days"
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Due Date"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Close"
      Height          =   495
      Left            =   3720
      TabIndex        =   18
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmbNewRec 
      Caption         =   "Mark as Returned"
      Height          =   495
      Left            =   1320
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtItemName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtItemCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtAuthor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Frame fmStudentInfo 
      Caption         =   "Student info"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   6735
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "Student Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "Section"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FF80&
         Caption         =   "Adviser"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label txtStudentName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label txtSection 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label txtAdviser 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4440
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label txtLRN 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FF80&
         Caption         =   "LRN"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Item Name"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "Type"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblName 
      BackColor       =   &H0080FF80&
      Caption         =   "Item Code"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Caption         =   "Author"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Category"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmItemReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Public transactionID As Integer
Private itemID As Integer

Private Sub cmbNewRec_Click()
  Dim response As String
  response = MsgBox("Are your sure you want to proceed?", vbOKCancel, "Question")
  If (response = vbOK) Then
      Set tempRs = InventoryDao.getTransaction(transactionID)
      itemID = tempRs!ITEM_ID
      tempRs!RETURN_DATE = Now
      tempRs!RECEIVED_BY = UserSession.getLoginUser
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      
      Set tempRs = InventoryDao.getRsByID(itemID)
      tempRs!Status = "Available"
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      
      MsgBox "Update Compalte"
      Unload Me
      
  End If
End Sub

Private Sub cmdClosed_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Set rs = InventoryDao.getTransactionInfo(transactionID)
  If (rs.RecordCount > 0) Then
    'Item
    txtItemCode = CommonHelper.extractStringValue(rs!ITEM_CODE)
    txtItemType = CommonHelper.extractStringValue(rs!ITEM_TYPE)
    txtCategory = CommonHelper.extractStringValue(rs!CATEGORY)
    txtItemName = CommonHelper.extractStringValue(rs!ITEM_NAME)
    txtAuthor = CommonHelper.extractStringValue(rs!author)
    'Student
    txtLRN = CommonHelper.extractStringValue(rs!lrn)
    txtStudentName = CommonHelper.extractStringValue(rs!STUDENT_NAME)
    txtAdviser = CommonHelper.extractStringValue(rs!Adviser)
    txtSection = CommonHelper.extractStringValue(rs!Section)
    'TRANSACTION
    txtBorrowedDate = CommonHelper.extractDateValue(rs!BORROWED_DATE)
    txtDueDate = CommonHelper.extractDateValue(rs!DUE_DATE)
    txtRemainingDays = CommonHelper.extractStringValue(rs!REMAINING_DAYS)
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call DbInstance.closeRecordSet(rs)
End Sub
