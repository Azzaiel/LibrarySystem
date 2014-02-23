VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOverdueItems 
   Caption         =   "Overdue Items"
   ClientHeight    =   3765
   ClientLeft      =   435
   ClientTop       =   2805
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   14550
   Begin MSDataGridLib.DataGrid dgTransactionDash 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOverdueItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private Sub Form_Load()
  Set rs = InventoryDao.getOverdueTransactions
  If (rs.RecordCount > 0) Then
    Set dgTransactionDash.DataSource = rs
    Call formatTransactionDashDatagrid
  Else
    Unload Me
  End If
End Sub
Private Sub formatTransactionDashDatagrid()
    With dgTransactionDash
     'LRN - 0
    .Columns(0).Width = 1500
    .Columns(0).Alignment = dbgCenter


     'DUE DATE - 5
    .Columns(7).Width = 1500
    .Columns(7).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(7).Alignment = dbgCenter
    
    'TRANSACTION_ID
    .Columns(8).Visible = False
    
  End With
End Sub

