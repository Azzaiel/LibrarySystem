VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmQuantityReport 
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Search Form"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.TextBox txtSearchItemCode 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmSearchType 
         Height          =   315
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtSearchName 
         Height          =   285
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cmSearchCategory 
         Height          =   315
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearSearch 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmbExport 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "ISBN"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FF80&
         Caption         =   "Type"
         Height          =   255
         Left            =   7320
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080FF80&
         Caption         =   " Title"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FF80&
         Caption         =   "Category"
         Height          =   255
         Left            =   10080
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid dgQuantityReport 
      Height          =   5895
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   10398
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
Attribute VB_Name = "frmQuantityReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Call populateDataGrid
End Sub
Public Sub populateDataGrid()
  Set rs = InventoryDao.getQuantityReport
  Set dgQuantityReport.DataSource = rs
  dgQuantityReport.Refresh
  Call formatDataGrid
End Sub
Private Sub formatDataGrid()
  With dgQuantityReport
    .Columns(4).Alignment = dbgCenter
    .Columns(4).Width = 1100
    .Columns(5).Alignment = dbgCenter
    .Columns(5).Width = 1100
    .Columns(6).Alignment = dbgCenter
    .Columns(6).Width = 1100
    .Columns(7).Alignment = dbgCenter
    .Columns(7).Width = 1100
    .Columns(8).Alignment = dbgCenter
    .Columns(8).Width = 1100
  End With
End Sub

