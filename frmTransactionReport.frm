VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransactionReport 
   Caption         =   "Tansaction Report"
   ClientHeight    =   8220
   ClientLeft      =   495
   ClientTop       =   450
   ClientWidth     =   18165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   18165
   Begin VB.Frame Frame1 
      Caption         =   "Select Date Range by Borrow Date"
      Height          =   975
      Left            =   3960
      TabIndex        =   4
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dpStartDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110690305
         CurrentDate     =   41650
      End
      Begin MSComCtl2.DTPicker dpEndDate 
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110690305
         CurrentDate     =   41650
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dgReport 
      Height          =   6975
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   12303
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
Attribute VB_Name = "frmTransactionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset

Private Sub cmdExport_Click()
  Dim excelApp As New Excel.Application
  Dim oBook As New Excel.Workbook
  Dim oSheet As New Excel.Worksheet
  
  Set excelApp = CreateObject("Excel.Application")
  Set oBook = excelApp.Workbooks.Open(CommonHelper.getTemplatesPath & "\" & Constants.TRANSACTION_REPORT_TEMPLATE)
  Set oSheet = excelApp.Worksheets(1)
  
  oSheet.name = "Transaction Report"
  
  oSheet.Range("A2").CopyFromRecordset dgReport.DataSource
  oSheet.Columns.AutoFit

  excelApp.DisplayAlerts = False
  oBook.SaveAs CommonHelper.getTempPath & "\" & Constants.TEMP_WORK_BOOK
  'excelApp.Visible = True
  
    Dim pdfFilePat As String
    pdfFilePat = CommonHelper.getTempPath & "\temp_" & Format(Now, "mmhhyysssh") & ".pdf"
    Call oBook.ExportAsFixedFormat(xlTypePDF, pdfFilePat, xlQualityStandard, False, True)
    oBook.Close
    Call CommonHelper.openFile(pdfFilePat, Me.hWnd)

End Sub

Private Sub cmdSearch_Click()
  Set dgReport.DataSource = Nothing
  Call DbInstance.closeRecordSet(rs)
  Set rs = InventoryDao.getTransactionReport(dpStartDate.value, DateAdd("d", 1, dpEndDate.value))
  Set dgReport.DataSource = rs
  dgReport.Refresh
  Call formatDataGrid
  If (rs.RecordCount = 0) Then
    MsgBox "No Record found", vbInformation
  End If
  
End Sub
Private Sub formatDataGrid()
  With dgReport
    .Columns(0).Width = 1500
    .Columns(1).Width = 1250
    .Columns(2).Width = 1250
    .Columns(3).Width = 1800
    .Columns(4).Width = 2250
    .Columns(5).Width = 1200
    .Columns(6).Width = 2250
    .Columns(7).Width = 2250
    .Columns(8).Width = 1250
    .Columns(9).Width = 1250
    .Columns(10).Width = 1600
    .Columns(10).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(11).Width = 1500
    .Columns(11).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(12).Width = 1500
    .Columns(12).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(13).Width = 1200
  End With
End Sub

Private Sub Form_Load()
  dpStartDate = DateAdd("m", -1, Now)
  dpEndDate = Now
  Set dgReport.DataSource = Nothing
  Call DbInstance.closeRecordSet(rs)
  Set rs = InventoryDao.getFakeTransactionReportRs
  Set dgReport.DataSource = rs
  dgReport.Refresh
  Call formatDataGrid
  End Sub
