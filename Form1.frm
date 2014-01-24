VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Library System"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   20250
   Begin VB.Frame frmControl 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      Begin VB.Frame Frame6 
         Caption         =   "Stats (Double click to view full list of items status)"
         Height          =   2295
         Left            =   13920
         TabIndex        =   50
         Top             =   7440
         Width           =   6135
         Begin MSDataGridLib.DataGrid dgStat 
            Height          =   1335
            Left            =   840
            TabIndex        =   51
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2355
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
            RowDividerStyle =   3
            AllowDelete     =   -1  'True
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
         Begin VB.Label lblTotalBooks 
            Alignment       =   2  'Center
            Caption         =   "totalBooks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   52
            Top             =   1680
            Width           =   4815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Quick Search"
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         Begin VB.ComboBox cmbSearchStatus 
            Height          =   315
            ItemData        =   "Form1.frx":0000
            Left            =   4200
            List            =   "Form1.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   960
            Width           =   1935
         End
         Begin VB.ComboBox cmSearchType 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtSearchItemCode 
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtSearchName 
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   1080
            Width           =   1935
         End
         Begin VB.ComboBox cmSearchCategory 
            Height          =   315
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtSearchAuthor 
            Height          =   285
            Left            =   4200
            TabIndex        =   9
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmItemsQuickSearch 
            Caption         =   "Search"
            Height          =   435
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   2895
         End
         Begin VB.CommandButton cmdClearSearch 
            Caption         =   "Clear"
            Height          =   435
            Left            =   3360
            TabIndex        =   7
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label Label18 
            BackColor       =   &H0080FF80&
            Caption         =   "Status"
            Height          =   255
            Left            =   3240
            TabIndex        =   59
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label15 
            BackColor       =   &H0080FF80&
            Caption         =   " Title"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label14 
            BackColor       =   &H0080FF80&
            Caption         =   "Type"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label8 
            BackColor       =   &H0080FF80&
            Caption         =   "ISBN"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label16 
            BackColor       =   &H0080FF80&
            Caption         =   "Category"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            BackColor       =   &H0080FF80&
            Caption         =   "Author"
            Height          =   255
            Left            =   3240
            TabIndex        =   13
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Detail"
         Height          =   5535
         Left            =   6840
         TabIndex        =   5
         Top             =   120
         Width           =   6975
         Begin VB.Frame Frame7 
            ClipControls    =   0   'False
            Height          =   1455
            Left            =   3960
            TabIndex        =   54
            Top             =   360
            Width           =   1935
            Begin VB.OptionButton optAvailable 
               Caption         =   "Available"
               Enabled         =   0   'False
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
               TabIndex        =   57
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton optLost 
               Caption         =   "Lost"
               Enabled         =   0   'False
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
               TabIndex        =   56
               Top             =   600
               Width           =   1335
            End
            Begin VB.OptionButton optDamage 
               Caption         =   "Damanged"
               Enabled         =   0   'False
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
               TabIndex        =   55
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame fmStudentInfo 
            Caption         =   "Student info"
            Height          =   1335
            Left            =   120
            TabIndex        =   38
            Top             =   3720
            Width           =   6735
            Begin VB.Label Label13 
               BackColor       =   &H0080FF80&
               Caption         =   "LRN"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   1095
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
               TabIndex        =   48
               Top             =   240
               Width           =   2295
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
               TabIndex        =   45
               Top             =   600
               Width           =   2175
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
               TabIndex        =   44
               Top             =   240
               Width           =   2175
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
               TabIndex        =   43
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label lblSelectStudent 
               Caption         =   "Select Student"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2520
               TabIndex        =   42
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label11 
               BackColor       =   &H0080FF80&
               Caption         =   "Adviser"
               Height          =   255
               Left            =   3720
               TabIndex        =   41
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label10 
               BackColor       =   &H0080FF80&
               Caption         =   "Section"
               Height          =   255
               Left            =   3720
               TabIndex        =   40
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label1 
               BackColor       =   &H0080FF80&
               Caption         =   "Student Name"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.ComboBox cmCategory 
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
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   32
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txtDescription 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2040
            Width           =   5055
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
            TabIndex        =   30
            Top             =   2640
            Width           =   5055
         End
         Begin VB.TextBox txtDonatedBy 
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
            TabIndex        =   29
            Top             =   3000
            Width           =   5055
         End
         Begin VB.ComboBox cmStatus 
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
            ItemData        =   "Form1.frx":0049
            Left            =   1440
            List            =   "Form1.frx":0059
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   28
            Top             =   3360
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
            TabIndex        =   23
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox cmItemType 
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
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   22
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtName 
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
            TabIndex        =   21
            Top             =   960
            Width           =   1935
         End
         Begin VB.ComboBox cmLocation 
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
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   20
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label lblChekOut 
            Caption         =   "Check out Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   47
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080FF80&
            Caption         =   "Category"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080FF80&
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label6 
            BackColor       =   &H0080FF80&
            Caption         =   "Author"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label7 
            BackColor       =   &H0080FF80&
            Caption         =   "Donated By"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label9 
            BackColor       =   &H0080FF80&
            Caption         =   "Status"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label lblName 
            BackColor       =   &H0080FF80&
            Caption         =   "ISBN"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080FF80&
            Caption         =   "Type"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080FF80&
            Caption         =   "Title"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl 
            BackColor       =   &H0080FF80&
            Caption         =   "Location"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Location Map"
         Height          =   4095
         Left            =   6960
         TabIndex        =   4
         Top             =   5640
         Width           =   6495
         Begin VB.PictureBox imgLoc 
            Height          =   3690
            Left            =   120
            Picture         =   "Form1.frx":0081
            ScaleHeight     =   3630
            ScaleWidth      =   6195
            TabIndex        =   6
            Top             =   240
            Width           =   6255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dashboard - (Double click to open detail form)"
         ClipControls    =   0   'False
         Height          =   7335
         Left            =   13920
         TabIndex        =   3
         Top             =   120
         Width           =   6135
         Begin MSDataGridLib.DataGrid dgTransactionDash 
            Height          =   6975
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   12303
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
            RowDividerStyle =   3
            AllowDelete     =   -1  'True
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
      Begin VB.Frame Frame2 
         Caption         =   "Result"
         Height          =   6975
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   6495
         Begin MSDataGridLib.DataGrid dgItems 
            Height          =   6615
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   11668
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
      Begin VB.Label lblIUser 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   0
         TabIndex        =   53
         Top             =   9360
         Width           =   6855
      End
   End
   Begin VB.Menu mnLookups 
      Caption         =   "Library"
      Begin VB.Menu mnItemType 
         Caption         =   "Library Materials"
      End
      Begin VB.Menu Categuries 
         Caption         =   "Categories"
      End
      Begin VB.Menu mnLocationMapping 
         Caption         =   "Location Map"
      End
   End
   Begin VB.Menu mnName 
      Caption         =   "School Data"
      Begin VB.Menu sections 
         Caption         =   "Sections"
      End
      Begin VB.Menu mnStudents 
         Caption         =   "Students"
      End
   End
   Begin VB.Menu mnInvetory 
      Caption         =   "Inventory"
   End
   Begin VB.Menu mnTransaction 
      Caption         =   "Transaction Report"
   End
   Begin VB.Menu Account 
      Caption         =   "Account"
      Begin VB.Menu mnUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu mnChangePassword 
         Caption         =   "Changes Password"
      End
   End
   Begin VB.Menu mnAppSession 
      Caption         =   "App Session"
      Begin VB.Menu mnLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private itemsRs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Private transactionRS As ADODB.Recordset
Private statRS As ADODB.Recordset

Private itemTypeItemList() As Variant
Private locationItemList() As Variant
Private categoriesItemList() As Variant

Public selectedStudentID As Integer
Public selectedReturnDate As Date

Private isStatusChangedEnabled As Boolean

Private Sub Categuries_Click()
  frmCategories.Show vbModal
End Sub

Private Sub cmdClearSearch_Click()
  txtSearchItemCode.Text = ""
  cmSearchType.ListIndex = -1
  txtSearchName.Text = ""
  cmSearchCategory.ListIndex = -1
  cmbSearchStatus.ListIndex = -1
  txtSearchAuthor = ""
End Sub
Private Sub cmItemsQuickSearch_Click()
  Set dgItems.DataSource = Nothing
  Call DbInstance.closeRecordSet(itemsRs)
  Set itemsRs = InventoryDao.dashboardSearch(txtSearchItemCode, getSearchItemTypeID, txtSearchAuthor, txtSearchName, getSearchCategoryID, cmbSearchStatus.Text)
  Set dgItems.DataSource = itemsRs
  If (itemsRs.RecordCount = 0) Then
    MsgBox "No record found", vbInformation
  Else
    itemsRs.MoveFirst
    Call showSelectedItem
  End If
  dgItems.Refresh
  'Call clearForm
  Call formatIemsDataGrid
End Sub

Private Sub cmLocation_Click()
 Dim FileName As String
  FileName = LookupDao.getLocationImgName(getLocationID)
  If (FileName <> vbNullString) Then
    imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & FileName)
  Else
     imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & Constants.MISSING_LOC_IMAGE_NAME)
  End If
End Sub

Private Sub cmSearchCategory_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub

Private Sub cmSearchType_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub
Private Sub clearDetailForm()

   txtName.Text = ""
   txtItemCode.Text = ""
   txtDescription.Text = ""
   txtDonatedBy.Text = ""
   txtAuthor.Text = ""
   cmStatus.Text = ""
    
   cmItemType.ListIndex = -1
   cmLocation.ListIndex = -1
   cmCategory.ListIndex = -1
   
End Sub
Private Sub dgItems_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call showSelectedItem
End Sub

Private Sub dgItems_SelChange(Cancel As Integer)
  Call showSelectedItem
End Sub
Private Sub showSelectedItem()

    Call clearStudentInfo
    Call clearDetailForm

    If (itemsRs.RecordCount = 0) Then
      Exit Sub
    End If
 
    txtName.Text = CommonHelper.extractStringValue(itemsRs!name)
    txtItemCode.Text = CommonHelper.extractStringValue(itemsRs!ITEM_CODE)
    txtDescription.Text = CommonHelper.extractStringValue(itemsRs!Description)
    txtDonatedBy.Text = CommonHelper.extractStringValue(itemsRs!DONATED_BY)
    txtAuthor.Text = CommonHelper.extractStringValue(itemsRs!author)
    cmStatus.Text = CommonHelper.extractStringValue(itemsRs!status)
    
    cmItemType.ListIndex = -1
    cmLocation.ListIndex = -1
    cmCategory.ListIndex = -1
    
    Dim index As Integer

   For index = 0 To UBound(itemTypeItemList)
     If (itemsRs!ITEM_TYPE_ID = Val(itemTypeItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmItemType.ListIndex = index
     End If
   Next index
   
   For index = 0 To UBound(locationItemList)
     If (itemsRs!LOCATION_ID = Val(locationItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmLocation.ListIndex = index
     End If
   Next index
   
   For index = 0 To UBound(categoriesItemList)
     If (itemsRs!CATEGORY_ID = Val(categoriesItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmCategory.ListIndex = index
     End If
   Next index
   
   isStatusChangedEnabled = False
   Call toogelItemCheckOutUI(False)
   optDamage.Enabled = False
   optLost.Enabled = False
   optAvailable.Enabled = False
   optDamage.value = False
   optLost.value = False
   optAvailable.value = False
   
   If (cmStatus = "Available") Then
      Call toogelItemCheckOutUI(True)
      optAvailable.value = True
      optDamage.Enabled = True
      optLost.Enabled = True
   ElseIf (cmStatus = "Borrowed") Then
      Set tempRs = InventoryDao.getStudentBorrower(itemsRs!id)
      txtLRN = tempRs!lrn
      txtStudentName = tempRs!STUDENT_NAME
      txtAdviser = tempRs!Adviser
      txtSection = tempRs!Section
      Call DbInstance.closeRecordSet(tempRs)
   ElseIf (cmStatus = "Damaged") Then
      optDamage.value = True
      optLost.Enabled = True
      optAvailable.Enabled = True
   ElseIf (cmStatus = "Loss") Then
      optLost.value = True
      optDamage.Enabled = True
      optAvailable.Enabled = True
   End If
   
   isStatusChangedEnabled = True
   
End Sub

Private Sub toogelItemCheckOutUI(isAvailable As Boolean)
  fmStudentInfo.Enabled = isAvailable
  lblChekOut.Enabled = isAvailable
  lblSelectStudent.Enabled = isAvailable
  If (isAvailable) Then
  
    txtStudentName.BackColor = vbWhite
    txtAdviser.BackColor = vbWhite
    txtSection.BackColor = vbWhite
    txtLRN.BackColor = vbWhite
    
    txtStudentName.ForeColor = vbBlack
    txtAdviser.ForeColor = vbBlack
    txtSection.ForeColor = vbBlack
    txtLRN.ForeColor = vbBlack
    
  Else
  
    txtStudentName.BackColor = vbGrayText
    txtAdviser.BackColor = vbGrayText
    txtSection.BackColor = vbGrayText
    txtLRN.BackColor = vbGrayText
    
    txtStudentName.ForeColor = vbWhite
    txtAdviser.ForeColor = vbWhite
    txtSection.ForeColor = vbWhite
    txtLRN.ForeColor = vbWhite
    
  End If
End Sub

Private Sub dgStat_DblClick()
  If (statRS.State <> adStateClosed) Then
    If (statRS.RecordCount <> 0) Then
      frmInventory.cmbSearchStatus.Text = statRS!Books
      frmInventory.Show vbModal
    End If
  End If
End Sub

Private Sub dgTransactionDash_DblClick()
   If (transactionRS.RecordCount > 0) Then
     frmItemReturn.transactionID = transactionRS!Transaction_ID
     frmItemReturn.Show vbModal
     Call populateTransactionDatagrid
     Call cmItemsQuickSearch_Click
   End If
End Sub

Private Sub fmStudentInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lblSelectStudent.ForeColor = vbBlue
End Sub
Private Sub Form_Load()
  Call populateDropDown
  Call initiateItemsRs
  Call populateTransactionDatagrid
  Call reloadBookStats
End Sub
Public Sub reloadBookStats()
   Set statRS = InventoryDao.getBookStatRs
   Set dgStat.DataSource = statRS
   Dim totalBooks As Long
   totalBooks = 0
   While Not statRS.EOF
     totalBooks = totalBooks + Val(statRS!Total)
     statRS.MoveNext
   Wend
   
   lblTotalBooks = "Total Books: " & totalBooks
   
   dgStat.Refresh
End Sub

Private Sub populateTransactionDatagrid()
  Set transactionRS = InventoryDao.getTransactionDashboardRs
  Set dgTransactionDash.DataSource = transactionRS
  dgTransactionDash.Refresh
  Call formatTransactionDashDatagrid
End Sub
Private Sub formatTransactionDashDatagrid()
    With dgTransactionDash
     'LRN - 0
    .Columns(0).Width = 1500
    .Columns(0).Alignment = dbgCenter
    
    .Columns(1).Caption = "Title"

     'DUE DATE - 5
    .Columns(5).Width = 1500
    .Columns(5).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(5).Alignment = dbgCenter
    
    'TRANSACTION_ID
    .Columns(6).Visible = False
    
  End With
End Sub
Private Sub populateDropDown()
  Dim index As Integer

  itemTypeItemList = LookupDao.getItemTypeItemList
  cmSearchType.Clear
  cmItemType.Clear
  For index = 0 To UBound(itemTypeItemList)
    cmSearchType.AddItem (itemTypeItemList(index, Constants.ITEM_LABEL_INDEX))
    cmItemType.AddItem (itemTypeItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
  locationItemList = LookupDao.getLocationMappingItemList
  cmLocation.Clear
  For index = 0 To UBound(locationItemList)
     cmLocation.AddItem (locationItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
  categoriesItemList = LookupDao.getCategoriesItemList
  cmSearchCategory.Clear
  cmCategory.Clear
  For index = 0 To UBound(categoriesItemList)
    cmSearchCategory.AddItem (categoriesItemList(index, Constants.ITEM_LABEL_INDEX))
    cmCategory.AddItem (categoriesItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
End Sub
Private Sub initiateItemsRs()
  Set itemsRs = InventoryDao.getDashboardEmptyRs
  Set dgItems.DataSource = itemsRs
  dgItems.Refresh
  Call formatIemsDataGrid
End Sub
Private Function getSearchCategoryID() As Integer
  Dim index As Integer
  index = cmSearchCategory.ListIndex
  If (index <> -1) Then
    getSearchCategoryID = categoriesItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getSearchCategoryID = 0
  End If
End Function

Private Function getSearchItemTypeID() As Integer
  Dim index As Integer
  index = cmSearchType.ListIndex
  If (index <> -1) Then
    getSearchItemTypeID = itemTypeItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getSearchItemTypeID = 0
  End If
End Function
Private Sub formatIemsDataGrid()
  If (itemsRs.RecordCount > 0) Then
    itemsRs.MoveFirst
    'Call showSelectedItem
  Else
    'Call clearForm
  End If
  
  With dgItems
  
     .Columns(1).Caption = "ISBN"
  
    .Columns(3).Caption = "Title"

     'CREATED DATE - 11
    .Columns(11).Width = 1500
    .Columns(11).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(11).Alignment = dbgCenter
    
    'LAST MOD DATE - 13
    .Columns(13).Width = 1500
    .Columns(13).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(13).Alignment = dbgCenter
    
    .Columns(13).Visible = False
    .Columns(14).Visible = False
    .Columns(15).Visible = False
    .Columns(16).Visible = False
    
  End With
End Sub
Private Function getLocationID() As Integer
  Dim index As Integer
  index = cmLocation.ListIndex
  If (index <> -1) Then
    getLocationID = locationItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getLocationID = 0
  End If
End Function



Private Sub clearStudentInfo()
    txtStudentName = ""
    txtSection = ""
    txtAdviser = ""
    selectedStudentID = 0
    txtLRN = ""
End Sub



Private Sub lblChekOut_Click()
  If (Not CommonHelper.hasValidValue(txtItemCode.Text)) Then
    MsgBox "Please select an Item", vbCritical
    Exit Sub
  End If
  If (selectedStudentID > 0) Then
    selectedReturnDate = vbNull
    frmReturnDate.Show vbModal
    If (selectedReturnDate <> vbNull) Then
      Set tempRs = InventoryDao.getFakeTransactionRS
      tempRs.AddNew
      tempRs!ITEM_ID = itemsRs!id
      tempRs!STUDENT_ID = selectedStudentID
      tempRs!LEND_DATE = Now
      tempRs!LEND_BY = UserSession.getLoginUser
      tempRs!REQUESTED_RETURN_DATE = selectedReturnDate
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      Set tempRs = InventoryDao.getRsByID(itemsRs!id)
      tempRs!status = "Borrowed"
      tempRs!LAST_MOD_BY = UserSession.getLoginUser
      tempRs!LAST_MOD_DATE = Now
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      MsgBox "Transaction Successful"
      Call reloadBookStats
      Call cmItemsQuickSearch_Click
      Call clearDetailForm
      Call populateTransactionDatagrid
    Else
      MsgBox "System cannot procced without retrun date", vbCritical
    End If
  Else
    MsgBox "Please select a Student", vbCritical
  End If
End Sub

Private Sub lblChekOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblChekOut.ForeColor = vbRed
End Sub

Private Sub lblLost_Click()

End Sub

Private Sub lblMarkAvailable_Click()
 
End Sub

Private Sub lblMarkAvailable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMarkAvailable.ForeColor = vbRed
End Sub

Private Sub lblMarkDamage_Click()
  
End Sub

Private Sub lblMarkDamage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMarkDamage.ForeColor = vbRed
End Sub

Private Sub lblMarkLost_Click()
 
End Sub

Private Sub lblMarkLost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMarkLost.ForeColor = vbRed
End Sub

Private Sub lblSelectStudent_Click()
 Call clearStudentInfo
 frmStudentSelect.Show vbModal
 lblSelectStudent.ForeColor = vbBlue
End Sub

Private Sub lblSelectStudent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblSelectStudent.ForeColor = vbRed
End Sub

Private Sub mnChangePassword_Click()
  frmChagePass.Show vbModal
End Sub

Private Sub mnExit_Click()
    Dim response As String
    response = MsgBox("Are you sure you want to exit system?", vbYesNo, "Question")
    If (response = vbYes) Then
      End
    End If
End Sub

Private Sub mnInvetory_Click()
  frmInventory.Show vbModal
End Sub

Private Sub mnItemType_Click()
  frmItemTypes.Show vbModal
End Sub

Private Sub mnLocationMapping_Click()
  frmLocationMapping.Show vbModal
End Sub

Private Sub mnLogout_Click()
  Dim response As String
    response = MsgBox("Are you sure you want to logout from system?", vbYesNo, "Question")
    If (response = vbYes) Then
      frmControl.Visible = False
      frmlogin.Show vbModal
    End If
End Sub

Private Sub mnStudents_Click()
  frmStudents.Show vbModal
End Sub

Private Sub mnTransaction_Click()
  frmTransactionReport.Show vbModal
End Sub

Private Sub mnUsers_Click()
  frmAccount.Show vbModal
End Sub

Private Sub optAvailable_Click()
   If (isStatusChangedEnabled) Then
    Call DbInstance.closeRecordSet(tempRs)
    Set tempRs = InventoryDao.getRsByID(itemsRs!id)
    tempRs!status = "Available"
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Item Status updated", vbInformation
    Call cmItemsQuickSearch_Click
    Call reloadBookStats
  End If
End Sub

Private Sub optDamage_Click()
  If (isStatusChangedEnabled) Then
    Call DbInstance.closeRecordSet(tempRs)
    Set tempRs = InventoryDao.getRsByID(itemsRs!id)
    tempRs!status = "Damaged"
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Item Status updated", vbInformation
    Call cmItemsQuickSearch_Click
    Call reloadBookStats
  End If
  
End Sub

Private Sub optLost_Click()
   If (isStatusChangedEnabled) Then
    Call DbInstance.closeRecordSet(tempRs)
    Set tempRs = InventoryDao.getRsByID(itemsRs!id)
    tempRs!status = "Loss"
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Item Status updated", vbInformation
    Call cmItemsQuickSearch_Click
    Call reloadBookStats
  End If
   
End Sub

Private Sub sections_Click()
  frmSections.Show vbModal
End Sub

Private Sub txtSearchAuthor_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub

Private Sub txtSearchItemCode_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub

Private Sub txtSearchName_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub
