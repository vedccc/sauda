VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmCloseRate 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   18330
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "ERGGHE"
      Height          =   2895
      Left            =   2400
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   12135
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   11895
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   11400
         TabIndex        =   17
         ToolTipText     =   "Close"
         Top             =   -15
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Check Trade"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   80
         TabIndex        =   19
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   16695
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   16695
         Begin VB.Label Label7 
            BackColor       =   &H00FF8080&
            Caption         =   "Closing Rate Entry"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   615
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   16695
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7845
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   16695
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copy Rates To Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   21
         Top             =   120
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   600
         Top             =   7080
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2865
         Left            =   3360
         TabIndex        =   4
         Top             =   4560
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5054
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
         TabAction       =   2
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Daily Closing Rates"
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "CONTDATE"
            Caption         =   "     Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "SAUDACODE"
            Caption         =   "Sauda Code"
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
         BeginProperty Column02 
            DataField       =   "SAUDANAME"
            Caption         =   "Sauda Name"
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
         BeginProperty Column03 
            DataField       =   "CLOSING"
            Caption         =   "Settle Rate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "OPENING"
            Caption         =   "Open"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "LOW"
            Caption         =   "Low"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "HIGH"
            Caption         =   "High "
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "CLRate"
            Caption         =   "Close"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "DataImport"
            Caption         =   "DataImport"
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
         BeginProperty Column09 
            DataField       =   "itemcode"
            Caption         =   "ItemCode"
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
         BeginProperty Column10 
            DataField       =   "EXCODE"
            Caption         =   "ExCode"
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   4004.788
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   2894.74
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filter Contracts With Standing"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12960
         TabIndex        =   15
         Top             =   0
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filter Zero Closing Rates "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12960
         TabIndex        =   14
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox TxtFilterSauda 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         TabIndex        =   3
         Top             =   120
         Width           =   3135
      End
      Begin MSDataListLib.DataCombo DComboExchnage 
         Height          =   360
         Left            =   3120
         TabIndex        =   2
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DComboSauda 
         Height          =   420
         Left            =   600
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711680
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthForeColor  =   0
         Value           =   37905.9259606482
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   7290
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   16695
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   7065
            Left            =   -120
            TabIndex        =   20
            Top             =   240
            Width           =   16455
            _ExtentX        =   29025
            _ExtentY        =   12462
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   21
            TabAction       =   2
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Daily Closing Rates"
            ColumnCount     =   15
            BeginProperty Column00 
               DataField       =   "CONTDATE"
               Caption         =   "    DATE"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "SAUDACODE"
               Caption         =   "Sauda Code"
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
            BeginProperty Column02 
               DataField       =   "SAUDANAME"
               Caption         =   "Sauda name"
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
            BeginProperty Column03 
               DataField       =   "CLOSING"
               Caption         =   "Settle Rate"
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
            BeginProperty Column04 
               DataField       =   "OPENING"
               Caption         =   "Open"
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
            BeginProperty Column05 
               DataField       =   "LOW"
               Caption         =   "Low"
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
            BeginProperty Column06 
               DataField       =   "HIGH"
               Caption         =   "High"
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
            BeginProperty Column07 
               DataField       =   "CLRATE"
               Caption         =   "Close"
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
            BeginProperty Column08 
               DataField       =   "DataImport"
               Caption         =   "DataImport"
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
            BeginProperty Column09 
               DataField       =   "ITEMCODE"
               Caption         =   "Item Code"
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
            BeginProperty Column10 
               DataField       =   "EXCODE"
               Caption         =   "Ex. Code"
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
            BeginProperty Column11 
               DataField       =   "SAUDAID"
               Caption         =   "SAUDAID"
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
            BeginProperty Column12 
               DataField       =   "ITEMID"
               Caption         =   "ITEMID"
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
            BeginProperty Column13 
               DataField       =   "EXID"
               Caption         =   "EXID"
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
            BeginProperty Column14 
               DataField       =   "PRATE"
               Caption         =   "PRate"
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
               MarqueeStyle    =   2
               BeginProperty Column00 
                  ColumnWidth     =   2385.071
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2085.166
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   2085.166
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column07 
                  Locked          =   -1  'True
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1065.26
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column10 
                  Locked          =   -1  'True
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column14 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1500.095
               EndProperty
            EndProperty
         End
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   11160
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthForeColor  =   0
         Value           =   37905.9259606482
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Sauda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   4560
         TabIndex        =   12
         Top             =   165
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ExCode"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   2280
         TabIndex        =   11
         Top             =   165
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   165
         Width           =   465
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I t e m   L i s t"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17280
      TabIndex        =   13
      Top             =   960
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000011&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   7740
      Left            =   0
      Top             =   840
      Width           =   16620
   End
End
Attribute VB_Name = "FrmCloseRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LSSaudas  As String:              Public Fb_Press As Byte:            Dim Rec As ADODB.Recordset
Dim RECGRID As ADODB.Recordset:       Dim RecSauda As ADODB.Recordset:     Dim ExRec As ADODB.Recordset
Dim LSExCode As String:               Dim ldate As Date
Dim LSaudaIDS As String
Dim LFilterRec  As ADODB.Recordset
Dim LFilterSaudaID As String
Sub Add_Rec()
    If RecSauda.RecordCount > 0 Then
        Frame1.Enabled = True:        Fb_Press = 1
        Label1.Visible = True:        vcDTP1.Visible = True
        Call Get_Selection(1):
        vcDTP1.Enabled = True:        DComboExchnage.Enabled = True
        vcDTP1.SetFocus:              DataGrid2.LeftCol = 0
        DataGrid2.Columns(0).text = vcDTP1.Value  'Date
    Else
       ' MsgBox "Please Select Sauda", vbCritical
       ' Call CANCEL_REC
    End If
End Sub
Sub CANCEL_REC()
    LSExCode = vbNullString
    Fb_Press = 0:                           Call Get_Selection(10)
    
    'DataGrid2.Columns(0).Locked = False:
    
    DComboSauda.Visible = False
    vcDTP1.Enabled = True:                  DComboExchnage.Enabled = True
    Frame1.Enabled = False:                 Call RecSet
    Set RecSauda = Nothing:    Set RecSauda = New ADODB.Recordset
    mysql = "SELECT I.EXID,I.ITEMID,I.EXCHANGECODE,S.SAUDAID,S.SAUDACODE,S.SAUDANAME,I.ITEMCODE,S.MATURITY FROM SAUDAMAST AS S,ITEMMAST AS I "
    mysql = mysql & " WHERE S.COMPCODE=" & GCompCode & " AND S.COMPCODE =I.COMPCODE AND S.ITEMID =I.ITEMID "
    mysql = mysql & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY I.ITEMCODE,MATURITY "
    RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RecSauda.EOF Then
        Set DComboSauda.RowSource = RecSauda
        DComboSauda.ListField = "SAUDANAME"
        DComboSauda.BoundColumn = "SAUDACODE"
    End If
End Sub
Sub MODIFY_REC()
    If RecSauda.RecordCount > 0 Then
        Fb_Press = 2
        Frame1.Enabled = True
        Call Get_Selection(2)
        DataGrid2.Columns(0).Locked = True:
        Label1.Visible = True: vcDTP1.Visible = True: vcDTP1.Enabled = True: Frame1.Enabled = True:
        DComboExchnage.Enabled = True
        If vcDTP1.Enabled Then vcDTP1.SetFocus
    Else
        MsgBox "Please Select Sauda.", vbCritical
        Call CANCEL_REC
    End If
End Sub
Sub Save_Rec()
    On Error GoTo err1
    Dim LFromDt As Date:    Dim MSauda As String:   Dim LSaudaID As Long:    Dim LConSno  As Long
    Dim LExID As Integer
    Dim LItemID As Integer
    Dim LExIDS As String
    Dim LSaudaIDS  As String
    DComboSauda.Visible = False
    Cnn.BeginTrans
    CNNERR = True
    RECGRID.Sort = "CONTDATE"
    RECGRID.MoveFirst
    LFromDt = DateValue(RECGRID!CONTDATE)
    Do While Not RECGRID.EOF
        MSauda = RECGRID!saudacode
        LFromDt = RECGRID!CONTDATE
        LSaudaID = RECGRID!SAUDAID
        LItemID = RECGRID!itemid
        If Check2.Value = 1 Then
            LFromDt = vcDTP2.Value
        End If
        If LenB(MSauda) <> 0 Then
            'LSaudaID = Get_SaudaID(MSauda)
            mysql = "DELETE FROM CTR_R WHERE COMPCODE=" & GCompCode & "  AND CONDATE='" & Format(LFromDt, "yyyy/MM/dd") & "' AND SAUDAID=" & LSaudaID & " "
            Cnn.Execute mysql
            'If Val(RECGRID!CLOSING) <> 0 Then
                If Not InStr(LSaudaIDS, LSaudaID) Then
                    If LenB(LSaudaIDS) > 0 Then LSaudaIDS = LSaudaIDS & ","
                    LSaudaIDS = LSaudaIDS & str(LSaudaID)
                End If
                LExID = RECGRID!EXID
                If Not InStr(LExIDS, LExID) Then
                    If LenB(LExIDS) > 0 Then LExIDS = LExIDS & ","
                    LExIDS = LExIDS & str(LExID)
                End If
                LConSno = 0:
                If LenB(LFromDt) > 0 And LenB(RECGRID!saudacode) > 0 And RECGRID!CLOSING <> 0 Then
                    LConSno = Get_ConSNo(LFromDt, RECGRID!saudacode, RECGRID!ITEMCODE, RECGRID!excode, LSaudaID, LItemID, LExID)
                    Call PInsert_Ctr_R(LConSno, MSauda, DateValue(LFromDt), IIf(IsNull(RECGRID!OPENING), 0, RECGRID!OPENING), IIf(IsNull(RECGRID!HIGH), 0, RECGRID!HIGH), IIf(IsNull(RECGRID!LOW), 0, RECGRID!LOW), IIf(IsNull(RECGRID!CLOSING), 0, RECGRID!CLOSING), IIf(IsNull(RECGRID!ClRate), 0, RECGRID!ClRate), (RECGRID!excode & vbNullString), (RECGRID!ITEMCODE & vbNullString), LSaudaID, LItemID, LExID, IIf(IsNull(RECGRID!PRATE), 0, RECGRID!PRATE), "")
                End If
            'End If
        End If
        RECGRID.MoveNext
    Loop
    Cnn.CommitTrans: CNNERR = False
    Dim LBillParties As String
    Dim TRec As ADODB.Recordset
    If LenB(LExIDS) > 0 Then
        LBillParties = vbNullString
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        mysql = "SELECT DISTINCT PARTY FROM CTR_D WHERE EXID IN (" & LExIDS & ") "
        If LenB(LSaudaIDS) > 0 Then mysql = mysql & " AND SAUDAID  IN (" & LSaudaIDS & ")"
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        Do While Not TRec.EOF
            If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ","
            LBillParties = LBillParties & "'" & TRec!PARTY & "'"
            TRec.MoveNext
        Loop
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        mysql = "SELECT DISTINCT PARTY FROM INV_D WHERE EXID IN (" & LExIDS & ") "
        If LenB(LSaudaIDS) > 0 Then mysql = mysql & " AND SAUDAID  IN (" & LSaudaIDS & ")"
        If LenB(LBillParties) > 0 Then mysql = mysql & " AND PARTY NOT IN (" & LBillParties & ")"
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        Do While Not TRec.EOF
            If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ","
            LBillParties = LBillParties & "'" & TRec!PARTY & "'"
            TRec.MoveNext
        Loop
        If LenB(LBillParties) > 0 Then
            Cnn.BeginTrans
            CNNERR = True
            Call RATE_TEST(LFromDt, LExIDS, , FrmCloseRate)
            'Call Delete_Inv_D(vbNullString, Trim(Str(LEXID)), vbNullString, LFromDt)
            If BILL_GENERATION(CDate(LFromDt), CDate(GFinEnd), LSaudaIDS, LBillParties, LExIDS) Then
                Cnn.CommitTrans: CNNERR = False
            Else
                Cnn.RollbackTrans: CNNERR = False:
                Call CANCEL_REC: Exit Sub
            End If
        End If
    End If
    LSExCode = vbNullString
    'Call Chk_Billing
    Call CANCEL_REC
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
       
       Cnn.RollbackTrans: CNNERR = False
    End If
End Sub

Private Sub Check1_Click()
    Call Fill_Grid
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    vcDTP2.Value = vcDTP1.Value + 1
    vcDTP2.Visible = True
Else
    vcDTP2.Visible = False
End If
End Sub

Private Sub Check3_Click()

LFilterSaudaID = vbNullString

If Check3.Value = 1 Then
    Set LFilterRec = Nothing
    Set LFilterRec = New ADODB.Recordset
    mysql = "SELECT ACCID,  SAUDAID, SUM(CASE CONTYPE WHEN 'B' THEN QTY WHEN 'S' THEN QTY*-1 END  ) FROM CTR_D  WHERE COMPCODE =" & GCompCode & ""
    mysql = mysql & " AND SAUDAID IN (SELECT DISTINCT SAUDAID FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
    mysql = mysql & " AND CONDATE <='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    mysql = mysql & " GROUP BY ACCID, SAUDAID"
    mysql = mysql & " HAVING  ROUND(SUM(CASE CONTYPE WHEN 'B' THEN QTY WHEN 'S' THEN QTY*-1 END  ),2)<>0"
    LFilterRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LFilterRec.EOF Then
        Do While Not LFilterRec.EOF
            If LenB(LFilterSaudaID) > 0 Then LFilterSaudaID = LFilterSaudaID & ","
            LFilterSaudaID = LFilterSaudaID & LFilterRec!SAUDAID
            LFilterRec.MoveNext
        Loop
    End If
End If

    Set RecSauda = Nothing:    Set RecSauda = New ADODB.Recordset
    mysql = "SELECT ITEMID,SAUDAID,EXID,EXCODE,SAUDACODE,SAUDANAME,ITEMCODE,MATURITY FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " "
    mysql = mysql & " AND MATURITY>='" & Format(ldate, "YYYY/MM/DD") & "'"
    If LenB(LFilterSaudaID) > 0 Then mysql = mysql & " AND SAUDAID IN ( " & LFilterSaudaID & ")"
    If LenB(LSExCode) > 0 Then mysql = mysql & " AND EXCODE ='" & LSExCode & "'"
    mysql = mysql & " ORDER BY ITEMCODE,MATURITY"
    RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RecSauda.EOF Then
        Set DComboSauda.RowSource = RecSauda
        DComboSauda.ListField = "SAUDANAME"
        DComboSauda.BoundColumn = "SAUDACODE"
    Else
        MsgBox "No Contracts Fould ", vbInformation
        If Frame1.Enabled = True Then
            vcDTP1.SetFocus
        End If
        'Call CANCEL_REC
    End If

Call Fill_Grid


End Sub

Private Sub Command1_Click()
Frame8.Visible = False
End Sub

Private Sub DComboSauda_GotFocus()
    DComboSauda.Top = Val(DataGrid2.Top) + Val(DataGrid2.RowTop(DataGrid2.Row))
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboSauda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If LenB(DComboSauda.BoundText) = 0 Then
            MsgBox "Please Select Sauda.", vbCritical
        Else
            RECGRID!saudacode = DComboSauda.BoundText
            RECGRID!SAUDANAME = DComboSauda.text
            
            RecSauda.MoveFirst
            RecSauda.Find "SAUDACODE='" & DComboSauda.BoundText & "'", , adSearchForward
            If Not RecSauda.EOF Then
                RECGRID!ITEMCODE = RecSauda!ITEMCODE
                RECGRID!excode = RecSauda!excode
                RECGRID!SAUDAID = RecSauda!SAUDAID
            Else
                RECGRID!ITEMCODE = vbNullString
            End If
            Call DataGrid2_AfterColEdit(1)
            DComboSauda.Visible = False
            DataGrid2.Col = 3
            DataGrid2.SetFocus
        End If
    ElseIf KeyCode = 27 Then
        DComboSauda.Visible = False
    End If
End Sub

Private Sub DataGrid2_AfterColEdit(ByVal ColIndex As Integer)
    Dim TRec As ADODB.Recordset
    If ColIndex = 1 Then
        If IsDate(DataGrid2.Columns(0).text) Then    'RECGRID!CONTDATE
            If DateValue(DataGrid2.Columns(0).text) < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical:  Exit Sub
            If DateValue(DataGrid2.Columns(0).text) > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
        Else
            MsgBox "Please Enter Date.", vbCritical: Exit Sub
        End If
        
        If LenB(RECGRID!saudacode) = 0 Then MsgBox "Please Select Sauda ", vbCritical: Exit Sub
        mysql = "SELECT SAUDA FROM CTR_R WHERE COMPCODE=" & GCompCode & " AND CONDATE='" & Format(RECGRID!CONTDATE, "yyyy/MM/dd") & "' AND SAUDA='" & RECGRID!saudacode & "'"
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Closing Rate already exists for " & RECGRID!saudacode & ".", vbExclamation
            DataGrid2.Col = 1
            RECGRID!saudacode = vbNullString
            RECGRID!SAUDANAME = vbNullString
            RECGRID!SAUDAID = 0
        End If
        ''TO FIND WHETHER THE RECORD EXISTS IN RECORDSET
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        Set TRec = RECGRID.Clone
        TRec.Filter = "CONTDATE='" & RECGRID!CONTDATE & "'"
        TRec.MoveFirst
        TRec.Find "SAUDACODE='" & RECGRID!saudacode & "'", , adSearchForward
        If Not TRec.EOF Then
            ''THE ORIGINAL RECORD ALWAYS FINDS IN RECORDSET SO FIND NEXT TIME
            TRec.MoveNext
            TRec.Find "SAUDACODE='" & RECGRID!saudacode & "'", , adSearchForward
            If Not TRec.EOF Then
                MsgBox "Closing Rate already exists for " & RECGRID!saudacode & " in above records.", vbExclamation
                DataGrid2.Col = 1
                RECGRID!saudacode = vbNullString
                RECGRID!SAUDANAME = vbNullString
                RECGRID!SAUDAID = 0
            End If
        End If
        Set TRec = Nothing
    ElseIf DataGrid2.Columns(14).Visible And DataGrid2.Col = 6 Then
            DataGrid2.Col = 14
    End If
End Sub
Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = Val(1) Then
        RecSauda.MoveFirst
        RecSauda.Find "SAUDACODE='" & RECGRID!saudacode & "'", , adSearchForward
        If RecSauda.EOF Then
            RECGRID!saudacode = vbNullString
            RECGRID!SAUDANAME = vbNullString
            RECGRID!SAUDAID = 0
            DComboSauda.Visible = True
            DComboSauda.SetFocus
        Else
            RECGRID!saudacode = RecSauda!saudacode
            RECGRID!SAUDANAME = RecSauda!SAUDANAME
            RECGRID!SAUDAID = RecSauda!SAUDAID
        End If
    End If
End Sub
Private Sub DataGrid2_GotFocus()
    vcDTP1.Enabled = False
    DComboExchnage.Enabled = False
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LSDate As Date
    If KeyCode = 13 And (DataGrid2.Col <> 6 And DataGrid2.Col <> 14) Then
        If LenB(DataGrid2.Columns(0).text) = 0 Then
            MsgBox "Please enter date."
            DataGrid2.Columns(0).text = CStr(Date)
        ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid2.Col = 1 And LenB(RECGRID!saudacode) < 1) Then
            DComboSauda.Visible = True
            DComboSauda.SetFocus
        ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid2.Col = Val(1) And LenB(RECGRID!saudacode) > Val(1)) Then
          '  DataGrid2.Col = 2
          Sendkeys "{TAB}"
        Else
            If Not IsDate(DataGrid2.Columns(0).text) Then
                MsgBox "Please Enter Valid Date."
                DataGrid2.Columns(0).text = CStr(Date)
                DataGrid2.Col = 0
            Else
                If DateValue(DataGrid2.Columns(0).text) < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: DataGrid2.SetFocus: Exit Sub
                If DateValue(DataGrid2.Columns(0).text) > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: DataGrid2.SetFocus: Exit Sub
                Sendkeys "{TAB}"
            End If
        End If
    ElseIf KeyCode = 13 And DataGrid2.Col = 6 Then
        If DataGrid2.Columns(14).Visible Then
            Sendkeys "{TAB}"
            'DataGrid2.Col = 14
        Else
            LSDate = RECGRID!CONTDATE
            RECGRID.MoveNext
            If RECGRID.EOF Then
                RECGRID.AddNew
                RECGRID!CONTDATE = LSDate
                RECGRID!saudacode = vbNullString
                RECGRID!SAUDAID = 0
                RECGRID!itemid = 0
                RECGRID!EXID = 0
                RECGRID.Update
            End If
            DataGrid2.Col = 0
        End If
    ElseIf KeyCode = 13 And DataGrid2.Col = 14 Then
        LSDate = RECGRID!CONTDATE
        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            RECGRID!CONTDATE = LSDate
            RECGRID!saudacode = vbNullString
            RECGRID!SAUDAID = 0
            RECGRID!itemid = 0
            RECGRID!EXID = 0
            
            RECGRID.Update
        End If
        DataGrid2.Col = 0
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Fb_Press <> 0 Then
        If Me.ActiveControl.NAME = "vcDTP1" Then If KeyCode = 13 Then Sendkeys "{tab}"
        If Me.ActiveControl.NAME = "DComboExchnage" Then
            If KeyCode = 13 Then Sendkeys "{tab}"
        End If
    End If
End Sub

Private Sub Form_Load()
    LFilterSaudaID = vbNullString
    If Date <= GFinEnd Then
        vcDTP1.Value = Date
    Else
        vcDTP1.Value = DateValue(GFinEnd)
    End If
    vcDTP1.MaxDate = GFinEnd:    vcDTP1.MinDate = GFinBegin
    ldate = vcDTP1.Value:
    Call CANCEL_REC:
    Call Get_Selection(10)
    mysql = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE=" & GCompCode & "ORDER BY EXNAME"
    Set ExRec = Nothing
    Set ExRec = New ADODB.Recordset
    ExRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not ExRec.EOF Then
        Set DComboExchnage.RowSource = ExRec: DComboExchnage.BoundColumn = "EXCODE": DComboExchnage.ListField = "EXCODE"
    End If
    If ExRec.RecordCount = 1 Then
        DComboExchnage.BoundText = ExRec!excode
        DComboExchnage.Enabled = False
    End If
    Call colhide
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "CONTDATE", adDate, , adFldIsNullable
    'RECGRID.Fields.Append "CONTDATE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SAUDACODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SAUDANAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "CLOSING", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "OPENING", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LOW", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "HIGH", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "CLRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "DataImport", adBoolean, , adFldIsNullable
    RECGRID.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
    RECGRID.Fields.Append "EXCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "SAUDAID", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "ITEMID", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "EXID", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "PRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "CONTIME", adVarChar, 50, adFldIsNullable
    
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
    RECGRID.AddNew
    RECGRID!DATAIMPORT = False
    RECGRID.Update
    Set DataGrid2.DataSource = RECGRID
    DataGrid2.ReBind
    DataGrid2.Refresh
End Sub

Private Sub TxtFilterSauda_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtFilterSauda_Validate(Cancel As Boolean)
    If LenB(TxtFilterSauda.text) <> 0 Then
        Call Fill_Grid
       End If
End Sub
Private Sub vcDTP1_Validate(Cancel As Boolean)
    If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    ldate = vcDTP1.Value
    Set RecSauda = Nothing:    Set RecSauda = New ADODB.Recordset
    mysql = "SELECT I.EXID,I.ITEMID,I.EXCHANGECODE,S.SAUDAID,S.SAUDACODE, SAUDANAME,I.ITEMCODE,MATURITY FROM SAUDAMAST AS S, ITEMMAST AS I "
    mysql = mysql & "  WHERE S.COMPCODE=" & GCompCode & " AND  S.ITEMID =I.ITEMID "
    mysql = mysql & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY I.ITEMCODE,MATURITY "
    RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RecSauda.EOF Then
        Set DComboSauda.RowSource = RecSauda
        DComboSauda.ListField = "SAUDANAME"
        DComboSauda.BoundColumn = "SAUDACODE"
    Else
        MsgBox "Please Create Sauda.", vbInformation
        If Frame1.Enabled = True Then
            vcDTP1.SetFocus
        End If
        'Call CANCEL_REC
    End If
End Sub
Sub DELETE_REC()
    If RecSauda.RecordCount > 0 Then
        Fb_Press = 3
        Call Get_Selection(3)
        DataGrid2.Columns(0).Locked = True
        vcDTP1.Enabled = True
        DComboExchnage.Enabled = True
        Label1.Visible = True: vcDTP1.Visible = True: Frame1.Enabled = True: vcDTP1.SetFocus
    Else
        MsgBox "Please SELECT SAUDA.", vbCritical
        Call CANCEL_REC
    End If
End Sub

Private Sub DComboExchnage_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub DComboExchnage_Validate(Cancel As Boolean)
    LSExCode = DComboExchnage.BoundText
    ldate = DateValue(vcDTP1.Value)
    Set RecSauda = Nothing:    Set RecSauda = New ADODB.Recordset
    mysql = "SELECT ITEMID,SAUDAID,EXID,EXCODE,SAUDACODE,SAUDANAME,ITEMCODE,MATURITY FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " "
    mysql = mysql & " AND MATURITY>='" & Format(ldate, "YYYY/MM/DD") & "'"
    If LenB(LSExCode) > 0 Then mysql = mysql & " AND EXCODE ='" & LSExCode & "'"
    mysql = mysql & " ORDER BY ITEMCODE,MATURITY"
    RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RecSauda.EOF Then
        Set DComboSauda.RowSource = RecSauda
        DComboSauda.ListField = "SAUDANAME"
        DComboSauda.BoundColumn = "SAUDACODE"
    Else
        MsgBox "Please Create Sauda.", vbInformation
        If Frame1.Enabled = True Then
            vcDTP1.SetFocus
        End If
        'Call CANCEL_REC
    End If
    LSSaudas = vbNullString
    Fill_Grid
End Sub
Sub colhide()
    Dim TRec As ADODB.Recordset
    DataGrid2.Columns(14).Visible = True
    mysql = "select count(BILLBY) as 'BILLBY' from ACCT_EX where BILLBY='P'"
    Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        If TRec!billby = 0 Then
            DataGrid2.Columns(14).Visible = False
        End If
    End If
    DataGrid2.Columns(14).Visible = False
End Sub
 Sub Fill_Grid()
    Dim TRec As ADODB.Recordset
    Dim LExIDS As String
    Dim LCONDT_TIME As String
    Call RecSet
    
    Call colhide
        
    If SYSTEMLOCK(DateValue(vcDTP1.Value)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
    Else
        If Fb_Press <> 1 Then
            'mysql = "SELECT convert(varchar,day(A.condate)) +'/'+ convert(varchar,month(A.condate)) +'/'+ convert(varchar,year(A.condate)) + ' ' + isnull(A.contime,'') as 'CONDATE'"
            
            mysql = "SELECT CONDATE "
            mysql = mysql & ",S.ITEMID,S.EXID ,S.SAUDAID,A.SAUDA,A.CONTIME,A.CLOSERATE,A.OPRATE,A.LOWRATE,A.HGRATE,A.CLRATE,S.ITEMCODE,S.EXCODE, ISNULL(A.PRATE,0) AS 'PRATE' "
            mysql = mysql & " FROM CTR_R AS A,SAUDAMAST AS S "
            mysql = mysql & " WHERE A.COMPCODE=" & GCompCode & " AND A.COMPCODE =S.COMPCODE AND A.SAUDAID=S.SAUDAID AND S.MATURITY>='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND A.CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
            If LenB(LSExCode) > 0 Then mysql = mysql & " AND S.EXCODE='" & LSExCode & "'"
            If LenB(TxtFilterSauda.text) <> 0 Then mysql = mysql & " AND UPPER(S.SAUDACODE) LIKE '" & Trim(UCase(TxtFilterSauda.text)) & "%'"
            mysql = mysql & " ORDER BY A.EXCODE, A.ITEMCODE,S.MATURITY"
            Set TRec = Nothing
            Set TRec = New ADODB.Recordset
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                RECGRID.Delete
                DoEvents
                Do While Not TRec.EOF
                    RECGRID.AddNew
                    
                    RECGRID!CONTDATE = Format(TRec!Condate, "dd/MM/yyyy"): RECGRID!saudacode = TRec!Sauda
                    RECGRID!SAUDANAME = TRec!Sauda:                     RECGRID!CLOSING = TRec!CLOSERATE
                    RECGRID!OPENING = TRec!OPRATE:                      RECGRID!LOW = TRec!LOWRATE
                    RECGRID!HIGH = TRec!HGRATE:                         RECGRID!ClRate = TRec!ClRate
                    RECGRID!ITEMCODE = TRec!ITEMCODE:                   RECGRID!excode = TRec!excode
                    RECGRID!SAUDAID = TRec!SAUDAID:                     RECGRID!itemid = TRec!itemid
                    RECGRID!EXID = TRec!EXID:                           RECGRID!PRATE = TRec!PRATE
                    RECGRID!contime = TRec!contime
                    
                    RECGRID.Update
                    If Fb_Press = 3 Then
                        If LenB(LSSaudas) <> 0 Then LSSaudas = LSSaudas & ", "
                        LSSaudas = LSSaudas & TRec!SAUDAID & ""
                    End If
                    TRec.MoveNext
                Loop
                DoEvents
                TRec.MoveFirst: RECGRID.MoveFirst
                Call DataGrid2_AfterColEdit(0)
                DataGrid2.Col = 0: DataGrid2.Row = 0: 'DataGrid2.SetFocus
                If Fb_Press = 3 Then
                    If MsgBox("Confirm DELETE?", vbYesNo) = vbYes Then
                        Cnn.BeginTrans
                        CNNERR = True
                        mysql = "DELETE FROM CTR_R WHERE COMPCODE=" & GCompCode & " AND   SAUDAID  IN (" & LSSaudas & ") AND CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
                        Cnn.Execute mysql
                        If LenB(LSSaudas) = 0 Then
                            Cnn.CommitTrans
                            CNNERR = False
                        Else
                            LExIDS = vbNullString
                            If LenB(LSExCode) > 0 Then
                                LExIDS = CStr(Get_ExID(LSExCode))
                            End If
                            
                            'Call Delete_Inv_D(vbNullString, LExIDS, LSSaudas, vcDTP1.Value)
                            If BILL_GENERATION(CDate(vcDTP1.Value), CDate(GFinEnd), LSSaudas, vbNullString, LExIDS) Then
                                Cnn.CommitTrans
                                CNNERR = False
                            Else
                                Cnn.RollbackTrans
                                CNNERR = False
                            End If
                            LSExCode = vbNullString
                           ' Call Chk_Billing
                        End If
                    End If
                    Call CANCEL_REC
                End If
            End If
            Set TRec = Nothing
        Else
            RECGRID.Delete
            If RecSauda.RecordCount > 0 Then RecSauda.MoveFirst
                        
            Do While Not RecSauda.EOF
                If Check1.Value = 0 Then
                    RECGRID.AddNew
                    RECGRID!CONTDATE = Format(vcDTP1.Value, "dd/MM/yyyy"):                   RECGRID!saudacode = RecSauda!saudacode
                    RECGRID!SAUDANAME = RecSauda!SAUDANAME:             RECGRID!ITEMCODE = RecSauda!ITEMCODE
                    RECGRID!excode = RecSauda!excode:                   RECGRID!SAUDAID = RecSauda!SAUDAID
                    RECGRID!itemid = RecSauda!itemid:                   RECGRID!EXID = RecSauda!EXID

                                    
                    mysql = "SELECT CONTIME,CLOSERATE,OPRATE,LOWRATE,HGRATE,CLRATE,isnull(CTR_R.PRATE,0) AS 'PRATE' FROM CTR_R WHERE CTR_R.COMPCODE=" & GCompCode & " "
                    mysql = mysql & " AND  CTR_R.SAUDA='" & RecSauda!saudacode & "'"
                    mysql = mysql & " AND CTR_R.CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' "
                    Set TRec = Nothing
                    Set TRec = New ADODB.Recordset
                    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                    If TRec.EOF Then
                        RECGRID!CLOSING = 0:                RECGRID!OPENING = 0
                        RECGRID!LOW = 0:                    RECGRID!HIGH = 0
                        RECGRID!ClRate = 0:                 RECGRID!PRATE = 0
                        RECGRID!contime = ""
                    Else
                        RECGRID!CLOSING = IIf(IsNull(TRec!CLOSERATE), 0, TRec!CLOSERATE)
                        RECGRID!OPENING = IIf(IsNull(TRec!OPRATE), 0, TRec!OPRATE)
                        RECGRID!LOW = IIf(IsNull(TRec!LOWRATE), 0, TRec!LOWRATE)
                        RECGRID!HIGH = IIf(IsNull(TRec!HGRATE), 0, TRec!HGRATE)
                        RECGRID!ClRate = IIf(IsNull(TRec!ClRate), 0, TRec!ClRate)
                        RECGRID!PRATE = TRec!PRATE
                        RECGRID!contime = TRec!contime
                    End If
                    RECGRID.Update
                Else
                    mysql = "SELECT CONTIME,CLOSERATE,OPRATE,LOWRATE,HGRATE,CLRATE,isnull(CTR_R.PRATE,0) AS 'PRATE' FROM CTR_R WHERE CTR_R.COMPCODE=" & GCompCode & " "
                    mysql = mysql & " AND  CTR_R.SAUDA='" & RecSauda!saudacode & "'"
                    mysql = mysql & " AND CTR_R.CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' "
                    mysql = mysql & " AND ISNULL(CLOSERATE,0)>0  "
                    
                    
                    
                    
                    Set TRec = Nothing
                    Set TRec = New ADODB.Recordset
                    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                    If TRec.EOF Then
                        RECGRID.AddNew
                        RECGRID!CONTDATE = Format(vcDTP1.Value, "dd/MM/yyyy"):                        RECGRID!saudacode = RecSauda!saudacode
                        RECGRID!SAUDANAME = RecSauda!SAUDANAME:                 RECGRID!ITEMCODE = RecSauda!ITEMCODE
                        RECGRID!excode = RecSauda!excode:                       RECGRID!SAUDAID = RecSauda!SAUDAID
                        RECGRID!itemid = RecSauda!itemid:                       RECGRID!EXID = RecSauda!EXID
                        RECGRID!CLOSING = 0:                                    RECGRID!OPENING = 0
                        RECGRID!LOW = 0:                                        RECGRID!HIGH = 0
                        RECGRID!ClRate = 0:                                     RECGRID!PRATE = 0 '>>>RECGRID!PRATE = TRec!PRATE
                        RECGRID!contime = "" '>>>TRec!contime
                        
                        RECGRID.Update
                    End If
                End If
                RecSauda.MoveNext
            Loop
            If RECGRID.RecordCount > 0 Then
                RECGRID.MoveLast
                Call DataGrid2_AfterColEdit(0)
                DataGrid2.Refresh
                DataGrid2.Col = 0: DataGrid2.SetFocus
            End If
        End If
    End If
 End Sub
