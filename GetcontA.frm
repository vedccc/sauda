VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form GETContA 
   BackColor       =   &H80000000&
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Unused"
      Height          =   855
      Left            =   4080
      TabIndex        =   19
      Top             =   7560
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   3840
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   360
         Width           =   255
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   "DataCombo2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda Code"
         Height          =   225
         Index           =   2
         Left            =   720
         TabIndex        =   25
         Top             =   195
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item  Name"
         Height          =   225
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Rate"
         Height          =   225
         Index           =   4
         Left            =   2760
         TabIndex        =   22
         Top             =   435
         Width           =   1020
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GetcontA.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GetcontA.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   7680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   200
      TabIndex        =   5
      Top             =   960
      Width           =   11465
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5040
         TabIndex        =   36
         Top             =   -75
         Visible         =   0   'False
         Width           =   6495
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   345
            Left            =   720
            TabIndex        =   2
            Top             =   120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   609
            _Version        =   393216
            ForeColor       =   16711680
            Text            =   "DataCombo4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Height          =   345
            Left            =   3840
            TabIndex        =   3
            Top             =   120
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            _Version        =   393216
            ForeColor       =   12582912
            Text            =   "DataCombo5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Party"
            Height          =   225
            Index           =   15
            Left            =   3240
            TabIndex        =   38
            Top             =   173
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
            Height          =   225
            Index           =   14
            Left            =   120
            TabIndex        =   37
            Top             =   173
            Width           =   495
         End
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Text10"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Text9"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Text8"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Text7"
         Top             =   5280
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   12582912
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   5760
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   5760
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   360
         Left            =   3120
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   12582912
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7858
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "SRNO"
            Caption         =   "Trade No."
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
            DataField       =   "SCODE"
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
            DataField       =   "CONTYPE"
            Caption         =   "B/S"
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
            DataField       =   "BCODE"
            Caption         =   "Party"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "BNAME"
            Caption         =   "Party Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "BQNTY"
            Caption         =   "Qnty."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "BRATE"
            Caption         =   "Rate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
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
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2459.906
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   4169.764
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF0000&
         Height          =   345
         ItemData        =   "GetcontA.frx":08A4
         Left            =   3600
         List            =   "GetcontA.frx":08AE
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   0
         Width           =   1335
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   360
         Left            =   1440
         TabIndex        =   0
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37860.8625462963
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Height          =   135
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   360
         Width           =   11295
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11400
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda Total"
         Height          =   225
         Index           =   13
         Left            =   0
         TabIndex        =   35
         Top             =   5400
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         Height          =   225
         Index           =   12
         Left            =   0
         TabIndex        =   34
         Top             =   5835
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bought Qty"
         Height          =   225
         Index           =   11
         Left            =   4920
         TabIndex        =   33
         Top             =   5355
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sold Qty"
         Height          =   225
         Index           =   10
         Left            =   7200
         TabIndex        =   32
         Top             =   5355
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Difference Qty"
         Height          =   225
         Index           =   9
         Left            =   9120
         TabIndex        =   31
         Top             =   5355
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shree Balance"
         Height          =   225
         Index           =   8
         Left            =   1440
         TabIndex        =   30
         Top             =   5355
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   1
         Left            =   1080
         Picture         =   "GetcontA.frx":08C5
         Stretch         =   -1  'True
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   0
         Left            =   720
         Picture         =   "GetcontA.frx":0BCF
         Stretch         =   -1  'True
         Top             =   510
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shree Balance"
         Height          =   225
         Index           =   7
         Left            =   1440
         TabIndex        =   18
         Top             =   5835
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Difference Qty"
         Height          =   225
         Index           =   6
         Left            =   9120
         TabIndex        =   15
         Top             =   5835
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sold Qty"
         Height          =   225
         Index           =   5
         Left            =   7200
         TabIndex        =   14
         Top             =   5835
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bought Qty"
         Height          =   225
         Index           =   0
         Left            =   4920
         TabIndex        =   13
         Top             =   5835
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   225
         Index           =   18
         Left            =   3000
         TabIndex        =   7
         Top             =   75
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract Date"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   75
         Width           =   1110
      End
   End
   Begin VB.TextBox Text11 
      Height          =   330
      Left            =   5640
      TabIndex        =   39
      Text            =   "Text11"
      Top             =   2280
      Width           =   270
   End
   Begin VB.Label LblSave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save (F4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4200
      TabIndex        =   45
      Top             =   480
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label LblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   44
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label LblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close (F6)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10680
      TabIndex        =   43
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label LblDelete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete (F3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6960
      TabIndex        =   42
      Top             =   480
      Width           =   1170
   End
   Begin VB.Label LblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit (F2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   41
      Top             =   480
      Width           =   885
   End
   Begin VB.Label LblNew 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&New (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4200
      TabIndex        =   40
      Top             =   480
      Width           =   930
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      BorderWidth     =   12
      Height          =   6300
      Left            =   80
      Top             =   840
      Width           =   11685
   End
   Begin VB.Image Image1 
      Height          =   795
      Index           =   2
      Left            =   0
      Picture         =   "GetcontA.frx":0ED9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "GETContA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FlagFilter As Boolean
Dim FilterCol As Byte
Dim GridSrNo As Integer
Dim FLOWDIR As Byte
Dim VCHNO As String
Public fb_press As Byte
Dim REC As ADODB.Recordset
Dim RECGRID As ADODB.Recordset
Dim REC_SAUDA As ADODB.Recordset
Dim REC_ACCOUNT As ADODB.Recordset
Sub ADD_REC()
    Frame1.Enabled = True
    Call Get_Selection(1)
    vcDTP1.SetFocus
End Sub
Sub SAVE_REC()
    On Error GoTo ERR1

    If vcDTP1.Value < MFIN_BEG Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > MFIN_END Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    
    CNN.BeginTrans: CNNERR = True
    
    RECGRID.Sort = "SRNO"
    If RECGRID.RecordCount > 0 Then
        If fb_press = 1 Then
            VCHNO = VOUCHER_NUMBER("CONT", FIN_YEAR(vcDTP1.Value))
            If Not Adodc1.Recordset.EOF Then
                Set REC = Nothing
                Set REC = New ADODB.Recordset
                REC.Open "SELECT MAX(CONSNO) FROM CTR_M WHERE compcode=" & MC_CODE & "", CNN, adOpenForwardOnly, adLockReadOnly
                CONSNO = REC.Fields(0)
                Set REC = Nothing
            End If
    
            CONSNO = Val(CONSNO & "") + Val(1)
    
        Else
            CONSNO = Adodc1.Recordset!CONSNO
            VCHNO = Adodc1.Recordset!VOU_NO & ""
            If Len(Trim(VCHNO)) < Val(1) Then
                VCHNO = VOUCHER_NUMBER("CONT", FIN_YEAR(vcDTP1.Value))
            End If
        End If
    
        Call DELETE_VOUCHER(VCHNO)
    
        If fb_press <> 1 Then
            If Len(DataCombo4.Text) < 1 And Len(DataCombo5.Text) < 1 Then
                MYSQL = "DELETE FROM CTR_D WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
    
            ElseIf Len(DataCombo4.Text) > 1 And Len(DataCombo5.Text) < 1 Then
                MYSQL = "DELETE FROM CTR_D WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND SAUDA='" & DataCombo4.BoundText & "'"
    
            Else
                MYSQL = "DELETE FROM CTR_D WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND SAUDA='" & DataCombo4.BoundText & "' AND PARTY='" & DataCombo5.BoundText & "'"
    
            End If
    
            CNN.Execute MYSQL
            
            CNN.Execute "DELETE FROM CTR_R WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
            CNN.Execute "DELETE FROM CTR_M WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
        End If
    
        If SWTYPE = "SQL" Then
            MYSQL = "INSERT INTO CTR_M(CompCode,CONSNO, CONDATE, SAUDA, ITEMCODE, CLOSERATE, VOU_NO, PATTAN) VALUES(" & MC_CODE & "," & CONSNO & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "', '" & Text2.Text & "', ' ', " & Val(Text3.Text) & ", '" & VCHNO & "', '" & Mid(Combo1.Text, 1, 1) & "')"
        Else
            MYSQL = "INSERT INTO CTR_M(CompCode,CONSNO, CONDATE, SAUDA, ITEMCODE, CLOSERATE, VOU_NO, PATTAN) VALUES(" & MC_CODE & "," & CONSNO & ", DATEVALUE('" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'), '" & Text2.Text & "', ' ', " & Val(Text3.Text) & ", '" & VCHNO & "', '" & Mid(Combo1.Text, 1, 1) & "')"
        End If
        CNN.Execute MYSQL
    
        Dim BOOLAC As String * 1
        Dim RC As ADODB.Recordset
    
        RECGRID.MoveFirst
        Do While Not RECGRID.EOF
            ''RECORDSET RC IS CHECKING WHETHER THE PARTY IS PERSONNEL OR NOT
            BOOLAC = "N"
    '        MYSQL = "SELECT PERSONNELAC FROM PITBROK WHERE compcode=" & MC_CODE & " AND AC_CODE='" & RECGRID!BCODE & "'"
    '        Set RC = New ADODB.Recordset
    '        RC.Open MYSQL, CNN, adOpenKeyset, adLockReadOnly
    '        If RC!PERSONNELAC = "Y" Then
    '            MYSQL = "SELECT PERSONNELAC FROM PITBROK WHERE compcode=" & MC_CODE & " AND AC_CODE='" & RECGRID!SCODE & "'"
    '            Set RC = New ADODB.Recordset
    '            RC.Open MYSQL, CNN, adOpenKeyset, adLockReadOnly
    '
    '            If RC!PERSONNELAC = "Y" Then BOOLAC = "Y"
    '
    '        End If
    
            If Len(RECGRID!BNAME & "") > Val(0) Then   ''WHEN BUYER AND SELLER BOTH ARE THERE
                If RECGRID!BQNTY > Val(0) And RECGRID!BRate > Val(0) Then                   ''QNTY AND RATE REQUIRED
                    If SWTYPE = "SQL" Then
                        MYSQL = "INSERT INTO CTR_D (CompCode,CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT) VALUES(" & MC_CODE & "," & Val(CONSNO) & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'," & Val(RECGRID.AbsolutePosition) & ",'" & RECGRID!scode & "', '', '" & RECGRID!BCODE & "', " & Val(RECGRID!BQNTY) & "," & Val(RECGRID!BRate) & ",'" & RECGRID!CONTYPE & "', '" & BOOLAC & "')"
                    Else
                        MYSQL = "INSERT INTO CTR_D (CompCode,CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT) VALUES(" & MC_CODE & "," & Val(CONSNO) & ", DATEVALUE('" & Format(vcDTP1.Value, "yyyy/MM/dd") & "')," & Val(RECGRID.AbsolutePosition) & ",'" & RECGRID!scode & "', '', '" & RECGRID!BCODE & "', " & Val(RECGRID!BQNTY) & "," & Val(RECGRID!BRate) & ",'" & RECGRID!CONTYPE & "', '" & BOOLAC & "')"
                    End If
                    CNN.Execute MYSQL
                End If
            End If
            RECGRID.MoveNext
        Loop
    
        If Val(Text3.Text) > Val(0) Then
            If SWTYPE = "SQL" Then
                MYSQL = "INSERT INTO CTR_R(CompCode,CONSNO, SAUDA, CONDATE, PATTAN, OPRATE, HGRATE, LOWRATE, CLOSERATE) VALUES(" & MC_CODE & "," & CONSNO & ",'" & DataCombo1.BoundText & "', '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & Mid(Combo1.Text, 1, 1) & "', 0, 0, 0, " & Val(Text3.Text) & ")"
            Else
                MYSQL = "INSERT INTO CTR_R(CompCode,CONSNO, SAUDA, CONDATE, PATTAN, OPRATE, HGRATE, LOWRATE, CLOSERATE) VALUES(" & MC_CODE & "," & CONSNO & ",'" & DataCombo1.BoundText & "', DATEVALUE('" & vcDTP1.Value & "'),'" & Mid(Combo1.Text, 1, 1) & "', 0, 0, 0, " & Val(Text3.Text) & ")"
            End If
            CNN.Execute MYSQL
        End If
        
        Call GENERATE_VOUCHER
    
        CNN.CommitTrans
        CNNERR = False
        Adodc1.Refresh
    
     ''BILL GENERATION
     ''TO FIND FROMDATE
        MYSQL = "SELECT MAX(SETDATE) FROM SETTLE WHERE compcode=" & MC_CODE & " AND SeTDATE < '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
        Set REC = Nothing: Set REC = New ADODB.Recordset: REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
        If Not REC.EOF Then MFROMDATE = REC.Fields(0) + Val(1): MFROMDATE = Format(vcDTP1.Value, "yyyy/MM/dd")
                    
      ''TO FIND TODATE
        MYSQL = "SELECT MAX(CONDATE) FROM CTR_M WHERE compcode=" & MC_CODE & " AND SAUDA='" & Text2.Text & "'"
        Set REC = Nothing: Set REC = New ADODB.Recordset: REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
        If Len(Trim(REC.Fields(0) & "")) > Val(10) Then
            MTODATE = REC.Fields(0)
        Else
            MTODATE = MFROMDATE
        End If
    
        CNN.BeginTrans: CNNERR = False
        If BILL_GENERATION(CDate(MFROMDATE), CDate(MTODATE), "'" & Text2.Text & "'") Then
            CNN.CommitTrans: CNNERR = False
        Else
            CNN.RollbackTrans: CNNERR = False
        End If
        
    End If
    Call CANCEL_REC
    'Call lblcancel_Click
    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
    If CNNERR = True Then CNN.RollbackTrans: CNNERR = False
End Sub
Sub CANCEL_REC()
    GridSrNo = 1
    Call RECSET
    FlagFilter = False
    ''modify & then save then again modify giving problem
    'Unload Me
    'GETCont.Show

    Set DataGrid1.DataSource = RECGRID
    DataGrid1.Refresh

    DataCombo1.Visible = False
    DataCombo3.Visible = False

    Frame4.Visible = False

    Call ClearFormFn(GETContA)
    Call Get_Selection(10)
    Frame1.Enabled = False
End Sub
Sub MODIFY_REC()
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    
    If SWTYPE = "SQL" Then
        MYSQL = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND PATTAN='" & Mid(Combo1.Text, 1, 1) & "'"
    Else
        MYSQL = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND CONDATE=DATEVALUE('" & Format(vcDTP1.Value, "yyyy/MM/dd") & "') AND PATTAN='" & Mid(Combo1.Text, 1, 1) & "'"
    End If

    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    
    If REC.EOF Then
        If fb_press = 2 Then
            MsgBox "Transaction does not exists for the selected creteria?", vbExclamation
            Call CANCEL_REC
        End If
        Exit Sub
    End If
    
    'Adodc1.RecordSource = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(REC!CONSNO) & ""
    'Adodc1.Refresh
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find "CONSNO=" & Val(REC!CONSNO & "") & "", , adSearchForward
    
    If fb_press = 1 Then fb_press = 2

    With Adodc1.Recordset
        vcDTP1.Value = !CONDATE
        DataCombo1.BoundText = !Sauda
        Text2.Text = !Sauda
        DataCombo2.BoundText = !ItemCode
        If !PATTAN = "C" Then
            Combo1.ListIndex = Val(0)
        Else
            Combo1.ListIndex = Val(1)
        End If
        Text3.Text = Format(!CLOSERATE, "0.00")
    End With

    If Len(DataCombo4.Text) < 1 And Len(DataCombo5.Text) < 1 Then
        MYSQL = "SELECT CTR_D.*, A.NAME AS NAME, SAUDA.SAUDACODE, SAUDA.SAUDANAME FROM CTR_D, ACCOUNTD AS A, SAUDAMAST AS SAUDA WHERE CTR_D.compcode=" & MC_CODE & " AND CTR_D.PARTY=A.AC_CODE AND CTR_D.SAUDA=SAUDA.SAUDACODE AND CTR_D.CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " ORDER BY CONNO, CONTYPE"

    ElseIf Len(DataCombo4.Text) > 1 And Len(DataCombo5.Text) < 1 Then
        MYSQL = "SELECT CTR_D.*, A.NAME AS NAME, SAUDA.SAUDACODE, SAUDA.SAUDANAME FROM CTR_D, ACCOUNTD AS A, SAUDAMAST AS SAUDA WHERE CTR_D.compcode=" & MC_CODE & " AND CTR_D.PARTY=A.AC_CODE AND CTR_D.SAUDA=SAUDA.SAUDACODE AND CTR_D.CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND CTR_D.SAUDA='" & DataCombo4.BoundText & "' ORDER BY CONNO, CONTYPE"

    ElseIf Len(DataCombo4.Text) < 1 And Len(DataCombo5.Text) > 1 Then
        MYSQL = "SELECT CTR_D.*, A.NAME AS NAME, SAUDA.SAUDACODE, SAUDA.SAUDANAME FROM CTR_D, ACCOUNTD AS A, SAUDAMAST AS SAUDA WHERE CTR_D.compcode=" & MC_CODE & " AND CTR_D.PARTY=A.AC_CODE AND CTR_D.SAUDA=SAUDA.SAUDACODE AND CTR_D.CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND CTR_D.PARTY='" & DataCombo5.BoundText & "' ORDER BY CONNO, CONTYPE"

    Else
        MYSQL = "SELECT CTR_D.*, A.NAME AS NAME, SAUDA.SAUDACODE, SAUDA.SAUDANAME FROM CTR_D, ACCOUNTD AS A, SAUDAMAST AS SAUDA WHERE CTR_D.compcode=" & MC_CODE & " AND CTR_D.PARTY=A.AC_CODE AND CTR_D.SAUDA=SAUDA.SAUDACODE AND CTR_D.CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND CTR_D.SAUDA='" & DataCombo4.BoundText & "' AND CTR_D.PARTY='" & DataCombo5.BoundText & "' ORDER BY CONNO, CONTYPE"

    End If

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly

    Call RECSET
    RECGRID.Delete

    Do While Not REC.EOF
        RECGRID.AddNew
        RECGRID!SRNO = RECGRID.AbsolutePosition
        RECGRID!BCODE = REC!PARTY & ""
        RECGRID!BNAME = REC!Name
        RECGRID!BQNTY = REC!QTY
        RECGRID!BRate = REC!Rate
        RECGRID!CONTYPE = REC!CONTYPE
        RECGRID!scode = REC!SAUDACODE
        RECGRID!SNAME = REC!SAUDANAME

        RECGRID.Update

        REC.MoveNext
    Loop

    Set DataGrid1.DataSource = RECGRID

'''''    Call DataGrid1_AfterColEdit(0)

    If fb_press = 3 Then
        If MsgBox("You are about to delete this record. Confirm Delete?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            On Error GoTo ERR1
            CNN.BeginTrans
            CNNERR = True

            If Len(DataCombo4.Text) < 1 And Len(DataCombo5.Text) < 1 Then
                MYSQL = "DELETE FROM CTR_D WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""

            ElseIf Len(DataCombo4.Text) > 1 And Len(DataCombo5.Text) < 1 Then
                MYSQL = "DELETE FROM CTR_D WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND SAUDA='" & DataCombo4.BoundText & "'"

            Else
                MYSQL = "DELETE FROM CTR_D WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND SAUDA='" & DataCombo4.BoundText & "' AND PARTY='" & DataCombo5.BoundText & "'"

            End If
            CNN.Execute MYSQL
        
            Set REC = Nothing
            Set REC = New ADODB.Recordset
            REC.Open "SELECT * FROM CTR_D WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & "", CNN, adOpenForwardOnly, adLockReadOnly
            If REC.EOF Then
                CNN.Execute "DELETE FROM CTR_R WHERE compcode=" & MC_CODE & "VCONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
                CNN.Execute "DELETE FROM CTR_M WHERE compcode=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
                Call DELETE_VOUCHER(Adodc1.Recordset!VOU_NO & "")
'            Else
'                Call GENERATE_VOUCHER
            End If
            Adodc1.Refresh
            ''REGENERATING SETTLEMENTS
              MFROMDATE = Format(vcDTP1.Value, "yyyy/MM/dd")

            ''TO FIND TODATE
              MYSQL = "SELECT MAX(CONDATE) FROM CTR_M WHERE compcode=" & MC_CODE & " AND SAUDA='" & Text2.Text & "'"
              Set REC = Nothing
              Set REC = New ADODB.Recordset
              REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
              If Len(Trim(REC.Fields(0) & "")) > Val(10) Then
                  MTODATE = REC.Fields(0)
              Else
                  MTODATE = MFROMDATE
              End If
              If BILL_GENERATION(CDate(MFROMDATE), CDate(MTODATE), "'" & Text2.Text & "'") Then
                CNN.CommitTrans: CNNERR = False
              Else
                CNN.RollbackTrans: CNNERR = False
              End If
            ''REGENERATING UPTO HERE

ERR1:
            If Err.Number <> 0 Then
                MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
            End If
            If CNNERR = True Then CNN.RollbackTrans: CNNERR = False
            Call CANCEL_REC
        End If
    End If
End Sub
Private Sub Combo1_GotFocus()
    If FLOWDIR = 1 Then
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND SAUDA='" & DataCombo1.BoundText & "'", CNN, adOpenForwardOnly, adLockReadOnly
        If REC.EOF Then SendKeys "%{DOWN}"
    Else
'       Text2.SetFocus
    End If
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    MYSQL = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND PATTAN='" & Left(Combo1.Text, 1) & "' AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    If Not REC.EOF Then
        If vcDTP1.Value < DateValue(REC!CONDATE) Then
            MsgBox "Opening for this SAUDA has been already entered on " & Format(REC!CONDATE, "yyyy/MM/dd"), vbExclamation, "Warning"
            vcDTP1.Value = Date
            Cancel = True
            Exit Sub
        End If
    Else
        If Not REC.EOF Then
            If REC!CONDATE > vcDTP1.Value Then
                MsgBox "Opening for this SAUDA has been already entered on " & Format(REC!CONDATE, "yyyy/MM/dd"), vbExclamation, "Warning"
                vcDTP1.Value = Date
                Exit Sub
            End If
        End If
    End If

    If Combo1.ListIndex = Val(0) Then   ''CONTRACT
        
    Else                                ''OPENING
'        DataGrid1.Columns(7).Locked = False
    End If

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    If SWTYPE = "SQL" Then
        MYSQL = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND PATTAN='" & Mid(Combo1.Text, 1, 1) & "'"
    Else
        MYSQL = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND CONDATE=DATEVALUE('" & Format(vcDTP1.Value, "yyyy/MM/dd") & "') AND PATTAN='" & Mid(Combo1.Text, 1, 1) & "'"
    End If

    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly

    If Not REC.EOF Then
        Frame4.Visible = True
        
        Text11.SetFocus
    End If
End Sub
Private Sub DataCombo1_GotFocus()
    DataCombo1.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        DataGrid1.SetFocus
        DataCombo1.Visible = False

    ElseIf KeyCode = 13 Then
        If Len(Trim(DataCombo1.BoundText)) > 0 Then
            RECGRID!scode = DataCombo1.BoundText
            RECGRID!scode = DataCombo1.BoundText
            DataGrid1.SetFocus
            DataCombo1.Visible = False
        Else
            MsgBox "Please select sauda."
        End If
    End If
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    Set REC_SAUDA = Nothing
    Set REC_SAUDA = New ADODB.Recordset
    REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE compcode=" & MC_CODE & " AND SAUDACODE='" & DataCombo1.BoundText & "'", CNN, adOpenForwardOnly, adLockReadOnly
    If REC_SAUDA.EOF Then
        Cancel = True
    Else
        Text2.Text = REC_SAUDA!SAUDACODE
        DataCombo1.BoundText = Text2.Text
        DataCombo2.BoundText = REC_SAUDA!ItemCode
        Combo1.SetFocus
    End If
End Sub
Private Sub DataCombo2_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo3_GotFocus()
    DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        RECGRID!BCODE = DataCombo3.BoundText
        RECGRID!BNAME = DataCombo3.Text
        DataGrid1.Col = 4

        DataGrid1.SetFocus
        DataCombo3.Visible = False

    ElseIf KeyCode = 27 Then
        DataGrid1.SetFocus
        DataCombo3.Visible = False
    End If
End Sub
Private Sub DataCombo3_Validate(Cancel As Boolean)
    If DataCombo3.Visible = True Then Cancel = True
End Sub
Private Sub DataCombo4_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub DataCombo5_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub DataCombo5_Validate(Cancel As Boolean)
    Call MODIFY_REC
End Sub
Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)

    If DataGrid1.Col = Val(3) And Len(DataGrid1.Text) >= 0 Then
        REC_ACCOUNT.MoveFirst
        REC_ACCOUNT.Find "AC_CODE='" & DataGrid1.Text & "'", , adSearchForward
        If Not REC_ACCOUNT.EOF Then
            If DataGrid1.Col = Val(3) Then
                DataGrid1.Col = 4
                RECGRID!BCODE = REC_ACCOUNT!AC_CODE
                RECGRID!BNAME = REC_ACCOUNT!Name
            End If
        Else
            RECGRID!BCODE = ""
            RECGRID!BNAME = ""
            DataGrid1.SetFocus
        End If

    ElseIf DataGrid1.Col = Val(1) And Len(DataGrid1.Text) >= 0 Then
        If REC_SAUDA.RecordCount > 0 Then
            REC_SAUDA.MoveFirst
            REC_SAUDA.Find "SAUDACODE='" & DataGrid1.Text & "'", , adSearchForward
            If Not REC_SAUDA.EOF Then
                If DataGrid1.Col = Val(3) Then
                    DataGrid1.Col = 4
                    RECGRID!scode = REC_SAUDA!SAUDACODE
                    RECGRID!SNAME = REC_SAUDA!SAUDANAME
                End If
            Else
                RECGRID!scode = ""
                RECGRID!SNAME = ""
    
                DataGrid1.SetFocus
            End If
        End If
    ElseIf DataGrid1.Col = Val(2) Then
        If DataGrid1.Columns(2).Text = "B" Then
        ElseIf DataGrid1.Columns(2).Text = "S" Then
        Else
            DataGrid1.Columns(2).Text = "B"
            RECGRID!CONTYPE = "B"
            DataGrid1.Col = 3: DataGrid1.SetFocus
        End If
    End If

    Call GRID_SUMMARY
End Sub
Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If FlagFilter Then
        MsgBox "Please press F1 to refresh and than make any entries."
        Cancel = -1
        Exit Sub
    End If
    
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then KeyCode = 0
    If KeyCode = 32 And DataGrid1.Col = 2 Then
        KeyCode = 0
        If RECGRID!CONTYPE = "B" Then
            RECGRID!CONTYPE = "S"
        Else
            RECGRID!CONTYPE = "B"
        End If
        DataGrid1.Col = 2: DataGrid1.SetFocus: Exit Sub
    End If
    If KeyCode = 114 And (DataGrid1.Col = 1 Or DataGrid1.Col = 2 Or DataGrid1.Col = 3 Or DataGrid1.Col = 4) Then
        FlagFilter = True: FilterCol = DataGrid1.Col
        If FilterCol = 2 Then 'CONTYPE
            RECGRID.Filter = "CONTYPE='" & RECGRID!CONTYPE & "'"
        ElseIf FilterCol = 1 Then 'SCODE
            RECGRID.Filter = "SCODE = '" & RECGRID!scode & "'"
        ElseIf FilterCol = 4 Or FilterCol = 3 Then 'BCODE
            RECGRID.Filter = "BCODE = '" & RECGRID!BCODE & "'"
        End If
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        Call GRID_SUMMARY
        Exit Sub
    ElseIf KeyCode = 112 Then  'F1 REFRESH
        FlagFilter = True: FilterCol = 0
        RECGRID.Filter = "SRNO > 0 "
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        Call GRID_SUMMARY
        Exit Sub
    End If
    'If FlagFilter Then MsgBox "Please press F1 to refresh and than make any entries.": Exit Sub
    If DataGrid1.Col = 1 And Shift <> 2 Then
        If Len(RECGRID!scode & "") < 1 Then
            DataCombo1.Visible = True
            DataCombo1.SetFocus
        End If

    ElseIf DataGrid1.Col = 3 And Shift <> 2 Then
        If Len(RECGRID!BCODE & "") < 1 Then
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        End If
                    
    ElseIf KeyCode = 13 And DataGrid1.Col = 6 Then
        MSAUDA = RECGRID!scode
        MSAUDANAME = RECGRID!SNAME
        CONTYPE = RECGRID!CONTYPE
        BCODE = RECGRID!BCODE
        BNAME = RECGRID!BNAME

        RECGRID.MoveNext
        If RECGRID.EOF Then    'And Not FlagFilter
            RECGRID.AddNew
            If Combo1.ListIndex = Val(1) Then   ''OPENING
                RECGRID!CONTYPE = CONTYPE
            Else                        ''LAST INFORMATION CARIES
                RECGRID!scode = MSAUDA
                RECGRID!SNAME = MSAUDANAME
                RECGRID!CONTYPE = CONTYPE
                RECGRID!BCODE = BCODE
                RECGRID!BNAME = BNAME
            End If
            RECGRID.Update
            GridSrNo = GridSrNo + 1
            RECGRID!SRNO = GridSrNo 'IIf(RECGRID.AbsolutePosition < 0, RECGRID.RecordCount + 1, RECGRID.AbsolutePosition)
        End If
        DataGrid1.LeftCol = 0
        DataGrid1.Col = 2

    ElseIf KeyCode = 115 Then   ''F4 KEY
        RNO = InputBox("Enter the row number.", "Sauda")
        If Val(RNO) > Val(0) Then
            RECGRID.MoveFirst
            RECGRID.Find "SRNO=" & Val(RNO) & "", , adSearchForward
            If RECGRID.EOF Then
                MsgBox "Record not found.", vbCritical, "Error"
                RECGRID.MoveFirst
            End If
            DataGrid1.Col = 1
            DataGrid1.SetFocus
        End If

    ElseIf KeyCode = 46 And Shift = 2 Then
        RECGRID.Delete
        If RECGRID.RecordCount = 0 Then
            RECGRID.AddNew
            RECGRID!SRNO = RECGRID.RecordCount
            If Combo1.ListIndex = Val(1) Then
                RECGRID!BRate = Val(Text3.Text)
                RECGRID!SRate = Val(Text3.Text)
            End If

            RECGRID.Update
        End If
        Call DataGrid1_AfterColEdit(0)

    ElseIf DataGrid1.Col = 5 Then
        MROWNO = 0
    End If
End Sub
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If DataGrid1.Col = 2 Then
        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "B" Or KeyAscii = 13 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        
        Call GRID_SUMMARY
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 112 Then
'        Call lblnew_Click
'    ElseIf KeyCode = 113 Then
'        Call lbledit_Click
'    ElseIf KeyCode = 114 Then
'        Call LblDelete_Click
'    ElseIf KeyCode = 115 Then
'        Call lblsave_Click
'    ElseIf KeyCode = 116 Then
'        Call lblcancel_Click
'    ElseIf KeyCode = 117 Then
'        Call lblexit_Click
'    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.Name = "vcDTP1" Then
            SendKeys "{tab}"
        End If
    End If
End Sub
Private Sub Form_Load()
'    Call CANCEL_REC
'----------
'    vcDTP1.MaxDate = MFIN_END: vcDTP1.MinDate = MFIN_BEG
'    vcDTP2.MaxDate = MFIN_END: vcDTP2.MinDate = MFIN_BEG
    GridSrNo = 1
    LblNew.Visible = False: LblEdit.Visible = False: LblCancel.Visible = False: LblDelete.Visible = False: LblSave.Visible = False: LblExit.Visible = False
    Call RECSET
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.Refresh
    Call ClearFormFn(GETContA)
    Frame1.Enabled = False
'--------
    If SWTYPE = "SQL" Then
        MYSQL = "SELECT ITEMCODE, (ITEMCODE+','+ITEMCODE+','+ITEMNAME) AS ITEMNAME FROM ITEMMAST WHERE compcode=" & MC_CODE & "  ORDER BY ITEMCODE"
    Else
        MYSQL = "SELECT ITEMCODE, (ITEMCODE+','+ITEMCODE+','+ITEMNAME) AS ITEMNAME FROM ITEMMAST WHERE compcode=" & MC_CODE & " ORDER BY ITEMCODE"
    End If
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open MYSQL, CNN, adOpenKeyset, adLockReadOnly
    If Not REC.EOF Then
        Set DataCombo2.RowSource = REC
        DataCombo2.BoundColumn = "ITEMCODE"
        DataCombo2.ListField = "ITEMNAME"
    
        Set REC_ACCOUNT = Nothing
        Set REC_ACCOUNT = New ADODB.Recordset
        REC_ACCOUNT.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE compcode=" & MC_CODE & "  ORDER BY AC_CODE", CNN, adOpenKeyset, adLockReadOnly
    
        Set DataCombo3.RowSource = REC_ACCOUNT
        DataCombo3.BoundColumn = "AC_CODE"
        DataCombo3.ListField = "NAME"
    
        Set DataCombo5.RowSource = REC_ACCOUNT
        DataCombo5.BoundColumn = "AC_CODE"
        DataCombo5.ListField = "NAME"
    
        Set REC_SAUDA = Nothing
        Set REC_SAUDA = New ADODB.Recordset
        REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE compcode=" & MC_CODE & " ORDER BY SAUDANAME", CNN, adOpenKeyset, adLockReadOnly
        Set DataCombo4.RowSource = REC_SAUDA
        DataCombo4.BoundColumn = "SAUDACODE"
        DataCombo4.ListField = "SAUDANAME"
    
        Adodc1.ConnectionString = CNN
        Adodc1.RecordSource = "SELECT * FROM CTR_M where compcode=" & MC_CODE & " ORDER BY CONSNO"
        Adodc1.Refresh
    
        Set DataGrid1.DataSource = RECGRID
        DataGrid1.ReBind
        DataGrid1.Refresh
    Else
        Call Get_Selection(12)
    End If
    FlagFilter = False
End Sub
Private Sub Text11_GotFocus()
    If Frame4.Visible = True Then
        DataCombo4.SetFocus
    Else
        If Frame1.Enabled = True Then SetFocus
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    FLOWDIR = 1
    If Len(Trim(Text2.Text)) < 1 Then
        DataCombo1.SetFocus
    Else
        Set REC_SAUDA = Nothing
        Set REC_SAUDA = New ADODB.Recordset
        REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE compcode=" & MC_CODE & " AND SAUDACODE='" & Text2.Text & "'", CNN, adOpenKeyset, adLockReadOnly
        If REC_SAUDA.EOF Then
            MsgBox "Invalid sauda code.", vbExclamation, "Error"
            Cancel = True
        Else
            Text2.Text = REC_SAUDA!SAUDACODE
            DataCombo1.BoundText = CStr(Text2.Text)
            DataCombo2.BoundText = REC_SAUDA!ItemCode
        End If
    End If
End Sub
Private Sub Text3_GotFocus()
    FLOWDIR = 0
    Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Sub RECSET()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "BCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "BNAME", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "BQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "SNAME", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    RECGRID.AddNew
    RECGRID.Update
    RECGRID!SRNO = GridSrNo 'RECGRID.AbsolutePosition
    DataGrid1.Col = 1
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Text3.Text = Format(Text3.Text, "0.00")
    If Combo1.ListIndex = Val(1) Then
        RECGRID!BRate = Val(Text3.Text)
        RECGRID!SRate = Val(Text3.Text)
    End If
End Sub
Sub DELETE_VOUCHER(VOU_NO As String)
'    Set REC = Nothing
'    Set REC = New ADODB.Recordset
'    REC.Open "SELECT * FROM VCHAMT WHERE compcode=" & MC_CODE & " AND VOU_NO='" & VOU_NO & "'", CNN, adOpenForwardOnly, adLockReadOnly
'    Do While Not REC.EOF
'        If REC!DR_CR = "D" Then
'            MYSQL = "UPDATE ACCOUNTM SET DEBIT =DEBIT -" & REC!AMOUNT & " WHERE compcode = " & MC_CODE & " and AC_CODE='" & REC!AC_CODE & "'"
'        Else
'            MYSQL = "UPDATE ACCOUNTM SET CREDIT=CREDIT-" & REC!AMOUNT & " WHERE compcode = " & MC_CODE & " and AC_CODE='" & REC!AC_CODE & "'"
'        End If
'        CNN.Execute MYSQL
'
'        REC.MoveNext
'    Loop
'
'    Set REC = Nothing

    CNN.Execute "DELETE FROM VCHAMT  WHERE compcode=" & MC_CODE & " AND VOU_NO='" & VOU_NO & "'"
    CNN.Execute "DELETE FROM VOUCHER WHERE compcode=" & MC_CODE & " AND VOU_NO='" & VOU_NO & "'"
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static OLDVAL As Integer

    Select Case ColIndex
    Case 0
        If OLDVAL = -1 Then
            RECGRID.Sort = "SRNO DESC"
        Else
            RECGRID.Sort = "SRNO"
        End If
    
    Case 1
        If OLDVAL = -1 Then
            RECGRID.Sort = "BCODE DESC"
        Else
            RECGRID.Sort = "BCODE"
        End If

    Case 2
        If OLDVAL = -1 Then
            RECGRID.Sort = "BNAME DESC"
        Else
            RECGRID.Sort = "BNAME"
        End If

    Case 3
        If OLDVAL = -1 Then
            RECGRID.Sort = "BQNTY DESC"
        Else
            RECGRID.Sort = "BQNTY"
        End If

    Case 4
        If OLDVAL = -1 Then
            RECGRID.Sort = "BRATE DESC"
        Else
            RECGRID.Sort = "BRATE"
        End If
    Case 5
        If OLDVAL = -1 Then
            RECGRID.Sort = "SCODE DESC"
        Else
            RECGRID.Sort = "SCODE"
        End If

    Case 6
        If OLDVAL = -1 Then
            RECGRID.Sort = "SNAME DESC"
        Else
            RECGRID.Sort = "SNAME"
        End If

    Case 7
        If OLDVAL = -1 Then
            RECGRID.Sort = "SQNTY DESC"
        Else
            RECGRID.Sort = "SQNTY"
        End If

    Case 8
        If OLDVAL = -1 Then
            RECGRID.Sort = "SRATE DESC"
        Else
            RECGRID.Sort = "SRATE"
        End If
        
    End Select

    If OLDVAL = -1 Then
        Call VISIBLE_IMAGE(0)
    Else
        Call VISIBLE_IMAGE(1)
    End If

    If OLDVAL = ColIndex Then
        OLDVAL = -1
    Else
        OLDVAL = ColIndex
    End If
    Image1(0).Left = DataGrid1.Left + DataGrid1.Columns(ColIndex).Left + (DataGrid1.Columns(ColIndex).Width) / 2
    Image1(1).Left = DataGrid1.Left + DataGrid1.Columns(ColIndex).Left + (DataGrid1.Columns(ColIndex).Width) / 2
End Sub
Sub VISIBLE_IMAGE(SORT_ORDER As Byte)
    If SORT_ORDER = 1 Then
        Image1(0).Visible = False
        Image1(1).Visible = True
    Else
        Image1(0).Visible = True
        Image1(1).Visible = False
    End If
End Sub

Private Sub vcDTP1_GotFocus()
    If FlagFilter Then FlagFilter = False: DataGrid1.Col = 0: DataGrid1.SetFocus
End Sub
Private Sub vcDTP1_Validate(Cancel As Boolean)
    Set REC_SAUDA = Nothing: Set REC_SAUDA = New ADODB.Recordset
    REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE compcode=" & MC_CODE & " AND MATURITY >= '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'   ORDER BY SAUDANAME", CNN, adOpenKeyset, adLockReadOnly
    If Not REC_SAUDA.EOF Then
        Set DataCombo1.RowSource = REC_SAUDA
        DataCombo1.BoundColumn = "SAUDACODE"
        DataCombo1.ListField = "SAUDANAME"
    Else
        MsgBox "Please define sauda.", vbCritical
        Call CANCEL_REC
    End If
End Sub
Sub GRID_SUMMARY()
    Set REC = Nothing: Set REC = New ADODB.Recordset: Set REC = RECGRID.Clone
    
    If FlagFilter Then
        If FilterCol = 2 Then 'CONTYPE
            REC.Filter = "CONTYPE='" & RECGRID!CONTYPE & "'"
        ElseIf FilterCol = 1 Then 'SCODE
            REC.Filter = "SCODE = '" & RECGRID!scode & "'"
        ElseIf FilterCol = 4 Or FilterCol = 3 Then 'BCODE
            REC.Filter = "BCODE = '" & RECGRID!BCODE & "'"
        End If
    End If
    
    BQNTY = 0: SQNTY = 0: BAmt = 0: SAmt = 0
    Do While Not REC.EOF
        If REC!CONTYPE = "B" Then
            BQNTY = BQNTY + Val(REC!BQNTY & "")
            BAmt = BAmt + (Val(REC!BQNTY & "") * Val(REC!BRate & ""))
        Else
            SQNTY = SQNTY + Val(REC!BQNTY & "")
            SAmt = SAmt + (Val(REC!BQNTY & "") * Val(REC!BRate & ""))
        End If
        REC.MoveNext
    Loop
    Text1.Text = BQNTY: Text4.Text = SQNTY: Text5.Text = Val(Text1.Text) - Val(Text4.Text): Text6.Text = Format(Val(BAmt) - Val(SAmt), "0.00")
''''----------------------------
    If RECGRID.EOF Or RECGRID.BOF Then Exit Sub
    Set REC = Nothing: Set REC = New ADODB.Recordset: Set REC = RECGRID.Clone
    REC.Filter = "SCODE='" & RECGRID!scode & "'"
    BQNTY = 0: SQNTY = 0: BAmt = 0: SAmt = 0
    Do While Not REC.EOF
        If REC!CONTYPE = "B" Then
            BQNTY = BQNTY + Val(REC!BQNTY & "")
            BAmt = BAmt + (Val(REC!BQNTY & "") * Val(REC!BRate & ""))
        Else
            SQNTY = SQNTY + Val(REC!BQNTY & "")
            SAmt = SAmt + (Val(REC!BQNTY & "") * Val(REC!BRate & ""))
        End If
        REC.MoveNext
    Loop
    Text10.Text = BQNTY: Text9.Text = SQNTY
    Text8.Text = Val(Text1.Text) - Val(Text4.Text)
    Text7.Text = Format(Val(BAmt) - Val(SAmt), "0.00")
End Sub
Sub GENERATE_VOUCHER()
    LBAMT = 0: LSAMT = 0
    If RECGRID.RecordCount > 0 Then
        RECGRID.MoveFirst
        RECGRID.Sort = "SCODE Asc"
        LSaudaCode = RECGRID!scode
        Do While Not RECGRID.EOF
            Set GeneralRec1 = Nothing: Set GeneralRec1 = New ADODB.Recordset
            GeneralRec1.Open "SELECT EX.SHREEAC,EX.TRADINGACC,IT.LOT  FROM exmast  AS EX , ITEMMAST AS IM,Saudamast as SM WHERE EX.compcode=" & MC_CODE & " AND EX.COMPCODE= IM.COMPCODE  and EX.compcode = SM.compcode AND EX.EXCODE=IM.EXCHANGECODE  AND  SM.SAUDACODE = '" & RECGRID!scode & "' AND IM.ITEMCODE = SM.ITEMCODE ", CNN, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec1.EOF Then GSHREE = GeneralRec1!shreeac: GTRADING = GeneralRec1!TRADINGACC: GCALVAL = GeneralRec1!LOT
            
            If LSaudaCode = RECGRID!scode Then
                If RECGRID!CONTYPE = "B" Then
                    LBAMT = LBAMT + (RECGRID!BQNTY * RECGRID!BRate * GCALVAL)
                ElseIf RECGRID!CONTYPE = "S" Then
                    LSAMT = LSAMT + (RECGRID!BQNTY * RECGRID!BRate * GCALVAL)
                End If
            Else
                If LBAMT = LSAMT Then
                Else
                    MYSQL = "INSERT INTO VOUCHER(CompCode,VOU_NO, VOU_DT, VOU_TYPE, VOU_PR, BILLNO, BILLDT, USER_NAME, USER_DATE, USER_TIME, USER_ACTION) VALUES(" & MC_CODE & ",'" & VCHNO & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','O','','" & Text2.Text & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & USER_ID & "','" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD')"
                    CNN.Execute MYSQL
                    
                    MAMOUNT = Abs(Val(LBAMT - LSAMT))
                    If Val(LSAMT - LBAMT) > Val(0) Then
                        MCR = "D"
                        MDR = "C"
                        sql = "CREDIT=CREDIT+"
                        SQL1 = "DEBIT=DEBIT+"
                    Else
                        MCR = "C"
                        MDR = "D"
                        sql = "DEBIT=DEBIT+"
                        SQL1 = "CREDIT=CREDIT+"
                    End If
            
                    MNARATION = "Shree for : " & LSaudaCode & ", " & DateValue(vcDTP1.Value)
            
                    ''SHREE POSTING
                    MYSQL = "INSERT INTO VCHAMT(CompCode,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MDR & "'," & Val(GSHREE) & "," & Val(MAMOUNT) & ",'" & MNARATION & "')"
                    CNN.Execute MYSQL
            
                    ''TRADING AC POSTING
                    MYSQL = "INSERT INTO VCHAMT(CompCode,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MCR & "'," & Val(GTRADING) & "," & Val(MAMOUNT) & ",'" & MNARATION & "')"
                    CNN.Execute MYSQL
                End If
                LBAMT = 0: LSAMT = 0
                If Not RECGRID.EOF Then
                    LSaudaCode = RECGRID!scode
                    If RECGRID!CONTYPE = "B" Then
                        LBAMT = LBAMT + (RECGRID!BQNTY * RECGRID!BRate * GCALVAL)
                    ElseIf RECGRID!CONTYPE = "S" Then
                        LSAMT = LSAMT + (RECGRID!BQNTY * RECGRID!BRate * GCALVAL)
                    End If
                End If
            End If
            RECGRID.MoveNext
        Loop
        If LBAMT = LSAMT Then
        Else
            MYSQL = "INSERT INTO VOUCHER(CompCode,VOU_NO, VOU_DT, VOU_TYPE, VOU_PR, BILLNO, BILLDT, USER_NAME, USER_DATE, USER_TIME, USER_ACTION) VALUES(" & MC_CODE & ",'" & VCHNO & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','O','','" & Text2.Text & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & USER_ID & "','" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD')"
            CNN.Execute MYSQL
            
            MAMOUNT = Abs(Val(LBAMT - LSAMT))
            If Val(LSAMT - LBAMT) > Val(0) Then
                MCR = "D"
                MDR = "C"
                sql = "CREDIT=CREDIT+"
                SQL1 = "DEBIT=DEBIT+"
            Else
                MCR = "C"
                MDR = "D"
                sql = "DEBIT=DEBIT+"
                SQL1 = "CREDIT=CREDIT+"
            End If
            MNARATION = "Shree for : " & LSaudaCode & ", " & DateValue(vcDTP1.Value)
            ''SHREE POSTING
            MYSQL = "INSERT INTO VCHAMT(CompCode,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MDR & "'," & Val(GSHREE) & "," & Val(MAMOUNT) & ",'" & MNARATION & "')"
            CNN.Execute MYSQL
    
            ''TRADING AC POSTING
            MYSQL = "INSERT INTO VCHAMT(CompCode,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MCR & "'," & Val(GTRADING) & "," & Val(MAMOUNT) & ",'" & MNARATION & "')"
            CNN.Execute MYSQL
        End If
    End If
End Sub
Public Sub lblcancel_Click()
    Call GETMAIN.ButtonClick(5)
    LblNew.Visible = True: LblEdit.Visible = True: LblDelete.Visible = True
    LblSave.Visible = False: LblCancel.Visible = False
End Sub
Public Sub lblsave_Click()
    Call GETMAIN.ButtonClick(4)
'    LblNew.Visible = True: LblEdit.Visible = True: LblDelete.Visible = True
'    LblSave.Visible = False: LblCancel.Visible = False
End Sub
Public Sub lblexit_Click()
    Call GETMAIN.ButtonClick(6)
End Sub
Public Sub lblnew_Click()
    Call GETMAIN.ButtonClick(1)
    LblNew.Visible = False: LblEdit.Visible = False: LblDelete.Visible = False
    LblSave.Visible = True: LblCancel.Visible = True
End Sub
Public Sub lbledit_Click()
    Call GETMAIN.ButtonClick(2)
    LblNew.Visible = False: LblEdit.Visible = False: LblDelete.Visible = False
    LblSave.Visible = True: LblCancel.Visible = True
End Sub
Public Sub LblDelete_Click()
    Call GETMAIN.ButtonClick(3)
'    LblNew.Visible = False: LblEdit.Visible = False: LblDelete.Visible = False
'    LblSave.Visible = True: LblCancel.Visible = True
End Sub
