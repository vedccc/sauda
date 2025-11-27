VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form GETContBS 
   BackColor       =   &H80000000&
   ClientHeight    =   8490
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1815
      Begin VB.Line Line8 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
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
            Picture         =   "GetcontBS.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GetcontBS.frx":0452
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
      Height          =   6045
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   11585
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   360
         Left            =   6600
         TabIndex        =   4
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   12582912
         Text            =   "DataCombo4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   5520
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   360
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   12582912
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1_old 
         Height          =   540
         Left            =   45
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   953
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   16
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
            DataField       =   "BCODE"
            Caption         =   "Buyer"
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
            DataField       =   "BNAME"
            Caption         =   "Buyer Name"
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "SCODE"
            Caption         =   "Seller"
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
            DataField       =   "SNAME"
            Caption         =   "Seller Name"
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
            DataField       =   "SQNTY"
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
         BeginProperty Column08 
            DataField       =   "SRATE"
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
         BeginProperty Column09 
            DataField       =   "LInvNo"
            Caption         =   "LInvNo"
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
            DataField       =   "RInvNo"
            Caption         =   "RInvNo"
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
            DataField       =   "DImport"
            Caption         =   "DImport"
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
            DataField       =   "CONTIME"
            Caption         =   "CONTIME"
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
            DataField       =   "UserId"
            Caption         =   "UserId"
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
            DataField       =   "LCLCODE"
            Caption         =   "LCLCODE"
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
         BeginProperty Column15 
            DataField       =   "RCLCODE"
            Caption         =   "RCLCODE"
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
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "GetcontBS.frx":08A4
         Left            =   10080
         List            =   "GetcontBS.frx":08AE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   255
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   735
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   240
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   16711680
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37860.8625462963
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   1560
         TabIndex        =   6
         Top             =   735
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   "DataCombo2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Height          =   15
         Left            =   165
         TabIndex        =   13
         Top             =   1200
         Width           =   11415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7011
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
         ColumnCount     =   12
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
         BeginProperty Column07 
            DataField       =   "InvNo"
            Caption         =   "InvNo"
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
            DataField       =   "DImport"
            Caption         =   "DImport"
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
            DataField       =   "CONTIME"
            Caption         =   "CONTIME"
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
            DataField       =   "UserId"
            Caption         =   "UserId"
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
            DataField       =   "CLCODE"
            Caption         =   "CLCODE"
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
               Locked          =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   2459.906
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   6164.788
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserId"
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
         Left            =   5760
         TabIndex        =   31
         Top             =   780
         Width           =   705
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   0
         Left            =   720
         Picture         =   "GetcontBS.frx":08C5
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 to open new party"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5040
         TabIndex        =   24
         Top             =   1215
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   1
         Left            =   1080
         Picture         =   "GetcontBS.frx":0BCF
         Stretch         =   -1  'True
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shree"
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
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   5595
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Difference"
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
         Index           =   6
         Left            =   5640
         TabIndex        =   20
         Top             =   5595
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sold"
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
         Index           =   5
         Left            =   8760
         TabIndex        =   19
         Top             =   5595
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bought "
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
         Index           =   0
         Left            =   2520
         TabIndex        =   18
         Top             =   5595
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Index           =   18
         Left            =   9240
         TabIndex        =   12
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   9240
         TabIndex        =   11
         Top             =   690
         Width           =   1110
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item  Name"
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
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda"
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
         Index           =   2
         Left            =   3240
         TabIndex        =   9
         Top             =   315
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract Date"
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
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   1425
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1390
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Authors"
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
      TabIndex        =   30
      Top             =   480
      Width           =   930
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
      TabIndex        =   29
      Top             =   480
      Width           =   885
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
      Left            =   6840
      TabIndex        =   28
      Top             =   480
      Width           =   1170
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
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
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
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   1020
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
      Left            =   10560
      TabIndex        =   25
      Top             =   480
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      BorderWidth     =   12
      Height          =   6300
      Left            =   75
      Top             =   840
      Width           =   11685
   End
End
Attribute VB_Name = "GETContBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim Lparty As String
Dim LConNo As Long
Dim LUserId As String
Dim LConSno As Integer
Dim LDataImport As Byte
Dim OldDate As Date
Dim FLOWDIR As Byte
Dim VCHNO As String
Dim GRIDPOS As Byte
Dim LConType As String
Public fb_press As Byte
Dim REC As ADODB.Recordset
Dim RECEX As ADODB.Recordset
Dim RECGRID As ADODB.Recordset
Dim REC_SAUDA As ADODB.Recordset
Dim REC_ACCOUNT As ADODB.Recordset
Dim REC_CloRate As ADODB.Recordset
Sub ADD_REC()
    If REC_ACCOUNT.RecordCount > 0 Then
        Frame1.Enabled = True: Combo1.ListIndex = 0
        Call Get_Selection(1)
        If vcDTP1.Enabled Then vcDTP1.SetFocus
    Else
        Call CANCEL_REC
    End If
End Sub
Sub SAVE_REC()
    On Error GoTo ERR1
    'validation
    If vcDTP1.Value < MFIN_BEG Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > MFIN_END Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    
    If Val(Text1.Text) = 0 Then MsgBox "Please check entries.", vbCritical: Exit Sub
    If Val(Text4.Text) = 0 Then MsgBox "Please check entries.", vbCritical:  Exit Sub
        
    
    Set REC_SAUDA = Nothing: Set REC_SAUDA = New ADODB.Recordset
    REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE COMPCODE=" & MC_CODE & " AND SAUDACODE='" & Text2.Text & "'", cnn, adOpenForwardOnly, adLockReadOnly
    If REC_SAUDA.EOF Then
        MsgBox "Invalid Sauda Code.", vbExclamation, "Error": Text2.SetFocus: Exit Sub
    Else
        Set GeneralRec1 = Nothing: Set GeneralRec1 = New ADODB.Recordset
        GeneralRec1.Open "SELECT EX.SHREEAC,EX.TRADINGACC FROM EXMAST AS EX , ITEMMAST AS IM WHERE EX.COMPCODE=" & MC_CODE & " AND EX.COMPCODE= IM.COMPCODE  AND EX.EXCODE=IM.EXCHANGECODE  AND  IM.ITEMCODE = '" & REC_SAUDA!ItemCode & "'", cnn, adOpenForwardOnly, adLockReadOnly
        If Not GeneralRec1.EOF Then
             GSHREE = GeneralRec1!shreeac
             GTRADING = GeneralRec1!TRADINGACC
        End If
    End If

    RECGRID.Sort = "SRNO"
    If RECGRID.RecordCount > 0 Then
        cnn.BeginTrans
        CNNERR = True
        If fb_press = 1 Then
            VCHNO = VOUCHER_NUMBER("CONT", FIN_YEAR(vcDTP1.Value))
            If Not Adodc1.Recordset.EOF Then
                Set REC = Nothing: Set REC = New ADODB.Recordset
                REC.Open "SELECT CONSNO FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND SAUDA='" & DataCombo1.BoundText & "' AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", cnn, adOpenForwardOnly, adLockReadOnly
                If Not REC.EOF Then
                    CONSNO = REC!CONSNO
                Else
                    Set REC = Nothing: Set REC = New ADODB.Recordset
                    REC.Open "SELECT MAX(CONSNO) FROM CTR_M WHERE COMPCODE=" & MC_CODE & "", cnn, adOpenForwardOnly, adLockReadOnly
                    CONSNO = Val(REC.Fields(0) & "") + Val(1)
                End If
                Set REC = Nothing
            Else
                CONSNO = 1
            End If
        Else
            CONSNO = Adodc1.Recordset!CONSNO
            VCHNO = Adodc1.Recordset!VOU_NO & ""
            If Len(Trim(VCHNO)) < Val(1) Then
                VCHNO = VOUCHER_NUMBER("CONT", FIN_YEAR(vcDTP1.Value))
            End If
        End If
        Call DELETE_VOUCHER(VCHNO)
        If fb_press = 2 Then
            If Len(Trim(DataCombo4.BoundText)) > 0 Then
                cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " AND userid = '" & DataCombo4.BoundText & "'"
            Else
                cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
            End If
            If Text3.Locked Then
            Else
                cnn.Execute "DELETE FROM CTR_R WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
            End If
            cnn.Execute "DELETE FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
        End If
        cnn.Execute "DELETE FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & CONSNO & ""
        LDataImport = IIf(IsNull(LDataImport), 0, LDataImport)
        MYSQL = "INSERT INTO CTR_M(COMPCODE,CONSNO, CONDATE, SAUDA, ITEMCODE, CLOSERATE, VOU_NO, PATTAN,DataImport) VALUES(" & MC_CODE & "," & CONSNO & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "', '" & Text2.Text & "', '" & DataCombo2.BoundText & "', " & Val(Text3.Text) & ", '" & VCHNO & "', '" & Mid(Combo1.Text, 1, 1) & "'," & LDataImport & ")"
        cnn.Execute MYSQL
        Dim BOOLAC As String * 1
        Dim RC As ADODB.Recordset
        'do not initialized LPARTY here
        RECGRID.MoveFirst
        Do While Not RECGRID.EOF
            ''RECORDSET RC IS CHECKING WHETHER THE PARTY IS PERSONNEL OR NOT
            BOOLAC = "N"
            MCL = ""
            If Len(RECGRID!BNAME & "") > Val(0) Then    ''WHEN BUYER AND SELLER BOTH ARE THERE
                If RECGRID!BQNTY > Val(0) And RECGRID!BRate > Val(0) Then                   ''QNTY AND RATE REQUIRED
                    If InStr(Lparty, RECGRID!BCODE) Then
                    Else
                        If Lparty = "" Then
                            Lparty = "'" & RECGRID!BCODE & "'"
                        Else
                            Lparty = Lparty & ",'" & RECGRID!BCODE & "'"
                        End If
                    End If
                    If RECGRID!DIMPORT = 0 Then
                        MCL = RECGRID!BCODE
                    Else
                        MCL = RECGRID!LCLCODE
                    End If
                    
                    MYSQL = "INSERT INTO CTR_D (COMPCODE,CLCODE,CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT,DATAIMPORT,CONTIME,UserId) VALUES "
                    MYSQL = MYSQL & "(" & MC_CODE & ",'" & MCL & "'," & Val(CONSNO) & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'," & Val(RECGRID!SRNO) & ",'" & Text2.Text & "', '" & DataCombo2.BoundText & "', '" & RECGRID!BCODE & "', " & Val(RECGRID!BQNTY) & "," & Val(RECGRID!BRate) & ",'" & (RECGRID!CONTYPE & "") & "', '" & BOOLAC & "'," & RECGRID!DIMPORT & ",'" & RECGRID!CONTIME & "','" & (RECGRID!UserId & "") & "')"
                    cnn.Execute MYSQL
                End If
            End If
            RECGRID.MoveNext
        Loop
        If Text3.Locked Then
        Else
            If Val(Text3.Text) > Val(0) Then
                MYSQL = "INSERT INTO CTR_R(COMPCODE,CONSNO, SAUDA, CONDATE, PATTAN, OPRATE, HGRATE, LOWRATE, CLOSERATE,DataImport) VALUES(" & MC_CODE & "," & CONSNO & ",'" & DataCombo1.BoundText & "', '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & Mid(Combo1.Text, 1, 1) & "', 0, 0, 0, " & Val(Text3.Text) & ",0)"
                cnn.Execute MYSQL
            End If
        End If
        If Val(Text6.Text) <> Val(0) Then
            MYSQL = "INSERT INTO VOUCHER(COMPCODE,VOU_NO, VOU_DT, VOU_TYPE, VOU_PR, BILLNO, BILLDT, USER_NAME, USER_DATE, USER_TIME, USER_ACTION) VALUES(" & MC_CODE & ",'" & VCHNO & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','O','','" & Text2.Text & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & USER_ID & "','" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD')"
            cnn.Execute MYSQL
            MAMOUNT = Abs(Val(Text6.Text))
            If Val(Text6.Text) < Val(0) Then
                MCR = "C": MDR = "D": sql = "DEBIT=DEBIT+": SQL1 = "CREDIT=CREDIT+"
            Else
                MCR = "D": MDR = "C"
                sql = "CREDIT=CREDIT+": SQL1 = "DEBIT=DEBIT+"
            End If
            MNARATION = "Shree for : " & Text2.Text & ", " & Format(vcDTP1.Value, "yyyy/MM/dd")
            ''SHREE POSTING
            MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MDR & "','" & GSHREE & "'," & Val(MAMOUNT) & ",'" & MNARATION & "')"
            cnn.Execute MYSQL
            ''TRADING AC POSTING
            MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MCR & "','" & GTRADING & "'," & Val(MAMOUNT) & ",'" & MNARATION & "')"
            cnn.Execute MYSQL
        End If
        
        Call UpdateBrokRateType(True, True, Lparty, "'" & DataCombo2.BoundText & "'", vcDTP1.Value, vcDTP1.Value, vcDTP1.Value, "'" & DataCombo1.BoundText & "'")
        'Call UpdateStanding(Lparty, "'" & DataCombo1.BoundText & "'", vcDTP1.Value, vcDTP1.Value)
        
        cnn.CommitTrans
        CNNERR = False
         Adodc1.Refresh
     ''BILL GENERATION
        MFROMDATE = Format(vcDTP1.Value, "yyyy/MM/dd")
        MYSQL = "SELECT MATURITY FROM SAUDAMAST WHERE COMPCODE=" & MC_CODE & " AND SAUDACODE = '" & Text2.Text & "'"
        Set REC = Nothing: Set REC = New ADODB.Recordset: REC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
        If Not REC.EOF Then MTODATE = REC.Fields(0)
        cnn.BeginTrans
        CNNERR = False
        If BILL_GENERATION(CDate(MFROMDATE), CDate(MTODATE), "'" & Text2.Text & "'", "") Then
            cnn.CommitTrans: CNNERR = False
        Else
            cnn.RollbackTrans: CNNERR = False
        End If
    End If
    Call CANCEL_REC
    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
    If CNNERR = True Then cnn.RollbackTrans: CNNERR = False
End Sub
Sub CANCEL_REC()
    vcDTP1.Enabled = True: Text2.Enabled = True: DataCombo1.Enabled = True: Combo1.Enabled = True: DataCombo4.Enabled = True: Text3.Enabled = True
    LConType = "B"
    LConNo = 0
    Call RECSET
    fb_press = 0
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.Refresh
    Label2.Visible = False
    DataCombo3.Visible = False
    Call ClearFormFn(GETContBS)
    Call Get_Selection(10)
    Combo1.ListIndex = -1: Frame1.Enabled = False
End Sub
Function MODIFY_REC(LCONDATE As Date, LSAUDA As String, LPATTAN As String) As Boolean
    
    Set REC = Nothing: Set REC = New ADODB.Recordset
    MYSQL = "SELECT * FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND CONDATE='" & Format(LCONDATE, "yyyy/MM/dd") & "' AND SAUDA='" & LSAUDA & "' AND PATTAN='" & Mid(LPATTAN, 1, 1) & "'"
    REC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
    If REC.EOF Then
        If fb_press = 2 Then
            MsgBox "Transaction does not exists for the SELECTed creteria?", vbExclamation
            OldDate = vcDTP1.Value: GETContBS.fb_press = 1: vcDTP1.Value = OldDate: MODIFY_REC = True
            Exit Function
        ElseIf fb_press = 1 Then
            MODIFY_REC = True
        End If
        Exit Function
    Else
         If fb_press = 1 Then
            Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
            MYSQL = "SELECT * FROM CTR_D WHERE COMPCODE=" & MC_CODE & " AND CONDATE='" & Format(LCONDATE, "yyyy/MM/dd") & "' AND SAUDA = '" & LSAUDA & "'"
            GeneralRec.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                MsgBox "Contract already exists.Please press enter to modify contract.", vbInformation
                OldDate = vcDTP1.Value: GETContBS.fb_press = 2: vcDTP1.Value = OldDate
                GETMAIN.StatusBar1.Panels(2).Text = "Modify Record"
                MODIFY_REC = False
                Exit Function
            Else
                MODIFY_REC = True
                Exit Function
            End If
         Else
            Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
            MYSQL = "SELECT * FROM CTR_D WHERE COMPCODE=" & MC_CODE & " AND CONDATE='" & Format(LCONDATE, "yyyy/MM/dd") & "' AND SAUDA='" & LSAUDA & "' "
            GeneralRec.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                MODIFY_REC = True
            Else
                MsgBox "Contract does not exists.Please add new contract.", vbInformation
                OldDate = vcDTP1.Value
                Call CANCEL_REC
                GETContBS.fb_press = 1
                vcDTP1.Value = OldDate
                GETContBS.ADD_REC
                GETMAIN.StatusBar1.Panels(2).Text = "Add Record"
                MODIFY_REC = False
                Exit Function
            End If
        End If
        LDataImport = IIf(REC!DATAIMPORT & "", 1, 0)
    End If
    
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
            DataGrid1.Columns(7).Locked = True
        Else
            Combo1.ListIndex = Val(1)
            DataGrid1.Columns(7).Locked = False
        End If
    End With

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    MYSQL = "SELECT CTR_D.*, A.NAME AS NAME FROM CTR_D, ACCOUNTD AS A WHERE CTR_D.COMPCODE=" & MC_CODE & " AND CTR_D.COMPCODE=A.COMPCODE AND CTR_D.PARTY=A.AC_CODE AND CTR_D.CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " ORDER BY CONNO, CONTYPE"
    REC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly

    Call RECSET
    RECGRID.Delete
    Lparty = ""
    Do While Not REC.EOF
        If (REC!UserId = DataCombo4.BoundText) Or DataCombo4.Text = "" Then
            LConNo = REC!CONNO
            If Trim((REC!PARTY & "")) = "" Then
            Else
                If InStr(Lparty, REC!PARTY) Then
                Else
                    If Lparty = "" Then
                        Lparty = "'" & REC!PARTY & "'"
                    Else
                        Lparty = Lparty & ",'" & REC!PARTY & "'"
                    End If
                End If
                If InStr(Lparty, REC!PARTY) Then
                Else
                    If Lparty = "" Then
                        Lparty = "'" & REC!PARTY & "'"
                    Else
                        Lparty = Lparty & ",'" & REC!PARTY & "'"
                    End If
                End If
            End If
            RECGRID.AddNew
            RECGRID!SRNO = LConNo 'RECGRID.AbsolutePosition
            RECGRID!CONTYPE = REC!CONTYPE & ""
            RECGRID!BCODE = REC!PARTY & ""
            RECGRID!CLCODE = REC!CLCODE & ""
            RECGRID!BNAME = REC!Name
            RECGRID!BQNTY = REC!QTY
            RECGRID!BRate = REC!Rate
            RECGRID!INVNO = Val(REC!INVNO & "")
            RECGRID!DIMPORT = IIf(REC!DATAIMPORT & "", 1, 0)
            RECGRID!CONTIME = IIf(IsNull(REC!CONTIME), Time, REC!CONTIME)
            RECGRID!UserId = REC!UserId & ""
            
'            REC.MoveNext
'            If REC.EOF Then
'                MsgBox "Seller party missing.", vbInformation
'                RECGRID.Update
'                Exit Do
'            Else
'                RECGRID!scode = REC!PARTY & ""
'                RECGRID!RCLCODE = REC!CLCODE & ""
'                RECGRID!SNAME = REC!Name
'                RECGRID!SQNTY = REC!QTY
'                RECGRID!SRate = REC!Rate
'                RECGRID!RInvNo = Val(REC!INVNO & "")
'            End If
            RECGRID.Update
        End If
        REC.MoveNext
    Loop
    
    Set DataGrid1.DataSource = RECGRID
    Call DataGrid1_AfterColEdit(0)

    If fb_press = 3 Then
        If MsgBox("You are about to DELETE all contracts. Confirm Delete?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            On Error GoTo ERR1
            cnn.BeginTrans
            CNNERR = True
                MYSQL = "DELETE FROM CTR_D WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & Adodc1.Recordset!CONSNO & ""
                cnn.Execute MYSQL

                MYSQL = "DELETE FROM CTR_R WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & Adodc1.Recordset!CONSNO & ""
                cnn.Execute MYSQL

            Call DELETE_VOUCHER(Adodc1.Recordset!VOU_NO & "")

            MYSQL = "DELETE FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & Adodc1.Recordset!CONSNO & ""
            cnn.Execute MYSQL

            ''REGENERATING SETTLEMENTS
              MFROMDATE = Format(vcDTP1.Value, "yyyy/MM/dd")

            ''TO FIND TODATE
            MYSQL = "SELECT MATURITY FROM SAUDAMAST WHERE COMPCODE=" & MC_CODE & " AND SAUDACODE = '" & Text2.Text & "'"
            Set REC = Nothing: Set REC = New ADODB.Recordset: REC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
            If Not REC.EOF Then MTODATE = REC.Fields(0)

              Call UpdateBrokRateType(True, True, Lparty, "'" & DataCombo2.BoundText & "'", vcDTP1.Value, vcDTP1.Value, vcDTP1.Value)
              'Call UpdateStanding(Lparty, "'" & DataCombo1.BoundText & "'", vcDTP1.Value, vcDTP1.Value)
                
              If BILL_GENERATION(CDate(MFROMDATE), CDate(MTODATE), "'" & Text2.Text & "'") Then
                cnn.CommitTrans
                CNNERR = False
              Else
                cnn.RollbackTrans
                CNNERR = False
              End If
            ''REGENERATING UPTO HERE
            If fb_press = 2 Then
                GETMAIN.Toolbar1_Buttons(4).Enabled = True: GETMAIN.saverec.Enabled = True
            ElseIf fb_press = 3 Then
                Call CANCEL_REC
            End If
            MODIFY_REC = True
            Exit Function
ERR1:
            If Err.Number <> 0 Then
                MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
            End If
            If CNNERR = True Then cnn.RollbackTrans: CNNERR = False
            Call CANCEL_REC
        End If
    End If
End Function
Private Sub Combo1_GotFocus()
    If FLOWDIR = 1 Then
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT * FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND SAUDA='" & DataCombo1.BoundText & "'", cnn, adOpenForwardOnly, adLockReadOnly
        If REC.EOF Then SendKeys "%{DOWN}"
    Else
       Text2.SetFocus
    End If
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 1 Then flag = True
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If flag Then
        Text3.SetFocus
    Else
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT * FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND PATTAN='O' AND SAUDA='" & DataCombo1.BoundText & "'", cnn, adOpenForwardOnly, adLockReadOnly
    
        If Not REC.EOF Then
            If Format(vcDTP1.Value, "yyyy/MM/dd") < REC!CONDATE Then
                MsgBox "Opening for this SAUDA has been already entered on " & Format(REC!CONDATE, "yyyy/MM/dd"), vbExclamation, "Warning"
                vcDTP1.Value = Date
                Cancel = True
                Exit Sub
            End If
        Else
            If Not REC.EOF Then
                If REC!CONDATE > Format(vcDTP1.Value, "yyyy/MM/dd") Then
                    MsgBox "Opening for this SAUDA has been already entered on " & Format(REC!CONDATE, "yyyy/MM/dd"), vbExclamation, "Warning"
                    vcDTP1.Value = Date
                    Exit Sub
                End If
            End If
        End If
    
        If Combo1.ListIndex = Val(0) Then   ''CONTRACT
            DataGrid1.Columns(7).Locked = True
        Else                                ''OPENING
            DataGrid1.Columns(7).Locked = False
        End If
        
        'Check UserId*****
        LConSno = 0: Set REC = Nothing: Set REC = New ADODB.Recordset
        MYSQL = "SELECT ConSno FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND SAUDA='" & DataCombo1.BoundText & "' AND PATTAN='" & Mid(Combo1.Text, 1, 1) & "'"
        REC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
        If Not REC.EOF Then LConSno = REC!CONSNO
            
        Set REC = Nothing: Set REC = New ADODB.Recordset
        MYSQL = "SELECT distinct fmly.fmlyNAME,fmly.fmlycode FROM CTR_D, accfmly  AS fmly WHERE CTR_D.COMPCODE=" & MC_CODE & " AND CTR_D.COMPCODE=fmly.COMPCODE AND CTR_D.userid=fmly.fmlycode AND CTR_D.CONSNO = " & Val(LConSno) & " ORDER BY fmlyname"
        REC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
        If Not REC.EOF Then
            Set DataCombo4.RowSource = REC: DataCombo4.ListField = "Fmlyname": DataCombo4.BoundColumn = "FmlyCode"   ': DataCombo4.SetFocus
        Else
            If MODIFY_REC(vcDTP1.Value, DataCombo1.BoundText, Combo1.Text) Then
            Else
                Cancel = True
            End If
        End If
    End If
    flag = False
End Sub
Private Sub DataCombo1_GotFocus()
    SendKeys "%{DOWN}"
    If Len(Trim(Text2.Text)) > 0 Then Combo1.SetFocus
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    
    Set REC_SAUDA = Nothing: Set REC_SAUDA = New ADODB.Recordset
    REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE COMPCODE=" & MC_CODE & " AND SAUDACODE='" & DataCombo1.BoundText & "'", cnn, adOpenForwardOnly, adLockReadOnly
    If REC_SAUDA.EOF Then
        Cancel = True
    Else
        Text2.Text = REC_SAUDA!SAUDACODE
        DataCombo1.BoundText = Text2.Text
        DataCombo2.BoundText = REC_SAUDA!ItemCode
                
        MYSQL = "SELECT LOT FROM ITEMMAST WHERE COMPCODE=" & MC_CODE & " AND  ITEMCODE='" & REC_SAUDA!ItemCode & "'"
        Set RECEX = Nothing: Set RECEX = New ADODB.Recordset: RECEX.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
        If Not RECEX.EOF Then
            GCALVAL = RECEX!LOT
        Else
            GCALVAL = 0
        End If
        
                
                
        MYSQL = "SELECT fy.fmlycode FROM EXMAST AS EX , ITEMMAST AS IT , accfmly  as  fy WHERE ex.COMPCODE=" & MC_CODE & " AND ex.COMPCODE=it.COMPCODE AND ex.EXCODE=it.ExchangeCode AND IT.ITEMCODE = '" & REC_SAUDA!ItemCode & "' AND ex.COMPCODE=fy.COMPCODE AND EX.CONTRACTACC = fy.fmlyhead "
        Set RECEX = Nothing: Set RECEX = New ADODB.Recordset: RECEX.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
        If Not RECEX.EOF Then
            LUserId = RECEX!FMLYCode
        Else
            'create new branch with head of exchange contract account
            MYSQL = "SELECT EX.CONTRACTACC,ex.ExName FROM EXMAST AS EX , ITEMMAST AS IT  WHERE ex.COMPCODE=" & MC_CODE & " AND ex.COMPCODE=it.COMPCODE AND ex.EXCODE=it.ExchangeCode AND IT.ITEMCODE = '" & REC_SAUDA!ItemCode & "'  "
            Set RECEX = Nothing: Set RECEX = New ADODB.Recordset: RECEX.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
            If Not RECEX.EOF Then
                cnn.Execute "Insert Into AccFmly (COMPCODE,FMLYCode,FMLYName,FMLYHEAD) values (" & MC_CODE & ",'" & RECEX!ContractACC & "','" & RECEX!ContractACC & "','" & RECEX!ContractACC & "')"
                LUserId = RECEX!ContractACC
                MsgBox "Generated new default branch for " & RECEX!EXNAME, vbInformation
            End If
        End If
        Combo1.SetFocus
    End If
    Call GetCloseRate
End Sub
Private Sub DataCombo2_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo3_GotFocus()

    If DataGrid1.Col = 3 Then
'        DataGrid1.Text = "": DataGrid1.Col = 1
        Label2.Visible = True: Label2.Left = 1080
        DataCombo3.Left = Val(1910)
        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    End If
        SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        RECGRID!BCODE = DataCombo3.BoundText
        RECGRID!BNAME = DataCombo3.Text
        RECGRID!UserId = LUserId
        DataGrid1.Col = 4
        DataGrid1.Enabled = True
        DataGrid1.SetFocus
        DataCombo3.Visible = False: Label2.Visible = False
    ElseIf KeyCode = 27 Then
        DataGrid1.Enabled = True
        DataGrid1.SetFocus
        DataCombo3.Visible = False: Label2.Visible = False
    ElseIf KeyCode = 114 Then   'F3  NEW PARTY
        GETACNT.Show
        GETACNT.ZOrder
        GETACNT.add_record
    ElseIf KeyCode = 18 Then
        DataCombo3.Visible = True: DataCombo3.SetFocus
    End If
End Sub
Private Sub DataCombo3_Validate(Cancel As Boolean)
    If DataCombo3.Visible = True Then
        Cancel = True
    Else
        Label2.Visible = False
    End If
End Sub
Private Sub DataCombo4_Validate(Cancel As Boolean)
    If MODIFY_REC(vcDTP1.Value, DataCombo1.BoundText, Combo1.Text) Then
        If Len(Trim(DataCombo4.BoundText)) > 0 Then LUserId = DataCombo4.BoundText
    Else
        Cancel = True
    End If
End Sub
Public Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)

    If ColIndex = Val(3) Then
        REC_ACCOUNT.MoveFirst
        REC_ACCOUNT.Find "AC_CODE='" & DataGrid1.Text & "'", , adSearchForward
        If Not REC_ACCOUNT.EOF Then
                RECGRID!BCODE = REC_ACCOUNT!AC_CODE
                RECGRID!BNAME = REC_ACCOUNT!Name
        Else
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        End If
    ElseIf ColIndex = 2 Then
        DataGrid1.Text = UCase(DataGrid1.Text)
        LConType = UCase(DataGrid1.Text)
    
    ElseIf ColIndex = 3 Or ColIndex = 4 Then
        ''IF CONTRACT THEN ONLY CHANGE OCCURS
        If Val(RECGRID!BRate & "") > 0 Then
            If Combo1.ListIndex = Val(0) Then
                RECGRID!SQNTY = RECGRID!BQNTY
                If Val(RECGRID!SRate & "") = Val(0) Then RECGRID!SRate = RECGRID!BRate
            End If
        Else
            If ColIndex = 3 Then
            Else
                MsgBox "Rate can not be zero.Please enter rate.", vbCritical
                DataGrid1.Col = 4: DataGrid1.Enabled = True: DataGrid1.SetFocus
            End If
        End If
    End If
    Set REC = Nothing: Set REC = New ADODB.Recordset: Set REC = RECGRID.Clone
    BQNTY = 0: SQNTY = 0: BAmt = 0: SAmt = 0
    Do While Not REC.EOF
        If (UCase(REC!CONTYPE) = "B") Then
            BQNTY = BQNTY + Val(REC!BQNTY & "")
            BAmt = BAmt + (Val(REC!BQNTY & "") * Val(REC!BRate & "")) * GCALVAL
        ElseIf (UCase(REC!CONTYPE) = "S") Then
            SQNTY = SQNTY + Val(REC!BQNTY & "")
            SAmt = SAmt + (Val(REC!BQNTY & "") * Val(REC!BRate & "")) * GCALVAL
        End If
        REC.MoveNext
    Loop
    Text1.Text = BQNTY: Text4.Text = SQNTY
    Text5.Text = Val(Text1.Text) - Val(Text4.Text)
    Text6.Text = Format(Val(BAmt) - Val(SAmt), "0.00")
End Sub
Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If fb_press = 2 Then
        If Val(DataGrid1.Columns(9).Text) > 0 And (ColIndex = 1 Or ColIndex = 2 Or ColIndex = 3 Or ColIndex = 4) Then
            MsgBox "Invoice already generated.", vbCritical: Cancel = 1: DataGrid1.Col = ColIndex: DataGrid1.Enabled = True: DataGrid1.SetFocus: Exit Sub
        ElseIf Val(DataGrid1.Columns(10).Text) > 0 And (ColIndex = 5 Or ColIndex = 6 Or ColIndex = 7 Or ColIndex = 8) Then
            MsgBox "Invoice already generated.", vbCritical: Cancel = 1: DataGrid1.Col = ColIndex: DataGrid1.Enabled = True: DataGrid1.SetFocus: Exit Sub
        End If
    End If
End Sub
Private Sub DataGrid1_GotFocus()
    vcDTP1.Enabled = False
    Text2.Enabled = False
    DataCombo1.Enabled = False
    Combo1.Enabled = False
    DataCombo4.Enabled = False
    Text3.Enabled = False
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And DataGrid1.Col = 6 Then
        BCODE = RECGRID!BCODE
        BNAME = RECGRID!BNAME
        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            If Combo1.ListIndex = Val(1) Then   ''OPENING
                RECGRID!BRate = Val(Text3.Text)
            Else                        ''LAST INFORMATION CARIES
                RECGRID!BCODE = BCODE
                RECGRID!BNAME = BNAME
                RECGRID!CONTYPE = LConType
            End If
            RECGRID!DIMPORT = 0
            RECGRID!UserId = LUserId & ""
            RECGRID!CONTIME = Time
            LConNo = LConNo + 1
            RECGRID!SRNO = LConNo 'RECGRID.AbsolutePosition
            RECGRID.Update
        End If
        DataGrid1.LeftCol = 0
        DataGrid1.Col = 0
    ElseIf KeyCode = 114 Then   'F3  NEW PARTY
        GETACNT.Show
        GETACNT.ZOrder
        GETACNT.add_record
    ElseIf KeyCode = 118 Then   ''F7 KEY
        RNO = InputBox("Enter the row number.", "Sauda")
        If Val(RNO) > Val(0) Then
            RECGRID.MoveFirst
            RECGRID.Find "SRNO=" & Val(RNO) & "", , adSearchForward
            If RECGRID.EOF Then
                MsgBox "Record not found.", vbCritical, "Error"
                RECGRID.MoveFirst
            End If
            DataGrid1.Col = 1
            DataGrid1.Enabled = True
            DataGrid1.SetFocus
        End If
    ElseIf KeyCode = 46 And Shift = 2 Then
        RECGRID.Delete
        If RECGRID.RecordCount = 0 Then
            RECGRID.AddNew
            LConNo = LConNo + 1
            RECGRID!SRNO = LConNo 'RECGRID.RecordCount
            If Combo1.ListIndex = Val(1) Then
                RECGRID!BRate = Val(Text3.Text)
                RECGRID!SRate = Val(Text3.Text)
                RECGRID!UserId = LUserId
            End If
            RECGRID.Update
        End If
        Call DataGrid1_AfterColEdit(0)
    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 3) Then
        On Error GoTo Error1
        If DataGrid1.Text = "" Then
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        End If
Error1:
    ElseIf KeyCode = 27 Then
        KeyCode = 0
    ElseIf ColIndex = 2 Then
        
        LConType = UCase(DataGrid1.Text)
    End If

End Sub
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If (KeyAscii = 98 Or KeyAscii = 66 Or KeyAscii = 83 Or KeyAscii = 115) And DataGrid1.Col = 2 Then
    Else
        If KeyAscii = 13 Then
        Else
            If DataGrid1.Col = 2 Then
                DataGrid1.Text = LConType
                DataGrid1.Enabled = True
                DataGrid1.SetFocus
            End If
        End If
    End If
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
    Call CANCEL_REC
'----------
    LblNew.Visible = False: LblEdit.Visible = False: LblCancel.Visible = False: LblDelete.Visible = False: LblSave.Visible = False: LblExit.Visible = False
    Set DataGrid1.DataSource = RECGRID: DataGrid1.Refresh
    Call ClearFormFn(GETContBS)
    Frame1.Enabled = False
'--------
    DataCombo3.Top = Val(2055): DataCombo3.Left = Val(1910)
    MYSQL = "SELECT ITEMCODE, ITEMCODE+','+ITEMNAME AS ITEMNAME,Lot FROM ITEMMAST WHERE COMPCODE=" & MC_CODE & " ORDER BY ITEMCODE"
    Set REC = Nothing: Set REC = New ADODB.Recordset: REC.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
    If Not REC.EOF Then
        Set DataCombo2.RowSource = REC: DataCombo2.BoundColumn = "ITEMCODE": DataCombo2.ListField = "ITEMNAME"
                
        QACC_CHANGE = False: Set REC_ACCOUNT = Nothing: Set REC_ACCOUNT = New ADODB.Recordset
        REC_ACCOUNT.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & MC_CODE & " AND gcode in (12,14) ORDER BY NAME ", cnn, adOpenKeyset, adLockReadOnly
        If Not REC_ACCOUNT.EOF Then Set DataCombo3.RowSource = REC_ACCOUNT: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
        
        Adodc1.ConnectionString = cnn: Adodc1.RecordSource = "SELECT * FROM CTR_M where COMPCODE=" & MC_CODE & " ORDER BY CONSNO": Adodc1.Refresh
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    Else
        Call Get_Selection(12)
    End If
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    If QACC_CHANGE Then
        QACC_CHANGE = False: Set REC_ACCOUNT = Nothing
        Set REC_ACCOUNT = New ADODB.Recordset
        REC_ACCOUNT.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & MC_CODE & " AND gcode in (12,14) ORDER BY NAME ", cnn, adOpenKeyset, adLockReadOnly
        If Not REC_ACCOUNT.EOF Then
            Set DataCombo3.RowSource = REC_ACCOUNT
            DataCombo3.BoundColumn = "AC_CODE"
            DataCombo3.ListField = "NAME"
        Else
            MsgBox "Please create customer account", vbInformation
            Call Get_Selection(12)
        End If
    End If
    If fb_press > 0 Then Call Get_Selection(fb_press)
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then FrmSauda.Show
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    FLOWDIR = 1
    If Len(Trim(Text2.Text)) < 1 Then
        DataCombo1.SetFocus
    Else
        If Not GetCloseRate Then Text2.Text = "": DataCombo1.SetFocus
    End If
End Sub
Private Sub Text3_GotFocus()
    FLOWDIR = 0: Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Sub RECSET()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "SCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "BCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "BNAME", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "BQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BRATE", adDouble, , adFldIsNullable

    RECGRID.Fields.Append "InvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "DImport", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "CONTIME", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "UserId", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "CLCODE", adVarChar, 30, adFldIsNullable

    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    RECGRID.AddNew
    RECGRID!CONTYPE = "B"
    RECGRID!DIMPORT = 0
    RECGRID!CONTIME = Time
    RECGRID!UserId = LUserId
    RECGRID.Update
    
    LConNo = LConNo + 1
    RECGRID!SRNO = LConNo  'RECGRID.AbsolutePosition
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
    cnn.Execute "DELETE FROM VCHAMT  WHERE COMPCODE=" & MC_CODE & " AND VOU_NO='" & VOU_NO & "'"
    cnn.Execute "DELETE FROM VOUCHER WHERE COMPCODE=" & MC_CODE & " AND VOU_NO='" & VOU_NO & "'"
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
    Case 13
        If OLDVAL = -1 Then
            RECGRID.Sort = "UserId DESC"
        Else
            RECGRID.Sort = "UserId"
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
Private Sub vcDTP1_Validate(Cancel As Boolean)
    Set REC_SAUDA = Nothing
    Set REC_SAUDA = New ADODB.Recordset
    REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE COMPCODE=" & MC_CODE & " AND MATURITY >=  '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'   ORDER BY itemcode,MATURITY", cnn, adOpenKeyset, adLockReadOnly
    If Not REC_SAUDA.EOF Then
        Set DataCombo1.RowSource = REC_SAUDA
        DataCombo1.BoundColumn = "SAUDACODE"
        DataCombo1.ListField = "SAUDANAME"
    End If
End Sub
Function GetCloseRate() As Boolean
     Set REC_SAUDA = Nothing: Set REC_SAUDA = New ADODB.Recordset
     REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE COMPCODE=" & MC_CODE & " AND SAUDACODE='" & Text2.Text & "' AND MATURITY>= '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", cnn, adOpenForwardOnly, adLockReadOnly
     If REC_SAUDA.EOF Then
         MsgBox "Invalid SAUDA code.", vbExclamation, "Error"
         GetCloseRate = False
     Else
         GetCloseRate = True
         Set REC_CloRate = Nothing: Set REC_CloRate = New ADODB.Recordset
         REC_CloRate.Open "SELECT CloseRate,DataImport FROM CTR_R WHERE COMPCODE=" & MC_CODE & " AND SAUDA='" & Text2.Text & "' AND CONDATE =  '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", cnn, adOpenForwardOnly, adLockReadOnly
         If Not REC_CloRate.EOF Then
            Text3.Text = Format(REC_CloRate!CLOSERATE, "0.00")
         End If
         Text2.Text = REC_SAUDA!SAUDACODE
         DataCombo1.BoundText = CStr(Text2.Text)
         DataCombo2.BoundText = REC_SAUDA!ItemCode
    End If
End Function
Public Sub lblcancel_Click()
    Call GETMAIN.ButtonClick(5)
    LblNew.Visible = True: LblEdit.Visible = True: LblDelete.Visible = True
    LblSave.Visible = False: LblCancel.Visible = False
End Sub
Public Sub lblsave_Click()
    Call GETMAIN.ButtonClick(4)
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
End Sub
