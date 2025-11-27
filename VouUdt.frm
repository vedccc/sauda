VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form VouUdt 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "VouUdt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1695
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
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
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12135
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   495
         Left            =   5400
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   7260
      Left            =   195
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12806
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Voucher Detail"
      TabPicture(0)   =   "VouUdt.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11415
         Begin TabDlg.SSTab SSTab1 
            Height          =   2175
            Left            =   3960
            TabIndex        =   4
            Top             =   1800
            Visible         =   0   'False
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   3836
            _Version        =   393216
            Tabs            =   1
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Voucher"
            TabPicture(0)   =   "VouUdt.frx":0028
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label8"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label3"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Shape1"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "DTPicker1"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "CANCEL_CMD"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Command3"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "VOU_TYPE"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "pr_frame"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            Begin VB.Frame pr_frame 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1800
               TabIndex        =   8
               Top             =   960
               Visible         =   0   'False
               Width           =   2535
               Begin VB.OptionButton Rpt_opn 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF8080&
                  Caption         =   "Receipt"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   1320
                  TabIndex        =   10
                  Top             =   120
                  Width           =   1095
               End
               Begin VB.OptionButton pmt_opn 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF8080&
                  Caption         =   "Payment"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   120
                  TabIndex        =   9
                  Top             =   120
                  Width           =   1095
               End
            End
            Begin VB.ComboBox VOU_TYPE 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   345
               ItemData        =   "VouUdt.frx":0044
               Left            =   1800
               List            =   "VouUdt.frx":0054
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   600
               Width           =   1695
            End
            Begin VB.CommandButton Command3 
               Caption         =   "O&K"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   2640
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   0
               Width           =   855
            End
            Begin VB.CommandButton CANCEL_CMD 
               Caption         =   "Cance&l"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   3600
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   0
               Width           =   855
            End
            Begin vcDateTimePicker.vcDTP DTPicker1 
               Height          =   375
               Left            =   1800
               TabIndex        =   11
               Top             =   1440
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   37680
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   1  'Opaque
               Height          =   1575
               Left            =   120
               Top             =   480
               Width           =   4335
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Voucher Date"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   13
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Voucher Type"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   360
               TabIndex        =   12
               Top             =   600
               Width           =   1455
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid_detail 
            Bindings        =   "VouUdt.frx":0075
            Height          =   3975
            Left            =   120
            TabIndex        =   3
            Top             =   2160
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483633
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
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
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "*******  Voucher Amount  *******"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "AC_NAME"
               Caption         =   "      Account Description"
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
               DataField       =   "DR_CR"
               Caption         =   "Dr/Cr"
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
               DataField       =   "AMOUNT"
               Caption         =   "     Amount"
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
            BeginProperty Column03 
               DataField       =   "NARRATION"
               Caption         =   "                          Narration"
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
               ScrollBars      =   3
               BeginProperty Column00 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   3960
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   675.213
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   1514.835
               EndProperty
               BeginProperty Column03 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   4229.858
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid_voucher 
            Bindings        =   "VouUdt.frx":008E
            Height          =   1815
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   3201
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483633
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
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
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "*******  Voucher  *******"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "VOU_NO"
               Caption         =   "                      Voucher No."
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
               DataField       =   "VOU_DT"
               Caption         =   "         Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "d.MMMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "CHEQUE_NO"
               Caption         =   "           Cheque No."
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
               DataField       =   "CHEQUE_DT"
               Caption         =   "    Cheque Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "d.MMMM yyyy"
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
               ScrollBars      =   2
               BeginProperty Column00 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   4155.024
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   1904.882
               EndProperty
               BeginProperty Column02 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   2520
               EndProperty
               BeginProperty Column03 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   1800
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSAdodcLib.Adodc Ado_detail1 
      Height          =   330
      Left            =   6720
      Top             =   8160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      Caption         =   "Ado_Detail"
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
   Begin MSAdodcLib.Adodc Ado_voucher1 
      Height          =   375
      Left            =   3840
      Top             =   8160
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Ado_Voucher1"
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
   Begin MSAdodcLib.Adodc Ado_Voucher 
      Height          =   375
      Left            =   960
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc Ado_detail 
      Height          =   330
      Left            =   960
      Top             =   3360
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
      Caption         =   "Adodc2"
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000040&
      BorderWidth     =   12
      Height          =   7500
      Left            =   75
      Top             =   1080
      Width           =   11925
   End
End
Attribute VB_Name = "VouUdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_Payrpt As String
Dim F_Voutype As String
Public Fb_Press As Byte
Dim VCHREC As ADODB.Recordset
Dim VouRec As ADODB.Recordset
Dim VouDRec As ADODB.Recordset
Private Sub CANCEL_CMD_Click()
    VouFrm.Fb_Press = 0
    Call Get_Selection(5)
    Unload Me
End Sub
Private Sub Command1_Click()
    MVou_No = DataGrid_voucher.Columns(0).text: VouFrm.Fb_Press = VouUdt.Fb_Press: MVou_Type = F_Voutype
    VouFrm.Show
    If VouFrm.VOUCHER_ACCESS(VouRec!VOU_NO) Then
        If VouUdt.Fb_Press = 2 Then
            VouFrm.Frame6.Enabled = True
            If VouFrm.TXT_NARR.Visible = True Then VouFrm.TXT_NARR.SetFocus
        ElseIf VouUdt.Fb_Press = 3 Then
            If MsgBox("  Confirm Delete ?        ", vbYesNo + vbQuestion, "Confirmation") = 6 Then
                Call VouFrm.Delete_Entry
            End If
            Call VouFrm.CLEAR_SCREEN
            Call Get_Selection(5)
        End If
    End If
    Unload VouUdt
End Sub
Private Sub Command2_Click()
    SSTab2.Visible = False: Command1.Visible = False: Command2.Visible = False
End Sub

Private Sub DataGrid_voucher_Click()
    mysql = "SELECT A.Ac_Name,Vt.Dr_Cr,Vt.Amount,Vt.Narration From Vchamt As Vt,Account As A Where  A.COMPCODE =" & GCompCode & " AND A.COMPCODE =VT.COMPCODE AND Vt.Vou_No='" & DataGrid_voucher.Columns(0).text & "' And A.Ac_Code=Vt.Ac_Code ORDER BY VT.EntSeq"
    Set VouDRec = Nothing: Set VouRec = New ADODB.Recordset
    VouDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
End Sub
Private Sub DataGrid_voucher_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        mysql = "SELECT A.Ac_Name,Vt.Dr_Cr,Vt.Amount,Vt.Narration From Vchamt As Vt,Account As A Where  A.COMPCODE =" & GCompCode & " AND A.COMPCODE =VT.COMPCODE AND Vt.Vou_No='" & DataGrid_voucher.Columns(0).text & "' And A.Ac_Code=Vt.Ac_Code ORDER BY VT.EntSeq"
        Set VouDRec = Nothing: Set VouRec = New ADODB.Recordset
        VouDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    End If
        
End Sub
Private Sub DataGrid_voucher_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    mysql = "SELECT A.Ac_Name,Vt.Dr_Cr,Vt.Amount,Vt.Narration From Vchamt As Vt,Account As A Where A.COMPCODE =" & GCompCode & " AND A.COMPCODE =VT.COMPCODE Vt.Vou_No='" & DataGrid_voucher.Columns(0).text & "' And A.Ac_Code=Vt.Ac_Code ORDER BY VT.EntSeq"
    Set VouDRec = Nothing: Set VouRec = New ADODB.Recordset
    VouDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        VouFrm.Fb_Press = 0
        Call Get_Selection(5)
        Unload Me
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Form_Load()
    SSTab2.ZOrder
    SSTab1.Visible = True: SSTab2.Visible = False: DTPicker1.MinDate = SysLock_Date: DTPicker1.MaxDate = GFinEnd: DTPicker1.Value = Now
    Command1.Visible = False: Command2.Visible = False: CANCEL_CMD.Visible = True
    VOU_TYPE.text = VOU_TYPE.List(0)
End Sub
Private Sub Form_Paint()
    GETMAIN.StatusBar1.Panels(1).text = "Voucher Details"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    GETMAIN.StatusBar1.Panels(1).text = vbNullString
End Sub
Private Sub Command3_Click()
    Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
    mysql = "SELECT V.VOU_NO,V.VOU_DT FROM VOUCHER AS V WHERE V.COMPCODE =" & GCompCode & " AND V.VOU_TYPE='" & F_Voutype & "' AND V.VOU_PR='" & F_Payrpt & "' AND V.VOU_DT='" & Format(DTPicker1.Value, "yyyy/MM/dd") & "'"
    VouRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    If VouRec.RecordCount <> 0 Then
        Set VouDRec = Nothing: Set VouRec = New ADODB.Recordset
        mysql = "SELECT A.AC_NAME,VT.DR_CR,VT.AMOUNT,VT.NARRATION FROM VCHAMT AS VT, ACCOUNT AS A WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =VT.COMPCODE AND VT.VOU_NO='" & Ado_Voucher.Recordset!VOU_NO & "' "
        mysql = mysql & " AND A.AC_CODE=VT.AC_CODE ORDER BY VT.ENTSEQ"
        VouDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        SSTab2.Visible = True: Command1.Visible = True: Command2.Visible = True: DataGrid_voucher.Col = 0: DataGrid_voucher.Row = 0: DataGrid_voucher.SetFocus
    ElseIf VouRec.RecordCount = 0 Then
        MsgBox "Transaction not exist.", vbInformation, "Message"
        CANCEL_CMD.SetFocus
    End If
End Sub

Private Sub pmt_opn_LostFocus()
    F_Payrpt = "PAYMENT"
End Sub
Private Sub Rpt_opn_LostFocus()
    F_Payrpt = "RECEIPT"
End Sub
Private Sub VOU_TYPE_Click()
    If VOU_TYPE.text = "Cash" Or VOU_TYPE.text = "Bank" Then
        pr_frame.Visible = True
    Else
        pr_frame.Visible = False
    End If
   Select Case VOU_TYPE.ListIndex
    Case 0
        F_Voutype = "CV": MSEL_OPT = 1
    Case 1
        F_Voutype = "BV": MSEL_OPT = 2
    Case 2
        F_Voutype = "JV": F_Payrpt = "0": MSEL_OPT = 3
    Case 3
        F_Voutype = "CN": F_Payrpt = "0": MSEL_OPT = 4
    Case 4
        F_Voutype = "DN": F_Payrpt = "0": MSEL_OPT = 5
    End Select
End Sub
Private Sub VOU_TYPE_GotFocus()
    Sendkeys "%{DOWN}"
'    Call LSendKeys_Down
End Sub
Private Sub VOU_TYPE_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 38) Or (KeyCode = 40) Then
        If VOU_TYPE.text = "Cash" Or VOU_TYPE.text = "Bank" Then
            pr_frame.Visible = True
        Else
            pr_frame.Visible = False
        End If
    End If
End Sub
Sub MakeRec()
    Set VCHREC = Nothing
    Set VCHREC = New ADODB.Recordset
    VCHREC.Fields.Append "VCHNO", adVarChar, 18, adFldIsNullable
    VCHREC.Fields.Append "VCHDT", adVarChar, 10, adFldIsNullable
    VCHREC.Fields.Append "AC_NAME", adVarChar, 65, adFldIsNullable
    VCHREC.Fields.Append "DRCR", adVarChar, 1, adFldIsNullable
    VCHREC.Fields.Append "CRAMT", adDouble, , adFldIsNullable
    VCHREC.Fields.Append "DRAMT", adDouble, , adFldIsNullable
    VCHREC.Fields.Append "CHQNO", adVarChar, 8, adFldIsNullable
    VCHREC.Fields.Append "CHQDT", adVarChar, 10, adFldIsNullable
    VCHREC.Fields.Append "NARR", adVarChar, 100, adFldIsNullable
    VCHREC.Open , , adOpenKeyset, adLockOptimistic
End Sub
