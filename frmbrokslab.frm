VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form frmbrokslab 
   BackColor       =   &H00808080&
   Caption         =   "Brokerage Slab"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15165
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   1695
      Begin VB.Label Label2 
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
         TabIndex        =   22
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   11295
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   420
         Left            =   2040
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   741
         _Version        =   393216
         ForeColor       =   64
         Text            =   "DataCombo3"
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
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   6240
         Width           =   11055
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Left            =   4800
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   3240
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   11055
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            ItemData        =   "frmbrokslab.frx":0000
            Left            =   4320
            List            =   "frmbrokslab.frx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   135
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            Left            =   4320
            TabIndex        =   4
            Text            =   "Text3"
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text2"
            Top             =   600
            Width           =   1575
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   360
            Left            =   1680
            TabIndex        =   0
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37860.8625462963
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   360
            Left            =   7200
            TabIndex        =   2
            Top             =   135
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ForeColor       =   64
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exchange"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   0
            Left            =   5880
            TabIndex        =   18
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Period (in Days)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   645
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scheme"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   18
            Left            =   3480
            TabIndex        =   12
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   4
            Left            =   3480
            TabIndex        =   11
            Top             =   645
            Width           =   750
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   158
            Width           =   930
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4740
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   8361
         _Version        =   393216
         AllowArrows     =   -1  'True
         BackColor       =   12640511
         ForeColor       =   4194368
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "SRNO"
            Caption         =   "SNo"
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
            Caption         =   "Party"
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
            Caption         =   "Party Name"
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
            DataField       =   "Amount"
            Caption         =   "Amount"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "StartDate"
            Caption         =   "Start Date"
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
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "SRATE"
            Caption         =   "Rate"
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
            Caption         =   "Contime"
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
            DataField       =   "DiffAmt"
            Caption         =   "Diff Amt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column14 
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
            DataField       =   "ORDER_NO"
            Caption         =   "Order No"
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
         BeginProperty Column18 
            DataField       =   "TRADE_NO"
            Caption         =   "Trade No"
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
               ColumnWidth     =   3300.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Object.Visible         =   0   'False
               ColumnWidth     =   1200.189
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
               Object.Visible         =   0   'False
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column13 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
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
      TabIndex        =   6
      Top             =   0
      Width           =   11775
      Begin VB.Frame Frame7 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   3240
         TabIndex        =   19
         Top             =   120
         Width           =   3975
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "Brokerage Rate Slab Setup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
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
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7260
      Left            =   120
      Top             =   1080
      Width           =   11565
   End
End
Attribute VB_Name = "frmbrokslab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean:        Dim LParty As String
Dim LParties As String:     Dim MVchno  As String
Dim MVou_No As String:      Dim LUserId As String
Dim LContractAcc As String: Dim LExchangeCode As String
Dim MExCode As String:      Dim MExBrokAC As String
Dim MExSrvAc As String:     Dim LSrvRate As Double
Dim VchNo  As String:       Dim LBConfirm As Integer
Dim LSConfirm As Integer:   Dim LPeriod As Integer
Dim LConNo As Long:         Dim LConSno As Long
Dim OldDate As Date:        Dim LDate  As Date
Dim FLOWDIR As Byte:        Dim LDataImport As Byte
Dim GRIDPOS As Byte:        Public Fb_Press As Byte
Dim RecEx As ADODB.Recordset:        Dim TRec As ADODB.Recordset
Dim RECGRID As ADODB.Recordset:
Dim Rec_Account  As ADODB.Recordset:
Dim LAmount As Double
Sub Add_Rec()
    If Fb_Press = 2 Then
        Text2.Enabled = False
        Text3.Enabled = False
    End If
    If Rec_Account.RecordCount > 0 Then
        LDataImport = 0
        Frame1.Enabled = True: Combo1.ListIndex = 0
        Call Get_Selection(1)
        If vcDTP1.Enabled Then vcDTP1.SetFocus
    Else
        Call CANCEL_REC
    End If
End Sub
Sub Save_Rec()
    Dim LExCode As String
    On Error GoTo ERR1
    RECGRID.Sort = "SRNO"
    If RECGRID.RecordCount > 0 Then
        Cnn.BeginTrans
        CNNERR = True
        ConSno = 1
        MYSQL = "SELECT VOU_NO FROM RATESLAB WHERE COMPCODE =" & GCompCode & " AND UPTOSTDT ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND SLABNO  =" & Combo1.ListIndex & " AND EXCODE ='" & MExCode & "'"
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        If Not TRec.EOF Then
            Do While Not TRec.EOF
                If TRec!VOU_NO <> "" Then
                    MYSQL = "DELETE FROM VOUCHER WHERE COMPCODE =" & GCompCode & "  AND VOU_NO='" & TRec!VOU_NO & "'"
                    Cnn.Execute MYSQL
                End If
                TRec.MoveNext
            Loop
        End If
        Cnn.Execute "DELETE FROM RATESLAB WHERE COMPCODE =" & GCompCode & " AND UPTOSTDT ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND SLABNO  =" & Combo1.ListIndex & " AND EXCODE ='" & MExCode & "'"
        Dim RC As ADODB.Recordset
        LParties = vbNullString
        RECGRID.MoveFirst
        Do While Not RECGRID.EOF
            MVou_No = ""
            If Not IsNull(RECGRID!BCODE) Then
                If RECGRID!BCODE <> "" Then
                    LParty = RECGRID!BCODE
                    If InStr(LParties, "'" & LParty & "'") < 1 Then
                        If LenB(LParties) <> 0 Then LParties = LParties & ","
                        LParties = LParties & "'" & LParty & "'"
                    End If
                    If Combo1.ListIndex = 0 Then 'Monthly
                        LPeriod = 30
                    ElseIf Combo1.ListIndex = 1 Then ' Quarterly
                        LPeriod = 90
                    ElseIf Combo1.ListIndex = 0 Then ' Half yearly
                        LPeriod = 180
                    ElseIf Combo1.ListIndex = 0 Then ' Yearly
                        LPeriod = 360
                    End If
                    LAmount = RECGRID!AMOUNT
                    
                    LDate = RECGRID!StartDate
                    Set TRec = Nothing:        Set TRec = New ADODB.Recordset
                    MYSQL = "SELECT SERVICETAX FROM EXTAX WHERE COMPCODE =" & GCompCode & " AND EXCHANGECODE  ='" & MExCode & "' AND FROMDT<='" & Format(LDate, "YYYY/MM/DD") & "' AND TODT>='" & Format(LDate, "YYYY/MM/DD") & "'"
                    TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
                    If Not TRec.EOF Then
                        LSrvRate = TRec!servicetax
                    Else
                        LSrvRate = 0
                    End If
                    LSrvTaxAmt = LAmount * (LSrvRate / 100)
                    MVchno = Get_VouNo("BRSH", GFinYear)
                    MNARATION = "Brokerage Debited for " & Combo1.text & " Scheme for " & MExCode & " Exchange "
                    MNARATION2 = "Brokerage Credited for " & Combo1.text & " Scheme for " & MExCode & " Exchange of " & RECGRID!BNAME & ""
                    Call PInsert_Voucher(MVchno, LDate, "C", "P", Val(RECGRID!SrNo), "ADD", RECGRID!BCODE, 0, vbNullString, MExCode, "0")
                    'MYSQL = "EXEC INSERT_VOUCHER " & GCompCode  & ",'" & MVchno & "','C','" & Format(LDate, "yyyy/MM/dd") & "','','" & RECGRID!SrNo & "','" & Format(LDate, "yyyy/MM/dd") & "','" & GUserName & "','" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD','','" & RECGRID!BCODE & "',0,'" & Format(Date, "yyyy/MM/dd") & "','" & MExCode & "'"
                    'Cnn.Execute MYSQL
                    MYSQL = "EXEC INSERT_VCHAMT " & GCompCode & ",'" & MVchno & "','C','" & Format(LDate, "yyyy/MM/dd") & "','D','" & LParty & "'," & Abs(LAmount) & ",'','','" & MNARATION & "','','" & MExCode & "',0,'','" & MExCode & "'"
                    Cnn.Execute MYSQL
                    MYSQL = "EXEC INSERT_VCHAMT " & GCompCode & ",'" & MVchno & "','C','" & Format(LDate, "yyyy/MM/dd") & "','C','" & MExBrokAC & "'," & Abs(LAmount) & ",'','','" & MNARATION2 & "','','" & MExCode & "',0,'','" & MExCode & "'"
                    Cnn.Execute MYSQL
                    
                    MNARATION = "Service Tax Debited for " & Combo1.text & " Scheme for " & MExCode & " Exchange "
                    MNARATION2 = "Service Tax Credited for " & Combo1.text & " Scheme for " & MExCode & " Exchange of " & RECGRID!BNAME & ""
                    MYSQL = "EXEC INSERT_VCHAMT " & GCompCode & ",'" & MVchno & "','C','" & Format(LDate, "yyyy/MM/dd") & "','D','" & LParty & "'," & Abs(LSrvTaxAmt) & ",'','','" & MNARATION & "','','" & MExCode & "',0,'','" & MExCode & "'"
                    Cnn.Execute MYSQL
                    MYSQL = "EXEC INSERT_VCHAMT " & GCompCode & ",'" & MVchno & "','C','" & Format(LDate, "yyyy/MM/dd") & "','C','" & MExSrvAc & "'," & Abs(LSrvTaxAmt) & ",'','','" & MNARATION2 & "','','" & MExCode & "',0,'','" & MExCode & "'"
                    Cnn.Execute MYSQL
                    
                    MYSQL = "INSERT INTO RATESLAB (COMPCODE,SLABNO,PARTY,PERIOD,AMOUNT,UPTOSTDT,VOU_NO,EXCODE) VALUES "
                    MYSQL = MYSQL & "( " & GCompCode & "," & Combo1.ListIndex & ",'" & RECGRID!BCODE & "'," & Val(LPeriod) & "," & RECGRID!AMOUNT & ",'" & Format(RECGRID!StartDate, "YYYY/MM/DD") & "','" & MVchno & "','" & MExCode & "')"
                    Cnn.Execute MYSQL
                End If
            End If
            RECGRID.MoveNext
        Loop
        'cnn.BeginTrans
        MExCode = "'" & MExCode & "'"
        If GSrvYN = "Y" Then Call Updt_SrvTax(LParties, MExCode, vbNullString, vcDTP1.Value, GFinEnd)
        'Call UpdateBrokRateType(LParties, "", vcDTP1.Value, CDate(GFinEnd), "", MExCode)
        
        Cnn.CommitTrans
        
        DoEvents: Cnn.BeginTrans
        If BILL_GENERATION(vcDTP1.Value, CDate(GFinEnd), "", LParties, MExCode) Then
            Cnn.CommitTrans
            CNNERR = False
        Else
            Cnn.RollbackTrans
            CNNERR = False
        End If
        Call Chk_Billing
    End If
    Call CANCEL_REC
    Exit Sub
ERR1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
       'Resume
       Cnn.RollbackTrans: CNNERR = False
    End If
End Sub
Sub CANCEL_REC()
    vcDTP1.Enabled = True: Text2.Enabled = True: Combo1.Enabled = True: Text3.Enabled = True
    vcDTP1.Value = Now
    LConNo = 10000
    Text2.text = vbNullString
    Text3.text = vbNullString
    Call RecSet
    Fb_Press = 0
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.Refresh
    DataCombo3.Visible = False
    
    Call Get_Selection(10)
    Combo1.ListIndex = 0: Frame1.Enabled = False
End Sub
Function MODIFY_REC(LCondate As Date, LSAUDA As Integer) As Boolean
    Dim LREC As ADODB.Recordset
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    MYSQL = "SELECT * FROM RATESLAB  WHERE COMPCODE =" & GCompCode & " AND UPTOSTDT ='" & Format(LCondate, "yyyy/MM/dd") & "' AND SLABNO=" & LSAUDA & "AND EXCODE ='" & MExCode & "'"
    Rec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rec.EOF Then
        If Fb_Press = 2 Then
            MsgBox "Transaction does not exists for the Selected creteria?", vbExclamation
            OldDate = vcDTP1.Value
            GETCont.Fb_Press = 1
            vcDTP1.Value = OldDate
            MODIFY_REC = True
            Exit Function
        ElseIf Fb_Press = 1 Then
            MODIFY_REC = True
        End If
        Exit Function
    Else
        If Fb_Press = 1 Then
            Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
            MYSQL = "SELECT * FROM RATESLAB WHERE COMPCODE =" & GCompCode & " AND uptostdt='" & Format(LCondate, "yyyy/MM/dd") & "' AND SLABNO = " & LSAUDA & " and EXCODE ='" & MExCode & "'"
            GeneralRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                MsgBox "Entries already exists.Please Press Enter to Modify Entries.", vbInformation
                OldDate = vcDTP1.Value
                GETCont.Fb_Press = 2
                vcDTP1.Value = OldDate
                GETMAIN.StatusBar1.Panels(2).text = "Modify Record"
                MODIFY_REC = False
                Exit Function
            Else
                MODIFY_REC = True
                Exit Function
            End If
        Else
            Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
            MYSQL = "SELECT * FROM RATESLAB WHERE COMPCODE =" & GCompCode & " AND UPTOSTDT ='" & Format(LCondate, "yyyy/MM/dd") & "' AND SLABNO = " & LSAUDA & " and EXCODE ='" & MExCode & "'"
            GeneralRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                MODIFY_REC = True
            Else
                MsgBox "Entries does not exists.Please Enter to Add New Entries.", vbInformation
                OldDate = vcDTP1.Value
                Call CANCEL_REC
                GETCont.Fb_Press = 1
                vcDTP1.Value = OldDate
                GETCont.Add_Rec
                GETMAIN.StatusBar1.Panels(2).text = "Add Record"
                MODIFY_REC = False
                Exit Function
            End If
        End If
    End If
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    MYSQL = "SELECT A.SLABNO,A.UPTOSTDT,A.PERIOD,A.AMOUNT,A.PARTY,A.VOU_NO,B.NAME FROM RATESLAB AS A, ACCOUNTD AS B  WHERE A.COMPCODE =" & GCompCode & " AND A.UPTOSTDT ='" & Format(LCondate, "yyyy/MM/dd") & "' AND A.SLABNO=" & LSAUDA & " AND A.EXCODE ='" & MExCode & "'"
    MYSQL = MYSQL & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
    MYSQL = MYSQL & " ORDER BY A.PARTY,A.UPTOSTDT"
    Rec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    Call RecSet
    RECGRID.Delete
    LParty = vbNullString
    LConNo = 0
    Text2.Enabled = True
    Text3.Enabled = True
    If Not Rec.EOF Then
        Rec.MoveFirst
        Text2.text = Rec!Period
        Text3.text = Rec!AMOUNT
        Text3.text = Format(Text3.text, "0.00")
        Text2.Enabled = False
        Text3.Enabled = False
    End If
    Do While Not Rec.EOF
        LConNo = LConNo + 1
        RECGRID.AddNew
        
        RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
        RECGRID!BCODE = Rec!PARTY & ""
        RECGRID!BNAME = Rec!NAME
        RECGRID!AMOUNT = Rec!AMOUNT
        RECGRID!StartDate = Rec!UPTOSTDT
        RECGRID.Update
        Rec.MoveNext
    Loop
    
    Set DataGrid1.DataSource = RECGRID
    Call DataGrid1_AfterColEdit(0)
    If Fb_Press = 3 Then
        If MsgBox("You are about to Delete all Entries. Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
            On Error GoTo ERR1
            Cnn.BeginTrans
            CNNERR = True
            MYSQL = "DELETE FROM RATESLAB WHERE COMPCODE=" & GCompCode & " AND SLABNO =" & LSAUDA & " AND UPTOSTDT='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
            Cnn.Execute MYSQL
            Cnn.CommitTrans
            
            ''REGENERATING UPTO HERE
            If Fb_Press = 2 Then
                GETMAIN.Toolbar1_Buttons(4).Enabled = True: GETMAIN.saverec.Enabled = True
            ElseIf Fb_Press = 3 Then
                Call CANCEL_REC
            End If
            MODIFY_REC = True
            Exit Function
ERR1:
            If err.Number <> 0 Then
                MsgBox err.Number, vbCritical, "Error Number : " & err.Number
            End If
            If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
            Call CANCEL_REC
        End If
    End If
End Function


Private Sub Combo1_GotFocus()
 Sendkeys "%{DOWN}"
End Sub

Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    MYSQL = "SELECT * FROM RATESLAB WHERE COMPCODE =" & GCompCode & "  AND UPTOSTDT='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND SLABNO=" & Combo1.ListIndex & ""
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not TRec.EOF And Fb_Press = 1 Then
        If MODIFY_REC(vcDTP1.Value, Combo1.ListIndex) Then
        Else
            Cancel = True
        End If
    End If
    If Combo1.ListIndex = 0 Then 'Monthly
        Text2.text = 30
    ElseIf Combo1.ListIndex = 1 Then ' Quarterly
        Text2.text = 90
    ElseIf Combo1.ListIndex = 0 Then ' Half yearly
        Text2.text = 180
    ElseIf Combo1.ListIndex = 0 Then ' Yearly
        Text2.text = 360
    End If
    
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    If DataCombo1.BoundText = "" Then
        MsgBox "Please Select Exchange "
        Cancel = True
    Else
        MExCode = DataCombo1.BoundText
        Set TRec = Nothing:        Set TRec = New ADODB.Recordset
        MYSQL = "SELECT BROKAC,SRVTAXACC FROM EXMAST WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & MExCode & "'"
        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        If Not TRec.EOF Then
            MExBrokAC = TRec!BROKAC
            MExSrvAc = SrvTaxACC
        Else
            MsgBox "Invalid ExCode "
            Cancel = True
        End If
    End If
    If MODIFY_REC(vcDTP1.Value, Combo1.ListIndex) Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub DataCombo2_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_GotFocus()
    Sendkeys "%{DOWN}"
    If DataGrid1.Col = 1 Or DataGrid1.Col = 2 Then
        DataGrid1.text = ""
        DataGrid1.Col = 1
    '    Label2.Visible = True: Label2.Left = 1080
        DataCombo3.Left = Val(1080)
        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    ElseIf DataGrid1.Col = 5 Or DataGrid1.Col = 6 Or DataGrid1.Col = 7 Then
        DataGrid1.text = ""
        DataGrid1.Col = 5: DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
        DataCombo3.Left = Val(7200)
     '   Label2.Visible = True: Label2.Left = 7200
    End If
    Sendkeys "%{DOWN}"
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If DataGrid1.Col = 1 Then
            RECGRID!BCODE = DataCombo3.BoundText
            RECGRID!BNAME = DataCombo3.text
            DataGrid1.Col = 2
        ElseIf DataGrid1.Col = 5 Or DataGrid1.Col = 6 Or DataGrid1.Col = 7 Then
            RECGRID!scode = DataCombo3.BoundText
            RECGRID!SNAME = DataCombo3.text
            DataGrid1.Col = 6
        End If
        DataGrid1.SetFocus
        DataCombo3.Visible = False
    ElseIf KeyCode = 27 Then
        DataGrid1.SetFocus
        DataCombo3.Visible = False
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
Public Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = Val(1) Then
        Rec_Account.MoveFirst
        Rec_Account.Find "AC_CODE='" & DataGrid1.text & "'", , adSearchForward
        If Not Rec_Account.EOF Then
            If ColIndex = Val(1) Then
                DataGrid1.Col = 2
                RECGRID!BCODE = Rec_Account!AC_CODE
                RECGRID!BNAME = Rec_Account!NAME
            End If
        Else
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        End If
    ElseIf ColIndex = 3 Then
        ''IF CONTRACT THEN ONLY CHANGE OCCURS
        If Val(RECGRID!AMOUNT & "") <= 0 Then
            MsgBox "Amount not be zero.Please enter Amount", vbCritical
            DataGrid1.Col = 4: DataGrid1.SetFocus
        End If
    ElseIf ColIndex = 4 Then
        'RECGRID.AddNew
        'LConNo = LConNo + 1
        'RECGRID!SRNO = Val(LConNo)
        'RECGRID!AMOUNT = Val(Text3.Text)
        'RECGRID!STARTDATE = vcDTP1.Value
        'RECGRID.Update
        DataGrid1.Col = 1: DataGrid1.SetFocus
    End If
    Set Rec = Nothing: Set Rec = New ADODB.Recordset: Set Rec = RECGRID.Clone
    BQnty = 0: SQnty = 0: BAmt = 0: SAmt = 0
    
End Sub
Private Sub DataGrid1_GotFocus()
    vcDTP1.Enabled = False
    Text2.Enabled = False
    'DataCombo1.Enabled = False
    Combo1.Enabled = False
    'DataCombo4.Enabled = False
    Text3.Enabled = False
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And DataGrid1.Col = 4 Then
        BCODE = RECGRID!BCODE
        BNAME = RECGRID!BNAME
        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            RECGRID!AMOUNT = Val(Text3.text)
            RECGRID!StartDate = vcDTP1.Value
            LConNo = LConNo + 1
            RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
            RECGRID.Update
        End If
        DataGrid1.LeftCol = 0
        DataGrid1.Col = 0
        Call DataGrid1_AfterColEdit(0)
    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 1 Or DataGrid1.Col = 5) Then
        If Len(Trim(DataGrid1.text)) < 1 Then
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        Else
        End If
    ElseIf KeyCode = 27 Then
        KeyCode = 0
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "vcDTP1" Then
            Sendkeys "{tab}"
        End If
    End If
End Sub
Private Sub Form_Load()
    Call CANCEL_REC
'----------
    Set DataGrid1.DataSource = RECGRID: DataGrid1.Refresh
    Frame1.Enabled = False
'--------
    LDataImport = 0
    DataCombo3.Top = Val(2055): DataCombo3.Left = Val(1080)
    QACC_CHANGE = False: Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
    Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Account.EOF Then
        Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    End If
    Set RecEx = Nothing
    Set RecEx = New ADODB.Recordset
    MYSQL = "SELECT EXCODE,EXNAME,BROKAC FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME"
    RecEx.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not RecEx.EOF Then
        Set DataCombo1.RowSource = RecEx: DataCombo1.BoundColumn = "EXCODE": DataCombo1.ListField = "EXNAME"
    End If
    
    
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    If QACC_CHANGE Then
        QACC_CHANGE = False: Set Rec_Account = Nothing
        Set Rec_Account = New ADODB.Recordset
        Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
        If Not Rec_Account.EOF Then
            Set DataCombo3.RowSource = Rec_Account
            DataCombo3.BoundColumn = "AC_CODE"
            DataCombo3.ListField = "NAME"
        Else
            MsgBox "Please create customer account", vbInformation
            Call Get_Selection(12)
        End If
    End If
    If Fb_Press > 0 Then Call Get_Selection(Fb_Press)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
    Text11.text = Format(Text11, text, "0.0000")
End Sub

Private Sub Text18_Validate(Cancel As Boolean)
    LPassword = EncryptNEW(Text18.text, 13)
    If LPassword = GRegNo2 Then
        MsgBox ""
    Else
        Cancel = True
    End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)

If GRegNo2 = EncryptNEW(Text12.text, 13) Then
    Combo1.Enabled = True
    Combo1.SetFocus
Else
    MsgBox "Invalid Password No Modificatiobn Allowed"
    Cancel = True
End If
End Sub

Private Sub Text2_GotFocus()
    Text2.SelLength = Len(Text2.text)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Text2.text <> "" Then
    If Val(Text2.text) < 1 Then
        MsgBox "Period can not be less than 1 Day"
        Cancel = True
    End If
Else
    MsgBox "Period can not be blank"
    Cancel = True
End If
End Sub
Private Sub Text3_GotFocus()
    Text3.SelLength = Len(Text3.text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
Text3.text = Format(Text3.text, "0.00")
If Text3.text <> "" Then
    If Val(Text3.text) < 0 Then
        MsgBox "Amount can not be less than 0"
        Cancel = True
    End If
Else
    MsgBox "Amount can not be blank"
    Cancel = True
End If
If Val(Text3.text) > 0 Then
    RECGRID!AMOUNT = Val(Text3.text)
    RECGRID.Update
End If
End Sub


Sub RecSet()
    
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "BNAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "Amount", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "StartDate", adDate, , adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    RECGRID.AddNew
    RECGRID!StartDate = vcDTP1.Value
    RECGRID!AMOUNT = Val(Text3.text)
    RECGRID.Update
    LConNo = LConNo + 1
    RECGRID!SrNo = LConNo  'RECGRID.AbsolutePosition
    DataGrid1.Col = 1
End Sub
Sub Delete_Voucher(VOU_NO As String)
    Cnn.Execute "DELETE FROM VOUCHER WHERE COMPCODE=" & GCompCode & " AND VOU_NO='" & VOU_NO & "'"
    Cnn.Execute "DELETE FROM VCHAMT  WHERE COMPCODE=" & GCompCode & " AND VOU_NO='" & VOU_NO & "'"
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

    End Select
    
    
End Sub


Private Sub vcDTP1_Validate(Cancel As Boolean)
    If SYSTEMLOCK(DateValue(vcDTP1.Value)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
        Cancel = True
    End If
End Sub
