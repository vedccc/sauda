VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmcont4 
   Caption         =   "Contract Entry"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
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
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame15 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Caption         =   "Frame15"
      Height          =   855
      Left            =   7320
      TabIndex        =   114
      Top             =   4800
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label Label36 
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   115
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "ERGGHE"
      Height          =   4095
      Left            =   3120
      TabIndex        =   96
      Top             =   4200
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
         TabIndex        =   98
         Top             =   480
         Width           =   11895
      End
      Begin VB.CommandButton Command3 
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
         TabIndex        =   97
         ToolTipText     =   "Close"
         Top             =   -15
         Width           =   615
      End
      Begin VB.Label Label29 
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
         TabIndex        =   99
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4575
      Left            =   17400
      TabIndex        =   85
      Top             =   4680
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "Apply"
         Height          =   375
         Left            =   2640
         TabIndex        =   90
         Top             =   120
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3975
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7011
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   89
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4575
      Left            =   17400
      TabIndex        =   83
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   375
         Left            =   2640
         TabIndex        =   88
         Top             =   120
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7011
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SaudaCode"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda/Contract"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   87
         Top             =   120
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   68
      Top             =   2520
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Trade"
      TabPicture(0)   =   "frmcont4.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Standing Position"
      TabPicture(1)   =   "frmcont4.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   6255
         Left            =   240
         TabIndex        =   81
         Top             =   480
         Width           =   16335
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   6015
            Left            =   120
            TabIndex        =   82
            Top             =   120
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   10610
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   21
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
               Name            =   "Times New Roman"
               Size            =   11.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6255
         Left            =   -74880
         TabIndex        =   69
         Top             =   360
         Width           =   16695
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4935
            Left            =   240
            TabIndex        =   72
            Top             =   600
            Width           =   16455
            _ExtentX        =   29025
            _ExtentY        =   8705
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   21
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
               Name            =   "Times New Roman"
               Size            =   11.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Caption         =   "ERGGHE"
            Height          =   3375
            Left            =   5160
            TabIndex        =   105
            Top             =   840
            Visible         =   0   'False
            Width           =   6255
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "With high low"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1680
               TabIndex        =   111
               Top             =   1680
               Width           =   2415
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Without high low"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1680
               TabIndex        =   110
               Top             =   2160
               Value           =   -1  'True
               Width           =   2655
            End
            Begin VB.CommandButton Command6 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
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
               Left            =   3000
               Style           =   1  'Graphical
               TabIndex        =   109
               Top             =   2880
               Width           =   975
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
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
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label Label35 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   113
               Top             =   2160
               Visible         =   0   'False
               Width           =   4695
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Order#:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1680
               TabIndex        =   112
               Top             =   1080
               Width           =   4695
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Do you want to execute this order?"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   720
               TabIndex        =   107
               Top             =   600
               Width           =   4695
            End
            Begin VB.Shape Shape1 
               BackStyle       =   1  'Opaque
               Height          =   2175
               Left            =   0
               Top             =   480
               Width           =   6255
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Order execution"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   75
               TabIndex        =   106
               Top             =   120
               Width           =   1815
            End
         End
         Begin VB.ComboBox cmbtype 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "frmcont4.frx":0038
            Left            =   11520
            List            =   "frmcont4.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   120
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DComboBroker 
            Height          =   390
            Left            =   8340
            TabIndex        =   76
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DComboSauda 
            Height          =   390
            Left            =   4680
            TabIndex        =   75
            Top             =   120
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DComboParty 
            Height          =   390
            Left            =   1680
            TabIndex        =   74
            Top             =   120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox TxtDiffAmt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13800
            TabIndex        =   71
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Delete All Trades"
            Height          =   400
            Left            =   15120
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   120
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   4815
            Left            =   240
            TabIndex        =   73
            Top             =   600
            Width           =   16335
            _ExtentX        =   28813
            _ExtentY        =   8493
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   18
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
               Name            =   "Times New Roman"
               Size            =   9.75
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo DComboCode 
            Height          =   390
            Left            =   720
            TabIndex        =   95
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridOrder 
            Height          =   4815
            Left            =   240
            TabIndex        =   101
            Top             =   600
            Width           =   16455
            _ExtentX        =   29025
            _ExtentY        =   8493
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648447
            HeadLines       =   1
            RowHeight       =   21
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
               Name            =   "Times New Roman"
               Size            =   11.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10920
            TabIndex        =   103
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Broker"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   80
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   79
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Party"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Diff Amt"
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
            Height          =   255
            Left            =   12960
            TabIndex        =   77
            Top             =   180
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   6975
      Left            =   0
      TabIndex        =   67
      Top             =   2400
      Width           =   17295
   End
   Begin VB.TextBox TxtBrokerConfirm 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12480
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   66
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox TxtConfirm 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11160
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   65
      Top             =   9000
      Width           =   855
   End
   Begin VB.Frame Frame11 
      Caption         =   "Frame11"
      Height          =   975
      Left            =   1440
      TabIndex        =   59
      Top             =   9600
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox TxtFileType 
         Height          =   375
         Left            =   1200
         TabIndex        =   60
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2400
      Left            =   11880
      TabIndex        =   41
      Top             =   9480
      Visible         =   0   'False
      Width           =   9855
      Begin VB.TextBox TxtSaudaID 
         Height          =   735
         Left            =   3000
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox Text7 
         Height          =   330
         Left            =   5520
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtContime 
         Height          =   330
         Left            =   3240
         TabIndex        =   54
         Text            =   "Text8"
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtOrdNo 
         Height          =   330
         Left            =   600
         TabIndex        =   53
         Text            =   "OrdNo"
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtDataImport 
         Height          =   330
         Left            =   840
         TabIndex        =   52
         Text            =   "Text16"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtCalVal 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   510
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   50
         Width           =   975
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   5520
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2160
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   46
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   45
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   17295
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   16935
         Begin VB.Frame Frame13 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   100
            Top             =   120
            Width           =   2415
            Begin VB.OptionButton OptOrder 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Order"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   1
               Top             =   45
               Width           =   1335
            End
            Begin VB.OptionButton OptTrade 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Trade"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               TabIndex        =   0
               Top             =   45
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.ComboBox InstCombo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "frmcont4.frx":005C
            Left            =   7200
            List            =   "frmcont4.frx":005E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtAdminPass 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   15600
            PasswordChar    =   "*"
            TabIndex        =   63
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   10440
            TabIndex        =   6
            Top             =   120
            Width           =   1000
         End
         Begin VB.CommandButton CmdModify 
            Caption         =   "Modify"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   11520
            TabIndex        =   7
            Top             =   135
            Width           =   1000
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "Cancel"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12600
            TabIndex        =   8
            Top             =   120
            Width           =   1000
         End
         Begin VB.CheckBox ChkShowContract 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Show All Contracts"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4800
            TabIndex        =   3
            Top             =   80
            Width           =   2175
         End
         Begin vcDateTimePicker.vcDTP DtpCondate 
            Height          =   360
            Left            =   3240
            TabIndex        =   2
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   3
            Value           =   41160.4222453704
         End
         Begin MSDataListLib.DataCombo DComboExchange 
            Height          =   420
            Left            =   8880
            TabIndex        =   5
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   393216
            Style           =   2
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
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Ad Paswd"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   14280
            TabIndex        =   64
            Top             =   195
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Exch"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8280
            TabIndex        =   58
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   57
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1575
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   16935
         Begin VB.ComboBox CmbOrderStatus 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "frmcont4.frx":0060
            Left            =   13440
            List            =   "frmcont4.frx":0070
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Frame FrameOpt 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   635
            Left            =   120
            TabIndex        =   92
            Top             =   960
            Visible         =   0   'False
            Width           =   3975
            Begin VB.TextBox TxtOptType 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               MaxLength       =   2
               TabIndex        =   20
               Text            =   "CE"
               Top             =   80
               Width           =   855
            End
            Begin VB.TextBox TxtStrike 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2520
               TabIndex        =   21
               Top             =   80
               Width           =   1215
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Strilke"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1920
               TabIndex        =   94
               Top             =   140
               Width           =   615
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Call/Put"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   140
               Width           =   855
            End
         End
         Begin VB.CheckBox ChkCarry 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Carry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9000
            TabIndex        =   26
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox ChkAppBrok 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apply Brokerage"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10200
            TabIndex        =   27
            Top             =   960
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   490
            Left            =   4320
            TabIndex        =   49
            Top             =   960
            Width           =   4575
            Begin VB.TextBox TxtRefLot 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   960
               TabIndex        =   61
               Top             =   80
               Width           =   1095
            End
            Begin VB.TextBox TxtSettleRate 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   80
               Width           =   1215
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "Lot Size"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   135
               Width           =   855
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "Settle Rate"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   51
               Top             =   135
               Width           =   1095
            End
         End
         Begin VB.TextBox TxtLot 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9195
            TabIndex        =   14
            Top             =   480
            Width           =   650
         End
         Begin VB.TextBox TxtBrokerCode 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12360
            MaxLength       =   15
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtPtyCode 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   10
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox TxtContype 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   8640
            MaxLength       =   3
            TabIndex        =   13
            Text            =   "Buy"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TxtQty 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9885
            TabIndex        =   15
            Top             =   480
            Width           =   1020
         End
         Begin VB.TextBox TxtRate 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   10935
            TabIndex        =   16
            Top             =   480
            Width           =   1320
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   14640
            TabIndex        =   23
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtConNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   9
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox TxtConRate 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13200
            TabIndex        =   22
            Top             =   960
            Width           =   1335
         End
         Begin vcDateTimePicker.vcDTP vcDTP2 
            Height          =   360
            Left            =   2640
            TabIndex        =   25
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   41178.8109953704
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Height          =   420
            Left            =   15960
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   741
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo DComboTSauda 
            Height          =   390
            Left            =   4800
            TabIndex        =   12
            Top             =   480
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   420
            Left            =   2430
            TabIndex        =   11
            Top             =   480
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   741
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   420
            Left            =   13440
            TabIndex        =   18
            Top             =   480
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   741
            _Version        =   393216
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
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Qty."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9960
            TabIndex        =   102
            Top             =   0
            Width           =   840
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Lot"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9195
            TabIndex        =   48
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12360
            TabIndex        =   40
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Broker Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13440
            TabIndex        =   39
            Top             =   0
            Width           =   2415
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   38
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2550
            TabIndex        =   37
            Top             =   0
            Width           =   2115
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sauda"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4875
            TabIndex        =   36
            Top             =   0
            Width           =   3735
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "B/S"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8640
            TabIndex        =   35
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   34
            Top             =   -120
            Width           =   1020
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Rate"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10935
            TabIndex        =   33
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Con No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   15720
            TabIndex        =   31
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Con Rate"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12120
            TabIndex        =   30
            Top             =   1020
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmcont4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LFExCode As String:                  Dim LFParty As String:              Dim LFBroker As String:            Dim LBillExCodes As String
Dim LBillParties As String:              Dim LBillSaudas As String:          Dim LSExCodes As String:           Dim LSPNames As String
Dim LSType As String:                    Dim LSUserIds As String:            Dim LItemCodeDBCombo As String:    Dim LFSauda As String
Dim LFBPress As Integer:                 Dim LPDataImport As Byte:           Dim SaveCalled As Boolean:         Dim LBillItems As String
Dim LOldParty As String:                 Dim LOldBroker As String:           Dim LOldContype As String:         Dim LOldSauda As String
Dim LOldEXCode As String:                Dim LOldRate2 As Double:            Dim LOldQty As Double:             Dim LOldRate As Double
Dim LOldConno As Long:                   Dim ExRec As ADODB.Recordset:       Dim PartyRec As ADODB.Recordset:   Dim ItemRec As ADODB.Recordset
Dim AllSaudaRec As ADODB.Recordset:      Dim SaudaRec As ADODB.Recordset:    Dim LFPartyRec As ADODB.Recordset: Dim ContRec As ADODB.Recordset
Dim LFSaudaRec As ADODB.Recordset:       Dim LFBrokerRec As ADODB.Recordset: Dim LListSaudas As String:         Dim LlistParties As String
Dim LFExID As Integer


Public Sub ShowStanding()
Dim NStandRec As ADODB.Recordset
If LenB(LFParty) > 1 Then
    mysql = "EXEC Get_PartyNetQtyPARTY " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ",'" & LFParty & "' "
Else
    mysql = "EXEC Get_PartyNetQty " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & " "
End If
Set NStandRec = Nothing: Set NStandRec = New ADODB.Recordset
NStandRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not NStandRec.EOF Then
    Set DataGrid2.DataSource = NStandRec
    DataGrid2.ReBind
    DataGrid2.Refresh
    DataGrid2.Columns(0).Width = 3000:
    DataGrid2.Columns(1).Width = 3000
    DataGrid2.Columns(2).Width = 1200

    DataGrid2.Columns(2).Alignment = dbgRight:
    DataGrid2.Columns(2).NumberFormat = "0.00"
Else
    Set DataGrid2.DataSource = Nothing
    DataGrid2.Refresh
End If
End Sub

Private Sub ChkShowContract_LostFocus()
    Call FillTradeSaudaCombo
End Sub

Private Sub cmbtype_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub cmbtype_Validate(Cancel As Boolean)
 Call DATA_GRID_REFRESH
End Sub

Private Sub CmdSave_Click()
    Dim LExCode As String:          Dim LDelFlag As Boolean:        Dim LOConNo As String:      Dim LContime As String:     Dim LCSauda As String
    Dim LCItemCode As String:       Dim LConType As String:         Dim LSInstType As String:   Dim LStatus As String:      Dim LST_Time As String
    Dim mparty As String:           Dim LCLot As Double:            Dim LCRefLot As Double:     Dim LCBrokLot As Double:    Dim LSCondate As Date
    Dim LConNo As Long:             Dim LClient As String:          Dim LExCont As String:      Dim MSaudaCode As String:   Dim LItemCode As String
    Dim MQty As Double:             Dim MRate As Double:            Dim LCalval As Double:      Dim MConRate As Double:     Dim LSConSno As Long:
    Dim LSOptType  As String:       Dim LSStrike As Double:         Dim LBSParty As String:     Dim LBrokFlag As String:    Dim LATime As String
    Dim TRec As ADODB.Recordset:    Dim NRec As ADODB.Recordset:    Dim LOrdNo As String:       Dim LSaudaID As Long
    Dim LExID As Integer:           Dim LItemID  As Integer
    
    On Error GoTo err1
    LDelFlag = False
    DoEvents
    LSCondate = DtpCondate.Value
    If GSysLockDt > LSCondate Then
        MsgBox "Can Not Add/Modfify/ Delete Trades. Settlement Locked Till " & GSysLockDt & ""
        Exit Sub
    End If
    
    If OptOrder.Value And CmbOrderStatus.ListIndex = 3 Then  'executed
        MsgBox "Executed order can't save!!!!", vbInformation
        Exit Sub
    End If

    'Check Opening
    Frame1.Enabled = False
    If LenB(TxtConNo.text) < 1 Then
        MsgBox "Trade No can not be Blank":        Frame1.Enabled = True
        TxtConNo.Locked = False
        TxtConNo.SetFocus:
        Exit Sub
    Else
        LConNo = Val(TxtConNo.text)
    End If
    If LenB(TxtPtyCode.text) = 0 Then
        MsgBox "Party Code can not be Blank":        Frame1.Enabled = True
        TxtPtyCode.SetFocus
        Exit Sub
    Else
        mparty = Get_AccountDCode(TxtPtyCode.text)
        If LenB(mparty) < 1 Then
            MsgBox "Invalid Party Code":            Frame1.Enabled = True
            TxtPtyCode.SetFocus
            Exit Sub
        Else
            LClient = mparty
        End If
    End If
    If OptTrade.Value Then
        If LenB(TxtBrokerCode.text) = 0 Then
            MsgBox "Broker A/c can not be Blank":        Frame1.Enabled = True
            TxtBrokerCode.SetFocus
            Exit Sub
        Else
            LExCont = Get_AccountDCode(TxtBrokerCode.text)
            If LenB(LExCont) < 1 Then
                MsgBox "Invalid Broker Code":            Frame1.Enabled = True
                TxtBrokerCode.SetFocus
                Exit Sub
            End If
        End If
    Else
        LExCont = 0
    End If
    
    If mparty = LExCont Then
        MsgBox " Buyer and Seller party Can not be same Pls Correct"
        Frame1.Enabled = True:
        TxtBrokerCode.SetFocus:               Exit Sub
    End If
    LCSauda = vbNullString
    If LenB(DComboTSauda.BoundText) = 0 Then
        MsgBox "Sauda Code can not be Blank"
        Frame1.Enabled = True
        TxtPtyCode.SetFocus:
        Exit Sub
    Else
        If ChkShowContract.Value = 1 Then
            AllSaudaRec.MoveFirst
            AllSaudaRec.Find "SAUDACODE='" & DComboTSauda.BoundText & "'", , adSearchForward
            If Not AllSaudaRec.EOF Then
                LCLot = AllSaudaRec!lot:                LCRefLot = AllSaudaRec!REFLOT
                LCBrokLot = AllSaudaRec!BROKLOT:        LSInstType = AllSaudaRec!INSTTYPE
                LCItemCode = vbNullString:              LCSauda = vbNullString
                LExID = Get_ExID(AllSaudaRec!excode)
                LCItemCode = Get_ItemMaster(LExID, AllSaudaRec!EX_SYMBOL)
                
                If LenB(LCItemCode) < 1 Then
                    If AllSaudaRec!LOTWISE = "Y" Then
                        LCItemCode = Create_TItemMast(AllSaudaRec!ITEMCODE, AllSaudaRec!ITEMName, AllSaudaRec!EX_SYMBOL, AllSaudaRec!TRADEABLELOT, AllSaudaRec!excode)
                    Else
                        LCItemCode = Create_TItemMast(AllSaudaRec!ITEMCODE, AllSaudaRec!ITEMName, AllSaudaRec!EX_SYMBOL, AllSaudaRec!lot, AllSaudaRec!excode)
                    End If
                End If
                If LenB(LCItemCode) < 1 Then
                    Frame1.Enabled = True
                    MsgBox "Import new Contracts First"
                    Exit Sub
                End If
                LItemID = Get_ITEMID(LCItemCode)
                LCSauda = Get_SaudaMaster(LExID, LItemID, AllSaudaRec!MATURITY, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
                If LenB(LCSauda) < 1 Then LCSauda = Create_TSaudaMast(LCItemCode, AllSaudaRec!MATURITY, AllSaudaRec!excode, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
                If LenB(LCSauda) < 1 Then
                    MsgBox "Import new Contracts First"
                    Frame1.Enabled = True
                    Exit Sub
                Else
                    MSaudaCode = LCSauda:                    LItemCode = LCItemCode
                    LExCode = AllSaudaRec!excode:            LSOptType = AllSaudaRec!OPTTYPE:
                    LSStrike = AllSaudaRec!STRIKEPRICE:      LCalval = AllSaudaRec!lot
                    LSaudaID = Get_SaudaID(MSaudaCode)
                End If
            Else
                MsgBox "Invalid Contract"
                Frame1.Enabled = True
                Exit Sub
            End If
            SaudaRec.Requery
        Else
            SaudaRec.MoveFirst
            SaudaRec.Find "SAUDACODE='" & DComboTSauda.BoundText & "'", , adSearchForward
            If Not SaudaRec.EOF Then
                MSaudaCode = SaudaRec!saudacode:            LItemCode = Trim(SaudaRec!ITEMCODE)
                LExCode = SaudaRec!excode:                  LSInstType = SaudaRec!INSTTYPE
                LSOptType = SaudaRec!OPTTYPE:               LSStrike = SaudaRec!STRIKEPRICE
                LCalval = SaudaRec!lot
                LSaudaID = SaudaRec!SAUDAID
                LExID = SaudaRec!EXID
                LItemID = SaudaRec!itemid
                If SaudaRec!excode = "NSE" Then LCalval = SaudaRec!TRADEABLELOT
            Else
                MsgBox "Invalid Sauda Code":
                Frame1.Enabled = True
                DComboTSauda.SetFocus
                Exit Sub
            End If
        End If
    End If
    LCalval = Get_LotSize(LItemID, LSaudaID, LExID)
    If Val(TxtQty.text) = 0 Then
        If LFBPress = 1 Then
            MsgBox "Trade Qty can not be Zero "
            Frame1.Enabled = True:
            TxtQty.SetFocus:        Exit Sub
        Else
            If MsgBox("You are about to Delete this Trade. Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
                LDelFlag = True
            Else
                MsgBox "Trade Qty can not be Zero ":
                Frame1.Enabled = True:
                TxtQty.SetFocus:
                Exit Sub
            End If
        End If
    Else
        MQty = Round(Val(TxtQty.text), 2)
    End If
    If Val(TxtRate.text) = 0 Then
        MsgBox "Trade Rate can not be Zero ":
        Frame1.Enabled = True
        TxtRate.SetFocus
        Exit Sub
    Else
        MRate = Round(Val(TxtRate.text), 4)
    End If
    MConRate = 0
    
    If OptTrade.Value Then
        If Val(TxtConRate.text) = 0 Then
            MsgBox "Trade Con Rate can not be Zero "
            Frame1.Enabled = True:
            TxtConRate.SetFocus:          Exit Sub
        Else
            MConRate = Round(Val(TxtConRate.text), 4)
        End If
    End If
    
    If OptTrade.Value Then 'trade
        LSConSno = Get_ConSNo(LSCondate, MSaudaCode, LItemCode, LExCode, LSaudaID, LItemID, LExID)
    Else
        LSConSno = 0
    End If
    
    LOConNo = LConNo:    LContime = Time:    LOrdNo = LTrim$(RTrim$(str(LConNo)))
        
    DoEvents
    CNNERR = True
    Cnn.BeginTrans
    
    If LFBPress = 2 Then
        LConNo = Val(TxtConNo.text):        LOConNo = Trim(Text7.text)
        LOrdNo = Trim(TxtOrdNo.text)
                        
        'sACHIN -- to move contrat entry in log table before edit
        'mysql = "INSERT INTO CTR_D_LOG (CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,loginuser,datetm,[tran])  "
        'mysql = mysql & "SELECT CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,'" & GUserName & "',getdate() "
        'If Val(TxtQty.text) = 0 Then
        '    mysql = mysql & ",'4' "
        'Else
        '    mysql = mysql & ",'" & LFBPress & "' "
        'End If
        
       ' mysql = mysql & "FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONNO=" & Val(TxtConNo.text) & "  AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
       ' Cnn.Execute mysql

        mysql = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONNO=" & Val(TxtConNo.text) & "  AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
        Cnn.Execute mysql
        LATime = CStr(Date) & " " & CStr(Time)
        'Call PInsert_Ctr_Log(LOldConno, LOConNo, "Delete", LATime, DtpCondate.Value, LOldEXCode, LOldSauda, LOldParty, LOldBroker, LOldContype, LOldQty, LOldRate, LOldRate2, GUserName)
    Else
        TxtConfirm.text = "0"
        TxtBrokerConfirm.text = "0"
    End If
    If LDelFlag = False Then
        If LFBPress = 1 Then
            If OptTrade.Value Then 'trade
                If GConNoType <> 0 Then
                    LConNo = Get_Max_ConNo(LSCondate, LFExID)
                Else
                    LConNo = Get_Max_ConNo(LSCondate, 0)
                End If
                LConNo = LConNo + 1
                LOConNo = LConNo
            End If
        End If
        If TxtContype.text = "Buy" Then
            LConType = "B"
        Else
            LConType = "S"
        End If
        If ChkCarry.Value = 1 Then
            LOrdNo = "Carry"
        End If
        If TxtContype.text = "Buy" Then
            LConType = "B"
        Else
            LConType = "S"
        End If
        If ChkAppBrok.Value = 1 Then
            LBrokFlag = "Y"
        Else
            LBrokFlag = "N"
        End If
        
        If OptTrade.Value Then 'TRADE
            Call Add_To_Ctr_D2(LConType, LClient, LSConSno, LSCondate, LConNo, MSaudaCode, LItemCode, mparty, MQty, MRate, MConRate, LExCont, LContime, LOrdNo, vbNullString, _
            LOConNo, LExCode, LCalval, LPDataImport, vbNullString, LSInstType, LSOptType, LSStrike, Left$(TxtFileType.text, 3), LBrokFlag, LExID, LItemID, LSaudaID)
                        
            If LFBPress = 1 Then
                mysql = "UPDATE CTR_D SET ContractType='T' WHERE COMPCODE =" & GCompCode & " AND CONSNO =" & LSConSno & " "
                mysql = mysql & " And CONNO = " & LConNo & " AND CONDATE ='" & Format(LSCondate, "YYYY/MM/dd") & "'AND PARTY='" & mparty & "'"
                Cnn.Execute mysql
            End If
            
        ElseIf OptOrder.Value Then
            Call Add_To_ORD_M_D(LConType, LClient, CmbOrderStatus.ListIndex, LSCondate, LConNo, MSaudaCode, LItemCode, mparty, MQty, MRate, MConRate, LExCont, LContime, LOrdNo, vbNullString, _
            LOConNo, LExCode, LCalval, LPDataImport, vbNullString, LSInstType, LSOptType, LSStrike, Left$(TxtFileType.text, 3), LBrokFlag, LExID, LItemID, LSaudaID)
            
        End If
                                    
        LATime = CStr(Date) & " " & CStr(Time)
        'Call PInsert_Ctr_Log(LConNo, LOConNo, "Add", LATime, DtpCondate.Value, LExCode, MSaudaCode, MParty, LExCont, LConType, MQty, MRate, MConRate, GUserName)
        If LFBPress = 2 Then
            mysql = "UPDATE CTR_D SET CONFIRM=" & Val(TxtConfirm.text) & " WHERE COMPCODE =" & GCompCode & " AND CONSNO =" & LSConSno & " "
            mysql = mysql & " And CONNO = " & LConNo & " AND CONDATE ='" & Format(LSCondate, "YYYY/MM/dd") & "'AND PARTY='" & mparty & "'"
            Cnn.Execute mysql
        
            mysql = "UPDATE CTR_D SET CONFIRM=" & Val(TxtBrokerConfirm.text) & " WHERE COMPCODE =" & GCompCode & " AND CONSNO =" & LSConSno & " "
            mysql = mysql & " And CONNO = " & LConNo & " AND CONDATE ='" & Format(LSCondate, "YYYY/MM/DD") & "'AND PARTY='" & LExCont & "'"
            Cnn.Execute mysql
        ElseIf LFBPress = 1 Then
            'sACHIN -- to move contrat entry in log table before edit
            'mysql = "INSERT INTO CTR_D_LOG (CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,loginuser,datetm,[tran])  "
            'mysql = mysql & "SELECT CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,'" & GUserName & "',getdate(),'" & LFBPress & "' FROM CTR_D  "
            'mysql = mysql & "WHERE COMPCODE =" & GCompCode & " AND CONNO=" & LConNo & "  AND CONDATE ='" & Format(LSCondate, "YYYY/MM/DD") & "'"
            'Cnn.Execute mysql
        
        
        End If
        LBSParty = "'" & mparty & "','" & LExCont & "'"
        'Call Delete_Inv_D(LBSParty, "'" & LEXCODE & "'", "'" & MSaudaCode & "'", DtpCondate.Value)
    End If
    Cnn.CommitTrans
    CNNERR = False
    If LenB(LBillParties) < 1 Then
        LBillParties = "'" & mparty & "','" & LExCont & "'"
    Else
        If InStr(LBillParties, "'" & mparty & "'") < 1 Then LBillParties = LBillParties & ",'" & mparty & "'"
        If InStr(LBillParties, "'" & LExCont & "'") < 1 Then LBillParties = LBillParties & ",'" & LExCont & "'"
    End If
    If LenB(LBillExCodes) < 1 Then
        LBillExCodes = Trim(str(LExID))
    Else
        If LStr_Exists(LBillExCodes, str(LExID)) = False Then LBillExCodes = LBillExCodes & "," & Trim(str(LExID))
    End If
    If LenB(LBillItems) < 1 Then
        LBillItems = "'" & LItemID & "'"
    Else
        If InStr(LBillItems, LItemID) < 1 Then LBillItems = LBillItems & "," & "'" & LItemID & "'"
    End If
    
    If LenB(LBillSaudas) < 1 Then
        LBillSaudas = Trim(str(LSaudaID))
    Else
        If LStr_Exists(LBillSaudas, str(LSaudaID)) = False Then LBillSaudas = LBillSaudas & "," & Trim(str(LSaudaID))
    End If
    
    LConNo = LConNo + 1
    Call DATA_GRID_REFRESH
    If GShowStd = "Y" Then Call ShowStanding
    
    '>>> ORDER_EXECUTION
''    If OptOrder.Value Then
''        If CmbOrderStatus.ListIndex = 1 Then '>>> pending
''            Frame14.Visible = True
''            Label34.Caption = "Order#:" & TxtConNo.text
''            Label35.Caption = "EXEC ORDER_EXECUTION '" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & GCompCode & ",'" & TxtConNo.text & "'"
''            Option2.SetFocus
''
''
''        End If
''    End If
''
    TxtLot = vbNullString:          TxtQty.text = vbNullString:
    TxtRate.text = vbNullString:    TxtConRate.text = vbNullString
    TxtOptType.text = "CE":         TxtStrike.text = vbNullString
    TxtConfirm.text = "0":          TxtBrokerConfirm.text = "0"
    Frame1.Enabled = True
    If LFBPress = 2 Then
        TxtConNo.text = vbNullString:               TxtPtyCode.text = vbNullString
        DataCombo2.BoundText = vbNullString:        DComboTSauda.BoundText = vbNullString
        TxtOrdNo = vbNullString:                    TxtConNo.SetFocus
    Else
        TxtConNo.text = LConNo:
        If GUniqClientId = "2175AHM" Or GCINNo = "3000" Then
            TxtContype.SetFocus
        Else
            TxtPtyCode.SetFocus
        End If
        
    End If
    SaveCalled = True
    ChkCarry.Value = 0
    
    Exit Sub
err1:
If err.Number <> 0 Then
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    Frame1.Enabled = True
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End If
End Sub
Private Sub CmdAdd_Click()
   ' Call Chk_Brokerage(DtpCondate.Value)
   
'    '---------------
'    Call FillTradeSaudaCombo
'    Call FillLFPartyCombo
'    Call FillFSaudaCombo
'    Call DATA_GRID_REFRESH
'    Call FillFBrokerCombo
'    If GShowStd = "Y" Then Call ShowStanding
'    DtpCondate.Enabled = False
'    ChkShowContract.Enabled = False
'    InstCombo.Enabled = False
'    DComboExchange.Enabled = False
'    '---------------


    If OptOrder.Value Then 'order
        TxtConNo.text = Trim(Get_Max_OrderNo(DtpCondate.Value, 0) + 1)
    Else
        If GConNoType <> 0 Then
            TxtConNo.text = Trim(Get_Max_ConNo(DtpCondate.Value, LFExID) + 1)
        Else
            TxtConNo.text = Trim(Get_Max_ConNo(DtpCondate.Value, 0) + 1)
        End If
    End If
    
    TxtConNo.Locked = True:                             TxtQty.Locked = False
    TxtPtyCode.text = vbNullString:                     DataCombo2.BoundText = vbNullString
    TxtLot = vbNullString:                              TxtQty.text = vbNullString
    TxtRate.text = vbNullString:                        TxtConRate.text = vbNullString
    CmdModify.Enabled = False:                          Frame2.Enabled = True
    CmdAdd.Enabled = False
    CmdCancel.Enabled = True
    DtpCondate.Enabled = False
    LFBPress = 1
    LPDataImport = "0"
    
    If OptOrder.Value Then 'order
        Label12.Caption = "Add Order"
        CmbOrderStatus.ListIndex = 1
    Else 'trade
        Label12.Caption = "Add Trades"
    End If
            
    If GUniqClientId = "2175AHM" Or GCINNo = "3000" Then
        TxtContype.SetFocus
    Else
        TxtPtyCode.SetFocus
    End If
End Sub
Private Sub CmdModify_Click()
    '---------------
    Call FillTradeSaudaCombo
    Call FillLFPartyCombo
    Call FillFSaudaCombo
    Call DATA_GRID_REFRESH
    Call FillFBrokerCombo
    If GShowStd = "Y" Then Call ShowStanding
    DtpCondate.Enabled = False
    ChkShowContract.Enabled = False
    InstCombo.Enabled = False
    DComboExchange.Enabled = False
    '---------------
    
    Call Mod_Rec
End Sub
Private Sub CmdCancel_Click()

Call CANCEL_REC
End Sub

Private Sub Command1_Click()
Dim I As Integer
LListSaudas = vbNullString
For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(I).Checked = True Then
        If LenB(LListSaudas) > 1 Then
            LListSaudas = LListSaudas & ",'" & ListView1.ListItems(I).text & "'"
        Else
            LListSaudas = "'" & ListView1.ListItems(I).text & "'"
        End If
    End If
Next
Call DATA_GRID_REFRESH
End Sub

Private Sub Command2_Click()

Dim I As Integer
LlistParties = vbNullString
For I = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(I).Checked = True Then
        If LenB(LlistParties) > 1 Then
            LlistParties = LlistParties & ",'" & ListView2.ListItems(I).text & "'"
        Else
            LlistParties = "'" & ListView2.ListItems(I).text & "'"
        End If
    End If
Next
Call DATA_GRID_REFRESH

End Sub

Private Sub Command3_Click()
    Frame12.Visible = False
End Sub

Private Sub Command5_Click()
    If Option1.Value Then '>>> without high low
        mysql = Label35.Caption & ",'N'"
    Else '>>> with high low
        mysql = Label35.Caption & ",'Y'"
    End If
    Cnn.Execute mysql
    Frame14.Visible = False
End Sub

Private Sub Command6_Click()
    Frame14.Visible = False
End Sub

Private Sub Command9_Click()
Dim LDel As Boolean
Dim LATime As String
Dim LSCondate As Date
Dim LSaudaID As Double
Dim LExID As Double
If ContRec.RecordCount > 0 Then
    LSCondate = DtpCondate.Value
    If GSysLockDt > LSCondate Then
        MsgBox "Can Not Add/Modfify/Delete Trades. Settlement Locked Till " & GSysLockDt & ""
        Exit Sub
    End If
    If OptTrade.Value Then
        If MsgBox("Are You Sure You Want to Delete all Trades of " & DtpCondate.Value & " of " & DComboExchange.BoundText & "", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
            If Not ContRec.EOF Then
                ContRec.MoveFirst
                Do While Not ContRec.EOF
                
                    'sACHIN -- to move contrat entry in log table before edit
                    'mysql = "INSERT INTO CTR_D_LOG (CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,loginuser,datetm,[tran])  "
                    'mysql = mysql & "SELECT CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,'" & GUserName & "',getdate(),'3' FROM CTR_D  "
                    'mysql = mysql & "WHERE COMPCODE =" & GCompCode & " AND CONNO=" & ContRec!CONNO & "  AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
                    'Cnn.Execute mysql
                        
                    mysql = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & ContRec!excode & "' AND CONNO = " & ContRec!CONNO & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    LATime = CStr(Date) & " " & CStr(Time)
                    
                    LSaudaID = Get_SaudaID(ContRec!Sauda)
                    LExID = Get_ExID(ContRec!excode)
                    
                    If LenB(LBillParties) < 1 Then
                        LBillParties = "'" & ContRec!PARTY & "','" & ContRec!Code & "'"
                     '   If TRec!PARTY <> TRec!CONCODE Then
                     '       LBillParties = "'" & TRec!PARTY & "','" & TRec!CONCODE & "'"
                     '   Else
                     '       LBillParties = "'" & TRec!PARTY & "'"
                     '   End If
                    Else
                        If InStr(LBillParties, "'" & ContRec!PARTY & "'") < 1 Then LBillParties = LBillParties & ",'" & ContRec!PARTY & "'"
                        If InStr(LBillParties, "'" & ContRec!Code & "") < 1 Then LBillParties = LBillParties & ",'" & ContRec!Code & "'"
                    End If
                    
                   ' If LenB(LBillItems) < 1 Then
                   '     LBillItems = "'" & TRec!ITEMCODE & "'"
                   ' Else
                   '     If InStr(LBillItems, TRec!ITEMCODE) < 1 Then LBillItems = LBillItems & "," & "'" & TRec!ITEMCODE & "'"
                   ' End If
                   
                    If LenB(LBillExCodes) < 1 Then
                        LBillExCodes = Trim(str(LExID))
                    Else
                        If LStr_Exists(LBillExCodes, str(LExID)) = False Then LBillExCodes = LBillExCodes & "," & Trim(str(LExID))
                    End If
        
                    If LenB(LBillSaudas) < 1 Then
                        LBillSaudas = Trim(str(LSaudaID))
                    Else
                        If LStr_Exists(LBillSaudas, LSaudaID) = False Then LBillSaudas = LBillSaudas & "," & Trim(str(LSaudaID)) & ""
                    End If
                                                        
                    'Call PInsert_Ctr_Log(ContRec!CONNO, ContRec!TRADENO, "Delete", LATime, DtpCondate.Value, ContRec!EXCODE, ContRec!Sauda, ContRec!PARTY, ContRec!BROKER, ContRec!BS, ContRec!QTY, ContRec!Rate, ContRec!CONRATE, GUserName)
                    ContRec.MoveNext
                Loop
                SaveCalled = True
                Call CANCEL_REC
            End If
            'Call Delete_Inv_D(vbNullString, vbNullString, vbNullString, DtpCondate.Value)
        End If
    End If
    DATA_GRID_REFRESH
End If
End Sub

Private Sub DComboExchange_Validate(Cancel As Boolean)
DoEvents
Frame15.Visible = True

    If LenB(DComboExchange.BoundText) = 0 Then
        LFExCode = vbNullString
        LFExID = 0
        If GConNoType <> 0 Then
            MsgBox "Please Select Exchange"
            Cancel = True
            Exit Sub
        End If
    Else
        LFExCode = DComboExchange.BoundText
        ExRec.Filter = adFilterNone
        ExRec.Filter = "EXCODE='" & LFExCode & "'"
        LFExID = ExRec!EXID
    End If
    
    If LFExCode = "LME" Then
        DataCombo5.Visible = True
        vcDTP2.Visible = True
        ChkCarry.Visible = True
        vcDTP2.Value = DtpCondate.Value + 90
        Set ItemRec = Nothing
        mysql = "SELECT ITEMCODE,ITEMNAME,EXCHANGECODE,LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND EXCHANGECODE ='LME' ORDER BY ITEMCODE "
        Set ItemRec = New ADODB.Recordset
        ItemRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Set DataCombo5.RowSource = ItemRec:
        DataCombo5.BoundColumn = "ITEMCODE"
        DataCombo5.ListField = "ITEMNAME"
    Else
        DataCombo5.Visible = False
        vcDTP2.Visible = False
        ChkCarry.Visible = False
    End If
    Call FillTradeSaudaCombo
    Call FillLFPartyCombo
    Call FillFSaudaCombo
    Call DATA_GRID_REFRESH

    If GShowStd = "Y" Then Call ShowStanding
DoEvents
Frame15.Visible = False
End Sub
Private Sub DataCombo4_Validate(Cancel As Boolean)
'Dim NRec As ADODB.Recordset
Dim LBroker As String
If LenB(DataCombo4.text) = 0 Then
    MsgBox "Broker A/c can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    LBroker = Get_AccountDCode(DataCombo4.BoundText)
    If LenB(LBroker) > 0 Then
        TxtBrokerCode.text = LBroker
    Else
        DataCombo4.SetFocus
        Cancel = True
        Sendkeys "%{DOWN}"
    End If
End If
End Sub
Private Sub DataCombo2_Validate(Cancel As Boolean)
'Dim NRec As ADODB.Recordset
Dim LAcCode As String

If LenB(DataCombo2.text) = 0 Then
    MsgBox "Party can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    LAcCode = Get_AccountDCode(DataCombo2.BoundText)
    If LenB(LAcCode) > 1 Then
        TxtPtyCode.text = LAcCode
        If Frame2.Enabled = False Then
            Frame10.Enabled = True
            DtpCondate.Enabled = True
            DtpCondate.SetFocus
        ElseIf FrameOpt.Visible = True Then
            TxtOptType.SetFocus
        Else
           DComboTSauda.SetFocus
        End If
    Else
        DataCombo2.SetFocus
        Cancel = True
        Sendkeys "%{DOWN}"
    End If
End If
End Sub
Private Sub DComboTSauda_Validate(Cancel As Boolean)
DoEvents
Frame15.Visible = True
Dim LSaudaID As Long
If LenB(DComboTSauda.text) = 0 Then
   ' MsgBox "Sauda can not be blank"
    'Cancel = True
  '  Sendkeys "%{DOWN}"
Else
    Call Get_Value
End If
    LSaudaID = Get_SaudaID(DComboTSauda.BoundText)
    TxtSettleRate.text = Format(SDCLRATE(LSaudaID, DtpCondate.Value, "C"), "0.00")
    If GCINNo = "3000" Then Call Get_Value
DoEvents
Frame15.Visible = False
End Sub
Private Sub DataCombo5_Validate(Cancel As Boolean)
If LenB(DataCombo5.text) = 0 Then
    MsgBox "Itemcode can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    LItemCodeDBCombo = DataCombo5.BoundText
End If
End Sub
Private Sub DComboParty_Validate(Cancel As Boolean)
    If LenB(DComboParty.BoundText) <> 0 Then
        LFParty = DComboParty.BoundText
    Else
        LFParty = vbNullString
    End If
    
    If GShowStd = "Y" Then Call ShowStanding
    Call FillFSaudaCombo
    Call FillFBrokerCombo
    Call DATA_GRID_REFRESH
End Sub

Private Sub DComboCode_Validate(Cancel As Boolean)
If LenB(DComboCode.BoundText) <> 0 Then
    LFParty = DComboCode.BoundText
    DComboParty.BoundText = LFParty
Else
    LFParty = vbNullString
End If
If GShowStd = "Y" Then Call ShowStanding
Call FillFSaudaCombo
Call FillFBrokerCombo
Call DATA_GRID_REFRESH
End Sub

Private Sub DComboSauda_Validate(Cancel As Boolean)
FillFBrokerCombo
'If LenB(DComboSauda.BoundText) <> 0 Then
'    LFSauda = DComboSauda.BoundText
'Else
'    LFSauda = vbNullString
'End If
'Set LFBrokerRec = Nothing
'Set LFBrokerRec = New ADODB.Recordset
'MYSQL = "SELECT DISTINCT A.AC_CODE,A.NAME FROM ACCOUNTD  AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
'MYSQL = MYSQL & " AND A.AC_CODE  =B.CONCODE  AND B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
'If LenB(LFExCode) > 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
'If LenB(LFParty) > 0 Then MYSQL = MYSQL & " AND B.PARTY ='" & LFParty & "'"
'If LenB(LFSauda) > 0 Then MYSQL = MYSQL & " AND B.SAUDA ='" & LFSauda & "'"
'MYSQL = MYSQL & " ORDER BY A.NAME"
'LFBrokerRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'If Not LFBrokerRec.EOF Then
'    DComboBroker.Enabled = True
'    Set DComboBroker.RowSource = LFBrokerRec
'    DComboBroker.BoundColumn = "AC_CODE"
'    DComboBroker.ListField = "NAME"
'Else
'    DComboBroker.Enabled = False
''End If
'Call DATA_GRID_REFRESH
End Sub
Private Sub DComboBroker_Validate(Cancel As Boolean)
If LenB(DComboBroker.BoundText) <> 0 Then
    LFBroker = DComboBroker.BoundText
Else
    LFBroker = vbNullString
End If
Call DATA_GRID_REFRESH
End Sub
Private Sub DataGrid1_DblClick()
    Dim LPConNo As Long:         Dim LPSauda As String:      Dim LPConType As String:        Dim TRec As ADODB.Recordset
    If GCINNo = "2000" Then
        DataGrid1.Col = 4:              LPSauda = DataGrid1.text
        DataGrid1.Col = 10:             LPConNo = DataGrid1.text
        DataGrid1.Col = 2:              LPConType = DataGrid1.text
        'DataGrid1.Col = 9:              LPDataImport = Trim(DataGrid1.text)
    Else
        DataGrid1.Col = 2:              LPSauda = DataGrid1.text
        DataGrid1.Col = 11:             LPConNo = DataGrid1.text
        DataGrid1.Col = 3:              LPConType = DataGrid1.text
    End If

    Call Mod_Rec
    Call Get_Trade_Details(LPConNo)
    
    
    CmdAdd.Enabled = True:                          CmdModify.Enabled = False
    CmdCancel.Enabled = True:                       DtpCondate.Enabled = False
    DComboExchange.Enabled = False:
    Frame2.Enabled = True:                          LFBPress = 2
    Label12.Caption = "Modifty Trade"

    
    If (GSysLockDt >= DtpCondate.Value) Then
        CmdAdd.Enabled = False
        CmdModify.Enabled = False
        CmdSave.Enabled = False
    End If
    
    TxtConNo.SetFocus
'End If
Set TRec = Nothing
End Sub

Private Sub DataGridOrder_DblClick()
    Dim LPConNo As Long:         Dim LPSauda As String:      Dim LPConType As String:        Dim TRec As ADODB.Recordset
'    If GCINNo = "2000" Then
'        DataGrid1.Col = 4:              LPSauda = DataGrid1.text
'        DataGrid1.Col = 10:             LPConNo = DataGrid1.text
'        DataGrid1.Col = 2:              LPConType = DataGrid1.text
'        'DataGrid1.Col = 9:              LPDataImport = Trim(DataGrid1.text)
'    Else
        DataGridOrder.Col = 2:              LPSauda = DataGridOrder.text
        DataGridOrder.Col = 10:             LPConNo = DataGridOrder.text
        DataGridOrder.Col = 3:              LPConType = DataGridOrder.text
'    End If

    Call Mod_Rec
    Call Get_Trade_Details(LPConNo)
    
    
    
    CmdAdd.Enabled = True:                          CmdModify.Enabled = False
    CmdCancel.Enabled = True:                       DtpCondate.Enabled = False
    'DComboExchange.Enabled = False:
    Frame2.Enabled = True:                          LFBPress = 2
    Label12.Caption = "Modifty Order"
    
    
    CmdSave.Enabled = True
    If CmbOrderStatus.ListIndex = 3 Then 'executed
        CmdSave.Enabled = False
    End If
    
    If (GSysLockDt >= DtpCondate.Value) Then
        CmdAdd.Enabled = False
        CmdModify.Enabled = False
        CmdSave.Enabled = False
    End If
    
    
    TxtConNo.SetFocus
'End If
Set TRec = Nothing
End Sub


Private Sub DataGrid3_DblClick()
Dim LPConNo As Long:            Dim LPSauda As String:          Dim LPConType As String:        Dim TRec As ADODB.Recordset
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "DtpCondate" Or Me.ActiveControl.NAME = "vcDTP2" Then
            Sendkeys "{tab}"
        End If
    End If
End Sub
Private Sub Form_Load()
Dim TRec As ADODB.Recordset
LPDataImport = 0:           LSExCodes = vbNullString:    LSPNames = vbNullString:    LSType = vbNullString
LSUserIds = vbNullString:   LFExCode = vbNullString:     LFParty = vbNullString:     LFSauda = vbNullString
LFExID = 0:                 LFBroker = vbNullString:     ChkAppBrok.Value = 1:       TxtFileType.text = "0"
LListSaudas = vbNullString: LlistParties = vbNullString: SSTab1.Tab = 0
If GShowLotEntry = "Y" Then
    TxtLot.Visible = True: Label27.Visible = True
Else
    TxtLot.Visible = False: Label27.Visible = False
End If

OptOrder.Visible = False
If GOrderEntryYN = "Y" Then
    OptOrder.Visible = True
Else
    Frame13.Visible = False
    OptTrade.Width = 2415
End If

DtpCondate.Value = Date
Frame2.Visible = True
InstCombo.Clear
InstCombo.AddItem "All"
Set TRec = Nothing
Set TRec = New ADODB.Recordset
mysql = "SELECT DISTINCT INSTTYPE FROM SCRIPTMASTER  ORDER BY INSTTYPE"
TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not TRec.EOF Then
    If TRec.RecordCount > 0 Then
        TRec.MoveFirst
        Do While Not TRec.EOF
            InstCombo.AddItem (TRec!INSTTYPE)
            TRec.MoveNext
        Loop
        InstCombo.Visible = True
    End If
Else
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT INSTTYPE FROM SAUDAMAST ORDER BY INSTTYPE"
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not TRec.EOF Then
        If TRec.RecordCount > 0 Then
            TRec.MoveFirst
            Do While Not TRec.EOF
                InstCombo.AddItem (TRec!INSTTYPE)
                TRec.MoveNext
            Loop
            InstCombo.Visible = True
        End If
    End If
End If
InstCombo.ListIndex = 0:
If TRec.RecordCount = 1 Then
    InstCombo.Locked = True
End If
InstCombo.Visible = True

If GCINNo = "2000" Then
    DataCombo5.Left = 7300:         DComboTSauda.Left = 7300
    Label7.Left = 5280:             Label8.Left = 6080
    Label6.Left = 7300:             TxtContype.Left = 5280
    TxtQty.Left = 6080:             TxtContype.TabIndex = 11
    TxtQty.TabIndex = 12:           DComboTSauda.TabIndex = 13
ElseIf GUniqClientId = "2175AHM" Then
    Label7.Left = 1080:             TxtContype.Left = 1080
    TxtContype.TabIndex = 10:       Label27.Visible = False
    TxtLot.Visible = False:         Label8.Left = 1700
    TxtQty.Left = 1700:             TxtQty.TabIndex = 12
    Label6.Left = 2800:             DComboTSauda.Left = 2800
    Label6.Width = 3000
    Label9.Left = 5800:             DComboTSauda.TabIndex = 13
    TxtRate.Left = 5800:            TxtRate.TabIndex = 14
    Label3.Left = 7200:             TxtPtyCode.Left = 7200
    TxtPtyCode.TabIndex = 15:       Label5.Left = 8150
    DataCombo2.Left = 8150:         DataCombo2.TabIndex = 16
    Label13.Left = 11000:           TxtBrokerCode.Left = 11000
    Label10.Left = 12050:           TxtBrokerCode.TabIndex = 17
    DataCombo4.Left = 12050:        DataCombo4.Width = 3000
    Label10.Width = 3000:           DataCombo4.TabIndex = 18
ElseIf GUniqClientId = "1207BIK" Then
    Label7.Left = 1200:             TxtContype.Left = 1200
    Label8.Left = 1845:             TxtQty.Left = 1845
    Label6.Left = 3000:             DComboTSauda.Left = 3000
    Label9.Left = 6600:             TxtRate.Left = 6600
    Label27.Visible = False:        TxtLot.Visible = False
    Label3.Left = 8100:             TxtPtyCode.Left = 8100
    Label5.Left = 9100:             DataCombo2.Left = 9100
    Label5.Width = 3250:            DataCombo2.Width = 3250
    Label13.Left = 12400:           TxtBrokerCode.Left = 12400
    DataCombo4.Left = 13500:        Label10.Width = 3250
    Label10.Left = 13500:           DataCombo4.Width = 3250
    TxtContype.TabIndex = 8:        TxtQty.TabIndex = 9
    TxtRate.TabIndex = 11:          DComboTSauda.TabIndex = 10
    TxtPtyCode.TabIndex = 12:       DataCombo2.TabIndex = 13
    TxtBrokerCode.TabIndex = 14:    DataCombo4.TabIndex = 15
    
End If
Set ExRec = Nothing: Set ExRec = New ADODB.Recordset
mysql = "SELECT EXID,EXCODE,EXNAME,CONTRACTACC FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME"
ExRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not ExRec.EOF Then
    Set DComboExchange.RowSource = ExRec
    DComboExchange.BoundColumn = "EXCODE"
    DComboExchange.ListField = "EXCODE"
End If
If ExRec.RecordCount = 1 Then
    ExRec.MoveFirst
    DComboExchange.BoundText = ExRec!excode
End If
Set PartyRec = Nothing: Set PartyRec = New ADODB.Recordset
mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
PartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not PartyRec.EOF Then
    Set DataCombo2.RowSource = PartyRec
    DataCombo2.BoundColumn = "AC_CODE"
    DataCombo2.ListField = "NAME"
    Set DataCombo4.RowSource = PartyRec
    DataCombo4.BoundColumn = "AC_CODE"
    DataCombo4.ListField = "NAME"
End If

'If cmbtype.Visible Then
    cmbtype.ListIndex = 0
'End If

End Sub
Private Sub DComboExchange_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboParty_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboCode_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub DComboSauda_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboBroker_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo4_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo2_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboTSauda_GotFocus()
    Sendkeys "%{DOWN}"
    If LenB(DComboTSauda.BoundText) > 0 Then
        DComboTSauda.BoundText = DComboTSauda.BoundText
    End If
    
End Sub
Private Sub DataCombo5_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CANCEL_REC
End Sub

Private Sub InstCombo_Validate(Cancel As Boolean)
    If ChkShowContract.Value = 1 Then
        'If InstCombo.text = "OPT" Then
        '    FrameOpt.Visible = True
        'Else
            FrameOpt.Visible = False
        'End If
    Else
       FrameOpt.Visible = False
    End If
End Sub
Private Sub OptOrder_Click()
    If OptOrder.Value Then
        Label11.Caption = "Ord. No"
        Label13.Visible = False
        ChkShowContract.Caption = "Show All Orders"
        TxtBrokerCode.Visible = False
        Label10.Caption = "Status"
        DataCombo4.Visible = False
        CmbOrderStatus.Visible = True
        If LFBPress = 1 Then
            CmbOrderStatus.ListIndex = 0
        End If
        DataGrid1.Visible = False
        DataGrid3.Visible = False
        DataGridOrder.Visible = True
        Command9.Caption = "Delete All Orders"
        DComboCode.Visible = False
        DComboParty.Visible = False
        DComboSauda.Visible = False
        DComboBroker.Visible = False
        Label31.Visible = False
        
        Label16.Visible = False
        Label21.Visible = False
        Label22.Visible = False
        
        cmbtype.Visible = False
        FrameOpt.Visible = False
        TxtRefLot.Visible = False
        Label26.Visible = False
        TxtSettleRate.Visible = False
        Label28.Visible = False
        ChkCarry.Visible = False
        ChkAppBrok.Visible = False
        TxtConRate.Visible = False
        Label17.Visible = False
        SSTab1.Tab = 0
        SSTab1.Caption = "Orders"
        SSTab1.TabVisible(1) = False
        
    End If
End Sub
Private Sub OptTrade_Click()
    If OptTrade.Value Then
        Label11.Caption = "Con. No"
        Label13.Visible = True
        TxtBrokerCode.Visible = True
        ChkShowContract.Caption = "Show All Contracts"
        Label10.Caption = "Broker Name"
        DataCombo4.Visible = True
        CmbOrderStatus.Visible = False
        DataGridOrder.Visible = False
        DataGrid1.Visible = True
        DataGrid3.Visible = True
        Command9.Caption = "Delete All Trades"
        DComboCode.Visible = True
        DComboParty.Visible = True
        DComboSauda.Visible = True
        DComboBroker.Visible = True
        Label31.Visible = True
        
        Label16.Visible = True
        Label21.Visible = True
        Label22.Visible = True
        
        cmbtype.Visible = True
        cmbtype.ListIndex = 0
        FrameOpt.Visible = False
        TxtRefLot.Visible = True
        Label26.Visible = True
        TxtSettleRate.Visible = True
        Label28.Visible = True
        ChkCarry.Visible = True
        ChkAppBrok.Visible = True
        TxtConRate.Visible = True
        Label17.Visible = True
        SSTab1.Tab = 0
        SSTab1.Caption = "Trade"
        SSTab1.TabVisible(1) = True
    End If
End Sub

Private Sub TxtOptType_GotFocus()
    TxtOptType.SelStart = 0
    TxtOptType.SelLength = Len(TxtOptType.text)
End Sub
Private Sub TxtOptType_KeyPress(KeyAscii As Integer)
If Val(KeyAscii) >= 48 And KeyAscii <= 122 Then
    If Val(KeyAscii) = 66 Or Val(KeyAscii) = 98 Or Val(KeyAscii) = 83 Or Val(KeyAscii) = 115 Then
    Else
        If TxtOptType.text = "CE" Then
            TxtOptType.text = "PE"
        Else
            TxtOptType.text = "CE"
        End If
    End If
End If
If KeyAscii = 32 Then
    If TxtOptType.text = "CE" Then
        TxtOptType.text = "PE"
    Else
        TxtOptType.text = "CE"
    End If
End If
If KeyAscii = 43 Then TxtOptType.text = "CE"
If KeyAscii = 45 Then TxtOptType.text = "PE"

End Sub

Private Sub TxtOptType_Validate(Cancel As Boolean)
If TxtOptType.text = "CE" Or TxtOptType.text = "PE" Then
Else
    MsgBox "Please Select Valid Option Type (CE,PE)"
    Cancel = True
End If
End Sub

Private Sub TxtPtyCode_Validate(Cancel As Boolean)
'Dim NRec As ADODB.Recordset
Dim LAcCode As String
If Frame2.Enabled = True Then
    If LenB(TxtPtyCode.text) = 0 Then
        DataCombo2.SetFocus
    Else
        LAcCode = Get_AccountDCode(TxtPtyCode.text)
        If LenB(LAcCode) > 1 Then
            DataCombo2.BoundText = LAcCode
            DComboTSauda.SetFocus
        Else
            DataCombo2.SetFocus
        End If
    End If
Else
    Frame1.Enabled = True
    Frame10.Enabled = True
    DtpCondate.Enabled = True
    DtpCondate.SetFocus
End If
End Sub
Private Sub TxtConNo_Validate(Cancel As Boolean)
    Dim NewRec As ADODB.Recordset
    TxtQty.Locked = False
    If LFBPress = 2 Then
        If LenB(TxtConNo.text) = 0 Then
            TxtQty.Locked = True
        Else
            TxtQty.Locked = False
            Call Get_Trade_Details(Val(TxtConNo.text))
        End If
    End If
End Sub
Private Sub TxtContype_KeyPress(KeyAscii As Integer)
If Val(KeyAscii) >= 48 And KeyAscii <= 122 Then
    If Val(KeyAscii) = 66 Or Val(KeyAscii) = 98 Or Val(KeyAscii) = 83 Or Val(KeyAscii) = 115 Then
    Else
        If TxtContype.text = "Buy" Then
            TxtContype.text = "Sel"
        Else
            TxtContype.text = "Buy"
        End If
    End If
End If
If KeyAscii = 32 Then
    If TxtContype.text = "Buy" Then
        TxtContype.text = "Sel"
    Else
        TxtContype.text = "Buy"
    End If
End If
If KeyAscii = 43 Then TxtContype.text = "Buy"
If KeyAscii = 45 Then TxtContype.text = "Sel"
    
End Sub

Private Sub TxtContype_Validate(Cancel As Boolean)
If TxtContype.text <> "Buy" Then
    If TxtContype.text <> "Sel" Then
        TxtContype.text = "Buy"
        Cancel = True
        TxtContype.SetFocus
    End If
End If
End Sub
Private Sub TxtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtLot_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtQty_Validate(Cancel As Boolean)
    TxtQty.text = Format(TxtQty.text, "0.00")
    TxtLot.text = CStr(Val(TxtQty.text) / Val(TxtRefLot.text))
End Sub
Private Sub TxtRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtRate_Validate(Cancel As Boolean)
    TxtRate.text = Format(TxtRate.text, "0.0000")
    If LFBPress = 1 Then TxtConRate.text = Format(TxtRate.text, "0.0000")
  If GCINNo <> "3000" Then Call Get_Value
End Sub
Private Sub TxtBrokerCode_GotFocus()
TxtBrokerCode.SelStart = 0
TxtBrokerCode.SelLength = Len(TxtBrokerCode.text)
End Sub
Private Sub TxtBrokerCode_Validate(Cancel As Boolean)
Dim LBroker  As String
If LenB(TxtBrokerCode.text) = 0 Then
    DataCombo4.SetFocus
Else
    LBroker = Get_AccountDCode(TxtBrokerCode.text)
    If LenB(LBroker) > 0 Then
        DataCombo4.BoundText = LBroker
        If LBroker = TxtPtyCode.text Then
            MsgBox "Broker A/c can not be Same As Party A/c"
            TxtConRate.SetFocus
            Exit Sub
        End If
        DataCombo4.SetFocus
    Else
        DataCombo4.SetFocus
    End If
    'Set NRec = Nothing
End If
End Sub
Private Sub TxtConRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtConRate_Validate(Cancel As Boolean)
    TxtConRate.text = Format(TxtConRate.text, "0.0000")
End Sub
Private Sub TxtQty_GotFocus()
    TxtQty.SelStart = 0
    TxtQty.SelLength = Len(TxtQty.text)
End Sub

Private Sub TxtLot_GotFocus()
    TxtLot.SelStart = 0
    TxtLot.SelLength = Len(TxtLot.text)
End Sub

Private Sub TxtConRate_GotFocus()
    TxtConRate.SelStart = 0
    TxtConRate.SelLength = Len(TxtConRate.text)
End Sub

Private Sub TxtRate_GotFocus()
    TxtRate.SelStart = 0
    TxtRate.SelLength = Len(TxtRate.text)
End Sub
Private Sub TxtPtyCode_GotFocus()
    TxtPtyCode.SelStart = 0
    TxtPtyCode.SelLength = Len(TxtPtyCode.text)
End Sub
Public Sub DATA_GRID_REFRESH()
Dim LShreeRec As ADODB.Recordset
            
    If GCINNo = "2000" Then
        mysql = "SELECT A.Party,B.Name,A.CONTYPE AS Type ,A.QTY AS 'Lot', A.Qty,A.Sauda ,A.Rate,A.UserID,A.ConTime,A.ROWNO1 AS TradeNo,"
        mysql = mysql & " A.ConNo,A.CONCODE AS Code,C.NAME AS Broker,A.BROKAMT AS ConRate,A.DATAIMPORT,A.EXCODE FROM "
        mysql = mysql & " CTR_D A, ACCOUNTD B, ACCOUNTD AS C  WHERE A.COMPCODE=" & GCompCode & ""
        mysql = mysql & " AND A.COMPCODE =C.COMPCODE AND A.CONCODE =C.AC_CODE "
        mysql = mysql & " AND A.PARTY NOT IN (SELECT DISTINCT CONCODE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' "
        If LenB(LFExCode) <> 0 Then mysql = mysql & " AND EXID =" & LFExID & ""
        mysql = mysql & ")"
        mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
        If LenB(LFExCode) <> 0 Then mysql = mysql & " AND A.EXID =" & LFExID & ""
        mysql = mysql & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' ORDER BY ROWNO1 DESC "
        Set ContRec = Nothing
        Set ContRec = New ADODB.Recordset
        ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Set DataGrid1.DataSource = ContRec
        DataGrid1.ReBind
        DataGrid1.Refresh
    Else
        '0  Party,         '1  Name,:               '2  Sauda,:               '3    CONTYPE AS Type ,   '4 lot '5  Qty,
        '6  Rate,          '7  CONCODE AS Code,     '8 NAME AS Broker,        '9    BROKAMT AS ConRate, '10 RATE-A.BROKAMT AS DiffRate,
        '11 conno          '12  contime              13  rowno1               '14   dataimport
        
'        If OptOrder.Value Then 'order
'            LFExID = -1
'        End If
        
        If OptOrder.Value Then 'order
            mysql = " GET_GRIDDATA " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "',-1"
        Else
            mysql = " GET_GRIDDATA " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ""
        End If
        
        Set ContRec = Nothing:  Set ContRec = New ADODB.Recordset
        ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'
''        If LenB(LFParty) > 0 And LenB(LFSauda) > 0 And LenB(LFBroker) > 0 Then
''            ContRec.Filter = "PARTY ='" & LFParty & "' AND SAUDA ='" & LFSauda & "' AND CODE ='" & LFBroker & "'"
''        ElseIf LenB(LFParty) > 0 And LenB(LFSauda) < 1 And LenB(LFBroker) < 1 Then
''            ContRec.Filter = "PARTY ='" & LFParty & "'"
''        ElseIf LenB(LFParty) > 0 And LenB(LFSauda) > 0 And LenB(LFBroker) < 1 Then
''            ContRec.Filter = "PARTY ='" & LFParty & "' AND SAUDA ='" & LFSauda & "'"
''        ElseIf LenB(LFParty) > 0 And LenB(LFSauda) < 1 And LenB(LFBroker) > 0 Then
''            ContRec.Filter = "PARTY ='" & LFParty & "' AND CODE ='" & LFBroker & "'"
''        ElseIf LenB(LFParty) < 1 And LenB(LFSauda) > 0 And LenB(LFBroker) > 0 Then
''            ContRec.Filter = " SAUDA ='" & LFSauda & "' AND CODE ='" & LFBroker & "'"
''        ElseIf LenB(LFParty) < 1 And LenB(LFSauda) < 1 And LenB(LFBroker) > 0 Then
''            ContRec.Filter = "  CODE ='" & LFBroker & "'"
''        ElseIf LenB(LFParty) < 1 And LenB(LFSauda) > 0 And LenB(LFBroker) < 1 Then
''            ContRec.Filter = "SAUDA ='" & LFSauda & "'"
''        End If
'
        Dim LFILTERSTE As String
        LFILTERSTE = ""
        If LenB(LFParty) > 0 Then
            LFILTERSTE = "PARTY ='" & LFParty & "'"
        End If
        If LenB(LFSauda) > 0 Then
            If LFILTERSTE = "" Then
                LFILTERSTE = "SAUDA ='" & LFSauda & "'"
            Else
                LFILTERSTE = LFILTERSTE & " AND SAUDA ='" & LFSauda & "'"
            End If
        End If
        If LenB(LFBroker) > 0 Then
            If LFILTERSTE = "" Then
                LFILTERSTE = "CODE ='" & LFBroker & "'"
            Else
                LFILTERSTE = LFILTERSTE & " AND CODE ='" & LFBroker & "'"
            End If
        End If
        If cmbtype.ListIndex = 1 Then 'TRADE
            If LFILTERSTE = "" Then
                LFILTERSTE = "ContractType <> 'O'"
            Else
                LFILTERSTE = LFILTERSTE & " AND ContractType <> 'O'"
            End If
        ElseIf cmbtype.ListIndex = 2 Then 'ORDER
            If LFILTERSTE = "" Then
                LFILTERSTE = "ContractType = 'O'"
            Else
                LFILTERSTE = LFILTERSTE & " AND ContractType = 'O'"
            End If
        End If
        If LFILTERSTE <> "" Then
            ContRec.Filter = LFILTERSTE
        End If
                
        If OptOrder.Value Then 'order
        
            If LenB(LlistParties) > 0 Or LenB(LListSaudas) > 0 Then
                mysql = " SELECT A.Party,B.Name,A.Sauda,A.CONTYPE AS BS, A.Qty, A.Rate,A.CONCODE AS Code,C.NAME AS Broker, A.BROKAMT AS ConRate,A.RATE-A.BROKAMT AS DiffRate, "
                mysql = mysql & " A.ConNo,A.ConTime,A.ROWNO1 AS TradeNo,A.DATAIMPORT,A.EXCODE  "
                mysql = mysql & " FROM ORD_D A, ACCOUNTD B, ACCOUNTD AS C "
                mysql = mysql & " WHERE A.COMPCODE=" & GCompCode & " AND A.COMPCODE = C.COMPCODE AND A.CONCODE = C.AC_CODE  AND A.COMPCODE = B.COMPCODE "
                mysql = mysql & " AND A.PARTY=B.AC_CODE  AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' AND PERCONT ='P' "
                If LenB(LFExCode) > 0 Then mysql = mysql & " AND A.EXID =" & LFExID & ""
                If LenB(LlistParties) > 0 Then mysql = mysql & " AND B.AC_CODE IN (" & LlistParties & ")"
                If LenB(LListSaudas) > 0 Then mysql = mysql & " AND A.SAUDA IN( " & LListSaudas & ")"
                mysql = mysql & " ORDER BY CONNO DESC "
                
                Set ContRec = Nothing: Set ContRec = New ADODB.Recordset
                ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            End If
            
            Set DataGridOrder.DataSource = ContRec
            DataGridOrder.ReBind
            DataGridOrder.Refresh
            
        ElseIf OptTrade.Value Then 'trade
        
            If LenB(LlistParties) > 0 Or LenB(LListSaudas) > 0 Then
                mysql = " SELECT A.Party,B.Name,A.Sauda,A.CONTYPE AS BS,IIF(A.EXCODE='NSE',convert(decimal(18,4),A.QTY/S.REFLOT),convert(decimal(18,4),A.Qty)) AS 'Lot' ,a.qty, A.Rate,A.CONCODE AS Code,C.NAME AS Broker, A.BROKAMT AS ConRate,CASE A.CONTYPE when 'B' THEN A.RATE-A.BROKAMT WHEN 'S' THEN A.BROKAMT-A.RATE END  AS DiffRate, "
                mysql = mysql & " A.ConNo,A.ConTime,A.ROWNO1 AS TradeNo,A.DATAIMPORT,A.EXCODE  "
                mysql = mysql & " FROM CTR_D A, ACCOUNTD B, ACCOUNTD AS C ,SAUDAMAST S"
                mysql = mysql & " WHERE A.COMPCODE=" & GCompCode & " AND A.COMPCODE = C.COMPCODE AND A.CONCODE = C.AC_CODE  AND A.COMPCODE = B.COMPCODE "
                mysql = mysql & " AND A.PARTY=B.AC_CODE  AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' AND PERCONT ='P' AND A.SAUDAID=S.SAUDAID "
                If LenB(LFExCode) > 0 Then mysql = mysql & " AND A.EXID =" & LFExID & ""
                If LenB(LlistParties) > 0 Then mysql = mysql & " AND B.AC_CODE IN (" & LlistParties & ")"
                If LenB(LListSaudas) > 0 Then mysql = mysql & " AND A.SAUDA IN( " & LListSaudas & ")"
                mysql = mysql & " ORDER BY CONNO DESC "
                
                Set ContRec = Nothing:            Set ContRec = New ADODB.Recordset
                ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            End If
            Set DataGrid1.DataSource = ContRec
            DataGrid1.ReBind
            DataGrid1.Refresh
            
        End If
    End If
    
    Call Resize_Grid
    TxtDiffAmt.text = "0.00"
    
    If OptTrade.Value Then 'trade
            mysql = " EXEC Get_Diffamt " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ""
        Set LShreeRec = Nothing
        Set LShreeRec = New ADODB.Recordset
        LShreeRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not LShreeRec.EOF Then
            If Not IsNull(LShreeRec!DIFFAMT) Then TxtDiffAmt.text = Format(Val(LShreeRec!DIFFAMT), "0.00")
        End If
        Set LShreeRec = Nothing
    End If
    
'End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    'SORTING ***
    Dim LSortOrder As String
    On Error GoTo Error1
    DataGrid1.MarqueeStyle = dbgHighlightCell
    DoEvents
    If Left$(Label13.Caption, 1) = "A" Then
        LSortOrder = DataGrid1.Columns.Item(ColIndex).DataField & "  DESC"
        ContRec.Sort = ("" & LSortOrder & "")
        Label13.Caption = "Desc. ORDER ON  " & DataGrid1.Columns.Item(ColIndex).Caption
    Else
        LSortOrder = DataGrid1.Columns.Item(ColIndex).DataField & "  ASC"
        ContRec.Sort = ("" & LSortOrder & "")
        Label13.Caption = "Asc. ORDER ON " & DataGrid1.Columns.Item(ColIndex).Caption
    End If
    DoEvents
    If Not ContRec.EOF Then Set DataGrid1.DataSource = ContRec: DataGrid1.ReBind: DataGrid1.Refresh
    Call Resize_Grid
    
Error1:    Exit Sub
End Sub
Private Sub DataGridOrder_HeadClick(ByVal ColIndex As Integer)
    'SORTING ***
    Dim LSortOrder As String
    On Error GoTo Error1
    '---------------
    DtpCondate.Enabled = True
    ChkShowContract.Enabled = True
    InstCombo.Enabled = True
    DComboExchange.Enabled = True
    '---------------
    
    DataGridOrder.MarqueeStyle = dbgHighlightCell
    DoEvents
    If Left$(Label13.Caption, 1) = "A" Then
        LSortOrder = DataGridOrder.Columns.Item(ColIndex).DataField & "  DESC"
        ContRec.Sort = ("" & LSortOrder & "")
        Label13.Caption = "Desc. ORDER ON  " & DataGridOrder.Columns.Item(ColIndex).Caption
    Else
        LSortOrder = DataGridOrder.Columns.Item(ColIndex).DataField & "  ASC"
        ContRec.Sort = ("" & LSortOrder & "")
        Label13.Caption = "Asc. ORDER ON " & DataGridOrder.Columns.Item(ColIndex).Caption
    End If
    DoEvents
    If Not ContRec.EOF Then Set DataGridOrder.DataSource = ContRec: DataGridOrder.ReBind: DataGrid1.Refresh
    Call Resize_Grid
    
Error1:    Exit Sub
End Sub
Private Sub CANCEL_REC()

Dim SREC As ADODB.Recordset:        Dim PREC As ADODB.Recordset
'Dim LBSaudas As String:                     Dim LBParties As String:            Dim LBItems As String
ChkCarry.Value = 0
LFParty = vbNullString:                     LFExCode = vbNullString:            LFSauda = vbNullString
LFExID = 0
LFBroker = vbNullString:                    TxtCalval.text = vbNullString:      TxtValue.text = vbNullString:
LFExCode = vbNullString:                    Text15.text = vbNullString:         TxtBrokerConfirm.text = "0"
TxtSettleRate.text = vbNullString:          Frame2.Enabled = False:             TxtConfirm.text = "0"
TxtFileType.text = "0":                     SSTab1.Tab = 0:                     LListSaudas = vbNullString
LlistParties = vbNullString:                ListView1.ListItems.Clear:          ListView2.ListItems.Clear
TxtOptType.text = "CE"
TxtStrike.text = vbNullString
Label12.Caption = "Updating Bills"
GETMAIN.Toolbar1_Buttons(6).Enabled = False
    On Error GoTo err1
    CmdAdd.Enabled = True:                      CmdModify.Enabled = True: CmdSave.Enabled = True
    CmdCancel.Enabled = False:                  DtpCondate.Enabled = True
    DComboExchange.BoundText = vbNullString:    DataCombo2.BoundText = vbNullString
    DComboTSauda.BoundText = vbNullString:      DataCombo4.BoundText = vbNullString
    DComboParty.BoundText = vbNullString:       DComboSauda.BoundText = vbNullString
    DComboBroker.BoundText = vbNullString:      TxtBrokerCode.text = vbNullString
    DComboCode.BoundText = vbNullString
    TxtConNo.text = vbNullString:               TxtPtyCode.text = vbNullString
    TxtSaudaID.text = vbNullString
    
    DComboExchange.Enabled = True
    If ExRec.RecordCount = 1 Then
        ExRec.MoveFirst
        DComboExchange.BoundText = ExRec!excode
    End If
    ChkShowContract.Enabled = True
    TxtConNo.Locked = False
    Frame2.Enabled = False
    LFBPress = 0:
    
    If OptOrder.Value Then
        SaveCalled = False
    End If
    
    If SaveCalled = True Then
        Frame1.Enabled = False:
        DataGrid1.Enabled = False
        DataGrid2.Enabled = False
        DataGridOrder.Enabled = False
        Call RATE_TEST(DtpCondate.Value, , , frmcont4)
        
        Call Shree_Posting(DateValue(DtpCondate.Value))
        
        CNNERR = True:                 Cnn.BeginTrans
        Call Update_Charges(LBillParties, LBillExCodes, LBillSaudas, vbNullString, DtpCondate.Value, DtpCondate.Value, True)
        GETMAIN.Label1.Caption = "Updating Brokerage Rate Itemwise Complete"
        Cnn.CommitTrans
        Cnn.BeginTrans
        If BILL_GENERATION(DtpCondate.Value, GFinEnd, LBillSaudas, LBillParties, LBillExCodes) Then
            Cnn.CommitTrans
            CNNERR = False
        End If
        'Call Chk_Billing
        DataGrid1.Enabled = True
        DataGrid2.Enabled = True
        DataGridOrder.Enabled = True
    End If
    LBillParties = vbNullString:    LBillExCodes = vbNullString
    LBillSaudas = vbNullString:     LBillItems = vbNullString
    SaveCalled = False:             Frame1.Enabled = True
    Frame3.Enabled = True:          Frame10.Enabled = True
    Frame2.Enabled = True:          GETMAIN.Toolbar1_Buttons(6).Enabled = True
    DtpCondate.Enabled = True:      DtpCondate.SetFocus
    ChkAppBrok.Value = 1
    Label12.Caption = "Bills Updated"
    Call DATA_GRID_REFRESH
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
        Cnn.RollbackTrans: CNNERR = False
        Frame1.Enabled = True
        
    End If
End Sub

Private Sub TxtLot_Validate(Cancel As Boolean)
If Val(TxtLot) > 0 Then TxtQty.text = CStr(Val(TxtLot.text) * Val(TxtRefLot.text))
End Sub
Private Sub DtpCondate_Validate(Cancel As Boolean)
    Dim NRec As ADODB.Recordset
  '  Dim lstr  As String
    Dim MDt As String
    
    If (GSysLockDt >= DtpCondate.Value) Then
        CmdAdd.Enabled = False
        CmdModify.Enabled = False
        CmdSave.Enabled = False
        If MsgBox("System is locked till date " & GSysLockDt & vbNewLine & "Do you still want to view document?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    Else
        CmdAdd.Enabled = True
        CmdModify.Enabled = True
        CmdSave.Enabled = True
    End If
        
    vcDTP2.MinDate = DtpCondate.Value
    vcDTP2.Value = DtpCondate.Value + 90
    MDt = Format(Now, "dd/mm/yyyy")
    
    If DtpCondate.Value > DateValue(MDt) Then
        MsgBox "Trade Date Is Greater than Current Date"
    End If
    
    Call FillLFPartyCombo
    Call FillFSaudaCombo
    Call FillFBrokerCombo
    FillTradeSaudaCombo
End Sub

Private Sub TxtStrike_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub

Private Sub TxtStrike_Validate(Cancel As Boolean)
TxtStrike = Format(TxtStrike.text, "0.00")
Call Connect_TSaudaCombo
End Sub

Private Sub vcDTP2_Validate(Cancel As Boolean)
    Dim TRec As ADODB.Recordset:        Dim LSaudaCode As String:       Dim LFLAG As Boolean:       Dim LTExCode  As String
    Dim LExID As Integer: Dim LItemID As Integer
    mysql = "SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE ='" & DataCombo5.BoundText & "' AND MATURITY='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        MsgBox "Creating New Contract for " & vcDTP2.Value & " Prompt Date"
        LSaudaCode = DataCombo5.text & " PD " & vcDTP2.Value
        ItemRec.MoveFirst
        ItemRec.Find "ITEMCODE ='" & DataCombo5.BoundText & "'"
        If ItemRec.EOF Then
            MsgBox "Invalid Item"
            Exit Sub
        Else
            LTExCode = ItemRec!EXCHANGECODE
        End If
        LExID = Get_ExID(LTExCode)
        LItemID = Get_ITEMID(LItemCodeDBCombo)
        Call PInsert_Saudamast(LSaudaCode, LSaudaCode, LItemCodeDBCombo, vcDTP2.Value, 1, 1, 0, "FUT", vbNullString, 0, LTExCode, 1, LExID, LItemID)
        LFLAG = True
    Else
        LSaudaCode = TRec!saudacode
    End If
    Set TRec = Nothing
    Call FillTradeSaudaCombo
    DComboTSauda.BoundText = LSaudaCode
End Sub
Private Sub Mod_Rec()
    If ContRec.RecordCount > 0 Then
        Frame2.Enabled = True:              CmdAdd.Enabled = False
        CmdModify.Enabled = False:          CmdCancel.Enabled = True
        DtpCondate.Enabled = False:         DComboExchange.Enabled = False
        ChkShowContract.Enabled = False:    TxtConNo.text = vbNullString
        TxtConNo.Locked = False:            TxtQty.Locked = False
        TxtQty.text = vbNullString:         TxtRate.text = vbNullString
        TxtConRate.text = vbNullString:
        LFBPress = 2
        Call Connect_TSaudaCombo
        Label12.Caption = "Modify Trades"
        TxtConNo.SetFocus
    Else
        MsgBox "No Records to Modify "
    End If
End Sub
Private Sub Get_Value()
    Call Connect_TSaudaCombo
    TxtSaudaID.text = vbNullString
    If ChkShowContract.Value = 1 Then
        AllSaudaRec.Filter = adFilterNone
        If Not AllSaudaRec.EOF Then
            AllSaudaRec.MoveFirst
            AllSaudaRec.Find "SAUDACODE ='" & DComboTSauda.BoundText & "'"
            If AllSaudaRec.EOF Then
                MsgBox "Invalid Contract"
            Else
                TxtCalval.text = Format(AllSaudaRec!lot, "0.00")
                Text15.text = Format(Val(TxtCalval.text) * Val(TxtQty.text), "0.00")
                TxtValue.text = Format(Val(Text15.text) * Val(TxtRate.text), "0.00")
                TxtRefLot.text = Format(AllSaudaRec!REFLOT, "0.00")
            End If
        Else
            MsgBox "No Contract"
        End If
    Else
        SaudaRec.Filter = adFilterNone
        If SaudaRec.RecordCount > 0 Then SaudaRec.MoveFirst
        SaudaRec.Find "SAUDACODE ='" & DComboTSauda.BoundText & "'", , adSearchForward
        If SaudaRec.EOF Then
            MsgBox "Invalid Contract"
            Sendkeys "%{DOWN}"
        Else
            If SaudaRec!excode = "MCX" Or SaudaRec!excode = "NSE" Then
                If SaudaRec!LOTWISE = "Y" Then
                    TxtCalval.text = Format(SaudaRec!TRADEABLELOT, "0.00")
                Else
                    TxtCalval.text = Format(SaudaRec!lot, "0.00")
                End If
            Else
                TxtCalval.text = Format(SaudaRec!lot, "0.00")
            End If
            Text15.text = Format(Val(TxtCalval.text) * Val(TxtQty.text), "0.00")
            TxtValue.text = Format(Val(Text15.text) * Val(TxtRate.text), "0.00")
            TxtRefLot.text = Format(SaudaRec!REFLOT, "0.00")
            TxtSaudaID.text = SaudaRec!SAUDAID
        End If
    End If
    If DComboTSauda.Enabled Then
    DComboTSauda.SetFocus
    End If
'    If LenB(TxtSaudaID.text) > 1 Then
'        DComboTSauda.BoundText = TxtSaudaID.text
'        DComboTSauda.SetFocus
 '   End If
End Sub
Public Sub FillTradeSaudaCombo()
    'MYSQL = " EXEC Get_ScriptContract " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "','" & LFExCode & "','" & InstCombo.text & "'"
    'Set AllSaudaRec = Nothing
    'Set AllSaudaRec = New ADODB.Recordset
    'AllSaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    '
    'MYSQL = "EXEC Get_SaudaContract " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ",'" & InstCombo.text & "'"
    'Set SaudaRec = Nothing
    'Set SaudaRec = New ADODB.Recordset
    'SaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    
    Call Connect_TSaudaCombo
End Sub

Public Sub Connect_TSaudaCombo()

If ChkShowContract.Value = 1 Then
    TxtOptType.text = ""
    'mysql = " EXEC Get_ScriptContract " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "','" & LFExCode & "','" & InstCombo.text & "','" & TxtOptType.text & "'," & Val(TxtStrike.text) & ""
    If InstCombo.text = "All" Then
        mysql = " EXEC Get_ScriptContract " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "','" & LFExCode & "','" & InstCombo.text & "','" & TxtOptType.text & "'," & Val(TxtStrike.text) & ""
    Else
        mysql = " SELECT S.SAUDACODE,S.SAUDANAME,S.ITEMCODE,C.ITEMNAME,S.MATURITY,S.EXCODE,S.INSTTYPE,S.OPTTYPE,S.STRIKEPRICE,C.LOT, S.LOT AS TRADEABLELOT,S.BROKLOT,S.REFLOT,EX.LOTWISE,S.EX_SYMBOL "
        mysql = mysql & " FROM SCRIPTMASTER AS S,EXMAST AS EX ,CONTRACTMASTER AS C WHERE EX.EXCODE =S.EXCODE AND EX.COMPCODE =" & GCompCode & ""
        mysql = mysql & " AND C.EX_SYMBOL=S.EX_SYMBOL AND C.EXCODE=S.EXCODE AND S.MATURITY>='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
        If LenB(LFExCode) > 0 Then mysql = mysql & " AND EX.EXCODE ='" & LFExCode & "'"
        If InstCombo.text <> "All" Then mysql = mysql & " AND  S.INSTTYPE ='" & InstCombo.text & "'"
        If InstCombo.text = "OPT" Then
            If LenB(TxtOptType.text) > 0 Then mysql = mysql & " AND OPTTYPE ='" & TxtOptType.text & "'"
            If Val(TxtStrike.text & vbNullString) <> 0 Then mysql = mysql & "  AND STRIKEPRICE =" & Val(TxtStrike.text & vbNullString) & ""
        End If
        mysql = mysql & " ORDER BY S.ITEMCODE ,S.INSTTYPE,S.OPTTYPE,S.STRIKEPRICE,S.MATURITY  "
    End If

    Set AllSaudaRec = Nothing
    Set AllSaudaRec = New ADODB.Recordset
    AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AllSaudaRec.EOF Then
        Set DComboTSauda.RowSource = AllSaudaRec:
        DComboTSauda.BoundColumn = "SAUDACODE"
        DComboTSauda.ListField = "SAUDANAME"
    Else
        Set DComboTSauda.RowSource = Nothing
        MsgBox "No Contract Exists "
        If Frame2.Enabled Then
            TxtPtyCode.SetFocus
        Else
            Frame10.Enabled = True
            DtpCondate.SetFocus
        End If
    End If
    
Else
        mysql = "EXEC Get_SaudaContract " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ""
        Set SaudaRec = Nothing
        Set SaudaRec = New ADODB.Recordset
        SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Set DComboTSauda.RowSource = SaudaRec
        DComboTSauda.BoundColumn = "SAUDACODE"
        DComboTSauda.ListField = "SAUDANAME"
        If Not SaudaRec.EOF Then
        
        Set DComboTSauda.RowSource = SaudaRec:
            DComboTSauda.BoundColumn = "SAUDACODE"
            DComboTSauda.ListField = "SAUDANAME"
        Else
            If LFExID <> -1 Then  ' If order then LFExID = -1
                MsgBox "No Contract Exists "
                Set DComboTSauda.RowSource = Nothing
                If Frame2.Enabled Then
                    TxtPtyCode.SetFocus
                Else
                    Frame10.Enabled = True
                    DtpCondate.SetFocus
                End If
            End If
        End If
    
    
End If

End Sub

Private Sub Resize_Grid()

If GShowLot <> "Y" Then
    DataGrid1.Columns(4).Visible = False
End If

If GCINNo = "2000" Then
        DataGrid1.Columns(1).Width = 2500:              DataGrid1.Columns(2).Width = 800
        DataGrid1.Columns(3).Width = 900:               DataGrid1.Columns(4).Width = 3000
        DataGrid1.Columns(5).Width = 1000:              DataGrid1.Columns(6).Width = 1200
        DataGrid1.Columns(7).Width = 1200:              DataGrid1.Columns(8).Width = 1200
        
        DataGrid1.Columns(7).Alignment = dbgRight:      DataGrid1.Columns(8).Alignment = dbgRight
        DataGrid1.Columns(2).Alignment = dbgCenter:     DataGrid1.Columns(9).Alignment = dbgRight
        DataGrid1.Columns(10).Alignment = dbgRight:     DataGrid1.Columns(3).Alignment = dbgCenter
        DataGrid1.Columns(5).NumberFormat = "0.00":     DataGrid1.Columns(5).Alignment = dbgRight:
        DataGrid1.Columns(6).Alignment = dbgRight
    Else
        '0  Party,         '1  Name,:               '2  Sauda,:               '3  CONTYPE AS Type ,   '4  Qty,
        '5  Rate,          '6  CONCODE AS Code,     '7 NAME AS Broker,        '8 BROKAMT AS ConRate, '9 RATE-A.BROKAMT AS DiffRate,
        '10 conno          '11  contime              12  rowno1                13'dataimport
                
        If OptTrade.Value Then 'Trade
            DataGrid1.Columns(0).Width = 900
            DataGrid1.Columns(1).Width = 2300:              DataGrid1.Columns(2).Width = 3400
            DataGrid1.Columns(3).Width = 500:               DataGrid1.Columns(4).Width = 1000
            DataGrid1.Columns(5).Width = 1200:              DataGrid1.Columns(6).Width = 1200: DataGrid1.Columns(7).Width = 800
            DataGrid1.Columns(8).Width = 2000:              DataGrid1.Columns(9).Width = 1200
            DataGrid1.Columns(10).Width = 1200:              DataGrid1.Columns(11).Width = 900
            DataGrid1.Columns(12).Width = 1300:             DataGrid1.Columns(13).Width = 900:
            DataGrid1.Columns(14).Visible = False:          DataGrid1.Columns(3).Alignment = dbgCenter
            DataGrid1.Columns(5).Alignment = dbgRight:      DataGrid1.Columns(6).Alignment = dbgRight: DataGrid1.Columns(7).Alignment = dbgCenter
            
            DataGrid1.Columns(8).Alignment = dbgLeft:       DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(4).Alignment = dbgRight:      DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(11).Alignment = dbgRight
            DataGrid1.Columns(5).NumberFormat = "0.0000":  DataGrid1.Columns(6).NumberFormat = "0.0000":    DataGrid1.Columns(10).NumberFormat = "0.00"
            DataGrid1.Columns(9).NumberFormat = "0.0000":     DataGrid1.Columns(12).Alignment = dbgLeft
            DataGrid1.Columns(5).Alignment = dbgRight:
            
            If GQty_Decimal = True Then
                DataGrid1.Columns(5).NumberFormat = "0.0000"
                DataGrid1.Columns(11).NumberFormat = "0.0000"
                DataGrid1.Columns(4).NumberFormat = "0.00"
                DataGrid1.Columns(10).NumberFormat = "0.00"
            End If
        ElseIf OptOrder.Value Then 'Order
            DataGridOrder.Columns(0).Width = 900: DataGridOrder.Columns(1).Width = 2300: DataGridOrder.Columns(2).Width = 3400
            DataGridOrder.Columns(3).Width = 500: DataGridOrder.Columns(4).Width = 1000: DataGridOrder.Columns(5).Width = 1200:
            
            DataGridOrder.Columns(6).Width = 0: DataGridOrder.Columns(7).Width = 0: DataGridOrder.Columns(8).Width = 0
            DataGridOrder.Columns(9).Width = 0: DataGridOrder.Columns(10).Width = 0: DataGridOrder.Columns(11).Width = 0:
            DataGridOrder.Columns(12).Width = 0: DataGridOrder.Columns(3).Alignment = dbgCenter
            
            DataGridOrder.Columns(6).Visible = False: DataGridOrder.Columns(7).Visible = False
            DataGridOrder.Columns(8).Visible = False: DataGridOrder.Columns(9).Visible = False
            DataGridOrder.Columns(10).Visible = False: DataGridOrder.Columns(11).Visible = False
            DataGridOrder.Columns(12).Visible = False: DataGridOrder.Columns(13).Visible = False: DataGridOrder.Columns(14).Visible = False: DataGridOrder.Columns(15).Visible = False
            
            
            DataGridOrder.Columns(5).Alignment = dbgRight:
'            DataGridOrder.Columns(6).Alignment = dbgCenter
'            DataGridOrder.Columns(7).Alignment = dbgLeft:       DataGridOrder.Columns(8).Alignment = dbgRight
'            DataGridOrder.Columns(4).Alignment = dbgRight:      DataGridOrder.Columns(9).Alignment = dbgRight
'            DataGridOrder.Columns(10).Alignment = dbgRight
            DataGridOrder.Columns(5).NumberFormat = "0.0000":     'DataGridOrder.Columns(9).NumberFormat = "0.00"
            'DataGridOrder.Columns(8).NumberFormat = "0.0000":     DataGridOrder.Columns(11).Alignment = dbgLeft
            
            If GQty_Decimal = True Then
                DataGridOrder.Columns(5).NumberFormat = "0.0000"
                DataGridOrder.Columns(11).NumberFormat = "0.0000"
                DataGridOrder.Columns(4).NumberFormat = "0.00"
                DataGridOrder.Columns(10).NumberFormat = "0.00"
            End If
        End If
    End If
End Sub

Private Sub Get_Trade_Details(LZConNo As Long)
    Dim TRec As ADODB.Recordset
    Dim TRec1 As ADODB.Recordset
    Dim LSaudaID As Long
    
    If OptTrade.Value Then
        mysql = " EXEC Get_Trade " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & "," & LZConNo & ""
    ElseIf OptOrder.Value Then
        mysql = " EXEC Get_Trade " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "',-1," & LZConNo & ""
    End If
    
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.RecordCount > 1 Then
        MsgBox " Duplicated Trade Nos Pls contact Sauda Support Staff"
        Exit Sub
    End If
    If Not TRec.EOF Then
        TxtConNo.text = TRec!CONNO
        Text7.text = (TRec!ROWNO1 & vbNullString)
        ChkAppBrok.Value = IIf(TRec!BROKFLAG = "Y", 1, 0)
        TxtPtyCode.text = TRec!PARTY
        DataCombo2.BoundText = TRec!PARTY
        TxtRate.text = Format(TRec!Rate, "0.0000")
        TxtConRate.text = Format(TRec!BROKAMT, "0.0000")
        TxtBrokerCode.text = TRec!CONCODE
        DataCombo4.BoundText = TRec!CONCODE
        DComboTSauda.BoundText = TRec!Sauda
        TxtSaudaID.text = TRec!SAUDAID
        If TRec!CONTYPE = "B" Then
            TxtContype.text = "Buy"
        Else
            TxtContype.text = "Sel"
        End If
        TxtConfirm.text = TRec!CONFIRM
        TxtQty.text = TRec!QTY
        LOldEXCode = TRec!excode
        LOldParty = TRec!PARTY
        LOldBroker = TRec!CONCODE
        LOldQty = TRec!QTY
        LOldContype = TRec!CONTYPE
        LOldRate = TRec!Rate
        LOldRate2 = TRec!BROKAMT
        LOldSauda = TRec!Sauda
        LOldConno = TRec!CONNO
        TxtFileType.text = TRec!filetype
        
        If OptTrade.Value Then
            If LenB(LBillExCodes) < 1 Then
                LBillExCodes = str(TRec!EXID)
            Else
                If LStr_Exists(LBillExCodes, str(TRec!EXID)) < 1 Then LBillExCodes = LBillExCodes & "," & str(TRec!EXID) & ""
            End If
        ElseIf OptOrder.Value Then
            LBillExCodes = "0"
            CmbOrderStatus.ListIndex = TRec!EXID
        End If
        
        If LenB(LBillParties) < 1 Then
            If TRec!PARTY <> TRec!CONCODE Then
                LBillParties = "'" & TRec!PARTY & "','" & TRec!CONCODE & "'"
            Else
                LBillParties = "'" & TRec!PARTY & "'"
            End If
        Else
            If InStr(LBillParties, "'" & TRec!PARTY & "'") < 1 Then LBillParties = LBillParties & ",'" & TRec!PARTY & "'"
            If InStr(LBillParties, "'" & TRec!CONCODE & "") < 1 Then LBillParties = LBillParties & ",'" & TRec!CONCODE & "'"
        End If
        If LenB(LBillItems) < 1 Then
            LBillItems = "'" & TRec!ITEMCODE & "'"
        Else
            If InStr(LBillItems, TRec!ITEMCODE) < 1 Then LBillItems = LBillItems & "," & "'" & TRec!ITEMCODE & "'"
        End If
        
        If LenB(LBillSaudas) < 1 Then
            LBillSaudas = Trim(str(TRec!SAUDAID))
        Else
            If LStr_Exists(LBillSaudas, TRec!SAUDAID) = False Then LBillSaudas = LBillSaudas & "," & Trim(str(TRec!SAUDAID)) & ""
        End If
        LSaudaID = TRec!SAUDAID
        'LSaudaId = Get_SaudaID(TRec!Sauda)
        TxtSettleRate.text = Format(SDCLRATE(LSaudaID, DtpCondate.Value, "C"), "0.00")
        Call Get_Value:
        mysql = "SELECT CONFIRM FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
        mysql = mysql & " AND CONNO=" & TRec!CONNO & " AND PARTY='" & TRec!CONCODE & "'"
        Set TRec1 = Nothing
        Set TRec1 = New ADODB.Recordset
        TRec1.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec1.EOF Then
            TxtBrokerConfirm.text = TRec1!CONFIRM
        Else
            TxtBrokerConfirm.text = "0"
        End If
        Set TRec = Nothing
        Set TRec1 = Nothing
    End If
    If Val(TxtRefLot.text) <> 0 Then TxtLot = CStr(Val(TxtQty.text) / Val(TxtRefLot.text))
End Sub
Private Sub FillFSaudaCombo()
    Set LFSaudaRec = Nothing
    Set LFSaudaRec = New ADODB.Recordset
    mysql = "EXEC Get_SaudaCtr_d " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ",'" & LFParty & "'"
    LFSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LFSaudaRec.EOF Then
        DComboSauda.Enabled = True:                 Set DComboSauda.RowSource = LFSaudaRec
        DComboSauda.BoundColumn = "SAUDACODE":      DComboSauda.ListField = "SAUDANAME"
    Else
        DComboSauda.Enabled = False
    End If
    ListView1.Visible = False
    ListView1.ListItems.Clear
    Do While Not LFSaudaRec.EOF
        ListView1.ListItems.Add , , LFSaudaRec!saudacode
        LFSaudaRec.MoveNext
    Loop
    ListView1.Visible = True
    
End Sub
Private Sub FillFBrokerCombo()
    
    If LenB(DComboSauda.BoundText) <> 0 Then
        LFSauda = DComboSauda.BoundText
    Else
        LFSauda = vbNullString
    End If
    Set LFBrokerRec = Nothing
    Set LFBrokerRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT A.AC_CODE,A.NAME FROM ACCOUNTD  AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
    mysql = mysql & " AND A.AC_CODE  =B.CONCODE  AND B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    If LenB(LFExCode) > 0 Then mysql = mysql & " AND B.EXID=" & LFExID & ""
    If LenB(LFParty) > 0 Then mysql = mysql & " AND B.PARTY ='" & LFParty & "'"
    If LenB(LFSauda) > 0 Then mysql = mysql & " AND B.SAUDA ='" & LFSauda & "'"
    mysql = mysql & " ORDER BY A.NAME"
    LFBrokerRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LFBrokerRec.EOF Then
        DComboBroker.Enabled = True
        Set DComboBroker.RowSource = LFBrokerRec
        DComboBroker.BoundColumn = "AC_CODE"
        DComboBroker.ListField = "NAME"
    Else
        DComboBroker.Enabled = False
    End If
    Call DATA_GRID_REFRESH
End Sub

Private Sub FillLFPartyCombo()
    Set LFPartyRec = Nothing
    
    Set LFPartyRec = New ADODB.Recordset
    mysql = " EXEC Get_PartyCtr_d " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ""
    LFPartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LFPartyRec.EOF Then
        DComboParty.Enabled = True
        Set DComboParty.RowSource = LFPartyRec
        DComboParty.BoundColumn = "AC_CODE"
        DComboParty.ListField = "NAME"
        DComboCode.Enabled = True
        Set DComboCode.RowSource = LFPartyRec
        DComboCode.BoundColumn = "AC_CODE"
        DComboCode.ListField = "AC_CODE"
    Else
        DComboParty.Enabled = False
        DComboCode.Enabled = False
    End If
    ListView2.Visible = False
    ListView2.ListItems.Clear
    Do While Not LFPartyRec.EOF
        ListView2.ListItems.Add , , LFPartyRec!AC_CODE
        ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , LFPartyRec!NAME
        LFPartyRec.MoveNext
    Loop
    ListView2.Visible = True
End Sub

