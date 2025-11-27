VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ExMFrm 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   10935
   ClientLeft      =   0
   ClientTop       =   -1905
   ClientWidth     =   15960
   ForeColor       =   &H00400000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame17 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0FF&
      Height          =   1335
      Left            =   6600
      TabIndex        =   124
      Top             =   3360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         TabIndex        =   125
         Top             =   120
         Width           =   4575
         Begin VB.TextBox TxtAdminPass 
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
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   126
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label34 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Admin Password"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1200
            TabIndex        =   127
            Top             =   120
            Width           =   2295
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   3600
      TabIndex        =   44
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General "
      TabPicture(0)   =   "ExMFrm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame10"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Accounts"
      TabPicture(1)   =   "ExMFrm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame11"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tax Rates"
      TabPicture(2)   =   "ExMFrm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Client Wise Broker"
      TabPicture(3)   =   "ExMFrm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame16"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Trade File Details"
      TabPicture(4)   =   "ExMFrm.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame13"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame12"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame12 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   120
         TabIndex        =   121
         Top             =   2520
         Width           =   10095
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "ExMFrm.frx":008C
            Height          =   5895
            Left            =   120
            TabIndex        =   122
            Top             =   0
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   10398
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   21
            AllowAddNew     =   -1  'True
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
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Exchange File "
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
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   1575
         Left            =   120
         TabIndex        =   101
         Top             =   840
         Width           =   10095
         Begin VB.CommandButton Command2 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   7080
            TabIndex        =   123
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   9120
            TabIndex        =   116
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtFolderName 
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
            Left            =   1560
            TabIndex        =   103
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox TxtBroker 
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
            Left            =   5640
            TabIndex        =   104
            Top             =   120
            Width           =   855
         End
         Begin VB.ComboBox PartyAsCombo 
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
            ItemData        =   "ExMFrm.frx":00A1
            Left            =   1560
            List            =   "ExMFrm.frx":00AE
            TabIndex        =   106
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox TxtClient 
            Enabled         =   0   'False
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
            Left            =   5640
            TabIndex        =   107
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   8280
            TabIndex        =   115
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox BrokerTypeCombo 
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
            ItemData        =   "ExMFrm.frx":00D0
            Left            =   1560
            List            =   "ExMFrm.frx":00EF
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox TxtFileId 
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
            Left            =   840
            TabIndex        =   102
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   5640
            TabIndex        =   111
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton CmdMod 
            Caption         =   "Mod"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   6360
            TabIndex        =   113
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TxtFileType 
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
            Left            =   4680
            TabIndex        =   110
            Top             =   1080
            Width           =   855
         End
         Begin MSDataListLib.DataCombo BrokerCombo 
            Height          =   390
            Left            =   6720
            TabIndex        =   105
            Top             =   120
            Width           =   3255
            _ExtentX        =   5741
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
         Begin MSDataListLib.DataCombo ClientCombo 
            Height          =   390
            Left            =   6720
            TabIndex        =   108
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   688
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Folder"
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
            TabIndex        =   120
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Broker Ac"
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
            Left            =   4560
            TabIndex        =   119
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Brok Type"
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
            TabIndex        =   118
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Party As"
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
            TabIndex        =   117
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Client Ac"
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
            Left            =   4560
            TabIndex        =   114
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "File Type"
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
            Left            =   3720
            TabIndex        =   112
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   -74880
         TabIndex        =   89
         Top             =   5400
         Width           =   10095
         Begin MSDataGridLib.DataGrid ExBrokGrid 
            Bindings        =   "ExMFrm.frx":0150
            Height          =   2895
            Left            =   120
            TabIndex        =   90
            Top             =   120
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5106
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   21
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
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
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Exchange wise Broker Client "
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "brokcode"
               Caption         =   "Broker"
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
               DataField       =   "BrokName"
               Caption         =   "Broker Name"
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
               DataField       =   "ClCode"
               Caption         =   "Client"
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
               DataField       =   "ClName"
               Caption         =   "Client Name"
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
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   3000.189
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   4575
         Left            =   -74880
         TabIndex        =   82
         Top             =   720
         Width           =   10095
         Begin MSComctlLib.ListView ListView1 
            Height          =   3255
            Left            =   120
            TabIndex        =   97
            Top             =   1080
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5741
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
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
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.TextBox TxtBrokCode 
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
            TabIndex        =   83
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton CmdSave2 
            Appearance      =   0  'Flat
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
            Height          =   425
            Left            =   8880
            TabIndex        =   87
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton CmdAdd2 
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
            Height          =   425
            Left            =   6480
            TabIndex        =   85
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton CmdMod2 
            Caption         =   "Mod"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Left            =   7680
            TabIndex        =   86
            Top             =   120
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo BrokerCombo2 
            Height          =   420
            Left            =   1800
            TabIndex        =   84
            Top             =   120
            Width           =   3615
            _ExtentX        =   6376
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
         Begin MSComctlLib.ListView ListView2 
            Height          =   3255
            Left            =   5400
            TabIndex        =   99
            Top             =   1080
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   5741
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
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
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "Selected Client List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   100
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "Client List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   98
            Top             =   600
            Width           =   5055
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Broker Ac"
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
            Left            =   240
            TabIndex        =   88
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   -74880
         TabIndex        =   72
         Top             =   780
         Width           =   10335
         Begin MSDataGridLib.DataGrid ExTaxGrid 
            Bindings        =   "ExMFrm.frx":0165
            Height          =   2895
            Left            =   120
            TabIndex        =   73
            Top             =   120
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   5106
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   21
            AllowAddNew     =   -1  'True
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
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Exchange Wise Tax"
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
      Begin VB.Frame Frame11 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   6375
         Left            =   -74880
         TabIndex        =   56
         Top             =   780
         Width           =   10335
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   6135
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Width           =   10095
            Begin MSDataListLib.DataCombo ContractCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   16
               Top             =   780
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo BankCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   15
               Top             =   240
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo TradindCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   14
               Top             =   240
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo BrokCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   17
               Top             =   780
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo StandingCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   18
               Top             =   5640
               Width           =   3765
               _ExtentX        =   6641
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo TransactionCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   19
               Top             =   4020
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo ShreeCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   20
               Top             =   1860
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo MarginCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   22
               Top             =   2400
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo SrvTaxCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   24
               Top             =   2940
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo CTTCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   26
               Top             =   3480
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo RskMCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   21
               Top             =   1860
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo StamDutyCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   23
               Top             =   2400
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo BankContraCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   25
               Top             =   2940
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo SEBITaxCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   27
               Top             =   3480
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo SBCTaxCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   28
               Top             =   4020
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo InterestCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   29
               Top             =   1320
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo CGSTCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   30
               Top             =   4560
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo SGSTCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   31
               Top             =   4560
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo IGSTCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   32
               Top             =   5100
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo UTTCombo 
               Height          =   420
               Left            =   6240
               TabIndex        =   33
               Top             =   5100
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo ClgCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   34
               Top             =   5640
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo CashCombo 
               Height          =   420
               Left            =   1320
               TabIndex        =   35
               Top             =   1320
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
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
            Begin VB.Label Label35 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Tran Fees"
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
               Height          =   225
               Left            =   5040
               TabIndex        =   93
               Top             =   4080
               Width           =   1110
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label32 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Standing"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   345
               Left            =   5040
               TabIndex        =   92
               Top             =   5760
               Width           =   870
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label31 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Cash"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   225
               Left            =   120
               TabIndex        =   91
               Top             =   1418
               Width           =   1110
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label3 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Clearing Fees"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   5730
               Width           =   1335
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label30 
               BackColor       =   &H00C0E0FF&
               Caption         =   "ICGST"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   5183
               Width           =   1095
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label29 
               BackColor       =   &H00C0E0FF&
               Caption         =   "SGST"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   225
               Left            =   5040
               TabIndex        =   76
               Top             =   4658
               Width           =   1110
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label28 
               BackColor       =   &H00C0E0FF&
               Caption         =   "UTT"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   225
               Left            =   5040
               TabIndex        =   75
               Top             =   5198
               Width           =   1110
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label25 
               BackColor       =   &H00C0E0FF&
               Caption         =   "CGST"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   4643
               Width           =   1095
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "Bank Contra"
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
               Left            =   5040
               TabIndex        =   71
               Top             =   3008
               Width           =   1170
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "Service"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   120
               TabIndex        =   70
               Top             =   3008
               Width           =   675
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "StampDuty"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   5040
               TabIndex        =   69
               Top             =   2468
               Width           =   1020
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "Margin"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   120
               TabIndex        =   68
               Top             =   2468
               Width           =   645
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "Risk Mngt"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   5040
               TabIndex        =   67
               Top             =   1928
               Width           =   945
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "Bank "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   5040
               TabIndex        =   66
               Top             =   315
               Width           =   900
            End
            Begin VB.Label Label4 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Contract"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   863
               Width           =   1095
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "Shree"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   64
               Top             =   1928
               Width           =   525
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "Brokerage "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Index           =   1
               Left            =   5040
               TabIndex        =   63
               Top             =   848
               Width           =   1020
            End
            Begin VB.Label Label6 
               BackColor       =   &H00C0E0FF&
               Caption         =   "CTT Tax"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   3563
               Width           =   1095
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label19 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Trading"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label20 
               BackColor       =   &H00C0E0FF&
               Caption         =   "SEBI Tax"
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
               Height          =   225
               Left            =   5040
               TabIndex        =   60
               Top             =   3578
               Width           =   1110
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label21 
               BackColor       =   &H00C0E0FF&
               Caption         =   "SBC Tax"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   4103
               Width           =   1095
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label22 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Interest"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   225
               Left            =   5040
               TabIndex        =   58
               Top             =   1440
               Width           =   1110
               WordWrap        =   -1  'True
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   3975
         Left            =   -74760
         TabIndex        =   45
         Top             =   780
         Width           =   10095
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Auto Update Nse Fo Lot"
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
            TabIndex        =   128
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   615
            Left            =   120
            TabIndex        =   78
            Top             =   2640
            Width           =   9855
            Begin VB.CheckBox Check3 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Currency Conversion"
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
               TabIndex        =   12
               Top             =   120
               Width           =   2295
            End
            Begin MSDataListLib.DataCombo CurrDCombo 
               Height          =   360
               Left            =   4920
               TabIndex        =   13
               Top             =   120
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   635
               _Version        =   393216
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label24 
               BackColor       =   &H00FFFFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Currency Contract"
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
               Left            =   3000
               TabIndex        =   79
               Top             =   173
               Width           =   1695
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   9855
            Begin VB.CheckBox ChkOptions 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Options"
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
               Left            =   8520
               TabIndex        =   94
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox TxtTradeFile 
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
               Left            =   6360
               TabIndex        =   8
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox Text3 
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
               Left            =   1800
               TabIndex        =   7
               Top             =   600
               Width           =   3375
            End
            Begin VB.ComboBox Combo1 
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
               ItemData        =   "ExMFrm.frx":017A
               Left            =   1800
               List            =   "ExMFrm.frx":01A2
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   120
               Width           =   2055
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Pro As User ID"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   6360
               TabIndex        =   5
               Top             =   120
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.TextBox UidTxt 
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
               Left            =   3840
               MaxLength       =   30
               TabIndex        =   10
               Top             =   1080
               Width           =   1335
            End
            Begin VB.ComboBox Combo2 
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
               ItemData        =   "ExMFrm.frx":0219
               Left            =   4920
               List            =   "ExMFrm.frx":0223
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   120
               Width           =   1335
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Lotwise"
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
               Height          =   255
               Left            =   8520
               TabIndex        =   6
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox TxtCloseFile 
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
               Left            =   1800
               TabIndex        =   9
               Top             =   1080
               Width           =   975
            End
            Begin vcDateTimePicker.vcDTP vcDTP1 
               Height          =   375
               Left            =   8160
               TabIndex        =   11
               Top             =   1080
               Width           =   1455
               _ExtentX        =   2566
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
               Value           =   39688.8027893519
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Trade File "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   5280
               TabIndex        =   80
               Top             =   660
               Width           =   990
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FMC UniqCode"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   330
               Left            =   120
               TabIndex        =   55
               Top             =   660
               Width           =   1575
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Broker Type"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   120
               TabIndex        =   54
               Top             =   150
               Width           =   1170
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Stamp Duty Tax From"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   360
               Left            =   5280
               TabIndex        =   53
               Top             =   1140
               Width           =   2100
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mem Id"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   3000
               TabIndex        =   52
               Top             =   1155
               Width           =   735
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Party as"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   3960
               TabIndex        =   51
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Closing File Type"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   120
               TabIndex        =   50
               Top             =   1125
               Width           =   1590
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   600
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   9855
            Begin VB.TextBox TxtExID 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   405
               Left            =   2640
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   96
               TabStop         =   0   'False
               Top             =   120
               Width           =   855
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   840
               MaxLength       =   10
               TabIndex        =   1
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   405
               Left            =   4800
               MaxLength       =   100
               TabIndex        =   2
               Top             =   120
               Width           =   4935
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Id"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   2160
               TabIndex        =   95
               Top             =   195
               Width           =   195
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ex Name"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   3720
               TabIndex        =   48
               Top             =   195
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00400040&
               Height          =   285
               Left            =   120
               TabIndex        =   47
               Top             =   195
               Width           =   510
            End
         End
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0FF&
      Height          =   1095
      Left            =   15720
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame15 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   4695
         Begin VB.TextBox Text18 
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
            IMEMode         =   3  'DISABLE
            Left            =   2280
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   40
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Enter Admin Password"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   2295
         End
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
      Height          =   735
      Left            =   0
      TabIndex        =   37
      Top             =   -120
      Width           =   14595
      Begin VB.Frame Frame4 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   42
         Top             =   120
         Width           =   14565
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Exchange Master Setup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   14295
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   10815
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         Height          =   4815
         Left            =   240
         Top             =   2160
         Width           =   7815
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "ExMFrm.frx":0235
      Height          =   9060
      Left            =   240
      TabIndex        =   36
      Top             =   840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   15981
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   -2147483634
      ForeColor       =   4194304
      ListField       =   "Author"
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   9420
      Left            =   120
      Top             =   720
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   9420
      Left            =   3480
      Top             =   720
      Width           =   11085
   End
End
Attribute VB_Name = "ExMFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte:                Dim RecExTax As ADODB.Recordset:    Dim old_item As String:                 Dim LSrvTax As Double
Dim LTranTax As Double:                 Dim LRiskMFees As Double:           Dim LSEBITax As Double:                 Dim LOpt_TurnOverTax As Double
Dim LSBC_Tax As Double:                 Dim LStampDutyTax As Double:        Dim LCGSTRate As Double:                Dim LSGSTRate As Double
Dim LIGSTRate As Double:                Dim LUTTRate As Double:             Dim LCLGRate As Double:                 Dim LFilePress As Integer
Dim LSEBITaxOpt As Double:              Dim LCurrSauda As String:           Dim AccRec As ADODB.Recordset:          Dim ExchangeRec As ADODB.Recordset
Dim BankRec As ADODB.Recordset:         Dim CurrRec  As ADODB.Recordset:    Dim ExFileRec As ADODB.Recordset:       Dim RecExFile As ADODB.Recordset
Dim CashRec As ADODB.Recordset:         Dim RecExBrokClient As ADODB.Recordset
Dim LOpt_StampDuty As Double: Dim MAUTOUPDLOT As String
Dim LClientRec As ADODB.Recordset
Sub Add_Rec()
    Fb_Press = 1: old_item = vbNullString:        Call Get_Selection(1)
    Frame1.Enabled = True::             Text1.text = vbNullString
    Text2.text = vbNullString:                    DataList1.Locked = True
    SSTab1.Enabled = True
    Text1.SetFocus
    LCurrSauda = vbNullString
End Sub
Sub Save_Rec()
    Dim LExID As Integer
    Dim LEQSTampdTax As Double
    Dim LEQSTT As Double
    If LenB(Text1.text) < 1 Then MsgBox "Exchange Code required before saving record.", vbCritical, "Error": Text1.SetFocus: Exit Sub
    If LenB(TradindCombo.BoundText) < 1 Then MsgBox "Trading A/c required before saving record.", vbCritical, "Error": TradindCombo.SetFocus: Exit Sub
    If LenB(BankCombo.BoundText) < 1 Then MsgBox "Bank A/c required before saving record.", vbCritical, "Error": BankCombo.SetFocus: Exit Sub
    If LenB(CashCombo.BoundText) < 1 Then MsgBox "Cash A/c required before saving record.", vbCritical, "Error": CashCombo.SetFocus: Exit Sub
    If LenB(ContractCombo.BoundText) < 1 Then MsgBox "Contract A/c required before saving record.", vbCritical, "Error": ContractCombo.SetFocus: Exit Sub
    If LenB(BrokCombo.BoundText) < 1 Then MsgBox "Brokerage A/c required before saving record.", vbCritical, "Error": BrokCombo.SetFocus: Exit Sub
    If LenB(ShreeCombo.BoundText) < 1 Then MsgBox "Shree A/c required before saving record.", vbCritical, "Error": ShreeCombo.SetFocus: Exit Sub
    If LenB(InterestCombo.BoundText) < 1 Then MsgBox "Interest A/c required before saving record.", vbCritical, "Error": InterestCombo.SetFocus: Exit Sub
    
    If GOnlyBrok = 0 Then
        If LenB(StandingCombo.BoundText) < 1 Then MsgBox "Standing A/c required before saving record.", vbCritical, "Error": StandingCombo.SetFocus: Exit Sub
        If LenB(TransactionCombo.BoundText) < 1 Then MsgBox "Transaction A/c required before saving record.", vbCritical, "Error": TransactionCombo.SetFocus: Exit Sub
        If LenB(RskMCombo.BoundText) < 1 Then MsgBox "Risk Management Fees A/c required before saving record.", vbCritical, "Error": RskMCombo.SetFocus: Exit Sub
        If LenB(MarginCombo.BoundText) < 1 Then MsgBox "Margin A/c required before saving record.", vbCritical, "Error": MarginCombo.SetFocus: Exit Sub
        If LenB(SrvTaxCombo.BoundText) < 1 Then MsgBox "Service Tax A/c required before saving record.", vbCritical, "Error": SrvTaxCombo.SetFocus: Exit Sub
        If LenB(StamDutyCombo.BoundText) < 1 Then MsgBox "Stamp Duty Tax A/c required before saving record.", vbCritical, "Error": StamDutyCombo.SetFocus: Exit Sub
        If LenB(BankContraCombo.BoundText) < 1 Then MsgBox "Bank Contra A/c required before saving record.", vbCritical, "Error": BankContraCombo.SetFocus: Exit Sub
        If LenB(CTTCombo.BoundText) < 1 Then MsgBox "Commodity Transaction Tax A/c  required before saving record.", vbCritical, "Error": CTTCombo.SetFocus: Exit Sub
        If LenB(SEBITaxCombo.BoundText) < 1 Then MsgBox "SEBI Transaction Charges A/c  required before saving record.", vbCritical, "Error": SEBITaxCombo.SetFocus: Exit Sub
        If LenB(SBCTaxCombo.BoundText) < 1 Then MsgBox "SBC Tax A/c required before saving record.", vbCritical, "Error": SBCTaxCombo.SetFocus: Exit Sub
        
        If LenB(CGSTCombo.BoundText) < 1 Then MsgBox "CGST A/c required before saving record.", vbCritical, "Error": CGSTCombo.SetFocus: Exit Sub
        If LenB(SGSTCombo.BoundText) < 1 Then MsgBox "SGST A/c required before saving record.", vbCritical, "Error": SGSTCombo.SetFocus: Exit Sub
        If LenB(IGSTCombo.BoundText) < 1 Then MsgBox "IGST A/c required before saving record.", vbCritical, "Error": IGSTCombo.SetFocus: Exit Sub
        If LenB(UTTCombo.BoundText) < 1 Then MsgBox "UTT A/c required before saving record.", vbCritical, "Error": UTTCombo.SetFocus: Exit Sub
        If LenB(ClgCombo.BoundText) < 1 Then MsgBox "Clearing A/c required before saving record.", vbCritical, "Error": ClgCombo.SetFocus: Exit Sub
    End If
    If Fb_Press = 1 Then
        mysql = "INSERT INTO EXMAST (COMPCODE,EXCODE)"
        mysql = mysql & " VALUES(" & GCompCode & ",'" & Text1.text & "')"
        Cnn.Execute mysql
    End If
    If Text1.text = "NSE" Then
        MAUTOUPDLOT = IIf(Check4.Value = 1, "Y", "N")
        mysql = "UPDATE COMPANY SET AUTOUPDLOT = '" & MAUTOUPDLOT & "' "
        Cnn.Execute mysql
    End If
    mysql = "UPDATE EXMAST SET "
    mysql = mysql & " EXNAME ='" & Text2.text & "'"
    mysql = mysql & " ,FMCCODE  ='" & Text3.text & "'"
    mysql = mysql & " ,TRADINGACC  ='" & TradindCombo.BoundText & "'"
    mysql = mysql & " ,BANKACC  ='" & BankCombo.BoundText & "'"
    mysql = mysql & " ,ContractACC ='" & ContractCombo.BoundText & "'"
    mysql = mysql & " ,BROKAC ='" & BrokCombo.BoundText & "'"
    mysql = mysql & " ,STANDAC ='" & StandingCombo.BoundText & "'"
    mysql = mysql & " ,TRANAC ='" & TransactionCombo.BoundText & "'"
    mysql = mysql & " ,SHREEAC ='" & ShreeCombo.BoundText & "'"
    mysql = mysql & " ,RISKMACC  ='" & RskMCombo.BoundText & "'"
    mysql = mysql & " ,MarginACC  ='" & MarginCombo.BoundText & "'"
    mysql = mysql & " ,SrvTaxACC  ='" & SrvTaxCombo.BoundText & "'"
    mysql = mysql & " ,SDutyTaxACC  ='" & StamDutyCombo.BoundText & "'"
    mysql = mysql & " ,BANKCLI  ='" & BankContraCombo.BoundText & "'"
    mysql = mysql & " ,TRANTAX  ='" & CTTCombo.BoundText & "'"
    mysql = mysql & " ,STCACC  ='" & SEBITaxCombo.BoundText & "'"
    mysql = mysql & " ,SBCACC  ='" & SBCTaxCombo.BoundText & "'"
    mysql = mysql & " ,InterestAcc  ='" & InterestCombo.BoundText & "'"
    mysql = mysql & " ,CGSTACC  ='" & CGSTCombo.BoundText & "'"
    mysql = mysql & " ,SGSTACC  ='" & SGSTCombo.BoundText & "'"
    mysql = mysql & " ,IGSTACC  ='" & IGSTCombo.BoundText & "'"
    mysql = mysql & " ,UTTACC  ='" & UTTCombo.BoundText & "'"
    mysql = mysql & " ,CASHACC  ='" & CashCombo.BoundText & "'"
    mysql = mysql & " ,CLGACC  ='" & ClgCombo.BoundText & "'"
    mysql = mysql & " ,CLOSEFILE  ='" & TxtCloseFile.text & "'"
    mysql = mysql & " ,TFILETYPE  ='" & Trim(TxtTradeFile.text) & "'"
    mysql = mysql & " ,STMPDATEAPP = '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    If Check2.Value Then
        mysql = mysql & " ,LOTWISE  = 'Y'"
    Else
        mysql = mysql & " ,LOTWISE  = 'N'"
    End If
    If Check1.Value Then
        mysql = mysql & " ,TRADEDI  = 'U'"
    Else
        mysql = mysql & " ,TRADEDI  = 'P'"
    End If
    If Check3.Value Then
        mysql = mysql & ",CURR='Y'"
    Else
        mysql = mysql & ",CURR='N'"
    End If
    If ChkOptions.Value = 1 Then
        mysql = mysql & ",OPTIONS ='Y'"
    Else
        mysql = mysql & ",OPTIONS ='N'"
    End If
    If CurrDCombo.Enabled = True Then
        If LenB(CurrDCombo.BoundText) > 0 Then
            mysql = mysql & ",CURRSAUDA='" & LCurrSauda & "'"
        End If
    End If
    If Combo1.ListIndex = 0 Then
        mysql = mysql & " ,BrokerType = 'M'"
    ElseIf Combo1.ListIndex = 1 Then
        mysql = mysql & " ,BrokerType = 'S'"
    ElseIf Combo1.ListIndex = 2 Then
        mysql = mysql & " ,BrokerType = 'C'"
    ElseIf Combo1.ListIndex = 3 Then
        mysql = mysql & " ,BrokerType = 'X'"
    ElseIf Combo1.ListIndex = 4 Then
        mysql = mysql & " ,BrokerType = 'O'"
    ElseIf Combo1.ListIndex = 5 Then
        mysql = mysql & " ,BrokerType = 'T'"
    ElseIf Combo1.ListIndex = 6 Then
        mysql = mysql & " ,BrokerType = 'Z'"
    ElseIf Combo1.ListIndex = 7 Then
        mysql = mysql & " ,BrokerType = 'N'"
    ElseIf Combo1.ListIndex = 8 Then
        mysql = mysql & " ,BrokerType = 'E'"
    ElseIf Combo1.ListIndex = 9 Then
        mysql = mysql & " ,BrokerType = 'A'"
    ElseIf Combo1.ListIndex = 10 Then
        mysql = mysql & " ,BrokerType = 'W'"
    End If
    
    mysql = mysql & " ,MEMBERID ='" & UidTxt.text & "'"
    mysql = mysql & " ,PartyAs ='" & Combo2.text & "'"
    mysql = mysql & " WHERE COMPCODE =" & GCompCode & "  AND EXCODE ='" & Text1.text & "'"
    Cnn.Execute mysql
    Cnn.BeginTrans
    Cnn.Execute "DELETE FROM EXTAX WHERE COMPCODE =" & GCompCode & " AND EXCHANGECODE  ='" & Text1.text & "'"
    Cnn.CommitTrans
    Cnn.BeginTrans
        RecExTax.MoveFirst
        While Not RecExTax.EOF
            If IsNull(RecExTax!StartDate) Or IsNull(RecExTax!ENDDATE) Then
            Else
                If IsDate(RecExTax!StartDate) And IsDate(RecExTax!ENDDATE) Then
                    LTranTax = IIf(IsNull(RecExTax!CTT_TAX), 0, RecExTax!CTT_TAX)
                    LOpt_TurnOverTax = IIf(IsNull(RecExTax!OPT_CTT_TAX), 0, RecExTax!OPT_CTT_TAX)
                    LStampDutyTax = IIf(IsNull(RecExTax!STAMPDUTY), 0, RecExTax!STAMPDUTY)
                    LOpt_StampDuty = IIf(IsNull(RecExTax!OPT_STAMPDUTY), 0, RecExTax!OPT_STAMPDUTY)
                    LRiskMFees = IIf(IsNull(RecExTax!RISKMFEES), 0, RecExTax!RISKMFEES)
                    LSEBITax = IIf(IsNull(RecExTax!SEBITAX), 0, RecExTax!SEBITAX)
                    LSEBITaxOpt = IIf(IsNull(RecExTax!OPT_SEBITax), 0, RecExTax!OPT_SEBITax)
                    LCGSTRate = IIf(IsNull(RecExTax!CGSTRATE), 0, RecExTax!CGSTRATE)
                    LSGSTRate = IIf(IsNull(RecExTax!SGSTRATE), 0, RecExTax!SGSTRATE)
                    LIGSTRate = IIf(IsNull(RecExTax!IGSTRATE), 0, RecExTax!IGSTRATE)
                    LUTTRate = IIf(IsNull(RecExTax!UTTRATE), 0, RecExTax!UTTRATE)
                    LSBC_Tax = IIf(IsNull(RecExTax!SBC_TAX), 0, RecExTax!SBC_TAX)
                    LSrvTax = IIf(IsNull(RecExTax!SRV_TAX), 0, RecExTax!SRV_TAX)
                    LCLGRate = IIf(IsNull(RecExTax!CLGRATE), 0, RecExTax!CLGRATE)
                    LEQSTampdTax = IIf(IsNull(RecExTax!EQ_STAMPDUTY), 0, RecExTax!EQ_STAMPDUTY)
                    LEQSTT = IIf(IsNull(RecExTax!EQ_STT_TAX), 0, RecExTax!EQ_STT_TAX)
                    LExID = Get_ExID(Text1.text)
                    mysql = "INSERT INTO EXTAX (COMPCODE ,EXCHANGECODE,FROMDT,TODT,SERVICETAX,TURNOVERTAX,STAMPDTAX,STMPRATE,STCTAX,"
                    mysql = mysql & " SEBITAXOPT,OPT_TURNOVERTAX,SBC_TAX,CGSTRATE,SGSTRATE,IGSTRATE,UTTRATE,CLGRATE,OPT_STAMPDTAX,EXID,EQ_STT,EQ_STAMPDTAX )"
                    mysql = mysql & " VALUES (" & GCompCode & ",'" & Text1.text & "','" & Format(RecExTax!StartDate, "yyyy/MM/dd") & "','" & Format(RecExTax!ENDDATE, "yyyy/MM/dd") & "'"
                    mysql = mysql & "," & LSrvTax & "," & LTranTax & "," & LStampDutyTax & "," & LRiskMFees & "," & LSEBITax & "," & Val(LSEBITaxOpt) & ","
                    mysql = mysql & " " & LOpt_TurnOverTax & "," & LSBC_Tax & "," & LCGSTRate & "," & LSGSTRate & "," & LIGSTRate & "," & LUTTRate & "," & LCLGRate & "," & LOpt_StampDuty & "," & LExID & "," & LEQSTT & "," & LEQSTampdTax & ") "
                    Cnn.Execute mysql
                End If
            End If
            RecExTax.MoveNext
        Wend
    Cnn.CommitTrans
    CNNERR = False
    Set ExchangeRec = Nothing:    Set ExchangeRec = New ADODB.Recordset
    mysql = "SELECT * FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME"
    ExchangeRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not ExchangeRec.EOF Then
        Set DataList1.RowSource = ExchangeRec
        DataList1.ListField = "EXNAME"
        DataList1.BoundColumn = "EXCODE"
    End If
    Call CANCEL_REC
End Sub
Sub cancel_trade()
        CmdAdd.Enabled = True:                  CmdMod.Enabled = True
        TxtFileId.text = vbNullString:          txtFolderName.text = vbNullString
        TxtBroker.text = vbNullString:          TxtClient.text = vbNullString
        BrokerCombo.BoundText = vbNullString:   ClientCombo.BoundText = vbNullString
        BrokerTypeCombo.ListIndex = 0:          PartyAsCombo.ListIndex = 0
        TxtFileType.text = vbNullString:
        
End Sub
Sub CANCEL_REC()
    Text1.text = vbNullString:              Text2.text = vbNullString:              Text3.text = vbNullString
    TxtCloseFile.text = vbNullString:       UidTxt.text = vbNullString:             TxtTradeFile.text = vbNullString
    CurrDCombo.text = vbNullString:         CurrDCombo.BoundText = vbNullString:    TradindCombo.BoundText = vbNullString
    BankCombo.BoundText = vbNullString:     CashCombo.BoundText = vbNullString:     ContractCombo.BoundText = vbNullString
    BrokCombo.BoundText = vbNullString:     StandingCombo.BoundText = vbNullString: TransactionCombo.BoundText = vbNullString
    ShreeCombo.BoundText = vbNullString:    RskMCombo.BoundText = vbNullString:     MarginCombo.BoundText = vbNullString
    SrvTaxCombo.BoundText = vbNullString:   StamDutyCombo.BoundText = vbNullString: BankContraCombo.BoundText = vbNullString
    CTTCombo.BoundText = vbNullString:      SEBITaxCombo.BoundText = vbNullString:  SBCTaxCombo.BoundText = vbNullString
    InterestCombo.BoundText = vbNullString: CGSTCombo.BoundText = vbNullString:     SGSTCombo.BoundText = vbNullString
    IGSTCombo.BoundText = vbNullString:     UTTCombo.BoundText = vbNullString:    ClgCombo.BoundText = vbNullString
    Text1.Enabled = True: Text2.Enabled = True
    TxtFileType.text = vbNullString
    ChkOptions.Value = 0
    Fb_Press = 0
    Call Get_Selection(10)
    Frame1.Enabled = False
    SSTab1.Enabled = False
    If UCase(GUserName) = "ANKAN" Then
        GETMAIN.Toolbar1_Buttons(1).Enabled = True
        GETMAIN.Toolbar1_Buttons(3).Enabled = True
    Else
        GETMAIN.Toolbar1_Buttons(1).Enabled = False
        GETMAIN.Toolbar1_Buttons(3).Enabled = False
    End If
    Call SetRec
    Set ExTaxGrid.DataSource = RecExTax
    ExTaxGrid.ReBind: ExTaxGrid.Refresh
    
    Call cancel_trade

    PartyAsCombo.text = ""
    TxtAdminPass.text = "": Frame17.Visible = False: SSTab1.Tab = 0
    Call Set_EXTax_Grid
    DataList1.Locked = False
End Sub
Sub MODIFY_REC()
    Dim LBrokAcc As String
    
    Dim TRec As ADODB.Recordset
    If Trim(DataList1.BoundText) <> "" Then
        DataList1.Locked = True
        SSTab1.Enabled = True
        ExchangeRec.MoveFirst
        ExchangeRec.Find "EXCODE='" & DataList1.BoundText & "'", , adSearchForward
        If Not ExchangeRec.EOF Then
            Text1.text = ExchangeRec!excode
            TxtExID.text = ExchangeRec!EXID
            Text1.Enabled = False
            Text2.text = ExchangeRec!EXNAME
            Text3.text = IIf(IsNull(ExchangeRec!FMCCODE), vbNullString, ExchangeRec!FMCCODE)
            TxtCloseFile.text = Trim(IIf(IsNull(ExchangeRec!CLOSEFILE), vbNullString, ExchangeRec!CLOSEFILE))
            TxtTradeFile.text = Trim(IIf(IsNull(ExchangeRec!TFILETYPE), vbNullString, ExchangeRec!TFILETYPE))
            TradindCombo.BoundText = IIf(IsNull(ExchangeRec!TRADINGACC), vbNullString, ExchangeRec!TRADINGACC)
            BankCombo.BoundText = IIf(IsNull(ExchangeRec!BANKACC), vbNullString, ExchangeRec!BANKACC)
            CashCombo.BoundText = IIf(IsNull(ExchangeRec!CASHACC), vbNullString, ExchangeRec!CASHACC)
            ContractCombo.BoundText = IIf(IsNull(ExchangeRec!CONTRACTACC), vbNullString, ExchangeRec!CONTRACTACC)
            BrokCombo.BoundText = ExchangeRec!BROKAC & vbNullString
            LBrokAcc = ExchangeRec!BROKAC & vbNullString
            If IsNull(ExchangeRec!standac) Then
                StandingCombo.BoundText = LBrokAcc
            Else
                StandingCombo.BoundText = ExchangeRec!standac
            End If
            If IsNull(ExchangeRec!tranac) Then
                TransactionCombo.BoundText = LBrokAcc
            Else
                TransactionCombo.BoundText = ExchangeRec!tranac
            End If
            ShreeCombo.BoundText = ExchangeRec!shreeac & vbNullString
            If IsNull(ExchangeRec!RISKMACC) Then
                RskMCombo.BoundText = LBrokAcc
            Else
                RskMCombo.BoundText = ExchangeRec!RISKMACC
            End If
            If IsNull(ExchangeRec!MarginACC) Then
                MarginCombo.BoundText = LBrokAcc
            Else
                MarginCombo.BoundText = ExchangeRec!MarginACC
            End If
            If IsNull(ExchangeRec!SrvTaxACC) Then
                SrvTaxCombo.BoundText = LBrokAcc
            Else
                SrvTaxCombo.BoundText = ExchangeRec!SrvTaxACC
            End If
            If IsNull(ExchangeRec!SDutyTaxAcc) Then
                StamDutyCombo.BoundText = LBrokAcc
            Else
                StamDutyCombo.BoundText = ExchangeRec!SDutyTaxAcc
            End If
            If IsNull(ExchangeRec!BANKCLI) Then
                BankContraCombo.BoundText = ExchangeRec!BANKACC & vbNullString
            Else
                BankContraCombo.BoundText = ExchangeRec!BANKCLI
            End If
            If IsNull(ExchangeRec!TRANTAX) Then
                CTTCombo.BoundText = LBrokAcc
            Else
                CTTCombo.BoundText = ExchangeRec!TRANTAX
            End If
            If IsNull(ExchangeRec!STCACC) Then
                SEBITaxCombo.BoundText = LBrokAcc
            Else
                SEBITaxCombo.BoundText = ExchangeRec!STCACC
            End If
            If IsNull(ExchangeRec!SBCACC) Then
                SBCTaxCombo.BoundText = LBrokAcc
            Else
                SBCTaxCombo.BoundText = ExchangeRec!SBCACC
            End If
            If IsNull(ExchangeRec!InterestAcc) Then
                InterestCombo.BoundText = LBrokAcc
            Else
                InterestCombo.BoundText = ExchangeRec!InterestAcc
            End If
            If IsNull(ExchangeRec!CGSTACC) Then
                CGSTCombo.BoundText = LBrokAcc
            Else
                CGSTCombo.BoundText = ExchangeRec!CGSTACC
            End If
            If IsNull(ExchangeRec!SGSTACC) Then
                SGSTCombo.BoundText = LBrokAcc
            Else
                SGSTCombo.BoundText = ExchangeRec!SGSTACC
            End If
            If IsNull(ExchangeRec!IGSTACC) Then
                IGSTCombo.BoundText = LBrokAcc
            Else
                IGSTCombo.BoundText = ExchangeRec!IGSTACC
            End If
            If IsNull(ExchangeRec!UTTACC) Then
                UTTCombo.BoundText = LBrokAcc
            Else
                UTTCombo.BoundText = ExchangeRec!UTTACC
            End If
            If IsNull(ExchangeRec!CLGACC) Then
                ClgCombo.BoundText = LBrokAcc
            Else
                ClgCombo.BoundText = ExchangeRec!CLGACC
            End If
            vcDTP1.Value = IIf(IsNull(ExchangeRec!STMPDATEAPP), Date, ExchangeRec!STMPDATEAPP)
            If UCase(ExchangeRec!BROKERTYPE) = "M" Then
                Combo1.ListIndex = 0
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "S" Then
                Combo1.ListIndex = 1
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "C" Then
                Combo1.ListIndex = 2
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "X" Then
                Combo1.ListIndex = 3
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "O" Then
                Combo1.ListIndex = 4
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "T" Then
                Combo1.ListIndex = 5
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "Z" Then
                Combo1.ListIndex = 6
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "N" Then
                Combo1.ListIndex = 7
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "E" Then
                Combo1.ListIndex = 8
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "A" Then
                Combo1.ListIndex = 9
            ElseIf UCase(ExchangeRec!BROKERTYPE) = "W" Then
                Combo1.ListIndex = 10
            End If
            If IsNull(ExchangeRec!TRADEDI) Then
               Check1.Value = 1
            Else
                If ExchangeRec!TRADEDI = "P" Then
                    Check1.Value = 0
                Else
                    Check1.Value = 1
                End If
            End If
            If IsNull(ExchangeRec!Options) Then
                ChkOptions.Value = 0
            Else
                If ExchangeRec!Options = "Y" Then
                    ChkOptions.Value = 1
                Else
                    ChkOptions.Value = 0
                End If
            End If
            If IsNull(ExchangeRec!LOTWISE) Then
               Check2.Value = 0
            Else
                If ExchangeRec!LOTWISE = "Y" Then
                    Check2.Value = 1
                Else
                    Check2.Value = 0
                End If
            End If
            
            If Text1.text = "NSE" Then
                Check4.Visible = True
                Check4.Value = IIf(MAUTOUPDLOT = "Y", 1, 0)
            Else
                Check4.Visible = False
            End If
            
            UidTxt.text = ExchangeRec!MemberId & vbNullString
            If ExchangeRec!PARTYAS = "Client" Then
                Combo2.ListIndex = 0
            Else
                Combo2.ListIndex = 1
            End If
            If ExchangeRec!CURR = "Y" Then
                Check3.Value = 1
                If Not IsNull(ExchangeRec!CURRSAUDA) Then
                    CurrDCombo.BoundText = ExchangeRec!CURRSAUDA
                End If
            Else
                Check3.Value = 0
                CurrDCombo.BoundText = vbNullString
            End If
            Frame1.Enabled = True
            'Exchange Wise Tax
            Call Refresh_FileGrid
            Call Refresh_ExBrokGrid
            
            Call SetRec
            
            mysql = "SELECT * FROM EXTAX WHERE COMPCODE =" & GCompCode & " AND EXCHANGECODE ='" & Text1.text & "'"
            Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            If Not TRec.EOF Then
                While Not TRec.EOF
                    RecExTax.AddNew
                    RecExTax!StartDate = TRec!FROMDT:                           RecExTax!ENDDATE = TRec!ToDt
                    RecExTax!SRV_TAX = Val(TRec!servicetax & vbNullString):     RecExTax!CTT_TAX = Val(TRec!TURNOVERTAX & vbNullString)
                    RecExTax!STAMPDUTY = Val(TRec!stampdtax & vbNullString):    RecExTax!OPT_STAMPDUTY = Val(TRec!OPT_stampdtax & vbNullString):
                    RecExTax!RISKMFEES = Val(TRec!STMPRATE & vbNullString):     RecExTax!SEBITAX = Val(TRec!STCTAX & vbNullString):
                    RecExTax!SBC_TAX = Val(TRec!SBC_TAX & vbNullString):        RecExTax!CGSTRATE = Val(TRec!CGSTRATE & vbNullString)
                    RecExTax!SGSTRATE = Val(TRec!SGSTRATE & vbNullString):      RecExTax!IGSTRATE = Val(TRec!IGSTRATE & vbNullString)
                    RecExTax!UTTRATE = Val(TRec!UTTRATE & vbNullString):        RecExTax!OPT_SEBITax = Val(TRec!SEBITaxOPT & vbNullString)
                    RecExTax!CLGRATE = Val(TRec!CLGRATE & vbNullString):        RecExTax!OPT_CTT_TAX = Val(TRec!OPT_TURNOVERTAX & vbNullString)
                    RecExTax!EQ_STT_TAX = Val(TRec!EQ_STT & vbNullString):      RecExTax!EQ_STAMPDUTY = Val(TRec!EQ_STAMPDTAX & vbNullString)
                    RecExTax.Update
                    TRec.MoveNext
                Wend
            Else
                RecExTax.AddNew
                RecExTax!StartDate = GFinBegin:     RecExTax!ENDDATE = GFinEnd
                RecExTax!SRV_TAX = 0:               RecExTax!CTT_TAX = 0
                RecExTax!STAMPDUTY = 0:             RecExTax!OPT_STAMPDUTY = 0
                RecExTax!RISKMFEES = 0:             RecExTax!SEBITAX = 0
                RecExTax!SBC_TAX = 0:               RecExTax!CGSTRATE = 0
                RecExTax!SGSTRATE = 0:              RecExTax!IGSTRATE = 0
                RecExTax!UTTRATE = 0:               RecExTax!OPT_SEBITax = 0
                RecExTax!CLGRATE = 0:               RecExTax!OPT_CTT_TAX = 0
                RecExTax!EQ_STT_TAX = 0
                RecExTax!EQ_STAMPDUTY = 0
                
            End If
            Set ExTaxGrid.DataSource = RecExTax
            Call Set_EXTax_Grid
            Text2.SetFocus
        End If
    Else
        MsgBox "Please Select Exchange.", vbInformation
        Call CANCEL_REC
        DataList1.Locked = False
        DataList1.SetFocus
    End If
End Sub
Private Sub ExNameDb_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub BankCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub CASHCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub


Private Sub BrokerCombo_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub BrokerCombo_Validate(Cancel As Boolean)
If LenB(BrokerCombo.BoundText) > 0 Then
    AccRec.MoveFirst
    AccRec.Find "AC_CODE='" & BrokerCombo.BoundText & "'"
    If Not AccRec.EOF Then
        TxtBroker.text = AccRec!AC_CODE
    Else
        MsgBox "Invalid Broker"
        Cancel = True
    End If
Else
    MsgBox "Invalid Broker"
    Cancel = True
End If
End Sub

Private Sub BrokerCombo2_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub BrokerCombo2_Validate(Cancel As Boolean)
Dim LBrokerCode As String
Dim TRec As ADODB.Recordset
Dim I As Integer
If LenB(BrokerCombo2.BoundText) > 0 Then
    LBrokerCode = Get_AccountMCode(BrokerCombo2.BoundText)
    If LenB(LBrokerCode) > 0 Then
        TxtBrokCode.text = LBrokerCode
        mysql = "SELECT A.CLIENT,B.NAME FROM EXBROKCLIENT AS A, ACCOUNTD AS B WHERE A.COMPCODE =" & GCompCode & " AND A.BROKER ='" & LBrokerCode & "'"
        mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.CLIENT  =B.AC_CODE ORDER BY B.NAME "
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        ListView2.ListItems.Clear
        Do While Not TRec.EOF
            ListView2.ListItems.Add , , TRec!CLIENT
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , TRec!NAME
            ListView2.ListItems(ListView2.ListItems.Count).Checked = True
            For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).ListSubItems(1).text = TRec!CLIENT Then
                    ListView1.ListItems(I).Checked = True
                End If
            Next
            TRec.MoveNext
        Loop
    Else
        MsgBox "Invalid Broker"
        'Cancel = True
    End If
Else
    MsgBox "Invalid Broker"
   ' Cancel = True
End If
End Sub

Private Sub BrokerTypeCombo_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub BrokerTypeCombo_Validate(Cancel As Boolean)
If BrokerTypeCombo.ListIndex < 0 Then
    MsgBox "Please select Broker Type"
    Cancel = True
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check1.Caption = "Pro as UserID"
Else
    Check1.Caption = "Pro as Pro"
End If
End Sub

Private Sub ClientCombo_GotFocus()
Sendkeys "%{DOWN}"
End Sub
Private Sub ClientCombo2_GotFocus()
Sendkeys "%{DOWN}"
End Sub


'Private Sub ClientCombo2_Validate(Cancel As Boolean)
'If LenB(ClientCombo2.BoundText) > 0 Then
'    AccRec.MoveFirst
'    AccRec.Find "AC_CODE='" & ClientCombo2.BoundText & "'"
'    If Not AccRec.EOF Then
'        TxtClCode.text = AccRec!AC_CODE
'    Else
'        MsgBox "Invalid Client"
'        Cancel = True
'    End If
'Else
'    MsgBox "Invalid Client"
'    Cancel = True
'End If
'
'End Sub

Private Sub ClientCombo_Validate(Cancel As Boolean)
If LenB(ClientCombo.BoundText) > 0 Then
    AccRec.MoveFirst
    AccRec.Find "AC_CODE='" & ClientCombo.BoundText & "'"
    If Not AccRec.EOF Then
        TxtClient.text = AccRec!AC_CODE
    Else
        MsgBox "Invalid Client"
        Cancel = True
    End If
Else
    MsgBox "Invalid Client"
    Cancel = True
End If

End Sub

Private Sub CmdAdd_Click()
Dim TRec  As ADODB.Recordset
Dim LMaxId As Integer
LFilePress = 1
Set TRec = Nothing
Set TRec = New ADODB.Recordset
mysql = "SELECT MAX(FILEID ) AS MID FROM EXFILE WHERE COMPCODE =" & GCompCode & ""
TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
If TRec.EOF Then
    LMaxId = 1
Else
    If IsNull(TRec!Mid) Then
        LMaxId = 1
    Else
        LMaxId = TRec!Mid + 1
    End If
End If
TxtFileId.text = LMaxId:                txtFolderName.text = vbNullString
TxtBroker.text = vbNullString:          TxtClient.text = vbNullString
BrokerCombo.BoundText = vbNullString:   ClientCombo.BoundText = vbNullString
BrokerTypeCombo.ListIndex = 0:          PartyAsCombo.ListIndex = 0
TxtFileType.text = vbNullString
CmdAdd.Enabled = False:                 CmdMod.Enabled = False
CmdSave.Enabled = True:                 TxtFileId.Locked = True
txtFolderName.SetFocus

End Sub
Private Sub CmdMod_Click()
    LFilePress = 2
    TxtFileId.Locked = False:               TxtFileId.text = vbNullString
    txtFolderName.text = vbNullString:      TxtBroker.text = vbNullString
    TxtClient.text = vbNullString:          BrokerCombo.BoundText = vbNullString
    ClientCombo.BoundText = vbNullString:    TxtFileType.text = vbNullString
    BrokerTypeCombo.ListIndex = 0:          PartyAsCombo.ListIndex = 0
    CmdAdd.Enabled = False:                 CmdMod.Enabled = False
    CmdSave.Enabled = True:                 TxtFileId.SetFocus
    TxtFileId.SetFocus
End Sub

Private Sub CmdSave_Click()
Dim LBrokerType As String:      Dim LPartyAs As String: Dim LExFileType  As String
Dim lclientid As Integer: Dim Lbrokerid As Integer
If LFilePress = 2 Then
    mysql = "DELETE FROM EXFILE WHERE COMPCODE =" & GCompCode & " AND EXCODE='" & Text1.text & "'AND FILEID =" & Trim(TxtFileId.text) & ""
    Cnn.Execute mysql
End If
If LenB(txtFolderName.text) > 1 Then
    mysql = "UPDATE EXMAST SET SAUDATRADE = 'Y' WHERE COMPCODE =" & GCompCode & " AND EXCODE='" & Text1.text & "'"
    Cnn.Execute mysql
End If
If LenB(txtFolderName.text) > 1 Then
    
    If BrokerTypeCombo.ListIndex = 0 Then
        LBrokerType = "M"
    ElseIf BrokerTypeCombo.ListIndex = 1 Then
        LBrokerType = "S"
    ElseIf BrokerTypeCombo.ListIndex = 2 Then
        LBrokerType = "C"
    ElseIf BrokerTypeCombo.ListIndex = 3 Then
        LBrokerType = "X"
    ElseIf BrokerTypeCombo.ListIndex = 4 Then
        LBrokerType = "O"
    ElseIf BrokerTypeCombo.ListIndex = 5 Then
        LBrokerType = "T"
    ElseIf BrokerTypeCombo.ListIndex = 6 Then
        LBrokerType = "E"
    ElseIf BrokerTypeCombo.ListIndex = 7 Then
        LBrokerType = "W"
    ElseIf BrokerTypeCombo.ListIndex = 8 Then
        LBrokerType = "1"
    End If
    LExFileType = Trim(TxtFileType.text)
    LPartyAs = Left$(PartyAsCombo.text, 1)
    Lbrokerid = Get_AccID(TxtBroker.text)
    lclientid = Get_AccID(TxtClient.text)

    
    mysql = "INSERT INTO EXFILE(COMPCODE,EXCODE,FILEID,FOLDERNAME,BROKERTYPE,PARTYAS,CONTRACTACC,CLIENTCODE,EXFILETYPE,EXID,CLIENTID,BROKERID)"
    mysql = mysql & " VALUES (" & GCompCode & ",'" & Text1.text & "'," & Trim(TxtFileId.text) & ",'" & txtFolderName & "','" & LBrokerType & "','" & LPartyAs & "','" & TxtBroker.text & "','" & TxtClient.text & "','" & LExFileType & "'," & Val(TxtExID.text) & "," & lclientid & "," & Lbrokerid & ")"
    Cnn.Execute mysql
End If
Call cancel_trade
Call Refresh_FileGrid
CmdAdd.SetFocus
End Sub

Private Sub Combo1_GotFocus()
    Sendkeys "%{down}"
End Sub


Private Sub Combo2_GotFocus()
    Sendkeys "%{down}"
End Sub

Private Sub Command1_Click()
    Call cancel_trade
End Sub

Private Sub Command2_Click()
If Trim(TxtFileId.text) <> "" Then
    If MsgBox("Confirm DELETE?", vbYesNo) = vbYes Then
        mysql = "DELETE FROM EXFILE WHERE COMPCODE =" & GCompCode & " AND EXCODE='" & Text1.text & "'AND FILEID =" & Trim(TxtFileId.text) & ""
        Cnn.Execute mysql
                       
        Call cancel_trade
        Call Refresh_FileGrid
        CmdAdd.SetFocus
    End If
    
End If
End Sub

Private Sub ContractCombo_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub BrokCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub CurrDCombo_GotFocus()
Sendkeys "%{DOWN}"
End Sub
Private Sub CurrDCombo_Validate(Cancel As Boolean)
If LenB(CurrDCombo.BoundText) > 0 Then
    LCurrSauda = CurrDCombo.BoundText
    CurrRec.MoveFirst
    CurrRec.Find "SAUDACODE='" & LCurrSauda & "'", , adSearchForward
    If CurrRec.EOF Then
        CurrDCombo.BoundText = vbNullString
        LCurrSauda = vbNullString
        MsgBox "Please Select Valid Contract"
        Cancel = True
    End If
End If
End Sub
Private Sub DataCombo6_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
'Private Sub DataCombo6_Validate(Cancel As Boolean)
'If LenB(DataCombo6.BoundText) < 1 Then
'    MsgBox " Please Select Broker"
'    Cancel = True
'Else
'    AccRec.MoveFirst
'    AccRec.Find " AC_CODE='" & DataCombo6.BoundText & "'"
'    If AccRec.EOF Then
'        MsgBox " Invalid broker  "
'        Cancel = True
'    Else
'        Text10.text = AccRec!AC_CODE
'    End If
'End If
'End Sub

'Private Sub DataCombo7_Validate(Cancel As Boolean)
'If LenB(DataCombo7.BoundText) < 1 Then
'Else
'    AccRec.MoveFirst
'    AccRec.Find " AC_CODE='" & DataCombo7.BoundText & "'"
'    If AccRec.EOF Then
'        MsgBox " Invalid Client  "
'        Cancel = True
'    Else
'        Text11.text = AccRec!AC_CODE
'    End If
'End If
'End Sub

Private Sub DataCombo7_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub PartyAsCombo_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub PartyAsCombo_Validate(Cancel As Boolean)
If PartyAsCombo.ListIndex < 0 Then
    MsgBox "Please select Party As"
    Cancel = True
Else
    If PartyAsCombo.ListIndex = 2 Then
        TxtClient.Enabled = True
        ClientCombo.Enabled = True
    Else
        TxtClient.Enabled = False
        ClientCombo.Enabled = False
    End If
End If
End Sub

Private Sub SrvTaxCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 4 Then
        TxtAdminPass.text = ""
        Frame17.Visible = True
        SSTab1.Enabled = False
        TxtAdminPass.SetFocus
    End If
End Sub

Private Sub StamDutyCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub BankContraCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub CTTCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub CGSTCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub SGSTCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub IGSTCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub TxtAdminPass_Validate(Cancel As Boolean)
    If TxtAdminPass.text = VAdminPass Then
        SSTab1.Enabled = True
        TxtFileId.SetFocus
    ElseIf TxtAdminPass.text <> "" Then
        MsgBox "Invalid password!!!"
        TxtAdminPass.text = ""
    End If
    TxtAdminPass.text = vbNullString
End Sub
Private Sub TxtAdminPass_LostFocus()
    Frame17.Visible = False
End Sub
'Private Sub Text10_Validate(Cancel As Boolean)
'If LenB(Text10.text) > 0 Then
'    AccRec.MoveFirst
'    AccRec.Find "AC_CODE ='" & Text10.text & "'"
'    If Not AccRec.EOF Then
'        DataCombo6.BoundText = Text10.text
'    End If
'End If
'End Sub
'Private Sub Text11_Validate(Cancel As Boolean)
'If LenB(Text11.text) > 0 Then
'    AccRec.MoveFirst
'    AccRec.Find "AC_CODE ='" & Text11.text & "'"
'    If Not AccRec.EOF Then
'        DataCombo7.BoundText = Text11.text
'    End If
'End If
'End Sub
Private Sub TxtBroker_Validate(Cancel As Boolean)
If LenB(TxtBroker.text) > 1 Then
    BrokerCombo.BoundText = TxtBroker.text
End If
End Sub
Private Sub TxtClient_Validate(Cancel As Boolean)
If LenB(TxtClient.text) > 1 Then
    ClientCombo.BoundText = TxtClient.text
End If
End Sub
Private Sub TxtFileId_Validate(Cancel As Boolean)
Dim TRec As ADODB.Recordset

If Trim(TxtFileId.text) <> "" Then

    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT FILEID,FOLDERNAME,CONTRACTACC,BROKERTYPE,PARTYAS,CLIENTCODE,EXFILETYPE FROM EXFILE WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & Text1.text & "' AND FILEID =" & Trim(TxtFileId.text) & ""
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not TRec.EOF Then
        txtFolderName.text = TRec!FOLDERNAME
        TxtBroker.text = TRec!CONTRACTACC
        BrokerCombo.BoundText = TRec!CONTRACTACC
        TxtClient.text = TRec!CLIENTCODE
        TxtFileType.text = (TRec!EXFILETYPE & vbNullString)
        If Not IsNull(TRec!CLIENTCODE) Then
            TxtClient.text = TRec!CLIENTCODE
            ClientCombo.BoundText = TRec!CLIENTCODE
        Else
            TxtClient.text = vbNullString
            ClientCombo.BoundText = vbNullString
        End If
        If TRec!BROKERTYPE = "M" Then
            BrokerTypeCombo.ListIndex = 0
        ElseIf TRec!BROKERTYPE = "S" Then
            BrokerTypeCombo.ListIndex = 1
        ElseIf TRec!BROKERTYPE = "C" Then
            BrokerTypeCombo.ListIndex = 2
        ElseIf TRec!BROKERTYPE = "X" Then
            BrokerTypeCombo.ListIndex = 3
        ElseIf TRec!BROKERTYPE = "O" Then
            BrokerTypeCombo.ListIndex = 4
        ElseIf TRec!BROKERTYPE = "T" Then
            BrokerTypeCombo.ListIndex = 5
        ElseIf TRec!BROKERTYPE = "E" Then
            BrokerTypeCombo.ListIndex = 6
        ElseIf TRec!BROKERTYPE = "W" Then
            BrokerTypeCombo.ListIndex = 7
        ElseIf TRec!BROKERTYPE = "1" Then
            BrokerTypeCombo.ListIndex = 8
        End If
        If TRec!PARTYAS = "C" Then
            PartyAsCombo.ListIndex = 0
        ElseIf TRec!PARTYAS = "U" Then
            PartyAsCombo.ListIndex = 1
        ElseIf TRec!PARTYAS = "F" Then
            PartyAsCombo.ListIndex = 2
        End If
            
    Else
        MsgBox "No Entry for this File Id "
        TxtFileId.SetFocus
    End If
End If
End Sub
Private Sub txtFolderName_Validate(Cancel As Boolean)
If LenB(txtFolderName.text) > 0 Then
    txtFolderName.text = UCase(txtFolderName.text)
Else
    If LFilePress = 1 Then
        MsgBox "Folder Name can not be blank "
    End If
End If
End Sub
Private Sub UTTCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub CLGCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub InterestCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub StandingCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub TransactionCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub ShreeCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub RskMCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub DataList1_Click()
    
    Text1.text = DataList1.BoundText
    Text2.text = DataList1.text
    SSTab1.Tab = 0
End Sub
Private Sub DataList1_DblClick()
    If DataList1.Locked Then
    Else
        Call Get_Selection(2)
        Fb_Press = 2
        Call MODIFY_REC
    End If
End Sub
Private Sub DataList1_KeyPress(KeyAscii As Integer)
    If DataList1.Locked Then
    Else
        If KeyAscii = 13 Then
            Call Get_Selection(2)
            Fb_Press = 2
            Call MODIFY_REC
        End If
    End If
End Sub
Private Sub ExTaxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = 13 And ExTaxGrid.Col = 1 Then
        RecExTax.MoveNext
        If RecExTax.EOF Then
            RecExTax.AddNew
            RecExTax!StartDate = Format(GFinBegin, "YYYY/MM/DD")
            RecExTax!ENDDATE = Format(GFinEnd, "YYYY/MM/DD")
            RecExTax!CTT_TAX = 0
            RecExTax!OPT_CTT_TAX = 0
            RecExTax!STAMPDUTY = 0:
            RecExTax!OPT_STAMPDUTY = 0:
            RecExTax!RISKMFEES = 0
            RecExTax!SEBITAX = 0:
            RecExTax!OPT_SEBITax = 0
            RecExTax!CGSTRATE = 0
            RecExTax!SGSTRATE = 0:
            RecExTax!IGSTRATE = 0
            RecExTax!UTTRATE = 0:
            RecExTax!SBC_TAX = 0:
            RecExTax!SRV_TAX = 0:
            RecExTax!CLGRATE = 0
            RecExTax!CLGRATE = 0
            RecExTax!EQ_STT_TAX = 0
            RecExTax!EQ_STAMPDUTY = 0
            
            
            RecExTax.Update
        End If
        ExTaxGrid.Col = 0
    ElseIf KeyCode = 13 Then
          Sendkeys "{tab}"
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    If KeyCode = 27 And Frame14.Visible Then Frame14.Visible = False
'    If KeyCode = 121 Then
'        Frame14.Visible = True
'        Text18.SetFocus
'    End If

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Text18_Validate(Cancel As Boolean)
    If Text18.text = "minku2010" Then
        Combo1.Locked = False
    End If
    Text18.text = ""
End Sub
Private Sub text18_LostFocus()
    Frame14.Visible = False
End Sub
Private Sub Form_Load()
    LFilePress = 0
    CmdSave.Enabled = False
    Check4.Visible = False
    
    
    Call CANCEL_REC
    
    Set ExchangeRec = Nothing:    Set ExchangeRec = New ADODB.Recordset
    mysql = "SELECT AUTOUPDLOT FROM COMPANY WHERE COMPCODE =" & GCompCode & ""
    ExchangeRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not ExchangeRec.EOF Then MAUTOUPDLOT = ExchangeRec!AUTOUPDLOT
    
    Set ExchangeRec = Nothing:    Set ExchangeRec = New ADODB.Recordset
    mysql = "SELECT * FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME"
    ExchangeRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    If Not ExchangeRec.EOF Then
        Set DataList1.RowSource = ExchangeRec
        DataList1.ListField = "EXNAME"
        DataList1.BoundColumn = "EXCODE"
    End If
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    AccRec.Open "SELECT NAME,AC_CODE FROM ACCOUNTM WHERE COMPCODE = " & GCompCode & " ORDER BY NAME", Cnn, adOpenKeyset, adLockReadOnly
    If Not AccRec.EOF Then
        Set TradindCombo.RowSource = AccRec:        TradindCombo.ListField = "NAME":    TradindCombo.BoundColumn = "AC_CODE"
        Set ContractCombo.RowSource = AccRec:       ContractCombo.ListField = "NAME":   ContractCombo.BoundColumn = "AC_CODE"
        Set BrokCombo.RowSource = AccRec:           BrokCombo.ListField = "NAME":       BrokCombo.BoundColumn = "AC_CODE"
        Set InterestCombo.RowSource = AccRec:       InterestCombo.ListField = "NAME":   InterestCombo.BoundColumn = "AC_CODE"
        Set ShreeCombo.RowSource = AccRec:          ShreeCombo.ListField = "NAME":      ShreeCombo.BoundColumn = "AC_CODE"
        Set BrokerCombo.RowSource = AccRec:         BrokerCombo.ListField = "NAME":     BrokerCombo.BoundColumn = "AC_CODE"
        Set ClientCombo.RowSource = AccRec:         ClientCombo.ListField = "NAME":     ClientCombo.BoundColumn = "AC_CODE"
        Set BrokerCombo2.RowSource = AccRec:        BrokerCombo2.ListField = "NAME":    BrokerCombo2.BoundColumn = "AC_CODE"
        'Set ClientCombo2.RowSource = AccRec:        ClientCombo2.ListField = "NAME":    ClientCombo2.BoundColumn = "AC_CODE"
        If GOnlyBrok = 0 Then
            Set StandingCombo.RowSource = AccRec:       StandingCombo.ListField = "NAME":       StandingCombo.BoundColumn = "AC_CODE"
            Set TransactionCombo.RowSource = AccRec:    TransactionCombo.ListField = "NAME":    TransactionCombo.BoundColumn = "AC_CODE"
            Set RskMCombo.RowSource = AccRec:           RskMCombo.ListField = "NAME":           RskMCombo.BoundColumn = "AC_CODE"
            Set MarginCombo.RowSource = AccRec:         MarginCombo.ListField = "NAME":         MarginCombo.BoundColumn = "AC_CODE"
            Set SrvTaxCombo.RowSource = AccRec:         SrvTaxCombo.ListField = "NAME":         SrvTaxCombo.BoundColumn = "AC_CODE"
            Set StamDutyCombo.RowSource = AccRec:       StamDutyCombo.ListField = "NAME":       StamDutyCombo.BoundColumn = "AC_CODE"
            Set BankContraCombo.RowSource = AccRec:     BankContraCombo.ListField = "NAME":     BankContraCombo.BoundColumn = "AC_CODE"
            Set CTTCombo.RowSource = AccRec:            CTTCombo.ListField = "NAME":            CTTCombo.BoundColumn = "AC_CODE"
            Set SEBITaxCombo.RowSource = AccRec:        SEBITaxCombo.ListField = "NAME":        SEBITaxCombo.BoundColumn = "AC_CODE"
            Set SBCTaxCombo.RowSource = AccRec:         SBCTaxCombo.ListField = "NAME":         SBCTaxCombo.BoundColumn = "AC_CODE"
            Set CGSTCombo.RowSource = AccRec:           CGSTCombo.ListField = "NAME":           CGSTCombo.BoundColumn = "AC_CODE"
            Set SGSTCombo.RowSource = AccRec:           SGSTCombo.ListField = "NAME":           SGSTCombo.BoundColumn = "AC_CODE"
            Set IGSTCombo.RowSource = AccRec:           IGSTCombo.ListField = "NAME":           IGSTCombo.BoundColumn = "AC_CODE"
            Set UTTCombo.RowSource = AccRec:            UTTCombo.ListField = "NAME":            UTTCombo.BoundColumn = "AC_CODE"
            Set ClgCombo.RowSource = AccRec:            ClgCombo.ListField = "NAME":            ClgCombo.BoundColumn = "AC_CODE"
        Else
            StandingCombo.Visible = False:      TransactionCombo.Visible = False:   RskMCombo.Visible = False
            MarginCombo.Visible = False:        SrvTaxCombo.Visible = False:        StamDutyCombo.Visible = False
            BankContraCombo.Visible = False:    CTTCombo.Visible = False:           SEBITaxCombo.Visible = False
            SBCTaxCombo.Visible = False:        CGSTCombo.Visible = False:          SGSTCombo.Visible = False
            IGSTCombo.Visible = False:          UTTCombo.Visible = False:           ClgCombo.Visible = False
            Label3.Visible = False:             Label5.Visible = False:             Label6.Visible = False
            Label7.Visible = False:             Label10.Visible = False:            Label12.Visible = False
            Label14.Visible = False:            Label20.Visible = False:            Label21.Visible = False
            Label25.Visible = False:            Label28.Visible = False:            Label29.Visible = False
            Label30.Visible = False:            Label32.Visible = False:            Label35.Visible = False
        End If
        
        
    End If
    Set CurrRec = Nothing
    Set CurrRec = New ADODB.Recordset
    mysql = "SELECT SAUDACODE, SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY ='" & Format(GFinEnd, "YYYY/MM/DD") & "'"
    CurrRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not CurrRec.EOF Then
        Set CurrDCombo.RowSource = CurrRec
        CurrDCombo.ListField = "SAUDANAME"
        CurrDCombo.BoundColumn = "SAUDACODE"
    Else
        CurrDCombo.Enabled = False
    End If
    
    Set BankRec = Nothing: Set BankRec = New ADODB.Recordset
    BankRec.Open "SELECT NAME,AC_CODE FROM ACCOUNTM WHERE COMPCODE = " & GCompCode & " AND  GCODE = 11", Cnn, adOpenKeyset, adLockReadOnly
    If Not BankRec.EOF Then Set BankCombo.RowSource = BankRec: BankCombo.ListField = "NAME": BankCombo.BoundColumn = "AC_CODE"
    
    Set CashRec = Nothing: Set CashRec = New ADODB.Recordset
    CashRec.Open "SELECT NAME,AC_CODE FROM ACCOUNTM WHERE COMPCODE = " & GCompCode & " AND  GCODE = 10", Cnn, adOpenKeyset, adLockReadOnly
    If Not CashRec.EOF Then Set CashCombo.RowSource = CashRec: CashCombo.ListField = "NAME": CashCombo.BoundColumn = "AC_CODE"
    Call Refresh_ExBrokGrid
    SSTab1.Tab = 0
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
End Sub
Private Sub stmrate_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERChk(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim TRec As ADODB.Recordset
    If Fb_Press = 1 Then
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT EXCODE FROM EXMAST WHERE COMPCODE =" & GCompCode & " AND EXCODE='" & Text1.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then MsgBox "Exchange code already exists.", vbExclamation, "Warning": Cancel = True
    End If
End Sub
Sub LIST_ITEM()
'Dim TREC As ADODB.Recordset
'    Screen.MousePointer = 11
'
'    Call Get_Selection(12)
'
'    MYSQL = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME"
'    Set TREC = Nothing
'    Set TREC = New ADODB.Recordset
'    TREC.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
'
'    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "RptExList.RPT", 1)
'
'    RDCREPO.DiscardSavedData
'    RDCREPO.Database.SetDataSource TREC
'
'    CRViewer1.Width = CInt(GETMAIN.Width - 100)
'    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
'    CRViewer1.Top = 0
'    CRViewer1.Left = 0
'
'    CRViewer1.Visible = True
'    CRViewer1.ReportSource = RDCREPO
'
'    CRViewer1.ViewReport
'
'    Set RPT = Nothing
'    Screen.MousePointer = 0
End Sub
Private Sub TradindCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Sub SetRec()
    Set RecExTax = Nothing: Set RecExTax = New ADODB.Recordset
    RecExTax.Fields.Append "StartDate", adDate, , adFldIsNullable
    RecExTax.Fields.Append "EndDate", adDate, , adFldIsNullable
    RecExTax.Fields.Append "CTT_Tax", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "Opt_CTT_Tax", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "EQ_STT_Tax", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "StampDuty", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "OPT_StampDuty", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "EQ_StampDuty", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "RiskMFees", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "SEBITAX", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "OPT_SEBITAX", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "CGSTRate", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "SGSTRate", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "IGSTRate", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "UTTRate", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "SBC_Tax", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "Srv_Tax", adDouble, , adFldIsNullable
    RecExTax.Fields.Append "ClgRate", adDouble, , adFldIsNullable
    RecExTax.Open , , adOpenKeyset, adLockBatchOptimistic
End Sub

Private Sub Refresh_FileGrid()
    Dim TRec As ADODB.Recordset
    Set RecExFile = Nothing
    Set RecExFile = New ADODB.Recordset
    RecExFile.Fields.Append "FileId", adInteger, , adFldIsNullable
    RecExFile.Fields.Append "FolderName", adVarChar, 30, adFldIsNullable
    RecExFile.Fields.Append "ContAcc", adVarChar, 15, adFldIsNullable
    RecExFile.Fields.Append "ContName", adVarChar, 100, adFldIsNullable
    RecExFile.Fields.Append "BrokerType", adVarChar, 30, adFldIsNullable
    RecExFile.Fields.Append "PartyAs", adVarChar, 30, adFldIsNullable
    RecExFile.Fields.Append "ExFileType", adVarChar, 3, adFldIsNullable
    RecExFile.Fields.Append "ClientCode", adVarChar, 15, adFldIsNullable
    RecExFile.Fields.Append "ClientName", adVarChar, 100, adFldIsNullable
    RecExFile.Open , , adOpenKeyset, adLockBatchOptimistic
    mysql = "SELECT FILEID,FOLDERNAME,PARTYAS ,BROKERTYPE,CONTRACTACC,CLIENTCODE,B.NAME,EXFILETYPE FROM EXFILE AS A,ACCOUNTD AS B  "
    mysql = mysql & " WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE AND A.CONTRACTACC=B.AC_CODE "
    mysql = mysql & " AND A.EXCODE ='" & Text1.text & "' ORDER BY A.FILEID"
    Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            RecExFile.AddNew
            RecExFile!FILEID = TRec!FILEID
            RecExFile!FOLDERNAME = TRec!FOLDERNAME
            RecExFile!CONTACC = TRec!CONTRACTACC
            RecExFile!CONTNAME = TRec!NAME
            RecExFile!BROKERTYPE = TRec!BROKERTYPE
            RecExFile!PARTYAS = TRec!PARTYAS
            RecExFile!CLIENTCODE = TRec!CLIENTCODE
            RecExFile!EXFILETYPE = TRec!EXFILETYPE
            RecExFile.Update
            TRec.MoveNext
        Loop
    Else
        RecExFile.AddNew
        RecExFile!FILEID = 1
        RecExFile!FOLDERNAME = vbNullString
        RecExFile!CONTACC = vbNullString
        RecExFile!CONTNAME = vbNullString
        RecExFile!BROKERTYPE = vbNullString
        RecExFile!PARTYAS = vbNullString
        RecExFile!CLIENTCODE = vbNullString
        RecExFile!EXFILETYPE = vbNullString
        RecExFile.Update
    End If
    Set DataGrid1.DataSource = RecExFile
    DataGrid1.ReBind: DataGrid1.Refresh
    DataGrid1.Columns(0).Width = 900:
    DataGrid1.Columns(4).Width = 900:
    DataGrid1.Columns(5).Width = 900:
End Sub

Private Sub Refresh_ExBrokGrid()
    Dim TRec As ADODB.Recordset
    Set RecExBrokClient = Nothing
    Set RecExBrokClient = New ADODB.Recordset
    'RecExBrokClient.Fields.Append "SNo", adInteger, , adFldIsNullable
    RecExBrokClient.Fields.Append "BrokCode", adVarChar, 15, adFldIsNullable
    RecExBrokClient.Fields.Append "BrokName", adVarChar, 100, adFldIsNullable
    RecExBrokClient.Fields.Append "ClCode", adVarChar, 15, adFldIsNullable
    RecExBrokClient.Fields.Append "ClName", adVarChar, 100, adFldIsNullable
    RecExBrokClient.Open , , adOpenKeyset, adLockBatchOptimistic
    mysql = "SELECT A.BROKER,C.NAME AS BROKNAME,A.CLIENT,B.NAME AS CLNAME FROM EXBROKCLIENT AS A ,ACCOUNTD AS B, ACCOUNTD AS C"
    mysql = mysql & " WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE AND A.COMPCODE=C.COMPCODE AND A.CLIENT=B.AC_CODE AND A.BROKER = C.AC_CODE "
    mysql = mysql & " ORDER BY C.NAME,B.NAME "
    Set TRec = Nothing: Set TRec = New ADODB.Recordset:
    TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            RecExBrokClient.AddNew
            'RecExBrokClient!SNO = TRec!SNO
            RecExBrokClient!CLCODE = TRec!CLIENT
            RecExBrokClient!CLNAME = TRec!CLNAME
            RecExBrokClient!BROKCODE = TRec!BROKER
            RecExBrokClient!BROKNAME = TRec!BROKNAME
            RecExBrokClient.Update
            
            TRec.MoveNext
        Loop
    Else
        RecExBrokClient.AddNew
        RecExBrokClient!CLCODE = vbNullString
        RecExBrokClient!CLNAME = vbNullString
        RecExBrokClient!BROKCODE = vbNullString
        RecExBrokClient!BROKNAME = vbNullString
        RecExBrokClient.Update
    End If
    Set ExBrokGrid.DataSource = RecExBrokClient
    ExBrokGrid.ReBind: ExBrokGrid.Refresh
    
End Sub
Private Sub CmdAdd2_Click()
Dim TRec  As ADODB.Recordset
Dim LMaxId As Integer
LFilePress = 1
Set LClientRec = Nothing
Set LClientRec = New ADODB.Recordset
mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & ""
mysql = mysql & " AND AC_CODE NOT IN (SELECT DISTINCT CLIENT FROM EXBROKCLIENT  WHERE COMPCODE =" & GCompCode & ")"
mysql = mysql & " ORDER BY NAME "
LClientRec.Open mysql, Cnn, adOpenForwardOnly
ListView1.ListItems.Clear
ListView2.ListItems.Clear
ListView1.Visible = False
Do While Not LClientRec.EOF
    ListView1.ListItems.Add , , LClientRec!AC_CODE
    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LClientRec!NAME
    LClientRec.MoveNext
Loop
ListView1.Visible = True
ListView2.Visible = True

TxtBrokCode.text = vbNullString:
BrokerCombo2.BoundText = vbNullString:
CmdAdd2.Enabled = False:                 CmdMod2.Enabled = False
CmdSave2.Enabled = True:
'TxtSno.Locked = True
'TxtClCode.SetFocus
End Sub
Private Sub CmdMod2_Click()
    LFilePress = 2
    'TxtSno.Locked = False
    
    'TxtSno.text = vbNullString
    TxtBrokCode.text = vbNullString
    'TxtClCode.text = vbNullString
    BrokerCombo2.BoundText = vbNullString
    'ClientCombo2.BoundText = vbNullString
    CmdAdd2.Enabled = False
    CmdMod2.Enabled = False
    CmdSave2.Enabled = True
    Set LClientRec = Nothing
    Set LClientRec = New ADODB.Recordset
    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & ""
    mysql = mysql & " AND AC_CODE NOT IN (SELECT DISTINCT CLIENT FROM EXBROKCLIENT WHERE COMPCODE =" & GCompCode & ")"
    mysql = mysql & " ORDER BY NAME "
    LClientRec.Open mysql, Cnn, adOpenForwardOnly
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView1.Visible = False
    Do While Not LClientRec.EOF
        ListView1.ListItems.Add , , LClientRec!AC_CODE
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LClientRec!NAME
        LClientRec.MoveNext
    Loop
    ListView1.Visible = True
    ListView2.Visible = True
    TxtBrokCode.SetFocus
End Sub
Private Sub CmdSave2_Click()
Dim LBrokerCode     As String
Dim I As Integer
Dim LSNO As Long
Dim LBrokID As Long
Dim lclientid As Long
If LenB(BrokerCombo2.BoundText) > 0 Then
    LBrokerCode = Get_AccountMCode(BrokerCombo2.BoundText)
    If LenB(LBrokerCode) > 0 Then
        Cnn.BeginTrans
        mysql = "DELETE FROM EXBROKCLIENT WHERE COMPCODE =" & GCompCode & " AND BROKER='" & LBrokerCode & " '"
        Cnn.Execute mysql
        LBrokID = Get_AccID(LBrokerCode)
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Checked = True Then
                lclientid = Get_AccID(ListView1.ListItems(I).text)
                mysql = "EXEC PINSERT_EXBROKCLIENT " & GCompCode & ",'" & LBrokerCode & "','" & ListView1.ListItems(I).text & "'," & LBrokID & "," & lclientid & ""
                Cnn.Execute mysql
            End If
        Next
        For I = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(I).Checked = True Then
                
                lclientid = Get_AccID(ListView2.ListItems(I).text)
                mysql = "EXEC PINSERT_EXBROKCLIENT " & GCompCode & ",'" & LBrokerCode & "','" & ListView2.ListItems(I).text & "'," & LBrokID & "," & lclientid & ""
                Cnn.Execute mysql
            End If
        Next
        Cnn.CommitTrans
    End If
    
Else
    MsgBox "Please Select Valid Broker"
End If
CmdAdd2.Enabled = True:                 CmdMod2.Enabled = True
'TxtSno.text = vbNullString:
TxtBrokCode.text = vbNullString
'TxtClCode.text = vbNullString:
BrokerCombo2.BoundText = vbNullString
'ClientCombo2.BoundText = vbNullString:
ListView1.ListItems.Clear
ListView2.ListItems.Clear
Call Refresh_ExBrokGrid
CmdAdd2.SetFocus
End Sub
'Private Sub TxtSno_Validate(Cancel As Boolean)
'Dim TRec As ADODB.Recordset
'If LFilePress = 2 Then
'    Set TRec = Nothing
'    Set TRec = New ADODB.Recordset
'    MYSQL = "SELECT SNO,CLIENT ,BROKER FROM EXBROKCLIENT  WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & Text1.text & "' AND SNO =" & Val(TxtSno.text) & ""
'    TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'    If Not TRec.EOF Then
'        TxtBrokCode.text = TRec!BROKER
'        BrokerCombo2.BoundText = TRec!BROKER
'        TxtClCode.text = TRec!CLIENT
'        ClientCombo2.BoundText = TRec!CLIENT
'    Else
'        MsgBox "No Entry for this Sno  "
'        TxtSno.SetFocus
'    End If
'End If
'End Sub
Private Sub TxtBrokcode_Validate(Cancel As Boolean)
If LenB(TxtBrokCode.text) > 1 Then
    BrokerCombo2.BoundText = TxtBrokCode.text
End If
End Sub
'Private Sub TxtClcode_Validate(Cancel As Boolean)
'If LenB(TxtClCode.text) > 1 Then
'    ClientCombo2.BoundText = TxtClCode.text
'End If
'End Sub
Private Sub Set_EXTax_Grid()
    ExTaxGrid.ReBind: ExTaxGrid.Refresh
    ExTaxGrid.Columns(0).Width = 1300:                  ExTaxGrid.Columns(1).Width = 1300
    ExTaxGrid.Columns(2).Width = 1300:                  ExTaxGrid.Columns(3).Width = 1300
    ExTaxGrid.Columns(4).Width = 1300:                  ExTaxGrid.Columns(5).Width = 1300
    ExTaxGrid.Columns(6).Width = 1300:                  ExTaxGrid.Columns(7).Width = 1300
    ExTaxGrid.Columns(8).Width = 1300:                  ExTaxGrid.Columns(9).Width = 1300
    ExTaxGrid.Columns(10).Width = 1300:                 ExTaxGrid.Columns(11).Width = 1300
    ExTaxGrid.Columns(12).Width = 1300:                 ExTaxGrid.Columns(13).Width = 1300:
    ExTaxGrid.Columns(14).Width = 1300:                 ExTaxGrid.Columns(15).Width = 1300:
    ExTaxGrid.Columns(16).Width = 1300:                 ExTaxGrid.Columns(17).Width = 1300:
    ExTaxGrid.Columns(2).Alignment = dbgRight:          ExTaxGrid.Columns(3).Alignment = dbgRight
    ExTaxGrid.Columns(4).Alignment = dbgRight:          ExTaxGrid.Columns(5).Alignment = dbgRight
    ExTaxGrid.Columns(6).Alignment = dbgRight:          ExTaxGrid.Columns(7).Alignment = dbgRight
    ExTaxGrid.Columns(8).Alignment = dbgRight:          ExTaxGrid.Columns(9).Alignment = dbgRight
    ExTaxGrid.Columns(10).Alignment = dbgRight:         ExTaxGrid.Columns(11).Alignment = dbgRight
    ExTaxGrid.Columns(12).Alignment = dbgRight:         ExTaxGrid.Columns(13).Alignment = dbgRight
    ExTaxGrid.Columns(14).Alignment = dbgRight:         ExTaxGrid.Columns(15).Alignment = dbgRight
    ExTaxGrid.Columns(16).Alignment = dbgRight:         ExTaxGrid.Columns(17).Alignment = dbgRight
    
    ExTaxGrid.Columns(2).NumberFormat = "0.00000":      ExTaxGrid.Columns(3).NumberFormat = "0.00000"
    ExTaxGrid.Columns(4).NumberFormat = "0.00000":      ExTaxGrid.Columns(5).NumberFormat = "0.00000"
    ExTaxGrid.Columns(6).NumberFormat = "0.00000":      ExTaxGrid.Columns(7).NumberFormat = "0.00000"
    ExTaxGrid.Columns(8).NumberFormat = "0.00000":      ExTaxGrid.Columns(9).NumberFormat = "0.00000"
    ExTaxGrid.Columns(10).NumberFormat = "0.00000":     ExTaxGrid.Columns(11).NumberFormat = "0.00000"
    ExTaxGrid.Columns(12).NumberFormat = "0.00000":     ExTaxGrid.Columns(13).NumberFormat = "0.00000"
    ExTaxGrid.Columns(14).NumberFormat = "0.00000"
    ExTaxGrid.Columns(15).NumberFormat = "0.00000"
    ExTaxGrid.Columns(16).NumberFormat = "0.00000"
    ExTaxGrid.Columns(17).NumberFormat = "0.00000"
End Sub
