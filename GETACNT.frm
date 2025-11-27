VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form GETACNT 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GETACNT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MouseIcon       =   "GETACNT.frx":0442
   ScaleHeight     =   10350
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame25 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "ERGGHE"
      Height          =   4095
      Left            =   2160
      TabIndex        =   193
      Top             =   3840
      Visible         =   0   'False
      Width           =   9855
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
         ItemData        =   "GETACNT.frx":0884
         Left            =   120
         List            =   "GETACNT.frx":0886
         TabIndex        =   195
         Top             =   480
         Width           =   9615
      End
      Begin VB.CommandButton Command5 
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
         Height          =   480
         Left            =   9120
         TabIndex        =   194
         ToolTipText     =   "Close"
         Top             =   -15
         Width           =   615
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Check"
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
         Left            =   80
         TabIndex        =   196
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   9000
      Width           =   10935
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   13800
      TabIndex        =   142
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Assign NewCode as Exchange Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   154
         Top             =   4680
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.TextBox TxtnewName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   100
         TabIndex        =   148
         Top             =   4200
         Width           =   3735
      End
      Begin VB.TextBox TxtoldName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   146
         Top             =   2040
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   150
         Top             =   5640
         Width           =   3735
      End
      Begin VB.TextBox TxtNewCode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   15
         TabIndex        =   147
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox TxtOldCode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   15
         TabIndex        =   145
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "New Account Name"
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
         TabIndex        =   152
         Top             =   3720
         Width           =   3735
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Old Account Name"
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
         TabIndex        =   151
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "New Account Code"
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
         TabIndex        =   149
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Old Account Code"
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
         TabIndex        =   144
         Top             =   615
         Width           =   3735
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Change Account Code"
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
         Left            =   120
         TabIndex        =   143
         Top             =   120
         Width           =   3735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   2895
      Left            =   480
      TabIndex        =   137
      Top             =   6000
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
         Name            =   "Verdana"
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
   Begin VB.Frame Frame24 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   480
      TabIndex        =   93
      Top             =   5160
      Width           =   12855
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Change Account Code"
         Height          =   405
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox TxtFilterCode 
         Height          =   405
         Left            =   1560
         TabIndex        =   94
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TxtFilterName 
         Height          =   405
         Left            =   3480
         TabIndex        =   95
         Top             =   120
         Width           =   4935
      End
      Begin VB.CommandButton CmdFilterOK 
         BackColor       =   &H00FF8080&
         Caption         =   "Go"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   136
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   135
         Top             =   195
         Width           =   495
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   195
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   13695
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   13695
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Account Setup"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   13215
         End
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   1335
      Left            =   16920
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   870
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
   Begin TabDlg.SSTab ACCSETUP 
      Height          =   4260
      Left            =   240
      TabIndex        =   30
      Top             =   840
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   7514
      _Version        =   393216
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Account Details"
      TabPicture(0)   =   "GETACNT.frx":0888
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame11"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Party Details"
      TabPicture(1)   =   "GETACNT.frx":08A4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame18"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Brokerage"
      TabPicture(2)   =   "GETACNT.frx":08C0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame19"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Exchange Details"
      TabPicture(3)   =   "GETACNT.frx":08DC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Brokerage and Taxes"
      TabPicture(4)   =   "GETACNT.frx":08F8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame17"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Party Multiplier"
      TabPicture(5)   =   "GETACNT.frx":0914
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame2"
      Tab(5).Control(1)=   "Frame8"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Party Brokerage"
      TabPicture(6)   =   "GETACNT.frx":0930
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame14"
      Tab(6).Control(1)=   "Frame15"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Sharing"
      TabPicture(7)   =   "GETACNT.frx":094C
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame20"
      Tab(7).Control(1)=   "Frame23"
      Tab(7).ControlCount=   2
      Begin VB.Frame Frame23 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame23"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74760
         TabIndex        =   184
         Top             =   2040
         Width           =   12735
         Begin MSDataGridLib.DataGrid ShareGrid 
            Height          =   1815
            Left            =   120
            TabIndex        =   185
            Top             =   120
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   3201
            _Version        =   393216
            BackColor       =   16761024
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
                  LCID            =   16393
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
                  LCID            =   16393
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
      Begin VB.Frame Frame20 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   168
         Top             =   720
         Width           =   12735
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   840
            TabIndex        =   183
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox Combo6 
            Height          =   360
            Left            =   2280
            TabIndex        =   172
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FF8080&
            Caption         =   "Update"
            Height          =   405
            Left            =   11280
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   11280
            TabIndex        =   170
            Top             =   120
            Width           =   1335
         End
         Begin VB.ComboBox Combo5 
            Height          =   360
            Left            =   8400
            TabIndex        =   169
            Text            =   "Combo3"
            Top             =   120
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   360
            Left            =   840
            TabIndex        =   173
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
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   360
            Left            =   4680
            TabIndex        =   174
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
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
         Begin MSDataListLib.DataCombo DataCombo5 
            Height          =   360
            Left            =   1800
            TabIndex        =   182
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
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
         Begin MSDataListLib.DataCombo DataCombo6 
            Height          =   360
            Left            =   4680
            TabIndex        =   186
            Top             =   600
            Width           =   2535
            _ExtentX        =   4471
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
         Begin MSDataListLib.DataCombo DataCombo7 
            Height          =   360
            Left            =   8400
            TabIndex        =   187
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
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
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   181
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "Share From"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7200
            TabIndex        =   180
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "Share  To"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3720
            TabIndex        =   179
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "ExCode"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7200
            TabIndex        =   177
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "ShareRate"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10200
            TabIndex        =   176
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3720
            TabIndex        =   175
            Top             =   195
            Width           =   495
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -74760
         TabIndex        =   166
         Top             =   1560
         Width           =   12735
         Begin MSDataGridLib.DataGrid BrokGrid 
            Height          =   2295
            Left            =   120
            TabIndex        =   167
            Top             =   120
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   4048
            _Version        =   393216
            BackColor       =   16761024
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
                  LCID            =   16393
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
                  LCID            =   16393
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
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74760
         TabIndex        =   155
         Top             =   720
         Width           =   12735
         Begin VB.ComboBox Combo4 
            Height          =   360
            Left            =   7200
            TabIndex        =   163
            Top             =   120
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   10440
            TabIndex        =   162
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FF8080&
            Caption         =   "Update"
            Height          =   405
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   120
            Width           =   1095
         End
         Begin VB.ComboBox Combo3 
            Height          =   360
            Left            =   2280
            TabIndex        =   158
            Top             =   120
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   360
            Left            =   960
            TabIndex        =   157
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   360
            Left            =   4200
            TabIndex        =   165
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
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
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
            Height          =   255
            Left            =   3720
            TabIndex        =   164
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "BrokRate"
            Height          =   255
            Left            =   9480
            TabIndex        =   160
            Top             =   195
            Width           =   975
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "BrokType"
            Height          =   255
            Left            =   6240
            TabIndex        =   159
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "ExCode"
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   210
            Width           =   735
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -74760
         TabIndex        =   132
         Top             =   1560
         Width           =   12615
         Begin MSDataGridLib.DataGrid Multigrid 
            Height          =   2295
            Left            =   120
            TabIndex        =   133
            Top             =   120
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   4048
            _Version        =   393216
            BackColor       =   16761024
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74760
         TabIndex        =   120
         Top             =   720
         Width           =   12615
         Begin VB.CommandButton CmdMod 
            BackColor       =   &H00FF8080&
            Caption         =   "Mod"
            Height          =   475
            Left            =   10560
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton CmdSave 
            BackColor       =   &H00FF8080&
            Caption         =   "Save"
            Height          =   475
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton CmdAdd 
            BackColor       =   &H00FF8080&
            Caption         =   "Add"
            Height          =   475
            Left            =   9480
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox TxtMultiplier 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   8040
            TabIndex        =   125
            Top             =   120
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo ExCombo 
            Height          =   360
            Left            =   2520
            TabIndex        =   123
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.TextBox TxtSno 
            Height          =   405
            Left            =   720
            TabIndex        =   122
            Top             =   120
            Width           =   855
         End
         Begin MSDataListLib.DataCombo ItemCombo 
            Height          =   360
            Left            =   4680
            TabIndex        =   124
            Top             =   120
            Width           =   2295
            _ExtentX        =   4048
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
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier"
            Height          =   255
            Left            =   7080
            TabIndex        =   131
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
            Height          =   255
            Left            =   4080
            TabIndex        =   130
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ex Code"
            Height          =   255
            Left            =   1680
            TabIndex        =   129
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "S.No"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74760
         TabIndex        =   76
         Top             =   720
         Width           =   12735
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   1575
            Left            =   120
            TabIndex        =   91
            Top             =   120
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   2778
            _Version        =   393216
            BackColor       =   16761024
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   11
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "INSTTYPE"
               Caption         =   "InstType"
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
               DataField       =   "BROKTYPE"
               Caption         =   "BrokType"
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
               DataField       =   "BrokRate"
               Caption         =   "BrokRate"
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
               DataField       =   "BrokRate2"
               Caption         =   "Brokrate2"
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
               DataField       =   "trantype"
               Caption         =   "TranType"
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
               DataField       =   "TranRate"
               Caption         =   "TranRate"
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
               DataField       =   "STDRATE"
               Caption         =   "StdRate"
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
               DataField       =   "MARTYPE"
               Caption         =   "MarType"
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
               DataField       =   "marrate"
               Caption         =   "MarRate"
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
            BeginProperty Column10 
               DataField       =   "UPTOSTDT"
               Caption         =   "UptoDate"
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
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column10 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   92
            Top             =   1800
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   2566
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BackColor       =   16761024
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            TabAction       =   1
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   12
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "ITEMNAME"
               Caption         =   "Item"
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
               DataField       =   "BROKTYPE"
               Caption         =   "Brok. Type"
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
               DataField       =   "BROKRATE"
               Caption         =   "Brok. Rate"
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
               DataField       =   "STDRATE"
               Caption         =   "Std Rate"
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
               DataField       =   "PARTYTYPE"
               Caption         =   "Party Type"
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
               DataField       =   "BCYCLE"
               Caption         =   "Billing Cycle"
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
               DataField       =   "TRANRATE"
               Caption         =   "Tran. Fees"
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
               DataField       =   "TranType"
               Caption         =   "TranType"
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
               DataField       =   "MarginType"
               Caption         =   "MarginType"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0.0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "MarginRate"
               Caption         =   "MarginRate"
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
            BeginProperty Column11 
               DataField       =   "UptoStdt"
               Caption         =   "UptoStdt"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   2
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   2025.071
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1544.882
               EndProperty
               BeginProperty Column10 
                  Alignment       =   1
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   1214.929
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   73
         Top             =   720
         Width           =   12855
         Begin VB.Frame Frame13 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   12735
            Begin MSDataListLib.DataCombo BrokerCombo 
               Height          =   360
               Left            =   8880
               TabIndex        =   29
               Top             =   2760
               Width           =   3735
               _ExtentX        =   6588
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
            Begin VB.TextBox TxtCovertRate 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   4920
               TabIndex        =   28
               Top             =   2760
               Width           =   2295
            End
            Begin VB.ComboBox Combo2 
               Height          =   360
               ItemData        =   "GETACNT.frx":0968
               Left            =   1560
               List            =   "GETACNT.frx":0978
               TabIndex        =   27
               Top             =   2760
               Width           =   2175
            End
            Begin VB.TextBox TxtPanno 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   4920
               MaxLength       =   15
               TabIndex        =   19
               ToolTipText     =   "Enter Residence No."
               Top             =   1200
               Width           =   2295
            End
            Begin VB.TextBox TxtMobile 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1560
               MaxLength       =   25
               TabIndex        =   21
               ToolTipText     =   "Enter Mobile No."
               Top             =   1740
               Width           =   2175
            End
            Begin VB.TextBox TxtUCC 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   8880
               MaxLength       =   50
               TabIndex        =   26
               ToolTipText     =   "Enter Mail Id."
               Top             =   2280
               Width           =   3735
            End
            Begin VB.TextBox TxtGST 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1560
               MaxLength       =   50
               TabIndex        =   24
               ToolTipText     =   "Enter Mail Id."
               Top             =   2280
               Width           =   2175
            End
            Begin VB.TextBox TxtCINNo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   4920
               MaxLength       =   50
               TabIndex        =   25
               ToolTipText     =   "Enter Mail Id."
               Top             =   2280
               Width           =   2295
            End
            Begin VB.TextBox TxtEmail 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   4920
               MaxLength       =   50
               TabIndex        =   22
               ToolTipText     =   "Enter Mail Id."
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox TxtStateCode 
               Height          =   405
               Left            =   4920
               MaxLength       =   10
               TabIndex        =   16
               Top             =   660
               Width           =   2295
            End
            Begin VB.TextBox TxtState 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   15
               ToolTipText     =   "Enter City Name"
               Top             =   660
               Width           =   2175
            End
            Begin VB.TextBox TxtFax 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   8880
               MaxLength       =   15
               TabIndex        =   23
               ToolTipText     =   "Enter Fax No."
               Top             =   1680
               Width           =   3735
            End
            Begin VB.TextBox TxtDirector 
               Height          =   405
               Left            =   8880
               TabIndex        =   20
               Top             =   1200
               Width           =   3735
            End
            Begin VB.TextBox TxtPhoneO 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   8880
               MaxLength       =   15
               TabIndex        =   17
               ToolTipText     =   "Enter Office Phone No."
               Top             =   660
               Width           =   3735
            End
            Begin VB.TextBox TxtPhoneR 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   18
               ToolTipText     =   "Enter Residence No."
               Top             =   1200
               Width           =   2175
            End
            Begin VB.TextBox TxtAdd 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1560
               MaxLength       =   100
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               ToolTipText     =   "Enter Party Address"
               Top             =   120
               Width           =   5655
            End
            Begin VB.TextBox TxtCity 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   8880
               MaxLength       =   15
               TabIndex        =   13
               ToolTipText     =   "Enter City Name"
               Top             =   120
               Width           =   2565
            End
            Begin VB.TextBox txtpin 
               Height          =   360
               Left            =   11520
               MaxLength       =   10
               TabIndex        =   14
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label60 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Broker"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7440
               TabIndex        =   188
               Top             =   2850
               Width           =   975
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Curr Diff"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   3840
               TabIndex        =   141
               Top             =   2760
               Width           =   810
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Convert Type"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   140
               Top             =   2760
               Width           =   1350
            End
            Begin VB.Label Label10 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "PAN No"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   3840
               TabIndex        =   116
               Top             =   1275
               Width           =   1095
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   173
               Width           =   855
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone (R)"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   114
               Top             =   1260
               Width           =   975
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile No."
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   113
               Top             =   1800
               Width           =   1020
            End
            Begin VB.Label Label36 
               BackStyle       =   0  'Transparent
               Caption         =   "State"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   735
               Width           =   495
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GSTIN"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   111
               Top             =   2370
               Width           =   615
            End
            Begin VB.Label Label35 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "UCC"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7440
               TabIndex        =   108
               Top             =   2355
               Width           =   495
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CIN No."
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   3840
               TabIndex        =   107
               Top             =   2340
               Width           =   750
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   3840
               TabIndex        =   106
               Top             =   1800
               Width           =   600
            End
            Begin VB.Label Label10 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "State Code"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   3840
               TabIndex        =   90
               Top             =   735
               Width           =   1095
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "TAccess Hash"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7440
               TabIndex        =   81
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label15 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "T User Id"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7440
               TabIndex        =   80
               Top             =   1275
               Width           =   1695
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone (O)"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   7440
               TabIndex        =   79
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "City && PIN"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7440
               TabIndex        =   75
               Top             =   180
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74880
         TabIndex        =   51
         Top             =   720
         Width           =   12735
         Begin VB.Frame Frame21 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   120
            TabIndex        =   77
            Top             =   840
            Width           =   12735
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Create Self Branch"
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   120
               TabIndex        =   191
               Top             =   2040
               Width           =   2175
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   10680
               TabIndex        =   190
               Top             =   1200
               Width           =   1695
            End
            Begin VB.TextBox TxtSebiRate 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   7200
               TabIndex        =   138
               Top             =   1200
               Width           =   975
            End
            Begin VB.CheckBox ChkMultiplier 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Apply Multiplier"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   7200
               TabIndex        =   134
               Top             =   720
               Width           =   1875
            End
            Begin VB.CheckBox ChkFutBrok 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Brok On Future Expiry"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   5160
               TabIndex        =   119
               Top             =   1680
               Width           =   2475
            End
            Begin VB.CheckBox ChkCutBrok 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Brok On Options Expiry"
               ForeColor       =   &H00000000&
               Height          =   495
               Left            =   2400
               TabIndex        =   118
               Top             =   1680
               Width           =   2595
            End
            Begin VB.CheckBox ChkPersonnelAc 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Personnel Account"
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   120
               TabIndex        =   117
               Top             =   1680
               Width           =   2175
            End
            Begin VB.ComboBox InterestCombo 
               Height          =   360
               ItemData        =   "GETACNT.frx":09AB
               Left            =   9000
               List            =   "GETACNT.frx":09C1
               TabIndex        =   63
               Top             =   720
               Width           =   1455
            End
            Begin VB.CheckBox ChkUTT 
               BackColor       =   &H00FFC0C0&
               Caption         =   "UTT Applicable"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   120
               TabIndex        =   65
               Top             =   1200
               Width           =   2055
            End
            Begin VB.CheckBox ChkIGST 
               BackColor       =   &H00FFC0C0&
               Caption         =   "IGST Applicable"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   120
               TabIndex        =   60
               Top             =   720
               Width           =   2055
            End
            Begin VB.CheckBox ChkSGST 
               BackColor       =   &H00FFC0C0&
               Caption         =   "SGST Applicable"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   4920
               TabIndex        =   67
               Top             =   1200
               Width           =   2115
            End
            Begin VB.CheckBox ChkCGST 
               BackColor       =   &H00FFC0C0&
               Caption         =   "CGST Applicable"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   2280
               TabIndex        =   66
               Top             =   1200
               Width           =   2475
            End
            Begin VB.CheckBox ChkServiceTax 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Service Tax Apply"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   2280
               TabIndex        =   61
               Top             =   720
               Width           =   2475
            End
            Begin VB.CheckBox CHKInterest 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Interest Applicable"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   4920
               TabIndex        =   62
               Top             =   720
               Width           =   2235
            End
            Begin VB.TextBox TxtIntRate 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   11400
               TabIndex        =   64
               Top             =   720
               Width           =   975
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   8280
               TabIndex        =   86
               Top             =   1680
               Visible         =   0   'False
               Width           =   4215
               Begin VB.TextBox TxtStampRate 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   2880
                  TabIndex        =   70
                  Top             =   120
                  Width           =   1095
               End
               Begin vcDateTimePicker.vcDTP vcDTP2 
                  Height          =   450
                  Left            =   720
                  TabIndex        =   68
                  Top             =   120
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   794
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   41937.4853009259
               End
               Begin VB.Label Label42 
                  BackColor       =   &H00C0E0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "From Date"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   88
                  Top             =   165
                  Width           =   855
               End
               Begin VB.Label Label43 
                  BackColor       =   &H00C0E0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rate"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   87
                  Top             =   180
                  Width           =   495
               End
            End
            Begin VB.ComboBox StampDutyCombo 
               Height          =   360
               ItemData        =   "GETACNT.frx":09F8
               Left            =   10560
               List            =   "GETACNT.frx":0A0E
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   120
               Width           =   1815
            End
            Begin VB.ComboBox SEBITaxCombo 
               Height          =   360
               ItemData        =   "GETACNT.frx":0A6A
               Left            =   3960
               List            =   "GETACNT.frx":0A7A
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   120
               Width           =   1935
            End
            Begin VB.ComboBox RiskMCombo 
               Height          =   360
               ItemData        =   "GETACNT.frx":0AAE
               Left            =   7200
               List            =   "GETACNT.frx":0ABB
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   120
               Width           =   1935
            End
            Begin VB.ComboBox CTTCmb 
               Height          =   360
               ItemData        =   "GETACNT.frx":0AE3
               Left            =   720
               List            =   "GETACNT.frx":0AF0
               Style           =   2  'Dropdown List
               TabIndex        =   56
               Top             =   120
               Width           =   1935
            End
            Begin VB.Label Label61 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Credit Limit"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   9360
               TabIndex        =   189
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label Label45 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Int Rate"
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
               Left            =   10560
               TabIndex        =   89
               Top             =   795
               Width           =   735
            End
            Begin VB.Label Label24 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "StampDuty"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   9360
               TabIndex        =   85
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label44 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "SEBI Tax"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2760
               TabIndex        =   84
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label41 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Risk Mngt."
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6000
               TabIndex        =   83
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label16 
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "CTT "
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   180
               Width           =   495
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   12615
            Begin VB.TextBox TxtBrokRate 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   5520
               TabIndex        =   54
               Top             =   120
               Width           =   2415
            End
            Begin VB.ComboBox ComboBrokType 
               Height          =   360
               ItemData        =   "GETACNT.frx":0B18
               Left            =   1680
               List            =   "GETACNT.frx":0B2E
               TabIndex        =   53
               Top             =   120
               Width           =   2415
            End
            Begin VB.ComboBox ComboBCycle 
               Height          =   360
               ItemData        =   "GETACNT.frx":0B9C
               Left            =   9600
               List            =   "GETACNT.frx":0BA6
               TabIndex        =   55
               Top             =   120
               Width           =   2775
            End
            Begin VB.Label Label17 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Brokerage Type"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   195
               Width           =   1935
            End
            Begin VB.Label Label20 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Brokerage Rate"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4440
               TabIndex        =   71
               Top             =   195
               Width           =   1455
            End
            Begin VB.Label Label22 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Billing Cycle"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   8280
               TabIndex        =   69
               Top             =   180
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   -74760
         TabIndex        =   49
         Top             =   780
         Width           =   12735
         Begin VB.ComboBox billbycmb 
            ForeColor       =   &H00000000&
            Height          =   360
            ItemData        =   "GETACNT.frx":0BC2
            Left            =   5280
            List            =   "GETACNT.frx":0BCF
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   1200
            Visible         =   0   'False
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid ExGrid 
            Height          =   3135
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            BackColor       =   16761024
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            TabAction       =   1
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Exchange Client Code"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "ExCode"
               Caption         =   "Exchange"
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
               DataField       =   "ACEXCODE"
               Caption         =   "Party Exchange Code"
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
               DataField       =   "BILLINGTYPE"
               Caption         =   "Bill By"
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
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3525.166
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2505.26
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         TabIndex        =   35
         Top             =   720
         Width           =   12735
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   1560
            Width           =   12495
            Begin MSDataListLib.DataCombo DComboPartyHead 
               Height          =   360
               Left            =   9120
               TabIndex        =   11
               Top             =   120
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   635
               _Version        =   393216
               Text            =   "DataCombo3"
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
            Begin VB.Frame Frame1 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1560
               TabIndex        =   47
               Top             =   0
               Width           =   4575
               Begin VB.OptionButton OptPType_Pro 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Pro"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2400
                  TabIndex        =   9
                  Top             =   180
                  Width           =   735
               End
               Begin VB.OptionButton OptPType_Board 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Board"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   10
                  Top             =   180
                  Width           =   975
               End
               Begin VB.OptionButton OptPType_Broker 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Broker"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   120
                  TabIndex        =   7
                  Top             =   180
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton OptPType_Client 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Cleint"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1240
                  TabIndex        =   8
                  Top             =   180
                  Width           =   975
               End
            End
            Begin VB.Label Label39 
               BackStyle       =   0  'Transparent
               Caption         =   "Party Head"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7680
               TabIndex        =   78
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Party Type"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   855
            Left            =   120
            TabIndex        =   39
            Top             =   2520
            Width           =   12495
            Begin VB.TextBox TODr 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   405
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   120
               Width           =   1935
            End
            Begin VB.TextBox TOCr 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   405
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   120
               Width           =   1815
            End
            Begin VB.TextBox DIFF 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               ForeColor       =   &H00FF0000&
               Height          =   405
               Left            =   10680
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   120
               Width           =   1695
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               Caption         =   "Opening Trial Balance Total"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   120
               TabIndex        =   109
               Top             =   120
               Width           =   3015
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Debit"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   3360
               TabIndex        =   45
               Top             =   180
               Width           =   510
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Credit"
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   6120
               TabIndex        =   44
               Top             =   165
               Width           =   615
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Opening Difference "
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   8760
               TabIndex        =   43
               Top             =   180
               Width           =   1920
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   12495
            Begin VB.TextBox TxtAccID 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FF0000&
               Height          =   405
               Left            =   1560
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   139
               TabStop         =   0   'False
               ToolTipText     =   "Enter Account Name (50)"
               Top             =   120
               Width           =   735
            End
            Begin VB.CheckBox ChkActive 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Active"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   4920
               TabIndex        =   6
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox TxtClBal 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   9120
               Locked          =   -1  'True
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   720
               Width           =   1815
            End
            Begin VB.Frame Frame22 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   11040
               TabIndex        =   101
               Top             =   600
               Width           =   1365
               Begin VB.OptionButton OptDr_Clbal 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Dr"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   103
                  TabStop         =   0   'False
                  Top             =   195
                  Value           =   -1  'True
                  Width           =   615
               End
               Begin VB.OptionButton OptCr_ClBal 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Cr"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   720
                  TabIndex        =   102
                  TabStop         =   0   'False
                  Top             =   195
                  Width           =   615
               End
            End
            Begin VB.Frame DCr_Frame 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   3240
               TabIndex        =   100
               Top             =   600
               Width           =   1485
               Begin VB.OptionButton Cr_OPN 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Cr"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   840
                  TabIndex        =   5
                  Top             =   195
                  Width           =   615
               End
               Begin VB.OptionButton Dr_OPN 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Dr"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   4
                  Top             =   195
                  Value           =   -1  'True
                  Width           =   615
               End
            End
            Begin VB.TextBox TxtOpBal 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   1560
               TabIndex        =   3
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox TxtAcCode 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FF0000&
               Height          =   405
               Left            =   2400
               MaxLength       =   15
               TabIndex        =   0
               ToolTipText     =   "Enter Account Name (50)"
               Top             =   120
               Width           =   1695
            End
            Begin VB.TextBox TxtAcName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FF0000&
               Height          =   405
               Left            =   4200
               MaxLength       =   100
               TabIndex        =   1
               ToolTipText     =   "Enter Account Name (50)"
               Top             =   120
               Width           =   3615
            End
            Begin MSDataListLib.DataCombo GRP_DBCOM 
               Bindings        =   "GETACNT.frx":0BFE
               Height          =   360
               Left            =   9360
               TabIndex        =   2
               Top             =   120
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   635
               _Version        =   393216
               ForeColor       =   16711680
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
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
               Caption         =   "Closing Bal"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7680
               TabIndex        =   105
               Top             =   795
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Opening Bal"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   120
               TabIndex        =   99
               Top             =   735
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Code && Name"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   38
               Top             =   180
               Width           =   1320
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Type"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   7920
               TabIndex        =   37
               Top             =   195
               Width           =   1395
            End
         End
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account  List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   14040
      MouseIcon       =   "GETACNT.frx":0C12
      MousePointer    =   99  'Custom
      TabIndex        =   98
      Top             =   480
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   3060
      Index           =   0
      Left            =   360
      Top             =   5880
      Width           =   12975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   8820
      Left            =   360
      Top             =   960
      Width           =   13485
   End
End
Attribute VB_Name = "GETACNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte:            Dim LAcCode As String
Dim Last_MG_Code As Integer:        Dim MG_Code As Integer:                 Dim LPtyHead As Integer
Dim MBOOL As Boolean:               Dim LActive As Boolean:                 Dim LCGSTType As String
Dim GrpCode As Integer:             Public LAc_Name As String:              Public AccountDRec As ADODB.Recordset
Public TRec As ADODB.Recordset:     Dim RECGRID As ADODB.Recordset:         Dim ExRecGrid As ADODB.Recordset
Dim LSettlementDt As String:        Dim PartyHeadRec As ADODB.Recordset:    Dim ExBrokRec As ADODB.Recordset
Dim GrpRec As ADODB.Recordset:      Public AccountMRec As ADODB.Recordset:  Dim RecMulti  As ADODB.Recordset
Dim ExRec As ADODB.Recordset:       Dim ItemRec As ADODB.Recordset:         Dim LFilePress  As Integer
Dim GRIDREC As ADODB.Recordset:     Dim BrokRec As ADODB.Recordset:         Dim LBroker As String

Sub add_record()
    Dim TRec As ADODB.Recordset
    Fb_Press = 1:
    InterestCombo.ListIndex = 0
    
    ACCSETUP.Enabled = True:             CTTCmb.ListIndex = 2
    StampDutyCombo.ListIndex = 2:        RiskMCombo.ListIndex = 2:
    ChkCGST.Value = 0:                   ChkSGST.Value = 0
    ChkIGST.Value = 0:                   ChkUTT.Value = 0
    TxtStampRate.text = "0.000000":        vcDTP2.Value = DateValue(GFinBegin)
    SEBITaxCombo.ListIndex = 2:
    If UCase(GAcCodeType) = "N" Then            ''IF SET PARTYCODE=NUMERIC IN COMPANY SETUP
        mysql = "SELECT MAX(CAST(AC_CODE AS INT)) AS AC_CODE FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & ""
        Set TRec = Nothing:        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        LAcCode = Val(TRec!AC_CODE & vbNullString) + 1
        TxtAcCode.text = LAcCode
        TxtAcCode.Alignment = 1
    Else
        TxtAcCode.text = vbNullString
    End If
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT EXID,EXCODE FROM EXMAST WHERE COMPCODE=" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        Call ExRECSET
        Do While Not TRec.EOF
            ExRecGrid.AddNew::                      ExRecGrid!excode = TRec!excode
            ExRecGrid!ACEXCODE = vbNullString:      ExRecGrid!BILLINGTYPE = "Settle Rate"
            ExRecGrid!EXID = TRec!EXID
            ExRecGrid.Update
            TRec.MoveNext
        Loop
        Set ExGrid.DataSource = ExRecGrid: ExGrid.ReBind: ExGrid.Refresh
    End If
    Set TRec = Nothing
    If Left$(GBrokType, 1) = "P" Then
        ComboBrokType.ListIndex = 0
    ElseIf Left$(GBrokType, 1) = "T" Then
        ComboBrokType.ListIndex = 1
    ElseIf Left$(GBrokType, 1) = "O" Then
        ComboBrokType.ListIndex = 2
    ElseIf Left$(GBrokType, 1) = "I" Then
        ComboBrokType.ListIndex = 3
    ElseIf Left$(GBrokType, 1) = "C" Then
        ComboBrokType.ListIndex = 4
    End If
    Call Get_Selection(Fb_Press)
    TxtAcName.text = vbNullString
    ACCSETUP.Enabled = True
    TxtAcCode.SetFocus
    ComboBCycle.ListIndex = 1
    ChkServiceTax.Value = 0
    GRP_DBCOM.BoundText = 12
End Sub
Sub Save_Record()
    On Error GoTo err1
    Dim LPartyType As String:     Dim LMTranType As String:    Dim LMBrokType As String:     Dim LMBCycle As String
    Dim LCttType As String:       Dim LRiskMType As String:    Dim LSEBIType As String:      Dim LStmType As String
    Dim LSrvApp  As String:       Dim TRec As ADODB.Recordset: Dim LBillType As String:      Dim LFutCutBrok As String
    Dim TRec2 As ADODB.Recordset: Dim LPAcCode As String::     Dim CheckAcc As String:       Dim LMultiplier As String
    Dim LOp_Bal As Double:        Dim LActive As Byte:         Dim LXAc_Code  As String
    Dim LOptCutBrok As String
    Dim LACCID  As Long:     Dim TRec1 As ADODB.Recordset
    If TxtAcCode.text = "0" Or LenB(TxtAcCode.text) = 0 Then MsgBox "Account code can not be zero length.", vbCritical: If TxtAcCode.Enabled Then TxtAcCode.SetFocus: Exit Sub
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    
    TRec.Open "SELECT NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND NAME='" & TxtAcName.text & "' AND AC_CODE <> '" & TxtAcCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then MsgBox "Duplicate account " & TRec!NAME, vbCritical: ACCSETUP.Tab = 0: TxtAcName.SetFocus: Exit Sub
    If Last_MG_Code <> MG_Code And Fb_Press = 2 Then  'in case of modify
        If Last_MG_Code = 12 Or Last_MG_Code = 13 Or Last_MG_Code = 14 Then
            If MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14 Then
            Else
                Set TRec = Nothing: Set TRec = New ADODB.Recordset
                TRec.Open "SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PARTY  = '" & TxtAcCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then MsgBox "Can not change Account Type.Contract exists.", vbCritical: ACCSETUP.Tab = 0: TxtAcName.SetFocus: Exit Sub
            End If
        End If
    End If
    GrpRec.MoveFirst
    GrpRec.Find "CODE=" & Val(GRP_DBCOM.BoundText) & "", , adSearchForward
    If Not GrpRec.EOF Then MG_Code = GrpRec!G_CODE
    If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) Then
        If ExRecGrid.RecordCount > 0 Then ExRecGrid.MoveFirst
        Do While Not ExRecGrid.EOF
            If IsNull(ExRecGrid!ACEXCODE) Then ExRecGrid!ACEXCODE = vbNullString: ExRecGrid.Update
                
            If LenB(ExRecGrid!ACEXCODE) = 0 Or Trim(ExRecGrid!ACEXCODE) = "0" Then
            Else
                Set TRec = Nothing: Set TRec = New ADODB.Recordset
                TRec.Open "SELECT COMPCODE FROM ACCT_EX WHERE  EXID=" & ExRecGrid!EXID & " AND ACEXCODE ='" & ExRecGrid!ACEXCODE & "' AND AC_CODE <> '" & TxtAcCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then
                    MsgBox "Client code  " & ExRecGrid!ACEXCODE & " already exists for another Account.", vbCritical
                    Exit Sub
                End If
            End If
            ExRecGrid.MoveNext
        Loop
    End If
    If Fb_Press = 1 Then
        Set TRec = Nothing:                  Set TRec = New ADODB.Recordset
        TRec.Open "SELECT AC_CODE FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " AND AC_CODE='" & TxtAcCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Account Code Already Assigned.", vbExclamation, "Warning"
            Exit Sub
        End If
    End If
    Cnn.BeginTrans
    CNNERR = True
    LOp_Bal = IIf(Dr_OPN.Value, (Val(TxtOpBal.text & vbNullString) * -1), Val(TxtOpBal.text & vbNullString))
    LActive = IIf(ChkActive.Value = 1, 1, 0)
        
    If Fb_Press = 1 Then
        Call PInsert_AccountM(TxtAcCode.text, TxtAcName.text, Val(GRP_DBCOM.BoundText), MG_Code, LOp_Bal, LActive, LPtyHead)
    End If
    If Last_MG_Code <> MG_Code And Fb_Press = 2 Then 'in case of modify
        If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) Then
        ElseIf (Last_MG_Code = 12 Or Last_MG_Code = 13 Or Last_MG_Code = 14) Then
            Cnn.Execute "DELETE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & TxtAcCode.text & "'"
        End If
        If (MG_Code = 36 Or MG_Code = 37) Then
            Cnn.Execute "DELETE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & TxtAcCode.text & "'"
        End If
    End If
    If LenB(TxtAcName.text) <> 0 And MG_Code <> 0 Then
        If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14 Or MG_Code = 36 Or MG_Code = 37) Then
            If Fb_Press = 2 Then Cnn.Execute "DELETE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & TxtAcCode.text & "'"
            If ComboBrokType.ListIndex = 0 Then
                LMBrokType = "P"
            ElseIf ComboBrokType.ListIndex = 1 Then
                LMBrokType = "T"
            ElseIf ComboBrokType.ListIndex = 2 Then
                LMBrokType = "O"
            ElseIf ComboBrokType.ListIndex = 3 Then
                LMBrokType = "I"
            ElseIf ComboBrokType.ListIndex = 4 Then
                LMBrokType = "C"
            Else
                LMBrokType = "V"
            End If
            LMTranType = "P"
            If ComboBCycle.ListIndex = 0 Then
                LMBCycle = "S"
            Else
                LMBCycle = "D"
            End If
            If OptPType_Board.Value = True Then
                LPartyType = "3"
            ElseIf OptPType_Client.Value = True Then
                LPartyType = "1"
            ElseIf OptPType_Broker.Value = True Then
                LPartyType = "2"
            ElseIf OptPType_Pro.Value = True Then
                LPartyType = "4"
            Else
                LPartyType = "1"
            End If
            If CTTCmb.ListIndex = 0 Then
                LCttType = "R"
            ElseIf CTTCmb.ListIndex = 1 Then
                LCttType = "P"
            Else
                LCttType = "N"
            End If
            If RiskMCombo.ListIndex = 0 Then
                LRiskMType = "R"
            ElseIf RiskMCombo.ListIndex = 1 Then
                LRiskMType = "P"
            Else
                LRiskMType = "N"
            End If
            If SEBITaxCombo.ListIndex = 0 Then
                LSEBIType = "R"
            ElseIf SEBITaxCombo.ListIndex = 1 Then
                LSEBIType = "P"
            ElseIf SEBITaxCombo.ListIndex = 3 Then
                LSEBIType = "C"
            Else
                LSEBIType = "N"
            End If
            If StampDutyCombo.ListIndex = 0 Then
                LStmType = "R"
            ElseIf StampDutyCombo.ListIndex = 1 Then
                LStmType = "P"
            ElseIf StampDutyCombo.ListIndex = 3 Then
                LStmType = "S"
            ElseIf StampDutyCombo.ListIndex = 4 Then
                LStmType = "I"
            ElseIf StampDutyCombo.ListIndex = 5 Then
                LStmType = "B"
            Else
                LStmType = "N"
            End If
            
          
            If ComboBCycle.ListIndex = 0 Then
                LMBCycle = "S"
            Else
                LMBCycle = "D"
            End If
            
            If LenB(TxtStampRate.text) = 0 Then TxtStampRate.text = "0.000000"
            
            LSrvApp = IIf(ChkServiceTax.Value = 1, "Y", "N")
            LOptCutBrok = IIf(ChkCutBrok.Value, "1", "0")
            LFutCutBrok = IIf(ChkFutBrok.Value, "1", "0")
            LMultiplier = IIf(ChkMultiplier.Value, "1", "0")
            LACCID = Get_AccID(TxtAcCode.text)
            Call PInsert_Accountd(TxtAcCode.text, MG_Code, TxtAcName.text, Trim$(TxtAdd.text), Trim$(TxtCity.text), Trim$(TxtPin.text), Trim$(TxtPhoneO.text), _
            Trim(TxtPhoneR.text), Trim$(TxtFax.text), Trim$(TxtMobile.text), Trim$(TxtEmail.text), IIf(ChkPersonnelAc.Value = 1, "Y", "N"), _
            CStr(ChkInterest.Value), CStr(ChkServiceTax.Value), Trim$(TxtPANNo.text), Trim(TxtDirector.text), LMBrokType, Val(TxtBrokRate.text), "P", 0, 0, LMBCycle, _
            LPartyType, vbNullString, Trim$(TxtCINNo.text), Val(TxtStampRate.text), LCttType, vcDTP2.Value, LStmType, LPtyHead, LRiskMType, LSEBIType, _
            Val(TxtIntRate.text), Trim$(TxtGST.text), CStr(ChkCGST.Value), CStr(ChkSGST.Value), CStr(ChkIGST.Value), CStr(ChkUTT.Value), TxtUCC.text, Trim$(TxtState.text), Trim(TxtStateCode.text), Left(InterestCombo.text, 1), LFutCutBrok, LOptCutBrok, LMultiplier, Val(TxtSebiRate.text & vbNullString), Val(TxtCovertRate), Left(Combo2.text, 1), LACCID, Val(Text4.text))

            'MYSQL = "UPDATE ACCOUNTD SET OPTCUTBROK ='" & LSrvApp & "',FUTCUTBROK ='" & LFutCutBrok & "',MULTIPLIER='" & LMultiplier & "',SEBIRATE =" & Val(TxtSebiRate.text & vbNullString) & ", FROMDATE ='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
            'MYSQL = MYSQL & ",CURRTYPE='" & Left(Combo2.text, 1) & "',CURRRATE=" & Val(TxtCovertRate) & ""
            'MYSQL = MYSQL & " WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & TxtAcCode.text & "'"
            'Cnn.Execute MYSQL
        End If
        If Fb_Press = 1 Then
            'MYSQL = "INSERT INTO ACCOUNTM (COMPCODE,AC_CODE,NAME,GCODE,GRPCODE,PTYHEAD,OP_BAL,ACTIVE )"
            'MYSQL = MYSQL & " VALUES (" & GCompCode & ",'" & TxtAcCode.text & "','" & TxtAcName.text & " '," & Val(GRP_DBCOM.BoundText) & "," & MG_Code & "," & LPtyHead & "," & LOp_Bal & ",'" & LActive & "' )"
            'Cnn.Execute MYSQL
        Else
            mysql = "UPDATE ACCOUNTM SET "
            mysql = mysql & "  NAME ='" & TxtAcName.text & "'"
            mysql = mysql & " ,GCODE ='" & Val(GRP_DBCOM.BoundText) & "'"
            mysql = mysql & " ,GRPCODE ='" & MG_Code & "'"
            mysql = mysql & " ,PTYHEAD =" & LPtyHead & ""
            mysql = mysql & " ,OP_BAL =" & LOp_Bal & ""
            mysql = mysql & " ,ACTIVE  ='" & LActive & "'"
            mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & TxtAcCode.text & "'"
            Cnn.Execute mysql
        End If
        Cnn.Execute "DELETE FROM ACCT_EX WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & TxtAcCode.text & "'"
        LPAcCode = TxtAcCode.text
        LACCID = Get_AccID(TxtAcCode.text)
        If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) Then
            If Fb_Press = 1 Then
                mysql = "SELECT EXID,EXCODE,OPTIONS  FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE "
                Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                While Not TRec.EOF
                    If TRec!excode = "EQ" Or TRec!excode = "BEQ" Then
                       ' If Fb_Press = 2 Then
                        '    mysql = "SELECT count(*) FROM PEXBROK WHERE UPTOSTDT = '" & Format(GFinEnd, "YYYY/MM/DD") & "'  AND EXID = '" & TRec!EXID & " ' AND ACCID = '" & LACCID & "' AND INSTTYPE = 'CSH' "
                         '   Set TRec1 = Nothing: Set TRec1 = New ADODB.Recordset: TRec1.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                          '  If TRec1.EOF Then
                           '     Call PInsert_PExBrok(LPAcCode, TRec!excode, "P", 0, 0, 0, 0, "P", 0, 0, "P", 0, 0, "I", 0, GFinEnd, "CSH", 0, 0, TRec!EXID, LACCID)
                            'End If
                        'Else
                            Call PInsert_PExBrok(LPAcCode, TRec!excode, "P", 0, 0, 0, 0, "P", 0, 0, "P", 0, 0, "I", 0, GFinEnd, "CSH", 0, 0, TRec!EXID, LACCID)
                       ' End If
                    ElseIf TRec!excode = "CME OP" Then
                        If TRec!Options = "Y" Then
                            Call PInsert_PExBrok(LPAcCode, TRec!excode, "P", 0, 0, 0, 0, "P", 0, 0, "P", 0, 0, "I", 0, GFinEnd, "OPT", 0, 0, TRec!EXID, LACCID)
                        End If
                    Else
                        If TRec!Options = "Y" Then
                        '    If Fb_Press = 2 Then
                          '      mysql = "SELECT count(*) FROM PEXBROK WHERE UPTOSTDT = '" & Format(GFinEnd, "YYYY/MM/DD") & "'  AND EXID = '" & TRec!EXID & " ' AND ACCID = '" & LACCID & "' AND INSTTYPE = 'OPT' "
                           '     Set TRec1 = Nothing: Set TRec1 = New ADODB.Recordset: TRec1.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                            '    If TRec1.EOF Then
                              '      Call PInsert_PExBrok(LPAcCode, TRec!excode, "P", 0, 0, 0, 0, "P", 0, 0, "P", 0, 0, "I", 0, GFinEnd, "OPT", 0, 0, TRec!EXID, LACCID)
                             '   End If
                           ' Else
                                Call PInsert_PExBrok(LPAcCode, TRec!excode, "P", 0, 0, 0, 0, "P", 0, 0, "P", 0, 0, "I", 0, GFinEnd, "OPT", 0, 0, TRec!EXID, LACCID)
                            'End If
                        End If
                        Call PInsert_PExBrok(LPAcCode, TRec!excode, "P", 0, 0, 0, 0, "P", 0, 0, "P", 0, 0, "I", 0, GFinEnd, "FUT", 0, 0, TRec!EXID, LACCID)
                    End If
                    TRec.MoveNext
                Wend
            End If
            If ExRecGrid.RecordCount > 0 Then ExRecGrid.MoveFirst
                Do While Not ExRecGrid.EOF
                    If IsNull(ExRecGrid!BILLINGTYPE) Then
                        LBillType = "S"
                    Else
                        If ExRecGrid!BILLINGTYPE = "Settle Rate" Then
                            LBillType = "S"
                        ElseIf ExRecGrid!BILLINGTYPE = "Bid Ask Rate" Then
                            LBillType = "B"
                        ElseIf ExRecGrid!BILLINGTYPE = "Party Own Rate" Then
                            LBillType = "P"
                        Else
                            LBillType = "S"
                        End If
                    End If
                    Call PINSERT_ACCT_EX(TxtAcCode.text, ExRecGrid!excode, Trim(ExRecGrid!ACEXCODE), LBillType, ExRecGrid!EXID, LACCID)
                    ExRecGrid.MoveNext
                Loop
            End If
            
            'Create Self Branch
            If Check2.Value Then
                Dim MFMLYCODE As String
                Dim LContraCode As String
                Dim LContraID As Long
                Dim LFmlyID As Long
                
                MFMLYCODE = TxtAcCode.text
                Dim lrec2 As ADODB.Recordset
                Set lrec2 = Nothing
                Set lrec2 = New ADODB.Recordset
                mysql = "SELECT CONTRACTACC FROM EXMAST  WHERE COMPCODE =" & GCompCode & " "
                lrec2.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                If Not lrec2.EOF Then
                    LContraCode = lrec2!CONTRACTACC
                    LContraID = Get_AccID(LContraCode)
                End If
                
                Set TRec = Nothing: Set TRec = New ADODB.Recordset
                TRec.Open "SELECT FMLYID,FMLYCODE,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " AND (FMLYCODE ='" & MFMLYCODE & "' OR FMLYNAME ='" & MFMLYCODE & "') ", Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then
                    MsgBox "Duplicate family code. Already exists " & MFMLYCODE, vbExclamation, "Warning"
                Else
                    mysql = "INSERT INTO ACCFMLY( COMPCODE,FMLYCODE,FMLYNAME,FMLYHEAD,CONTRA_AC,POSTSETTLE,HEADID,CONTRAID )"
                    mysql = mysql & " VALUES (" & GCompCode & ",'" & MFMLYCODE & "','" & MFMLYCODE & "','" & TxtAcCode.text & "','" & LContraCode & "','N' ," & LACCID & "," & LContraID & " )"
                    Cnn.Execute mysql
                    
                    Dim LREC As ADODB.Recordset
                    Dim LAC_CODE2 As String
                    Dim I As Integer
                    Dim LParty As String
                    LFmlyID = Get_Fmlyid(MFMLYCODE)
                    
                    mysql = "INSERT INTO ACCFMLYD (COMPCODE,FMLYCODE,PARTY,FMLYID,ACCID) VALUES "
                    mysql = mysql & "(" & GCompCode & ",'" & MFMLYCODE & "','" & TxtAcCode.text & "'," & LFmlyID & "," & LACCID & ")"
                    Cnn.Execute mysql
            
                    mysql = "SELECT EXID,EXCODE,OPTIONS FROM EXMAST  WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE "
                    Set lrec2 = Nothing
                    Set lrec2 = New ADODB.Recordset
                    lrec2.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                    Do While Not lrec2.EOF
                        LAC_CODE2 = Get_PEXSBROK_AC_CODE(LFmlyID, LACCID, lrec2!EXID, GFinEnd, "FUT")
                        If LenB(LAC_CODE2) < 1 Then
                            If lrec2!excode = "BEQ" Or lrec2!excode = "EQ" Then
                                LAC_CODE2 = Get_PEXSBROK_AC_CODE(LFmlyID, LACCID, lrec2!EXID, GFinEnd, "CSH")
                                If LenB(LAC_CODE2) < 1 Then Call PINSERT_PEXSBROK(TxtAcCode.text, MFMLYCODE, lrec2!excode, "P", 0, "N", 0, GFinEnd, "CSH", 0, lrec2!EXID, LFmlyID, LACCID)
                            Else
                                Call PINSERT_PEXSBROK(TxtAcCode.text, MFMLYCODE, lrec2!excode, "P", 0, "N", 0, GFinEnd, "FUT", 0, lrec2!EXID, LFmlyID, LACCID)
                                If lrec2!Options = "Y" Then
                                    If lrec2!excode = "NSE" Or lrec2!excode = "MCX" Or lrec2!excode = "NCDX" Then
                                        LAC_CODE2 = Get_PEXSBROK_AC_CODE(LFmlyID, LACCID, lrec2!EXID, GFinEnd, "OPT")
                                        If LenB(LAC_CODE2) < 1 Then
                                            Call PINSERT_PEXSBROK(TxtAcCode.text, MFMLYCODE, lrec2!excode, "P", 0, "N", 0, GFinEnd, "OPT", 0, lrec2!EXID, LFmlyID, LACCID)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        lrec2.MoveNext
                   Loop
              End If
            End If
            ACCSETUP.Enabled = False
            Call CANCEL_RECORD
        Else
            MsgBox "Check Entries", vbExclamation, "Error"
            If TxtAcName.Enabled Then TxtAcName.SetFocus
        End If
        
        Cnn.CommitTrans
        CNNERR = False

        Exit Sub
err1:
    
    MsgBox err.Description, vbCritical, err.HelpFile
    If CNNERR = True Then
       Cnn.RollbackTrans
    End If
    
End Sub
Sub Total_Dcr()
    Dim TRec As ADODB.Recordset
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    Set TRec = AccountMRec.Clone
    If Not TRec.EOF Then TRec.MoveFirst
    TODr.text = "0": TOCr.text = "0"
    While Not TRec.EOF
        If TRec!OP_BAL < 0 Then   'DEBIT
           TODr.text = CStr(CDbl(TODr.text) + (TRec!OP_BAL * (-1)))
        ElseIf TRec!OP_BAL > 0 Then  'CREDIT
           TOCr.text = CStr(CDbl(TOCr.text) + TRec!OP_BAL)
        End If
        TRec.MoveNext
    Wend
    DIFF.text = "0.00"
    If CDbl(TODr.text) > CDbl(TOCr.text) Then
        DIFF.text = CStr(CDbl(TODr.text) - CDbl(TOCr.text)) + " Dr"
    ElseIf CDbl(TODr.text) < CDbl(TOCr.text) Then
        DIFF.text = CStr(CDbl(TOCr.text) - CDbl(TODr.text))
    End If
    TODr.text = Format(Round(Val(TODr.text), 2), "0.00")
    TOCr.text = Format(Round(Val(TOCr.text), 2), "0.00")
    DIFF.text = Format(Round(Val(DIFF.text), 2), "0.00")
End Sub
Sub CANCEL_RECORD()
    'Call ClearFormFn(GETACNT)
    TxtAccID.text = vbNullString
    TxtAcCode.text = vbNullString:              TxtAcName.text = vbNullString
    GRP_DBCOM.BoundText = vbNullString:         TxtOpBal.text = vbNullString
    TxtClBal.text = vbNullString:               ChkActive.Value = 1
    OptPType_Client.Value = True:               TxtFilterName.text = vbNullString
    DComboPartyHead.BoundText = vbNullString:   TxtAdd.text = vbNullString
    TxtCity.text = vbNullString:                TxtState.text = vbNullString
    TxtStateCode.text = vbNullString:           TxtPhoneO.text = vbNullString
    TxtPhoneR.text = vbNullString:              TxtPANNo.text = vbNullString
    TxtDirector.text = vbNullString:            TxtMobile.text = vbNullString
    TxtFax.text = vbNullString:                 TxtEmail.text = vbNullString
    TxtGST.text = vbNullString:                 TxtCINNo.text = vbNullString
    TxtUCC.text = vbNullString:                 ComboBrokType.ListIndex = 0
    TxtBrokRate.text = "0.00":                  TxtIntRate.text = "0.00"
    TxtSebiRate.text = "0.00":                   Text4.text = "0.00"
    Text5.text = ""
    TxtCovertRate.text = "0.00"
    Set DataGrid1.DataSource = Nothing:         Set DataGrid2.DataSource = Nothing
    TxtStampRate = "0.00":                      ChkMultiplier.Value = 0
    CmdAdd.Enabled = True:                      CmdMod.Enabled = True
    CmdSave.Enabled = False
    Combo2.ListIndex = 0
    Refresh_MultiGrid
    
    Frame4.Visible = True:
    Frame24.Visible = True
    DataGrid3.Visible = True
    Dr_OPN.Value = True:            Cr_OPN.Value = False
    LAc_Name = vbNullString:        Fb_Press = 0:                   LAcCode = "0":
    MG_Code = 0:                    Last_MG_Code = 0:
    GRP_DBCOM.Locked = False:       TxtAcName.Enabled = True:       TxtAcCode.Enabled = True:
    GRP_DBCOM.Enabled = True:       Fb_Press = 0
    CmdFilterOK.Caption = "Go"
    Call Get_Selection(10)
    TxtAcCode.Enabled = True
    ACCSETUP.Tab = 0: ACCSETUP.TabVisible(1) = False: ACCSETUP.TabVisible(2) = False: ACCSETUP.TabVisible(3) = False: ACCSETUP.TabVisible(4) = False
    ACCSETUP.Enabled = False:
    Call FillDataGrid
    Call Total_Dcr
    End Sub
Private Sub BrokerCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub BrokerCombo_Validate(Cancel As Boolean)
If LenB(BrokerCombo.BoundText) > 0 Then
    LBroker = Get_AccountMCode(BrokerCombo.BoundText)
    If LenB(LBroker) < 0 Then
        MsgBox "Invalid broker"
        BrokerCombo.BoundText = vbNullString
        Cancel = True
    End If
End If
End Sub

Private Sub ChkMultiplier_Click()
If ChkMultiplier.Value = 1 Then
    Frame2.Enabled = True
Else
    Frame2.Enabled = False
End If
End Sub

Private Sub billbycmb_GotFocus()
    If ExGrid.Col = 2 Then
        If ExGrid.text = "Settle Rate" Then
          billbycmb.ListIndex = 0
        ElseIf ExGrid.text = "BiD Ask Rate" Then
          billbycmb.ListIndex = 1
        ElseIf ExGrid.text = "Party Own Rate" Then
          billbycmb.ListIndex = 2
        End If
        billbycmb.Left = 5500
        billbycmb.Top = ExGrid.Top + Val(ExGrid.RowTop(ExGrid.Row))
        Sendkeys "%{DOWN}"
    End If
End Sub

Private Sub billbycmb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ExGrid.Col = 2
        If billbycmb.ListIndex <= 0 Then
            ExGrid.text = "Settle Rate"
        ElseIf billbycmb.ListIndex = 1 Then
            ExGrid.text = "Bid Ask Rate"
        ElseIf billbycmb.ListIndex = 2 Then
            ExGrid.text = "Party Own Rate"
        End If
        billbycmb.Visible = False
    ElseIf KeyAscii = 27 Then
        billbycmb.Visible = False
    End If
End Sub

Private Sub billbycmb_Validate(Cancel As Boolean)
    If billbycmb.ListIndex = 0 Then
        ExGrid.Col = 2
        ExGrid.text = "Settle Rate"
    ElseIf billbycmb.ListIndex = 1 Then
        ExGrid.Col = 2
        ExGrid.text = "Bid Ask Rate"
    ElseIf billbycmb.ListIndex = 2 Then
        ExGrid.Col = 2
        ExGrid.text = "Party Own Rate"
    End If
End Sub


Private Sub Command2_Click()

If Frame6.Visible = True Then
    Frame6.Visible = False
Else
    Frame6.Visible = True
End If

End Sub

Private Sub Command5_Click()
    Frame25.Visible = False
End Sub

Private Sub CRViewer1_CloseButtonClicked(UseDefault As Boolean)
Frame24.Visible = True
DataGrid3.Visible = True
End Sub

Private Sub ExCombo_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub ExCombo_Validate(Cancel As Boolean)
If LenB(ExCombo.BoundText) < 1 Then
    Cancel = True
End If
End Sub

Private Sub ExGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Sendkeys "{tab}"
    '    If ExGrid.Col = 1 Then
    '        billbycmb.Visible = True
    '        billbycmb.SetFocus
    '    Else
        If ExGrid.Col = 2 Then
            'If LenB(ExGrid.Columns(2).text) < 1 Then
                billbycmb.Visible = True
                billbycmb.SetFocus

            'End If
        End If
    End If
End Sub


Private Sub ItemCombo_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub TxtAcName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtAcName_Validate(Cancel As Boolean)
    Dim VALRETB As Long
    Dim TRec As ADODB.Recordset
    If Len(Trim(TxtAcName.text)) > 1 Then
        LAc_Name = TxtAcName.text
        mysql = "SELECT COMPCODE FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " AND NAME='" & TxtAcName.text & "'"
        Set TRec = Nothing:        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF And Fb_Press = 1 Then
            MsgBox "Account Name already exists.", vbExclamation, "Warning"
            LAc_Name = vbNullString
            Cancel = True
            Set TRec = Nothing
            Exit Sub
        End If
        Set TRec = Nothing
    End If
    If Fb_Press = 0 Then
    Else
        If Len(Trim(TxtAcName.text)) < 1 Then
            Beep 500, 500
            Cancel = True
        End If
    End If
End Sub
Private Sub ACCSETUP_LostFocus()
    If ACCSETUP.Enabled = True Then
        If ACCSETUP.Tab = 0 Then
            If Fb_Press = 1 Then
                TxtAcCode.SetFocus
            Else
                TxtAcName.SetFocus
            End If
        ElseIf ACCSETUP.Tab = 1 Then
            TxtAdd.SetFocus
        ElseIf ACCSETUP.Tab = 2 Then
        
        End If
    End If
End Sub
Private Sub CSTDATE_Validate(Cancel As Boolean)
    ACCSETUP.Tab = 2
    TxtPANNo.SetFocus
End Sub
Private Sub ChkPersonnelAc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub ClientCmb_GotFocus()
Sendkeys "%{down}"
End Sub
Private Sub ClientCmb_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub ComboBrokType_GotFocus()
Sendkeys "%{DOWN}"
End Sub
Private Sub ComboBrokType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Combo6_GotFocus()
Sendkeys "%{DOWN}"
End Sub
Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub ComboBCycle_GotFocus()
Sendkeys "%{DOWN}"
End Sub
Private Sub StampDutyCombo_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub ComboBCycle_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub StampDutyCombo_Validate(Cancel As Boolean)
If StampDutyCombo.ListIndex >= 3 Then
    Frame5.Visible = True
End If
End Sub
Private Sub RiskMCombo_GotFocus()
Sendkeys "%{down}"
End Sub
Private Sub SEBITaxCombo_GotFocus()
Sendkeys "%{down}"
End Sub
Private Sub RiskMCombo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub SEBITaxCombo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub CmdFilterOK_Click()
If CmdFilterOK.Caption = "Go" Then
    CmdFilterOK.Caption = "Clear"
Else
    TxtFilterName.text = vbNullString
    TxtFilterCode.text = vbNullString
    CmdFilterOK.Caption = "Go"
End If
Call FillDataGrid
End Sub

Private Sub Cr_OPN_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub cttcmb_GotFocus()
Sendkeys "%{down}"
End Sub


Private Sub cttcmb_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub StampDutyCombo_GotFocus()
Sendkeys "%{down}"
End Sub
Private Sub StampDutyCombo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub DComboPartyHead_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub DComboPartyHead_GotFocus()
Sendkeys "%{DOWN}"
End Sub
Private Sub DComboPartyHead_Validate(Cancel As Boolean)
LPtyHead = DComboPartyHead.BoundText
ACCSETUP.Tab = 1
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
End Sub

Private Sub Dr_OPN_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Private Sub GRP_DBCOM_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label12.ForeColor = &H0&
End Sub
Private Sub Form_Paint()
    On Error GoTo err1
    Dim LGCode  As Long
    If GETMAIN.ActiveForm.NAME = Me.NAME Then
        LGCode = Val(GRP_DBCOM.BoundText)
        mysql = "SELECT CODE, G_NAME,G_CODE FROM AC_GROUP ORDER BY G_NAME"
        Set GrpRec = Nothing: Set GrpRec = New ADODB.Recordset
        GrpRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not GrpRec.EOF Then
            Set GRP_DBCOM.RowSource = GrpRec
            GRP_DBCOM.ListField = "G_NAME"
            GRP_DBCOM.BoundColumn = "CODE"
            GRP_DBCOM.BoundText = Val(LGCode)
        End If
        Call Get_Selection(Fb_Press)
        Me.BackColor = GETMAIN.BackColor
        ACCSETUP.BackColor = GETMAIN.BackColor
    End If
    Exit Sub
err1:
    If err.Number = Val(91) Then
        Resume Next
    Else
        MsgBox err.Description, vbCritical, err.HelpFile
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        Call Get_Selection(10)
        CRViewer1.Visible = False
        Frame4.Visible = True
        Frame24.Visible = True
        DataGrid3.Visible = True
        Cancel = 1
    Else
        Call CANCEL_RECORD
        GETMAIN.StatusBar1.Panels(1).text = vbNullString
        Unload Me
    End If
End Sub
Private Sub GRP_DBCOM_Change()
    GRP_DBCOM.DataChanged = True
End Sub
Private Sub GRP_DBCOM_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub GRP_DBCOM_LostFocus()
    If LenB(TxtAcName) <> 0 Then
        If LenB(GRP_DBCOM.text) = 0 Then
            GRP_DBCOM.SetFocus
        ElseIf GRP_DBCOM.MatchedWithList And GRP_DBCOM.DataChanged Then
            GrpRec.MoveFirst
            GrpRec.Find "CODE=" & Val(GRP_DBCOM.BoundText) & "", , adSearchForward
            If Not GrpRec.EOF Then
                MG_Code = GrpRec!G_CODE
                GrpCode = GRP_DBCOM.BoundText
            End If
            If (Last_MG_Code <> MG_Code And Fb_Press = 2 And (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14)) Or ((MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) And Fb_Press = 1) Then
                Set TRec = Nothing: Set TRec = New ADODB.Recordset
                TRec.Open "SELECT EXID,EXCODE FROM EXMAST WHERE COMPCODE=" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then
                    Call ExRECSET
                    Do While Not TRec.EOF
                        ExRecGrid.AddNew
                        ExRecGrid!excode = TRec!excode
                        ExRecGrid!EXID = TRec!EXID
                        ExRecGrid!ACEXCODE = vbNullString
                        ExRecGrid!BILLINGTYPE = "Settle Rate"
                        ExRecGrid.Update
                        TRec.MoveNext
                    Loop
                    Set TRec = Nothing
                    Set ExGrid.DataSource = ExRecGrid: ExGrid.ReBind: ExGrid.Refresh
                End If
                ACCSETUP.TabVisible(1) = True:  ACCSETUP.TabVisible(2) = True: ACCSETUP.TabVisible(3) = True:: ACCSETUP.TabVisible(4) = True
            ElseIf (MG_Code = 36 Or MG_Code = 37) Then
                ACCSETUP.TabVisible(1) = True:  ACCSETUP.TabVisible(2) = False: ACCSETUP.TabVisible(3) = False:: ACCSETUP.TabVisible(4) = False
            Else
                ACCSETUP.TabVisible(1) = False: ACCSETUP.TabVisible(2) = False: ACCSETUP.TabVisible(3) = False:: ACCSETUP.TabVisible(4) = False
            End If
        ElseIf Not GRP_DBCOM.MatchedWithList And GRP_DBCOM.DataChanged Then
            MsgBox "Invalid Group", vbInformation, "Error"
            GRP_DBCOM.text = ""
            GRP_DBCOM.SetFocus
        End If
    End If
    GRP_DBCOM.DataChanged = False
End Sub
Private Sub Form_Load()
    billbycmb.ListIndex = 0
    vcDTP2.Value = DateValue(GFinBegin)
    Call Get_Selection(10)
    CmdAdd.Enabled = True
    CmdMod.Enabled = True
    CmdSave.Enabled = False

    CTTCmb.ListIndex = 2
    StampDutyCombo.ListIndex = 2
    ACCSETUP.Enabled = False
    'Last Settlement Date
    LSettlementDt = vbNullString: Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT MAX(SETDATE) AS  MaxSettleDate FROM SETTLE WHERE COMPCODE = " & GCompCode & "", Cnn, adOpenKeyset, adLockReadOnly
    If Not TRec.EOF Then LSettlementDt = IIf(IsNull(TRec!MaxSettleDate), GFinEnd, TRec!MaxSettleDate)
    mysql = "SELECT EXCODE FROM EXMAST WHERE COMPCODE =" & GCompCode & ""
    Set ExRec = Nothing
    Set ExRec = New ADODB.Recordset
    ExRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Set ExCombo.RowSource = ExRec
    ExCombo.ListField = "EXCODE"
    ExCombo.BoundColumn = "EXCODE"
    
    mysql = "SELECT ITEMCODE,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " ORDER BY ITEMNAME"
    Set ItemRec = Nothing
    Set ItemRec = New ADODB.Recordset
    ItemRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Set ItemCombo.RowSource = ItemRec
    ItemCombo.ListField = "ITEMCODE"
    ItemCombo.BoundColumn = "ITEMCODE"
    
    
    mysql = "SELECT * FROM PARTYHEAD ORDER BY HEADCODE"
    Set PartyHeadRec = Nothing
    Set PartyHeadRec = New ADODB.Recordset
    PartyHeadRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If PartyHeadRec.EOF Then
        Label39.Visible = False
        DComboPartyHead.Visible = False
    Else
        Set DComboPartyHead.RowSource = PartyHeadRec
        DComboPartyHead.ListField = "HEADNAME"
        DComboPartyHead.BoundColumn = "HEADCODE"
    End If
    mysql = "SELECT * FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
    Set AccountDRec = Nothing
    Set AccountDRec = New ADODB.Recordset
    AccountDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
    Set BrokRec = Nothing
    Set BrokRec = New ADODB.Recordset
    BrokRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Set BrokerCombo.RowSource = BrokRec
    BrokerCombo.ListField = "NAME"
    BrokerCombo.BoundColumn = "AC_CODE"
    
    Call RecSet
    Call CANCEL_RECORD
End Sub
Private Sub Label12_Click()
    Call List_Rec
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label12.ForeColor = &HC00000
End Sub

Private Sub TxtIntRate_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub TxtMultiplier_Validate(Cancel As Boolean)
TxtMultiplier.text = Format(Val(TxtMultiplier.text), "0.00")
End Sub

Private Sub TxtMultiplier_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub

Private Sub TxtMultiplier_GotFocus()
    TxtMultiplier.SelLength = Len(TxtMultiplier.text)
End Sub

Private Sub TxtOldCode_Validate(Cancel As Boolean)
Dim LName As String
Dim LACCID As Long
LACCID = Get_AccID(TxtOldCode.text)
If LACCID = 0 Then
    MsgBox "Invalid old Code "
    TxtOldCode.text = vbNullString
    Cancel = True
Else
    LName = Get_AccountName(LACCID)
    If LenB(LName) < 1 Then
        MsgBox "Invalid old Code "
        TxtOldCode.text = vbNullString
        Cancel = True
    Else
        TxtoldName.text = LName
        TxtnewName.text = LName
    End If
End If
End Sub

Private Sub TxtOpBal_GotFocus()
    TxtOpBal.SelLength = Len(TxtOpBal.text)
End Sub
Private Sub TxtOpBal_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtOpBal_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtOpBal_LostFocus()
    TxtOpBal.text = Format(Val(TxtOpBal.text), "0.00")
End Sub

Private Sub TxtCovertRate_GotFocus()
    TxtCovertRate.SelLength = Len(TxtCovertRate.text)
End Sub
Private Sub TxtCovertRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtCovertRate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtCovertRate_LostFocus()
    TxtCovertRate.text = Format(Val(TxtCovertRate.text & vbNullString), "0.00")
End Sub

Private Sub TxtSno_Validate(Cancel As Boolean)
Dim TRec As ADODB.Recordset
Set TRec = Nothing
Set TRec = New ADODB.Recordset
mysql = "SELECT * FROM PARTYMULTI  WHERE COMPCODE =" & GCompCode & " AND PARTY  ='" & TxtAcCode.text & "' AND SNO  =" & Val(Trim(TxtSno.text)) & ""
TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not TRec.EOF Then
    ExCombo.BoundText = TRec!excode
    ItemCombo.BoundText = TRec!ITEMCODE
    TxtMultiplier.text = Format(TRec!Rate, "0.00")
    ExCombo.SetFocus
Else
    MsgBox "No Entry for this SNo "
    TxtSno.SetFocus
End If

End Sub
Private Sub TxtStampRate_GotFocus()
    TxtStampRate.SelLength = Len(TxtStampRate.text)
End Sub
Private Sub TxtStampRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtStampRate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtStampRate_LostFocus()
    TxtStampRate.text = Format(Val(TxtStampRate.text), "0.000000")
End Sub
Private Sub TxtSEBIRATE_GotFocus()
    TxtSebiRate.SelLength = Len(TxtSebiRate.text)
End Sub
Private Sub TxtSEBIRATE_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtSEBIRATE_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtSEBIRATE_LostFocus()
    TxtSebiRate.text = Format(Val(TxtSebiRate.text), "0.000000")
End Sub
Private Sub TxtOpBal_Validate(Cancel As Boolean)
    If SYSTEMLOCK(Date) Then
        MsgBox "Sorry System Locked.  No Modification Allowed for Opening Balances"
        Cancel = True
    End If
End Sub
Private Sub OptPType_Broker_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub OptPType_Client_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtPANNo_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtEmail_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtMobile_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtFilterName_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtAcCode_GotFocus()
    TxtAcCode.SelLength = Len(TxtAcCode.text)
End Sub
Private Sub TxtAcCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If UCase(GAcCodeType) = "N" Then
        If IsNumeric(TxtAcCode.text) Or LenB(TxtAcCode.text) = 0 Then
            If KeyCode = 13 Then Sendkeys "{tab}"
        Else
            MsgBox "Numeric account code defined in company setup.", vbCritical: TxtAcCode.text = LAcCode: TxtAcCode.SetFocus
        End If
    Else
        If KeyCode = 13 Then Sendkeys "{tab}"
    End If
End Sub
Private Sub TxtAcCode_Validate(Cancel As Boolean)
    If Fb_Press = 1 Then
        mysql = "SELECT AC_CODE FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " AND AC_CODE ='" & TxtAcCode.text & "'"
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Account Code already exists "
            TxtAcCode.text = vbNullString
            Cancel = True
            TxtAcCode.SetFocus
            Set TRec = Nothing
            Exit Sub
        End If
        Set TRec = Nothing
    End If
    If Fb_Press <> 0 Then TxtAcName.SetFocus
End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtDirector_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtBrokRate_GotFocus()
    TxtBrokRate.SelLength = Len(TxtBrokRate.text)
End Sub
Private Sub TxtBrokRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtBrokRate_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtBrokRate_Validate(Cancel As Boolean)
    TxtBrokRate.text = Format(TxtBrokRate.text, "0.00000")
End Sub
Private Sub TxtPhoneO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtPhoneR_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Sub Delete_Record()

Call Delete_Record_loop
Text5.text = ""
Exit Sub

    LAcCode = TxtAcCode.text
    mysql = "SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PARTY='" & TxtAcCode.text & "'"
    Set TRec = Nothing:    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        MsgBox "Transaction Exists Can't Delete Account.", vbExclamation, "Error"
        Call CANCEL_RECORD
        Exit Sub
    End If
    mysql = "SELECT TOP 1 COMPCODE FROM VCHAMT WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & TxtAcCode.text & "'"
    Set TRec = Nothing:    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        MsgBox "Transaction Exists Can't Delete Account.", vbExclamation, "Error"
        Call CANCEL_RECORD
        Exit Sub
    End If
    Set TRec = Nothing
    If MsgBox(String(5, " ") & "You are about to Delete this record. Are you sure . ." & String(10, " "), vbYesNo + vbCritical + vbDefaultButton1, "Confirmation") = vbYes Then
        If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) Then
        Cnn.Execute "DELETE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM ACCT_EX  WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND PARTY='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM PEXBROK  WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM PEXSBROK WHERE COMPCODE =" & GCompCode & " AND PARTY='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM PITBROK  WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM PITSBROK WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & LAcCode & "'"
        Cnn.Execute "DELETE FROM PARTYMULTI WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & LAcCode & "'"
    End If
    End If
    Call CANCEL_RECORD
    Call Total_Dcr
    Exit Sub
    
End Sub
Sub Delete_Record_loop()

    Dim Arr() As String
    Dim flag As Boolean
    Dim I As Integer
    flag = False
    Arr = Split(Text5.text, ",")
    
    For I = 0 To UBound(Arr)
    
        If Arr(I) <> "" Then
            LAcCode = Arr(I)
            mysql = "SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PARTY='" & Arr(I) & "'"
            Set TRec = Nothing:    Set TRec = New ADODB.Recordset
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                ImportIssues ("Account:" + Arr(I) + ", Transaction Exists Can't Delete Account.")
                flag = True
            Else
                mysql = "SELECT TOP 1 COMPCODE FROM VCHAMT WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & Arr(I) & "'"
                Set TRec = Nothing:    Set TRec = New ADODB.Recordset
                TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then
                    flag = True
                    ImportIssues ("Account:" + Arr(I) + ", Transaction Exists Can't Delete Account.")
                End If
            End If
            
            If Not flag Then
                Set TRec = Nothing
                'If MsgBox(String(5, " ") & "You are about to Delete this record. Are you sure . ." & String(10, " "), vbYesNo + vbCritical + vbDefaultButton1, "Confirmation") = vbYes Then
                    If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) Then Cnn.Execute "DELETE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM ACCT_EX  WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND PARTY='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM PEXBROK  WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM PEXSBROK WHERE COMPCODE =" & GCompCode & " AND PARTY='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM PITBROK  WHERE COMPCODE =" & GCompCode & " AND AC_CODE='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM PITSBROK WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & LAcCode & "'"
                    Cnn.Execute "DELETE FROM PARTYMULTI WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & LAcCode & "'"
                'End If
                Call CANCEL_RECORD
                Call Total_Dcr
                ImportIssues ("Account:" + Arr(I) + ", Account deleted.")
            End If
        End If
    Next
    Frame25.Visible = True
    Exit Sub
    
End Sub
Sub ACCOUNT_ACCESS()
On Error GoTo Error1
    Dim LAC_CODE As String
    Dim LClBal As Double
    If LenB(TxtAcCode.text) <> 0 Then
        LAC_CODE = TxtAcCode.text
        'DataList1.Locked = True
        AccountMRec.MoveFirst
        AccountMRec.Find "AC_CODE = '" & LAC_CODE & "'", , adSearchForward
        If Not AccountMRec.EOF Then
            With AccountMRec
                LAc_Name = !NAME:
                TxtAccID.text = !ACCID
                'OldName = !NAME
                GrpCode = !GrpCode & vbNullString
                TxtAcCode.text = !AC_CODE & vbNullString
                LAcCode = !AC_CODE & vbNullString
                TxtAcCode.Enabled = False
                TxtAcName.text = LAc_Name
                Last_MG_Code = Val(!GrpCode & vbNullString)
                If .Fields("OP_BAL") < 0 Then
                    TxtOpBal.text = Format(Abs(!OP_BAL), "0.00")
                    Dr_OPN.Value = True
                    Cr_OPN.Value = False
                Else
                    TxtOpBal.text = Format(Abs(!OP_BAL), "0.00")
                    Cr_OPN.Value = True
                    Dr_OPN.Value = False
                End If
                
                LClBal = Net_DrCr(LAC_CODE, Date)
                LClBal = LClBal + !OP_BAL
                Frame22.Enabled = True
                If LClBal < 0 Then
                    TxtClBal.text = Format(Abs(LClBal), "0.00")
                    OptDr_Clbal.Value = True
                    OptCr_ClBal.Value = False
                Else
                    TxtClBal.text = Format(Abs(LClBal), "0.00")
                    OptCr_ClBal.Value = True
                    OptDr_Clbal.Value = False
                End If
                Frame22.Enabled = False
                MG_Code = !GrpCode & ""
                If !PTYHEAD <> 0 Then DComboPartyHead.BoundText = !PTYHEAD
                LPtyHead = IIf(IsNull(!PTYHEAD), 0, !PTYHEAD)
                If MG_Code <> 0 Then
                    GRP_DBCOM.BoundText = Val(!GCODE & vbNullString)
                    GRP_DBCOM.DataChanged = False
                Else
                    MsgBox String(10, " ") & "Group Missing" & String(15, " "), vbInformation, "Error"
                End If
                LActive = IIf(IsNull(!Active), True, !Active)
                If LActive = True Then
                    ChkActive.Value = 1
                Else
                    ChkActive.Value = 0
                End If
            End With
            ''''From Party_DB Ado
            If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14 Or MG_Code = 36 Or MG_Code = 37) Then
                AccountDRec.Requery
                AccountDRec.MoveFirst
                AccountDRec.Find "AC_CODE ='" & LAcCode & "'", , adSearchForward
                If Not AccountDRec.EOF Then
                    With AccountDRec
                        TxtAcName.text = !NAME & vbNullString
                        TxtAdd.text = !ADD1 & vbNullString:                               TxtCity.text = !City & vbNullString
                        TxtPin.text = !Pin & vbNullString:                                TxtPhoneO.text = !PhoneO & vbNullString
                        TxtPhoneR.text = !PhoneR & vbNullString:                          TxtFax.text = !Fax & vbNullString
                        TxtMobile.text = !Mobile & vbNullString:                          TxtEmail.text = !Email & vbNullString
                        TxtPANNo.text = IIf(IsNull(!PANNO), vbNullString, !PANNO):        TxtDirector.text = IIf(IsNull(!DIRECTOR), vbNullString, !DIRECTOR)
                        'TxtCINNo.text = IIf(IsNull(!CSTNO), vbNullString, Trim(!CSTNO)):
                        TxtGST.text = IIf(IsNull(!GSTIN), vbNullString, !GSTIN)
                        TxtUCC.text = !UCC & vbNullString:                                TxtState.text = !State & vbNullString
                        TxtStateCode = !StateCode & vbNullString
                        If IsNull(!PARTYTYPE) Then
                            OptPType_Board.Value = True
                        Else
                            If !PARTYTYPE = "3" Then
                                OptPType_Board.Value = True
                            ElseIf !PARTYTYPE = "2" Then
                                OptPType_Broker.Value = True
                            ElseIf !PARTYTYPE = "1" Then
                                OptPType_Client.Value = True
                            ElseIf !PARTYTYPE = "4" Then
                                OptPType_Pro.Value = True
                            Else
                                OptPType_Client.Value = True
                            End If
                        End If
                        If !CURRTYPE = "S" Then
                            Combo2.ListIndex = 0
                        ElseIf !CURRTYPE = "H" Then
                            Combo2.ListIndex = 1
                        ElseIf !CURRTYPE = "R" Then
                            Combo2.ListIndex = 2
                        ElseIf !CURRTYPE = "F" Then
                            Combo2.ListIndex = 3
                        End If
                        TxtCovertRate.text = Format(Val(!CURRRATE), "0.00")
                        If IsNull(!PERSONNELAC) Then
                            ChkPersonnelAc.Value = 0
                        Else
                            ChkPersonnelAc.Value = IIf(!PERSONNELAC = "Y", 1, 0)
                        End If
                        
                        ChkInterest.Value = IIf(IsNull(!INTAPP), 0, !INTAPP)
                        If IsNull(!SRTAXAPP) Then
                            ChkServiceTax.Value = 0
                        Else
                            If !SRTAXAPP = "1" Then
                                ChkServiceTax.Value = 1
                            Else
                                ChkServiceTax.Value = 0
                            End If
                        End If
                        If IsNull(!CTTTYPE) Then
                            CTTCmb.ListIndex = 2
                        Else
                            If !CTTTYPE = "R" Then
                                CTTCmb.ListIndex = 0
                            ElseIf !CTTTYPE = "P" Then
                                CTTCmb.ListIndex = 1
                            Else
                                CTTCmb.ListIndex = 2
                            End If
                        End If
                        If IsNull(!RISKMTYPE) Then
                            RiskMCombo.ListIndex = 2
                        Else
                            If !RISKMTYPE = "R" Then
                                RiskMCombo.ListIndex = 0
                            ElseIf !RISKMTYPE = "P" Then
                                RiskMCombo.ListIndex = 1
                            Else
                                RiskMCombo.ListIndex = 2
                            End If
                        End If
                        If IsNull(!SEBITYPE) Then
                            SEBITaxCombo.ListIndex = 2
                        Else
                            If !SEBITYPE = "R" Then
                                SEBITaxCombo.ListIndex = 0
                            ElseIf !SEBITYPE = "P" Then
                                SEBITaxCombo.ListIndex = 1
                            ElseIf !SEBITYPE = "C" Then
                                SEBITaxCombo.ListIndex = 3
                            Else
                                SEBITaxCombo.ListIndex = 2
                            End If
                        End If
                        ChkCGST.Value = IIf(IsNull(!CGST), 0, !CGST)
                        ChkSGST.Value = IIf(IsNull(!SGST), 0, !SGST)
                        ChkCutBrok.Value = IIf(IsNull(!OptCutBrok), 1, !OptCutBrok)
                        ChkFutBrok.Value = IIf(IsNull(!FUTCutBrok), 1, !FUTCutBrok)
                        ChkMultiplier.Value = IIf(IsNull(!MULTIPLIER), 0, !MULTIPLIER)
                        
                        'txtsenirate.text = format(txtseb
                        If ChkMultiplier.Value = 1 Then
                            Frame2.Enabled = True
                            Call Refresh_MultiGrid
                        End If
                        
                        If IsNull(!INTTYPE) Then
                            InterestCombo.ListIndex = 0
                        Else
                            If !INTTYPE = "W" Then
                                InterestCombo.ListIndex = 0
                            ElseIf !INTTYPE = "D" Then
                                InterestCombo.ListIndex = 2
                            ElseIf !INTTYPE = "P" Then
                                InterestCombo.ListIndex = 3
                            ElseIf !INTTYPE = "R" Then
                                InterestCombo.ListIndex = 4
                            ElseIf !INTTYPE = "B" Then
                                InterestCombo.ListIndex = 5
                            Else
                                InterestCombo.ListIndex = 1
                            End If
                        End If
                        If IsNull(!IGST) Then
                            ChkIGST.Value = 0
                        Else
                            If !IGST = "0" Then
                                ChkIGST.Value = 0
                            ElseIf !IGST = "1" Then
                                ChkIGST.Value = 1
                            Else
                                ChkIGST.Value = 0
                            End If
                        End If
                        If IsNull(!UTT) Then
                            ChkUTT.Value = 0
                        Else
                            If !UTT = "0" Then
                                ChkUTT.Value = 0
                            ElseIf !UTT = "1" Then
                                ChkUTT.Value = 1
                            Else
                                ChkUTT.Value = 0
                            End If
                        End If
                        If IsNull(!APPLYON) Then
                            StampDutyCombo.ListIndex = 2
                        Else
                            If !APPLYON = "R" Then
                                StampDutyCombo.ListIndex = 0
                            ElseIf !APPLYON = "P" Then
                                StampDutyCombo.ListIndex = 1
                            ElseIf !APPLYON = "S" Then
                                StampDutyCombo.ListIndex = 3
                            ElseIf !APPLYON = "I" Then
                                StampDutyCombo.ListIndex = 4
                            ElseIf !APPLYON = "B" Then
                                StampDutyCombo.ListIndex = 5
                            Else
                                StampDutyCombo.ListIndex = 2
                            End If
                        End If
                        If Not IsNull(!FROMDATE) Then
                            If IsDate(!FROMDATE) Then
                                vcDTP2.Value = !FROMDATE
                            Else
                               vcDTP2.Value = GFinBegin
                            End If
                        Else
                            vcDTP2.Value = GFinBegin
                        End If
                        If StampDutyCombo.ListIndex >= 3 Then
                            Frame5.Visible = True
                        Else
                            Frame5.Visible = False
                        End If
                        If !broktype = "P" Then
                            ComboBrokType.ListIndex = 0
                        ElseIf !broktype = "T" Then
                            ComboBrokType.ListIndex = 1
                        ElseIf !broktype = "O" Then
                            ComboBrokType.ListIndex = 2
                        ElseIf !broktype = "I" Then
                            ComboBrokType.ListIndex = 3
                        ElseIf !broktype = "C" Then
                            ComboBrokType.ListIndex = 4
                        Else
                            ComboBrokType.ListIndex = 5
                        End If
                        If IsNull(!BCYCLE) Then
                            ComboBCycle.ListIndex = 1
                        Else
                            If !BCYCLE = "S" Then
                                ComboBCycle.ListIndex = 0
                            Else
                                ComboBCycle.ListIndex = 1
                            End If
                        End If
                        TxtIntRate.text = Format(!INTRATE, "0.00000")
                        TxtBrokRate.text = Format(!brokrate, "0.00000")
                        TxtStampRate.text = Format(!STTRATE, "0.000000")
                        TxtSebiRate.text = Format(!SEBIRATE, "0.000000")
                        Text4.text = Format(!CrLimit, "0.00000")
                    End With
                    If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) Then
                        Call RecSet
                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
                        mysql = "SELECT EXCODE,BROKTYPE,BROKRATE,STDRATE,TRANTYPE,TRANRATE,MARTYPE,MARRATE,INSTTYPE,BROKRATE2,MBROKRATE,MBROKRATE2,UPTOSTDT"
                        mysql = mysql & " FROM PEXBROK AS PB WHERE PB.COMPCODE =" & GCompCode & " "
                        mysql = mysql & " AND  PB.AC_CODE='" & LAcCode & "' ORDER BY PB.EXCODE,INSTTYPE,PB.UPTOSTDT"
                        TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                        If Not TRec.EOF Then
                            Do While Not TRec.EOF
                                ExBrokRec.AddNew
                                ExBrokRec!excode = TRec!excode:         ExBrokRec!INSTTYPE = TRec!INSTTYPE
                                ExBrokRec!brokrate = TRec!brokrate:     ExBrokRec!BROKRATE2 = TRec!BROKRATE2
                                ExBrokRec!MBROKRATE = TRec!MBROKRATE:   ExBrokRec!MBROKRATE2 = TRec!MBROKRATE2
                                ExBrokRec!TRANRATE = TRec!TRANRATE:     ExBrokRec!STDRATE = TRec!STDRATE
                                ExBrokRec!MARRATE = TRec!MARRATE:       ExBrokRec!UPTOSTDT = TRec!UPTOSTDT
                                Select Case TRec!broktype
                                Case ""
                                    ExBrokRec!broktype = "Transaction Wise"
                                Case "T"
                                    ExBrokRec!broktype = "Transaction Wise"
                                Case "O"
                                    ExBrokRec!broktype = "Opening Sauda"
                                Case "P"
                                    ExBrokRec!broktype = "Percentage Wise"
                                Case "I"
                                    ExBrokRec!broktype = "Intraday Wise"
                                Case "C"
                                    ExBrokRec!broktype = "Closing Sauda"
                                Case "V"
                                    ExBrokRec!broktype = "Value Wise Intraday"
                                Case "Q"
                                    ExBrokRec!broktype = "Qtywise Intraday"
                                Case "D"
                                    ExBrokRec!broktype = "Delivery Wise"
                                Case "H"
                                    ExBrokRec!broktype = "Higher Value Wise Intraday"
                                Case "W"
                                    ExBrokRec!broktype = "WHigher Value Wise Intraday"
                                Case "X"
                                    ExBrokRec!broktype = "XIntraDay Wise"
                                Case "Z"
                                    ExBrokRec!broktype = "Zlotwise"
                                End Select
                                ExBrokRec!TRANTYPE = "Percentage Wise"
                                If TRec!MARTYPE = "Q" Then
                                    ExBrokRec!MARTYPE = "Qtywise (Per Unit)"
                                ElseIf TRec!MARTYPE = "I" Then
                                    ExBrokRec!MARTYPE = "Import Rates"
                                ElseIf TRec!MARTYPE = "C" Then
                                    ExBrokRec!MARTYPE = "Client Wise"
                                Else
                                    ExBrokRec!MARTYPE = "Value Wise (In %)"
                                End If
                                ExBrokRec.Update
                                TRec.MoveNext
                            Loop
                            Set TRec = Nothing
                            Set DataGrid2.DataSource = ExBrokRec: DataGrid2.ReBind: DataGrid2.Refresh
                        End If
                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
                        mysql = "SELECT PB.BROKRATE,PB.STDRATE,PB.BROKTYPE,PB.TRANTYPE,PB.TRANRATE,PB.MARTYPE,PB.MARRATE,PB.UPTOSTDT,IT.ITEMCODE,IT.ITEMNAME "
                        mysql = mysql & " FROM PITBROK AS PB,ITEMMAST AS IT WHERE PB.COMPCODE =" & GCompCode & " AND PB.COMPCODE=IT.COMPCODE AND PB.ITEMCODE=IT.ITEMCODE "
                        mysql = mysql & " AND PB.AC_CODE='" & LAcCode & "' ORDER BY IT.ITEMCODE,PB.UPTOSTDT"
                        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                        If Not TRec.EOF Then
                            Do While Not TRec.EOF
                                RECGRID.AddNew
                                RECGRID!ITEMCODE = TRec!ITEMCODE & vbNullString
                                RECGRID!ITEMName = TRec!ITEMName & vbNullString
                                RECGRID!brokrate = Val(TRec!brokrate & vbNullString)
                                RECGRID!STDRATE = Val(TRec!STDRATE & vbNullString)
                                RECGRID!TRANRATE = Val(TRec!TRANRATE & vbNullString)
                                Select Case TRec!broktype
                                Case ""
                                    RECGRID!broktype = "Transaction Wise"
                                Case "T"
                                    RECGRID!broktype = "Transaction Wise"
                                Case "O"
                                    RECGRID!broktype = "Opening Sauda"
                                Case "P"
                                    RECGRID!broktype = "Percentage Wise"
                                Case "I"
                                    RECGRID!broktype = "Intraday Wise"
                                Case "C"
                                    RECGRID!broktype = "Closing Sauda"
                                Case "V"
                                    RECGRID!broktype = "Value Wise Intraday"
                                Case "Q"
                                    RECGRID!broktype = "Qtywise Intraday"
                                Case "D"
                                    RECGRID!broktype = "Delivery Wise"
                                Case "H"
                                    RECGRID!broktype = "Higher Value Wise Intraday"
                                Case "W"
                                    RECGRID!broktype = "WHigher Value Wise Intraday"
                                Case "X"
                                    RECGRID!broktype = "XIntraDay Wise"
                                Case "Z"
                                    RECGRID!broktype = "Zlotwise"
                                End Select
                                RECGRID!TRANTYPE = "Percentgae Wise"
                                If TRec!MARTYPE = "Q" Then
                                    RECGRID!MARGINTYPE = "Qtywise (Per Unit)"
                                ElseIf TRec!MARTYPE = "I" Then
                                    RECGRID!MARGINTYPE = "Import Rates"
                                ElseIf TRec!MARTYPE = "C" Then
                                    RECGRID!MARGINTYPE = "Client Wise"
                                Else
                                    RECGRID!MARGINTYPE = "Value Wise (In %)"
                                End If
                                RECGRID!MarginRate = Val(TRec!MARRATE & vbNullString)
                                If IsNull(TRec!UPTOSTDT) Then
                                    RECGRID!UPTOSTDT = DateValue(LSettlementDt)
                                Else
                                    If TRec!UPTOSTDT = "" Then
                                        RECGRID!UPTOSTDT = DateValue(LSettlementDt)
                                    Else
                                        RECGRID!UPTOSTDT = TRec!UPTOSTDT
                                    End If
                                End If
                                RECGRID.Update
                                TRec.MoveNext
                            Loop
                            Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
                        End If
                        'exchange code *************
                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
                        TRec.Open "SELECT EXCODE,EXID  FROM EXMAST WHERE COMPCODE=" & GCompCode & " ", Cnn, adOpenForwardOnly, adLockReadOnly      '' left outer join acct_ex on EXMAST.excode=acct_ex.excode AND EXMAST.AC_CODE='" & LACCode & "' AND EXMAST.COMPCODE=" & GCompCode  & "  "
                        If Not TRec.EOF Then
                            Call ExRECSET
                            Do While Not TRec.EOF
                                ExRecGrid.AddNew
                                ExRecGrid!excode = TRec!excode
                                ExRecGrid!EXID = TRec!EXID
                                ExRecGrid!ACEXCODE = vbNullString
                                ExRecGrid!BILLINGTYPE = "Settle Rate"
                                ExRecGrid.Update
                                TRec.MoveNext
                            Loop
                            Set TRec = Nothing: Set TRec = New ADODB.Recordset
                            TRec.Open "SELECT ACEXCODE,BILLBY,EXCODE FROM ACCT_EX WHERE COMPCODE = " & GCompCode & " AND AC_CODE='" & LAcCode & "'", Cnn, adOpenForwardOnly, adLockReadOnly
                            If TRec.EOF Then
                                Set TRec = Nothing: Set TRec = New ADODB.Recordset
                                TRec.Open "SELECT EXCODE FROM EXMAST WHERE COMPCODE = " & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
                                While Not TRec.EOF
                                    ExRecGrid.MoveFirst
                                    ExRecGrid.Find "EXCODE = '" & TRec!excode & "'", , adSearchForward
                                    If Not ExRecGrid.EOF Then
                                        ExRecGrid!ACEXCODE = vbNullString
                                        ExRecGrid.Update
                                    End If
                                    TRec.MoveNext
                                Wend
                            Else
                                While Not TRec.EOF
                                    ExRecGrid.MoveFirst
                                    ExRecGrid.Find "EXCODE = '" & TRec!excode & "'", , adSearchForward
                                    If Not ExRecGrid.EOF Then
                                        ExRecGrid!ACEXCODE = TRec!ACEXCODE & ""
                                        If TRec!billby = "S" Then
                                            ExRecGrid!BILLINGTYPE = "Settle Rate"
                                        ElseIf TRec!billby = "B" Then
                                            ExRecGrid!BILLINGTYPE = "Bid Ask Rate"
                                        ElseIf TRec!billby = "P" Then
                                            ExRecGrid!BILLINGTYPE = "Party Own Rate"
                                        End If
                                        ExRecGrid.Update
                                    End If
                                    TRec.MoveNext
                                Wend
                            End If
                            Set ExGrid.DataSource = ExRecGrid: ExGrid.ReBind: ExGrid.Refresh
                        End If
                    End If
                End If
                ACCSETUP.TabVisible(1) = True: ACCSETUP.TabVisible(2) = True:   ACCSETUP.TabVisible(3) = True:                ACCSETUP.TabVisible(4) = True
                If (MG_Code = 36 Or MG_Code = 37) Then
                    ACCSETUP.TabVisible(2) = False
                    ACCSETUP.TabVisible(3) = False
                    ACCSETUP.TabVisible(4) = False
                End If
            Else
                ACCSETUP.TabVisible(1) = False: ACCSETUP.TabVisible(2) = False: ACCSETUP.TabVisible(3) = False: ACCSETUP.TabVisible(4) = False
            End If
        End If
        'Check Branch Exists Or Not
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT COMPCODE FROM ACCFMLY WHERE COMPCODE = " & GCompCode & " AND FMLYHEAD = '" & LAC_CODE & "' ", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            Check2.Visible = False
        Else
            Check2.Visible = True
        End If
        
        If Fb_Press = 3 Then
            Call Delete_Record
            Exit Sub
        End If
        ACCSETUP.Enabled = True
        ACCSETUP.Tab = 0: TxtAcName.SetFocus
    Else
        MsgBox "Please Select account", vbInformation
        Call CANCEL_RECORD
    End If
    Exit Sub
Error1: If err.Number <> 0 Then
            MsgBox err.Description
            Dim Ret As Byte
            If Ret = 5 Then
                Resume Next
            ElseIf Ret = 4 Then
            Else
                Call Get_Selection(10)
                GETMAIN.ProgressBar1.Value = 0
                GETMAIN.ProgressBar1.Visible = False
                Unload Me
            End If
        Else
            Exit Sub
        End If
End Sub
Private Sub DIFF_LOSTFOCUS()
    If (MG_Code = 12 Or MG_Code = 13 Or MG_Code = 14) Then
        ACCSETUP.Tab = 1
        TxtAdd.SetFocus
    End If
End Sub
Sub List_Rec()
    Screen.MousePointer = 11
    Frame24.Visible = False
    DataGrid3.Visible = False
    Call Get_Selection(12)
    mysql = "SELECT DISTINCT GR.G_NAME,A.NAME,AD.CITY,A.OP_BAL,A.ACTIVE,A.AC_CODE,AD.ADD1,AD.PIN,AD.PANNO,AX.ACEXCODE,AD.PHONEO,AD.PHONER,AD.MOBILE,AD.EMAIL"
    mysql = mysql & " FROM AC_GROUP AS GR, ACCOUNTM AS A, ACCOUNTD AS AD, ACCT_EX AS AX "
    mysql = mysql & "Where AD.COMPCODE = A.COMPCODE  AND AD.AC_CODE=A.AC_CODE AND AX.COMPCODE =AD.COMPCODE AND AD.AC_CODE =AX.AC_CODE AND A.COMPCODE=" & GCompCode & " AND GR.CODE=A.GCODE ORDER BY GR.G_NAME, A.NAME"
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If GETMAIN.PRINTTOGGLE.Checked = True Then
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "ACLIST_W.RPT", 1)
    Else
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "ACLIST_D.RPT", 1)
    End If
    Frame4.Visible = False
    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource TRec
    RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'" & GCompanyName & "'"
    CRViewer1.Width = CInt(GETMAIN.Width - 100)
    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Visible = True
    CRViewer1.ReportSource = RDCREPO
    CRViewer1.ViewReport
    Screen.MousePointer = 0
End Sub
Sub RecSet()
    Set ExBrokRec = Nothing
    Set ExBrokRec = New ADODB.Recordset
    ExBrokRec.Fields.Append "EXCODE", adVarChar, "10", adFldIsNullable
    ExBrokRec.Fields.Append "INSTTYPE", adVarChar, "3", adFldIsNullable
    ExBrokRec.Fields.Append "AC_CODE", adVarChar, "15", adFldIsNullable
    ExBrokRec.Fields.Append "BROKTYPE", adVarChar, "50", adFldIsNullable
    ExBrokRec.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "MINRATE", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "MBROKTYPE", adVarChar, "50", adFldIsNullable
    ExBrokRec.Fields.Append "MBROKRATE", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "MBROKRATE2", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "TRANTYPE", adVarChar, "50", adFldIsNullable
    ExBrokRec.Fields.Append "TRANRATE", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "STDRATE", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "MARTYPE", adVarChar, "50", adFldIsNullable
    ExBrokRec.Fields.Append "MARRATE", adDouble, , adFldIsNullable
    ExBrokRec.Fields.Append "UPTOSTDT", adVarChar, "10", adFldIsNullable
    ExBrokRec.Open , , adOpenKeyset, adLockOptimistic
    
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
    RECGRID.Fields.Append "ITEMNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "STDRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "TRANRATE", adDouble, , adFldIsNullable
    
    RECGRID.Fields.Append "TRANTYPE", adVarChar, 50, adFldIsNullable
    
    RECGRID.Fields.Append "PARTYTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BCYCLE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "MARGINTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "MARGINRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "UptoStdt", adVarChar, 10, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub
Sub ExRECSET()
    Set ExRecGrid = Nothing
    Set ExRecGrid = New ADODB.Recordset
    ExRecGrid.Fields.Append "EXCODE", adVarChar, 10, adFldIsNullable
    ExRecGrid.Fields.Append "ACEXCODE", adVarChar, 30, adFldIsNullable
    ExRecGrid.Fields.Append "BillingType", adVarChar, 50, adFldIsNullable
    ExRecGrid.Fields.Append "EXID", adInteger, , adFldIsNullable
    
    ExRecGrid.Open , , adOpenKeyset, adLockOptimistic
End Sub
Private Sub TxtFax_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub txt19_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub TxtIntRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtIntRate_LostFocus()
    TxtIntRate.text = Format(Val(TxtIntRate.text), "0.00")
End Sub
Private Sub TxtIntRate_Validate(Cancel As Boolean)
    TxtIntRate.text = Format(Val(TxtIntRate.text), "0.0000")
End Sub

Private Sub CmdAdd_Click()
Dim TRec  As ADODB.Recordset
Dim LMaxId As Integer
LFilePress = 1
Set TRec = Nothing
Set TRec = New ADODB.Recordset
mysql = "SELECT MAX(SNO ) AS MID FROM PARTYMULTI WHERE COMPCODE =" & GCompCode & ""
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

TxtSno.text = LMaxId:                   ExCombo.BoundText = vbNullString
ItemCombo.BoundText = vbNullString:     TxtMultiplier.text = vbNullString
CmdAdd.Enabled = False:                 CmdMod.Enabled = False
CmdSave.Enabled = True:                 TxtSno.Locked = True
ExCombo.SetFocus

End Sub
Private Sub CmdMod_Click()
LFilePress = 2
TxtSno.Locked = False
TxtSno.text = vbNullString
ExCombo.BoundText = vbNullString
ItemCombo.BoundText = vbNullString
TxtMultiplier.text = vbNullString
CmdAdd.Enabled = False
CmdMod.Enabled = False
CmdSave.Enabled = True
TxtSno.SetFocus
End Sub

Private Sub CmdSave_Click()
Dim LBrokerType As String:      Dim LPartyAs As String
Dim TRec   As ADODB.Recordset:  Dim LExID As Integer
Dim LItemID As Integer:         Dim LACCID As Long

If LenB(ExCombo.BoundText) < 1 Then
    MsgBox "Please Select Exchange"
    ExCombo.SetFocus
    Exit Sub
End If

If IsNull(ItemCombo.BoundText) Then ItemCombo.BoundText = vbNullString
If LFilePress = 2 Then
    mysql = "DELETE FROM PARTYMULTI WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & TxtAcCode.text & "'"
    mysql = mysql & " AND EXCODE='" & ExCombo.BoundText & "' "
    mysql = mysql & " AND ITEMCODE ='" & ItemCombo.BoundText & "'"
    mysql = mysql & " AND SNO =" & Val(TxtSno.text) & ""
    
    Cnn.Execute mysql
Else
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT SNO FROM PARTYMULTI WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & TxtAcCode.text & "'"
    mysql = mysql & " AND EXCODE='" & ExCombo.BoundText & "'"
    mysql = mysql & " AND ITEMCODE ='" & ItemCombo.BoundText & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        MsgBox "Entry Already Exist For " & (ExCombo.BoundText & vbNullString) & " For " & (ItemCombo.BoundText & vbNullString) & ""
        Exit Sub
        ExCombo.SetFocus
    End If
End If
LExID = Get_ExID(ExCombo.BoundText)
LItemID = Get_ITEMID(ItemCombo.text)
LACCID = Get_AccID(TxtAcCode.text)
If Val(TxtMultiplier.text & vbNullString) <> 0 Then
    mysql = "INSERT INTO PARTYMULTI (COMPCODE,SNO,PARTY,EXCODE,ITEMCODE,RATE,EXID,ITEMID,ACCID)"
    mysql = mysql & " VALUES (" & GCompCode & "," & Val(TxtSno.text) & ",'" & TxtAcCode.text & "','" & ExCombo.BoundText & "','" & ItemCombo.BoundText & "'," & Val(TxtMultiplier.text) & "," & LExID & "," & LItemID & "," & LACCID & " )"
    Cnn.Execute mysql
End If
CmdAdd.Enabled = True
CmdMod.Enabled = True
CmdSave.Enabled = False
TxtSno.text = vbNullString
ExCombo.BoundText = vbNullString
ItemCombo.BoundText = vbNullString
TxtMultiplier.text = vbNullString
Call Refresh_MultiGrid
CmdAdd.SetFocus

End Sub

Private Sub Refresh_MultiGrid()
    Dim TRec As ADODB.Recordset
    Set RecMulti = Nothing
    Set RecMulti = New ADODB.Recordset
    RecMulti.Fields.Append "SNo", adInteger, , adFldIsNullable
    RecMulti.Fields.Append "ExCode", adVarChar, 10, adFldIsNullable
    RecMulti.Fields.Append "ItemCode", adVarChar, 20, adFldIsNullable
    RecMulti.Fields.Append "Multiplier", adDouble, , adFldIsNullable
    RecMulti.Open , , adOpenKeyset, adLockBatchOptimistic
    mysql = "SELECT SNO,EXCODE ,ITEMCODE  ,RATE FROM PARTYMULTI  "
    mysql = mysql & " WHERE COMPCODE =" & GCompCode & " "
    mysql = mysql & " AND PARTY  ='" & TxtAcCode.text & "' ORDER BY SNO "
    Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            RecMulti.AddNew
            RecMulti!sno = TRec!sno
            RecMulti!excode = TRec!excode
            RecMulti!ITEMCODE = (TRec!ITEMCODE & vbNullString)
            RecMulti!MULTIPLIER = TRec!Rate
            RecMulti.Update
            TRec.MoveNext
        Loop
    Else
        RecMulti.AddNew
        RecMulti!sno = 1
        RecMulti!excode = vbNullString
        RecMulti!ITEMCODE = vbNullString
        RecMulti!MULTIPLIER = 0
        RecMulti.Update
    End If
    Set Multigrid.DataSource = RecMulti
    Multigrid.ReBind: DataGrid1.Refresh
    Multigrid.Columns(0).Width = 900:
    Multigrid.Columns(1).Width = 1500:
    Multigrid.Columns(2).Width = 2000:
    Multigrid.Columns(3).Width = 1000:
    Multigrid.Columns(3).Alignment = dbgRight
    Multigrid.Columns(3).NumberFormat = "0.00"
End Sub


Private Sub FillDataGrid()
    mysql = "SELECT A.ACCID,A.Ac_Code,A.NAME,B.G_NAME AS GroupName,A.OP_BAL,A.GCODE,A.GRPCODE,A.ACTIVE,a.PTYHEAD FROM ACCOUNTM AS A , AC_GROUP AS B WHERE A.COMPCODE =" & GCompCode & " AND A.GCODE=B.CODE "
    If LenB(TxtFilterName.text) <> 0 Then mysql = mysql & " AND UPPER(A.NAME) LIKE '" & Trim(UCase(TxtFilterName.text)) & "%' "
    If LenB(TxtFilterCode.text) <> 0 Then mysql = mysql & " AND A.AC_CODE  LIKE '" & Trim(UCase(TxtFilterCode.text)) & "%' "
    mysql = mysql & " ORDER BY A.NAME"
    Set AccountMRec = Nothing:            Set AccountMRec = New ADODB.Recordset
    AccountMRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If AccountMRec.EOF Then
        'MsgBox "No Records Found"
        TxtFilterName.text = vbNullString
        'TxtFilterName.SetFocus
    End If
    If Not AccountMRec.EOF Then
        Set DataGrid3.DataSource = AccountMRec
        DataGrid3.ReBind
        DataGrid3.Refresh
        DataGrid3.Columns(0).Width = 900:
        DataGrid3.Columns(1).Width = 1200:
        DataGrid3.Columns(2).Width = 5000:
        DataGrid3.Columns(3).Width = 2500
        DataGrid3.Columns(4).Width = 2500
        DataGrid3.Columns(5).Visible = False
        DataGrid3.Columns(6).Visible = False
        DataGrid3.Columns(7).Visible = False
        DataGrid3.Columns(8).Visible = False
        DataGrid3.Columns(4).Alignment = dbgRight
        DataGrid3.Columns(4).NumberFormat = "0.00"
        DataGrid3.Refresh
    End If
End Sub

Private Sub DataGrid3_Click()
If AccountMRec.RecordCount > 0 Then
    If AccountMRec.EOF Then AccountMRec.MoveFirst
    DataGrid3.Col = 1
    TxtAcCode.text = DataGrid3.text
    DataGrid3.Col = 2
    TxtAcName.text = DataGrid3.text
    
    If InStr(1, Text5.text, TxtAcCode.text) <= 0 Then
        Text5.text = Text5.text + "," + TxtAcCode.text
    Else
        Text5.text = Replace(Text5.text, "," + TxtAcCode.text, "")
    End If
End If
End Sub
Private Sub DataGrid3_DblClick()
    Call Get_Selection(2)
    Fb_Press = 2
    Call ACCOUNT_ACCESS
End Sub
Private Sub DataGrid3_KeyPress(KeyAscii As Integer)
    Dim LChar As String
    LChar = UCase(Chr(KeyAscii))
    If KeyAscii = 13 Then
        Call Get_Selection(2)
        Fb_Press = 2
        Call ACCOUNT_ACCESS
    Else
        AccountMRec.MoveFirst
        Do While Not AccountMRec.EOF
            If Left$(AccountMRec!NAME, 1) <> LChar Then
                AccountMRec.MoveNext
            Else
                Exit Do
            End If
        Loop
        If AccountMRec.EOF Then AccountMRec.MoveFirst
    End If
End Sub
Private Sub Command1_Click()
Dim TRec As ADODB.Recordset
Dim LOldCode As String
Dim LNewCode As String
Dim LNewName As String
If LenB(TxtOldCode.text) > 0 Then
    LOldCode = Get_AccountMCode(TxtOldCode.text)
    If LenB(LOldCode) < 1 Then
        MsgBox "Invalid Old Code  "
        TxtOldCode.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Old Code can not be Blank "
    TxtOldCode.SetFocus
    Exit Sub
End If

If LenB(TxtNewCode.text) > 0 Then
    LNewCode = Get_AccountMCode(TxtNewCode.text)
    If LenB(LNewCode) > 0 Then
        MsgBox "Invalid New Code, New Old Already Exists   "
        TxtNewCode.text = vbNullString
        TxtNewCode.SetFocus
        Exit Sub
    End If
Else
    MsgBox "New Code can not be Blank "
    TxtNewCode.SetFocus
    Exit Sub
End If

If LenB(TxtoldName.text) < 1 Then
    MsgBox "Old name Blank "
    TxtOldCode.SetFocus
    Exit Sub
End If
If LenB(TxtnewName.text) > 0 Then
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT AC_CODE FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND NAME ='" & TxtnewName.text & "' AND AC_CODE<>'" & TxtOldCode.text & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        MsgBox "Invalid New Name, New Name Already Assigned to " & TRec!AC_CODE & ""
        TxtnewName.text = vbNullString
        TxtnewName.SetFocus
        Exit Sub
    End If
Else
    MsgBox "New Name can not be Blank "
    TxtnewName.SetFocus
    Exit Sub
End If

If Check1.Value = 1 And Len(LNewCode) > 0 Then
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT AC_CODE FROM ACCT_EX WHERE COMPCODE =" & GCompCode & " AND ACEXCODE = '" & LNewCode & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        MsgBox "Invalid Exchnage Code , New Code Already Assigned to " & TRec!AC_CODE & ""
        TxtNewCode.text = vbNullString
        TxtNewCode.SetFocus
        Exit Sub
    End If
End If
    
    Cnn.BeginTrans
    CNNERR = True
    mysql = "UPDATE ACCOUNTM SET AC_CODE ='" & TxtNewCode.text & "',NAME ='" & TxtnewName.text & "' "
    mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE ACCOUNTD SET AC_CODE ='" & TxtNewCode.text & "',NAME ='" & TxtnewName.text & "' "
    mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    
    mysql = "UPDATE ACCT_EX SET AC_CODE ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND AC_CODE  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    
    If Check1.Value = 1 Then
        mysql = "UPDATE ACCT_EX SET ACEXCODE ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND AC_CODE  ='" & TxtNewCode.text & "'"
        Cnn.Execute mysql
    End If
    mysql = "UPDATE ACCFMLY SET FMLYHEAD  ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND FMLYHEAD ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE ACCFMLY SET CONTRA_AC  ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND CONTRA_AC  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE ACCFMLYD SET PARTY ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE PEXBROK SET AC_CODE ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE PEXSBROK SET PARTY ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE PITBROK SET AC_CODE ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND AC_CODE  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE PITSBROK SET PARTY ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE CTR_D SET PARTY ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE CTR_D SET CLCODE ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND CLCODE  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE CTR_D SET CONCODE ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND CONCODE  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE INV_D SET PARTY ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE INV_D1 SET PARTY ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE VCHAMT  SET AC_CODE  ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE DMARGIN SET PARTY ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE EXBROKCLIENT SET CLIENT  ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND CLIENT  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    mysql = "UPDATE PARTYMULTI SET PARTY   ='" & TxtNewCode.text & "' WHERE COMPCODE =" & GCompCode & " AND PARTY  ='" & TxtOldCode.text & "'"
    Cnn.Execute mysql
    
    MsgBox "Account Code Changed Succesfully"
    Cnn.CommitTrans
    DataGrid3.Refresh
    
End Sub
Sub ImportIssues(LMsg As String)

    Dim iloop As Integer
    
    For iloop = 0 To List1.ListCount
        If LMsg = List1.List(iloop) Then
            LMsg = ""
            Exit For
        End If
    Next
    If LMsg <> "" Then
        List1.AddItem (LMsg)
    End If
    
End Sub
