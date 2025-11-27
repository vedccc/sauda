VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmExSBrok 
   Caption         =   "Exchange Wise Sub Brokerage and Sharing"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
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
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   1695
      Left            =   5760
      TabIndex        =   40
      Top             =   5640
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1080
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label label7 
         Caption         =   "Formula"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13080
      TabIndex        =   29
      Top             =   2640
      Width           =   4815
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply"
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
         Left            =   3240
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox ChkUpdLastSettle 
         BackColor       =   &H00FFFF80&
         Caption         =   "Upd Last Settlement"
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
         Left            =   2040
         TabIndex        =   31
         Top             =   120
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox ChkBrokLock 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Lock Brokerage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   495
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   43531.5349074074
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "Upto Date "
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
         Left            =   120
         TabIndex        =   33
         Top             =   555
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
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
      Left            =   12960
      TabIndex        =   26
      Top             =   3720
      Width           =   4935
      Begin VB.OptionButton OptExchange 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Exchange Wise"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton OptItem 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Item Wise"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   27
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   12960
      TabIndex        =   20
      Top             =   960
      Width           =   5055
      Begin VB.CheckBox ChkUpdBrok 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Update Client Brok Also "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   39
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1575
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
         ItemData        =   "FrmExSBrok.frx":0000
         Left            =   1080
         List            =   "FrmExSBrok.frx":000D
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo UptoDateCombo 
         Height          =   420
         Left            =   3240
         TabIndex        =   23
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "To Lock Brokerage first Select Above Date"
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
         TabIndex        =   34
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "B Date"
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
         Height          =   285
         Left            =   2520
         TabIndex        =   24
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Inst Type"
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
         TabIndex        =   21
         Top             =   195
         Width           =   855
      End
   End
   Begin VB.ComboBox ShareTypeCombo 
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
      ItemData        =   "FrmExSBrok.frx":0029
      Left            =   18480
      List            =   "FrmExSBrok.frx":0036
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox BrokTypeCombo 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "FrmExSBrok.frx":005E
      Left            =   18480
      List            =   "FrmExSBrok.frx":0095
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   240
      TabIndex        =   14
      Top             =   4320
      Width           =   17775
      Begin TabDlg.SSTab SSTab1 
         Height          =   5055
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   17535
         _ExtentX        =   30930
         _ExtentY        =   8916
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         BackColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Exchange Wise Sub Brokerage"
         TabPicture(0)   =   "FrmExSBrok.frx":01E5
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ExGrid"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "ItemWise Sub Brokerage"
         TabPicture(1)   =   "FrmExSBrok.frx":0201
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ItemGrid"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid ExGrid 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   37
            Top             =   720
            Width           =   17295
            _ExtentX        =   30506
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            BackColor       =   -2147483628
            ForeColor       =   4194368
            HeadLines       =   1
            RowHeight       =   23
            TabAction       =   1
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   14
            BeginProperty Column00 
               DataField       =   "FMLYNAME"
               Caption         =   "Branch Name"
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
               DataField       =   "Party"
               Caption         =   "Code"
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
               DataField       =   "PartyName"
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
               DataField       =   "EXCODE"
               Caption         =   "ExCode"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
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
            BeginProperty Column05 
               DataField       =   "BROKRATE"
               Caption         =   "Brok Rate"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "BBROKTYPE"
               Caption         =   "Sub BrokType"
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
               DataField       =   "BBROKRATE"
               Caption         =   "SubBrokRate"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "BrokRate2"
               Caption         =   "BrokRate2"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "UPTOSTDT"
               Caption         =   "Set. Date"
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
            BeginProperty Column10 
               DataField       =   "Share"
               Caption         =   "Share(%)"
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
               DataField       =   "Applyon"
               Caption         =   "Apply On"
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
               DataField       =   "New"
               Caption         =   "New"
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
               DataField       =   "FMLYCODE"
               Caption         =   "Branch"
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
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column10 
                  Alignment       =   1
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   2520
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid ItemGrid 
            Height          =   4335
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   17160
            _ExtentX        =   30268
            _ExtentY        =   7646
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            BackColor       =   -2147483628
            ForeColor       =   4194368
            HeadLines       =   1
            RowHeight       =   23
            TabAction       =   1
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   16
            BeginProperty Column00 
               DataField       =   "FMLYNAME"
               Caption         =   "BranchName"
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
               DataField       =   "Party"
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
               DataField       =   "PartyName"
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
               DataField       =   "ITEMCODE"
               Caption         =   "Item Code"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
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
            BeginProperty Column05 
               DataField       =   "BROKRATE"
               Caption         =   "Brok Rate"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "BBROKTYPE"
               Caption         =   "Sub BrokType"
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
               DataField       =   "BBROKRATE"
               Caption         =   "SubBrokRate"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "BROKRATE2"
               Caption         =   "BrokRate2"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "UPTOSTDT"
               Caption         =   "Set. Date"
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
            BeginProperty Column10 
               DataField       =   "Share"
               Caption         =   "Share(%)"
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
               DataField       =   "Applyon"
               Caption         =   "Apply On"
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
               DataField       =   "New"
               Caption         =   "New"
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
               DataField       =   "FMLYCODE"
               Caption         =   "Branch"
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
               DataField       =   "EXCODE"
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
            BeginProperty Column15 
               DataField       =   "DeleteRow"
               Caption         =   "DeleteRow"
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
                  ColumnWidth     =   1604.976
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column10 
                  Alignment       =   1
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   2520
               EndProperty
               BeginProperty Column14 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   1005.165
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000040&
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
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18255
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
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   18015
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sub Brokerage Sharing Setup "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   17895
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
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
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   17775
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   4
         Top             =   120
         Width           =   12615
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Select All"
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
            Left            =   11160
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Select All"
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
            Left            =   5760
            TabIndex        =   7
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Select All"
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
            Left            =   8160
            TabIndex        =   6
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Select All"
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
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            Top             =   120
            Width           =   1215
         End
         Begin MSComctlLib.ListView PartyLst 
            Height          =   2700
            Left            =   3600
            TabIndex        =   8
            Top             =   480
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   4763
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   4194304
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Party Name"
               Object.Width           =   6350
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ACCID"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView ExList 
            Height          =   2700
            Left            =   7200
            TabIndex        =   9
            Top             =   480
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   4763
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   4194304
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   5185
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Exchange Name"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView BranchList 
            Height          =   2700
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   4763
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   4194304
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Branch Name"
               Object.Width           =   5185
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "FMLYCODE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "FMLYID"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView ItemList 
            Height          =   2700
            Left            =   9600
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   4763
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   4194304
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   5185
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ExCode"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ItemID"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
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
            Height          =   360
            Left            =   7200
            TabIndex        =   17
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Party List"
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
            Left            =   3600
            TabIndex        =   16
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Branch List"
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
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Item List"
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
            Height          =   360
            Left            =   9600
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   9060
      Left            =   120
      Top             =   720
      Width           =   18075
   End
End
Attribute VB_Name = "FrmExSBrok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LExCodes As String:             Dim LFmlyIDs  As String:       Dim LParties As String:             Dim LInstType As String
Dim Items As String:                Dim ListIt As ListItem:         Dim GridColVal As String:           Dim CountRow As Double
Dim SearchRow As Double:            Public FlagBrok As Boolean:     Dim LSettlementDt As Date:          Dim RECGRID As ADODB.Recordset
Dim TempRec As ADODB.Recordset:     Dim flag As Boolean:            Public LDataCol As Integer:         Public Fb_Press As Byte
Dim AddMode As Boolean:             Dim RecAcc As ADODB.Recordset:  Dim LSItems As String:              Dim ItemRecGrid As ADODB.Recordset
Dim UptoDateRec As ADODB.Recordset
Dim AllExcodes As Boolean
Dim LActive_Grid As Integer
Dim GGridRow As Integer
Dim GGridCol As Integer


Sub ADD_NEW()
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
    Frame4.Enabled = True
    
    Call Get_Selection(1)
    BranchList.SetFocus
End Sub
Sub CANCEL_REC()
    Dim I As Integer
    Fb_Press = 0
    
    Frame3.Enabled = True
    Frame2.Enabled = True
    CmdOk.Enabled = True:   UptoDateCombo.Enabled = True
    ChkUpdLastSettle.Visible = False
    Label9.Visible = False
    vcDTP1.Visible = False
    CmdApply.Visible = False
    CmdOk.Enabled = True
    PartyLst.ListItems.Clear
    For I = 1 To BranchList.ListItems.Count
        BranchList.ListItems.Item(I).Checked = False
    Next I
    For I = 1 To ExList.ListItems.Count
        ExList.ListItems.Item(I).Checked = False
    Next I
    Check1.Value = 0:    Check4.Value = 0:    Check5.Value = 0
    vcDTP1.Value = Date
    ChkBrokLock.Value = 0
    Call RecSet
    Call ItemRecSet
    UptoDateCombo.BoundText = vbNullString
    Frame9.Enabled = True
    Set ExGrid.DataSource = RECGRID: ExGrid.ReBind: ExGrid.Refresh: ExGrid.Enabled = False
    Set ItemGrid.DataSource = ItemRecGrid: ItemGrid.ReBind: ItemGrid.Refresh: ItemGrid.Enabled = False
    CmdOk.Enabled = True:  BrokTypeCombo.Visible = False
    SSTab1.Tab = 0
    Frame2.Enabled = False: Frame3.Enabled = False
    Frame8.Enabled = False
    Frame9.Enabled = False
    Frame4.Enabled = False
    Call Get_Selection(13)
End Sub
Sub Save_Rec()
    On Error GoTo err1
    Frame3.Enabled = False
    Frame2.Enabled = False
    Frame1.Enabled = False
    
    CmdOk.Enabled = False
    UptoDateCombo.Enabled = False
    Dim LastStDate, LStdTDate As Date
    mysql = "DELETE FROM PEXSBROK WHERE UPTOSTDT IS NULL"
    Cnn.Execute mysql
    mysql = "DELETE FROM PITSBROK WHERE UPTOSTDT IS NULL"
    Cnn.Execute mysql
    If InstCombo.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf InstCombo.ListIndex = 1 Then
        LInstType = "OPT"
    Else
        LInstType = "CSH"
    End If
    If OptExchange.Value Then
        Save_ExBrok
    Else
        Save_ItemBrok
    End If
    Cnn.BeginTrans: CNNERR = True
    If ChkUpdBrok.Value = 1 Then
        Call Update_Charges(LParties, LExCodes, vbNullString, vbNullString, GFinBegin, GFinEnd, False)
    End If
    'Call Delete_Inv_D(LParties, LExCodes, vbNullString, GFinBegin)
    If BILL_GENERATION(GFinBegin, GFinEnd, vbNullString, LParties, LExCodes) Then
        Cnn.CommitTrans: CNNERR = False
    Else
        Cnn.RollbackTrans: CNNERR = False
    End If
    'Call Chk_Billing
    Call CANCEL_REC
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
        Cnn.RollbackTrans
        CNNERR = False
    End If
End Sub
Private Sub Check1_Click()
    Dim I As Integer
    For I = 1 To PartyLst.ListItems.Count
        If Check1.Value = 1 Then
            PartyLst.ListItems.Item(I).Checked = True
        Else
            PartyLst.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub Check2_Click()
'    If Check2.Value = 1 Then
'        MYSQL = "SELECT DISTINCT ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC, CTR_D AS CT ,ACCFMLY AS AF WHERE ACC.COMPCODE=" & GCompCode & " AND ACC.COMPCODE = CT.COMPCODE AND ACC.COMPCODE = AF.COMPCODE AND CT.USERID = AF.FMLYCODE AND AF.FMLYHEAD = ACC.AC_CODE ORDER BY ACC.NAME"
'    Else
'        MYSQL = "SELECT DISTINCT ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC, PEXBROK AS PB WHERE ACC.COMPCODE=" & GCompCode & " AND ACC.COMPCODE = PB.COMPCODE AND ACC.AC_CODE = PB.AC_CODE ORDER BY ACC.NAME"
'    End If
'    Set RecAcc = Nothing
'    Set RecAcc = New ADODB.Recordset
'    RecAcc.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'    If Not RecAcc.EOF Then
'        PartyLst.ListItems.Clear
'        While Not RecAcc.EOF
'            Set ListIt = PartyLst.ListItems.Add(, , RecAcc!NAME)
'            ListIt.SubItems(1) = RecAcc!AC_CODE
'            ListIt.SubItems(1) = RecAcc!ACCID
'            RecAcc.MoveNext
'        Wend
'    End If
End Sub
Private Sub Check4_Click()
    Dim I As Integer
    For I = 1 To BranchList.ListItems.Count
        If Check4.Value = 1 Then
            BranchList.ListItems.Item(I).Checked = True
        Else
            BranchList.ListItems.Item(I).Checked = False
        End If
    Next I
    Call BranchList_Click
End Sub
Private Sub Check5_Click()
    Dim I As Integer
For I = 1 To ExList.ListItems.Count
    If Check5.Value = 1 Then
        ExList.ListItems.Item(I).Checked = True
    Else
        ExList.ListItems.Item(I).Checked = False
    End If
Next I
Call Get_ExCodes
Call Fill_ItemList
End Sub

Private Sub ChkBrokLock_Click()
If ChkBrokLock.Value = 1 Then
    vcDTP1.Visible = True
    ChkUpdLastSettle.Visible = True
    Label9.Visible = True
    CmdApply.Visible = True
    
Else
    vcDTP1.Visible = False
    ChkUpdLastSettle.Visible = False
    Label9.Visible = False
    CmdApply.Visible = False
End If
End Sub

Private Sub Command1_Click()

End Sub


Private Sub ShareTypeCombo_GotFocus()
    If LActive_Grid = 1 Then
        If Mid(RECGRID!APPLYON, 1, 1) = "N" Then
            ShareTypeCombo.ListIndex = 0
        ElseIf Mid(RECGRID!APPLYON, 1, 1) = "G" Then
            ShareTypeCombo.ListIndex = 1
        ElseIf Mid(RECGRID!APPLYON, 1, 1) = "S" Then
            ShareTypeCombo.ListIndex = 2
        End If
        ShareTypeCombo.Top = 4800 + Val(ExGrid.Top) + Val(ExGrid.RowTop(ExGrid.Row))
        ShareTypeCombo.Width = Val(ExGrid.Columns(ExGrid.Col).Width)
        ShareTypeCombo.Left = 360 + Val(ExGrid.Left) + Val(ExGrid.Columns(ExGrid.Col).Left)
    Else
        If Mid(ItemRecGrid!APPLYON, 1, 1) = "N" Then
            ShareTypeCombo.ListIndex = 0
        ElseIf Mid(ItemRecGrid!APPLYON, 1, 1) = "G" Then
            ShareTypeCombo.ListIndex = 1
        ElseIf Mid(ItemRecGrid!APPLYON, 1, 1) = "S" Then
            ShareTypeCombo.ListIndex = 2
        End If
        ShareTypeCombo.Top = 4800 + Val(ItemGrid.Top) + Val(ItemGrid.RowTop(ItemGrid.Row))
        ShareTypeCombo.Width = Val(ItemGrid.Columns(ExGrid.Col).Width)
        ShareTypeCombo.Left = 360 + Val(ItemGrid.Left) + Val(ItemGrid.Columns(ItemGrid.Col).Left)
    End If
        Sendkeys "%{DOWN}"
End Sub
Private Sub ShareTypeCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If LActive_Grid = 1 Then
            RECGRID!APPLYON = ShareTypeCombo.text
            ExGrid.Col = 11
            ShareTypeCombo.Visible = False: ExGrid.SetFocus
        Else
            ItemRecGrid!APPLYON = ShareTypeCombo.text
            ItemGrid.Col = 11
            ShareTypeCombo.Visible = False: ItemGrid.SetFocus
        End If
        
    ElseIf KeyCode = 27 Then
        ShareTypeCombo.Visible = False
    End If
End Sub
Private Sub ShareTypeCombo_Validate(Cancel As Boolean)
    If Len(Trim(ShareTypeCombo.text)) < 1 Then Cancel = True: Exit Sub
End Sub
Private Sub InstCombo_Validate(Cancel As Boolean)
If InstCombo.ListIndex < 0 Then
    MsgBox "Please Select Instrument Type "
    Cancel = True
End If
End Sub
Public Sub CmdOk_Click()
    Dim TRec As ADODB.Recordset
    Dim MDefBrokType As String
    Dim J As Integer
    
    Frame3.Enabled = False
    LParties = vbNullString
    For J = 1 To PartyLst.ListItems.Count
        If PartyLst.ListItems(J).Checked = True Then
            If LenB(LParties) > 0 Then LParties = LParties & ", "
            LParties = LParties & "'" & PartyLst.ListItems(J).SubItems(1) & "'"
        End If
    Next
    LInstType = vbNullString
    If InstCombo.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf InstCombo.ListIndex = 1 Then
        LInstType = "OPT"
    ElseIf InstCombo.ListIndex = 2 Then
        LInstType = "CSH"
    End If
    
    If LenB(LParties) < 1 Then
        Frame3.Enabled = True
        MsgBox "Please Select Party.", vbCritical:
        
        PartyLst.SetFocus:
        Exit Sub
    End If
    'Call Check_PExSBrok
    LFmlyIDs = vbNullString
    For J = 1 To BranchList.ListItems.Count
        If BranchList.ListItems(J).Checked = True Then
            If LenB(LFmlyIDs) > 0 Then LFmlyIDs = LFmlyIDs & ", "
            LFmlyIDs = LFmlyIDs & "'" & BranchList.ListItems(J).SubItems(2) & "'"
        End If
    Next
    LExCodes = vbNullString
    For J = 1 To ExList.ListItems.Count
        If ExList.ListItems(J).Checked = True Then
            If LenB(LExCodes) > 0 Then LExCodes = LExCodes & ", "
            LExCodes = LExCodes & "" & ExList.ListItems(J).SubItems(2) & ""
        End If
    Next
    If LenB(LExCodes) < 1 Then
        Frame3.Enabled = True
        MsgBox "Please Select Exchange ", vbCritical:
        ExList.SetFocus::
        Exit Sub
    End If
    LSItems = vbNullString
    If OptItem = True Then
        LSItems = vbNullString
        For J = 1 To ItemList.ListItems.Count
            If ItemList.ListItems(J).Checked = True Then
                If LenB(LSItems) > 0 Then LSItems = LSItems & ", "
                LSItems = LSItems & ItemList.ListItems(J).SubItems(2) & ""
            End If
        Next
        If LenB(LSItems) < 1 Then
            Frame3.Enabled = True
            MsgBox "Please Select Items  ", vbCritical:
            ItemList.SetFocus::
            Exit Sub
        End If
    End If
    Call Fill_ExchangeGrid
    Call Fill_ItemGrid
    If LenB(UptoDateCombo.BoundText) > 1 Then ChkBrokLock.Visible = True
    Frame1.Enabled = True
    CmdOk.Enabled = False:
    If RECGRID.RecordCount > 0 Then RECGRID.MoveFirst:
    
    'ExGrid.SetFocus
    ExGrid.LeftCol = 0
    Label3.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim scriptcontrol As New scriptcontrol
scriptcontrol.Language = "vbscript"
    If KeyAscii = 13 Then
        If Text1.text <> "" Then
            Dim Lformula As String
            Dim Lbbrokrate As Double
            Dim LSHARE As Double
            Dim Lvar As Variant
            Dim mno1 As Double
            
            If LActive_Grid = 1 Then
                RECGRID.MoveFirst
                While Not RECGRID.EOF
                    Lformula = UCase(Text1.text)
                    If InStr(1, Lformula, "BROKRATE") > 0 Then
                        Lformula = Replace(Lformula, "BROKRATE", RECGRID!brokrate)
                    End If
                    If InStr(1, Lformula, "SHARE") > 1 Then
                        Lformula = Replace(Lformula, "SHARE", RECGRID!SHARE)
                    End If
                    Lvar = Lformula
                    RECGRID!bbrokrate = Val(scriptcontrol.Eval(Lvar))
                    RECGRID.MoveNext
                Wend
                Set ExGrid.DataSource = RECGRID: ExGrid.ReBind: ExGrid.Refresh
                ExGrid.Row = 1: ExGrid.Col = 1: ExGrid.SetFocus
            Else
                ItemRecGrid.MoveFirst
                While Not ItemRecGrid.EOF
                    Lformula = UCase(Text1.text)
                    If InStr(1, Lformula, "BROKRATE") > 0 Then
                        Lformula = Replace(Lformula, "BROKRATE", ItemRecGrid!brokrate)
                    End If
                    If InStr(1, Lformula, "SHARE") > 1 Then
                        Lformula = Replace(Lformula, "SHARE", ItemRecGrid!SHARE)
                    End If
                    Lvar = Lformula
                    ItemRecGrid!bbrokrate = Val(scriptcontrol.Eval(Lvar))
                    ItemRecGrid.MoveNext
                Wend
                Set ItemGrid.DataSource = ItemRecGrid: ItemGrid.ReBind: ItemGrid.Refresh
                ItemGrid.Row = 1: ItemGrid.Col = 1: ItemGrid.SetFocus
            End If
            Frame6.Visible = False
        End If
    End If
End Sub

Private Sub UptoDateCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub UptoDateCombo_Validate(Cancel As Boolean)
    If IsDate(UptoDateCombo.text) Then
        If SYSTEMLOCK(DateValue(UptoDateCombo.text)) Then
            MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
            Cancel = True
        End If
    End If
End Sub


Private Sub ExList_Click()
    Call Get_ExCodes
    Call Fill_ItemList
End Sub

Private Sub Form_Load()
    Dim TRec As ADODB.Recordset
    FlagBrok = False
    'Last Settlement Date
    InstCombo.ListIndex = 0
    SSTab1.Tab = 0
    Text1.text = GEmail
    
'    Set TRec = Nothing
'    Set TRec = New ADODB.Recordset
'    mysql = "SELECT DISTINCT INSTTYPE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & ""
'    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Not TRec.EOF Then
'        If TRec.RecordCount = 1 Then
'            If TRec!INSTTYPE = "FUT" Then
'                InstCombo.Locked = True
'            End If
'        End If
'    End If
    
    LSettlementDt = GFinEnd: Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT MAX(SETDATE) AS MAXSETTLEDATE FROM SETTLE WHERE COMPCODE = " & GCompCode & "", Cnn, adOpenKeyset, adLockReadOnly
    If Not TRec.EOF Then LSettlementDt = TRec!MaxSettleDate
    Set RecAcc = Nothing
    Set RecAcc = New ADODB.Recordset
    mysql = "SELECT DISTINCT PB.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC,PEXBROK AS PB WHERE ACC.COMPCODE =" & GCompCode & " AND ACC.COMPCODE =PB.COMPCODE AND ACC.AC_CODE =PB.AC_CODE ORDER BY ACC.NAME"
    RecAcc.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    ExGrid.Enabled = False
    If Not RecAcc.EOF Then
        'While Not RecAcc.EOF
        '    Set ListIt = PartyLst.ListItems.Add(, , RecAcc!NAME)
        '    ListIt.SubItems(1) = RecAcc!AC_CODE
        '    RecAcc.MoveNext
        'Wend
        Call Get_Selection(13)
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT FMLYID,FMLYCODE,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " ORDER BY FMLYNAME ", Cnn, adOpenKeyset, adLockReadOnly
        If Not TRec.EOF Then
            While Not TRec.EOF
                Set ListIt = BranchList.ListItems.Add(, , TRec!FmlyNAME)
                ListIt.SubItems(1) = TRec!FMLYCODE
                ListIt.SubItems(2) = TRec!FMLYID
                TRec.MoveNext
            Wend
        End If
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        mysql = "SELECT EXID,EXCODE,EXNAME FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE "
        TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not TRec.EOF Then
            TRec.MoveFirst
            ExList.ListItems.Clear
            ExList.Enabled = True: Check4.Enabled = True
            Do While Not TRec.EOF
                If (TRec!excode = "EQ" Or TRec!excode = "BEQ") Then
                    InstCombo.Visible = True
                End If
                ExList.ListItems.Add , , TRec!excode
                ExList.ListItems(ExList.ListItems.Count).ListSubItems.Add , , TRec!EXNAME
                ExList.ListItems(ExList.ListItems.Count).ListSubItems.Add , , TRec!EXID
                TRec.MoveNext
            Loop
            
            If TRec.RecordCount = 1 Then
                Check4.Value = 1: ExList.TabStop = False: Check4.TabStop = False
                Call Check4_Click
            Else
                ExList.TabStop = True: Check4.TabStop = True
            End If
        Else
            ExList.Enabled = False: Check4.Enabled = False
        End If
        Set UptoDateRec = Nothing: Set UptoDateRec = New ADODB.Recordset
        mysql = "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PEXSBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT"
        UptoDateRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not UptoDateRec.EOF Then
            Set UptoDateCombo.RowSource = UptoDateRec
            UptoDateCombo.ListField = "CONDATE"
            UptoDateCombo.BoundColumn = "CONDATE"
        End If
    End If
    Call CANCEL_REC
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "PARTY", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "PARTYNAME", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
    RECGRID.Fields.Append "ITEMNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BBROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BBROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "UPTOSTDT", adDate, , adFldIsNullable
    RECGRID.Fields.Append "APPLYON", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SHARE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "New", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "FMLYCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "FMLYNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "EXCODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "EXID", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "FMLYID", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "ACCID", adInteger, , adFldIsNullable
    
    
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub
Sub ItemRecSet()
    Set ItemRecGrid = Nothing
    Set ItemRecGrid = New ADODB.Recordset
    ItemRecGrid.Fields.Append "PARTY", adVarChar, 15, adFldIsNullable
    ItemRecGrid.Fields.Append "PARTYNAME", adVarChar, 100, adFldIsNullable
    ItemRecGrid.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
    ItemRecGrid.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    ItemRecGrid.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    ItemRecGrid.Fields.Append "BBROKTYPE", adVarChar, 50, adFldIsNullable
    ItemRecGrid.Fields.Append "BBROKRATE", adDouble, , adFldIsNullable
    ItemRecGrid.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    ItemRecGrid.Fields.Append "UPTOSTDT", adDate, , adFldIsNullable
    ItemRecGrid.Fields.Append "APPLYON", adVarChar, 50, adFldIsNullable
    ItemRecGrid.Fields.Append "SHARE", adDouble, , adFldIsNullable
    ItemRecGrid.Fields.Append "New", adDouble, , adFldIsNullable
    ItemRecGrid.Fields.Append "FMLYCODE", adVarChar, 6, adFldIsNullable
    ItemRecGrid.Fields.Append "FMLYNAME", adVarChar, 50, adFldIsNullable
    ItemRecGrid.Fields.Append "EXCODE", adVarChar, 50, adFldIsNullable
    ItemRecGrid.Fields.Append "DeleteRow", adVarChar, 1, adFldIsNullable
    ItemRecGrid.Fields.Append "EXID", adInteger, , adFldIsNullable
    ItemRecGrid.Fields.Append "ITEMID", adInteger, , adFldIsNullable
    ItemRecGrid.Fields.Append "FMLYID", adInteger, , adFldIsNullable
    ItemRecGrid.Fields.Append "ACCID", adInteger, , adFldIsNullable
    
    ItemRecGrid.Open , , adOpenKeyset, adLockOptimistic
End Sub
Private Sub ExGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim LGridRow As Integer
Dim LGridCol As Integer
    LActive_Grid = 1
    If KeyCode = 13 And ExGrid.Col = 6 Then ' BROKTYPE
        BrokTypeCombo.Visible = True: BrokTypeCombo.SetFocus
    ElseIf KeyCode = 13 And ExGrid.Col = 4 And ChkUpdBrok.Value = 1 Then ' BROKTYPE
        BrokTypeCombo.Visible = True: BrokTypeCombo.SetFocus
    ElseIf KeyCode = 13 And ExGrid.Col = 11 Then 'APPLYON
        ShareTypeCombo.Visible = True: ShareTypeCombo.SetFocus
    ElseIf KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
    'Press F7 to replace all rows with current cell value
0    If KeyCode = 118 Then   'F7
        LGridRow = ExGrid.Row
        LGridCol = ExGrid.Col
        If ExGrid.Col = 4 And ChkUpdBrok.Value = 1 Then 'BROKTYPE
            GridColVal = RECGRID!broktype
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!broktype = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf ExGrid.Col = 5 And ChkUpdBrok.Value = 1 Then 'BROKRATE
            GridColVal = RECGRID!brokrate
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!brokrate = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf ExGrid.Col = 6 Then 'BBROKTYPE
            GridColVal = RECGRID!bbroktype
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!bbroktype = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf ExGrid.Col = 7 Then 'BBROKRATE
            GridColVal = RECGRID!bbrokrate
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!bbrokrate = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf ExGrid.Col = 8 Then 'BBROKRATE 2
            GridColVal = RECGRID!BROKRATE2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BROKRATE2 = GridColVal
                RECGRID.MoveNext
            Wend
        'ElseIf ExGrid.Col = 8 Then 'UPTOSTDT
        '    GridColVal = RECGRID!UPTOSTDT
        '    RECGRID.MoveFirst
        '    While Not RECGRID.EOF
         '       RECGRID!UPTOSTDT = GridColVal
          '      RECGRID.MoveNext
           ' Wend
        ElseIf ExGrid.Col = 10 Then 'SHARE
            GridColVal = RECGRID!SHARE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!SHARE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf ExGrid.Col = 11 Then 'APPLYON
            GridColVal = RECGRID!APPLYON
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!APPLYON = GridColVal
                RECGRID.MoveNext
            Wend
        End If
        Set ExGrid.DataSource = RECGRID: ExGrid.ReBind: ExGrid.Refresh
        ExGrid.Row = LGridRow: ExGrid.Col = LGridCol: ExGrid.SetFocus
    ElseIf KeyCode = 119 Then   'F8
        Frame6.Visible = True
    End If
End Sub
Private Sub BrokTypeCombo_GotFocus()
Dim LGridRow As Integer
Dim LGridCol As Integer
    
    If LActive_Grid = 1 Then
        If ExGrid.Col = 6 Then
            If Mid(RECGRID!broktype, 1, 1) = "T" Then
                BrokTypeCombo.ListIndex = 0
            ElseIf Mid(RECGRID!broktype, 1, 1) = "P" Then
                BrokTypeCombo.ListIndex = 1
            ElseIf Mid(RECGRID!broktype, 1, 1) = "O" Then
                BrokTypeCombo.ListIndex = 2
            ElseIf Mid(RECGRID!broktype, 1, 1) = "I" Then
                BrokTypeCombo.ListIndex = 3
            ElseIf Mid(RECGRID!broktype, 1, 1) = "Z" Then
                BrokTypeCombo.ListIndex = 4
            ElseIf Mid(RECGRID!broktype, 1, 1) = "R" Then
                BrokTypeCombo.ListIndex = 5
            ElseIf Mid(RECGRID!broktype, 1, 1) = "X" Then
                BrokTypeCombo.ListIndex = 6
            ElseIf Mid(RECGRID!broktype, 1, 1) = "Q" Then
                BrokTypeCombo.ListIndex = 7
            ElseIf Mid(RECGRID!broktype, 1, 1) = "C" Then
                BrokTypeCombo.ListIndex = 8
            ElseIf Mid(RECGRID!broktype, 1, 1) = "H" Then
                BrokTypeCombo.ListIndex = 9
            ElseIf Mid(RECGRID!broktype, 1, 1) = "W" Then
                BrokTypeCombo.ListIndex = 10
            ElseIf Mid(RECGRID!broktype, 1, 1) = "S" Then
                BrokTypeCombo.ListIndex = 11
            ElseIf Mid(RECGRID!broktype, 1, 1) = "V" Then
                BrokTypeCombo.ListIndex = 13
            ElseIf Mid(RECGRID!broktype, 1, 1) = "3" Then
                BrokTypeCombo.ListIndex = 14
            ElseIf Mid(RECGRID!broktype, 1, 1) = "A" Then
                BrokTypeCombo.ListIndex = 15
            ElseIf Mid(RECGRID!broktype, 1, 1) = "5" Then
                BrokTypeCombo.ListIndex = 16
            End If
        End If
        BrokTypeCombo.Top = 4800 + Val(ExGrid.Top) + Val(ExGrid.RowTop(ExGrid.Row))
        BrokTypeCombo.Width = Val(ExGrid.Columns(ExGrid.Col).Width)
        BrokTypeCombo.Left = 360 + Val(ExGrid.Left) + Val(ExGrid.Columns(ExGrid.Col).Left)
    Else
        If ItemGrid.Col = 6 Then
            If Mid(ItemRecGrid!broktype, 1, 1) = "T" Then
                BrokTypeCombo.ListIndex = 0
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "P" Then
                BrokTypeCombo.ListIndex = 1
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "O" Then
                BrokTypeCombo.ListIndex = 2
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "I" Then
                BrokTypeCombo.ListIndex = 3
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "Z" Then
                BrokTypeCombo.ListIndex = 4
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "R" Then
                BrokTypeCombo.ListIndex = 5
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "X" Then
                BrokTypeCombo.ListIndex = Val(6)
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "Q" Then
                BrokTypeCombo.ListIndex = 7
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "C" Then
                BrokTypeCombo.ListIndex = 8
            ElseIf Mid(RECGRID!broktype, 1, 1) = "H" Then
                BrokTypeCombo.ListIndex = 9
            ElseIf Mid(RECGRID!broktype, 1, 1) = "W" Then
                BrokTypeCombo.ListIndex = 10
            ElseIf Mid(RECGRID!broktype, 1, 1) = "S" Then
                BrokTypeCombo.ListIndex = 11
            ElseIf Mid(RECGRID!broktype, 1, 1) = "V" Then
                BrokTypeCombo.ListIndex = 13
            ElseIf Mid(RECGRID!broktype, 1, 1) = "3" Then
                BrokTypeCombo.ListIndex = 14
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "A" Then
                BrokTypeCombo.ListIndex = 15
            ElseIf Mid(ItemRecGrid!broktype, 1, 1) = "5" Then
                BrokTypeCombo.ListIndex = 16
            End If
        End If
        BrokTypeCombo.Top = 4800 + Val(ItemGrid.Top) + Val(ItemGrid.RowTop(ItemGrid.Row))
        BrokTypeCombo.Width = Val(ItemGrid.Columns(ItemGrid.Col).Width)
        BrokTypeCombo.Left = 360 + Val(ItemGrid.Left) + Val(ItemGrid.Columns(ItemGrid.Col).Left)
    End If
    Sendkeys "%{DOWN}"
End Sub
Private Sub BrokTypeCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LGridRow As Integer
Dim LGridCol As Integer

    If KeyCode = 13 Then
        If LActive_Grid = 1 Then
            LGridRow = ExGrid.Row: LGridCol = ExGrid.Col:
            If ExGrid.Col = 6 Then RECGRID!bbroktype = BrokTypeCombo.text
            If ExGrid.Col = 4 Then RECGRID!broktype = BrokTypeCombo.text
            Set ExGrid.DataSource = RECGRID: ExGrid.ReBind: ExGrid.Refresh:   RECGRID.MoveFirst: ExGrid.SetFocus
            If LGridCol = 6 Then
                ExGrid.Row = LGridRow: ExGrid.Col = LGridCol + 1: BrokTypeCombo.Visible = False: ExGrid.SetFocus
            Else
                ExGrid.Row = LGridRow: ExGrid.Col = LGridCol + 1: BrokTypeCombo.Visible = False: ExGrid.SetFocus
            End If
        Else
            LGridRow = ItemGrid.Row: LGridCol = ItemGrid.Col:
            If LGridCol = 6 Then ItemRecGrid!bbroktype = BrokTypeCombo.text
            If LGridCol = 4 Then ItemRecGrid!broktype = BrokTypeCombo.text
            Set ItemGrid.DataSource = ItemRecGrid: ItemGrid.ReBind: ItemGrid.Refresh:  ItemRecGrid.MoveFirst: ItemGrid.SetFocus
            If LGridCol = 6 Then
                ItemGrid.Row = LGridRow: ItemGrid.Col = LGridCol + 1: BrokTypeCombo.Visible = False: ItemGrid.SetFocus
            Else
                ItemGrid.Row = LGridRow: ItemGrid.Col = LGridCol + 1: BrokTypeCombo.Visible = False: ItemGrid.SetFocus
            End If
        End If
    ElseIf KeyCode = 27 Then
        BrokTypeCombo.Visible = False
    End If
End Sub
Private Sub BrokTypeCombo_Validate(Cancel As Boolean)
    If Len(Trim(BrokTypeCombo.text)) < 1 Then
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub BranchList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        Call BranchList_Click
    End If
End Sub
Private Sub BranchList_Click()
    Dim I As Integer
    Dim RecSauda As ADODB.Recordset
    LFmlyIDs = vbNullString
    For I = 1 To BranchList.ListItems.Count
        If BranchList.ListItems(I).Checked = True Then
            If LenB(LFmlyIDs) > 0 Then LFmlyIDs = LFmlyIDs & ", "
            LFmlyIDs = LFmlyIDs & BranchList.ListItems(I).ListSubItems(2) & ""
        End If
  Next I
  PartyLst.ListItems.Clear
  If LFmlyIDs = "" Then Me.MousePointer = 0: Exit Sub
    mysql = "SELECT ACCID,AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND ACCID IN (SELECT DISTINCT ACCID FROM ACCFMLYD WHERE  FMLYID IN  (" & LFmlyIDs & "))  ORDER BY NAME "
    Set RecSauda = Nothing: Set RecSauda = New ADODB.Recordset: RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    While Not RecSauda.EOF
        Set ListIt = PartyLst.ListItems.Add(, , RecSauda!NAME)
        ListIt.SubItems(1) = RecSauda!AC_CODE
        ListIt.SubItems(2) = RecSauda!ACCID
        RecSauda.MoveNext
    Wend
End Sub
Private Sub Fill_ExchangeGrid()
    Dim BrokRec As ADODB.Recordset:     Dim PExBrokRec As ADODB.Recordset
    Dim LBrokType As String:            Dim LShType As String
    Dim LBrokRate As Double
    Call RecSet
    Set ExGrid.DataSource = RECGRID: ExGrid.ReBind: ExGrid.Refresh
    Set BrokRec = Nothing: Set BrokRec = New ADODB.Recordset
    mysql = "SELECT AM.NAME,AM.ACCID,A.PARTY,A.EXCODE,A.BROKTYPE AS BBROKTYPE,A.BROKRATE AS BBROKRATE,A.BROKRATE2,A.SHTYPE,A.SHRATE,"
    mysql = mysql & " A.FMLYID ,A.FMLYCODE,FM.FMLYNAME,A.UPTOSTDT,A.EXID "
    mysql = mysql & " FROM PEXSBROK AS A, ACCOUNTD AS AM,ACCFMLY FM"
    mysql = mysql & " WHERE A.COMPCODE=" & GCompCode & "  "
    mysql = mysql & " AND A.PARTY IN  (" & LParties & ") "
    mysql = mysql & " AND A.INSTTYPE='" & LInstType & "' "
    mysql = mysql & " AND A.ACCID =AM.ACCID "
    mysql = mysql & " AND FM.FMLYID = A.FMLYID  "
    If LenB(LFmlyIDs) > 0 Then mysql = mysql & " AND A.FMLYID   IN (" & LFmlyIDs & ")  "
    If LenB(LExCodes) > 0 Then mysql = mysql & "AND A.EXID  IN (" & LExCodes & ")  "
    If IsDate(UptoDateCombo.text) Then mysql = mysql & " AND A.UPTOSTDT = '" & Format(UptoDateCombo.text, "yyyy/MM/dd") & "'  "
    mysql = mysql & " ORDER BY AM.NAME,A.EXCODE,A.UPTOSTDT "
    BrokRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not BrokRec.EOF Then
        ExGrid.Enabled = True
        Do While Not BrokRec.EOF
            DoEvents
            RECGRID.AddNew
            RECGRID.Fields("EXCODE") = BrokRec!excode
            RECGRID!EXID = BrokRec!EXID
            RECGRID.Fields("PARTY") = BrokRec!PARTY
            RECGRID.Fields("PARTYNAME") = BrokRec!NAME
            
            Set PExBrokRec = Nothing
            Set PExBrokRec = New ADODB.Recordset
            mysql = "SELECT EXCODE,AC_CODE,BROKTYPE,BROKRATE,UPTOSTDT FROM PEXBROK WHERE INSTTYPE ='" & LInstType & "'"
            mysql = mysql & " AND EXID  =" & BrokRec!EXID & " AND ACCID =" & BrokRec!ACCID & ""
            mysql = mysql & " AND UPTOSTDT> = '" & Format(BrokRec!UPTOSTDT, "yyyy/MM/dd") & "'  "
            mysql = mysql & " ORDER BY UPTOSTDT "
            PExBrokRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            If Not PExBrokRec.EOF Then
                If IsNull(PExBrokRec!broktype) Or PExBrokRec!broktype = "" Then
                    LBrokType = "Transaction"
                Else
                    LBrokType = PExBrokRec!broktype
                End If
                 LBrokRate = IIf(IsNull(PExBrokRec!brokrate), 0, PExBrokRec!brokrate)
            Else
                LBrokRate = 0
                LBrokType = "T"
            End If
            RECGRID.Fields("BROKRATE") = LBrokRate
            Select Case LBrokType
                Case "B"
                    RECGRID.Fields("BROKTYPE") = "BuySell Intraday"
                Case "C"
                    RECGRID.Fields("BROKTYPE") = "Closing Sauda"
                Case "3"
                    RECGRID.Fields("BROKTYPE") = "3 Closing Sauda %"
                Case "D"
                    RECGRID.Fields("BROKTYPE") = "Delivery Wise Brokerage"
                Case "F"
                    RECGRID.Fields("BROKTYPE") = "Fixed Brokerage"
                Case "H"
                    RECGRID.Fields("BROKTYPE") = "Higher Value Percentage Wise"
                Case "I"
                    RECGRID.Fields("BROKTYPE") = "IntraDay Brokerage"
                Case "L"
                    RECGRID.Fields("BROKTYPE") = "LotWise Higher Value "
                Case "M"
                    RECGRID.Fields("BROKTYPE") = "MRate Wise IntraDay"
                Case "N"
                    RECGRID.Fields("BROKTYPE") = "N Per Trade Wise"
                Case "O"
                    RECGRID.Fields("BROKTYPE") = "Opening Sauda"
                Case "P"
                    RECGRID.Fields("BROKTYPE") = "Percentage wise"
                Case "Q"
                    RECGRID.Fields("BROKTYPE") = "Qtywise IntraDay"
                Case "R"
                    RECGRID.Fields("BROKTYPE") = "RZLotwise Intraday"
                Case "S"
                    RECGRID.Fields("BROKTYPE") = "Slab Wise Brokerage"
                Case "T"
                    RECGRID.Fields("BROKTYPE") = "Transaction"
                Case "U"
                    RECGRID.Fields("BROKTYPE") = "U ShareQty Wise"
                Case "V"
                    RECGRID.Fields("BROKTYPE") = "Valuewise Intraday"
                Case "X"
                    RECGRID.Fields("BROKTYPE") = "XIntraday Higher Wise"
                Case "Y"
                    RECGRID.Fields("BROKTYPE") = "Y Qtywise Intraday"
                Case "Z"
                    RECGRID.Fields("BROKTYPE") = "ZLotwise"
                Case "A"
                    RECGRID.Fields("BROKTYPE") = "A Opening ZLotwise"
                Case "5"
                    RECGRID.Fields("BROKTYPE") = "5 Closing Sauda ZLotwise"
            End Select
            'RECGRID.Fields("BBROKTYPE") = IIf(IsNull(BrokRec!BBROKTYPE), "T", BrokRec!BBROKTYPE)
            LBrokType = IIf(IsNull(BrokRec!bbroktype), "T", BrokRec!bbroktype)
            Select Case LBrokType
                Case "A"
                    RECGRID.Fields("BROKTYPE") = "A Opening ZLotwise"
                Case "5"
                    RECGRID.Fields("BROKTYPE") = "5 Closing Sauda Zlotwise"
                Case "D"
                    RECGRID.Fields("BBROKTYPE") = "Delivery Wise Brokerage"
                Case "C"
                    RECGRID.Fields("BBROKTYPE") = "Closing Sauda"
                Case "3"
                    RECGRID.Fields("BBROKTYPE") = "3 Closing Sauda %"
                Case "H"
                    RECGRID.Fields("BBROKTYPE") = "Higher Value Percentage Wise"
                Case "I"
                    RECGRID.Fields("BBROKTYPE") = "IntraDay Brokerage"
                Case "O"
                    RECGRID.Fields("BBROKTYPE") = "Opening Sauda"
                Case "P"
                    RECGRID.Fields("BBROKTYPE") = "Percentage wise"
                Case "Q"
                    RECGRID.Fields("BBROKTYPE") = "Qtywise IntraDay"
                Case "R"
                    RECGRID.Fields("BBROKTYPE") = "RZLotwise Intraday"
                Case "S"
                    RECGRID.Fields("BBROKTYPE") = "Sub Brokerage in %"
                Case "T"
                    RECGRID.Fields("BBROKTYPE") = "Transaction"
                Case "V"
                    RECGRID.Fields("BBROKTYPE") = "Valuewise Intraday"
                Case "X"
                    RECGRID.Fields("BBROKTYPE") = "XIntraday Higher Wise"
                Case "Z"
                    RECGRID.Fields("BBROKTYPE") = "ZLotwise"
            End Select
            RECGRID.Fields("BBROKRATE") = IIf(IsNull(BrokRec!bbrokrate), 0, BrokRec!bbrokrate)
            RECGRID.Fields("BROKRATE2") = IIf(IsNull(BrokRec!BROKRATE2), 0, BrokRec!BROKRATE2)
            LShType = IIf(IsNull(BrokRec!SHTYPE), "G", BrokRec!SHTYPE)
            If LShType = "G" Then
                RECGRID.Fields("APPLYON") = "Gross Amount"
            ElseIf LShType = "S" Then
                RECGRID.Fields("APPLYON") = "ShareNet"
            Else
                RECGRID.Fields("APPLYON") = "Net Amount"
            End If
            RECGRID.Fields("SHARE") = IIf(IsNull(BrokRec!SHRATE), 0, BrokRec!SHRATE)
            RECGRID.Fields("FMLYCODE") = IIf(IsNull(BrokRec!FMLYCODE), "G", BrokRec!FMLYCODE)
            RECGRID.Fields("FMLYNAME") = IIf(IsNull(BrokRec!FmlyNAME), 0, BrokRec!FmlyNAME)
            RECGRID.Fields("FMLYID") = BrokRec!FMLYID
            RECGRID.Fields("ACCID") = BrokRec!ACCID
            
            If IsNull(BrokRec!UPTOSTDT) Then
                RECGRID.Fields("UPTOSTDT") = Format(LSettlementDt, "YYYY/MM/DD")
            Else
                RECGRID.Fields("UPTOSTDT") = Format(BrokRec!UPTOSTDT, "DD/MM/YYYY")
            End If
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            RECGRID.Update
            BrokRec.MoveNext
        Loop
        Set ExGrid.DataSource = RECGRID: ExGrid.ReBind: ExGrid.Refresh: RECGRID.MoveFirst:
        
        'ExGrid.SetFocus
        ExGrid.LeftCol = 0
        Label3.Visible = True
    Else
        Label3.Visible = False
        
        MsgBox "Record does not exists.", vbExclamation
        Call CANCEL_REC
    End If
End Sub
Private Sub Fill_ItemGrid()
    Dim BrokRec As ADODB.Recordset:     Dim PItBrokRec As ADODB.Recordset
    Dim LFmlyRec As ADODB.Recordset:    Dim PartyRec As ADODB.Recordset
    Dim PExBrokRec As ADODB.Recordset:  Dim LBrokRate As Double
    Dim TRec As ADODB.Recordset:        Dim LBrokType As String
    Dim LShType As String
    
    Call ItemRecSet
    Call Get_Items
    LSettlementDt = GFinEnd
    If LenB(LSItems) > 1 Or OptItem.Value = False Then
        Set ItemGrid.DataSource = ItemRecGrid: ItemGrid.ReBind: ItemGrid.Refresh
        Set BrokRec = Nothing: Set BrokRec = New ADODB.Recordset
        If Len(LSItems) < 1 Then
            mysql = "SELECT A.ITEMID,A.EXID,AM.NAME,A.ACCID,A.PARTY,A.EXCODE,A.ITEMCODE,A.BROKTYPE AS BBROKTYPE,A.BROKRATE AS BBROKRATE,A.BROKRATE2,A.SHTYPE,A.SHRATE,"
            mysql = mysql & " A.FMLYID,A.FMLYCODE,FM.FMLYNAME,A.UPTOSTDT FROM PITSBROK AS A, ACCOUNTD AS AM,ACCFMLY FM WHERE A.COMPCODE=" & GCompCode & "  "
            mysql = mysql & " AND A.ACCID =AM.ACCID  "
            mysql = mysql & " AND A.PARTY IN  (" & LParties & ") "
            mysql = mysql & " AND A.FMLYID  IN  (" & LFmlyIDs & ") "
            mysql = mysql & " AND A.INSTTYPE='" & LInstType & "' "
            mysql = mysql & " AND FM.FMLYID  = A.FMLYID   "
            If LenB(LExCodes) > 0 Then mysql = mysql & " AND A.EXID  IN (" & LExCodes & ")  "
            If LenB(LSItems) > 0 Then mysql = mysql & " AND A.ITEMID  IN (" & LSItems & ")  "
            If IsDate(UptoDateCombo.text) Then mysql = mysql & " AND A.UPTOSTDT = '" & Format(UptoDateCombo.text, "yyyy/MM/dd") & "'  "
            mysql = mysql & " ORDER BY AM.NAME,A.EXCODE,A.ITEMCODE,A.UPTOSTDT "
        Else
            'mysql = "SELECT AM.ACCID,AM.NAME,AM.AC_CODE,A.tRANTYPE,A.BROKTYPE,A.BROKRATE,A.BROKRATE2,A.STDRATE,A.TRANRATE, A.UPTOSTDT , A.MARTYPE,A.MARRATE,A.MINRATE,"
            'mysql = mysql & " A.MBROKRATE , A.MBROKRATE2, A.MBROKTYPE, B.EXCHANGECODE, B.ITEMCODE, A.INSTTYPE, B.ITEMID, B.EXID "
            'mysql = mysql & " FROM ACCOUNTD AM INNER JOIN ITEMMAST B ON AM.COMPCODE=B.COMPCODE "
            'mysql = mysql & " LEFT OUTER JOIN PITBROK AS A ON B.ITEMID=A.ITEMID  AND AM.ACCID = A.ACCID "
            'mysql = mysql & " WHERE AM.COMPCODE= " & GCompCode & " and B.EXID IN (" & LExCodes & ") and AM.AC_CODE IN (" & LSParties & ")"
            'If LenB(LSItems) > 0 Then mysql = mysql & " AND B.ITEMID IN (" & LSItems & ")"
            'If ChkBrokLock.Value = 0 Then
            '    If IsDate(SettleDCombo.text) Then mysql = mysql & " AND A.UPTOSTDT = '" & Format(SettleDCombo.text, "yyyy/MM/dd") & "'"
            'End If
            'mysql = mysql & " ORDER BY AM.NAME,B.EXCHANGECODE, A.ITEMCODE,A.UPTOSTDT "
            'FROM PITSBROK AS A, ACCOUNTD AS AM,ACCFMLY FM WHERE A.COMPCODE=" & GCompCode & "
            mysql = "SELECT AM.ACCID,AM.NAME,AM.AC_CODE AS PARTY,B.ITEMID,B.EXID,B.EXCHANGECODE AS EXCODE,B.ITEMCODE,A.BROKTYPE AS BBROKTYPE,A.BROKRATE AS BBROKRATE,A.BROKRATE2,A.SHTYPE,A.SHRATE,"
            mysql = mysql & " FM.FMLYID,FM.FMLYCODE,FM.FMLYNAME,A.UPTOSTDT "
            mysql = mysql & " FROM ACCOUNTD AM INNER JOIN ITEMMAST B ON AM.COMPCODE=B.COMPCODE "
            If LenB(LExCodes) > 0 Then mysql = mysql & " AND B.EXID  IN (" & LExCodes & ")  "
            If LenB(LSItems) > 0 Then mysql = mysql & " AND B.ITEMID  IN (" & LSItems & ")  "
            mysql = mysql & " AND AM.AC_CODE IN  (" & LParties & ") "
            mysql = mysql & " LEFT OUTER JOIN ACCFMLY AS FM ON  "
            mysql = mysql & " FM.FMLYID  IN  (" & LFmlyIDs & ") "
            mysql = mysql & " LEFT OUTER JOIN PITSBROK AS A ON B.ITEMID=A.ITEMID  AND AM.ACCID = A.ACCID "
            mysql = mysql & " AND A.INSTTYPE='" & LInstType & "' "
            mysql = mysql & " AND FM.FMLYID  = A.FMLYID   "

            If IsDate(UptoDateCombo.text) Then mysql = mysql & " AND A.UPTOSTDT = '" & Format(UptoDateCombo.text, "yyyy/MM/dd") & "'  "
            mysql = mysql & " ORDER BY AM.NAME,A.EXCODE,A.ITEMCODE,A.UPTOSTDT "
        
        End If
        BrokRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not BrokRec.EOF Then
            ItemGrid.Enabled = True
            Do While Not BrokRec.EOF
                DoEvents
                ItemRecGrid.AddNew
                ItemRecGrid.Fields("EXCODE") = BrokRec!excode
                ItemRecGrid.Fields("EXID") = BrokRec!EXID
                ItemRecGrid.Fields("ITEMID") = BrokRec!itemid
                ItemRecGrid.Fields("PARTY") = BrokRec!PARTY
                ItemRecGrid.Fields("ITEMCODE") = BrokRec!ITEMCODE
                ItemRecGrid.Fields("PARTYNAME") = BrokRec!NAME
                ItemRecGrid.Fields("DELETEROW") = "N"
                Set PExBrokRec = Nothing
                Set PExBrokRec = New ADODB.Recordset
                mysql = " SELECT EXCODE,AC_CODE,ITEMCODE,BROKTYPE,BROKRATE,UPTOSTDT FROM PITBROK WHERE COMPCODE =" & GCompCode & " AND INSTTYPE ='" & LInstType & "'"
                mysql = mysql & " AND EXCODE  ='" & BrokRec!excode & "'AND AC_CODE ='" & BrokRec!PARTY & "'AND ITEMID=" & BrokRec!itemid & ""
                mysql = mysql & " AND UPTOSTDT> = '" & Format(BrokRec!UPTOSTDT, "yyyy/MM/dd") & "'  "
                mysql = mysql & " ORDER BY UPTOSTDT "
                PExBrokRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                If Not PExBrokRec.EOF Then
                    If IsNull(PExBrokRec!broktype) Or PExBrokRec!broktype = "" Then
                        LBrokType = "Transaction"
                    Else
                        LBrokType = PExBrokRec!broktype
                    End If
                    LBrokRate = IIf(IsNull(PExBrokRec!brokrate), 0, PExBrokRec!brokrate)
                Else
                    LBrokRate = 0
                    LBrokType = "T"
                End If
                ItemRecGrid.Fields("BROKRATE") = LBrokRate
                Select Case LBrokType
                    Case "A"
                        RECGRID.Fields("BROKTYPE") = "A Opening ZLotwise"
                    Case "5"
                        RECGRID.Fields("BROKTYPE") = "5 Closing Sauda ZLotwise"
                    Case "B"
                        ItemRecGrid.Fields("BROKTYPE") = "BuySell Intraday"
                    Case "C"
                        ItemRecGrid.Fields("BROKTYPE") = "Closing Sauda"
                    Case "3"
                        ItemRecGrid.Fields("BROKTYPE") = "3 Closing Sauda %"
                    Case "D"
                        ItemRecGrid.Fields("BROKTYPE") = "Delivery Wise Brokerage"
                    Case "F"
                        ItemRecGrid.Fields("BROKTYPE") = "Fixed Brokerage"
                    Case "H"
                        ItemRecGrid.Fields("BROKTYPE") = "Higher Value Percentage Wise"
                    Case "I"
                        ItemRecGrid.Fields("BROKTYPE") = "IntraDay Brokerage"
                    Case "L"
                        ItemRecGrid.Fields("BROKTYPE") = "LotWise Higher Value "
                    Case "M"
                        ItemRecGrid.Fields("BROKTYPE") = "MRate Wise IntraDay"
                    Case "N"
                        ItemRecGrid.Fields("BROKTYPE") = "N Per Trade Wise"
                    Case "O"
                        ItemRecGrid.Fields("BROKTYPE") = "Opening Sauda"
                    Case "P"
                        ItemRecGrid.Fields("BROKTYPE") = "Percentage wise"
                    Case "Q"
                        ItemRecGrid.Fields("BROKTYPE") = "Qtywise IntraDay"
                    Case "R"
                        ItemRecGrid.Fields("BROKTYPE") = "RZLotwise Intraday"
                    Case "S"
                        ItemRecGrid.Fields("BROKTYPE") = "Slab Wise Brokerage"
                    Case "T"
                        ItemRecGrid.Fields("BROKTYPE") = "Transaction"
                    Case "U"
                        ItemRecGrid.Fields("BROKTYPE") = "U ShareQty Wise"
                    Case "V"
                        ItemRecGrid.Fields("BROKTYPE") = "Valuewise Intraday"
                    Case "X"
                        ItemRecGrid.Fields("BROKTYPE") = "XIntraday Higher Wise"
                    Case "Y"
                        ItemRecGrid.Fields("BROKTYPE") = "Y Qtywise Intraday"
                    Case "Z"
                        ItemRecGrid.Fields("BROKTYPE") = "ZLotwise"
                End Select
                LBrokType = IIf(IsNull(BrokRec!bbroktype), "T", BrokRec!bbroktype)
                Select Case LBrokType
                    Case "A"
                        ItemRecGrid.Fields("BROKTYPE") = "A Opening ZLotwise"
                    Case "5"
                        ItemRecGrid.Fields("BROKTYPE") = "5 Closing Sauda ZLotwise"
                    Case "D"
                        ItemRecGrid.Fields("BBROKTYPE") = "Delivery Wise Brokerage"
                    Case "C"
                        ItemRecGrid.Fields("BBROKTYPE") = "Closing Sauda"
                    Case "3"
                        ItemRecGrid.Fields("BBROKTYPE") = "3 Closing Sauda %"
                    Case "H"
                        ItemRecGrid.Fields("BBROKTYPE") = "Higher Value Percentage Wise"
                    Case "I"
                        ItemRecGrid.Fields("BBROKTYPE") = "IntraDay Brokerage"
                    Case "O"
                        ItemRecGrid.Fields("BBROKTYPE") = "Opening Sauda"
                    Case "P"
                        ItemRecGrid.Fields("BBROKTYPE") = "Percentage wise"
                    Case "Q"
                        ItemRecGrid.Fields("BBROKTYPE") = "Qtywise IntraDay"
                    Case "R"
                        ItemRecGrid.Fields("BBROKTYPE") = "RZLotwise Intraday"
                    Case "S"
                        ItemRecGrid.Fields("BBROKTYPE") = "Sub Brokerage in %"
                    Case "T"
                        ItemRecGrid.Fields("BBROKTYPE") = "Transaction"
                    Case "V"
                        ItemRecGrid.Fields("BBROKTYPE") = "Valuewise Intraday"
                    Case "X"
                        ItemRecGrid.Fields("BBROKTYPE") = "XIntraday Higher Wise"
                    Case "Z"
                        ItemRecGrid.Fields("BBROKTYPE") = "ZLotwise"
                End Select
                ItemRecGrid.Fields("BBROKRATE") = IIf(IsNull(BrokRec!bbrokrate), 0, BrokRec!bbrokrate)
                ItemRecGrid.Fields("BROKRATE2") = IIf(IsNull(BrokRec!BROKRATE2), 0, BrokRec!BROKRATE2)
                LShType = IIf(IsNull(BrokRec!SHTYPE), "G", BrokRec!SHTYPE)
                If LShType = "G" Then
                    ItemRecGrid.Fields("APPLYON") = "Gross Amount"
                ElseIf LShType = "S" Then
                    ItemRecGrid.Fields("APPLYON") = "ShareNet"
                Else
                    ItemRecGrid.Fields("APPLYON") = "Net Amount"
                End If
                ItemRecGrid.Fields("SHARE") = IIf(IsNull(BrokRec!SHRATE), 0, BrokRec!SHRATE)
                ItemRecGrid.Fields("FMLYCODE") = BrokRec!FMLYCODE
                ItemRecGrid.Fields("FMLYNAME") = BrokRec!FmlyNAME
                ItemRecGrid.Fields("FMLYID") = BrokRec!FMLYID
                ItemRecGrid.Fields("ACCID") = BrokRec!ACCID
                If IsNull(BrokRec!UPTOSTDT) Then
                    ItemRecGrid.Fields("UPTOSTDT") = Format(LSettlementDt, "YYYY/MM/DD")
                Else
                    ItemRecGrid.Fields("UPTOSTDT") = Format(BrokRec!UPTOSTDT, "YYYY/MM/DD")
                End If
                CountRow = CountRow + 1
                ItemRecGrid.Fields("New") = CountRow
                ItemRecGrid.Update
                BrokRec.MoveNext
            Loop
            Set ItemGrid.DataSource = ItemRecGrid: ItemGrid.ReBind: ItemGrid.Refresh: ItemRecGrid.MoveFirst
            ItemGrid.LeftCol = 0
            Label3.Visible = True
        Else
            If OptItem.Value = True Then
                If MsgBox("No Records. Do you really want to apply Seperate Sub Brokergae for selected Item/Script", vbYesNo + vbQuestion, "Confirm New Records") = vbYes Then
                    mysql = " SELECT FMLYID,FMLYCODE FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " "
                    If LenB(LFmlyIDs) > 0 Then mysql = mysql & " AND FMLYID IN (" & LFmlyIDs & ") "
                    mysql = mysql & " ORDER BY FMLYCODE "
                    Set LFmlyRec = Nothing
                    Set LFmlyRec = New ADODB.Recordset
                    LFmlyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                    Do While Not LFmlyRec.EOF
                        mysql = " SELECT PARTY,ACCID FROM ACCFMLYD WHERE FMLYID =" & LFmlyRec!FMLYID & " "
                        If LenB(LParties) > 0 Then mysql = mysql & " AND PARTY IN (" & LParties & ") "
                        mysql = mysql & " ORDER BY PARTY "
                        Set PartyRec = Nothing
                        Set PartyRec = New ADODB.Recordset
                        PartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                        Do While Not PartyRec.EOF
                            mysql = "SELECT EXID,ITEMID,EXCHANGECODE,ITEMCODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMID IN (" & LSItems & ") ORDER BY ITEMCODE "
                            Set TRec = Nothing
                            Set TRec = New ADODB.Recordset
                            TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                            Do While Not TRec.EOF
                                DoEvents
                                Call PInsert_Pitsbrok(PartyRec!PARTY, LFmlyRec!FMLYCODE, TRec!EXCHANGECODE, TRec!ITEMCODE, "P", 0, "N", 0, LSettlementDt, LInstType, 0, TRec!EXID, TRec!itemid, LFmlyRec!FMLYID, PartyRec!ACCID)
                                TRec.MoveNext
                            Loop
                            PartyRec.MoveNext
                        Loop
                        LFmlyRec.MoveNext
                    Loop
                    Set PartyRec = Nothing
                    Set LFmlyRec = Nothing
                    Set TRec = Nothing
                End If
                'CANCEL_REC
                Call CmdOk_Click
                Exit Sub
            End If
            Exit Sub
        End If
    End If
End Sub
Sub Save_ExBrok()
    Dim TempRec As ADODB.Recordset:    Dim LSDate As Date
    Dim TRec As ADODB.Recordset:        Dim TRec1 As ADODB.Recordset
    Dim TRec2 As ADODB.Recordset:       Dim LParty  As String
    Dim LFmlyCode As String
    Dim LFmlyID As Long
    Dim LACCID  As Long
    
    Set TempRec = Nothing
    Set TempRec = New ADODB.Recordset
    Set TempRec = RECGRID.Clone
    If TempRec.RecordCount > 0 Then
        Cnn.BeginTrans: CNNERR = True
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If Not IsNull(TempRec!excode) Then
                If LenB(TempRec!excode) > 0 Then
                    mysql = "DELETE FROM PEXSBROK WHERE INSTTYPE='" & LInstType & "'"
                    mysql = mysql & " AND PARTY= '" & TempRec!PARTY & "'"
                    mysql = mysql & " AND FMLYCODE ='" & TempRec!FMLYCODE & "'"
                    mysql = mysql & " AND EXID =" & TempRec!EXID & ""
                    mysql = mysql & " AND UpToStDt = '" & Format(TempRec!UPTOSTDT, "yyyy/MM/dd") & "' "
                    mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    If DateValue(TempRec!UPTOSTDT) > DateValue(GSysLockDt) Then
                        If ChkUpdBrok.Value = 1 Then
                            mysql = "UPDATE PEXBROK SET BROKTYPE='" & Left(TempRec!broktype, 1) & "',BROKRATE =" & Val(TempRec!brokrate) & ""
                            mysql = mysql & " WHERE AC_CODE ='" & TempRec!PARTY & "' AND EXID =" & TempRec!EXID & " AND UPTOSTDT ='" & Format(TempRec!UPTOSTDT, "YYYY/MM/DD") & "' AND INSTTYPE ='" & LInstType & "'"
                            Cnn.Execute mysql
                            'Call PInsert_PExBrok(TempRec!PARTY, TempRec!EXCODE, Left$(TempRec!BrokType, 1), Val(TempRec!BBROKRATE), Left$(TempRec!APPLYON, 1), Val(TempRec!SHARE), TempRec!UPTOSTDT, LInstType, TempRec!BROKRATE2, TempRec!EXID)
                        End If

                        Call PINSERT_PEXSBROK(TempRec!PARTY, TempRec!FMLYCODE, TempRec!excode, Left$(TempRec!bbroktype, 1), Val(TempRec!bbrokrate), Left$(TempRec!APPLYON, 1), Val(TempRec!SHARE), TempRec!UPTOSTDT, LInstType, TempRec!BROKRATE2, TempRec!EXID, TempRec!FMLYID, TempRec!ACCID)
                    Else
                        MsgBox "Sorry System Locked.  No Modification Allowed"
                        Exit Do
                    End If
                End If
            End If
            DoEvents
            TempRec.MoveNext
        Loop
        LSettlementDt = GFinEnd
        Set TRec = Nothing:        Set TRec = New ADODB.Recordset
        mysql = "SELECT * FROM ACCFMLYD WHERE COMPCODE=" & GCompCode & " AND FMLYID IN (" & LFmlyIDs & ") AND PARTY IN (" & LParties & ") ORDER BY FMLYID ,PARTY "
        TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        While Not TRec.EOF
            LParty = TRec!PARTY
            LFmlyCode = TRec!FMLYCODE
            LFmlyID = TRec!FMLYID
            LACCID = TRec!ACCID
            mysql = "SELECT EXID,EXCODE FROM EXMAST  WHERE COMPCODE  = " & GCompCode & " AND EXID IN (" & LExCodes & ")"
            mysql = mysql & "  ORDER BY EXCODE  "
            Set TRec1 = Nothing: Set TRec1 = New ADODB.Recordset: TRec1.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            Do While Not TRec1.EOF
                DoEvents
                mysql = "SELECT COMPCODE FROM PEXSBROK WHERE FMLYID =" & LFmlyID & " AND PARTY  ='" & LParty & "' AND EXID=" & TRec1!EXID & " AND UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE='" & LInstType & "'"
                Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset:
                TRec2.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                If TRec2.EOF Then
                    If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                        Call PINSERT_PEXSBROK(LParty, LFmlyCode, TRec1!excode, "P", 0, "N", 0, LSettlementDt, LInstType, 0, TRec1!EXID, LFmlyID, LACCID)
                    End If
                End If
                TRec1.MoveNext
            Loop
            TRec.MoveNext
        Wend
        Cnn.CommitTrans: CNNERR = False
    End If
End Sub
Sub Save_ItemBrok()
Dim TRec2 As ADODB.Recordset:   Dim TRec As ADODB.Recordset
Dim LSDate As Date:             Dim LParty As String
Dim LFmlyCode As String:        Dim TRec1  As ADODB.Recordset
Dim TempRec As ADODB.Recordset

    Set TempRec = Nothing
    Set TempRec = New ADODB.Recordset
    Set TempRec = ItemRecGrid.Clone
    If TempRec.RecordCount > 0 Then
        Cnn.BeginTrans: CNNERR = True
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If Not IsNull(TempRec!ITEMCODE) Then
                If LenB(TempRec!ITEMCODE) > 0 Then
                    mysql = "DELETE FROM PITSBROK WHERE INSTTYPE='" & LInstType & "'"
                    mysql = mysql & " AND PARTY= '" & TempRec!PARTY & "'"
                    mysql = mysql & " AND FMLYCODE = '" & TempRec!FMLYCODE & "'"
                    mysql = mysql & " AND EXID =" & TempRec!EXID & ""
                    mysql = mysql & " AND ITEMCODE ='" & TempRec!ITEMCODE & "'"
                    mysql = mysql & " AND UpToStDt = '" & Format(TempRec!UPTOSTDT, "yyyy/MM/dd") & "' "
                    mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    If TempRec!DELETEROW <> "Y" Then
                        If DateValue(TempRec!UPTOSTDT) > DateValue(GSysLockDt) Then
                            
                            If ChkUpdBrok.Value = 1 Then
                                mysql = "SELECT ac_code FROM PITBROK WHERE EXID = " & TempRec!EXID & " AND AC_CODE = '" & TempRec!PARTY & "' AND ITEMCODE ='" & TempRec!ITEMCODE & "' AND UPTOSTDT ='" & Format(TempRec!UPTOSTDT, "YYYY/MM/DD") & "' AND INSTTYPE ='" & LInstType & "' "
                                Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset:
                                TRec2.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                                If Not TRec2.EOF Then
                                    mysql = "UPDATE PITBROK SET BROKTYPE='" & Left(TempRec!broktype, 1) & "',BROKRATE =" & Val(TempRec!brokrate) & ""
                                    mysql = mysql & " WHERE AC_CODE ='" & TempRec!PARTY & "' AND ITEMCODE ='" & TempRec!ITEMCODE & "' AND EXID =" & TempRec!EXID & " AND UPTOSTDT ='" & Format(TempRec!UPTOSTDT, "YYYY/MM/DD") & "' AND INSTTYPE ='" & LInstType & "'"
                                    Cnn.Execute mysql
                                Else
                                    Call PInsert_PitBrok(GCompCode, TempRec!PARTY, TempRec!ITEMCODE, Left$(TempRec!broktype, 1), Val(TempRec!brokrate & ""), 0, 0, "P", 0, "I", 0, Format(TempRec!UPTOSTDT, "yyyy/MM/dd"), 0, 0, 0, "P", 0, 0, LInstType, TempRec!excode, TempRec!EXID, TempRec!itemid, TempRec!ACCID)
                                End If
                                'Call PInsert_PExBrok(TempRec!PARTY, TempRec!EXCODE, Left$(TempRec!BrokType, 1), Val(TempRec!BBROKRATE), Left$(TempRec!APPLYON, 1), Val(TempRec!SHARE), TempRec!UPTOSTDT, LInstType, TempRec!BROKRATE2, TempRec!EXID)
                            End If

                            Call PInsert_Pitsbrok(TempRec!PARTY, TempRec!FMLYCODE, TempRec!excode, TempRec!ITEMCODE, Left$(TempRec!bbroktype, 1), Val(TempRec!bbrokrate), Left$(TempRec!APPLYON, 1), Val(TempRec!SHARE), TempRec!UPTOSTDT, LInstType, TempRec!BROKRATE2, TempRec!EXID, TempRec!itemid, TempRec!FMLYID, TempRec!ACCID)
                        Else
                            MsgBox "Sorry System Locked.  No Modification Allowed"
                            Exit Do
                        End If
                    End If
                End If
            End If
            DoEvents
            TempRec.MoveNext
        Loop
        Dim LFmlyID As Long
        Dim LACCID As Long
                
        If ChkUpdLastSettle.Value Then
            LSettlementDt = GFinEnd
            Set TRec = Nothing:        Set TRec = New ADODB.Recordset
            mysql = "SELECT * FROM ACCFMLYD WHERE COMPCODE=" & GCompCode & " AND FMLYID IN (" & LFmlyIDs & ") AND PARTY IN (" & LParties & ") ORDER BY FMLYID,PARTY "
            TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            While Not TRec.EOF
                LParty = TRec!PARTY
                LFmlyCode = TRec!FMLYCODE
                LFmlyID = TRec!FMLYID
                LACCID = TRec!ACCID
                
                mysql = "SELECT EXID,ITEMID,EXCHANGECODE,ITEMCODE FROM ITEMMAST WHERE COMPCODE = " & GCompCode & " "
                If Len(LExCodes) > 0 Then mysql = mysql & " AND EXID   IN (" & LExCodes & ")"
                If Len(LSItems) > 0 Then mysql = mysql & " AND ITEMID IN (" & LSItems & ")"
                mysql = mysql & "  ORDER BY EXCHANGECODE ,ITEMCODE  "
                Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                Do While Not TRec2.EOF
                    DoEvents
                    mysql = "SELECT COMPCODE FROM PITSBROK WHERE FMLYID =" & LFmlyID & " "
                    mysql = mysql & " AND ACCID =" & LACCID & " AND ITEMID =" & TRec2!itemid & " AND "
                    mysql = mysql & " UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE='" & LInstType & "'"
                    Set TRec1 = Nothing: Set TRec1 = New ADODB.Recordset: TRec1.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                    If TRec1.EOF Then
                        If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                            'MYSQL = "EXEC INSERT_PITSBROK " & GCompCode & " ,'" & LParty & "','" & LFmlyCode & "','" & TRec2!EXCHANGECODE & "','" & TRec2!ITEMCODE & "','P',0,'G',0,'" & Format(LSettlementDt, "YYYY/MM/DD") & "','" & LInstType & "',0"
                           ' Cnn.Execute MYSQL
                            Call PInsert_Pitsbrok(LParty, LFmlyCode, TRec2!EXCHANGECODE, TRec2!ITEMCODE, "P", 0, "N", 0, LSettlementDt, LInstType, 0, TRec2!EXID, TRec2!itemid, LFmlyID, LACCID)
                        End If
                    End If
                    TRec2.MoveNext
                Loop
                TRec.MoveNext
            Wend
        End If
        Cnn.CommitTrans: CNNERR = False
    End If
End Sub
Private Sub Get_Items()
    Dim ChkCount As Integer
    Dim I As Integer
    
    LSItems = vbNullString
    For I = 1 To ItemList.ListItems.Count
        If ItemList.ListItems(I).Checked = True Then
            If LenB(LSItems) <> 0 Then LSItems = LSItems & ","
            LSItems = LSItems & ItemList.ListItems(I).SubItems(2) & ""
        End If
    Next I
End Sub

Private Sub OptExchange_Click()
    If OptItem.Value = True Then
        ItemList.Visible = True
        ChkUpdLastSettle.Value = 0
    Else
        ItemList.Visible = False
        ChkUpdLastSettle.Value = 1
    End If
    Set UptoDateRec = Nothing: Set UptoDateRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PEXSBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT"
    UptoDateRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not UptoDateRec.EOF Then
        Set UptoDateCombo.RowSource = UptoDateRec
        UptoDateCombo.ListField = "CONDATE"
        UptoDateCombo.BoundColumn = "CONDATE"
    End If
    SSTab1.Tab = 0
End Sub
Private Sub OptItem_Click()
    If OptItem.Value = True Then
        ItemList.Visible = True
        ChkUpdLastSettle.Value = 0
    Else
        ItemList.Visible = False
        ChkUpdLastSettle.Value = 1
    End If
    Set UptoDateRec = Nothing: Set UptoDateRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PITSBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT"
    UptoDateRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not UptoDateRec.EOF Then
        Set UptoDateCombo.RowSource = UptoDateRec
        UptoDateCombo.ListField = "CONDATE"
        UptoDateCombo.BoundColumn = "CONDATE"
    End If
       SSTab1.Tab = 0
End Sub

Private Sub Fill_ItemList()
    Call Get_ExCodes
    Dim ItemRec As ADODB.Recordset
    Set ItemRec = Nothing: Set ItemRec = New ADODB.Recordset
    mysql = "SELECT ITEMCODE,ITEMNAME,EXCHANGECODE,ITEMID FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " "
    If LenB(LExCodes) <> 0 Then mysql = mysql & " AND EXID  IN (" & LExCodes & ")"
    mysql = mysql & " ORDER BY ITEMCODE "
    ItemRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not ItemRec.EOF Then
        ItemRec.MoveFirst
        ItemList.ListItems.Clear
        ItemList.Enabled = True:
        Do While Not ItemRec.EOF
            If (ItemRec!EXCHANGECODE = "EQ" Or ItemRec!EXCHANGECODE = "BEQ") Then InstCombo.Visible = True
            ItemList.ListItems.Add , , ItemRec!ITEMCODE
            ItemList.ListItems(ItemList.ListItems.Count).ListSubItems.Add , , ItemRec!EXCHANGECODE
            ItemList.ListItems(ItemList.ListItems.Count).ListSubItems.Add , , ItemRec!itemid
            ItemRec.MoveNext
        Loop
    End If
        
End Sub
Private Sub Get_ExCodes()
    Dim ChkCount As Integer
    Dim I As Integer
    LExCodes = vbNullString
    ChkCount = 0
    For I = 1 To ExList.ListItems.Count
        If ExList.ListItems(I).Checked = True Then
            ChkCount = ChkCount + 1
            If LenB(LExCodes) <> 0 Then LExCodes = LExCodes & ","
            LExCodes = LExCodes & ExList.ListItems(I).ListSubItems(2) & ""
        End If
    Next I
    If ChkCount = ExList.ListItems.Count Then
        AllExcodes = True
    Else
        AllExcodes = False
    End If
End Sub


Private Sub ItemGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LGridRow  As Integer
    Dim LGridCol As Integer
    LActive_Grid = 2
    
    If KeyCode = 13 And ItemGrid.Col = 6 Then ' BROKTYPE
        BrokTypeCombo.Visible = True: BrokTypeCombo.SetFocus
    ElseIf KeyCode = 13 And ItemGrid.Col = 4 And ChkUpdBrok.Value = 1 Then ' BROKTYPE
        BrokTypeCombo.Visible = True: BrokTypeCombo.SetFocus
    ElseIf KeyCode = 13 And ItemGrid.Col = 11 Then 'APPLYON
        ShareTypeCombo.Visible = True: ShareTypeCombo.SetFocus
    ElseIf KeyCode = 13 And ItemGrid.Col = 15 Then 'delete row
        ItemGrid.text = UCase(ItemGrid.text)
        If ItemGrid.text = "Y" Then
        Else
            ItemGrid.text = "N"
        End If
        ItemGrid.SetFocus
        ItemGrid.Col = 14
    ElseIf KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
    'Press F7 to replace all rows with current cell value
    If KeyCode = 118 Then   'F7
        LGridRow = ItemGrid.Row
        LGridCol = ItemGrid.Col
        If ItemGrid.Col = 4 Then 'BROKTYPE
            GridColVal = ItemRecGrid!broktype
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!broktype = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 5 Then 'BROKRATE
            GridColVal = ItemRecGrid!brokrate
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!brokrate = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 6 Then 'BBROKTYPE
        
            GridColVal = ItemRecGrid!bbroktype
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!bbroktype = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 7 Then 'BBROKRATE
            GridColVal = ItemRecGrid!bbrokrate
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!bbrokrate = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 8 Then 'BBROKRATE
            GridColVal = ItemRecGrid!BROKRATE2
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!BBROKRATE2 = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 9 Then 'UPTOSTDT
            GridColVal = ItemRecGrid!UPTOSTDT
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!UPTOSTDT = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 10 Then 'SHARE
            GridColVal = ItemRecGrid!SHARE
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!SHARE = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 11 Then 'APPLYON
            GridColVal = ItemRecGrid!APPLYON
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!APPLYON = GridColVal
                ItemRecGrid.MoveNext
            Wend
        ElseIf ItemGrid.Col = 15 Then 'deletw row
            GridColVal = ItemRecGrid!DELETEROW
            ItemRecGrid.MoveFirst
            While Not ItemRecGrid.EOF
                ItemRecGrid!DELETEROW = GridColVal
                ItemRecGrid.MoveNext
            Wend
        End If
        Set ItemGrid.DataSource = ItemRecGrid: ItemGrid.ReBind: ItemGrid.Refresh
        ItemGrid.Row = LGridRow: ItemGrid.Col = LGridCol: ItemGrid.SetFocus
    End If
    If KeyCode = 119 Then   'F8
        Frame6.Visible = True
    End If
End Sub

Private Sub CmdApply_Click()
Dim TempRec As ADODB.Recordset:     Dim TempRec2 As ADODB.Recordset
Dim LStdTDate As Date:              Dim AccRec As ADODB.Recordset
Dim TRec As ADODB.Recordset:        Dim TRec2 As ADODB.Recordset
Dim LXAc_Code  As String
UptoDateCombo.BoundText = vbNullString
Frame9.Enabled = False
If OptExchange.Value = True Then
    If RECGRID.RecordCount > 0 Then
        Set TempRec = Nothing
        Set TempRec = RECGRID.Clone
        TempRec.MoveFirst
        Cnn.BeginTrans: CNNERR = True
        LSItems = vbNullString
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If Not IsNull(TempRec!excode) Then
                If LenB(TempRec!excode) > 0 Then
                    mysql = "DELETE FROM PEXSBROK WHERE COMPCODE =" & GCompCode & " AND INSTTYPE='" & LInstType & "' AND PARTY ='" & TempRec!PARTY & "'"
                    mysql = mysql & " AND FMLYCODE ='" & TempRec!FMLYCODE & "' AND EXCODE ='" & TempRec!excode & "'"
                    mysql = mysql & " AND UpToStDt = '" & Format(TempRec!UPTOSTDT, "yyyy/MM/dd") & "' "
                    mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    If ChkBrokLock.Value = 1 Then
                        LStdTDate = DateValue(vcDTP1.Value)
                        If LStdTDate > DateValue(GSysLockDt) Then
                            Call PINSERT_PEXSBROK(TempRec!PARTY, TempRec!FMLYCODE, TempRec!excode, Left$(TempRec!bbroktype, 1), Val(TempRec!bbrokrate), Left$(TempRec!APPLYON, 1), Val(TempRec!SHARE), LStdTDate, LInstType, TempRec!BROKRATE2, TempRec!EXID, TempRec!FMLYID, TempRec!ACCID)
                        Else
                            MsgBox "Sorry System Locked.  No Modification Allowed"
                            Exit Do
                        End If
                    End If
                End If
            End If
            TempRec.MoveNext
        Loop
        LSettlementDt = GFinEnd
        Set TempRec = Nothing:        Set AccRec = Nothing:        Set AccRec = New ADODB.Recordset
        mysql = "SELECT FMLYCODE,PARTY,FMLYID,ACCID  FROM ACCFMLYD WHERE COMPCODE=" & GCompCode & " AND FMLYID IN (" & LFmlyIDs & ") "
        mysql = mysql & " AND PARTY IN (" & LParties & ") ORDER BY FMLYCODE,PARTY "
        AccRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        While Not AccRec.EOF
            mysql = "SELECT EXID,EXCODE FROM EXMAST WHERE COMPCODE  = " & GCompCode & " "
            If LenB(LExCodes) > 0 Then mysql = mysql & " AND EXID  IN (" & LExCodes & ")"
            mysql = mysql & "  ORDER BY EXCODE "
            Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not TRec.EOF
                DoEvents
                LXAc_Code = Get_PEXSBROK_AC_CODE(AccRec!FMLYID, AccRec!ACCID, TRec!EXID, LSettlementDt, LInstType)
                'MYSQL = "SELECT PARTY FROM PEXSBROK WHERE COMPCODE=" & GCompCode & " AND PARTY ='" & AccRec!PARTY & "' AND FMLYCODE ='" & AccRec!FMLYCODE & "'  "
                'MYSQL = MYSQL & " AND EXID =" & TRec!EXID & " AND UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE ='" & LInstType & "'"
                'Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
                If Len(LXAc_Code) < 1 Then
                    If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                        Call PINSERT_PEXSBROK(AccRec!PARTY, AccRec!FMLYCODE, TRec!excode, "P", 0, "N", 0, LSettlementDt, LInstType, 0, TRec!EXID, AccRec!FMLYID, AccRec!ACCID)
                    End If
                End If
                TRec.MoveNext
            Loop
            Set TRec = Nothing
            AccRec.MoveNext
        Wend
        Cnn.CommitTrans
        CNNERR = False
        Set AccRec = Nothing
    End If
    Fill_ExchangeGrid
Else
    If ItemRecGrid.RecordCount > 0 Then
        Set TempRec2 = Nothing:        Set TempRec2 = ItemRecGrid
        TempRec2.MoveFirst:        Cnn.BeginTrans: CNNERR = True
        Do While Not TempRec2.EOF
            If Not IsNull(TempRec2!ITEMCODE) Then
                If LenB(TempRec2!ITEMCODE) > 0 Then
                    mysql = "DELETE FROM PITSBROK WHERE COMPCODE =" & GCompCode & " AND INSTTYPE='" & LInstType & "' "
                    mysql = mysql & " AND PARTY = '" & TempRec2!PARTY & "'"
                    mysql = mysql & " AND ITEMCODE ='" & TempRec2!ITEMCODE & "'"
                    mysql = mysql & " AND FMLYCODE ='" & TempRec2!FMLYCODE & "'"
                    mysql = mysql & " AND UpToStDt = '" & Format(TempRec2!UPTOSTDT, "yyyy/MM/dd") & "' "
                    mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    If TempRec2!DELETEROW <> "Y" Then
                        If ChkBrokLock.Value = 1 Then
                            LStdTDate = DateValue(vcDTP1.Value)
                            If DateValue(LStdTDate) > DateValue(GSysLockDt) Then
                                Call PInsert_Pitsbrok(TempRec2!PARTY, TempRec2!FMLYCODE, TempRec2!excode, TempRec2!ITEMCODE, Left$(TempRec2!bbroktype, 1), Val(TempRec2!bbrokrate & vbNullString), Left$(TempRec2!APPLYON, 1), _
                                Val(TempRec2!SHARE & vbNullString), LStdTDate, LInstType, TempRec2!BROKRATE2, TempRec2!EXID, TempRec2!itemid, TempRec2!FMLYID, TempRec2!ACCID)
                            Else
                                MsgBox "Sorry System Locked.  No Modification Allowed"
                                Exit Do
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            TempRec2.MoveNext
        Loop
        Set TempRec2 = Nothing
        LSettlementDt = GFinEnd
        If ChkUpdLastSettle.Value = 1 Then
            
            Call Get_Items
            Set AccRec = Nothing
            Set AccRec = New ADODB.Recordset
            mysql = "SELECT FMLYID,ACCID,FMLYCODE,PARTY FROM ACCFMLYD WHERE COMPCODE=" & GCompCode & " AND FMLYID IN (" & LFmlyIDs & ") "
            mysql = mysql & " AND PARTY IN (" & LParties & ") ORDER BY FMLYCODE,PARTY "
            AccRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            While Not AccRec.EOF
                mysql = "SELECT EXID,ITEMID,EXCHANGECODE,ITEMCODE FROM ITEMMAST  WHERE COMPCODE  = " & GCompCode & " "
                If LenB(LSItems) > 1 Then mysql = mysql & " AND ITEMID   IN (" & LSItems & ")"
                mysql = mysql & "  ORDER BY EXCHANGECODE,ITEMCODE  "
                Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                Do While Not TRec.EOF
                    DoEvents
                    mysql = "SELECT PARTY FROM PITSBROK WHERE COMPCODE=" & GCompCode & " AND PARTY ='" & AccRec!PARTY & "' AND FMLYCODE ='" & AccRec!FMLYCODE & "' "
                    mysql = mysql & " AND ITEMID  =" & TRec!itemid & " AND UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE ='" & LInstType & "'"
                    Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                    If TRec2.EOF Then
                        If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                            Call PInsert_Pitsbrok(AccRec!PARTY, AccRec!FMLYCODE, TRec!EXCHANGECODE, TRec!ITEMCODE, "P", 0, "N", 0, LSettlementDt, LInstType, 0, TRec!EXID, TRec!itemid, AccRec!FMLYID, AccRec!ACCID)
                        End If
                    End If
                    TRec.MoveNext
                Loop
                AccRec.MoveNext
            Wend
            Set TRec = Nothing
            Set AccRec = Nothing
            Cnn.CommitTrans
            CNNERR = False
        End If
        Fill_ItemGrid
    End If
End If
                               
End Sub





Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim scriptcontrol As New scriptcontrol
scriptcontrol.Language = "vbscript"
    If KeyAscii = 13 Then
        If Text1.text <> "" Then
            Dim Lformula As String
            Dim Lbbrokrate As Double
            Dim LSHARE As Double
            Dim Lvar As Variant
            Dim mno1 As Double

            RECGRID.MoveFirst
            While Not RECGRID.EOF
                Lformula = UCase(Text1.text)
                If InStr(1, Lformula, "BROKRATE") > 0 Then
                    Lformula = Replace(Lformula, "BROKRATE", RECGRID!brokrate)
                End If
                If InStr(1, Lformula, "SHARE") > 1 Then
                    Lformula = Replace(Lformula, "SHARE", RECGRID!SHARE)
                End If
                Lvar = Lformula
                RECGRID!bbrokrate = Val(scriptcontrol.Eval(Lvar))
                RECGRID.MoveNext
            Wend
            Set ExGrid.DataSource = RECGRID: ExGrid.ReBind: ExGrid.Refresh
            ExGrid.Row = 1: ExGrid.Col = 1: ExGrid.SetFocus
            Frame6.Visible = False
        End If
    End If
End Sub

