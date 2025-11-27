VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PrtUpfrm 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PrtUpfrm"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   12780
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   13935
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Daily Trade Modify"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   13695
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12640511
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Party Contract Details"
      TabPicture(0)   =   "PrtUpfrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DataGrid1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DataCombo3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "vcDTP2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "vcDTP1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListView1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.CheckBox Check2 
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
         Height          =   285
         Left            =   5640
         TabIndex        =   5
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
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
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   5640
         TabIndex        =   4
         ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select members, F5 to select non members."
         Top             =   600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parties"
            Object.Width           =   8026
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OpBal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "FmlyId"
            Object.Width           =   0
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   330
         Left            =   8040
         TabIndex        =   0
         Top             =   -15
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
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
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   330
         Left            =   10620
         TabIndex        =   1
         Top             =   -15
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select item specific."
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5311
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SAUDAID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Lot"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "exchange"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SDutyType"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "SDutyRate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SDutyPer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "MTYPE"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   420
         Left            =   5040
         TabIndex        =   7
         Top             =   4920
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   741
         _Version        =   393216
         ForeColor       =   12582912
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   8281
         _Version        =   393216
         AllowArrows     =   -1  'True
         BackColor       =   16776960
         HeadLines       =   1
         RowHeight       =   23
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
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "Condate"
            Caption         =   "Condate"
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
            DataField       =   "Sauda"
            Caption         =   "Sauda"
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
         BeginProperty Column03 
            DataField       =   "Name"
            Caption         =   "Name"
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
            DataField       =   "ConType"
            Caption         =   "Type"
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
            DataField       =   "Qnty"
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
            DataField       =   "RATE"
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
            DataField       =   "SaudaCode"
            Caption         =   "SaudaCode"
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
            DataField       =   "CONNO"
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
         BeginProperty Column09 
            DataField       =   "OldParty"
            Caption         =   "OldParty"
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
         BeginProperty Column12 
            DataField       =   "trdtime"
            Caption         =   "Trade Time"
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
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   2594.835
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1019.906
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
               ColumnWidth     =   1800
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Index           =   1
         Left            =   9720
         TabIndex        =   13
         Top             =   45
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Index           =   0
         Left            =   6720
         TabIndex        =   12
         Top             =   0
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parties"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   2
         Left            =   5640
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   5760
         Width           =   9375
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   2640
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   13695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000040&
      BorderWidth     =   12
      Height          =   9420
      Left            =   75
      Top             =   840
      Width           =   14085
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3480
      TabIndex        =   14
      Top             =   360
      Width           =   105
   End
End
Attribute VB_Name = "PrtUpfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecLst As ADODB.Recordset
Dim Rec_Account As ADODB.Recordset
Dim REC1 As ADODB.Recordset
Dim LSaudaCode As String
Dim LPartyCode As String
Sub SaudaList()
    Me.MousePointer = 11: ListView2.TabStop = False: Check1.TabStop = False
    ListView2.ListItems.Clear
    Dim TRec As ADODB.Recordset
    
    Check1.Value = 0
       
    MYSQL = "SELECT SAUDAID,SAUDACODE,SAUDANAME,ITEMCODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & ""
    MYSQL = MYSQL & " AND SAUDACODE IN (SELECT DISTINCT SAUDA FROM CTR_M WHERE COMPCODE= " & GCompCode & " "
    MYSQL = MYSQL & " AND CONDATE >='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' "
    MYSQL = MYSQL & " AND CONDATE <='" & Format(vcDTP2.Value, "yyyy/MM/dd") & "') "
    MYSQL = MYSQL & " ORDER BY ITEMCODE,SAUDAMAST.MATURITY "
        
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            ListView2.ListItems.Add , , TRec!SAUDACODE
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , TRec!SAUDANAME
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , TRec!SAUDAID
            TRec.MoveNext
        Loop
        ListView2.TabStop = True: Check1.TabStop = True
    Else
        MsgBox "Sauda does not exists for Selected Date Range.", vbInformation
    End If
    Me.MousePointer = 0
End Sub
Sub PartyList()
    Me.MousePointer = 11: ListView1.TabStop = False: Check2.TabStop = False
    Dim I As Integer
    Dim TRec As ADODB.Recordset
    LSaudaCode = vbNullString
    For I = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(I).Checked = True Then LSaudaCode = LSaudaCode & "'" & ListView2.ListItems(I).text & "'"
        If I < ListView2.ListItems.Count Then
            If ListView2.ListItems(I + 1).Checked = True And Len(LSaudaCode) > Val(0) Then LSaudaCode = LSaudaCode & ", "
        End If
    Next
    LSaudaCode = Trim(LSaudaCode)
    
    If LSaudaCode = "" Then
        ListView1.ListItems.Clear
    Else
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        MYSQL = "SELECT  ACCOUNTM.NAME,ACCOUNTM.AC_CODE FROM ACCOUNTM,ACCOUNTD WHERE ACCOUNTM.COMPCODE=" & GCompCode & " "
        MYSQL = MYSQL & " AND  ACCOUNTM.COMPCODE=ACCOUNTD.COMPCODE AND ACCOUNTM.AC_CODE=ACCOUNTD.AC_CODE "
        MYSQL = MYSQL & " AND ACCOUNTD.AC_CODE IN (SELECT DISTINCT PARTY FROM CTR_D WHERE COMPCODE = " & GCompCode & "  AND CONDATE<='" & Format(vcDTP2.Value, "yyyy/MM/dd") & "'"
        MYSQL = MYSQL & " AND SAUDA IN (" & LSaudaCode & "))  ORDER BY ACCOUNTM.NAME"
        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            ListView1.ListItems.Clear
            Do While Not TRec.EOF
                ListView1.ListItems.Add , , TRec!NAME
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , TRec!AC_CODE
                TRec.MoveNext
            Loop
            ListView1.Visible = True: ListView1.TabStop = True: Check2.TabStop = True
        End If
    End If
    Me.MousePointer = 0
End Sub
Private Sub Check1_Click()
    
    Dim I As Integer
    For I = 1 To ListView2.ListItems.Count
        If Check1.Value = 1 Then
            ListView2.ListItems.Item(I).Checked = True
        Else
            ListView2.ListItems.Item(I).Checked = False
        End If
    Next I
    Call PartyList
End Sub
Private Sub Check2_Click()
Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        If Check2.Value = 1 Then
            ListView1.ListItems.Item(I).Checked = True
        Else
            ListView1.ListItems.Item(I).Checked = False
        End If
    Next I
    Call GenerateLst
End Sub
Private Sub DataCombo3_GotFocus()
    DataCombo3.Left = Val(4090): DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        RecLst!PARTY = DataCombo3.BoundText
        RecLst!NAME = DataCombo3.text
        DataGrid1.TabStop = True: DataGrid1.Col = 3: DataGrid1.SetFocus: DataCombo3.Visible = False
    ElseIf KeyCode = 27 Then
        DataGrid1.SetFocus: DataCombo3.Visible = False
    End If
End Sub
Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    Dim A As Long
    If ColIndex = Val(2) Then
        Rec_Account.MoveFirst: Rec_Account.Find "AC_CODE='" & DataGrid1.text & "'", , adSearchForward
        If Not Rec_Account.EOF Then
            RecLst!PARTY = Rec_Account!AC_CODE: RecLst!NAME = Rec_Account!NAME: DataGrid1.Col = 3: DataGrid1.SetFocus
            DataGrid1.Col = DataGrid1.Col - 1
            A = RecLst.AbsolutePosition
            If A < RecLst.RecordCount Then
                DataGrid1.Row = DataGrid1.Row + 1
            End If
            DataGrid1.Col = DataGrid1.Col - 1
        Else
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        End If
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static OLDVAL As Integer
    Select Case ColIndex
    Case 0
        If OLDVAL = -1 Then
            RecLst.Sort = "CONDATE DESC"
        Else
            RecLst.Sort = "CONDATE"
        End If
    Case 1
        If OLDVAL = -1 Then
            RecLst.Sort = "SAUDA DESC"
        Else
            RecLst.Sort = "SAUDA"
        End If
    Case 2
        If OLDVAL = -1 Then
            RecLst.Sort = "PARTY DESC"
        Else
            RecLst.Sort = "PARTY"
        End If
    Case 3
        If OLDVAL = -1 Then
            RecLst.Sort = "NAME DESC"
        Else
            RecLst.Sort = "NAME"
        End If
    Case 4
        If OLDVAL = -1 Then
            RecLst.Sort = "CONTYPE DESC"
        Else
            RecLst.Sort = "CONTYPE"
        End If
    Case 5
        If OLDVAL = -1 Then
            RecLst.Sort = "QNTY DESC"
        Else
            RecLst.Sort = "QNTY"
        End If
    Case 6
        If OLDVAL = -1 Then
            RecLst.Sort = "RATE DESC"
        Else
            RecLst.Sort = "RATE"
        End If
    Case 7
        If OLDVAL = -1 Then
            RecLst.Sort = "CONNO DESC"
        Else
            RecLst.Sort = "CONNO"
        End If
    Case 8
        If OLDVAL = -1 Then
            RecLst.Sort = "TRDTIME DESC"
        Else
            RecLst.Sort = "TRDTIME"
        End If
    End Select
    If OLDVAL = ColIndex Then
        OLDVAL = -1
    Else
        OLDVAL = ColIndex
    End If
End Sub

Public Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 Or KeyCode = 9) And DataGrid1.Col = 2 Then
        If DataGrid1.text = "" Then
            DataCombo3.Visible = True: DataCombo3.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    On Error Resume Next
    If Me.ActiveControl.NAME = "vcDTP1" Or Me.ActiveControl.NAME = "vcDTP2" Then
        Sendkeys "{tab}"
    End If
End If

End Sub

Private Sub Form_Load()
    vcDTP1.Value = Date: vcDTP2.Value = Date
    Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
    Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Account.EOF Then Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
End Sub

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub ListView1_Click()
    Call GenerateLst
End Sub
Private Sub ListView2_Click()
    Call PartyList
End Sub


Private Sub vcDTP1_Validate(Cancel As Boolean)
    If SYSTEMLOCK(DateValue(vcDTP1.Value)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
        Cancel = True
    Else
        Call SaudaList
    End If
End Sub

Sub RecSet()
    Set RecLst = Nothing
    Set RecLst = New ADODB.Recordset
    RecLst.Fields.Append "Condate", adVarChar, 15, adFldIsNullable
    RecLst.Fields.Append "Sauda", adVarChar, 50, adFldIsNullable
    RecLst.Fields.Append "Party", adVarChar, 50, adFldIsNullable
    RecLst.Fields.Append "Name", adVarChar, 150, adFldIsNullable
    RecLst.Fields.Append "ConType", adVarChar, 2, adFldIsNullable '0
    RecLst.Fields.Append "Qnty", adDouble, , adFldIsNullable
    RecLst.Fields.Append "RATE", adDouble, , adFldIsNullable
    RecLst.Fields.Append "SaudaCode", adVarChar, 50, adFldIsNullable
    RecLst.Fields.Append "CONNO", adVarChar, 15, adFldIsNullable
    RecLst.Fields.Append "OldParty", adVarChar, 50, adFldIsNullable
    RecLst.Fields.Append "TRDTIME", adVarChar, 35, adFldIsNullable
    RecLst.Fields.Append "SAUDAID", adInteger, , adFldIsNullable
    RecLst.Open , , adOpenKeyset, adLockOptimistic
End Sub
Sub GenerateLst()
    Dim J As Integer
    Dim CountJ As Integer
    Dim TRec As ADODB.Recordset
    LPartyCode = vbNullString
    CountJ = Val(0)
    For J = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(J).Checked = True Then
            If Len(LPartyCode) > 1 Then
                LPartyCode = LPartyCode & ",'" & ListView1.ListItems(J).SubItems(1) & "'"
            Else
                LPartyCode = "'" & ListView1.ListItems(J).SubItems(1) & "'"
            End If
        End If
    Next
    Call RecSet
    If Len(LPartyCode) > 0 Then
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        MYSQL = "SELECT DISTINCT C.CONDATE,C.SAUDA,C.SAUDAID,C.CONSNO,C.ITEMCODE,C.PARTY,C.ConType,C.Qty,C.RATE,C.CONNO,C.CONTIME,A.NAME  "
        MYSQL = MYSQL & "FROM CTR_D AS C ,ACCOUNTM AS A WHERE C.COMPCODE =" & GCompCode & " AND C.COMPCODE = A.COMPCODE "
        MYSQL = MYSQL & " AND C.CONDATE >= '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND C.CONDATE <= '" & Format(vcDTP2.Value, "yyyy/MM/dd") & "' "
        MYSQL = MYSQL & " AND C.COMPCODE=A.COMPCODE AND C.PARTY=A.AC_CODE "
        If LenB(LSaudaCode) > 1 Then MYSQL = MYSQL & " AND C.SAUDA IN (" & LSaudaCode & ") "
        MYSQL = MYSQL & "AND C.PARTY IN (" & LPartyCode & ") "
        MYSQL = MYSQL & "ORDER BY C.CONDATE,C.SAUDA "
        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            Do While Not TRec.EOF
                RecLst.AddNew
                RecLst!Condate = TRec!Condate
                RecLst!Sauda = TRec!Sauda
                RecLst!PARTY = TRec!PARTY
                RecLst!NAME = TRec!NAME
                RecLst!CONTYPE = TRec!CONTYPE
                RecLst!QNTY = TRec!QTY
                RecLst!Rate = TRec!Rate
                RecLst!SAUDACODE = TRec!Sauda
                RecLst!CONNO = TRec!CONNO
                RecLst!TRDTIME = TRec!CONTIME
                RecLst!SAUDAID = TRec!SAUDAID
                RecLst!OLDPARTY = TRec!PARTY
            RecLst.Update
            TRec.MoveNext
        Loop
        If Not RecLst.EOF Then RecLst.MoveFirst: Set DataGrid1.DataSource = RecLst: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus
        GETMAIN.Toolbar1_Buttons(4).Enabled = True: GETMAIN.saverec.Enabled = True
    End If
    End If
End Sub
Sub UpdateParty()
    Dim LParties As String:     Dim LTrfto As String
    Dim LSauda As String:       Dim CountJ As Integer
    Dim LAC_CODE As String
    Dim LACCID As Long
    Me.MousePointer = 11
    LSauda = vbNullString: LParties = vbNullString: LTrfto = vbNullString: CountJ = 0:  CNNERR = False
    Cnn.BeginTrans
    RecLst.MoveFirst
    While Not RecLst.EOF
        LAC_CODE = Get_AccountDCode(RecLst!PARTY)
        If LenB(LAC_CODE) > 1 Then
            If RecLst!PARTY <> RecLst!OLDPARTY Then
                LACCID = Get_AccID(RecLst!PARTY)
                MYSQL = "UPDATE  CTR_D SET  PARTY='" & RecLst!PARTY & "',ACCID=" & LACCID & ""
                MYSQL = MYSQL & " WHERE COMPCODE = " & GCompCode & ""
                MYSQL = MYSQL & " AND PARTY ='" & RecLst!OLDPARTY & "' "
                MYSQL = MYSQL & " AND CONDATE ='" & Format(RecLst!Condate, "yyyy/MM/dd") & "'"
                MYSQL = MYSQL & " AND SAUDAID ='" & RecLst!SAUDAID & "' "
                MYSQL = MYSQL & " AND CONTYPE  = '" & RecLst!CONTYPE & "' "
                MYSQL = MYSQL & " AND CONNO =" & RecLst!CONNO & ""
                Cnn.Execute MYSQL
                If LenB(LParties) < 1 Then
                    LParties = "'" & RecLst!PARTY & "'"
                Else
                    LParties = LParties & ",'" & RecLst!PARTY & "'"
                End If
                If LenB(LParties) < 1 Then
                    LParties = "'" & RecLst!OLDPARTY & "'"
                Else
                    LParties = LParties & ",'" & RecLst!OLDPARTY & "'"
                End If
                If LenB(LSauda) < 1 Then
                    LSauda = Str(RecLst!SAUDAID)
                Else
                    LSauda = LSauda & "," & Str(RecLst!SAUDAID) & ""
                End If
            End If
        Else
            MsgBox "Invalid Party for Trade No " & RecLst!CONNO & ""
        End If
        RecLst.MoveNext
    Wend
    Cnn.CommitTrans: CNNERR = False
    'Call Update_BrokTran(LParties, vbNullString, vbNullString, LSauda, vcDTP1.Value, vcDTP1.Value)
    Call Update_Charges(LParties, vbNullString, LSauda, vbNullString, vcDTP1.Value, vcDTP1.Value, True)
    CNNERR = True
    Cnn.BeginTrans
    If BILL_GENERATION(CDate(vcDTP1.Value), CDate(GFinEnd), LSauda, LParties, vbNullString) Then
        Cnn.CommitTrans: CNNERR = False
    Else
        Cnn.RollbackTrans: CNNERR = False
    End If
    LParties = vbNullString
    LSauda = vbNullString
    'Call Chk_Billing
    'CLEAR SCREEN
    Set RecLst = Nothing: Set DataGrid1.DataSource = RecLst: DataGrid1.ReBind: DataGrid1.Refresh
    Check1.Value = 0: Check2.Value = 0
    ListView1.ListItems.Clear: ListView2.ListItems.Clear
    GETMAIN.Toolbar1_Buttons(4).Enabled = False: GETMAIN.saverec.Enabled = False
    Me.MousePointer = 0

err1:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    End If

    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
    GETMAIN.ProgressBar1.Visible = False

End Sub
Private Sub vcDTP2_Validate(Cancel As Boolean)
    If SYSTEMLOCK(DateValue(vcDTP2.Value)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
        Cancel = True
    Else
        Call SaudaList
    End If
End Sub
