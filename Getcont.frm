VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form GETCont 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   18960
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame13 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "ERGGHE"
      Height          =   4095
      Left            =   3000
      TabIndex        =   72
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
         TabIndex        =   74
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
         TabIndex        =   73
         ToolTipText     =   "Close"
         Top             =   -15
         Width           =   615
      End
      Begin VB.Label Label15 
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
         TabIndex        =   75
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtCalval 
      Height          =   495
      Left            =   7560
      TabIndex        =   71
      Text            =   "Text2"
      Top             =   10800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox TxtSaudaCode 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   360
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   70
      Top             =   10800
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox TxtItemId 
      Height          =   405
      Left            =   9360
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtExid 
      Height          =   405
      Left            =   10920
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtSaudaID 
      Height          =   405
      Left            =   13080
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   7095
      Left            =   19080
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
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
         Left            =   4800
         TabIndex        =   55
         Top             =   120
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   11033
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "Select Parties to Filter"
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
         TabIndex        =   54
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   1215
      Left            =   240
      TabIndex        =   47
      Top             =   9240
      Visible         =   0   'False
      Width           =   16215
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   9720
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   3720
         TabIndex        =   49
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ForeColor       =   64
         Text            =   ""
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserId"
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
         Left            =   720
         TabIndex        =   48
         Top             =   540
         Width           =   630
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
      Height          =   615
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   18855
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   18855
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Contract Entry"
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
            TabIndex        =   44
            Top             =   120
            Width           =   18615
         End
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
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
         Left            =   9720
         TabIndex        =   59
         Top             =   480
         Width           =   6855
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   10
      Top             =   -720
      Width           =   9975
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
         Left            =   8640
         TabIndex        =   16
         Top             =   240
         Width           =   1080
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
         Left            =   5094
         TabIndex        =   15
         Top             =   240
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
         Left            =   6777
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
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
         Left            =   3261
         TabIndex        =   13
         Top             =   240
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
         Left            =   1713
         TabIndex        =   12
         Top             =   240
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   9600
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
            Picture         =   "Getcont.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getcont.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7725
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   18420
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   1095
         Left            =   15360
         TabIndex        =   56
         Top             =   120
         Width           =   2895
         Begin VB.CheckBox ChkCheckRate 
            BackColor       =   &H00FF8080&
            Caption         =   "Check Rate"
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
            Left            =   1440
            TabIndex        =   67
            Top             =   120
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Filter Party"
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
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   5895
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   18135
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   360
            Left            =   1080
            TabIndex        =   46
            Top             =   840
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   64
            Text            =   ""
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5775
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   17895
            _ExtentX        =   31565
            _ExtentY        =   10186
            _Version        =   393216
            AllowArrows     =   -1  'True
            BackColor       =   12640511
            ForeColor       =   4194368
            HeadLines       =   1
            RowHeight       =   21
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
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   21
            BeginProperty Column00 
               DataField       =   "SRNO"
               Caption         =   "S.No"
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
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "BRATE"
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
               DataField       =   "OLDTRD"
               Caption         =   "OLDTRD"
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
            BeginProperty Column19 
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
            BeginProperty Column20 
               DataField       =   "ORDTIME"
               Caption         =   "ORDERTIME"
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
                  WrapText        =   -1  'True
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   1395.213
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
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column13 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column15 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column16 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column17 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column18 
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column19 
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column20 
                  ColumnWidth     =   1995.024
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         TabIndex        =   24
         Top             =   7080
         Width           =   18015
         Begin VB.TextBox TxtTotSAmt 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   13695
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox TxtTotBuyAmt 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox TxtSellAvg 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   11640
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   120
            Width           =   1125
         End
         Begin VB.TextBox TxtBuyAvg 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   3615
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   120
            Width           =   1000
         End
         Begin VB.TextBox TxtDiffAmt 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   16560
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox TxtTotSQty 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   120
            Width           =   1000
         End
         Begin VB.TextBox TxtTotBQty 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   1000
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Totals"
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
            Left            =   120
            TabIndex        =   39
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Buy Qty"
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
            Left            =   720
            TabIndex        =   38
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Avg Rate"
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
            Left            =   10680
            TabIndex        =   37
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Sell Qty"
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
            Left            =   8640
            TabIndex        =   36
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Buy Amt"
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
            Left            =   4800
            TabIndex        =   35
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Avg. Rate"
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
            Left            =   2640
            TabIndex        =   34
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Sell Amt"
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
            Left            =   12840
            TabIndex        =   33
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label12 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   15600
            TabIndex        =   32
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   15015
         Begin VB.TextBox TxtStrike 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   600
            Width           =   1305
         End
         Begin VB.TextBox TxtOptType 
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
            Height          =   405
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox TxtInst 
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
            Height          =   405
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox TxtExCode 
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   600
            Width           =   825
         End
         Begin VB.CheckBox Check1 
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
            Height          =   360
            Left            =   2490
            TabIndex        =   2
            Top             =   120
            Width           =   2655
         End
         Begin VB.TextBox TxtItemName 
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
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   600
            Width           =   2175
         End
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   8760
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox TxtAdminPasswd 
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
            Left            =   13320
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   600
            Visible         =   0   'False
            Width           =   1530
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   1785
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
            ForeColor       =   &H00000040&
            Height          =   405
            ItemData        =   "Getcont.frx":08A4
            Left            =   10680
            List            =   "Getcont.frx":08AE
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Width           =   1815
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   360
            Left            =   960
            TabIndex        =   1
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
            Value           =   37860.8625462963
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   420
            Left            =   6000
            TabIndex        =   3
            Top             =   120
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   741
            _Version        =   393216
            ForeColor       =   64
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
            Left            =   13320
            TabIndex        =   5
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   741
            _Version        =   393216
            ForeColor       =   64
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ExCode"
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
            Index           =   5
            Left            =   120
            TabIndex        =   61
            Top             =   638
            Width           =   750
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UserId"
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
            Left            =   12480
            TabIndex        =   58
            Top             =   165
            Width           =   690
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Admin Pass"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   12600
            TabIndex        =   42
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   0
            Left            =   8280
            TabIndex        =   40
            Top             =   660
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   23
            Top             =   158
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   2
            Left            =   5280
            TabIndex        =   22
            Top             =   165
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
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
            Index           =   3
            Left            =   5280
            TabIndex        =   21
            Top             =   638
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set. Rate"
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
            Left            =   9675
            TabIndex        =   20
            Top             =   638
            Width           =   915
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   18
            Left            =   9675
            TabIndex        =   19
            Top             =   165
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   6840
         Width           =   14055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   165
         TabIndex        =   9
         Top             =   1200
         Width           =   11415
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   0
         Left            =   720
         Picture         =   "Getcont.frx":08C5
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   1
         Left            =   1080
         Picture         =   "Getcont.frx":0BCF
         Stretch         =   -1  'True
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000011&
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7980
      Left            =   120
      Top             =   720
      Width           =   18765
   End
End
Attribute VB_Name = "GETCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LDataImport As Byte:                Public Fb_Press As Byte:            Dim LUserId As String
Dim LFParties As String:                Dim LSParties As String:            Dim FCalval As Double
Dim LSConno As Long:                    Dim LUserIDRec As ADODB.Recordset:  Dim Rec_Account As ADODB.Recordset
Dim AllSaudaRec As ADODB.Recordset:     Dim RECGRID As ADODB.Recordset:     Dim Rec_Sauda As ADODB.Recordset
Sub Add_Rec()
    LDataImport = 0
    Frame1.Enabled = True: Combo1.ListIndex = 0: Frame7.Enabled = True
    Call Get_Selection(1)
    If vcDTP1.Enabled Then vcDTP1.SetFocus
    If Fb_Press > 1 Then
        Check1.Value = 0
        Check1.Enabled = False
    End If
End Sub
Sub Save_Rec()
    Dim LExCode As String:      Dim LSConSno As Long
    Dim MBAmt As Double:        Dim MSAmt As Double
    Dim LSClient As String:     Dim TRec As ADODB.Recordset
    Dim MToDate As Date:        Dim LSInstType As String
    Dim LSOptType As String:    Dim LSStrike As Double
    Dim LBuyAc_Code As String:  Dim LSellAc_Code As String
    Dim LSaudaID As Long:       Dim LExID  As Integer
    Dim LItemID As Integer
    Dim LSCondate As Date
    On Error GoTo err1
    If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If Val(TxtTotBQty.text) = 0 Then MsgBox "Please Check Entries.", vbCritical: Exit Sub
    If Val(TxtTotSQty.text) = 0 Then MsgBox "Please Check Entries.", vbCritical:  Exit Sub
    LExCode = TxtExCode.text
    LSCondate = vcDTP1.Value
    
    If GSysLockDt > LSCondate Then
        MsgBox "Can Not Add/Modfify/Delete Trades. Settlement Locked Till " & GSysLockDt & ""
        Exit Sub
    End If
    
    RECGRID.Sort = "SRNO"
    If RECGRID.RecordCount > 0 Then
        Cnn.BeginTrans
        CNNERR = True
        
        LExID = Get_ExID(LExCode)
        LItemID = Get_ITEMID(TxtItemName.text)
        LSaudaID = Get_SaudaID(TxtSaudaCode.text)
        LSConSno = Get_ConSNo(vcDTP1.Value, TxtSaudaCode.text, TxtItemName.text, LExCode, LSaudaID, LItemID, LExID)
        '
        
        Dim recparty As ADODB.Recordset
        Set recparty = Nothing
        Set recparty = New ADODB.Recordset
        mysql = "SELECT PARTY FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND SAUDAID =" & LSaudaID & ""
        recparty.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        
        While Not recparty.EOF
            If LenB(LSParties) = 0 Then
                LSParties = "'" & recparty!PARTY & "'"
            Else
                If InStr(LSParties, "'" & recparty!PARTY & "'") < 1 Then LSParties = LSParties & ",'" & recparty!PARTY & "'"
            End If
            recparty.MoveNext
        Wend
        
'        'sACHIN -- to move contrat entry in log table before edit
'        mysql = "INSERT INTO CTR_D_LOG (CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,loginuser,datetm,[tran])  "
'        mysql = mysql & "SELECT CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,'" & GUserName & "',getdate() "
'        mysql = mysql & ",'" & Fb_Press & "' "
'
'        mysql = mysql & "FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND SAUDAID =" & LSaudaID & "   "
'        If LenB(DataCombo4.BoundText) > 0 Then mysql = mysql & " AND USERID ='" & LUserId & "' "
'        If LenB(LFParties) > 0 Then mysql = mysql & "AND CONNO IN (SELECT DISTINCT Z.CONNO FROM CTR_D Z WITH(NOLOCK)  WHERE Z.COMPCODE =" & GCompCode & "  AND Z.PARTY IN (" & LFParties & ") AND Z.CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND Z.SAUDAID =" & LSaudaID & " )"
'
'        Cnn.Execute mysql
        
        
        mysql = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND SAUDAID =" & LSaudaID & ""
        If LenB(DataCombo4.BoundText) > 0 Then mysql = mysql & " AND USERID   ='" & LUserId & "'"
        If LenB(LFParties) > 0 Then mysql = mysql & " AND CONNO IN (SELECT DISTINCT CONNO FROM CTR_D WHERE COMPCODE =" & GCompCode & "  AND PARTY IN (" & LFParties & ") AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND SAUDAID =" & LSaudaID & " )"
        Cnn.Execute mysql
        
        
        LDataImport = IIf(IsNull(LDataImport), 0, LDataImport)
        If LenB(TxtItemName.text) = 0 Then
            MsgBox "Please Check Itemcode for this Contract. Entry Not Saved"
            Exit Sub
        Else
            RECGRID.MoveFirst
            MBAmt = 0:    MSAmt = 0
            Do While Not RECGRID.EOF
                If LenB(RECGRID!BNAME) > 0 And LenB(RECGRID!SNAME) > 0 Then   ''WHEN BUYER AND SELLER BOTH ARE THERE
                    If RECGRID!BQnty > 0 And RECGRID!BRate > 0 And RECGRID!SQnty > Val(0) And RECGRID!SRate > Val(0) Then                   ''QNTY AND RATE REQUIRED
                        If RECGRID!DIMPORT = 0 Then
                            LSClient = RECGRID!BCODE
                        Else
                            LSClient = (RECGRID!LCLCODE & vbNullString)
                        End If
                        LBuyAc_Code = Get_AccountDCode(RECGRID!BCODE)
                        If LenB(LBuyAc_Code) < 1 Then
                            MsgBox "Please Check Account Code for " & RECGRID!BCODE & "  of Trade No " & Val(RECGRID!SrNo) & ""
                            If CNNERR = True Then
                                CNNERR = False
                                Cnn.RollbackTrans
                                Exit Sub
                            End If
                        End If
                        LSellAc_Code = Get_AccountDCode(RECGRID!scode)
                        If LenB(LSellAc_Code) < 1 Then
                            MsgBox "Please Check Account Code for " & RECGRID!scode & "  of Trade No " & Val(RECGRID!SrNo) & ""
                            If CNNERR = True Then
                                CNNERR = False
                                Cnn.RollbackTrans
                                Exit Sub
                            End If
                        End If
                        If LenB(LSParties) = 0 Then
                            LSParties = "'" & RECGRID!BCODE & "'"
                        Else
                            If InStr(LSParties, "'" & RECGRID!BCODE & "'") < 1 Then LSParties = LSParties & ",'" & RECGRID!BCODE & "'"
                        End If
                        If InStr(LSParties, "'" & RECGRID!scode & "'") < 1 Then LSParties = LSParties & ",'" & RECGRID!scode & "'"
                        LDataImport = Abs(RECGRID!DIMPORT)
                        If IsNull(RECGRID!USERID) Then RECGRID!USERID = RECGRID!scode
                        
                        
                        Call Add_To_Ctr_D2("B", LSClient, LSConSno, vcDTP1.Value, Val(RECGRID!SrNo), TxtSaudaCode.text, TxtItemName.text, LBuyAc_Code, Val(RECGRID!BQnty), Val(RECGRID!BRate), Val(RECGRID!SRate), LSellAc_Code, RECGRID!contime, RECGRID!ORDER_NO, Trim$(RECGRID!USERID), RECGRID!TRADE_NO, LExCode, FCalval, LDataImport, RECGRID!ORDTIME, TxtInst.text, TxtOptType.text, Val(TxtStrike.text), IIf(IsNull(RECGRID!filetype), 0, RECGRID!filetype), RECGRID!BROKFLAG, LExID, LItemID, LSaudaID)
                        
                    End If
                End If
                RECGRID.MoveNext
            Loop
            
            Call Shree_Posting(DateValue(vcDTP1.Value))
            If ChkCheckRate.Value = 1 Then Call RATE_TEST(vcDTP1.Value, , str(LSaudaID), GETCont)
            
            LExID = Get_ExID(LExCode)
            LExCode = "'" & LExCode & "'"
            Call Update_Charges(LSParties, str(LExID), str(LSaudaID), Trim(str(LItemID)), vcDTP1.Value, vcDTP1.Value, True)
            Cnn.CommitTrans
            CNNERR = False
            Cnn.BeginTrans
            CNNERR = False
            If BILL_GENERATION(vcDTP1.Value, GFinEnd, str(LSaudaID), LSParties, str(LExID)) Then
                Cnn.CommitTrans: CNNERR = False
            Else
                Cnn.RollbackTrans: CNNERR = False
            End If
            'Call Chk_Billing
        End If
    End If
    Call CANCEL_REC
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number

    If CNNERR = True Then
        Cnn.RollbackTrans: CNNERR = False
    End If
End Sub
Sub CANCEL_REC()
    vcDTP1.Enabled = True: TxtSaudaCode.Enabled = True: DataCombo1.Enabled = True: Combo1.Enabled = True: DataCombo4.Enabled = True: TxtSettleRate.Enabled = True: TxtRefLot.Enabled = True
    Check1.Enabled = True
    Call RecSet
    Fb_Press = 0
    Combo1.ListIndex = 0
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.Refresh
    Frame9.Enabled = False
    Label2.Visible = False
    DataCombo3.Visible = False
    Check1.Enabled = True
    Call CLEAR_ITEM
    'Call ClearFormFn(GETCont)
    Call Get_Selection(10)
    Frame1.Enabled = False
End Sub
Function MODIFY_REC(LMCondate As Date, LMSauda As String, LMPattan As String) As Boolean
    Dim LCItemCode  As String:      Dim TRec As ADODB.Recordset:    Dim LCSaudacode As String
    Dim LDItemCode As String:       Dim LDExCode As String:         Dim MToDate As Date
    Dim LSConSno As Long:           Dim LConNo As Long:             Dim LSaudaID As Long
    Dim LExID As Integer
    Dim LItemID As Integer
    LSParties = vbNullString
    LSConSno = 0
    If LenB(LMSauda) < 1 Then Exit Function
    If Check1.Value = 1 Then
        LCItemCode = vbNullString:        LCSaudacode = vbNullString
        If AllSaudaRec.RecordCount > 0 Then AllSaudaRec.MoveFirst
        AllSaudaRec.Find "SAUDACODE='" & LMSauda & "'"
        If AllSaudaRec.EOF Then
            MsgBox "Invalid Sauda Code"
            MODIFY_REC = False
            Exit Function
        End If
        TxtRefLot.text = Format(AllSaudaRec!REFLOT, "0.00")
        LExID = Get_ExID(AllSaudaRec!excode)
        LCItemCode = Get_ItemMaster(LExID, AllSaudaRec!EX_SYMBOL)
        If LenB(LCItemCode) < 1 Then LCItemCode = Create_TItemMast(AllSaudaRec!ITEMCODE, AllSaudaRec!ITEMCODE, AllSaudaRec!EX_SYMBOL, AllSaudaRec!lot, AllSaudaRec!excode)
        LItemID = Get_ITEMID(LCItemCode)
        LCSaudacode = Get_SaudaMaster(LExID, LItemID, AllSaudaRec!MATURITY, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
        If LenB(LCSaudacode) < 1 Then LCSaudacode = Create_TSaudaMast(LCItemCode, AllSaudaRec!MATURITY, AllSaudaRec!excode, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
        LMSauda = LCSaudacode
        
    End If
    If GQty_Decimal = True Then
        DataGrid1.Columns(3).NumberFormat = "0.00"
        DataGrid1.Columns(7).NumberFormat = "0.00"
    End If
    vcDTP1.Value = LMCondate
    
    Set TRec = Nothing:    Set TRec = New ADODB.Recordset
   '
    mysql = "SELECT EX.EXID,IT.ITEMCODE,IT.ITEMID,IT.LOT,SD.INSTTYPE,SD.OPTTYPE,SD.STRIKEPRICE,SD.SAUDAID,SD.SAUDACODE,"
    mysql = mysql & " SD.TRADEABLELOT,EX.EXCODE,EX.LOTWISE,SD.REFLOT FROM ITEMMAST AS IT,SAUDAMAST AS SD,EXMAST AS EX "
    mysql = mysql & " WHERE IT.COMPCODE =" & GCompCode & " AND IT.COMPCODE = SD.COMPCODE AND IT.ITEMCODE=SD.ITEMCODE "
    mysql = mysql & " AND EX.COMPCODE =IT.COMPCODE AND IT.EXID =EX.EXID "
    mysql = mysql & " AND SD.SAUDACODE='" & LMSauda & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not TRec.EOF Then
        TxtSaudaCode.text = TRec!saudacode
        DataCombo1.BoundText = TRec!saudacode
        TxtItemName.text = TRec!ITEMCODE
        TxtExCode.text = TRec!excode
        TxtInst.text = TRec!INSTTYPE
        TxtOptType.text = TRec!OPTTYPE
        TxtStrike.text = Format(TRec!STRIKEPRICE, "0.00")
        TxtRefLot.text = Format(TRec!REFLOT, "0.00")
        LSaudaID = TRec!SAUDAID
        LItemID = TRec!itemid
        TxtItemID.text = TRec!itemid
        TxtExID.text = TRec!EXID
        LExID = TRec!EXID
        FCalval = Get_LotSize(LItemID, LSaudaID, LExID)
    Else
        MsgBox "Invalid Sauda Please Try Again"
        MODIFY_REC = False
        DataCombo1.SetFocus
        Get_Selection (0)
        Exit Function
    End If
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT CONSNO FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LMCondate, "yyyy/MM/dd") & "' AND SAUDA='" & LMSauda & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        If Fb_Press = 2 Then
            MsgBox "Transaction does not exists for the Selected creteria?", vbExclamation
            If AddPerm = True Then
                GETCont.Fb_Press = 1
                MODIFY_REC = True
            Else
                MODIFY_REC = False
                Get_Selection (0)
                Exit Function
            End If
        ElseIf Fb_Press = 1 Then
            MODIFY_REC = True
        End If
        Call Grid_Entry
        Exit Function
    Else
        LSConSno = TRec!CONSNO
        If Fb_Press = 1 Then
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            mysql = "SELECT TOP 1 CONNO FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LMCondate, "yyyy/MM/dd") & "' AND SAUDA = '" & LMSauda & "'"
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                MsgBox "Contract already Exists.Please Press Enter to Modify Contract.", vbInformation
                GETCont.Fb_Press = 2
                GETMAIN.StatusBar1.Panels(2).text = "Modify Record"
                MODIFY_REC = True
            Else
                MODIFY_REC = True
                Call Grid_Entry
                Exit Function
            End If
         Else
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            mysql = "SELECT TOP 1 CONNO FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LMCondate, "yyyy/MM/dd") & "' AND SAUDA='" & LMSauda & "' "
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                MODIFY_REC = True
            Else
                MsgBox "Contract does not exists.Please add New Contract.", vbInformation
                GETCont.Fb_Press = 1
                GETCont.Add_Rec
                GETMAIN.StatusBar1.Panels(2).text = "Add Record"
                MODIFY_REC = False:                   Get_Selection (0)
                Exit Function
            End If
        End If
        LDataImport = 0
    End If
    Set TRec = Nothing:    Set TRec = New ADODB.Recordset
    mysql = "SELECT C.CONNO,C.PARTY,C.USERID,C.CLCODE,C.CONCODE,C.QTY,C.CONFIRM,C.RATE,C.DATAIMPORT,C.CONTIME,C.ORDNO,C.INVNO,C.ROWNO1,C.CONTYPE,A.NAME,C.ROWNO,C.ORDTIME,BROKFLAG,C.FILETYPE "
    mysql = mysql & " FROM CTR_D AS C,ACCOUNTD AS A WHERE C.COMPCODE =" & GCompCode & " AND C.COMPCODE =A.COMPCODE AND C.PARTY=A.AC_CODE AND C.CONDATE ='" & Format(LMCondate, "YYYY/MM/DD") & "' AND C.SAUDA='" & TxtSaudaCode.text & "'"
    If Combo1.ListIndex = 0 Then
        mysql = mysql & " AND C.PATTAN='C'"
    Else
        mysql = mysql & " AND C.PATTAN='O'"
    End If
    If LenB(LFParties) <> 0 Then mysql = mysql & " AND C.CONNO IN (SELECT DISTINCT CONNO FROM CTR_D WHERE COMPCODE =" & GCompCode & "  AND CONDATE ='" & Format(LMCondate, "YYYY/MM/DD") & "' AND SAUDA='" & TxtSaudaCode.text & "' AND PARTY IN (" & LFParties & "))"
    mysql = mysql & " ORDER BY C.ROWNO,C.CONNO,C.CONTYPE"
    
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    Call RecSet
    RECGRID.Delete
    With TRec
        Do While Not .EOF
            LConNo = !CONNO
            If (Trim(!USERID) = Trim(DataCombo4.BoundText)) Or DataCombo4.text = vbNullString Then
                RECGRID.AddNew:                 RECGRID!SrNo = LConNo
                RECGRID!BCODE = !PARTY:         RECGRID!LCLCODE = (!CLCODE & vbNullString)
                RECGRID!BNAME = !NAME:          RECGRID!CONCODE = IIf(IsNull(!CONCODE), vbNullString, !CONCODE)
                RECGRID!BQnty = !QTY:           RECGRID!bconfirm = !CONFIRM
                RECGRID!BRate = !Rate:          RECGRID!LInvNo = Val(!invno)
                RECGRID!filetype = !filetype
                RECGRID!ORDER_NO = !ORDNO & vbNullString: RECGRID!contime = IIf(IsNull(!contime), Time, !contime)
                RECGRID!USERID = !USERID & vbNullString:  RECGRID!ORDTIME = !ORDTIME & vbNullString
                RECGRID!TRADE_NO = IIf(IsNull(!ROWNO1), Trim(str(LConNo)), !ROWNO1)
                RECGRID!BROKFLAG = !BROKFLAG
                If Not IsNull(!DATAIMPORT) Then
                    If !DATAIMPORT = True Then
                        RECGRID!DIMPORT = 1
                    Else
                        RECGRID!DIMPORT = 0
                    End If
                Else
                    RECGRID!DIMPORT = 0
                End If
                If InStr(LSParties, "'" & !PARTY & "'") < 1 Then
                    If LenB(LSParties) <> 0 Then LSParties = LSParties & ","
                    LSParties = LSParties & "'" & !PARTY & "'"
                End If
                .MoveNext
                If .EOF Then
                    MsgBox "Seller Party missing.", vbInformation
                    RECGRID.Update
                    Exit Do
                Else
                    RECGRID!scode = !PARTY:                    RECGRID!RCLCODE = (!CLCODE & vbNullString)
                    RECGRID!SNAME = !NAME:                     RECGRID!SQnty = !QTY
                    RECGRID!SCONFIRM = !CONFIRM:               RECGRID!SRate = !Rate
                    RECGRID!RInvNo = Val(!invno)
                    If InStr(LSParties, "'" & !PARTY & "'") < 1 Then
                        If LenB(LSParties) <> 0 Then LSParties = LSParties & ", "
                        LSParties = LSParties & "'" & !PARTY & "'"
                        LSParties = "'" & !PARTY & "'"
                    End If
                End If
                RECGRID!DIFFAMT = (Val((RECGRID!BQnty * RECGRID!BRate)) - Val((RECGRID!SQnty * RECGRID!SRate))) * FCalval
                RECGRID!OLDTRD = "Y"
                RECGRID.Update
            End If
            .MoveNext
        Loop
    End With
    LSConno = Get_Max_ConNo(LMCondate, 0)
    Set DataGrid1.DataSource = RECGRID
    Call DataGrid1_AfterColEdit(0)
    Dim LSCondate     As Date
    LSCondate = vcDTP1.Value
    If Fb_Press = 3 Then
        If GSysLockDt > LSCondate Then
            MsgBox "Can Not Add/Modfify/Delete Trades. Settlement Locked Till " & GSysLockDt & ""
            Exit Function
        End If
    
        If MsgBox("You are about to Delete all Contracts. Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
            On Error GoTo err1
            Cnn.BeginTrans
            CNNERR = True
            mysql = "DELETE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & LSConSno & ""
            Cnn.Execute mysql
            Cnn.CommitTrans
            CNNERR = False
            
            Cnn.BeginTrans
            CNNERR = True
            Rec_Sauda.MoveFirst
            Rec_Sauda.Find "SAUDACODE ='" & TxtSaudaCode.text & "'"
            If Rec_Sauda.EOF Then
                MsgBox "Invalid Sauda"
                Exit Function
            Else
                MToDate = DateValue(Rec_Sauda!MATURITY)
                LDItemCode = "'" & Rec_Sauda!ITEMCODE & "'"
                LDExCode = "'" & Rec_Sauda!excode & "'"
                LExID = Rec_Sauda!EXID
                LSaudaID = Rec_Sauda!SAUDAID
                LItemID = Rec_Sauda!itemid
            End If
            Set TRec = Nothing
            Call Update_Charges(vbNullString, str(LExID), str(LSaudaID), Trim(str(LItemID)), vcDTP1.Value, vcDTP1.Value, True)
            If BILL_GENERATION(LMCondate, MToDate, str(LSaudaID), vbNullString, str(LExID)) Then
                Cnn.CommitTrans
                CNNERR = False
            Else
                Cnn.RollbackTrans
                CNNERR = False
            End If
            'Call Chk_Billing
            ''REGENERATING UPTO HERE
            If Fb_Press = 2 Then
                GETMAIN.Toolbar1_Buttons(4).Enabled = True: GETMAIN.saverec.Enabled = True
            ElseIf Fb_Press = 3 Then
                Call CANCEL_REC
            End If
            MODIFY_REC = True
            Exit Function
err1:
            If err.Number <> 0 Then
                MsgBox err.Description, vbCritical, "Error Number : " & err.Description
                
            End If
            If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
            Call CANCEL_REC
        End If
    End If
    Call Grid_Entry
End Function


Private Sub Check2_Click()

End Sub

Private Sub Combo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    Dim TRec As ADODB.Recordset
    If Combo1.ListIndex = 0 Then
        TxtSettleRate.SetFocus
    Else
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open "SELECT TOP 1 CONDATE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PATTAN='O' AND SAUDA='" & TxtSaudaCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            If Format(vcDTP1.Value, "yyyy/MM/dd") < TRec!Condate Then
                MsgBox "Opening for this contract  has been already entered on " & Format(TRec!Condate, "yyyy/MM/dd"), vbExclamation, "Warning"
                vcDTP1.Value = Date
                Cancel = True
                Exit Sub
            End If
        End If
        DataGrid1.Columns(7).Locked = True
    End If
End Sub

Private Sub Command1_Click()
    Dim TRec As ADODB.Recordset
    LFParties = vbNullString
    If LenB(TxtSaudaCode.text) <> 0 Then
        mysql = "SELECT DISTINCT A.PARTY,B.NAME FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
        mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE AND SAUDA='" & TxtSaudaCode.text & "'"
        mysql = mysql & "  AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' "
        mysql = mysql & " ORDER BY B.NAME "
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not TRec.EOF Then
            ListView1.ListItems.Clear
            ListView1.Enabled = True
            Do While Not TRec.EOF
                ListView1.ListItems.Add , , TRec!PARTY
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , TRec!NAME
                TRec.MoveNext
            Loop
        End If
        Frame11.Left = 8400
        Frame11.Top = 1440
        Frame11.Visible = True
    End If
End Sub
Private Sub Command2_Click()
LFParties = Get_Parties
Frame11.Visible = False
Call TxtSettleRate_Validate(False)
End Sub

Private Sub Command3_Click()
Frame13.Visible = False
End Sub

Private Sub DataCombo1_GotFocus()
If Fb_Press <> 1 Then Check1.Value = 0
If Check1.Value = 1 Then
    If Not AllSaudaRec.EOF Then
        Set DataCombo1.RowSource = AllSaudaRec
        DataCombo1.BoundColumn = "SAUDACODE"
        DataCombo1.ListField = "SAUDANAME"
    Else
        MsgBox "No Contracts in Scriptmaster. Please Import Scriptmaster"
    End If
Else
    If Not Rec_Sauda.EOF Then
        Set DataCombo1.RowSource = Rec_Sauda
        DataCombo1.BoundColumn = "SAUDACODE"
        DataCombo1.ListField = "SAUDANAME"
    End If
End If
Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    Dim TRec As ADODB.Recordset:    Dim LSaudaID As Long
    Dim LItemCode As String:    Dim LExID  As Integer
    Dim LItemID  As Integer
    If Check1.Value = 1 Then
        If AllSaudaRec.RecordCount > 0 Then AllSaudaRec.MoveFirst
        AllSaudaRec.Find "SAUDACODE='" & DataCombo1.BoundText & "'", , adSearchForward
        If AllSaudaRec.EOF Then
            Cancel = True
            MsgBox "Please Select Contract"
            Exit Sub
        Else
            LFParties = vbNullString
            TxtSaudaCode.text = AllSaudaRec!saudacode
            TxtItemName.text = AllSaudaRec!ITEMCODE
            TxtInst.text = AllSaudaRec!INSTTYPE
            TxtOptType.text = AllSaudaRec!OPTTYPE
            TxtStrike.text = Format(AllSaudaRec!STRIKEPRICE, "0.00")
            If LenB(TxtItemName.text) = 0 Then
                MsgBox "Please Select Commodity for this Contract"
                Cancel = True
            End If
            TxtExCode.text = AllSaudaRec!excode
            LSaudaID = Get_SaudaID(TxtSaudaCode.text)
            TxtSaudaID.text = LSaudaID
            LExID = Get_ExID(TxtExCode.text)
            TxtRefLot = Format(AllSaudaRec!REFLOT, "0.00")
            LItemCode = AllSaudaRec!ITEMCODE
        End If
    Else
        If Rec_Sauda.RecordCount > 0 Then Rec_Sauda.MoveFirst
        
        Rec_Sauda.Find "SAUDACODE='" & DataCombo1.BoundText & "'"
        If Rec_Sauda.EOF Then
            MsgBox "No Contracts Exist"
            Cancel = True
            Exit Sub
        Else
            LFParties = vbNullString
            TxtSaudaCode.text = Rec_Sauda!saudacode
            TxtItemName.text = Rec_Sauda!ITEMCODE
            If LenB(TxtItemName.text) = 0 Then
                MsgBox "Please Select Commodity for this Contract"
                Cancel = True
            End If
            TxtExCode.text = Rec_Sauda!excode
            FCalval = Rec_Sauda!lot
            TxtInst.text = Rec_Sauda!INSTTYPE
            TxtOptType.text = Rec_Sauda!OPTTYPE
            TxtStrike.text = Format(Rec_Sauda!STRIKEPRICE, "0.00")
            LSaudaID = Rec_Sauda!SAUDAID
            TxtSaudaID.text = LSaudaID
            LExID = Rec_Sauda!EXID
            LItemCode = Rec_Sauda!ITEMCODE
            
            TxtRefLot = Format(Rec_Sauda!REFLOT, "0.00")
        End If
        LItemID = Get_ITEMID(LItemCode)
        FCalval = Get_LotSize(LItemID, LSaudaID, LExID)
        Set LUserIDRec = Nothing:        Set LUserIDRec = New ADODB.Recordset
        mysql = "SELECT DISTINCT USERID FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
        mysql = mysql & " AND SAUDA='" & Rec_Sauda!saudacode & "'AND USERID IS NOT NULL AND USERID <>'' ORDER BY USERID"
        LUserIDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not LUserIDRec.EOF Then
            Set DataCombo4.RowSource = LUserIDRec
            DataCombo4.ListField = "USERID"
            DataCombo4.BoundColumn = "USERID"
        End If
        Call GetCloseRate
    End If
    Combo1.SetFocus
    LSConno = Get_Max_ConNo(vcDTP1.Value, 0)
    LSConno = LSConno + 1
    RECGRID!SrNo = LSConno  'RECGRID.AbsolutePosition
    RECGRID.Update
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
End Sub
Private Sub DataCombo3_GotFocus()
    Sendkeys "%{DOWN}"
    If DataGrid1.Col = 1 Or DataGrid1.Col = 2 Then
        DataGrid1.text = vbNullString
        DataGrid1.Col = 1
        Label2.Visible = True: Label2.Left = 1080
        DataCombo3.Left = Val(1080)
        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    ElseIf DataGrid1.Col = 5 Or DataGrid1.Col = 6 Or DataGrid1.Col = 7 Then
        DataGrid1.text = vbNullString
        DataGrid1.Col = 5: DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
        DataCombo3.Left = Val(7200)
        Label2.Visible = True: Label2.Left = 7200
    End If
    Sendkeys "%{DOWN}"
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim LAcCode As String
Dim LAc_Name As String
    If KeyCode = 13 Then
        If DataCombo3.BoundText <> "" Then
            LAcCode = DataCombo3.BoundText
            LAc_Name = DataCombo3.text
            If InStr(LAcCode, "'") Then LAcCode = Replace(LAcCode, "'", "", 1, Len(LAcCode))
            Rec_Account.Filter = adFilterNone
            Rec_Account.Filter = "AC_CODE ='" & LAcCode & "'"
            If Rec_Account.EOF Then
                DataCombo3.BoundText = ""
                Sendkeys "%{DOWN}"
            Else
                Rec_Account.Filter = adFilterNone
                If DataGrid1.Col = 1 Then
                    RECGRID!BCODE = LAcCode
                    RECGRID!BNAME = LAc_Name
                    If RECGRID!USERID = "" Then RECGRID!USERID = LUserId
                    DataGrid1.Col = 2
                ElseIf DataGrid1.Col = 5 Or DataGrid1.Col = 6 Or DataGrid1.Col = 7 Then
                    RECGRID!scode = LAcCode
                    RECGRID!SNAME = LAc_Name
                    If RECGRID!USERID = "" Then RECGRID!USERID = LUserId
                    DataGrid1.Col = 6
                End If
                DataGrid1.SetFocus
                DataCombo3.Visible = False: Label2.Visible = False
            End If
        End If
    ElseIf KeyCode = 27 Then
        DataGrid1.SetFocus
        DataCombo3.Visible = False: Label2.Visible = False
    ElseIf KeyCode = 121 Then   'F3  NEW PARTY
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
    If Len(Trim(DataCombo4.BoundText)) > 0 Then LUserId = DataCombo4.BoundText
End Sub
Public Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    Dim LTotBQnty As Double:    Dim LTotSQnty As Double:    Dim LTotBAmt As Double:     Dim LTotSamt As Double
    Dim LBAmt As Double:        Dim LSAmt As Double:        Dim LDiffAmt As Double:      Dim TRec As ADODB.Recordset
    
    If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0
    If ColIndex = 1 Or ColIndex = 5 Then
        Rec_Account.MoveFirst
        Rec_Account.Find "AC_CODE='" & DataGrid1.text & "'", , adSearchForward
        If Not Rec_Account.EOF Then
            If ColIndex = Val(1) Then
                DataGrid1.Col = 2
                RECGRID!BCODE = Rec_Account!AC_CODE
                RECGRID!BNAME = Rec_Account!NAME
                If RECGRID!USERID = "" Then RECGRID!USERID = LUserId
            Else
                If Combo1.ListIndex = Val(0) Then
                    DataGrid1.Col = 7
                Else
                    DataGrid1.Col = 6
                End If
                RECGRID!scode = Rec_Account!AC_CODE
                RECGRID!SNAME = Rec_Account!NAME
                If RECGRID!USERID = "" Then RECGRID!USERID = LUserId
            End If
        Else
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        End If
    ElseIf ColIndex = 3 Or ColIndex = 4 Then
        ''IF CONTRACT THEN ONLY CHANGE OCCURS
        RECGRID!SQnty = RECGRID!BQnty
        If Val(RECGRID!BRate) > 0 Then
            If Val(RECGRID!SRate) = 0 Then RECGRID!SRate = RECGRID!BRate
        Else
            If ColIndex = 3 Then
            Else
                MsgBox "Rate can not be zero.Please enter rate.", vbCritical
                DataGrid1.Col = 4: DataGrid1.SetFocus
            End If
        End If
    End If
    LBAmt = 0:    LSAmt = 0:    LDiffAmt = 0
    If RECGRID.RecordCount > 0 Then
        RECGRID!BQnty = IIf(IsNull(RECGRID!BQnty), 0, RECGRID!BQnty)
        RECGRID!BRate = IIf(IsNull(RECGRID!BRate), 0, RECGRID!BRate)
        RECGRID!SQnty = IIf(IsNull(RECGRID!SQnty), 0, RECGRID!SQnty)
        RECGRID!SRate = IIf(IsNull(RECGRID!SRate), 0, RECGRID!SRate)
        If Val(RECGRID!BQnty & vbNullString) > 0 And Val(RECGRID!SQnty & vbNullString) Then
            LBAmt = Val(RECGRID!BQnty) * Val(RECGRID!BRate) * FCalval
            LSAmt = Val(RECGRID!SQnty * RECGRID!SRate) * FCalval
            LDiffAmt = LSAmt - LBAmt
            RECGRID!DIFFAMT = Format(LDiffAmt, "0.00")
        End If
    End If
    Call Calc_Totals
End Sub
Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If Fb_Press = 2 Then
        If Val(DataGrid1.Columns(9).text) > 0 And (ColIndex = 1 Or ColIndex = 2 Or ColIndex = 3 Or ColIndex = 4) Then
            MsgBox "Invoice already Generated.", vbCritical: Cancel = 1: DataGrid1.Col = ColIndex: DataGrid1.SetFocus: Exit Sub
        ElseIf Val(DataGrid1.Columns(10).text) > 0 And (ColIndex = 5 Or ColIndex = 6 Or ColIndex = 7 Or ColIndex = 8) Then
            MsgBox "Invoice already generated.", vbCritical: Cancel = 1: DataGrid1.Col = ColIndex: DataGrid1.SetFocus: Exit Sub
        End If
 
If ModiPerm = False And RECGRID!OLDTRD = "Y" Then
            MsgBox "No Modify Rights Available", vbCritical: Cancel = 1: DataGrid1.Col = ColIndex: DataGrid1.SetFocus: Exit Sub
        End If
    End If
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LBuyerCode As String: Dim LBuyerName As String
    Dim LSellerCode As String: Dim LSellerName As String
    Dim LRowNo As Integer
    
    If KeyCode = 13 And (DataGrid1.Col = 8 Or DataGrid1.Col = 18) Then
         If DataGrid1.Col = 8 Then
            If Val(DataGrid1.text) = 0 Then
                MsgBox "Rate can not be zero.Please Enter rate.", vbCritical
                DataGrid1.Col = 7: DataGrid1.SetFocus
            End If
        End If
        If Val(DataGrid1.Columns(8).text) <> 0 Then
            LBuyerCode = RECGRID!BCODE
            LBuyerName = RECGRID!BNAME
            LSellerCode = RECGRID!scode
            LSellerName = RECGRID!SNAME
            Rec_Account.MoveFirst
            Rec_Account.Find "AC_CODE='" & LBuyerCode & "'", , adSearchForward
            If Rec_Account.EOF Then
                MsgBox "Please Enter Valid Buyer Code"
                DataGrid1.Col = 0
                DataGrid1.SetFocus
                Exit Sub
            End If
            Rec_Account.MoveFirst
            Rec_Account.Find "AC_CODE='" & LSellerCode & "'", , adSearchForward
            If Rec_Account.EOF Then
                MsgBox "Please Enter Valid Seller Code"
                DataGrid1.Col = 4
                DataGrid1.SetFocus
                Exit Sub
            End If
            
            RECGRID.MoveNext
            If RECGRID.EOF Then
                RECGRID.AddNew
                If Combo1.ListIndex = Val(1) Then   ''OPENING
                    RECGRID!BRate = Val(TxtSettleRate.text)
                    RECGRID!SRate = Val(TxtSettleRate.text)
                Else                        ''LAST INFORMATION CARIES
                    RECGRID!BCODE = LBuyerCode
                    RECGRID!BNAME = LBuyerName
                    RECGRID!scode = LSellerCode
                    RECGRID!SNAME = LSellerName
                    RECGRID!BRate = 0
                    RECGRID!SRate = 0
                End If
                RECGRID!BROKFLAG = "Y"
                RECGRID!BQnty = 0
                RECGRID!SQnty = 0
                RECGRID!DIMPORT = 0
                RECGRID!bconfirm = 0
                RECGRID!SCONFIRM = 0
                If LenB(RECGRID!USERID) < 1 Then RECGRID!USERID = LUserId & vbNullString
                LSConno = LSConno + 1
                RECGRID!SrNo = LSConno 'RECGRID.AbsolutePosition
                RECGRID!OLDTRD = "N"
                RECGRID!ORDER_NO = CStr(LSConno)
                RECGRID!TRADE_NO = CStr(LSConno)
                RECGRID!ORDTIME = CStr(Date) & " " & CStr(Time)
                RECGRID!contime = CStr(Time)
                RECGRID.Update
            End If
            DataGrid1.LeftCol = 0
            DataGrid1.Col = 0
        End If
    ElseIf DataGrid1.Col = 4 And KeyCode = 13 Then
            If Val(DataGrid1.text) = 0 Then
                MsgBox "Rate can not be zero.Please Enter rate.", vbCritical
                DataGrid1.Col = 3: DataGrid1.SetFocus
            End If
        
    ElseIf KeyCode = 114 Then   'F3  NEW PARTY
        GETACNT.Show
        GETACNT.ZOrder
        GETACNT.add_record
    ElseIf KeyCode = 118 Then   ''F7 KEY
       ' LRowNo = InputBox("Enter the row number.", "Sauda")
       ' If Val(LRowNo) > Val(0) Then
        '    RECGRID.MoveFirst
        '    RECGRID.Find "SRNO=" & Val(LRowNo) & "", , adSearchForward
        '    If RECGRID.EOF Then
        '        MsgBox "Record not found.", vbCritical, "Error"
        '        RECGRID.MoveFirst
         '   End If
            DataGrid1.Col = 1
            DataGrid1.SetFocus
        'End If
    ElseIf KeyCode = 46 And Shift = 2 Then
        RECGRID.Delete
        If RECGRID.RecordCount = 0 Then
            RECGRID.AddNew
            LSConno = LSConno + 1
            RECGRID!SrNo = LSConno 'RECGRID.RecordCount
            If Combo1.ListIndex = Val(1) Then
                RECGRID!BRate = Val(TxtSettleRate.text)
                RECGRID!SRate = Val(TxtSettleRate.text)
                If RECGRID!USERID = "" Then RECGRID!USERID = LUserId
            End If
            RECGRID!BROKFLAG = "Y"
            RECGRID!BQnty = 0:            RECGRID!SQnty = 0
            RECGRID!BRate = 0:            RECGRID!SRate = 0
            RECGRID!bconfirm = 0:         RECGRID!SCONFIRM = 0
            RECGRID!OLDTRD = "N":         RECGRID!ORDER_NO = CStr(LSConno)
            RECGRID!TRADE_NO = CStr(LSConno)
            RECGRID!ORDTIME = CStr(Date) & " " & CStr(Time)
            RECGRID!contime = CStr(Time)
            RECGRID.Update
        End If
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
    Call Calc_Totals
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
    vcDTP1.Value = Date
    Call CANCEL_REC
    Set DataGrid1.DataSource = RECGRID: DataGrid1.Refresh
    Frame1.Enabled = False
    LDataImport = 0
    DataCombo3.Top = Val(2055): DataCombo3.Left = Val(1080)
    Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
    Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Account.EOF Then Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    Set Rec_Account = Nothing
    Set Rec_Account = New ADODB.Recordset
    Rec_Account.Open "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND GCODE IN (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Account.EOF Then
        Set DataCombo3.RowSource = Rec_Account
        DataCombo3.BoundColumn = "AC_CODE"
        DataCombo3.ListField = "NAME"
    Else
        MsgBox "Please create customer account", vbInformation
        Call Get_Selection(12)
    End If
    If Fb_Press > 0 Then Call Get_Selection(Fb_Press)
End Sub


Private Sub TxtRefLot_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtRefLot_Validate(Cancel As Boolean)
    TxtRefLot.text = Format(TxtRefLot.text, "0.00")
End Sub
Private Sub TxtSaudaCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then FrmSauda.Show
End Sub
Private Sub TxtSaudaCode_Validate(Cancel As Boolean)
    'FLOWDIR = 1
    If Len(Trim(TxtSaudaCode.text)) < 1 Then
        DataCombo1.SetFocus
    Else
        If Not GetCloseRate Then TxtSaudaCode.text = vbNullString: DataCombo1.SetFocus
    End If
End Sub
Private Sub TxtSettleRate_GotFocus()
    'FLOWDIR = 0:
    TxtSettleRate.SelLength = Len(TxtSettleRate.text)
End Sub
Private Sub TxtSettleRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "BNAME", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "BQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "SNAME", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "SQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "RInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "DImport", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "CONTIME", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "UserId", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "LCLCODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "RCLCODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "ORDER_NO", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "TRADE_NO", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "DiffAmt", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BConfirm", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "SConfirm", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "CONCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "OLDTRD", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "OrdTime", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKFLAG", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "FILETYPE", adDouble, , adFldIsNullable
    
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    RECGRID.AddNew
    RECGRID!DIMPORT = 0
    RECGRID!contime = Time
    RECGRID!USERID = LUserId
    RECGRID!BQnty = 0
    RECGRID!SQnty = 0
    RECGRID!BRate = 0
    RECGRID!SRate = 0
    RECGRID!bconfirm = 0
    RECGRID!BROKFLAG = "Y"
    RECGRID!SCONFIRM = 0
    RECGRID!OLDTRD = "N"
    RECGRID!filetype = 0
    LSConno = Get_Max_ConNo(vcDTP1.Value, 0)
    LSConno = LSConno + 1
    RECGRID!SrNo = LSConno  'RECGRID.AbsolutePosition
    RECGRID!ORDER_NO = CStr(LSConno)
    RECGRID!TRADE_NO = CStr(LSConno)
    RECGRID!ORDTIME = CStr(Date) & " " & CStr(Time)
    RECGRID!contime = CStr(Time)
    RECGRID.Update
    DataGrid1.Col = 1
End Sub
Private Sub TxtSettleRate_Validate(Cancel As Boolean)
    TxtSettleRate.text = Format(TxtSettleRate.text, "0.0000")
    If MODIFY_REC(vcDTP1.Value, TxtSaudaCode.text, Combo1.text) Then
    Else
        Cancel = True
    End If
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
    Case 16
        If OLDVAL = -1 Then
            RECGRID.Sort = "ORDER_NO DESC"
        Else
            RECGRID.Sort = "ORDER_NO"
        End If
    
    Case 17
        If OLDVAL = -1 Then
            RECGRID.Sort = "TRADE_NO DESC"
        Else
            RECGRID.Sort = "TRADE_NO"
        End If
    
    Case 12
        If OLDVAL = -1 Then
            RECGRID.Sort = "CONTIME DESC"
        Else
            RECGRID.Sort = "CONTIME"
        End If
    
    Case 13
        If OLDVAL = -1 Then
            RECGRID.Sort = "DiffAmt DESC"
        Else
            RECGRID.Sort = "DiffAmt"
        End If
    Case 14
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
    Dim LFExCode As String
    LFExCode = vbNullString
    Set AllSaudaRec = Nothing
    Set AllSaudaRec = New ADODB.Recordset
    mysql = " EXEC Get_ScriptContract " & GCompCode & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "','" & LFExCode & "'"
    
    AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    Set Rec_Sauda = Nothing
    Set Rec_Sauda = New ADODB.Recordset
    mysql = "EXEC Get_SaudaContract " & GCompCode & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "',0"
    'MYSQL = "EXEC Get_SaudaContract " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ",'" & InstCombo.text & "'"
    Rec_Sauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Sauda.EOF Then
        Set DataCombo1.RowSource = Rec_Sauda
        DataCombo1.BoundColumn = "SAUDACODE"
        DataCombo1.ListField = "SAUDANAME"
    End If
    If SYSTEMLOCK(DateValue(vcDTP1.Value)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
        Cancel = True
    End If
    If GQty_Decimal = True Then
        DataGrid1.Columns(3).NumberFormat = "0.00"
        DataGrid1.Columns(7).NumberFormat = "0.00"
    End If
    Combo1.SetFocus
    
    'If Fb_Press = 1 Then
    'Check1.SetFocus
End Sub
Function GetCloseRate() As Boolean
     GetCloseRate = True
     Dim LSaudaID As Long
     Dim LCLRATE As Double
     Dim TRec  As ADODB.Recordset
     LSaudaID = 0
     Set TRec = Nothing
     Set TRec = New ADODB.Recordset
     'MYSQL = "SELECT SAUDAID FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDACODE='" & TxtSaudaCode.text & "'"
     'TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
     'If TRec.EOF Then LSaudaID = TRec!SAUDAID
     LSaudaID = Get_SaudaID(TxtSaudaCode.text)
     LCLRATE = SDCLRATE(LSaudaID, vcDTP1.Value, "C")
     TxtSettleRate.text = Format(LCLRATE, "0.0000")
     
End Function
Private Sub Grid_Entry()
    Check1.Enabled = False:    vcDTP1.Enabled = False:      TxtSaudaCode.Enabled = False:    DataCombo1.Enabled = False
    Combo1.Enabled = False:    DataCombo4.Enabled = False:  TxtSettleRate.Enabled = False:    TxtRefLot.Enabled = False
    Frame7.Enabled = False:    Frame1.Enabled = True:       Frame9.Enabled = True:    Call GetCloseRate
    DoEvents:                  DataGrid1.SetFocus
End Sub
Public Function Get_Parties() As String
Dim LFParty_Codes As String
Dim I As Integer
LFParty_Codes = vbNullString
For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(I).Checked = True Then
        If LFParty_Codes <> "" Then LFParty_Codes = LFParty_Codes & ", "
        LFParty_Codes = LFParty_Codes & "'" & ListView1.ListItems(I) & "'"
    End If
Next
Get_Parties = LFParty_Codes
End Function

Private Sub CLEAR_ITEM()
    vcDTP1.Value = Date:    DataCombo1.text = vbNullString:    TxtExCode.text = vbNullString:    TxtSaudaCode.text = vbNullString:
    DataCombo4.text = vbNullString
    TxtItemName.text = vbNullString:       TxtRefLot.text = vbNullString:        TxtSettleRate.text = vbNullString:     TxtAdminPasswd.text = vbNullString:    DataCombo3.text = vbNullString
    TxtTotBQty.text = vbNullString:        TxtBuyAvg.text = vbNullString:         TxtTotBuyAmt.text = vbNullString:     TxtTotSQty.text = vbNullString:
    TxtTotSAmt.text = vbNullString:       TxtDiffAmt.text = vbNullString
End Sub


Private Sub Calc_Totals()
Dim LTotBQnty As Double:    Dim LTotSQnty As Double:    Dim LTotBAmt As Double:     Dim LTotSamt As Double
Dim LBAmt As Double:        Dim LSAmt As Double:        Dim LDiffAmt As Double:      Dim TRec As ADODB.Recordset
    
    LBAmt = 0: LSAmt = 0: LDiffAmt = 0
    Set TRec = Nothing: Set TRec = New ADODB.Recordset: Set TRec = RECGRID.Clone
    LTotBQnty = 0: LTotSQnty = 0: LTotBAmt = 0: LTotSamt = 0
    Do While Not TRec.EOF
        LTotBQnty = LTotBQnty + Val(IIf(IsNull(TRec!BQnty), 0, TRec!BQnty))
        LTotBAmt = LTotBAmt + IIf(IsNull(TRec!BQnty), 0, TRec!BQnty * TRec!BRate) * FCalval
        LTotSQnty = LTotSQnty + IIf(IsNull(TRec!SQnty), 0, TRec!SQnty)
        LTotSamt = LTotSamt + (IIf(IsNull(TRec!SQnty), 0, TRec!SQnty) * TRec!SRate) * FCalval
        TRec.MoveNext
    Loop
    Set TRec = Nothing
    
    If LTotBQnty <> 0 Then TxtBuyAvg.text = Format((LTotBAmt / (LTotBQnty) * FCalval), "0.00")
    If LTotSQnty <> 0 Then TxtSellAvg.text = Format((LTotSamt / (LTotSQnty) * FCalval), "0.00")
    TxtTotBQty.text = Format(LTotBQnty, "0.00")
    TxtTotSQty.text = Format(LTotSQnty, "0.00")
    TxtTotBuyAmt.text = Format(LTotBAmt, "0.00")
    TxtDiffAmt.text = Format(Val(LTotBAmt) - Val(LTotSamt), "0.00")
    TxtTotSAmt.text = Format(LTotSamt, "0.00")
End Sub
