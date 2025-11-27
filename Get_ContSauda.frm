VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frm_GetContSauda 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Contract Entry"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txtreflot 
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
      Left            =   8520
      TabIndex        =   65
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "ERGGHE"
      Height          =   4095
      Left            =   1440
      TabIndex        =   60
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
         TabIndex        =   62
         Top             =   480
         Width           =   11895
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   61
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
         TabIndex        =   63
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtItemID 
      Height          =   735
      Left            =   4800
      TabIndex        =   59
      TabStop         =   0   'False
      Text            =   "Text12"
      Top             =   9720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox TxtExID 
      Height          =   615
      Left            =   12720
      TabIndex        =   58
      TabStop         =   0   'False
      Text            =   "Text13"
      Top             =   9360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox TxtSaudaId 
      Height          =   735
      Left            =   9240
      TabIndex        =   57
      TabStop         =   0   'False
      Text            =   "Text12"
      Top             =   9480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8415
      Left            =   16320
      TabIndex        =   55
      Top             =   720
      Width           =   7455
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   8055
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   14208
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   23
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
            Size            =   12
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   6375
      Left            =   10200
      TabIndex        =   48
      Top             =   2640
      Visible         =   0   'False
      Width           =   6015
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "Get_ContSauda.frx":0000
         Left            =   1320
         List            =   "Get_ContSauda.frx":0013
         TabIndex        =   52
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3960
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4920
         TabIndex        =   50
         Top             =   720
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   240
         TabIndex        =   49
         Top             =   1200
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   8705
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Filter By"
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
         TabIndex        =   54
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "Filter Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   17280
      TabIndex        =   39
      Text            =   "Text8"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   17400
      TabIndex        =   38
      Text            =   "Text7"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   19320
      TabIndex        =   37
      Text            =   "Text6"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6255
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   16095
      Begin VB.CommandButton Command9 
         Caption         =   "Delete All Trades Below"
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
         Left            =   7200
         TabIndex        =   47
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   " Filter Trade"
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
         Left            =   1800
         TabIndex        =   46
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox Text10 
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
         Left            =   14280
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear All Filter"
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
         TabIndex        =   40
         Top             =   120
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5295
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   9340
         _Version        =   393216
         BackColor       =   -2147483634
         HeadLines       =   1
         RowHeight       =   23
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
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Difference Amount"
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
         Left            =   12480
         TabIndex        =   44
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label13 
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
         Left            =   3600
         TabIndex        =   36
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   16095
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
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
         Height          =   615
         Left            =   2160
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox Text11 
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
         Left            =   14160
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
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
         Left            =   12480
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Left            =   10560
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   15855
         Begin VB.TextBox lottext 
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
            Left            =   8880
            TabIndex        =   16
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text9 
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
            Left            =   12360
            TabIndex        =   20
            Top             =   480
            Width           =   1455
         End
         Begin vcDateTimePicker.vcDTP vcDTP2 
            Height          =   420
            Left            =   5160
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   741
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
            Height          =   360
            Left            =   5160
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   635
            _Version        =   393216
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
         Begin VB.TextBox Text2 
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
            MaxLength       =   6
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   360
            Left            =   5160
            TabIndex        =   12
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   635
            _Version        =   393216
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
            Height          =   405
            Left            =   14640
            TabIndex        =   19
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text5 
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
            Left            =   10815
            TabIndex        =   18
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text4 
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
            Left            =   9615
            TabIndex        =   17
            Top             =   480
            Width           =   1095
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
            Height          =   405
            Left            =   8160
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "Buy"
            Top             =   480
            Width           =   615
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   420
            Left            =   2280
            TabIndex        =   11
            Top             =   480
            Width           =   2775
            _ExtentX        =   4895
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
            Height          =   405
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label20 
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
            Height          =   255
            Left            =   8880
            TabIndex        =   64
            Top             =   120
            Width           =   615
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
            Height          =   255
            Left            =   12360
            TabIndex        =   42
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Maturity"
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
            Left            =   6960
            TabIndex        =   41
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   14055
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Trade No"
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
            TabIndex        =   34
            Top             =   120
            Width           =   1095
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
            Height          =   255
            Left            =   10800
            TabIndex        =   32
            Top             =   120
            Width           =   1455
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
            Height          =   255
            Left            =   9600
            TabIndex        =   31
            Top             =   120
            Width           =   1095
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
            Height          =   255
            Left            =   8160
            TabIndex        =   30
            Top             =   120
            Width           =   615
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
            Height          =   255
            Left            =   5400
            TabIndex        =   29
            Top             =   120
            Width           =   2655
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
            Height          =   255
            Left            =   2280
            TabIndex        =   28
            Top             =   120
            Width           =   2895
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
            Height          =   255
            Left            =   1320
            TabIndex        =   27
            Top             =   120
            Width           =   855
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   420
         Left            =   4560
         TabIndex        =   3
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
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
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   120
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
         Value           =   41160.4222453704
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   420
         Left            =   7440
         TabIndex        =   4
         Top             =   120
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Ad Pass"
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
         Left            =   13320
         TabIndex        =   45
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Broker A/c"
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
         Left            =   6360
         TabIndex        =   33
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   195
         Width           =   855
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
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Width           =   495
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sauda Software"
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
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   16080
   End
End
Attribute VB_Name = "Frm_GetContSauda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExRec As ADODB.Recordset:       Dim PartyRec As ADODB.Recordset:    Dim SaudaRec As ADODB.Recordset:    Dim ItemRec As ADODB.Recordset
Dim ContRec As ADODB.Recordset:     Dim NRec As ADODB.Recordset:        Dim ExContRec As ADODB.Recordset:   Dim AllSaudaRec As ADODB.Recordset
Dim SaveCalled As Boolean:          Dim LExCont As String:              Dim LExCode As String:              Dim MSaudaCode  As String
Dim LItemCode As String:            Dim LItemCodeDBCombo As String:     Dim LSaudaCodeDBCombo As String:    Dim LCalval As Double
Dim LPNames As String:              Dim Ltype As String:                Dim LSaudas As String:              Dim LContractAcc As String
Dim LFBPress As Integer:            Dim Condate  As Date:               Dim LConNo As Long:                 Dim MQty  As Double
Dim MRate  As Double:               Dim MConRate As Double:             Dim LCParties As String:            Dim LCSaudas As String
Dim LBillParties As String:         Dim LBillSaudas As String:          Dim LBSauda As String:              Dim LFParties As String
Dim LFSaudas As String:             Dim LSPNames As String
Dim LBillExIds  As String:          Dim LBillItems As String

Private Sub Combo1_Validate(Cancel As Boolean)
Dim TRec As ADODB.Recordset
    If Combo1.ListIndex = 3 Then
        ListView1.ListItems.Clear
        ListView1.Enabled = True
        ListView1.ListItems.Add , , "Buy"
        ListView1.ListItems.Add , , "Sell"
    Else
        If Combo1.ListIndex = 0 Then mysql = "SELECT DISTINCT A.PARTY FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
        If Combo1.ListIndex = 1 Then mysql = "SELECT DISTINCT B.NAME FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
        If Combo1.ListIndex = 2 Then mysql = "SELECT DISTINCT A.SAUDA FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
        If Combo1.ListIndex = 4 Then mysql = "SELECT DISTINCT A.USERID FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
        mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE AND A.PARTY <>'" & LExCont & "'  AND CONCODE ='" & LExCont & "'"
        If LenB(LExCode) <> 0 Then mysql = mysql & " AND A.EXCODE ='" & LExCode & "'"
        mysql = mysql & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' "
        If LenB(LFParties) <> 0 Then mysql = mysql & " AND A.PARTY IN (" & LFParties & ")"
        If LenB(LSPNames) <> 0 Then mysql = mysql & " AND B.NAME IN (" & LSPNames & ")"
        If LenB(LFSaudas) <> 0 Then mysql = mysql & " AND A.SAUDAID IN (" & LFSaudas & ")"
        If Combo1.ListIndex = 0 Then mysql = mysql & " ORDER BY Party"
        If Combo1.ListIndex = 1 Then mysql = mysql & " ORDER BY NAME"
        If Combo1.ListIndex = 2 Then mysql = mysql & " ORDER BY SAUDA"
        If Combo1.ListIndex = 4 Then mysql = mysql & " ORDER BY A.USERID "
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not TRec.EOF Then
            ListView1.ListItems.Clear
            ListView1.Enabled = True
            Do While Not TRec.EOF
                If Combo1.ListIndex = 0 Then ListView1.ListItems.Add , , TRec!PARTY
                If Combo1.ListIndex = 1 Then ListView1.ListItems.Add , , TRec!NAME
                If Combo1.ListIndex = 2 Then ListView1.ListItems.Add , , TRec!Sauda
                If Combo1.ListIndex = 2 Then ListView1.ListItems.Add , , TRec!SAUDAID
                
                If Combo1.ListIndex = 4 Then ListView1.ListItems.Add , , IIf(IsNull(TRec!USERID), "", TRec!USERID)
                TRec.MoveNext
            Loop
        End If
    End If

End Sub

Private Sub CmdSave_Click()
    Dim LDelFlag As Boolean:    Dim LOConNo As String:    Dim LContime As String:    Dim LCSauda As String
    Dim LCItemCode As String:   Dim LConType As String:   Dim LSInstType As String:
    Dim LSOptType As String:    Dim LSStrike  As Double:  Dim mparty As String
    Dim LClient As String:      Dim LSConSno As Long:     Dim LOrdNo As String
    Dim LSaudaID As Long:    Dim LExID As Integer:     Dim LItemID As Integer
    LItemID = 0
    LExID = 0
    LSaudaID = 0
    LDelFlag = False
    DoEvents
    On Error GoTo err1
    Condate = vcDTP1.Value
    If LenB(Text2.text) < 1 Then
        MsgBox "Trade No can not be Blank"
        Text2.Locked = False
        Text2.SetFocus
        Exit Sub
    Else
        LConNo = Val(Text2.text)
    End If
    
    If LenB(Text1.text) < 1 Then
        MsgBox "Party Code can not be Blank"
        Text1.SetFocus
        Exit Sub
    Else
        mparty = Get_AccountDCode(Text1.text)
        If LenB(mparty) < 1 Then
            MsgBox "Invalid Party Code"
            Text1.text = vbNullString
            Text1.SetFocus
            Exit Sub
        Else
            LClient = mparty
        End If
    End If
    LCSauda = vbNullString
    If LenB(DataCombo3.BoundText) < 1 Then
        MsgBox "Sauda Code can not be Blank"
        Text1.SetFocus
        Exit Sub
    Else
        If Check1.Value = 1 Then
            LCItemCode = vbNullString: LCSauda = vbNullString
            AllSaudaRec.MoveFirst
            AllSaudaRec.Find "SAUDACODE='" & DataCombo3.BoundText & "'", , adSearchForward
            If Not AllSaudaRec.EOF Then
                LExID = Get_ExID(AllSaudaRec!excode)
                LCItemCode = Get_ItemMaster(LExID, AllSaudaRec!EX_SYMBOL)
                LSInstType = AllSaudaRec!INSTTYPE
                If LenB(LCItemCode) < 1 Then LCItemCode = Create_TItemMast(AllSaudaRec!ITEMCODE, AllSaudaRec!ITEMCODE, AllSaudaRec!EX_SYMBOL, AllSaudaRec!lot, AllSaudaRec!excode)
                LItemID = Get_ITEMID(LCItemCode)
                LCSauda = Get_SaudaMaster(LExID, LItemID, AllSaudaRec!MATURITY, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
                If LenB(LCSauda) < 1 Then LCSauda = Create_TSaudaMast(LCItemCode, AllSaudaRec!MATURITY, AllSaudaRec!excode, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
            Else
                MsgBox "Please Import Closing Rates for all Exchanges to import all Contracts "
                'Exit Sub
                LCSauda = DataCombo3.BoundText
            End If
        Else
            LCSauda = DataCombo3.BoundText
        End If
        
        mysql = "SELECT C.EXID,B.ITEMID,A.SAUDAID,A.SAUDACODE,B.ITEMCODE,B.LOT,A.TRADEABLELOT,B.EXCHANGECODE,INSTTYPE,C.LOTWISE,A.OPTTYPE,A.STRIKEPRICE FROM SAUDAMAST A, ITEMMAST B,EXMAST AS C WHERE A.COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND A.COMPCODE =C.COMPCODE AND B.EXID =C.EXID AND A.SAUDACODE ='" & LCSauda & "'AND A.COMPCODE =B.COMPCODE AND A.ITEMCODE =B.ITEMCODE "
        Set NRec = Nothing
        Set NRec = New ADODB.Recordset
        NRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If NRec.EOF Then
            MsgBox "Invalid Sauda Code"
            DataCombo3.SetFocus
            Exit Sub
        Else
            MSaudaCode = NRec!saudacode
            LItemCode = Trim(NRec!ITEMCODE)
            LExCode = NRec!EXCHANGECODE
            LExID = NRec!EXID
            LSInstType = NRec!INSTTYPE
            LSOptType = NRec!OPTTYPE
            LSStrike = NRec!STRIKEPRICE
            LSaudaID = NRec!SAUDAID
            LItemID = NRec!itemid
            If NRec!EXCHANGECODE = "NSE" And NRec!LOTWISE = "Y" Then
                LCalval = NRec!TRADEABLELOT
            ElseIf NRec!EXCHANGECODE = "MCX" And NRec!LOTWISE = "Y" Then
                LCalval = NRec!TRADEABLELOT
            Else
                LCalval = NRec!lot
            End If
        End If
    End If
    If Val(Text4.text) = 0 Then
        If LFBPress = 1 Then
            MsgBox "Trade Qty can not be Zero "
            Text4.SetFocus
            Exit Sub
        Else
            If MsgBox("You are about to Delete this Trade. Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
                LDelFlag = True
            Else
                MsgBox "Trade Qty can not be Zero "
                Text4.SetFocus
                Exit Sub
            End If
        End If
    Else
        MQty = Round(Val(Text4.text), 2)
    End If
    If Val(Text5.text) = 0 Then
        MsgBox "Trade Rate can not be Zero "
        Text5.SetFocus
        Exit Sub
    Else
        MRate = Round(Val(Text5.text), 4)
    End If
    If Val(Text9.text) = 0 Then
        MsgBox "Trade Con Rate can not be Zero "
        Text9.SetFocus
        Exit Sub
    Else
        MConRate = Round(Val(Text9.text), 4)
    End If
    LSConSno = 1:           LOConNo = LConNo
    LContime = Time:        LOrdNo = str(LConNo)
    Cnn.BeginTrans
    LSConSno = Get_ConSNo(Condate, MSaudaCode, LItemCode, LExCode, LSaudaID, LItemID, LExID)
    DoEvents
    If LFBPress = 2 Then
        LConNo = Val(Text2.text):           LOConNo = Trim(Text6.text)
        LContime = Trim(Text7.text):        LOrdNo = Trim(Text8.text)
        mysql = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONNO=" & Val(Text2.text) & "  AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
        Cnn.Execute mysql
    End If
    If LDelFlag = False Then
        If Text3.text = "Buy" Then
            LConType = "B"
        Else
            LConType = "S"
        End If
        If Not InStr(LBillExIds, str(LExID)) Then
            If LenB(LBillExIds) > 0 Then
                LBillExIds = LBillExIds & "," & str(LExID)
            Else
                LBillExIds = str(LExID)
            End If
        End If
        
        Call Add_To_Ctr_D2(LConType, LClient, LSConSno, Condate, LConNo, MSaudaCode, LItemCode, mparty, MQty, MRate, MConRate, LExCont, LContime, LOrdNo, vbNullString, LOConNo, LExCode, LCalval, 0, vbNullString, LSInstType, LSOptType, LSStrike, "0", "Y", LExID, LItemID, LSaudaID)
    End If
    Cnn.CommitTrans
    If LenB(LBillParties) < 1 Then
        LBillParties = "'" & mparty & "','" & LExCont & "'"
    Else
        If InStr(LBillParties, "'" & mparty & "'") < 1 Then LBillParties = LBillParties & ",'" & mparty & "'"
        If InStr(LBillParties, "'" & LExCont & "'") < 1 Then LBillParties = LBillParties & ",'" & LExCont & "'"
    End If
    
    If LenB(LBillItems) < 1 Then
        LBillItems = "'" & LItemID & "'"
    Else
        If InStr(LBillItems, LItemID) < 1 Then LBillItems = LBillItems & "," & "'" & LItemID & "'"
    End If
    If LenB(LBillSaudas) > 0 Then
        If LStr_Exists(LBillSaudas, str(LSaudaID)) = False Then LBillSaudas = LBillSaudas & "," & str(LSaudaID)
    Else
        LBillSaudas = str(LSaudaID)
    End If
    LConNo = LConNo + 1
    LExCode = DataCombo1.BoundText
    Call DATA_GRID_REFRESH
    Call SHOW_STANDING
    Text4.text = "0.00"
    Text5.text = "0.0000"
    Text9.text = "0.0000"
    If LFBPress = 2 Then
        Text2.text = vbNullString
        Text1.text = vbNullString
        DataCombo2.BoundText = vbNullString
        DataCombo3.BoundText = vbNullString
        Text2.SetFocus
    Else
        Text2.text = LConNo
        Text1.SetFocus
    End If
    SaveCalled = True
    Exit Sub
err1:
If err.Number <> 0 Then
    
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
   If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End If

End Sub
Private Sub Command2_Click()
    If LenB(DataCombo4.BoundText) = 0 Then
        MsgBox "Please Select Broker A/c"
        DataCombo4.SetFocus
        Exit Sub
    End If
    LConNo = Get_Max_ConNo(vcDTP1.Value, 0):      LConNo = LConNo + 1
    Text2.text = LConNo:                        Text2.Locked = True
    Text1.text = vbNullString:                  DataCombo2.BoundText = vbNullString
    Text4.Locked = False:                       Text4.text = vbNullString
    Text5.text = vbNullString:                  Text9.text = vbNullString
    TxtSaudaID.text = vbNullString
    TxtExID.text = vbNullString
    Command3.Enabled = False:                   Frame2.Enabled = True
    Text1.SetFocus:                             Command2.Enabled = False
    Command3.Enabled = False:                   Command4.Enabled = True
    vcDTP1.Enabled = False:                     DataCombo1.Enabled = False
    Check1.Enabled = False:                     DataCombo4.Enabled = False
    LFBPress = 1:                               Label12.Caption = "Adding New Trades"
End Sub

Private Sub Command3_Click()
    Call Mod_Rec
End Sub

Private Sub Command4_Click()
Call CANCEL_REC
End Sub

Private Sub Command5_Click()
Dim I As Integer
If Combo1.ListIndex = 0 Then
    LFParties = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LFParties) <> 0 Then LFParties = LFParties & ", "
            LFParties = LFParties & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
ElseIf Combo1.ListIndex = 1 Then
    LPNames = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LPNames) <> 0 Then LPNames = LPNames & ", "
            LPNames = LPNames & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
ElseIf Combo1.ListIndex = 2 Then
    LFSaudas = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LFSaudas) <> 0 Then LFSaudas = LFSaudas & ", "
            LFSaudas = LFSaudas & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
ElseIf Combo1.ListIndex = 3 Then
    Ltype = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
             If LenB(Ltype) <> 0 Then Ltype = Ltype & ", "
             Ltype = Ltype & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
  
ElseIf Combo1.ListIndex = 4 Then
    
End If

If GCINNo = "2000" Then
    mysql = "SELECT A.Party,B.Name ,A.CONTYPE AS Type ,A.Qty,A.Sauda ,A.Rate,A.BROKAMT AS ConRate,A.UserID,A.ConTime,A.ROWNO1 AS TradeNo,A.ConNo FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
    mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE AND A.PARTY <>'" & LExCont & "'  AND CONCODE ='" & LExCont & "'"
    If LenB(LFParties) <> 0 Then mysql = mysql & " AND A.PARTY IN (" & LFParties & ")"
    If LenB(LPNames) <> 0 Then mysql = mysql & " AND B.NAME IN (" & LPNames & ")"
    If LenB(LFSaudas) <> 0 Then mysql = mysql & " AND A.SAUDA IN (" & LFSaudas & ")"
    If LenB(Ltype) <> 0 And Len(Ltype) < 4 Then mysql = mysql & " AND A.CONTYPE ='" & Left$(Ltype, 1) & "'"
    If LenB(LExCode) <> 0 Then mysql = mysql & " AND A.EXCODE ='" & LExCode & "'"
    mysql = mysql & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ROWNO1 DESC "
Else
    mysql = "SELECT A.Party,B.Name,A.Sauda,A.ConNo,A.CONTYPE AS Type ,A.Qty,A.Rate,A.Contime,A.ROWNO1 AS TradeNo,A.OrdNo ,A.USERID  FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
    mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE AND A.PARTY <>'" & LExCont & "'  AND CONCODE ='" & LExCont & "'"
    If LenB(LFParties) <> 0 Then mysql = mysql & " AND A.PARTY IN (" & LFParties & ")"
    If LenB(LPNames) <> 0 Then mysql = mysql & " AND B.NAME IN (" & LPNames & ")"
    If LenB(LFSaudas) <> 0 Then mysql = mysql & " AND A.SAUDA IN (" & LFSaudas & ")"
    If LenB(Ltype) <> 0 And Len(Ltype) < 4 Then mysql = mysql & " AND A.CONTYPE ='" & Left$(Ltype, 1) & "'"
    
    If LenB(LExCode) <> 0 Then mysql = mysql & " AND A.EXCODE ='" & LExCode & "'"
    mysql = mysql & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY CONNO,CONTIME"

End If
    Set ContRec = Nothing
    Set ContRec = New ADODB.Recordset
    ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Set DataGrid1.DataSource = ContRec
    Call DATA_GRID_REFRESH
    Call Data_Grid_Refresh2
    Frame4.Visible = False
End Sub

Private Sub Command6_Click()
Frame4.Visible = False
DataGrid1.SetFocus
End Sub

Private Sub Command7_Click()
LFParties = vbNullString:    LFSaudas = vbNullString:    LPNames = vbNullString:
Call DATA_GRID_REFRESH
End Sub

Private Sub Command8_Click()
    If Command2.Enabled = False Then
        If ContRec.RecordCount > 0 Then
            Frame4.Visible = True
            DoEvents
            ListView1.ListItems.Clear
            Combo1.SetFocus
        End If
    End If
End Sub

Private Sub Command9_Click()
Dim LDel As Boolean
If ContRec.RecordCount > 0 Then
    If MsgBox("Are You Sure You Want to Delte all Trades of " & vcDTP1.Value & " of " & DataCombo1.BoundText & "", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
        mysql = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & DataCombo1.BoundText & "' AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND CONCODE ='" & DataCombo4.BoundText & "'"
        Cnn.Execute mysql
    End If
    DATA_GRID_REFRESH
End If
End Sub

Private Sub DataCombo1_Validate(Cancel As Boolean)
    If Check1.Value = 1 Then
        mysql = "SELECT SAUDACODE,SAUDANAME,EX_SYMBOL,ITEMCODE,MATURITY,EXCODE,INSTTYPE,OPTTYPE,STRIKEPRICE,LOT FROM SCRIPTMASTER WHERE  "
        mysql = mysql & " MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY EXCODE,ITEMCODE,INSTTYPE,MATURITY"
        Set AllSaudaRec = Nothing
        Set AllSaudaRec = New ADODB.Recordset
        AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not AllSaudaRec.EOF Then
            Set DataCombo3.RowSource = AllSaudaRec
            DataCombo3.BoundColumn = "SAUDACODE"
            DataCombo3.ListField = "SAUDANAME"
        End If
    Else
        mysql = "SELECT SAUDACODE,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY ITEMCODE,INSTTYPE,MATURITY"
        Set SaudaRec = Nothing
        Set SaudaRec = New ADODB.Recordset
        SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not SaudaRec.EOF Then
            Set DataCombo3.RowSource = SaudaRec
            DataCombo3.BoundColumn = "SAUDACODE"
            DataCombo3.ListField = "SAUDANAME"
        End If
    End If
End Sub

Private Sub DataCombo3_Validate(Cancel As Boolean)
Dim MSauda As String
If DataCombo1.BoundText = "" Then
    LExCode = vbNullString
    If Check1.Value = 1 Then
        mysql = "SELECT SAUDACODE,SAUDANAME,EX_SYMBOL,ITEMCODE,MATURITY,EXCODE,INSTTYPE,OPTTYPE,STRIKEPRICE,LOT FROM SCRIPTMASTER WHERE  "
        mysql = mysql & " MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY EXCODE,ITEMCODE,INSTTYPE,MATURITY"
        Set AllSaudaRec = Nothing
        Set AllSaudaRec = New ADODB.Recordset
        AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not AllSaudaRec.EOF Then
            Set DataCombo3.RowSource = AllSaudaRec
            DataCombo3.BoundColumn = "SAUDACODE"
            DataCombo3.ListField = "SAUDANAME"
        End If
    Else
        mysql = "SELECT SAUDACODE,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY ITEMCODE,INSTTYPE,MATURITY"
        Set SaudaRec = Nothing
        Set SaudaRec = New ADODB.Recordset
        SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not SaudaRec.EOF Then
            Set DataCombo3.RowSource = SaudaRec
            DataCombo3.BoundColumn = "SAUDACODE"
            DataCombo3.ListField = "SAUDANAME"
        End If
    End If
Else
    LExCode = DataCombo1.BoundText
    ExRec.Filter = adFilterNone
    ExRec.Filter = "EXCODE='" & LExCode & "'"
    If Check1.Value = 1 Then
        mysql = "SELECT SAUDACODE,SAUDANAME,EX_SYMBOL,ITEMCODE,MATURITY,EXCODE,INSTTYPE,OPTTYPE,STRIKEPRICE,LOT FROM SCRIPTMASTER "
        mysql = mysql & " WHERE  EXCODE ='" & DataCombo1.BoundText & "' AND  MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,INSTTYPE,MATURITY"
        Set AllSaudaRec = Nothing
        Set AllSaudaRec = New ADODB.Recordset
        AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not AllSaudaRec.EOF Then
            Set DataCombo3.RowSource = AllSaudaRec
            DataCombo3.BoundColumn = "SAUDACODE"
            DataCombo3.ListField = "SAUDANAME"
        End If
        mysql = "SELECT ITEMCODE ,ITEMNAME,EXCODE FROM CONTRACTMASTER  WHERE EXCODE ='" & DataCombo1.BoundText & "' ORDER BY ITEMNAME"
        Set ItemRec = Nothing
        Set ItemRec = New ADODB.Recordset
        ItemRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not ItemRec.EOF Then
            Set DataCombo5.RowSource = ItemRec
            DataCombo5.BoundColumn = "ITEMCODE"
            DataCombo5.ListField = "ITEMNAME"
        End If
    Else
        mysql = "SELECT SAUDACODE,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & DataCombo1.BoundText & "'"
        mysql = mysql & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY ITEMCODE,INSTTYPE,MATURITY"
        Set SaudaRec = Nothing
        Set SaudaRec = New ADODB.Recordset
        SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not SaudaRec.EOF Then
            Set DataCombo3.RowSource = SaudaRec
            DataCombo3.BoundColumn = "SAUDACODE"
            DataCombo3.ListField = "SAUDANAME"
            
        End If
    
        mysql = "SELECT ITEMCODE ,ITEMNAME,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND  EXCHANGECODE ='" & DataCombo1.BoundText & "' ORDER BY ITEMNAME"
        Set ItemRec = Nothing
        Set ItemRec = New ADODB.Recordset
        ItemRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not ItemRec.EOF Then
            Set DataCombo5.RowSource = ItemRec
            DataCombo5.BoundColumn = "ITEMCODE"
            DataCombo5.ListField = "ITEMNAME"
        End If
        If LExCode = "NDF" Then
            DataCombo3.Visible = False
            vcDTP2.Visible = True
            DataCombo5.Visible = True
            Label16.Visible = True
            Label6.Caption = "Commodity"
        Else
            DataCombo3.Visible = True
            vcDTP2.Visible = False
            DataCombo5.Visible = False
            Label16.Visible = False
            Label6.Caption = "Contract"
        End If
    End If
End If
MSauda = DataCombo3.BoundText
If Len(MSauda) > 0 Then
    If Check1.Value = 1 Then
        mysql = "SELECT reflot FROM SCRIPTMASTER WHERE SAUDACODE = '" & MSauda & "' "
        Set NRec = Nothing
        Set NRec = New ADODB.Recordset
        NRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not NRec.EOF Then
            TxtRefLot.text = Format(NRec!REFLOT, "0.00")
        End If
   Else
        mysql = "SELECT reflot FROM SAUDAMAST WHERE  SAUDACODE = '" & MSauda & "' "
        Set NRec = Nothing
        Set NRec = New ADODB.Recordset
        NRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not NRec.EOF Then
            TxtRefLot.text = Format(NRec!REFLOT, "0.00")
        End If
   End If
Else
   TxtRefLot.text = "0.00"
End If
Call SHOW_STANDING
End Sub
Private Sub DataCombo4_Validate(Cancel As Boolean)
If DataCombo4.BoundText = "" Then
    MsgBox "Please Select Exchange Contract A/c"
    Cancel = True
Else
    LExCont = DataCombo4.BoundText
    Call DATA_GRID_REFRESH
End If
End Sub
Private Sub Check1_Validate(Cancel As Boolean)
If DataCombo1.Enabled = False Then
    Call DataCombo3_Validate(False)
End If
End Sub

Private Sub DataCombo2_Validate(Cancel As Boolean)
If DataCombo2.text = "" Then
    MsgBox "Party can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & DataCombo2.BoundText & "'"
    Set NRec = Nothing
    Set NRec = New ADODB.Recordset
    NRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not NRec.EOF Then
        If NRec!AC_CODE = LExCont Then
            MsgBox "Party Account Can Not Be Same As Broker A/c"
            Cancel = True
            DataCombo2.SetFocus
            Exit Sub
        End If
        DataCombo3.SetFocus
    Else
        DataCombo2.SetFocus
    End If
    Text1.text = DataCombo2.BoundText
End If

End Sub
'Private Sub DataCombo3_Validate(Cancel As Boolean)
'If DataCombo3.text = "" Then
'    MsgBox "Sauda can not be blank"
'    Cancel = True
 '   Sendkeys "%{DOWN}"
'End If
'End Sub

Private Sub DataCombo5_Validate(Cancel As Boolean)
If DataCombo5.text = "" Then
    MsgBox "Itemcode can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    LItemCodeDBCombo = DataCombo5.BoundText
End If
End Sub

Private Sub DataGrid1_DblClick()
Dim LPConNo As Integer
Dim LPSauda As String
Dim LPConType As String
Dim TRec As ADODB.Recordset
Dim LSaudaID As Long
If GCINNo = "2000" Then
    DataGrid1.Col = 4
    LPSauda = DataGrid1.text
    DataGrid1.Col = 10
    LPConNo = DataGrid1.text
    DataGrid1.Col = 2
    LPConType = DataGrid1.text
Else
    DataGrid1.Col = 2
    LPSauda = DataGrid1.text
    DataGrid1.Col = 3
    LPConNo = DataGrid1.text
    DataGrid1.Col = 4
    LPConType = DataGrid1.text
End If
Call Mod_Rec
mysql = "SELECT CONSNO,CONNO, QTY,RATE,PARTY,CONTYPE,SAUDA,ITEMCODE,EXCODE,CALVAL,CONCODE,SAUDAID,ITEMID,EXID FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND CONNO=" & LPConNo & ""
Set TRec = Nothing
Set TRec = New ADODB.Recordset
TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If TRec.EOF Then
Else
    Do While Not TRec.EOF
        Text2.text = TRec!CONNO
        If TRec!CONTYPE = LPConType Then
            Text1.text = TRec!PARTY
            DataCombo2.BoundText = TRec!PARTY
            Text5.text = Format(TRec!Rate, "0.0000")
        Else
            DataCombo4.BoundText = TRec!PARTY
            Text9.text = Format(TRec!Rate, "0.0000")
        End If
        LExCode = TRec!excode
        DataCombo3.BoundText = TRec!Sauda
        If LPConType = "B" Then
            Text3.text = "Buy"
        Else
            Text3.text = "Sel"
        End If
        Text4.text = TRec!QTY
        TxtSaudaID.text = TRec!SAUDAID
        LSaudaID = TRec!SAUDAID
        TxtItemID.text = TRec!itemid
        
        TxtExID.text = TRec!EXID
        If LenB(LBillParties) = 0 Then
            LBillParties = "'" & TRec!PARTY & "','" & TRec!CONCODE & "'"
        Else
            If InStr(LBillParties, "'" & TRec!PARTY & "'") < 1 Then LBillParties = LBillParties & ",'" & TRec!PARTY & "'"
            If InStr(LBillParties, "'" & TRec!CONCODE & "'") < 1 Then LBillParties = LBillParties & ",'" & TRec!CONCODE & "'"
        End If
        
        If LenB(LBillItems) < 1 Then
            LBillItems = "'" & TxtItemID.text & "'"
        Else
            If InStr(LBillItems, TxtItemID.text) < 1 Then LBillItems = LBillItems & "," & "'" & TxtItemID.text & "'"
        End If

        If LenB(LBillSaudas) > 0 Then
            If LStr_Exists(LBillSaudas, str(LSaudaID)) = False Then LBillSaudas = LBillSaudas & "," & str(LSaudaID)
        Else
            LBillSaudas = str(LSaudaID)
        End If
        
        'If LenB(LBillSaudas) = 0 Then
        '    LBillSaudas = "'" & TRec!Sauda & "'"
        'Else
        '    If InStr(LBillSaudas, TRec!Sauda) < 1 Then LBillSaudas = LBillSaudas & ",'" & TRec!Sauda & "'"
        'End If
    
        TRec.MoveNext
    Loop
    Command2.Enabled = True:                Command3.Enabled = False
    Command4.Enabled = True:                vcDTP1.Enabled = False
    DataCombo1.Enabled = False:             DataCombo4.Enabled = False
    Check1.Enabled = False:                 Frame2.Enabled = True
    LFBPress = 2
    Label12.Caption = "Modifying Existing Trades"
    Text2.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "vcDTP1" Or Me.ActiveControl.NAME = "vcDTP2" Then
            Sendkeys "{tab}"
        End If
    End If
    If Command2.Enabled = False Then
       If KeyCode = 121 Then Frame4.Visible = True
    End If
End Sub

Private Sub Form_Load()
LCParties = vbNullString
LCSaudas = vbNullString
vcDTP1.Value = Date
Frame2.Visible = True
If GCINNo = "2000" Then
    DataCombo5.Left = 7300:     DataCombo3.Left = 7300
    Label7.Left = 5280:         Label8.Left = 6080
    Label6.Left = 7300:         Text3.Left = 5280
    Text4.Left = 6080:          Text3.TabIndex = 12
    Text4.TabIndex = 13:        DataCombo3.TabIndex = 14
End If

Set ExRec = Nothing
Set ExRec = New ADODB.Recordset
mysql = "SELECT EXCODE,EXNAME,CONTRACTACC FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME"
ExRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not ExRec.EOF Then
  Set DataCombo1.RowSource = ExRec
    DataCombo1.BoundColumn = "EXCODE"
    DataCombo1.ListField = "EXCODE"
End If
If ExRec.RecordCount = 1 Then
    DataCombo1.BoundText = ExRec!excode
    DataCombo1.Enabled = False
End If
Set PartyRec = Nothing
Set PartyRec = New ADODB.Recordset
mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
PartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not PartyRec.EOF Then
  Set DataCombo2.RowSource = PartyRec
    DataCombo2.BoundColumn = "AC_CODE"
    DataCombo2.ListField = "NAME"
End If
Set ExContRec = Nothing
Set ExContRec = New ADODB.Recordset
mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND PARTYTYPE='1' ORDER BY NAME"
ExContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not ExContRec.EOF Then
    Set DataCombo4.RowSource = ExContRec
    DataCombo4.BoundColumn = "AC_CODE"
    DataCombo4.ListField = "NAME"
End If
End Sub
Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo4_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo2_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo5_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call CANCEL_REC
End Sub

Private Sub lottext_GotFocus()
    lottext.SelStart = 0
    lottext.SelLength = Len(lottext.text)
End Sub

Private Sub lottext_Validate(Cancel As Boolean)
    If Val(lottext) > 0 Then Text4.text = CStr(Val(lottext.text) * Val(TxtRefLot.text))
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Text1.text = "" Then
Else
    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & Text1.text & "'"
    Set NRec = Nothing
    Set NRec = New ADODB.Recordset
    NRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not NRec.EOF Then
        DataCombo2.BoundText = NRec!AC_CODE
        If NRec!AC_CODE = LExCont Then
            MsgBox "Party Account Can Not Be Same As Broker A/c"
            DataCombo2.SetFocus
            Exit Sub
        End If
        If GCINNo = "2000" Then
            DataCombo2.SetFocus
        Else
            DataCombo3.SetFocus
        End If
    End If
End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
Dim NewRec As ADODB.Recordset
Text4.Locked = False
If LFBPress = 2 Then
    If Text2.text = "" Then
        Text4.Locked = True
    Else
        Text4.Locked = False
        mysql = "SELECT A.PARTY,A.ITEMCODE,A.SAUDA,A.CONDATE,A.CONNO,A.CONSNO,A.CONTYPE,A.QTY,A.RATE,B.NAME,A.ROWNO1,A.ORDNO,A.CONTIME,BROKAMT FROM CTR_D AS A, ACCOUNTD AS B WHERE A.COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND A.COMPCODE=B.COMPCODE AND A.PARTY=B.AC_CODE AND A.CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'  AND A.CONNO=" & Val(Text2.text) & " AND A.CONCODE  ='" & LExCont & "'"
        Set NewRec = Nothing
        Set NewRec = New ADODB.Recordset
        NewRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not NewRec.EOF Then
            If NewRec!PARTY = LExCont Then
                Text9.text = Format(NewRec!Rate, "0.0000")
                NewRec.MoveNext
            End If
            Text2.text = NewRec!CONNO
            Text1.text = NewRec!PARTY
            DataCombo2.BoundText = NewRec!PARTY
            DataCombo3.BoundText = NewRec!Sauda
            If NewRec!CONTYPE = "B" Then
                Text3.text = "Buy"
            Else
                Text3.text = "Sel"
            End If
            Text4.text = NewRec!QTY
            Text5.text = Format(NewRec!Rate, "0.0000")
            Text6.text = NewRec!ROWNO1
            Text7.text = NewRec!contime
            Text8.text = NewRec!ORDNO
            Text9.text = Format(NewRec!BROKAMT, "0.0000")
            If Not NewRec.EOF Then
                If NewRec!PARTY = LExCont Then
                    Text9.text = Format(NewRec!Rate, "0.0000")
                    NewRec.MoveNext
                End If
            End If
        Else
            MsgBox " Invalid Trade No"
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Val(KeyAscii) >= 48 And KeyAscii <= 122 Then
    If Val(KeyAscii) = 66 Or Val(KeyAscii) = 98 Or Val(KeyAscii) = 83 Or Val(KeyAscii) = 115 Then
    Else
        If Text3.text = "Buy" Then
            Text3.text = "Sel"
        Else
            Text3.text = "Buy"
        End If
    End If
End If
If KeyAscii = 32 Then
    If Text3.text = "Buy" Then
        Text3.text = "Sel"
    Else
        Text3.text = "Buy"
    End If
End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Text3.text <> "Buy" Then
    If Text3.text <> "Sel" Then
        Text3.text = "Buy"
        Cancel = True
        Text3.SetFocus
    End If
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    Text4.text = Format(Text4.text, "0.00")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    Text5.text = Format(Text5.text, "0.0000")
    If LFBPress = 1 Then Text9.text = Format(Text5.text, "0.0000")
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    Text9.text = Format(Text9.text, "0.0000")
End Sub
Private Sub Text4_GotFocus()
    Text4.SelLength = Len(Text4.text)
End Sub
Private Sub Text9_GotFocus()
    Text9.SelLength = Len(Text9.text)
End Sub

Private Sub Text5_GotFocus()
    Text5.SelLength = Len(Text5.text)
End Sub
Private Sub Text1_GotFocus()
    Text1.SelLength = Len(Text1.text)
End Sub
Public Sub DATA_GRID_REFRESH()
Dim LShreeRec As ADODB.Recordset
If LenB(LExCont) <> 0 Then
    If GCINNo = "2000" Then
        mysql = "SELECT A.Party,B.Name ,A.CONTYPE AS Type ,A.Qty,A.Sauda ,A.Rate,A.BROKAMT AS ConRate,A.UserID,A.ConTime,A.ROWNO1 AS TradeNo,A.ConNo FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
        mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE AND A.PARTY <>'" & LExCont & "'  AND CONCODE ='" & LExCont & "'"
        If LenB(LExCode) <> 0 Then mysql = mysql & " AND A.EXCODE ='" & LExCode & "'"
        mysql = mysql & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY CONNO DESC "
        Set ContRec = Nothing
        Set ContRec = New ADODB.Recordset
        ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Set DataGrid1.DataSource = ContRec
        Call Data_Grid_Refresh2
    Else
        mysql = "SELECT A.Party,B.Name ,A.Sauda ,A.ConNo,A.CONTYPE AS Type ,A.Qty,A.Rate,A.BROKAMT AS ConRate,A.ConTime,A.ROWNO1 AS TradeNo,A.UserID FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
        mysql = mysql & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE AND A.PARTY <>'" & LExCont & "' AND CONCODE ='" & LExCont & "'"
        If LenB(LExCode) <> 0 Then mysql = mysql & " AND A.EXCODE ='" & LExCode & "'  "
        mysql = mysql & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY CONNO DESC "
        Set ContRec = Nothing
        Set ContRec = New ADODB.Recordset
        ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Set DataGrid1.DataSource = ContRec
        Call Data_Grid_Refresh2
    End If
    mysql = "SELECT SUM(CASE CONTYPE WHEN 'B' THEN (A.Qty*A.Rate*A.CALVAL)WHEN 'S'THEN (A.Qty*A.Rate*A.CALVAL)*-1 end) AS DIFFAMT  FROM CTR_D A WHERE A.COMPCODE=" & GCompCode & ""
    mysql = mysql & " AND CONCODE ='" & LExCont & "'"
    If LenB(LExCode) <> 0 Then mysql = mysql & " AND A.EXCODE ='" & LExCode & "'"
    mysql = mysql & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    Set LShreeRec = Nothing
    Set LShreeRec = New ADODB.Recordset
    LShreeRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LShreeRec.EOF Then
        If Not IsNull(LShreeRec!DIFFAMT) Then
            If LShreeRec!DIFFAMT <> 0 Then
                  Text10.text = Format(Val(LShreeRec!DIFFAMT), "0.00")
            End If
        End If
    End If
End If
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
    If Not ContRec.EOF Then Set DataGrid1.DataSource = ContRec:
    Call Data_Grid_Refresh2
    
Error1:    Exit Sub
End Sub

Private Sub CANCEL_REC()
Dim LMEXCODE As String:         Dim SREC As ADODB.Recordset:        Dim PREC As ADODB.Recordset
'Dim LBSaudas As String:         Dim LBParties As String:            Dim LBItems As String

'LBSaudas = vbNullString:        LBillParties = vbNullString:        LBItems = vbNullString
LFParties = vbNullString:       LSPNames = vbNullString:            Frame1.Enabled = False
Label12.Caption = "Updateing Bills Please Wait"
GETMAIN.Toolbar1_Buttons(6).Enabled = False
On Error GoTo err1
    Command2.Enabled = True:            Command3.Enabled = True:        Command4.Enabled = False
    vcDTP1.Enabled = True:              DataCombo4.BoundText = vbNullString
    If ExRec.RecordCount = 1 Then
        DataCombo1.BoundText = ExRec!excode
        DataCombo1.Enabled = False
    Else
        DataCombo1.Enabled = True
    End If
    Check1.Enabled = True:    DataCombo4.Enabled = True
    Text2.Locked = False:     Frame2.Enabled = False
    LFBPress = 0:            LSPNames = vbNullString:
    If SaveCalled = True Then
        Call Shree_Posting(DateValue(vcDTP1.Value))
        Cnn.BeginTrans
        Call RATE_TEST(vcDTP1.Value, , , Frm_GetContSauda)

        Call Update_Charges(LBillParties, LBillExIds, LBillSaudas, LBillItems, vcDTP1.Value, vcDTP1.Value, True)
        If BILL_GENERATION(vcDTP1.Value, GFinEnd, LBillSaudas, LBillParties, LBillExIds) Then Cnn.CommitTrans
        'Call Chk_Billing
    End If
    SaveCalled = False
    Frame1.Enabled = True
    GETMAIN.Toolbar1_Buttons(6).Enabled = True
    vcDTP1.SetFocus
    LExCode = vbNullString
    Call Command7_Click
    Label12.Caption = "Bills Updated Successfully "
    TxtSaudaID.text = vbNullString
    TxtExID.text = vbNullString
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
        
        Cnn.RollbackTrans: CNNERR = False
        Frame1.Enabled = True
    End If
End Sub
Private Sub vcDTP1_Validate(Cancel As Boolean)
    Dim NRec1 As ADODB.Recordset
    vcDTP2.MinDate = vcDTP1.Value
    vcDTP2.Value = vcDTP1.Value + 90
    'If GRateSlab = 1 Then
    '    MYSQL = "SELECT  TOP 1 COMPCODE  FROM  CTR_R WHERE COMPCODE  =" & GCompCode & "  AND CONDATE ='" & Format(vcDTP1.Value, "yyyy/mm/dd") & "'"
    '    Set NRec1 = Nothing
    '    Set NRec1 = New ADODB.Recordset
    '    NRec1.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    '    If Not NRec1.EOF Then
    '        Label19.Visible = True
    '        Text11.Visible = True
    '        Command3.Enabled = False
    '        Command2.Enabled = False
    '    Else
    '        Label19.Visible = False
    '        Text11.Visible = False
    '        Command3.Enabled = True
            Command2.Enabled = True
    '    End If
    'Else
        Label19.Visible = False
        Text11.Visible = False
    'End If
End Sub
Private Sub vcDTP2_Validate(Cancel As Boolean)
Dim TRec As ADODB.Recordset:        Dim LSaudaCode As String:       Dim LFLAG As Boolean:       Dim LTExCode As String
Dim LExID  As Integer: Dim LItemID As Integer
mysql = "SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE ='" & DataCombo5.BoundText & "' AND MATURITY='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
Set TRec = Nothing
Set TRec = New ADODB.Recordset
TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If TRec.EOF Then
    MsgBox "Creating New Contract for " & vcDTP2.Value & " Prompt Date"
    LSaudaCode = DataCombo5.text & " PD" & vcDTP2.Value
    ItemRec.MoveFirst
    ItemRec.Find "ITEMCODE ='" & DataCombo5.text & "'"
    If ItemRec.EOF Then
        MsgBox "Invalid Item "
    Else
        LTExCode = ItemRec!EXCHANGECODE
    End If
    LExID = Get_ExID(LTExCode)
    LItemID = Get_ITEMID(LItemCodeDBCombo)
    Call PInsert_Saudamast(LSaudaCode, LSaudaCode, LItemCodeDBCombo, vcDTP2.Value, 1, 1, 0, "FUT", vbNullString, 0, LTExCode, 1, LExID, LItemID)
    'MYSQL = "EXEC INSERT_SAUDAMAST " & GCompCode  & ",'" & LSaudaCode & "','" & LSaudaCode & "','" & LItemCodeDBCombo & "','" & Format(vcDTP2.Value, "yyyy/MM/dd") & "',1,'FUT','',0,'" & LTExCode & "',1,1"
    'Cnn.Execute MYSQL
    LSaudaCodeDBCombo = LSaudaCode
    LFLAG = True
Else
    LSaudaCodeDBCombo = TRec!saudacode
End If
If LFLAG = True Then
    mysql = "SELECT SAUDACODE,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND  EXCODE ='" & DataCombo1.BoundText & "') "
    mysql = mysql & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY ITEMCODE,MATURITY"
    Set SaudaRec = Nothing
    Set SaudaRec = New ADODB.Recordset
    SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not SaudaRec.EOF Then
        Set DataCombo3.RowSource = SaudaRec
        DataCombo3.BoundColumn = "SAUDACODE"
        DataCombo3.ListField = "SAUDANAME"
    End If
End If
DataCombo3.BoundText = LSaudaCodeDBCombo
End Sub
Private Sub Mod_Rec()
If LenB(DataCombo4.BoundText) < 1 Then
    MsgBox "Please Select Broker A/c"
    DataCombo4.SetFocus
    Exit Sub
End If
If ContRec.RecordCount > 0 Then
    Command2.Enabled = False:           Command3.Enabled = False
    Command4.Enabled = True:            vcDTP1.Enabled = False
    DataCombo1.Enabled = False:         Check1.Enabled = False
    DataCombo4.Enabled = False:         Text2.text = vbNullString
    Text2.Locked = False:               Text4.Locked = False
    Text4.text = vbNullString:          Text5.text = vbNullString
    Text9.text = vbNullString:          Frame2.Enabled = True
    TxtSaudaID.text = vbNullString
    TxtExID.text = vbNullString
    LFBPress = 2
    Label12.Caption = "Modifying Existing Trades"
    Text2.SetFocus
Else
    MsgBox "No Records to Modify "
End If
End Sub
Public Sub SHOW_STANDING()
Dim NStandRec As ADODB.Recordset
'Call DComboParty_Validate
'Call DComboSauda_Validate
mysql = "SELECT A.Name ,B.Sauda,SUM(CASE B.CONTYPE WHEN 'B' THEN B.QTY * 1 WHEN 'S' THEN B.QTY*-1 END) AS NetQty"
mysql = mysql & " FROM ACCOUNTD AS A,CTR_D AS B, SAUDAMAST AS S WHERE A.COMPCODE= " & GCompCode & "  AND A.COMPCODE =B.COMPCODE AND A.COMPCODE =S.COMPCODE AND S.MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND B.SAUDAID=S.SAUDAID AND A.AC_CODE =B.PARTY AND B.CONDATE <='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
If LenB(LExCode) <> 0 Then mysql = mysql & " AND S.EXCODE='" & LExCode & "'"
mysql = mysql & " GROUP BY A.NAME,B.SAUDA HAVING SUM(CASE B.CONTYPE WHEN 'B' THEN B.QTY * 1 WHEN 'S' THEN B.QTY*-1 END) <>0 ORDER BY A.NAME,B.SAUDA"
Set NStandRec = Nothing
Set NStandRec = New ADODB.Recordset
NStandRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not NStandRec.EOF Then
    Set DataGrid2.DataSource = NStandRec
    DataGrid2.ReBind
    DataGrid2.Refresh
    DataGrid2.Columns(0).Width = 2500:
    DataGrid2.Columns(1).Width = 2500
    DataGrid2.Columns(2).Width = 1000
    DataGrid2.Columns(2).Alignment = dbgRight:
End If
End Sub

Private Sub Data_Grid_Refresh2()
DataGrid1.ReBind: DataGrid1.Refresh
    If GCINNo = "2000" Then
        DataGrid1.Columns(1).Width = 3000:              DataGrid1.Columns(2).Width = 800
        DataGrid1.Columns(3).Width = 900:               DataGrid1.Columns(4).Width = 3000
        DataGrid1.Columns(5).Width = 1000:              DataGrid1.Columns(6).Width = 1200
        DataGrid1.Columns(7).Width = 1200:              DataGrid1.Columns(8).Width = 1200
        DataGrid1.Columns(7).Alignment = dbgRight:      DataGrid1.Columns(8).Alignment = dbgRight
        DataGrid1.Columns(2).Alignment = dbgCenter:     DataGrid1.Columns(9).Alignment = dbgRight
        DataGrid1.Columns(10).Alignment = dbgRight:     DataGrid1.Columns(3).Alignment = dbgCenter
        DataGrid1.Columns(7).NumberFormat = "0.0000":   DataGrid1.Columns(5).NumberFormat = "0.00"
        DataGrid1.Columns(5).Alignment = dbgRight:      DataGrid1.Columns(6).NumberFormat = "0.0000"
        DataGrid1.Columns(6).Alignment = dbgRight
    Else
        DataGrid1.Columns(0).Width = 800
        DataGrid1.Columns(1).Width = 3300
        DataGrid1.Columns(2).Width = 3300:              DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Width = 800:              DataGrid1.Columns(1).Width = 2500
        DataGrid1.Columns(2).Width = 3300:              DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Width = 1000:              DataGrid1.Columns(5).Width = 1000
        DataGrid1.Columns(6).Width = 1200:              DataGrid1.Columns(7).Width = 1200
        DataGrid1.Columns(8).Width = 1200:              DataGrid1.Columns(7).Alignment = dbgRight
        DataGrid1.Columns(8).Alignment = dbgRight:      DataGrid1.Columns(5).NumberFormat = "0.00"
        DataGrid1.Columns(5).Alignment = dbgRight:      DataGrid1.Columns(6).NumberFormat = "0.0000"
        DataGrid1.Columns(7).NumberFormat = "0.0000":   DataGrid1.Columns(6).Alignment = dbgRight
    End If
End Sub
