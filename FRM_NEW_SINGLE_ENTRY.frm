VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_NEW_SINGLE_ENTRY 
   BackColor       =   &H00FFC0C0&
   Caption         =   "New Contract Entry"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15960
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
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
      TabIndex        =   25
      Top             =   0
      Width           =   19335
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   19335
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "Single Contract Entry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   19095
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   255
      Left            =   18000
      TabIndex        =   17
      Top             =   9240
      Visible         =   0   'False
      Width           =   615
      Begin VB.TextBox Text2 
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
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Text2"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox Text6 
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
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox Text5 
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
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRM_NEW_SINGLE_ENTRY.frx":0000
         Left            =   120
         List            =   "FRM_NEW_SINGLE_ENTRY.frx":000A
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
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
      Height          =   8925
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   19140
      Begin VB.Frame Frame8 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   6975
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   18855
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6660
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   18570
            _ExtentX        =   32755
            _ExtentY        =   11748
            _Version        =   393216
            AllowArrows     =   -1  'True
            ForeColor       =   128
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
            ColumnCount     =   20
            BeginProperty Column00 
               DataField       =   "buysell"
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
            BeginProperty Column01 
               DataField       =   "Code"
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
               DataField       =   "name"
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
               DataField       =   "Saudacode"
               Caption         =   "SaudaCode"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "SaudaName"
               Caption         =   "Sauda Name"
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
            BeginProperty Column05 
               DataField       =   "Qnty"
               Caption         =   "Qnty"
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
            BeginProperty Column06 
               DataField       =   "Rate"
               Caption         =   "Rate"
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
            BeginProperty Column07 
               DataField       =   "Concode"
               Caption         =   "ConCode"
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
               DataField       =   "ConName"
               Caption         =   "Con. Name"
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
               DataField       =   "rate1"
               Caption         =   "Con. Rate"
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
            BeginProperty Column10 
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
            BeginProperty Column11 
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
            BeginProperty Column12 
               DataField       =   "CONTIME"
               Caption         =   "ConTime"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "H:mm:ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column13 
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
            BeginProperty Column14 
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
            BeginProperty Column15 
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
            BeginProperty Column18 
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
            BeginProperty Column19 
               DataField       =   "diffamt"
               Caption         =   "ShreeAmt"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   2
               BeginProperty Column00 
                  ColumnWidth     =   450.142
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column10 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column14 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column15 
                  Locked          =   -1  'True
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
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column19 
                  Alignment       =   1
                  ColumnWidth     =   1005.165
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   6660
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   18570
            _ExtentX        =   32755
            _ExtentY        =   11748
            _Version        =   393216
            AllowArrows     =   -1  'True
            ForeColor       =   128
            HeadLines       =   1
            RowHeight       =   21
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
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   20
            BeginProperty Column00 
               DataField       =   "Code"
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
            BeginProperty Column01 
               DataField       =   "name"
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
            BeginProperty Column02 
               DataField       =   "BUYSELL"
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
               DataField       =   "QNTY"
               Caption         =   "Qnty"
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
               DataField       =   "Saudacode"
               Caption         =   "SaudaCode"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "SaudaName"
               Caption         =   "Sauda Name"
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
            BeginProperty Column06 
               DataField       =   "Rate"
               Caption         =   "Rate"
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
            BeginProperty Column07 
               DataField       =   "Concode"
               Caption         =   "ConCode"
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
               DataField       =   "ConName"
               Caption         =   "Con. Name"
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
               DataField       =   "rate1"
               Caption         =   "Con. Rate"
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
            BeginProperty Column10 
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
            BeginProperty Column11 
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
            BeginProperty Column14 
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
            BeginProperty Column15 
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
            BeginProperty Column18 
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
            BeginProperty Column19 
               DataField       =   "diffamt"
               Caption         =   "ShreeAmt"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   2
               BeginProperty Column00 
                  Alignment       =   1
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   494.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   3195.213
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column10 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column14 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column15 
                  Locked          =   -1  'True
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
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column19 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo Saudacombo 
            Height          =   360
            Left            =   960
            TabIndex        =   45
            Top             =   0
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   4194304
            Text            =   "SAUDACOMBO"
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
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   23
         Top             =   8040
         Width           =   18975
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   13560
            TabIndex        =   46
            Top             =   0
            Width           =   5175
            Begin VB.TextBox Text11 
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
               Left            =   4200
               TabIndex        =   49
               Top             =   180
               Width           =   735
            End
            Begin VB.TextBox Text3 
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
               Left            =   1560
               TabIndex        =   47
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Shree Amount"
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
               TabIndex        =   50
               Top             =   250
               Width           =   1335
            End
            Begin VB.Label Label16 
               BackColor       =   &H00FFFFC0&
               Caption         =   "No. of Trades"
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
               Left            =   2880
               TabIndex        =   48
               Top             =   250
               Width           =   1335
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   495
            Left            =   7320
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   6255
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
               ForeColor       =   &H00400000&
               Height          =   405
               Left            =   4680
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   100
               Width           =   1335
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
               ForeColor       =   &H00400000&
               Height          =   405
               Left            =   600
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text8 
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
               ForeColor       =   &H00400000&
               Height          =   405
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   100
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00404000&
               BackStyle       =   0  'Transparent
               Caption         =   "Sell"
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
               Left            =   120
               TabIndex        =   42
               Top             =   180
               Width           =   330
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00404000&
               BackStyle       =   0  'Transparent
               Caption         =   "Sell Avg"
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
               Left            =   1800
               TabIndex        =   41
               Top             =   150
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Label Label12 
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
               Height          =   255
               Left            =   3840
               TabIndex        =   40
               Top             =   153
               Width           =   855
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   495
            Left            =   840
            TabIndex        =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   6255
            Begin VB.TextBox Text1 
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
               ForeColor       =   &H00400000&
               Height          =   405
               Left            =   600
               Locked          =   -1  'True
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   100
               Width           =   975
            End
            Begin VB.TextBox Text7 
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
               ForeColor       =   &H00400000&
               Height          =   405
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   100
               Visible         =   0   'False
               Width           =   1095
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
               ForeColor       =   &H00400000&
               Height          =   405
               Left            =   4680
               Locked          =   -1  'True
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   100
               Width           =   1335
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00404000&
               BackStyle       =   0  'Transparent
               Caption         =   "Buy"
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
               Left            =   120
               TabIndex        =   36
               Top             =   120
               Width           =   360
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00404000&
               BackStyle       =   0  'Transparent
               Caption         =   "Buy Avg"
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
               Left            =   1680
               TabIndex        =   35
               Top             =   120
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label Label10 
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
               Height          =   255
               Left            =   3840
               TabIndex        =   34
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00404000&
            BackStyle       =   0  'Transparent
            Caption         =   "Totals"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   285
            Width           =   540
         End
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
         Left            =   960
         TabIndex        =   13
         Top             =   960
         Width           =   11415
      End
      Begin VB.CommandButton cmdImportFromExcel 
         Caption         =   "..."
         Height          =   285
         Left            =   -360
         TabIndex        =   12
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   18855
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
            Height          =   375
            Left            =   2160
            TabIndex        =   2
            Top             =   120
            Width           =   2175
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
            ItemData        =   "FRM_NEW_SINGLE_ENTRY.frx":0019
            Left            =   17400
            List            =   "FRM_NEW_SINGLE_ENTRY.frx":0023
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   360
            Left            =   600
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
         Begin MSDataListLib.DataCombo Saudacmb 
            Bindings        =   "FRM_NEW_SINGLE_ENTRY.frx":003A
            Height          =   420
            Left            =   7485
            TabIndex        =   4
            Top             =   120
            Width           =   3495
            _ExtentX        =   6165
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
         Begin MSDataListLib.DataCombo PartyCmb 
            Bindings        =   "FRM_NEW_SINGLE_ENTRY.frx":0045
            Height          =   420
            Left            =   11880
            TabIndex        =   5
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
         Begin MSDataListLib.DataCombo DExCombo 
            Bindings        =   "FRM_NEW_SINGLE_ENTRY.frx":0050
            Height          =   420
            Left            =   5280
            TabIndex        =   3
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
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
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   4440
            TabIndex        =   51
            Top             =   195
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   11160
            TabIndex        =   29
            Top             =   195
            Width           =   480
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   6840
            TabIndex        =   11
            Top             =   195
            Width           =   570
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
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   16800
            TabIndex        =   10
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   195
            Width           =   435
         End
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   360
         Left            =   480
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4194304
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10320
         TabIndex        =   22
         Top             =   6840
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   1
         Left            =   1080
         Picture         =   "FRM_NEW_SINGLE_ENTRY.frx":005B
         Stretch         =   -1  'True
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F10 to open new party"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Left            =   1680
         TabIndex        =   16
         Top             =   840
         Width           =   1785
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   0
         Left            =   720
         Picture         =   "FRM_NEW_SINGLE_ENTRY.frx":0365
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1605
         TabIndex        =   15
         Top             =   5325
         Width           =   45
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   9180
      Left            =   75
      Top             =   720
      Width           =   19365
   End
End
Attribute VB_Name = "FRM_NEW_SINGLE_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte:                Dim FExCode As String:              Dim FParty As String:           Dim FSauda As String
Dim LConNo As Long:                     Dim LConSno As Long:                Dim LBillParties As String:     Dim LBillSaudas As String
Dim LDataImport As Byte:                Dim GRIDREC As ADODB.Recordset:     Dim RecEx As ADODB.Recordset:   Dim RECGRID As ADODB.Recordset:
Dim Rec_Account As ADODB.Recordset:     Dim AllSaudaRec As ADODB.Recordset: Dim SaudaRec As ADODB.Recordset:

Sub SaudaList()
If Check1.Value = 1 Then
    If LenB(FExCode) > 0 Then
        mysql = "SELECT EXCODE,EX_SYMBOL,SAUDACODE,ITEMCODE,MATURITY,LOT,INSTTYPE,OPTTYPE,STRIKEPRICE FROM SCRIPTMASTER WHERE  EXCODE='" & FExCode & "' AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,INSTTYPE,MATURITY"
    Else
        mysql = "SELECT EXCODE,EX_SYMBOL,SAUDACODE,ITEMCODE,MATURITY,LOT,INSTTYPE,OPTTYPE,STRIKEPRICE FROM SCRIPTMASTER WHERE MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,MATURITY"
    End If
    Set AllSaudaRec = Nothing: Set AllSaudaRec = New ADODB.Recordset
    AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AllSaudaRec.EOF Then
        Set Saudacmb.RowSource = AllSaudaRec
        Saudacmb.ListField = "SAUDACODE"
        Saudacmb.BoundColumn = "SAUDACODE"
        Set Saudacombo.RowSource = AllSaudaRec
        Saudacombo.ListField = "SAUDACODE"
        Saudacombo.BoundColumn = "SAUDACODE"
    End If
Else
    If LenB(FExCode) > 0 Then
        mysql = "SELECT SAUDAID,EXID,ITEMID,EXCODE,SAUDACODE,ITEMCODE,INSTTYPE,OPTTYPE,STRIKEPRICE,MATURITY,TRADEABLELOT  FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & FExCode & "' AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,INSTTYPE,MATURITY"
    Else
        mysql = "SELECT SAUDAID,EXID,ITEMID,EXCODE,SAUDACODE,ITEMCODE,INSTTYPE,OPTTYPE,STRIKEPRICE,MATURITY,TRADEABLELOT  FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,INSTTYPE,MATURITY"
    End If
    Set SaudaRec = Nothing: Set SaudaRec = New ADODB.Recordset
    SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not SaudaRec.EOF Then
        Set Saudacmb.RowSource = SaudaRec
        Saudacmb.ListField = "SAUDACODE"
        Saudacmb.BoundColumn = "SAUDACODE"
        Set Saudacombo.RowSource = SaudaRec
        Saudacombo.ListField = "SAUDACODE"
        Saudacombo.BoundColumn = "SAUDACODE"
    End If
End If
End Sub
Sub Add_Rec()
    If Rec_Account.RecordCount > 0 Then
        LDataImport = 0
        Frame1.Enabled = True: Combo1.ListIndex = 0: Frame7.Enabled = True:
        Call Get_Selection(1)
        If vcDTP1.Enabled Then vcDTP1.SetFocus
    Else
        Call CANCEL_REC
    End If
    LConNo = Get_Max_ConNo(DateValue(vcDTP1.Value), 0)
    LConNo = LConNo + 1
    RECGRID.AddNew
    RECGRID!DIMPORT = 0
    RECGRID!CONTIME = Time
    RECGRID!USERID = vbNullString
    RECGRID.Update
    LConNo = LConNo
    RECGRID!SrNo = LConNo  'RECGRID.AbsolutePosition
    If GCINNo = "2000" Then
        DataGrid2.Col = 0
    Else
        DataGrid1.Col = 0
    End If
End Sub
Sub Save_Rec()
    On Error GoTo err1
    Dim TRec As ADODB.Recordset:        Dim LTrades As Double:      Dim LTRD2 As Double:        Dim LExCode As String
    Dim MBCL As String:                 Dim MSCL  As String:        Dim LTrdNo As String:       Dim LConf As Boolean
    Dim BRate As Double:                Dim SRate As Double:        Dim LInstType As String:    Dim LOptType As String
    Dim LStrike As Double:              Dim MSAmt As Double:        Dim LBCODE As String:       Dim LSCode  As String
    Dim MBAmt  As Double:               Dim MSaudaCode  As String:  Dim LConfirm As Integer:    Dim RConfirm As Integer
    Dim LCITEM  As String:              Dim LCSauda  As String:     Dim LPattan  As String:     Dim LLOT As Double
    Dim LSConType As String:            Dim LLotWise As String:     Dim LSaudaID As Long:        Dim LExID As Integer
    Dim LItemID As Integer
    'validation
    If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.Enabled = True: vcDTP1.SetFocus: Exit Sub
    Cnn.BeginTrans
    CNNERR = True
    LConf = True
    If Fb_Press = 2 Then
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        mysql = "SELECT COUNT(*) AS  TRADES  FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
        If LenB(FExCode) > 0 Then mysql = mysql & " AND EXCODE ='" & FExCode & "'"
        If LenB(FSauda) > 0 Then mysql = mysql & " AND SAUDA='" & FSauda & "'"
        If LenB(FParty) > 0 Then mysql = mysql & " AND PARTY='" & FParty & "'"
        TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not TRec.EOF Then LTrades = IIf(IsNull(TRec!TRADES), 0, TRec!TRADES)
        LTRD2 = RECGRID.RecordCount * 2
        If (LTRD2) < (LTrades) Then LConf = MsgBox("You are deleting mores trades than entered in the grid. do you want to continue", vbYesNo)
    End If
    If LConf = True Then
        mysql = "DELETE  FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
        If LenB(FExCode) > 0 Then mysql = mysql & " AND EXCODE ='" & FExCode & "'"
        If LenB(FSauda) > 0 Then mysql = mysql & " AND SAUDA='" & FSauda & "'"
        If LenB(FParty) > 0 Then
            mysql = mysql & " AND CONNO IN (SELECT DISTINCT CONNO FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
            If LenB(FExCode) > 0 Then mysql = mysql & " AND EXCODE  ='" & FExCode & "'"
            If LenB(FSauda) > 0 Then mysql = mysql & " AND SAUDA='" & FSauda & "'"
            mysql = mysql & " AND PARTY='" & FParty & "')"
        End If
        Cnn.Execute mysql
        Set GRIDREC = Nothing: Set GRIDREC = New ADODB.Recordset
        Set GRIDREC = RECGRID.Clone
        GRIDREC.MoveFirst
        MSaudaCode = vbNullString:    MSAmt = 0:    MBAmt = 0
        While Not GRIDREC.EOF
            If Len(GRIDREC!NAME & vbNullString) > 0 And Len(GRIDREC!conName & vbNullString) > 0 Then
                If Val(GRIDREC!QNTY) > 0 And Val(GRIDREC!Rate) > 0 And Val(GRIDREC!Rate1) > 0 Then
                    LDataImport = IIf(IsNull(LDataImport), 0, LDataImport)
                    If GRIDREC!BUYSELL = "B" Then
                        LSConType = "B"
                        MBCL = GRIDREC!Code
                        MSCL = GRIDREC!CONCODE
                        LConfirm = Val(IIf(IsNull(GRIDREC!LConfirm), 0, GRIDREC!LConfirm))
                        RConfirm = Val(IIf(IsNull(GRIDREC!RConfirm), 0, GRIDREC!RConfirm))
                    Else
                        LSConType = "S"
                        MBCL = GRIDREC!CONCODE
                        MSCL = GRIDREC!Code
                        LConfirm = Val(IIf(IsNull(GRIDREC!RConfirm), 0, GRIDREC!RConfirm))
                        RConfirm = Val(IIf(IsNull(GRIDREC!LConfirm), 0, GRIDREC!LConfirm))
                    End If
                    Rec_Account.MoveFirst
                    Rec_Account.Find "AC_CODE='" & MBCL & "'"
                    If Rec_Account.EOF Then
                        MsgBox "Please Check Party Code " & MBCL & ""
                        CNNERR = False: Cnn.RollbackTrans
                        Exit Sub
                    End If
                    Rec_Account.MoveFirst
                    Rec_Account.Find "AC_CODE='" & MSCL & "'"
                    If Rec_Account.EOF Then
                        MsgBox "Please Check Party " & MSCL & ""
                        CNNERR = False: Cnn.RollbackTrans
                        Exit Sub
                    End If
                    If GRIDREC!BUYSELL = "B" Then
                        MBAmt = MBAmt + (Val(GRIDREC!QNTY & "") * (Round(Val(GRIDREC!Rate & ""), 2)) * GRIDREC!LOT)
                        LDataImport = Abs(GRIDREC!DIMPORT)
                        MSAmt = MSAmt + (Val(GRIDREC!QNTY & "") * Round(Val(GRIDREC!Rate1) & "", 2) * GRIDREC!LOT)
                    Else
                        MSAmt = MSAmt + (Val(GRIDREC!QNTY & "") * (Round(Val(GRIDREC!Rate & ""), 2)) * GRIDREC!LOT)
                        LDataImport = Abs(GRIDREC!DIMPORT)
                        MBAmt = MBAmt + (Val(GRIDREC!QNTY & "") * Round(Val(GRIDREC!Rate1) & "", 2) * GRIDREC!LOT)
                    End If
                    LCITEM = vbNullString
                    If Check1.Value = 1 Then
                        AllSaudaRec.MoveFirst
                        AllSaudaRec.Find "SAUDACODE='" & GRIDREC!saudacode & "'", , adSearchForward
                        If AllSaudaRec.EOF Then
                            MsgBox "Check Entry for  " & GRIDREC!saudacode & ""
                            Cnn.RollbackTrans: CNNERR = False
                            Exit Sub
                        Else
                            LExCode = AllSaudaRec!EXCODE
                            LExID = Get_ExID(LExCode)
                            LCITEM = Get_ItemMaster(LExID, AllSaudaRec!EX_SYMBOL)
                            If LenB(LCITEM) < 1 Then LCITEM = Create_TItemMast(AllSaudaRec!ITEMCODE, AllSaudaRec!ITEMCODE, AllSaudaRec!EX_SYMBOL, AllSaudaRec!LOT, AllSaudaRec!EXCODE)
                            LItemID = Get_ITEMID(LCITEM)
                            LCSauda = Get_SaudaMaster(LExID, LItemID, DateValue(AllSaudaRec!MATURITY), AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
                            If LenB(LCSauda) < 1 Then LCSauda = Create_TSaudaMast(LCITEM, AllSaudaRec!MATURITY, AllSaudaRec!EXCODE, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
                            LInstType = AllSaudaRec!INSTTYPE
                            LOptType = AllSaudaRec!OPTTYPE
                            LStrike = AllSaudaRec!STRIKEPRICE
                            RecEx.MoveFirst
                            RecEx.Find "EXCODE='" & LExCode & "'"
                            If Not RecEx.EOF Then LLotWise = RecEx!LOTWISE
                            LSaudaID = Get_SaudaID(LCSauda)
                        End If
                    Else
                        SaudaRec.MoveFirst
                        SaudaRec.Find "SAUDACODE='" & GRIDREC!saudacode & "'", , adSearchForward
                        If SaudaRec.EOF Then
                            MsgBox "Check Entry for  " & GRIDREC!saudacode & ""
                            Cnn.RollbackTrans: CNNERR = False
                            Exit Sub
                        Else
                            LExCode = SaudaRec!EXCODE
                            RecEx.MoveFirst
                            RecEx.Find "EXCODE='" & LExCode & "'"
                            If Not RecEx.EOF Then LLotWise = RecEx!LOTWISE
                            
                            LCITEM = SaudaRec!ITEMCODE
                            LInstType = SaudaRec!INSTTYPE
                            LOptType = SaudaRec!OPTTYPE
                            LStrike = SaudaRec!STRIKEPRICE
                            LItemID = SaudaRec!itemid
                            LSaudaID = SaudaRec!SAUDAID
                            LExID = SaudaRec!EXID
                        End If
                    End If
                    If Combo1.ListIndex = 0 Then
                        LPattan = "C"
                    Else
                        LPattan = "O"
                    End If
                    LLOT = Get_LotSize(LItemID, LSaudaID, LExID)
                    If InStr(LBillParties, "'" & LBCODE & "'") < 1 Then
                        If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ","
                        LBillParties = LBillParties & "'" & LBCODE & "'"
                    End If
                    If InStr(LBillParties, "'" & LSCode & "'") < 1 Then
                        If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ", "
                        LBillParties = LBillParties & "'" & LSCode & "'"
                    End If
                    If Len(LBillSaudas) > 0 Then
                        If LStr_Exists(LBillSaudas, Str(LSaudaID)) = False Then LBillSaudas = LBillSaudas & "," & Str(LSaudaID)
                    Else
                        LBillSaudas = Str(LSaudaID)
                    End If
                    'If InStr(LBillSaudas, GRIDREC!SAUDACODE) < 1 Then
                    '    If LenB(LBillSaudas) > 0 Then LBillSaudas = LBillSaudas & ","
                    '    LBillSaudas = LBillSaudas & "'" & GRIDREC!SAUDACODE & "'"
                    'End If
                    MSaudaCode = GRIDREC!saudacode
                    LSaudaID = Get_SaudaID(MSaudaCode)
                    LConSno = Get_ConSNo(vcDTP1.Value, GRIDREC!saudacode, LCITEM, LExCode, LSaudaID, LItemID, LExID)
                    If IsNull(GRIDREC!TRADE_NO) Then
                        LTrdNo = Val(GRIDREC!SrNo)
                    Else
                        LTrdNo = Val(GRIDREC!TRADE_NO)
                    End If
                    Call Add_To_Ctr_D2(LSConType, MBCL, Val(LConSno), Format(vcDTP1.Value, "YYYY/MM/DD"), Val(GRIDREC!SrNo), GRIDREC!saudacode, GRIDREC!ITEMCODE, GRIDREC!Code, Val(GRIDREC!QNTY), Val(GRIDREC!Rate), Val(GRIDREC!Rate1), GRIDREC!CONCODE, GRIDREC!CONTIME, vbNullString, vbNullString, LTrdNo, LExCode, LLOT, 0, vbNullString, LInstType, LOptType, LStrike, "0", "Y", LExID, LItemID, LSaudaID)
                    
                End If
            End If
            GRIDREC.MoveNext
        Wend
    End If
    
    Call RATE_TEST(vcDTP1.Value, , , FRM_NEW_SINGLE_ENTRY)
    Call Shree_Posting(DateValue(vcDTP1.Value))
    Call Update_Charges(LBillParties, vbNullString, LBillSaudas, vbNullString, vcDTP1.Value, vcDTP1.Value, True)
    Cnn.CommitTrans
    CNNERR = False
    Cnn.BeginTrans
    CNNERR = True
    If BILL_GENERATION(CDate(vcDTP1.Value), CDate(GFinEnd), LBillSaudas, LBillParties, vbNullString) Then
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
        
        Cnn.RollbackTrans: CNNERR = False
    End If
    
 End Sub
Sub CANCEL_REC()
    vcDTP1.Enabled = True:  Combo1.Enabled = True
    Call RecSet
    Fb_Press = 0:                 FSauda = vbNullString:      FParty = vbNullString
    FExCode = vbNullString:       Frame1.Enabled = True:      Frame7.Enabled = True:
    FExCode = vbNullString:       FSauda = vbNullString:      DExCombo.Enabled = True
    Check1.Enabled = True:        Saudacmb.Enabled = True:    Combo1.Enabled = True
    If GCINNo = "2000" Then
        DataGrid1.Visible = False
        DataGrid2.Visible = True
        Set DataGrid2.DataSource = RECGRID
        DataGrid2.Refresh
    Else
        DataGrid1.Visible = True
        DataGrid2.Visible = False
        Set DataGrid1.DataSource = RECGRID
        DataGrid1.Refresh
    End If
    
    Label2.Visible = False
    DataCombo3.Visible = False
    Call Get_Selection(10)
    Combo1.ListIndex = -1: Frame1.Enabled = False
End Sub
Function MODIFY_REC(LCondate As Date, LSaudaCode As String, LMEXCODE As String, LPattan As String) As Boolean
Dim LBuyAmt As Double:          Dim LSellAmt As Double
Dim TRec As ADODB.Recordset:    Dim LConTrd As Boolean
Dim LSaudaID As Long
Dim LGSauda As String
If Fb_Press = 1 Then
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PATTAN='" & Left$(LPattan, 1) & "'AND CONDATE='" & Format(LCondate, "YYYY/MM/DD") & "'  "
    If LenB(LMEXCODE) > 0 Then mysql = mysql & " AND EXCODE  ='" & LMEXCODE & "'"
    If LenB(LSaudaCode) > 0 Then mysql = mysql & " AND SAUDA ='" & LSaudaCode & "'"
    TRec.Open mysql, Cnn, , adLockReadOnly
    If TRec.EOF Then
        If GCINNo = "2000" Then
            DataGrid2.Col = 0
            DataGrid2.SetFocus
            If LenB(FSauda) > 0 Then
                DataGrid2.Columns(4).Locked = True
            Else
                DataGrid2.Columns(4).Locked = False
            End If
        Else
            DataGrid1.Col = 0
            DataGrid1.SetFocus
            If LenB(FSauda) > 0 Then
                DataGrid1.Columns(4).Locked = True
            Else
                DataGrid1.Columns(4).Locked = False
            End If
        End If
        Frame7.Enabled = False
        Fb_Press = 1
        MODIFY_REC = True
    Else
        If MsgBox("Contract Already Exist For Selected Criteria.Press OK To Modify The Existing Contracts.", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            Fb_Press = 2
        Else
            Call CANCEL_REC
            Exit Function
        End If
    End If
End If
    Call RecSet
    If Fb_Press = 1 Then
        RECGRID.AddNew
        RECGRID!DIMPORT = 0
        RECGRID!CONTIME = Time
        RECGRID!USERID = vbNullString
        RECGRID.Update
        LConNo = LConNo
        RECGRID!SrNo = LConNo  'RECGRID.AbsolutePosition
        If GCINNo = "2000" Then
            DataGrid2.Col = 0
            Set DataGrid2.DataSource = RECGRID: DataGrid2.ReBind: DataGrid2.Refresh
        Else
            DataGrid1.Col = 0
            Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        End If
    Else
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        mysql = "SELECT I.LOT,C.CONCODE,C.CONTYPE,C.PARTY,D.NAME,C.RATE,C.CONNO,C.ITEMCODE,C.SAUDA,C.CLCODE,C.INVNO,C.DATAIMPORT,C.CONTIME,C.QTY, C.USERID,I.EXCHANGECODE ,C.ROWNO1,C.CONFIRM,C.SAUDAID"
        mysql = mysql & " FROM CTR_D AS C ,ITEMMAST AS I , ACCOUNTD AS D WHERE C.COMPCODE =" & GCompCode & " AND C.COMPCODE = D.COMPCODE AND C.CONDATE='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' "
        mysql = mysql & " AND C.PARTY=D.AC_CODE AND I.COMPCODE=C.COMPCODE AND I.ITEMCODE=C.ITEMCODE AND C.PATTAN='" & Left$(LPattan, 1) & "' "
        If LenB(FExCode) > 0 Then mysql = mysql & "AND C.EXCODE  ='" & FExCode & "'"
        If LenB(FSauda) > 0 Then mysql = mysql & "AND C.SAUDA ='" & FSauda & "'"
        If LenB(FParty) > 0 Then mysql = mysql & " AND CONNO IN (SELECT DISTINCT CONNO FROM CTR_D WHERE COMPCODE= " & GCompCode & " AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND PARTY='" & FParty & "')"
        mysql = mysql & " ORDER BY C.CONNO,C.ROWNO"
        TRec.Open mysql, Cnn, , adLockReadOnly
        
        Dim FLAG_NEXT  As Boolean
        If Not TRec.EOF Then
            TRec.MoveFirst
            LBillParties = vbNullString:        LBillSaudas = vbNullString:
            While Not TRec.EOF
                LConTrd = False
                RECGRID.AddNew
                RECGRID!SrNo = TRec!CONNO 'RECGRID.AbsolutePositi
                RECGRID!TRADE_NO = Trim(CStr(TRec!ROWNO1))
                RECGRID!EXCODE = Trim(CStr(TRec!EXCHANGECODE))
                If Len(TRec!CONCODE & "") > 0 Then
                    If TRec!PARTY <> TRec!CONCODE Then
                        RECGRID!BUYSELL = TRec!CONTYPE
                        RECGRID!Code = TRec!PARTY
                        RECGRID!NAME = TRec!NAME
                        RECGRID!Rate = Round(TRec!Rate, 2)
                    Else
                        LConTrd = True
                        RECGRID!BUYSELL = "S"
                        RECGRID!CONCODE = TRec!PARTY
                        RECGRID!conName = TRec!NAME
                        RECGRID!Rate1 = Round(TRec!Rate, 2)
                    End If
                Else
                    RECGRID!BUYSELL = TRec!CONTYPE
                    RECGRID!Code = TRec!PARTY
                    RECGRID!NAME = TRec!NAME
                    RECGRID!Rate = Round(TRec!TRATE, 2)
                    RECGRID!Rate1 = Round(TRec!TRATE, 2)
                End If
                RECGRID!ITEMCODE = TRec!ITEMCODE
                RECGRID!EXCODE = TRec!EXCHANGECODE
                LSaudaID = TRec!SAUDAID
                RECGRID!LOT = Val(TRec!LOT)
                RECGRID!saudacode = TRec!Sauda
                RECGRID!LCLCODE = IIf(IsNull(TRec!CLCODE), "", TRec!CLCODE)
                RECGRID!QNTY = TRec!QTY
                RECGRID!LInvNo = TRec!invno
                RECGRID!LConfirm = Val(TRec!CONFIRM & "")
                If Not IsNull(TRec!DATAIMPORT) Then
                    If TRec!DATAIMPORT = True Then
                        RECGRID!DIMPORT = 1
                    Else
                        RECGRID!DIMPORT = 0
                    End If
                Else
                    RECGRID!DIMPORT = 0
                End If
                RECGRID!CONTIME = IIf(IsNull(TRec!CONTIME), Time, TRec!CONTIME)
                RECGRID!USERID = vbNullString
                If InStr(LBillParties, "'" & TRec!PARTY & "'") < 1 Then
                    If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ", "
                    LBillParties = LBillParties & "'" & TRec!PARTY & "'"
                End If
                If InStr(LBillParties, "'" & TRec!CONCODE & "'") < 1 Then
                    If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ", "
                    LBillParties = LBillParties & "'" & TRec!CONCODE & "'"
                End If
                If Len(LBillSaudas) > 0 Then
                    If LStr_Exists(LBillSaudas, Str(LSaudaID)) = False Then LBillSaudas = LBillSaudas & "," & Str(LSaudaID)
                Else
                    LBillSaudas = Str(LSaudaID)
                End If
                'If InStr(LBillSaudas, TRec!Sauda) < 1 Then
                '    If LenB(LBillSaudas) > 1 Then LBillSaudas = LBillSaudas & ","
                '    LBillSaudas = LBillSaudas & "'" & TRec!Sauda & "'"
                'End If
                TRec.MoveNext
                If InStr(LBillParties, "'" & TRec!PARTY & "'") < 1 Then
                    If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ","
                    LBillParties = LBillParties & "'" & TRec!PARTY & "'"
                End If
                If InStr(LBillParties, "'" & TRec!CONCODE & "'") = 0 Then
                    If LenB(LBillParties) > 0 Then LBillParties = LBillParties & ","
                    LBillParties = LBillParties & "'" & TRec!CONCODE & "'"
                End If
                If LenB(TRec!CONCODE) > 0 Then
                    If LConTrd = True Then
                        RECGRID!Code = TRec!PARTY
                        RECGRID!NAME = TRec!NAME
                        RECGRID!Rate = Round(TRec!Rate, 2)
                    Else
                        RECGRID!CONCODE = TRec!PARTY
                        RECGRID!conName = TRec!NAME
                        RECGRID!Rate1 = Round(TRec!Rate, 2)
                    End If
                Else
                    RECGRID!CONCODE = TRec!PARTY
                    RECGRID!conName = TRec!NAME
                End If
                RECGRID!RCLCODE = TRec!CLCODE & vbNullString
                RECGRID!RConfirm = TRec!CONFIRM & vbNullString
                RECGRID!RInvNo = TRec!invno
                If RECGRID!BUYSELL = "B" Then
                    LBuyAmt = RECGRID!QNTY * RECGRID!Rate * RECGRID!LOT
                    LSellAmt = RECGRID!QNTY * RECGRID!Rate1 * RECGRID!LOT
                Else
                    LBuyAmt = RECGRID!QNTY * RECGRID!Rate1 * RECGRID!LOT
                    LSellAmt = RECGRID!QNTY * RECGRID!Rate * RECGRID!LOT
                End If
                RECGRID!diffaMt = Val(LSellAmt - LBuyAmt)
                RECGRID.Update
                TRec.MoveNext
            Wend
            'LConNo= get_maxconno(
            LConNo = Get_Max_ConNo(DateValue(vcDTP1.Value), 0)
            LConNo = LConNo + 1
            If GCINNo = "2000" Then
                If FSauda <> "" Then
                    DataGrid2.Columns(4).Locked = True
                    Saudacmb.Enabled = False
                Else
                    DataGrid2.Columns(4).Locked = False
                End If
                Set DataGrid2.DataSource = RECGRID
                DataGrid2.ReBind:        DataGrid2.Col = 0
                Call DataGrid2_AfterColEdit(0)
            Else
                If FSauda <> "" Then
                    DataGrid1.Columns(4).Locked = True
                    Saudacmb.Enabled = False
                Else
                    DataGrid1.Columns(4).Locked = False
                End If
                Set DataGrid1.DataSource = RECGRID
                DataGrid1.ReBind:        DataGrid1.Col = 0
                Call DataGrid1_AfterColEdit(0)
            End If
            Frame7.Enabled = False
            MODIFY_REC = True
            If Fb_Press = 3 Then
                If MsgBox("Do You Want to Delete All Trades ", vbYesNo) = vbYes Then
                    Cnn.BeginTrans
                    CNNERR = True
                    If RECGRID.RecordCount > 0 Then
                        RECGRID.MoveFirst
                        Do While Not RECGRID.EOF
                            LConNo = RECGRID!CONNO
                            LGSauda = RECGRID!Sauda
                            mysql = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONNO=" & LConNo & " AND SAUDA ='" & LGSauda & "'"
                            Cnn.Execute mysql
                            RECGRID.MoveNext
                        Loop
                    End If
                    'Call Update_Charges(LBillParties, vbNullString, LBillSaudas, vbNullString, vcDTP1.Value, vcDTP1.Value, True)
                    Cnn.CommitTrans
                    CNNERR = False
                    Cnn.BeginTrans
                    CNNERR = True
                    If BILL_GENERATION(CDate(vcDTP1.Value), CDate(GFinEnd), vbNullString, vbNullString, vbNullString) Then
                        Cnn.CommitTrans: CNNERR = False
                    Else
                        Cnn.RollbackTrans: CNNERR = False
                    End If
                    'Call Chk_Billing
                    Call CANCEL_REC
                    Exit Function
                Else
                    Call CANCEL_REC
                    Exit Function
                End If
            End If
        End If
    End If
    Call Grid_Entry
End Function

Private Sub Combo1_LostFocus()
If MODIFY_REC(vcDTP1.Value, FSauda, FExCode, Combo1.text) Then

End If

End Sub

Private Sub DataCombo1_LostFocus()
If Combo1.Enabled = True Then Combo1.SetFocus

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim LBRate As Double:           Dim LSRate As Double:       Dim LBQty As Double
Dim LSQty As Double:            Dim LBAmt As Double:        Dim LSAmt As Double
Dim LTotBQty As Double:         Dim LTotSQty As Double:     Dim LLOT As Double:         Dim LDiffAmt As Double
Dim TRec As ADODB.Recordset
If Not RECGRID.EOF Then
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    Set TRec = RECGRID.Clone
    TRec.ActiveConnection = Nothing
    LSQty = 0: LBQty = 0: LSRate = 0: LBRate = 0: LDiffAmt = 0: LLOT = 0: LBAmt = 0: LSAmt = 0
    LTotBQty = 0: LTotSQty = 0
    TRec.Filter = adFilterNone
    'TRec.Filter = "SAUDACODE='" & RECGRID!SAUDACODE & "'"
    If Not TRec.EOF Then
        TRec.MoveFirst
        'Label3.Caption = (RECGRID!SAUDACODE & "")
        While Not TRec.EOF
            If TRec!BUYSELL = "B" Then
                LBRate = IIf(IsNull(TRec!Rate), 0, TRec!Rate)
                LSRate = IIf(IsNull(TRec!Rate1), 0, TRec!Rate1)
                LBQty = IIf(IsNull(TRec!QNTY), 0, TRec!QNTY)
                LSQty = LBQty
                LTotBQty = LTotBQty + LBQty
                LLOT = IIf(IsNull(TRec!LOT), 0, TRec!LOT)
                LBAmt = LBAmt + (LBQty * LBRate * LLOT)
                LSAmt = LSAmt + (LSRate * LSQty * LLOT)
            Else
                LSRate = IIf(IsNull(TRec!Rate), 0, TRec!Rate)
                LBRate = IIf(IsNull(TRec!Rate1), 0, TRec!Rate1)
                LSQty = IIf(IsNull(TRec!QNTY), 0, TRec!QNTY)
                LBQty = LSQty
                LTotSQty = LTotSQty + LSQty
                LBAmt = LBAmt + (LBQty * LBRate * LLOT)
                LSAmt = LSAmt + (LSRate * LSQty * LLOT)
            End If
            TRec.MoveNext
        Wend
    End If
    'Total Shree Caculation
    Text1.text = LTotBQty: Text4.text = LTotSQty
    Text9.text = Format(LBAmt, "0.00")
    Text10.text = Format(LSAmt, "0.00"):
    Text3.text = Format(LBAmt - LSAmt, "0.00")
End If

Text11.text = Val(RECGRID.RecordCount)
End Sub

Private Sub DataCombo3_GotFocus()
    Sendkeys "%{DOWN}"
    If GCINNo = "2000" Then
        If DataGrid2.Col = 2 Or DataGrid2.Col = 1 Then
            DataGrid2.Col = 1
            DataCombo3.Left = Val(1080)
            DataCombo3.Top = DataGrid2.Top + Val(DataGrid2.RowTop(DataGrid2.Row))
            Label2.Visible = True: Label2.Left = 1080
        ElseIf DataGrid2.Col = 7 Or DataGrid2.Col = 8 Then
            DataGrid2.Col = 8:
            DataCombo3.Top = DataGrid2.Top + Val(DataGrid2.RowTop(DataGrid2.Row))
            DataCombo3.Left = DataGrid2.Columns(7).Left
            Label2.Visible = True: Label2.Left = DataGrid2.Columns(9).Left
        End If
    Else
        If DataGrid1.Col = 2 Or DataGrid1.Col = 1 Then
            DataGrid1.Col = 1
            DataCombo3.Left = Val(1080)
            DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
            Label2.Visible = True: Label2.Left = 1080
        ElseIf DataGrid1.Col = 7 Or DataGrid1.Col = 8 Then
            DataGrid1.Col = 8:
            DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
            DataCombo3.Left = DataGrid1.Columns(7).Left
            Label2.Visible = True: Label2.Left = DataGrid1.Columns(9).Left
        End If
    
    End If
    
   Sendkeys "%{DOWN}"
   Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyPress(KeyAscii As Integer)
Dim TRec As ADODB.Recordset
If KeyAscii = 13 Then
    If InStr(DataCombo3.BoundText, "'") Then
        DataCombo3.BoundText = Replace(DataCombo3.BoundText, "'", "", 1, Len(DataCombo3.BoundText))
    End If
    If GCINNo = "2000" Then
        If DataGrid2.Col = 0 Or DataGrid2.Col = 1 Then
            If DataCombo3.BoundText <> "" Then
                Rec_Account.Filter = adFilterNone
                Rec_Account.Filter = "AC_CODE='" & DataCombo3.BoundText & "'"
                If Rec_Account.EOF Then
                    DataCombo3.BoundText = "'"
                    Rec_Account.Filter = adFilterNone
                Else
                    'LPtyHead = Rec_Account!PTYHEAD
                    Rec_Account.Filter = adFilterNone
                    RECGRID!Code = DataCombo3.BoundText
                    RECGRID!NAME = DataCombo3.text
                    RECGRID!USERID = vbNullString
                    DataGrid2.Col = 1
                    DataGrid2.SetFocus
                    DataCombo3.Visible = False: Label2.Visible = False
                End If
            End If
        ElseIf DataGrid2.Col = 8 Or DataGrid2.Col = 7 Then
            If DataCombo3.BoundText <> "" Then
                Rec_Account.Filter = adFilterNone
                Rec_Account.Filter = "AC_CODE='" & DataCombo3.BoundText & "'"
                If Rec_Account.EOF Then
                    DataCombo3.BoundText = "'"
                    Rec_Account.Filter = adFilterNone
                Else
                    Rec_Account.Filter = adFilterNone
                    RECGRID!CONCODE = DataCombo3.BoundText
                    RECGRID!conName = DataCombo3.text
                    RECGRID!USERID = vbNullString
                    DataCombo3.Visible = False: Label2.Visible = False
                    DataGrid2.Col = 8
                    DataGrid2.SetFocus
                End If
            End If
        End If
    Else
        If DataGrid1.Col = 1 Or DataGrid1.Col = 2 Then
            If DataCombo3.BoundText <> "" Then
                Rec_Account.Filter = adFilterNone
                Rec_Account.Filter = "AC_CODE='" & DataCombo3.BoundText & "'"
                If Rec_Account.EOF Then
                    DataCombo3.BoundText = ""
                    MsgBox "Please Select Valid Party"
                    RECGRID!Code = ""
                    DataGrid1.Col = 1
                    DataGrid1.SetFocus
                    Rec_Account.Filter = adFilterNone
                Else
                    'LPtyHead = Rec_Account!PTYHEAD
                    Rec_Account.Filter = adFilterNone
                    RECGRID!Code = DataCombo3.BoundText
                    RECGRID!NAME = DataCombo3.text
                    RECGRID!USERID = vbNullString
                    DataGrid1.Col = 2
                    DataGrid1.SetFocus
                    DataCombo3.Visible = False: Label2.Visible = False
                End If
            Else
                MsgBox "Please Select Party"
                RECGRID!Code = ""
                DataGrid1.Col = 1
                DataGrid1.SetFocus
            End If
        ElseIf DataGrid1.Col = 8 Or DataGrid1.Col = 7 Then
            If DataCombo3.BoundText <> "" Then
                Rec_Account.Filter = adFilterNone
                Rec_Account.Filter = "AC_CODE='" & DataCombo3.BoundText & "'"
                If Rec_Account.EOF Then
                    DataCombo3.BoundText = ""
                    MsgBox "Please Select Valid Contra Party"
                    RECGRID!CONCODE = ""
                    DataGrid1.Col = 7
                    DataGrid1.SetFocus
                    Rec_Account.Filter = adFilterNone
                Else
                    Rec_Account.Filter = adFilterNone
                    RECGRID!CONCODE = DataCombo3.BoundText
                    RECGRID!conName = DataCombo3.text
                    RECGRID!USERID = vbNullString
                    DataCombo3.Visible = False: Label2.Visible = False
                    DataGrid1.Col = 8
                    DataGrid1.SetFocus
                End If
            Else
                MsgBox "Please Select Contra Party"
                RECGRID!CONCODE = ""
                DataGrid1.Col = 7
                DataGrid1.SetFocus
            End If
        End If
    
    End If
ElseIf KeyAscii = 27 Then
    If GCINNo = "2000" Then
        DataGrid2.SetFocus
    Else
        DataGrid1.SetFocus
    End If
    
    DataCombo3.Visible = False: Label2.Visible = False
ElseIf KeyAscii = 121 Then   'F3  NEW PARTY
    GETACNT.Show
    GETACNT.ZOrder
    GETACNT.add_record
ElseIf KeyAscii = 18 Then
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
    Dim TempRec As ADODB.Recordset
    
    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)
    If ColIndex = Val(0) Then
        If UCase(Left$(Trim(DataGrid1.text), 1)) = "S" Or UCase(Left$(Trim(DataGrid1.text), 1)) = "B" Then
            DataGrid1.text = Left$(UCase(DataGrid1.text), 1)
            DataGrid1.Col = 0
        Else
           DataGrid1.text = "B"
           DataGrid1.Col = 0
        End If
    ElseIf ColIndex = 6 Then
        If Val(RECGRID!Rate & "") > 0 Then
            If Val(Round(RECGRID!Rate1, 2) & "") = Val(0) Then RECGRID!Rate1 = Round(RECGRID!Rate, 2)
        Else
            If ColIndex = 5 Then
            Else
                MsgBox "Rate can not be zero.Please enter rate.", vbCritical
                DataGrid1.Col = 6: DataGrid1.SetFocus
            End If
        End If
    ElseIf ColIndex = Val(5) Then
            mysql = "SELECT I.ITEMCODE,I.LOT,S.SAUDACODE,I.EXCHANGECODE FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & GCompCode & " AND I.COMPCODE=S.COMPCODE AND I.ITEMCODE=S.ITEMCODE AND S.SAUDACODE= '" & DataGrid1.text & "'"
        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open mysql, Cnn
        If TempRec.RecordCount > 0 Then
            TempRec.MoveFirst
            TempRec.Find "saudaCODE='" & DataGrid1.text & "'", , adSearchForward
            If Not TempRec.EOF Then
                RECGRID!ITEMCODE = TempRec!ITEMCODE
                RECGRID!saudacode = TempRec!saudacode
                RECGRID!LOT = TempRec!LOT
                RECGRID!EXCODE = TempRec!EXCHANGECODE
            Else
                RECGRID!saudacode = vbNullString
                DataGrid1.Col = 5
                Saudacombo.Visible = True
                Saudacombo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LConNoCHK As Double:    Dim TempRec As ADODB.Recordset:     Dim LBCODE As String
    Dim LBNAME As String:       Dim LItemCode As String:            Dim LLOT As Double
    Dim LSaudaCode As String:   Dim LLEXCODE As String:             Dim LBUYSELL As String
    Dim LSCode As String:       Dim LSNAME As String:               Dim LBAmt As Double
    Dim LSAmt  As Double:       Dim LDiffAmt  As Double
    If Not RECGRID.EOF Then
        If DataGrid1.Enabled = True Then
            If DataGrid1.Col = 6 Or DataGrid1.Col = 11 Then
                If KeyCode = 13 Or KeyCode = 9 Then
                    If Val(DataGrid1.text) = 0 Then
                        MsgBox "Rate Cannot Be Zero", vbCritical
                        DataGrid1.SetFocus
                        Exit Sub
                    End If
                End If
            ElseIf KeyCode = 13 And (DataGrid1.Col = 0) Then
                If UCase(Left$(Trim(DataGrid1.text), 1)) = "S" Or UCase(Left$(Trim(DataGrid1.text), 1)) = "B" Then
                    DataGrid1.text = Left$(UCase(DataGrid1.text), 1)
                    DataGrid1.SetFocus
                    DataGrid1.Col = 0
                Else
                    DataGrid1.text = "B"
                    DataGrid1.SetFocus
                    DataGrid1.Col = 0
                End If
            ElseIf KeyCode = 13 And ((DataGrid1.Col = 9) Or (DataGrid1.Col = 4 And FSauda <> "" And RECGRID!CONCODE <> "") Or (DataGrid1.Col = 6 And RECGRID!CONCODE <> "")) Then
                LBCODE = RECGRID!Code
                LBNAME = RECGRID!NAME
                LItemCode = RECGRID!ITEMCODE
                LLOT = RECGRID!LOT
                LSaudaCode = RECGRID!saudacode
                LLEXCODE = RECGRID!EXCODE
                LBUYSELL = RECGRID!BUYSELL
                LSCode = RECGRID!CONCODE
                LSNAME = RECGRID!conName
                DoEvents
                LBAmt = (RECGRID!LOT * RECGRID!QNTY * RECGRID!Rate)
                LSAmt = (RECGRID!LOT * RECGRID!QNTY * Val(DataGrid1.text))
                LDiffAmt = LBAmt - LSAmt
                RECGRID!diffaMt = LDiffAmt
                RECGRID.MoveNext
                If RECGRID.EOF Then
                    RECGRID.AddNew
                    RECGRID!Code = LBCODE
                    RECGRID!NAME = LBNAME
                    RECGRID!CONCODE = LSCode
                    RECGRID!conName = LSNAME
                    RECGRID!ITEMCODE = LItemCode
                    RECGRID!EXCODE = LLEXCODE
                    RECGRID!LOT = LLOT
                    RECGRID!BUYSELL = LBUYSELL
                    RECGRID!saudacode = LSaudaCode
                    RECGRID!QNTY = 0
                    RECGRID!Rate = 0
                    RECGRID!Rate1 = 0
                    RECGRID!DIMPORT = 0
                    RECGRID!USERID = vbNullString
                    RECGRID!CONTIME = Time
                    LConNoCHK = Get_Max_ConNo(DateValue(vcDTP1.Value), 0)
                    LConNo = LConNo + 1
                    If LConNoCHK > LConNo Then
                        MsgBox "Please Call Sauda Staff and Check Trade No"
                    End If
                    RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
                    RECGRID.Update
                End If
                DataGrid1.LeftCol = 0
                DataGrid1.Col = 0
            ElseIf DataGrid1.Col = Val(7) And (KeyCode = 13 Or KeyCode = 9) Then
                mysql = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode & " AND A.AC_CODE= '" & DataGrid1.text & "'"
                Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open mysql, Cnn
                If TempRec.RecordCount > 0 Then
                    RECGRID!CONCODE = TempRec!AC_CODE
                    RECGRID!conName = TempRec!NAME
                Else
                    RECGRID!CONCODE = vbNullString
                    DataGrid1.Col = 7
                    DataCombo3.Visible = True
                    DataCombo3.SetFocus
                    Exit Sub
                End If
            ElseIf DataGrid1.Col = Val(1) And (KeyCode = 13 Or KeyCode = 9) Then
                mysql = "SELECT A.AC_CODE,A.NAME,PTYHEAD FROM ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode & " AND A.AC_CODE= '" & DataGrid1.text & "'"
                Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open mysql, Cnn
                If TempRec.RecordCount > 0 Then
                    RECGRID!Code = TempRec!AC_CODE
                    RECGRID!NAME = TempRec!NAME
                Else
                    RECGRID!Code = vbNullString
                    DataGrid1.Col = 1
                    DataCombo3.Visible = True
                    DataCombo3.SetFocus
                    DataCombo3.SetFocus
                    Exit Sub
                End If
            ElseIf DataGrid1.Col = Val(3) And (KeyCode = 13 Or KeyCode = 9) Then
                If LenB(DataGrid1.text) > 1 Then
                    If Check1.Value = 1 Then
                        AllSaudaRec.MoveFirst
                        AllSaudaRec.Find "SAUDACODE='" & DataGrid1.text & "'"
                        If Not AllSaudaRec.EOF Then
                            RECGRID!ITEMCODE = AllSaudaRec!ITEMCODE
                            RECGRID!saudacode = AllSaudaRec!saudacode
                            RECGRID!LOT = 1
                            RECGRID!EXCODE = AllSaudaRec!EXCODE
                        Else
                            RECGRID!saudacode = vbNullString
                            DataGrid1.Col = 3
                            Saudacombo.Visible = True
                            Saudacombo.SetFocus
                        End If
                    Else
                        SaudaRec.MoveFirst
                        SaudaRec.Find "SAUDACODE='" & DataGrid1.text & "'"
                        If Not SaudaRec.EOF Then
                            RECGRID!ITEMCODE = SaudaRec!ITEMCODE
                            RECGRID!saudacode = SaudaRec!saudacode
                            RECGRID!LOT = 1
                            RECGRID!EXCODE = SaudaRec!EXCODE
                        Else
                            RECGRID!saudacode = vbNullString
                            DataGrid1.Col = 3
                            Saudacombo.Visible = True
                            Saudacombo.SetFocus
                        End If
                    End If
                Else
                    RECGRID!saudacode = vbNullString
                    DataGrid1.Col = 3
                    Saudacombo.Visible = True
                    Saudacombo.SetFocus
                End If
            ElseIf KeyCode = 114 Then   'F3  NEW PARTY
                GETACNT.Show
                GETACNT.ZOrder
                GETACNT.add_record
            ElseIf KeyCode = 46 And Shift = 2 Then
                'RECGRID.Delete
                If RECGRID.RecordCount = 0 Then
                    RECGRID.AddNew
                    LConNo = Get_Max_ConNo(DateValue(vcDTP1.Value), 0)
                    LConNo = LConNo + 1
                    RECGRID!SrNo = LConNo 'RECGRID.RecordCount
                    If Combo1.ListIndex = Val(1) Then
                        RECGRID!BRate = Round(Val(Text3.text), 2)
                        RECGRID!SRate = Round(Val(Text3.text), 2)
                        RECGRID!USERID = vbNullString
                    End If
                    RECGRID.Update
                End If
                Call DataGrid1_AfterColEdit(0)
            ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 0) Then
                If Len(Trim(DataGrid1.text)) < 1 Then
                    DataGrid1.Col = 0
                End If
            ElseIf (KeyCode = 13 Or KeyCode = 9) And ((DataGrid1.Col = 1) Or (DataGrid1.Col = 7)) Then
                If Len(Trim(DataGrid1.text)) < 1 Then
                    DataCombo3.Visible = True
                    DataCombo3.SetFocus
                End If
            ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 4) Then
            ElseIf (KeyCode = 13 Or KeyCode = 9) And DataGrid1.Col = 9 Then
                If Val(DataGrid1.text & "") <= 0 Then
                    MsgBox "Rate can not be zero.Please enter Rate.", vbCritical
                    DataGrid1.Col = 3: DataGrid1.SetFocus
                End If
        ElseIf KeyCode = 27 Then
            KeyCode = 0
        End If
        DataGrid1.CurrentCellVisible = True
    End If
End If
End Sub

Private Sub DExCombo_Click(Area As Integer)
Sendkeys "%{DOWN}"
End Sub

Private Sub DExCombo_Validate(Cancel As Boolean)
If LenB(DExCombo.BoundText) > 0 Then
    RecEx.MoveFirst
    RecEx.Find "EXCODE ='" & DExCombo.BoundText & "'", , adSearchForward
    If RecEx.EOF Then
        Cancel = True
        MsgBox "Select Valid Exchange"
        Exit Sub
    End If
    FExCode = DExCombo.BoundText
End If
    Call SaudaList
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Sendkeys "{tab}"
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
    Frame1.Enabled = False
    LDataImport = 0
    vcDTP1.Value = Date
    DataCombo3.Top = Val(2055): DataCombo3.Left = Val(1080)
    
    mysql = "SELECT EXCODE,SAUDACODE,EX_SYMBOL,ITEMCODE,LOT,MATURITY,INSTTYPE,OPTTYPE,STRIKEPRICE FROM SCRIPTMASTER ORDER BY ITEMCODE,INSTTYPE,MATURITY"
    Set AllSaudaRec = Nothing
    Set AllSaudaRec = New ADODB.Recordset
    AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    mysql = "SELECT EXCODE,LOTWISE FROM EXMAST WHERE COMPCODE=" & GCompCode & " ORDER BY EXCODE "
    Set RecEx = Nothing: Set RecEx = New ADODB.Recordset: RecEx.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RecEx.EOF Then
         Set DExCombo.RowSource = RecEx
         DExCombo.BoundColumn = "EXCODE"
         DExCombo.ListField = "EXCODE"
         
    End If
    
    DExCombo.Refresh
    Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
    Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Account.EOF Then
        Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
        Set PartyCmb.RowSource = Rec_Account:   PartyCmb.BoundColumn = "AC_CODE":    PartyCmb.ListField = "NAME"
    End If
        If GCINNo = "2000" Then
            Set DataGrid2.DataSource = RECGRID: DataGrid2.ReBind: DataGrid2.Refresh
        Else
            Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        End If
    
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    Set Rec_Account = Nothing
    Set Rec_Account = New ADODB.Recordset
    Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Account.EOF Then
        Set DataCombo3.RowSource = Rec_Account
        DataCombo3.BoundColumn = "AC_CODE"
        DataCombo3.ListField = "NAME"
        Set PartyCmb.RowSource = Rec_Account
        PartyCmb.BoundColumn = "AC_CODE"
        PartyCmb.ListField = "NAME"
    Else
        MsgBox "Please create customer account", vbInformation
        Call Get_Selection(12)
    End If
    If Fb_Press > 0 Then Call Get_Selection(Fb_Press)
End Sub

Private Sub Saudacmb_GotFocus()
Saudacmb.BoundText = vbNullString
Sendkeys "%{DOWN}"

End Sub

Private Sub partycmb_Validate(Cancel As Boolean)
If PartyCmb.BoundText <> "" Then
    Rec_Account.MoveFirst
    Rec_Account.Find "AC_CODE='" & PartyCmb.BoundText & "'"
    If Rec_Account.EOF Then
        FParty = ""
        MsgBox "Please Select Valid Party"
        Cancel = True
    Else
        FParty = Rec_Account!AC_CODE
    End If
End If

End Sub

Private Sub partycmb_GotFocus()
PartyCmb.BoundText = ""
Sendkeys "%{DOWN}"
End Sub

Private Sub Saudacmb_LostFocus()
    PartyCmb.SetFocus
End Sub

Private Sub Saudacmb_Validate(Cancel As Boolean)
If Check1.Value = 1 Then
    If LenB(Saudacmb.BoundText) > 1 Then
        AllSaudaRec.MoveFirst
        AllSaudaRec.Find "Saudacode ='" & Saudacmb.BoundText & "'"
        If Not AllSaudaRec.EOF Then
            FSauda = Saudacmb.BoundText
            RECGRID!saudacode = Saudacmb.BoundText
            Saudacombo.Visible = False
            If GCINNo = "2000" Then
                DataGrid2.Col = 0
                DataGrid2.Columns(12).Locked = True
                DataGrid2.SetFocus
            Else
                DataGrid1.Col = 0
                DataGrid1.Columns(12).Locked = True
                DataGrid1.SetFocus
            End If
        Else
            FSauda = vbNullString
            MsgBox "Please Select Valid Contract/Sauda"
            Cancel = True
        End If
    
    End If

Else
    If LenB(Saudacmb.BoundText) > 1 Then
        SaudaRec.MoveFirst
        SaudaRec.Find "Saudacode ='" & Saudacmb.BoundText & "'"
        If Not SaudaRec.EOF Then
            FSauda = Saudacmb.BoundText
            RECGRID!saudacode = Saudacmb.BoundText
            Saudacombo.Visible = False
            If GCINNo = "2000" Then
                DataGrid2.Col = 0
                DataGrid2.Columns(12).Locked = True
                DataGrid2.SetFocus
            Else
                DataGrid1.Col = 0
                DataGrid1.Columns(12).Locked = True
                DataGrid1.SetFocus
            End If
        Else
            FSauda = vbNullString
            MsgBox "Please Select Valid Contract/Sauda"
            Cancel = True
        End If
    
    End If
End If
End Sub

Private Sub SAUDACOMBO_GotFocus()
    Sendkeys "%{DOWN}"
    If GCINNo = "2000" Then
        Saudacombo.Left = DataGrid2.Columns(3).Left
        Saudacombo.Top = DataGrid2.Top + Val(DataGrid2.RowTop(DataGrid2.Row))
    Else
        Saudacombo.Left = DataGrid1.Columns(3).Left
        Saudacombo.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    End If
        
    Sendkeys "%{DOWN}"
    Sendkeys "%{DOWN}"
End Sub

Private Sub Saudacombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If LenB(Saudacombo.BoundText) > 0 Then
        RECGRID!saudacode = Saudacombo.BoundText
        If Check1.Value = 1 Then
            AllSaudaRec.MoveFirst
            AllSaudaRec.Find "SAUDACODE='" & Saudacombo.BoundText & "'", , adSearchForward
            If Not AllSaudaRec.EOF Then
                RECGRID!ITEMCODE = AllSaudaRec!ITEMCODE
                RECGRID!LOT = 1
                RECGRID!EXCODE = AllSaudaRec!EXCODE
                Saudacombo.Visible = False
            Else
                MsgBox "Please Check Entry for " & Saudacombo.BoundText & ""
                Exit Sub
            End If
        Else
            SaudaRec.MoveFirst
            SaudaRec.Find "SAUDACODE='" & Saudacombo.BoundText & "'"
            If Not SaudaRec.EOF Then
                RECGRID!ITEMCODE = SaudaRec!ITEMCODE
                RECGRID!LOT = 1
                RECGRID!EXCODE = SaudaRec!EXCODE
                Saudacombo.Visible = False
            Else
                MsgBox "Please Check Entry for " & Saudacombo.BoundText & ""
                Exit Sub
            End If
        End If
        If GCINNo = "2000" Then
            DataGrid2.Col = 4
            DataGrid2.SetFocus
        Else
            DataGrid1.Col = 4
            DataGrid1.SetFocus
        End If
        Saudacombo.Visible = False
    End If
End If


End Sub

Private Sub Saudacombo_Validate(Cancel As Boolean)
    If Saudacombo.Visible = True Then
        Cancel = True
    End If
    
End Sub

Private Sub Text3_GotFocus()
     Text3.SelLength = Len(Text3.text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Sub RecSet()
    
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BUYSELL", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "CODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "NAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "QNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "CONCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "CONNAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "RATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LOT", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
    RECGRID.Fields.Append "ITEMNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SAUDACODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SAUDANAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "RATE1", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "RInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "DImport", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "CONTIME", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "UserId", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "LCLCODE", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "RCLCODE", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "ORDER_NO", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "TRADE_NO", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "DIFFAMT", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "EXCODE", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "LCONFIRM", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "RCONFIRM", adInteger, , adFldIsNullable
    
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    Text3.text = Format(Text3.text, "0.00")
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
Function GetCloseRate() As Boolean
End Function
Private Sub Grid_Entry()
    vcDTP1.Enabled = False
    Text2.Enabled = False
    Saudacmb.Enabled = False
    Check1.Enabled = False
    Combo1.Enabled = False
    Frame7.Enabled = False
    Frame1.Enabled = True
    DoEvents
    If GCINNo = "2000" Then
        DataGrid2.SetFocus
    Else
        DataGrid1.SetFocus
    End If
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static OLDVAL As Integer
    Select Case ColIndex
    Case 0
        If OLDVAL = -1 Then
            RECGRID.Sort = "BUYSELL DESC"
        Else
            RECGRID.Sort = "BUYSELL"
        End If
    
    Case 1
        If OLDVAL = -1 Then
            RECGRID.Sort = "CODE DESC"
        Else
            RECGRID.Sort = "CODE"
        End If

    Case 2
        If OLDVAL = -1 Then
            RECGRID.Sort = "NAME DESC"
        Else
            RECGRID.Sort = "NAME"
        End If
    Case 3
        If OLDVAL = -1 Then
            RECGRID.Sort = "SAUDACODE DESC"
        Else
            RECGRID.Sort = "SAUDACODE"
        End If

    Case 5
        If OLDVAL = -1 Then
            RECGRID.Sort = "QNTY DESC"
        Else
            RECGRID.Sort = "QNTY"
        End If

    Case 6
        If OLDVAL = -1 Then
            RECGRID.Sort = "RATE DESC"
        Else
            RECGRID.Sort = "RATE"
        End If
    Case 6
        If OLDVAL = -1 Then
            RECGRID.Sort = "CONCODE DESC"
        Else
            RECGRID.Sort = "CONCODE"
        End If

    Case 7
        If OLDVAL = -1 Then
            RECGRID.Sort = "CONNAME DESC"
        Else
            RECGRID.Sort = "CONNAME"
        End If

    Case 8
        If OLDVAL = -1 Then
            RECGRID.Sort = "CONNAME DESC"
        Else
            RECGRID.Sort = "CONNAME"
        End If
    
    Case 9
        If OLDVAL = -1 Then
            RECGRID.Sort = "RATE1 DESC"
        Else
            RECGRID.Sort = "RATE1"
        End If
    
    Case 12
        If OLDVAL = -1 Then
            RECGRID.Sort = "CONTIME DESC"
        Else
            RECGRID.Sort = "CONTIME"
        End If
    
    Case 10
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
    
    Case 19
        If OLDVAL = -1 Then
            RECGRID.Sort = "DiffAmt DESC"
        Else
            RECGRID.Sort = "DiffAmt"
        End If
    Case 18
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
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
End Sub

Public Sub DataGrid2_AfterColEdit(ByVal ColIndex As Integer)
    Dim TempRec  As ADODB.Recordset
    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)
    If ColIndex = Val(2) Then
    ElseIf ColIndex = 6 Then
        If Val(RECGRID!Rate & "") > 0 Then
            If Val(Round(RECGRID!Rate1, 2) & "") = Val(0) Then RECGRID!Rate1 = Round(RECGRID!Rate, 2)
        Else
            If ColIndex = 5 Then
            Else
                MsgBox "Rate can not be zero.Please enter rate.", vbCritical
                DataGrid2.Col = 6: DataGrid2.SetFocus
            End If
        End If
    ElseIf ColIndex = Val(5) Then
        mysql = "SELECT I.ITEMCODE,I.LOT,S.SAUDACODE,I.EXCHANGECODE FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & GCompCode & " AND I.COMPCODE=S.COMPCODE AND I.ITEMCODE=S.ITEMCODE AND S.SAUDACODE= '" & DataGrid2.text & "'"
        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open mysql, Cnn
        If TempRec.RecordCount > 0 Then
            TempRec.MoveFirst
            TempRec.Find "SAUDACODE='" & DataGrid2.text & "'", , adSearchForward
            If Not TempRec.EOF Then
                RECGRID!ITEMCODE = TempRec!ITEMCODE
                RECGRID!saudacode = TempRec!saudacode
                RECGRID!LOT = TempRec!LOT
                RECGRID!EXCODE = TempRec!EXCHANGECODE
            Else
                RECGRID!saudacode = vbNullString
                DataGrid2.Col = 5
                Saudacombo.Visible = True
                Saudacombo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim LConNoCHK As Double:        Dim LBCODE As String:       Dim TempRec As ADODB.Recordset:     Dim LBNAME     As String
Dim LItemCode As String:           Dim LLOT As Double:                 Dim LSaudaCode  As String
   Dim LBUYSELL  As String:    Dim LSCode  As String:              Dim LSNAME  As String
            Dim LLEXCODE  As String
            
            Dim LBAmt  As Double
            Dim LSAmt  As Double
            Dim LDiffAmt As Double
            
    If DataGrid2.Enabled = True Then
        If DataGrid2.Col = 6 Or DataGrid2.Col = 11 Then
            If KeyCode = 13 Or KeyCode = 9 Then
                If Val(DataGrid2.text) = 0 Then
                    MsgBox "Rate Cannot Be Zero", vbCritical
                    DataGrid2.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf KeyCode = 13 And (DataGrid2.Col = 2) Then
            If UCase(Left$(Trim(DataGrid2.text), 1)) = "S" Or UCase(Left$(Trim(DataGrid2.text), 1)) = "B" Then
                DataGrid2.text = Left$(UCase(DataGrid2.text), 1)
                DataGrid2.SetFocus
                DataGrid2.Col = 2
            Else
                DataGrid2.text = "B"
                DataGrid2.SetFocus
                DataGrid2.Col = 2
            End If
        ElseIf KeyCode = 13 And ((DataGrid2.Col = 9) Or (DataGrid2.Col = 4 And FSauda <> "" And RECGRID!CONCODE <> "") Or (DataGrid2.Col = 6 And RECGRID!CONCODE <> "")) Then
            LBCODE = RECGRID!Code
            LBNAME = RECGRID!NAME
            LItemCode = RECGRID!ITEMCODE
            LLOT = RECGRID!LOT
            LSaudaCode = RECGRID!saudacode
            LBUYSELL = RECGRID!BUYSELL
            LSCode = RECGRID!CONCODE
            LSNAME = RECGRID!conName
            LLEXCODE = RECGRID!EXCODE
            DoEvents
            LBAmt = (RECGRID!LOT * RECGRID!QNTY * RECGRID!Rate)
            LSAmt = (RECGRID!LOT * RECGRID!QNTY * Val(RECGRID!Rate1))
            LDiffAmt = LBAmt - LSAmt
            RECGRID!diffaMt = LDiffAmt
            RECGRID.MoveNext
            If RECGRID.EOF Then
                RECGRID.AddNew
                RECGRID!Code = LBCODE
                RECGRID!NAME = LBNAME
                RECGRID!CONCODE = LSCode
                RECGRID!conName = LSNAME
                RECGRID!ITEMCODE = LItemCode
                RECGRID!EXCODE = LLEXCODE
                RECGRID!LOT = LLOT
                RECGRID!BUYSELL = LBUYSELL
                RECGRID!saudacode = LSaudaCode
                RECGRID!QNTY = 0
                RECGRID!Rate = 0
                RECGRID!Rate1 = 0
                RECGRID!DIMPORT = 0
                RECGRID!USERID = vbNullString
                RECGRID!CONTIME = Time
                LConNoCHK = Get_Max_ConNo(DateValue(vcDTP1.Value), 0)
                
                LConNo = LConNo + 1
                If LConNoCHK > LConNo Then
                    MsgBox "Please Call Sauda Staff and Check Trade No"
                End If

                RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
                RECGRID.Update
            End If
            
            DataGrid2.LeftCol = 0
            DataGrid2.Col = 0
        ElseIf DataGrid2.Col = Val(7) And (KeyCode = 13 Or KeyCode = 9) Then
            mysql = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode & " AND A.AC_CODE= '" & DataGrid2.text & "'"
            Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open mysql, Cnn
            If TempRec.RecordCount > 0 Then
                RECGRID!CONCODE = TempRec!AC_CODE
                RECGRID!conName = TempRec!NAME
            Else
                RECGRID!CONCODE = vbNullString
                DataGrid2.Col = 7
                DataCombo3.Visible = True
                DataCombo3.SetFocus
                DataCombo3.SetFocus
                Exit Sub
            End If
        ElseIf DataGrid2.Col = Val(0) And (KeyCode = 13 Or KeyCode = 9) Then
            
            
            mysql = "SELECT A.AC_CODE,A.NAME,PTYHEAD FROM ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode & " AND A.AC_CODE= '" & DataGrid2.text & "'"
            Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open mysql, Cnn
            If TempRec.RecordCount > 0 Then
                RECGRID!Code = TempRec!AC_CODE
                RECGRID!NAME = TempRec!NAME
                'LPtyHead = TempRec!PTYHEAD
            Else
                RECGRID!Code = ""
                DataGrid2.Col = 1
                DataCombo3.Visible = True
                DataCombo3.SetFocus
                DataCombo3.SetFocus
                Exit Sub
            End If
        ElseIf DataGrid2.Col = Val(4) And (KeyCode = 13 Or KeyCode = 9) Then
            mysql = "SELECT I.ITEMCODE,I.LOT,S.SAUDACODE,I.EXCHANGECODE FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & GCompCode & " AND I.COMPCODE=S.COMPCODE AND I.ITEMCODE=S.ITEMCODE AND S.SAUDACODE= '" & DataGrid2.text & "'"
            Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open mysql, Cnn
            If TempRec.RecordCount > 0 Then
                TempRec.MoveFirst
                TempRec.Find "saudaCODE='" & DataGrid2.text & "'", , adSearchForward
                If Not TempRec.EOF Then
                    RECGRID!ITEMCODE = TempRec!ITEMCODE
                    RECGRID!saudacode = TempRec!saudacode
                    RECGRID!LOT = TempRec!LOT
                    RECGRID!EXCODE = TempRec!EXCHANGECODE
                Else
                    RECGRID!saudacode = vbNullString
                    DataGrid2.Col = 4
                    Saudacombo.Visible = True
                    Saudacombo.SetFocus
                End If
            Else
                RECGRID!saudacode = vbNullString
                DataGrid2.Col = 4
                Saudacombo.Visible = True
                Saudacombo.SetFocus
            End If
        ElseIf KeyCode = 114 Then   'F3  NEW PARTY
            GETACNT.Show
            GETACNT.ZOrder
            GETACNT.add_record
        ElseIf KeyCode = 46 And Shift = 2 Then
            'RECGRID.Delete
            If RECGRID.RecordCount = 0 Then
                RECGRID.AddNew
                LConNo = Get_Max_ConNo(DateValue(vcDTP1.Value), 0)
                LConNo = LConNo + 1
                RECGRID!SrNo = LConNo 'RECGRID.RecordCount
                If Combo1.ListIndex = Val(1) Then
                    RECGRID!BRate = Round(Val(Text3.text), 2)
                    RECGRID!SRate = Round(Val(Text3.text), 2)
                    RECGRID!USERID = vbNullString
                End If
                RECGRID.Update
            End If
            Call DataGrid2_AfterColEdit(0)
        ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid2.Col = 0) Then
            If Len(Trim(DataGrid2.text)) < 1 Then
               DataGrid2.Col = 0
            Else
            End If
        ElseIf (KeyCode = 13 Or KeyCode = 9) And ((DataGrid2.Col = 0) Or (DataGrid2.Col = 7)) Then
            If Len(Trim(DataGrid2.text)) < 1 Then
                DataCombo3.Visible = True
                DataCombo3.SetFocus
            Else
            End If
        ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid2.Col = 4) Then
            If Len(Trim(DataGrid2.text)) < 1 Then
                Saudacombo.Visible = True
                Saudacombo.SetFocus
            End If
        ElseIf (KeyCode = 13 Or KeyCode = 9) And DataGrid2.Col = 9 Then
            If Val(DataGrid2.text & "") <= 0 Then
                MsgBox "Rate can not be zero.Please enter Rate.", vbCritical
                DataGrid2.Col = 3: DataGrid2.SetFocus
            End If
    ElseIf KeyCode = 27 Then
        KeyCode = 0
    End If
    DataGrid2.CurrentCellVisible = True
End If
End Sub


Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim LBRate As Double:       Dim LSRate As Double:       Dim LBQty As Double:        Dim LSQty As Double
Dim LBAmt As Double:        Dim LSAmt As Double:        Dim LTotBQty As Double:     Dim LTotSQty As Double
Dim LLOT As Double:         Dim LDiffAmt As Double:     Dim TRec As ADODB.Recordset
If Not RECGRID.EOF Then
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    Set TRec = RECGRID.Clone
    TRec.ActiveConnection = Nothing
    LSQty = 0: LBQty = 0: LSRate = 0: LBRate = 0: LDiffAmt = 0: LLOT = 0: LBAmt = 0: LSAmt = 0
    LTotBQty = 0: LTotSQty = 0
    TRec.Filter = adFilterNone
    'TRec.Filter = "SAUDACODE='" & RECGRID!SAUDACODE & "'"
    If Not TRec.EOF Then
        TRec.MoveFirst
        'Label3.Caption = (RECGRID!SAUDACODE & "")
        While Not TRec.EOF
            If TRec!BUYSELL = "B" Then
                LBRate = IIf(IsNull(TRec!Rate), 0, TRec!Rate)
                LBQty = IIf(IsNull(TRec!QNTY), 0, TRec!QNTY)
                LTotBQty = LTotBQty + LBQty
                LLOT = IIf(IsNull(TRec!LOT), 0, TRec!LOT)
                LBAmt = LBAmt + (LBQty * LBRate * LLOT)
                
            Else
                LSRate = IIf(IsNull(TRec!Rate), 0, TRec!Rate)
                LSQty = IIf(IsNull(TRec!QNTY), 0, TRec!QNTY)
                LTotSQty = LTotSQty + LSQty
                LSAmt = LSAmt + (LSRate * LSQty * LLOT)
            End If
            TRec.MoveNext
        Wend
    End If
    'Total Shree Caculation
    Text1.text = LTotBQty: Text4.text = LTotSQty
    
    Text9.text = Format(LBAmt, "0.00")
    Text10.text = Format(LSAmt, "0.00"):
    Text3.text = Format(LBAmt - LSAmt, "0.00")
End If
End Sub


Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
    Static OLDVAL As Integer
    Select Case ColIndex
    Case 2
        If OLDVAL = -1 Then
            RECGRID.Sort = "CONTYPE DESC"
        Else
            RECGRID.Sort = "CONTYPE"
        End If
    
    Case 0
        If OLDVAL = -1 Then
            RECGRID.Sort = "BCODE DESC"
        Else
            RECGRID.Sort = "BCODE"
        End If

    Case 1
        If OLDVAL = -1 Then
            RECGRID.Sort = "BNAME DESC"
        Else
            RECGRID.Sort = "BNAME"
        End If
    Case 4
        If OLDVAL = -1 Then
            RECGRID.Sort = "SAUDACODE DESC"
        Else
            RECGRID.Sort = "SAUDACODE"
        End If

    Case 3
        If OLDVAL = -1 Then
            RECGRID.Sort = "BQNTY DESC"
        Else
            RECGRID.Sort = "BQNTY"
        End If

    Case 5
        If OLDVAL = -1 Then
            RECGRID.Sort = "BRATE DESC"
        Else
            RECGRID.Sort = "BRATE"
        End If
    Case 6
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
End Sub
