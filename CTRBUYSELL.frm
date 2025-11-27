VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CTRBUYSELL 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Contract Entry"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   Begin VB.TextBox txtITEMID 
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
      Left            =   7680
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   9960
      Width           =   3255
   End
   Begin VB.TextBox TxtSaudaCode 
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
      Left            =   14160
      MaxLength       =   50
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
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
      TabIndex        =   32
      Top             =   0
      Width           =   13935
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
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
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -120
         TabIndex        =   33
         Top             =   0
         Width           =   14055
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Contract Entry"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   120
            Width           =   13695
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
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
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   13575
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
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
         Height          =   7845
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   13260
         Begin VB.TextBox TxtItem 
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   735
            Width           =   1935
         End
         Begin VB.TextBox TxtStrike 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   735
            Width           =   975
         End
         Begin VB.TextBox TxtOptType 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   735
            Width           =   975
         End
         Begin VB.TextBox TxtInst 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   735
            Width           =   975
         End
         Begin VB.TextBox TxtExCode 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Show All Contract"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2640
            TabIndex        =   2
            Top             =   15
            Width           =   1215
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
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
            TabIndex        =   19
            Top             =   7080
            Width           =   12975
            Begin VB.TextBox Text11 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   10450
               Locked          =   -1  'True
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   829
               Locked          =   -1  'True
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   5105
               Locked          =   -1  'True
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   9261
               Locked          =   -1  'True
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   11640
               Locked          =   -1  'True
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   2018
               Locked          =   -1  'True
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   6294
               Locked          =   -1  'True
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text9 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   3207
               Locked          =   -1  'True
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text10 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   7483
               Locked          =   -1  'True
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000003&
               BackStyle       =   0  'Transparent
               Caption         =   "Diff"
               ForeColor       =   &H00800080&
               Height          =   240
               Left            =   8672
               TabIndex        =   31
               Top             =   180
               Width           =   375
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000003&
               BackStyle       =   0  'Transparent
               Caption         =   "Sale"
               ForeColor       =   &H00800080&
               Height          =   240
               Left            =   4396
               TabIndex        =   30
               Top             =   180
               Width           =   495
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000003&
               BackStyle       =   0  'Transparent
               Caption         =   "Buy"
               ForeColor       =   &H00800080&
               Height          =   480
               Left            =   120
               TabIndex        =   29
               Top             =   180
               Width           =   495
            End
         End
         Begin VB.CommandButton cmdImportFromExcel 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13320
            TabIndex        =   10
            Top             =   5760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            ItemData        =   "CTRBUYSELL.frx":0000
            Left            =   8880
            List            =   "CTRBUYSELL.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   165
            Width           =   1575
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   11400
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   735
            Width           =   1545
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
            TabIndex        =   8
            Top             =   1200
            Width           =   11415
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   360
            Left            =   11400
            TabIndex        =   5
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   64
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
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   360
            Left            =   7200
            TabIndex        =   11
            Top             =   1680
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   64
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5580
            Left            =   120
            TabIndex        =   6
            Top             =   1440
            Width           =   12930
            _ExtentX        =   22807
            _ExtentY        =   9843
            _Version        =   393216
            AllowArrows     =   -1  'True
            ForeColor       =   128
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
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   17
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "Contype"
               Caption         =   "BuySell"
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
               DataField       =   "BCODE"
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
            BeginProperty Column03 
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
            BeginProperty Column04 
               DataField       =   "BQNTY"
               Caption         =   "Qnty"
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
            BeginProperty Column05 
               DataField       =   "BRATE"
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
            BeginProperty Column06 
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
            BeginProperty Column07 
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
            BeginProperty Column08 
               DataField       =   "SQNTY"
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
            BeginProperty Column09 
               DataField       =   "SRATE"
               Caption         =   "BrokRate"
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
            BeginProperty Column11 
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
            BeginProperty Column12 
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
            BeginProperty Column13 
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   2
               BeginProperty Column00 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  Object.Visible         =   0   'False
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column07 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   2294.929
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   915.024
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
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
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
            EndProperty
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   360
            Left            =   4560
            TabIndex        =   3
            Top             =   165
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   64
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
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   405
            Left            =   960
            TabIndex        =   1
            Top             =   165
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37860.8625462963
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Strike"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8760
            TabIndex        =   44
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "OptType"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6600
            TabIndex        =   43
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "InstType "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   42
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ExCode"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   780
            Width           =   720
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
            Left            =   10560
            TabIndex        =   18
            Top             =   195
            Width           =   630
         End
         Begin VB.Image Image1 
            Height          =   195
            Index           =   0
            Left            =   720
            Picture         =   "CTRBUYSELL.frx":0021
            Stretch         =   -1  'True
            Top             =   1230
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F10 to open new party"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   17
            Top             =   1215
            Width           =   2115
         End
         Begin VB.Image Image1 
            Height          =   195
            Index           =   1
            Left            =   1080
            Picture         =   "CTRBUYSELL.frx":032B
            Stretch         =   -1  'True
            Top             =   1230
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   18
            Left            =   8280
            TabIndex        =   16
            Top             =   225
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cl Rate"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   4
            Left            =   10560
            TabIndex        =   15
            Top             =   795
            Width           =   750
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   3
            Left            =   1920
            TabIndex        =   14
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   2
            Left            =   3840
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   465
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   9840
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
            Picture         =   "CTRBUYSELL.frx":0635
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CTRBUYSELL.frx":0A87
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   8265
      Left            =   0
      Top             =   720
      Width           =   13845
   End
End
Attribute VB_Name = "CTRBUYSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean:                        Dim LParty As String:               Dim LConNo As Long:                         Dim LConSno As Long
Dim LPattan As String:                      Dim LUserId As String:              Dim LContractAcc As String:                 Dim LExCode  As String
Dim LConType As String:                     Dim LParties As String:             Dim LDataImport As Byte:                    Dim VchNo As String
Dim FLOWDIR As Byte:                        Dim GRIDPOS As Byte:                Public Fb_Press As Byte:                    Dim RecEx As ADODB.Recordset
Dim RECGRID As ADODB.Recordset:             Dim Rec_Sauda As ADODB.Recordset:   Dim Rec_Account As ADODB.Recordset:         Dim REC_CloRate As ADODB.Recordset
Dim OldDate As Date:                        Dim FCalval As Double
Sub Add_Rec()
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
    On Error GoTo err1
    'validation
    Dim LSExCode As String:         Dim LSItemCode As String:       Dim LInstType  As String:       Dim LOptType As String
    Dim LStrike As Double:          Dim TSRec As ADODB.Recordset:   Dim TRec As ADODB.Recordset:    Dim LSaudaCode As String
    Dim MBAmt As Double:            Dim MSAmt As Double:            Dim LSaudaID As Long:        Dim MCL As String
    Dim LExID As Integer:           Dim LItemID As Integer
    Dim LACCID As Long
    If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If Val(Text1.text) + Val(Text4.text) = 0 Then MsgBox "Please Check Entries.", vbCritical: Exit Sub
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT EXID,SAUDAID,ITEMID,SAUDACODE,SAUDANAME,ITEMCODE,EXCODE,TRADEABLELOT,INSTTYPE,OPTTYPE,STRIKEPRICE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDACODE='" & TxtSaudaCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        MsgBox "Invalid Sauda Code.", vbExclamation, "Error": DataCombo1.SetFocus: Exit Sub
    Else
        LSExCode = TRec!excode:        LSItemCode = TRec!ITEMCODE
        LInstType = TRec!INSTTYPE:     LOptType = TRec!OPTTYPE
        LStrike = TRec!STRIKEPRICE:    LSaudaID = TRec!SAUDAID
        LExID = TRec!EXID
        LItemID = TRec!itemid
    End If
    RECGRID.Sort = "SRNO"
    If RECGRID.RecordCount > 0 Then
        Cnn.BeginTrans
        CNNERR = True
        LConSno = Get_ConSNo(vcDTP1.Value, DataCombo1.BoundText, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
        If Fb_Press = 2 Then
            If Len(Trim(DataCombo4.BoundText)) > 0 Then
                Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & LConSno & " AND userid = '" & DataCombo4.BoundText & "'"
            Else
                Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & LConSno & ""
            End If
        End If
        
        LDataImport = IIf(IsNull(LDataImport), 0, LDataImport)
        LPattan = Mid(Combo1.text, 1, 1)
        Dim RC As ADODB.Recordset
        RECGRID.MoveFirst
        MBAmt = 0:        MSAmt = 0
        Do While Not RECGRID.EOF
            MCL = vbNullString
            If Len(RECGRID!BNAME & "") > Val(0) Then   ''WHEN BUYER AND SELLER BOTH ARE THERE
                If RECGRID!BQnty > Val(0) And RECGRID!BRate > Val(0) Then                   ''QNTY AND RATE REQUIRED
                    If RECGRID!DIMPORT = 0 Then
                        MCL = RECGRID!BCODE
                    Else
                        MCL = RECGRID!LCLCODE
                    End If
                    If InStr(LParties, "'" & RECGRID!BCODE & "'") < 1 Then
                        If LenB(LParties) <> 0 Then LParties = LParties & ", "
                        LParties = LParties & "'" & RECGRID!BCODE & "'"
                    End If
                    If RECGRID!CONTYPE = "B" Then
                        LConType = "B"
                        MBAmt = MBAmt + (Val(RECGRID!BQnty & "") * Val(RECGRID!BRate & "")) * FCalval
                    Else
                        LConType = "S"
                        MSAmt = MSAmt + (Val(RECGRID!BQnty & "") * Val(RECGRID!BRate & "")) * FCalval
                    End If
                    LDataImport = Abs(RECGRID!DIMPORT)
                    LACCID = Get_AccID(RECGRID!BCODE)
                    mysql = "INSERT INTO CTR_D (COMPCODE ,CLCODE,CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT,DATAIMPORT,CONTIME,USERID,ORDNO,ROWNO1,EXCODE,PATTAN,CONCODE,CALVAL,CONFIRM,INSTTYPE,OPTTYPE,STRIKE,SAUDAID,EXID,ITEMID,ACCID ) "
                    mysql = mysql & "VALUES(" & GCompCode & ",'" & MCL & "'," & Val(LConSno) & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'," & Val(RECGRID!SrNo) & ",'" & TxtSaudaCode.text & "', '" & TxtItem.text & "', '" & RECGRID!BCODE & "', " & Val(RECGRID!BQnty) & "," & Val(RECGRID!BRate) & ",'" & LConType & "', 'N'," & LDataImport & ",'" & RECGRID!contime & "','" & (RECGRID!USERID & "") & "','','" & Val(RECGRID!SrNo) & "','" & LExCode & "','" & LPattan & "',''," & FCalval & ",0,'" & LInstType & "','" & LOptType & "'," & LStrike & "," & LSaudaID & "," & LExID & "," & LItemID & "," & LACCID & " )"
                    Cnn.Execute mysql
                End If
            End If
            RECGRID.MoveNext
        Loop
        'LSaudaId = Get_SaudaID(TxtSaudaCode.text)
        
        Call Shree_Posting(vcDTP1.Value)
        'Call Update_BrokTran(vbNullString, Str(LExID), Trim(Str(LItemID)), Str(LSaudaID), vcDTP1.Value, vcDTP1.Value)
        Call Update_Charges(LParties, vbNullString, str(LSaudaID), vbNullString, vcDTP1.Value, vcDTP1.Value, True)
        Cnn.CommitTrans
        CNNERR = False
        Cnn.BeginTrans
        CNNERR = True
        If BILL_GENERATION(CDate(vcDTP1.Value), CDate(GFinEnd), str(LSaudaID), LParties, vbNullString) Then
            Cnn.CommitTrans: CNNERR = False
        Else
            Cnn.RollbackTrans: CNNERR = False
        End If
        'Call Chk_Billing
    End If
    Call CANCEL_REC
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End Sub
Sub CANCEL_REC()
    vcDTP1.Enabled = True: TxtSaudaCode.Enabled = True: DataCombo1.Enabled = True: Combo1.Enabled = True: DataCombo4.Enabled = True: Text3.Enabled = True
    LConNo = 1
    Call RecSet
    Fb_Press = 0
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.Refresh
    Label2.Visible = False
    DataCombo3.Visible = False
    Call Get_Selection(10)
    Combo1.ListIndex = -1: Frame1.Enabled = False
End Sub
Function MODIFY_REC(LCondate As Date, LSauda As String, LPattan As String) As Boolean
    Dim MFromDate As Date:          Dim MToDate As Date
    Dim TRec As ADODB.Recordset:    Dim LConSno  As Long
    Dim LSaudaID As Long
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE  ='" & TxtItem.text & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then FCalval = TRec!lot
    
    
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT CONSNO FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LCondate, "yyyy/MM/dd") & "' AND SAUDA='" & LSauda & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        If Fb_Press = 2 Then
            MsgBox "Transaction does not exists for the Selected creteria?", vbExclamation
            CTRBUYSELL.Fb_Press = 1
            MODIFY_REC = True
            Exit Function
        ElseIf Fb_Press = 1 Then
            MODIFY_REC = True
        End If
        Exit Function
    Else
        LConSno = TRec!CONSNO
        If Fb_Press = 1 Then
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            mysql = "SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LCondate, "yyyy/MM/dd") & "' AND SAUDA = '" & LSauda & "'"
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                MsgBox "Contract already exists.Please press enter to modify contract.", vbInformation
                OldDate = vcDTP1.Value
                CTRBUYSELL.Fb_Press = 2
                vcDTP1.Value = OldDate
                GETMAIN.StatusBar1.Panels(2).text = "Modify Record"
                MODIFY_REC = False
                Exit Function
            Else
                MODIFY_REC = True
                Exit Function
            End If
         Else
            MODIFY_REC = True
            LConSno = TRec!CONSNO
        End If
        LDataImport = 0
    End If
    If Fb_Press = 1 Then Fb_Press = 2
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT C.CONNO,C.PARTY,C.SAUDA,C.CONTYPE,C.QTY,C.RATE,C.INVNO,C.DATAIMPORT,C.CONTIME, C.CLCODE,C.USERID,A.NAME AS NAME FROM CTR_D AS C , ACCOUNTD AS A "
    mysql = mysql & " WHERE C.COMPCODE =" & GCompCode & " AND C.COMPCODE =A.COMPCODE AND C.PARTY=A.AC_CODE AND C.CONSNO=" & LConSno & " ORDER BY C.CONNO,C.CONTYPE"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    Call RecSet
    LParties = vbNullString
    RECGRID.Delete
    LParty = vbNullString
    Do While Not TRec.EOF
        LConNo = TRec!CONNO
        If InStr(LParties, "'" & TRec!PARTY & "'") < 1 Then
            If LenB(LParties) <> 0 Then LParties = LParties & ", "
            LParties = LParties & "'" & TRec!PARTY & "'"
        End If
        RECGRID.AddNew
        RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
        RECGRID!CONTYPE = TRec!CONTYPE
        RECGRID!BCODE = TRec!PARTY
        RECGRID!LCLCODE = IIf(IsNull(TRec!CLCODE), "", TRec!CLCODE)
        RECGRID!BNAME = TRec!NAME
        RECGRID!BQnty = TRec!QTY
        RECGRID!BRate = TRec!Rate
        RECGRID!LInvNo = Val(TRec!invno & "")
        RECGRID!DIMPORT = IIf(IsNull(TRec!DATAIMPORT), 0, TRec!DATAIMPORT)
        RECGRID!contime = IIf(IsNull(TRec!contime), Time, TRec!contime)
        RECGRID!USERID = TRec!USERID & vbNullString
        RECGRID.Update
        TRec.MoveNext
    Loop
    
    Set DataGrid1.DataSource = RECGRID
    Call DataGrid1_AfterColEdit(0)

    If Fb_Press = 3 Then
        
        If MsgBox("You are about to Delete all Contracts. Confirm Delete?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            On Error GoTo err1
            Cnn.BeginTrans
            CNNERR = True
            mysql = "DELETE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & LConSno & ""
            Cnn.Execute mysql
            
            MFromDate = Format(vcDTP1.Value, "yyyy/MM/dd")
            ''TO FIND TODATE
            LSaudaID = Get_SaudaID(TxtSaudaCode.text)
            Call Update_Charges(vbNullString, vbNullString, str(LSaudaID), vbNullString, vcDTP1.Value, vcDTP1.Value, True)
            If BILL_GENERATION(vcDTP1.Value, GFinEnd, str(LSaudaID), vbNullString, vbNullString) Then
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
            If err.Number <> 0 Then MsgBox err.Description, vbCritical, "Error Number : " & err.Number
            If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
            Call CANCEL_REC
        End If
    End If
End Function


Private Sub Check1_Click()
Call UpdateSaudaCombo
End Sub

Private Sub Check1_Validate(Cancel As Boolean)
'Call UpdateSaudaCombo
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    Dim TRec As ADODB.Recordset
    If Combo1.ListIndex = 1 Then
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open "SELECT TOP 1 CONDATE FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND PATTAN='O' AND SAUDA='" & TxtSaudaCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            If DateValue(vcDTP1.Value) <> DateValue(TRec!Condate) Then
                MsgBox "Opening for this Sauda has been already entered on " & Format(TRec!Condate, "yyyy/MM/dd"), vbExclamation, "Warning"
                vcDTP1.Value = Date
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    LConSno = 0: Set TRec = Nothing: Set TRec = New ADODB.Recordset
    mysql = "SELECT CONSNO FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' "
    mysql = mysql & " AND SAUDA='" & TxtSaudaCode.text & "' AND PATTAN='" & Left$(Combo1.text, 1) & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then LConSno = TRec!CONSNO
    
        If MODIFY_REC(vcDTP1.Value, TxtSaudaCode.text, Combo1.text) Then
        Else
            Cancel = True
        End If
    
End Sub
Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
    'If Len(Trim(Text2.text)) > 0 Then Combo1.SetFocus
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    Dim LSSauda As String:    Dim LSItemCode As String:    Dim TRec As ADODB.Recordset
    Dim LExID As Integer
    If LenB(DataCombo1.BoundText) > 1 Then
        LSSauda = DataCombo1.BoundText
        If Check1.Value = 1 Then
            Rec_Sauda.MoveFirst
            Rec_Sauda.Find "SAUDACODE='" & LSSauda & "'"
            If Not Rec_Sauda.EOF Then
                LExID = Get_ExID(Rec_Sauda!excode)
                LSItemCode = Get_ItemMaster(LExID, Rec_Sauda!ITEMCODE)
                If LenB(LSItemCode) < 1 Then
                    mysql = "SELECT * FROM CONTRACTMASTER WHERE ITEMCODE='" & Rec_Sauda!ITEMCODE & "'"
                    Set TRec = Nothing
                    Set TRec = New ADODB.Recordset
                    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                    If Not TRec.EOF Then
                        LSItemCode = Create_TItemMast(TRec!ITEMCODE, TRec!ITEMName, TRec!EX_SYMBOL, TRec!lot, TRec!excode)
                    End If
                End If
                LSSauda = Create_TSaudaMast(Rec_Sauda!ITEMCODE, Rec_Sauda!MATURITY, Rec_Sauda!excode, Rec_Sauda!INSTTYPE, Rec_Sauda!OPTTYPE, Rec_Sauda!STRIKEPRICE)
                If LenB(LSSauda) > 1 Then
                    TxtSaudaCode.text = LSSauda
                    TxtExCode.text = Rec_Sauda!excode
                    TxtInst.text = Rec_Sauda!INSTTYPE
                    TxtOptType.text = Rec_Sauda!OPTTYPE
                    TxtStrike.text = Rec_Sauda!STRIKEPRICE
                    TxtItem.text = Rec_Sauda!ITEMCODE
                    FCalval = Rec_Sauda!lot
                Else
                    MsgBox " Invalid Sauda Selected "
                    Cancel = True
                    DataCombo1.SetFocus
                End If
            Else
                MsgBox " Invalid Sauda Selected "
                Cancel = True
                DataCombo1.BoundText = vbNullString
                DataCombo1.SetFocus
            End If
        Else
            Rec_Sauda.MoveFirst
            Rec_Sauda.Find "SAUDACODE='" & LSSauda & "'"
            If Not Rec_Sauda.EOF Then
                TxtSaudaCode.text = LSSauda
                TxtExCode.text = Rec_Sauda!excode
                TxtInst.text = Rec_Sauda!INSTTYPE
                TxtOptType.text = Rec_Sauda!OPTTYPE
                TxtStrike.text = Rec_Sauda!STRIKEPRICE
                TxtItem.text = Rec_Sauda!ITEMCODE
            Else
                MsgBox " Invalid Sauda Selected "
                Cancel = True
                DataCombo1.BoundText = vbNullString
                DataCombo1.SetFocus
            End If
        End If
    Else
        MsgBox " Invalid Sauda Selected "
        Cancel = True
        DataCombo1.BoundText = vbNullString
        DataCombo1.SetFocus
    End If
    mysql = "SELECT LOT,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE='" & Rec_Sauda!ITEMCODE & "'"
    Set RecEx = Nothing: Set RecEx = New ADODB.Recordset: RecEx.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RecEx.EOF Then
        FCalval = RecEx!lot
        LExCode = RecEx!EXCHANGECODE
    Else
        FCalval = 0
    End If
    
    Combo1.SetFocus
    Call GetCloseRate
End Sub
Private Sub DataCombo2_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_GotFocus()
    Sendkeys "%{DOWN}"
    If DataGrid1.Col = 1 Or DataGrid1.Col = 2 Then
        DataGrid1.text = ""
        DataGrid1.Col = 2
        Label2.Visible = True: Label2.Left = 2080
        DataCombo3.Left = Val(2080)
        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    ElseIf DataGrid1.Col = 5 Or DataGrid1.Col = 6 Or DataGrid1.Col = 7 Then
        DataGrid1.text = ""
        DataGrid1.Col = 5: DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
        DataCombo3.Left = Val(7200)
        Label2.Visible = True: Label2.Left = 7200
    End If
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If DataGrid1.Col = 2 Or DataGrid1.Col = 3 Then
            RECGRID!BCODE = DataCombo3.BoundText
            RECGRID!BNAME = DataCombo3.text
            RECGRID!USERID = LUserId
            DataGrid1.Col = 3
        End If
        DataGrid1.SetFocus
        DataCombo3.Visible = False: Label2.Visible = False
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
    If MODIFY_REC(vcDTP1.Value, DataCombo1.BoundText, Combo1.text) Then
        If Len(Trim(DataCombo4.BoundText)) > 0 Then LUserId = DataCombo4.BoundText
    Else
        Cancel = True
    End If
End Sub
Public Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)
    If ColIndex = Val(2) Then
        Rec_Account.MoveFirst
        Rec_Account.Find "AC_CODE='" & DataGrid1.text & "'", , adSearchForward
        If Not Rec_Account.EOF Then
            If ColIndex = Val(2) Then
                DataGrid1.Col = 3
                RECGRID!BCODE = Rec_Account!AC_CODE
                RECGRID!BNAME = Rec_Account!NAME
            Else
                If Combo1.ListIndex = Val(0) Then
                    DataGrid1.Col = 7
                Else
                    DataGrid1.Col = 6
                End If
                RECGRID!scode = Rec_Account!AC_CODE
                RECGRID!SNAME = Rec_Account!NAME
            End If
        Else
            DataCombo3.Visible = True
            DataCombo3.SetFocus
            'SendKeys "%{DOWN}"
        End If
    ElseIf ColIndex = 3 Or ColIndex = 4 Then
        ''IF CONTRACT THEN ONLY CHANGE OCCURS
        If Val(RECGRID!BRate & "") > 0 Then
        Else
            If ColIndex = 4 Then
            Else
                MsgBox "Rate can not be zero. Please enter Rate.", vbCritical
                DataGrid1.Col = 4: DataGrid1.SetFocus
            End If
        End If
    ElseIf ColIndex = 1 Then
        If DataGrid1.text = "B" Or DataGrid1.text = "S" Or DataGrid1.text = "b" Or DataGrid1.text = "s" Then
           DataGrid1.text = UCase(DataGrid1.text)
            LConType = UCase(DataGrid1.text)
        Else
            DataGrid1.Col = 1
            DataGrid1.text = "B"
            LConType = UCase(DataGrid1.text)
        End If
    End If
End Sub

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If Fb_Press = 2 Then
    End If
End Sub
Private Sub DataGrid1_GotFocus()
    vcDTP1.Enabled = False
    TxtSaudaCode.Enabled = False
    DataCombo1.Enabled = False
    Combo1.Enabled = False
    DataCombo4.Enabled = False
    Text3.Enabled = False
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRec As ADODB.Recordset:    Dim LBQnty  As Double:    Dim LSQnty As Double:
    Dim LBAmt  As Double:           Dim LSAmt As Double:      Dim LBDiffAmt  As Double
    Dim LBuyCode As String:         Dim LBuyName As String:   Dim LConType As String
    
    Set TRec = Nothing: Set TRec = New ADODB.Recordset: Set TRec = RECGRID.Clone
    LBQnty = 0: LSQnty = 0: LBAmt = 0: LSAmt = 0
    Do While Not TRec.EOF
        If TRec!CONTYPE = "B" Then
            LBQnty = LBQnty + Val(TRec!BQnty & "")
            LBAmt = LBAmt + (Val(TRec!BQnty & "") * Val(TRec!BRate & "")) * FCalval
        Else
            LSQnty = LSQnty + Val(TRec!BQnty & "")
            LSAmt = LSAmt + (Val(TRec!BQnty & "") * Val(TRec!BRate & "")) * FCalval
        End If
        TRec.MoveNext
    Loop
    Text1.text = LBQnty: Text4.text = LSQnty
    If LBQnty <> 0 Then
        Text7.text = LBAmt / (LBQnty * FCalval)
    Else
        Text7.text = 0
    End If
    If LSQnty <> 0 Then
        Text8.text = LSAmt / (LSQnty * FCalval)
    Else
        Text8.text = 0
    End If
    Text7.text = Format(Text7.text, "0.00")
    Text8.text = Format(Text8.text, "0.00")
    Text9.text = Format(LBAmt, "0.00") ' Bought Amount
    Text5.text = Format(Val(Text1.text) - Val(Text4.text), "0.00")
    Text6.text = Format(Val(LBAmt) - Val(LSAmt), "0.00")
    LBDiffAmt = LSAmt - LBAmt
    If Val(Text5.text) <> 0 Then
        Text11.text = Format(Val(LBDiffAmt) / Val(Text5.text), "0.00")
    End If
    Text10.text = Format(LSAmt, "0.00")
    If KeyCode = 13 And DataGrid1.Col = 5 Then
        LBuyCode = RECGRID!BCODE
        LBuyName = RECGRID!BNAME
        LConType = RECGRID!CONTYPE
        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            If Combo1.ListIndex = Val(1) Then   ''OPENING
                RECGRID!BRate = Val(Text3.text)
            Else                        ''LAST INFORMATION CARIES
                RECGRID!BCODE = LBuyCode
                RECGRID!BNAME = LBuyName
            End If
            RECGRID!CONTYPE = LConType
            RECGRID!DIMPORT = 0
            RECGRID!USERID = LUserId & ""
            RECGRID!contime = Time
            LConNo = LConNo + 1
            RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
            RECGRID.Update
        End If
        
        DataGrid1.LeftCol = 0
        DataGrid1.Col = 0
    ElseIf KeyCode = 114 Then   'F3  NEW PARTY
        GETACNT.Show
        GETACNT.ZOrder
        GETACNT.add_record
    ElseIf KeyCode = 46 And Shift = 2 Then
        RECGRID.Delete
        If RECGRID.RecordCount = 0 Then
            RECGRID.AddNew
            LConNo = LConNo + 1
            RECGRID!SrNo = LConNo 'RECGRID.RecordCount
            If Combo1.ListIndex = Val(1) Then
                RECGRID!BRate = Val(Text3.text)
                RECGRID!SRate = Val(Text3.text)
                RECGRID!USERID = LUserId
            End If
            RECGRID.Update
        End If
        Call DataGrid1_AfterColEdit(0)
    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 2 Or DataGrid1.Col = 5) Then
        If Len(Trim(DataGrid1.text)) < 1 Then
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        Else
        End If
    ElseIf KeyCode = 27 Then
        KeyCode = 0
    ElseIf KeyCode = 13 And DataGrid1.Col = 1 Then
        If DataGrid1.text = "B" Or DataGrid1.text = "S" Or DataGrid1.text = "b" Or DataGrid1.text = "s" Then
           DataGrid1.text = UCase(DataGrid1.text)
            LConType = UCase(DataGrid1.text)
        Else
            DataGrid1.Col = 1
            DataGrid1.text = "B"
            LConType = UCase(DataGrid1.text)
        End If
        DataGrid1.SetFocus
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
    vcDTP1.Value = Date
    Set DataGrid1.DataSource = RECGRID: DataGrid1.Refresh
    Frame1.Enabled = False
'--------
    LDataImport = 0
    Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
    Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND GCODE IN (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Account.EOF Then Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    If Fb_Press > 0 Then Call Get_Selection(Fb_Press)
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then FrmSauda.Show
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Sub RecSet()
    
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "BCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "BNAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "BQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "SNAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "SQNTY", adDouble, , adFldIsNullable
    
    
    RECGRID.Fields.Append "SRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "RInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "DImport", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "CONTIME", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "UserId", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "LCLCODE", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "RCLCODE", adVarChar, 30, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    RECGRID.AddNew
    RECGRID!DIMPORT = 0
    RECGRID!contime = Time
    RECGRID!USERID = LUserId
    RECGRID.Update
    
    LConNo = LConNo
    RECGRID!SrNo = LConNo  'RECGRID.AbsolutePosition
    RECGRID!CONTYPE = "B"
    DataGrid1.Col = 1
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Text3.text = Format(Text3.text, "0.00")
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
    Case 13
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
    Call UpdateSaudaCombo
End Sub
Function GetCloseRate() As Boolean
    Dim LSaudaID As Long
    LSaudaID = Val(TxtSaudaCode.text)
    Text3.text = Format(SDCLRATE(LSaudaID, vcDTP1.Value, "C"), "0.00")
End Function
Private Sub UpdateSaudaCombo()
Set Rec_Sauda = Nothing
Set Rec_Sauda = New ADODB.Recordset
If Check1.Value = 1 Then
    mysql = "SELECT SAUDACODE,ITEMCODE,MATURITY,INSTTYPE,OPTTYPE,STRIKEPRICE, EXCODE,LOT FROM SCRIPTMASTER"
    mysql = mysql & " WHERE MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY SAUDACODE"
Else
    mysql = "SELECT SAUDACODE,ITEMCODE,MATURITY,INSTTYPE,OPTTYPE,STRIKEPRICE,EXCODE FROM SAUDAMAST "
    mysql = mysql & " WHERE MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY SAUDACODE"
End If
Rec_Sauda.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not Rec_Sauda.EOF Then
    Set DataCombo1.RowSource = Rec_Sauda
    DataCombo1.BoundColumn = "SAUDACODE"
    DataCombo1.ListField = "SAUDACODE"
End If
End Sub
