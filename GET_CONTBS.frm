VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form GET_CONTBS 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Contract Entry"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   17820
   Begin VB.TextBox TxtItemId 
      Height          =   375
      Left            =   12240
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "Text18"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtSaudaId 
      Height          =   375
      Left            =   10680
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "Text18"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtExID 
      Height          =   375
      Left            =   8280
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "Text18"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtExCode 
      Height          =   495
      Left            =   5160
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Text17"
      Top             =   10080
      Visible         =   0   'False
      Width           =   2655
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
      TabIndex        =   20
      Top             =   0
      Width           =   17295
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   17295
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
            ForeColor       =   &H00800080&
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Width           =   17055
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   16935
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
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
         Height          =   8325
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   16500
         Begin VB.Frame Frame7 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   6975
            Left            =   9480
            TabIndex        =   36
            Top             =   240
            Width           =   6855
            Begin VB.TextBox Text16 
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
               Left            =   3480
               TabIndex        =   42
               Text            =   "Text14"
               Top             =   6480
               Width           =   1335
            End
            Begin VB.TextBox Text15 
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
               Left            =   2040
               TabIndex        =   41
               Text            =   "Text13"
               Top             =   6480
               Width           =   1335
            End
            Begin VB.TextBox Text14 
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
               Left            =   3480
               TabIndex        =   40
               Text            =   "Text14"
               Top             =   6000
               Width           =   1335
            End
            Begin VB.TextBox Text13 
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
               Left            =   2040
               TabIndex        =   39
               Text            =   "Text13"
               Top             =   6000
               Width           =   1335
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame8"
               Height          =   375
               Left            =   120
               TabIndex        =   37
               Top             =   120
               Width           =   6495
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Partywise Standing && Average Rates"
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
                  Left            =   960
                  TabIndex        =   38
                  Top             =   0
                  Width           =   4575
               End
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   5295
               Left            =   120
               TabIndex        =   45
               ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select members, F5 to select non members."
               Top             =   600
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   9340
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   8388736
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Parties"
                  Object.Width           =   2381
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Code"
                  Object.Width           =   1323
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Open Qty"
                  Object.Width           =   1323
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Buy Qty"
                  Object.Width           =   1323
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Sell Qty"
                  Object.Width           =   1323
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Close Qty"
                  Object.Width           =   1323
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "FmlyId"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "SRVTAXAPP"
                  Object.Width           =   0
               EndProperty
            End
            Begin VB.Label Label10 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Sell Qty and Rate"
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
               TabIndex        =   44
               Top             =   6480
               Width           =   1815
            End
            Begin VB.Label Label9 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Buy Qty and Rate"
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
               TabIndex        =   43
               Top             =   6000
               Width           =   1815
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   855
            Left            =   120
            TabIndex        =   21
            Top             =   7320
            Width           =   16215
            Begin VB.TextBox Text12 
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
               Height          =   360
               Left            =   1200
               TabIndex        =   35
               Text            =   "Text12"
               Top             =   240
               Width           =   2655
            End
            Begin VB.TextBox Text11 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   13365
               Locked          =   -1  'True
               TabIndex        =   30
               TabStop         =   0   'False
               Text            =   "Text11"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   4440
               Locked          =   -1  'True
               TabIndex        =   29
               TabStop         =   0   'False
               Text            =   "Text1"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   8385
               Locked          =   -1  'True
               TabIndex        =   28
               TabStop         =   0   'False
               Text            =   "Text4"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   12330
               Locked          =   -1  'True
               TabIndex        =   27
               TabStop         =   0   'False
               Text            =   "Text5"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   14400
               Locked          =   -1  'True
               TabIndex        =   26
               TabStop         =   0   'False
               Text            =   "Text6"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   5475
               Locked          =   -1  'True
               TabIndex        =   25
               TabStop         =   0   'False
               Text            =   "Text7"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   9540
               Locked          =   -1  'True
               TabIndex        =   24
               TabStop         =   0   'False
               Text            =   "Text8"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox Text9 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   23
               TabStop         =   0   'False
               Text            =   "Text9"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox Text10 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   10575
               Locked          =   -1  'True
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "Text10"
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contra A/c"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   34
               Top             =   278
               Width           =   1020
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000003&
               Caption         =   "Diff"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   240
               Left            =   11880
               TabIndex        =   33
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000003&
               Caption         =   "Sale"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   240
               Left            =   7920
               TabIndex        =   32
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000003&
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
               ForeColor       =   &H00800080&
               Height          =   240
               Left            =   3960
               TabIndex        =   31
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.CommandButton cmdImportFromExcel 
            Caption         =   "..."
            Height          =   285
            Left            =   11280
            TabIndex        =   9
            Top             =   7800
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox Combo1 
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
            ItemData        =   "GET_CONTBS.frx":0000
            Left            =   7680
            List            =   "GET_CONTBS.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   255
            Width           =   1575
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   735
            Width           =   1185
         End
         Begin VB.TextBox Text2 
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
            Left            =   3240
            MaxLength       =   50
            TabIndex        =   1
            Top             =   240
            Width           =   975
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
            TabIndex        =   7
            Top             =   1200
            Width           =   9135
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   360
            Left            =   4680
            TabIndex        =   4
            Top             =   720
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   420
            Left            =   4440
            TabIndex        =   10
            Top             =   2760
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5820
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   9210
            _ExtentX        =   16245
            _ExtentY        =   10266
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
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   2294.929
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1005.165
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
            Left            =   4680
            TabIndex        =   2
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
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
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   360
            Left            =   1080
            TabIndex        =   0
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37860.8625462963
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   360
            Left            =   1080
            TabIndex        =   12
            Top             =   735
            Width           =   3135
            _ExtentX        =   5530
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
            ForeColor       =   &H00800080&
            Height          =   285
            Left            =   4200
            TabIndex        =   19
            Top             =   765
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   195
            Index           =   0
            Left            =   720
            Picture         =   "GET_CONTBS.frx":0021
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
            TabIndex        =   18
            Top             =   1215
            Width           =   2115
         End
         Begin VB.Image Image1 
            Height          =   195
            Index           =   1
            Left            =   1080
            Picture         =   "GET_CONTBS.frx":032B
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
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   285
            Index           =   18
            Left            =   7200
            TabIndex        =   17
            Top             =   315
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cl Rate"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   285
            Index           =   4
            Left            =   7200
            TabIndex        =   16
            Top             =   750
            Width           =   1350
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ItemName"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   780
            Width           =   960
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
            ForeColor       =   &H00800080&
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   14
            Top             =   285
            Width           =   570
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
            ForeColor       =   &H00800080&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   315
            Width           =   435
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   18120
      Top             =   3000
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
            Picture         =   "GET_CONTBS.frx":0635
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GET_CONTBS.frx":0A87
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   10680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004080&
      BorderWidth     =   12
      Height          =   8865
      Left            =   0
      Top             =   720
      Width           =   17205
   End
End
Attribute VB_Name = "GET_CONTBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean:                Dim LParty As String:               Dim LConNo As Long:                 Dim LPattan As String
Dim LUserId As String:              Dim LContractAcc As String:         Dim LExCode  As String:             Dim LConType As String
Dim LConSno As Long:                Dim LDataImport As Byte:            Dim OldDate As Date:                Dim FLOWDIR As Byte
Dim VchNo As String:                Dim GRIDPOS As Byte:                Public Fb_Press As Byte:            Dim RecEx As ADODB.Recordset
Dim RECGRID As ADODB.Recordset:     Dim Rec_Sauda As ADODB.Recordset:   Dim Rec_Account As ADODB.Recordset: Dim REC_CloRate As ADODB.Recordset
Dim CtrMRec As ADODB.Recordset:     Dim FCalval As Double
Sub Add_Rec()
    If Rec.EOF Then
        MsgBox "Check Item Master And Sauda Mmster "
    Else
        If Rec_Account.RecordCount > 0 Then
            LDataImport = 0
            Frame1.Enabled = True: Combo1.ListIndex = 0
            Call Get_Selection(1)
            If vcDTP1.Enabled Then vcDTP1.SetFocus
        Else
            Call CANCEL_REC
        End If
    End If
End Sub
Sub Save_Rec()
    On Error GoTo err1
    'validation
    Dim MFromDate As Date
    Dim MToDate  As Date
    Dim LSExCode As String
    Dim LSaudaID As Long
    If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    LSaudaID = Get_SaudaID(Text2.text)
    If Val(Text1.text) + Val(Text4.text) = 0 Then MsgBox "Please Check Entries.", vbCritical: Exit Sub
    'If Val(Text4.Text) = 0 Then MsgBox "Please Check Entries.", vbCritical:  Exit Sub
    Set Rec_Sauda = Nothing: Set Rec_Sauda = New ADODB.Recordset
    Rec_Sauda.Open "SELECT SAUDACODE,ITEMCODE,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDAID=" & LSaudaID & "", Cnn, adOpenForwardOnly, adLockReadOnly
    If Rec_Sauda.EOF Then
        MsgBox "Invalid Sauda Code.", vbExclamation, "Error": Text2.SetFocus: Exit Sub
    Else
        Set GeneralRec1 = Nothing: Set GeneralRec1 = New ADODB.Recordset
        GeneralRec1.Open "SELECT EX.CONTRACTACC,EX.SHREEAC,EX.TRADINGACC FROM EXMAST AS EX , ITEMMAST AS IM WHERE EX.COMPCODE=" & GCompCode & " AND EX.COMPCODE=IM.COMPCODE AND EX.EXCODE=IM.EXCHANGECODE  AND  IM.ITEMCODE = '" & Rec_Sauda!ITEMCODE & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not GeneralRec1.EOF Then
             GShree = GeneralRec1!shreeac
             GTrading = GeneralRec1!TRADINGACC
             GCONTAC = GeneralRec1!CONTRACTACC
             LSExCode = GeneralRec1!EXCODE
        End If
    End If
    RECGRID.Sort = "SRNO"
    If RECGRID.RecordCount > 0 Then
        Cnn.BeginTrans
        CNNERR = True
        If Fb_Press = 1 Then
            VchNo = Get_VouNo("CONT", GFinYear)
            If Not CtrMRec.EOF Then
              Set Rec = Nothing: Set Rec = New ADODB.Recordset
               Rec.Open "SELECT CONSNO FROM CTR_M WHERE COMPCODE =" & GCompCode & "  AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND SAUDAID =" & LSaudaID & " AND PATTAN = '" & Mid(Combo1.text, 1, 1) & "' ", Cnn, adOpenForwardOnly, adLockReadOnly
                If Not Rec.EOF Then
                    CONSNO = Rec!CONSNO
              Else
                    Set Rec = Nothing: Set Rec = New ADODB.Recordset
                   Rec.Open "SELECT MAX(CONSNO) FROM CTR_M WHERE COMPCODE =" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
                   CONSNO = Val(Rec.Fields(0) & "") + Val(1)
                End If
                Set Rec = Nothing
            Else
                CONSNO = 1
            End If
        Else
            CONSNO = CtrMRec!CONSNO
        End If
        Call Delete_Voucher(VchNo)
        If Fb_Press = 2 Then
            If Len(Trim(DataCombo4.BoundText)) > 0 Then
                Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & Val(CtrMRec!CONSNO) & " AND USERID = '" & DataCombo4.BoundText & "'"
            Else
                Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & Val(CtrMRec!CONSNO) & ""
            End If
            If Text3.Locked Then
            Else
                Cnn.Execute "DELETE FROM CTR_R WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & Val(CtrMRec!CONSNO) & ""
            End If
            Cnn.Execute "DELETE FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & Val(CtrMRec!CONSNO) & ""
        End If
        Cnn.Execute "DELETE FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & CONSNO & ""
        
        LDataImport = IIf(IsNull(LDataImport), 0, LDataImport)
        'MYSQL = "EXEX INSERT_CTR_M "
        LConSno = Get_ConSNo(vcDTP1.Value, Text2.text, DataCombo2.BoundText, TxtExCode.text, Val(TxtSaudaID.text), Val(TxtItemID.text), Val(TxtExID.text))
        'MYSQL = "EXEC INSERT_CTR_M " & GCompCode & "," & ConSno & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "','" & Text2.text & "','" & DataCombo2.BoundText & "','" & Left(Combo1.text, 1) & "','" & LSExCode & "'"
        'MYSQL = "INSERT INTO CTR_M(COMPCODE,CONSNO, CONDATE, SAUDA, ITEMCODE, CLOSERATE, VOU_NO, PATTAN,DataImport) VALUES(" & GCompCode & "," & ConSno & ", '" &
        'Format(vcDTP1.Value, "yyyy/MM/dd") & "', '" & Text2.text & "', '" & DataCombo2.BoundText & "', " & Val(Text3.text) & ", '" & VchNo & "', '" & Mid(Combo1.text, 1, 1) & "'," & LDataImport & ")"
        Cnn.Execute mysql
        LPattan = Mid(Combo1.text, 1, 1)
        Dim RC As ADODB.Recordset
        'do not initialized LPARTY here
        RECGRID.MoveFirst
        MBAmt = 0:        MSAmt = 0:        MBQNTY = 0:        MSQNTY = 0
        Do While Not RECGRID.EOF
            MCL = ""
            If Len(RECGRID!BNAME & "") > Val(0) Then   ''WHEN BUYER AND SELLER BOTH ARE THERE
                If RECGRID!BQnty > Val(0) And RECGRID!BRate > Val(0) Then                   ''QNTY AND RATE REQUIRED
                    If RECGRID!DIMPORT = 0 Then
                        MCL = RECGRID!BCODE
                    Else
                        MCL = RECGRID!LCLCODE
                    End If
                    If RECGRID!CONTYPE = "B" Then
                        LConType = "B"
                        MBAmt = MBAmt + (Val(RECGRID!BQnty & "") * Val(RECGRID!BRate & "")) * FCalval
                        MBQNTY = MBQNTY + Val(RECGRID!BQnty)
                    Else
                        LConType = "S"
                        MSAmt = MSAmt + (Val(RECGRID!BQnty & "") * Val(RECGRID!BRate & "")) * FCalval
                        MSQNTY = MSQNTY + Val(RECGRID!BQnty)
                    End If
                    LDataImport = Abs(RECGRID!DIMPORT)
                    LUserId = RECGRID!USERID
                    mysql = "EXEC INSERT_CTR_DBS " & GCompCode & ",'" & MCL & "'," & Val(CONSNO) & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'," & Val(RECGRID!SrNo) & ",'" & Text2.text & "','" & DataCombo2.BoundText & "','" & RECGRID!BCODE & "'," & Val(RECGRID!BQnty) & "," & Val(RECGRID!BRate) & ",'" & LConType & "','N','" & LDataImport & "','" & RECGRID!CONTIME & "','" & LUserId & "',''," & Val(RECGRID!SrNo) & ",'" & LExCode & "','" & LPattan & "','N'," & FCalval & ", 0 ,''"
                    Cnn.Execute mysql
                End If
            End If
            LSRNO = RECGRID!SrNo
            RECGRID.MoveNext
        Loop
        LSRNO = LSRNO + 1
        LBAVGRATE = MBAmt / MBQNTY
        LSAVGRATE = MSAmt / MSQNTY
        mysql = "EXEC INSERT_CTR_DBS " & GCompCode & ",'" & MCL & "'," & Val(CONSNO) & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'," & Val(LSRNO) & ",'" & Text2.text & "','" & DataCombo2.BoundText & "','" & LContractAcc & "'," & Val(MBQNTY) & "," & Val(LBAVGRATE) & ",'S','N','" & LDataImport & "',''  ,'" & LUserId & "',''  ," & Val(LSRNO) & ",'" & LExCode & "','" & LPattan & "','Y'," & FCalval & ", 0 ,''"
        Cnn.Execute mysql
        LSRNO = LSRNO + 1
        mysql = "EXEC INSERT_CTR_DBS " & GCompCode & ",'" & MCL & "'," & Val(CONSNO) & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'," & Val(LSRNO) & ",'" & Text2.text & "','" & DataCombo2.BoundText & "','" & LContractAcc & "'," & Val(MSQNTY) & "," & Val(LSAVGRATE) & ",'B','N','" & LDataImport & "','','" & LUserId & "',''," & Val(LSRNO) & ",'" & LExCode & "','" & LPattan & "','Y'," & FCalval & ", 0 ,''"
        Cnn.Execute mysql
        LParty = vbNullString
        MFromDate = Format(vcDTP1.Value, "yyyy/MM/dd")
        
        mysql = "SELECT MATURITY FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDAID= " & LSaudaID & ""
        Set Rec = Nothing: Set Rec = New ADODB.Recordset: Rec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not Rec.EOF Then MToDate = Rec.Fields(0)
        Call Update_Charges(vbNullString, vbNullString, Str(LSaudaID), vbNullString, MFromDate, MFromDate, True)
        Cnn.CommitTrans
        CNNERR = False
        Cnn.BeginTrans
        CNNERR = False
        If BILL_GENERATION(CDate(MFromDate), CDate(MToDate), Str(LSaudaID), vbNullString, vbNullString) Then
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
    vcDTP1.Enabled = True: Text2.Enabled = True: DataCombo1.Enabled = True: Combo1.Enabled = True: DataCombo4.Enabled = True: Text3.Enabled = True
    LConNo = 10000
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
    Dim MFromDate   As Date
    Dim MToDate As Date
    Dim LSaudaID As Long
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    mysql = "SELECT IT.LOT FROM ITEMMAST AS IT,SAUDAMAST AS SD WHERE IT.COMPCODE =" & GCompCode & " AND IT.COMPCODE = SD.COMPCODE AND IT.ITEMCODE=SD.ITEMCODE AND SD.SAUDACODE='" & LSauda & "'"
    Rec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not Rec.EOF Then
        FCalval = Rec!lot
    End If
    
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    mysql = "SELECT * FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LCondate, "yyyy/MM/dd") & "' AND SAUDA='" & LSauda & "' AND PATTAN='" & Mid(LPattan, 1, 1) & "'"
    Rec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rec.EOF Then
        If Fb_Press = 2 Then
            MsgBox "Transaction does not exists for the Selected creteria?", vbExclamation
            OldDate = vcDTP1.Value
            GET_CONTBS.Fb_Press = 1
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
            mysql = "SELECT * FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LCondate, "yyyy/MM/dd") & "' AND SAUDA = '" & LSauda & "'"
            GeneralRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                MsgBox "Contract already exists.Please press enter to modify contract.", vbInformation
                OldDate = vcDTP1.Value
                GET_CONTBS.Fb_Press = 2
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
            mysql = "SELECT * FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(LCondate, "yyyy/MM/dd") & "' AND SAUDA='" & LSauda & "' "
            GeneralRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                MODIFY_REC = True
            Else
                MsgBox "Contract does not exists.Please add New Contract.", vbInformation
                OldDate = vcDTP1.Value
                Call CANCEL_REC
                GET_CONTBS.Fb_Press = 1
                vcDTP1.Value = OldDate
                GET_CONTBS.Add_Rec
                GETMAIN.StatusBar1.Panels(2).text = "Add Record"
                MODIFY_REC = False
                Exit Function
            End If
        End If
        LDataImport = IIf(IsNull(Rec!DATAIMPORT), 0, 1)
    End If
    CtrMRec.MoveFirst
    CtrMRec.Find "CONSNO=" & Val(Rec!CONSNO & "") & "", , adSearchForward
    If Fb_Press = 1 Then Fb_Press = 2
        With CtrMRec
            vcDTP1.Value = !Condate
            DataCombo1.BoundText = !Sauda
            Text2.text = !Sauda
            DataCombo2.BoundText = !ITEMCODE
            If !PATTAN = "C" Then
                Combo1.ListIndex = Val(0)
                DataGrid1.Columns(7).Locked = True
            Else
                Combo1.ListIndex = Val(1)
                DataGrid1.Columns(7).Locked = False
            End If
        End With
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    mysql = "SELECT CTR_D.COMPCODE,CTR_D.CONNO,CTR_D.CONDATE,CTR_D.CONSNO,CTR_D.CONTYPE,CTR_D.PARTY,CTR_D.QTY,CTR_D.CONNO,RATE,CTR_D.CLCODE,CTR_D.INVNO,CTR_D.DATAIMPORT,CTR_D.CONTIME,CTR_D.USERID,ctr_d.concode, A.NAME AS NAME FROM CTR_D, ACCOUNTD AS A WHERE CTR_D.COMPCODE =" & GCompCode & " AND CTR_D.COMPCODE =A.COMPCODE AND CTR_D.PARTY=A.AC_CODE AND CTR_D.CONSNO=" & Val(CtrMRec!CONSNO) & " ORDER BY CONNO,CONTYPE"
    Rec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    Call RecSet
    RECGRID.Delete
    LParty = vbNullString
    Do While Not Rec.EOF
        LConNo = Rec!CONNO
        If Rec!CONCODE <> "Y" Then
            RECGRID.AddNew
            RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
            RECGRID!CONTYPE = Rec!CONTYPE
            RECGRID!BCODE = Rec!PARTY & ""
            RECGRID!LCLCODE = Rec!CLCODE & ""
            RECGRID!BNAME = Rec!NAME
            RECGRID!BQnty = Rec!QTY
            RECGRID!BRate = Rec!Rate
            RECGRID!LInvNo = Val(Rec!invno & "")
            RECGRID!DIMPORT = IIf(IsNull(Rec!DATAIMPORT), 1, Rec!DATAIMPORT)
            RECGRID!CONTIME = IIf(IsNull(Rec!CONTIME), Time, Rec!CONTIME)
            RECGRID!USERID = Rec!USERID & ""
            RECGRID.Update
        Else
            Text12.text = Rec!NAME
        End If
        Rec.MoveNext
    Loop
    Set DataGrid1.DataSource = RECGRID
    Call DataGrid1_AfterColEdit(0)
    If Fb_Press = 3 Then
        If MsgBox("You are about to Delete all Contracts. Confirm Delete?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            On Error GoTo err1
            Cnn.BeginTrans
            CNNERR = True
            mysql = "DELETE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & CtrMRec!CONSNO & ""
            Cnn.Execute mysql
            mysql = "DELETE FROM CTR_R WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & CtrMRec!CONSNO & ""
            Cnn.Execute mysql
            Call Delete_Voucher(CtrMRec!VOU_NO & "")
            mysql = "DELETE FROM CTR_M WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & CtrMRec!CONSNO & ""
            Cnn.Execute mysql
            Cnn.CommitTrans
            'Adodc1.Refresh
            Cnn.BeginTrans
            ''REGENERATING SETTLEMENTS
            MFromDate = Format(vcDTP1.Value, "yyyy/MM/dd")
            ''TO FIND TODATE
            Str(LSaudaID) = Get_SaudaID(Text2.text)
            mysql = "SELECT MATURITY FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND SAUDAid =" & LSaudaID & ""
            Set Rec = Nothing: Set Rec = New ADODB.Recordset: Rec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not Rec.EOF Then MToDate = Rec.Fields(0)
            'Call UpdateBrokRateType("", "'" & DataCombo2.BoundText & "'", vcDTP1.Value, vcDTP1.Value, "", "")
            'If GMarginYN = "Y" Then Call UpdateMargin("'" & DataCombo2.BoundText & "'", "", vcDTP1.Value, CStr(GFinEnd), "")
            Call Update_Charges(vbNullString, vbNullString, Str(LSaudaID), vbNullString, MFromDate, MToDate, True)
            If BILL_GENERATION(CDate(MFromDate), CDate(MToDate), Str(LSaudaID), vbNullString, vbNullString) Then
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
                MsgBox err.Description, vbCritical, "Error Number : " & err.Number
            End If
            If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
            Call CANCEL_REC
        End If
    End If
End Function

Private Sub cmdImportFromExcel_Click()
    Dim exlCnn As New ADODB.Connection
    Dim conStr As String
    
    
    conStr = _
    "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=C:\cust.xls;DefaultDir=c:\;"
    exlCnn.ConnectionString = conStr
    exlCnn.Open
        
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    Rec.Open "SELECT * FROM [customers$]", exlCnn, adOpenDynamic, adLockOptimistic
    
    If (Rec.RecordCount > 0) Then
        Set DataGrid1.DataSource = Rec
    End If
    Set exlCnn = Nothing
            
End Sub

Private Sub Combo1_GotFocus()
    If FLOWDIR = 1 Then
        Set Rec = Nothing
        Set Rec = New ADODB.Recordset
        Rec.Open "SELECT * FROM CTR_M WHERE COMPCODE=" & GCompCode & " AND SAUDA='" & DataCombo1.BoundText & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Rec.EOF Then Sendkeys "%{DOWN}"
    Else
       Text2.SetFocus
    End If
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 1 Then
        flag = True
    End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If flag Then
        Text3.SetFocus
    Else
        Set Rec = Nothing
        Set Rec = New ADODB.Recordset
        Rec.Open "SELECT * FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND PATTAN='O' AND SAUDA='" & DataCombo1.BoundText & "'", Cnn, adOpenForwardOnly, adLockReadOnly
    
        If Not Rec.EOF Then
            If Format(vcDTP1.Value, "yyyy/MM/dd") < Rec!Condate Then
                MsgBox "Opening for this SAUDA has been already entered on " & Format(Rec!Condate, "yyyy/MM/dd"), vbExclamation, "Warning"
                vcDTP1.Value = Date
                Cancel = True
                Exit Sub
            End If
        Else
            If Not Rec.EOF Then
                If Rec!Condate > Format(vcDTP1.Value, "yyyy/MM/dd") Then
                    MsgBox "Opening for this Sauda has been already entered on " & Format(Rec!Condate, "yyyy/MM/dd"), vbExclamation, "Warning"
                    vcDTP1.Value = Date
                    Exit Sub
                End If
            End If
        End If
    'Check UserId*****
        LConSno = 0: Set Rec = Nothing: Set Rec = New ADODB.Recordset
        mysql = "SELECT CONSNO FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND SAUDA='" & DataCombo1.BoundText & "' AND PATTAN='" & Mid(Combo1.text, 1, 1) & "'"
        Rec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not Rec.EOF Then LConSno = Rec!CONSNO
            
        Set Rec = Nothing: Set Rec = New ADODB.Recordset
        mysql = "SELECT DISTINCT FMLY.FMLYNAME,FMLY.FMLYCODE FROM CTR_D, ACCFMLY AS FMLY WHERE CTR_D.COMPCODE =" & GCompCode & " AND CTR_D.COMPCODE =FMLY.COMPCODE AND CTR_D.USERID=FMLY.FMLYCODE AND CTR_D.CONSNO = " & Val(LConSno) & " ORDER BY FMLYNAME "
        Rec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not Rec.EOF Then
            Set DataCombo4.RowSource = Rec: DataCombo4.ListField = "Fmlyname": DataCombo4.BoundColumn = "FmlyCode"   ': DataCombo4.SetFocus
        Else
            If MODIFY_REC(vcDTP1.Value, DataCombo1.BoundText, Combo1.text) Then
            Else
                Cancel = True
            End If
        End If
    End If
    flag = False
End Sub
Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
    If Len(Trim(Text2.text)) > 0 Then Combo1.SetFocus
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    'Set Rec_Sauda = Nothing: Set Rec_Sauda = New ADODB.Recordset
    'Rec_Sauda.Open "SELECT SAUDACODE,SAUDANAME,ITEMCODE FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND SAUDACODE='" & DataCombo1.BoundText & "'", Cnn, adOpenForwardOnly, adLockReadOnly
    Rec_Sauda.Find "saudacode='" & DataCombo1.BoundText & "'"
    If Rec_Sauda.EOF Then
        MsgBox "Invalid Sauda "
        Cancel = True
    Else
        Text2.text = Rec_Sauda!saudacode
        Call GetCloseRate
        'DataCombo1.BoundText = Text2.text
        'DataCombo2.BoundText = Rec_Sauda!ITEMCODE
        mysql = "SELECT LOT,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE='" & Rec_Sauda!ITEMCODE & "'"
        Set RecEx = Nothing: Set RecEx = New ADODB.Recordset: RecEx.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not RecEx.EOF Then
            FCalval = RecEx!lot
            LExCode = RecEx!EXCHANGECODE
        Else
            FCalval = 0
        End If
        mysql = "SELECT EX.CONTRACTACC FROM EXMAST AS EX  WHERE EX.COMPCODE =" & GCompCode & " AND  EX.EXCODE = LEXCODE "
        Set RecEx = Nothing: Set RecEx = New ADODB.Recordset: RecEx.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not RecEx.EOF Then
        '    LUserId = VBNULLSTRING
            LContractAcc = RecEx!CONTRACTACC
        'Else
        '    'create new branch with head of exchange contract account
        '    MYSQL = "SELECT EX.CONTRACTACC,EX.EXNAME FROM EXMAST AS EX , ITEMMAST AS IT  WHERE EX.COMPCODE =" & GCompCode & " AND ex.COMPCODE=it.COMPCODE AND ex.EXCODE=it.ExchangeCode AND IT.ITEMCODE = '" & Rec_Sauda!ITEMCODE & "'  "
        '    Set RecEx = Nothing: Set RecEx = New ADODB.Recordset: RecEx.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
        '    If Not RecEx.EOF Then
        '        Cnn.Execute "INSERT INTO ACCFMLY (COMPCODE,FMLYCODE,FMLYNAME, FMLYHEAD) VALUES (" & GCompCode & ",'" & RecEx!ContractACC & "','" & RecEx!ContractACC & "','" & RecEx!ContractACC & "')"
        '        If IsNull(RecEx!ContractACC) Then
        '            MsgBox "Please Select or Create New Contract A/c in Exchange Setup "
        '        Else
        '            LUserId = RecEx!ContractACC
        '        End If
        '        MsgBox "Generated New Default Branch for " & RecEx!EXNAME, vbInformation
        '    End If
        End If
        Combo1.SetFocus
    End If
    
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
    'Set REC = Nothing: Set REC = New ADODB.Recordset: Set REC = RECGRID.Clone
    'BQNTY = 0: SQNTY = 0: BAmt = 0: SAmt = 0
    'Do While Not REC.EOF
    '    If REC!CONTYPE = "B" Then
    '        BQNTY = BQNTY + Val(REC!BQNTY & "")
    '        BAmt = BAmt + (Val(REC!BQNTY & "") * Val(REC!BRate & "")) * FCalVal
    '    Else
    '        SQNTY = SQNTY + Val(REC!SQNTY & "")
    '        SAmt = SAmt + (Val(REC!SQNTY & "") * Val(REC!SRate & "")) * FCalVal
    '    End If
    '    REC.MoveNext
    'Loop
    'Text1.Text = BQNTY: Text4.Text = SQNTY
    'If BQNTY <> 0 Then
    '    Text7.Text = BAmt / (BQNTY * FCalVal)
    'End If
    'If SQNTY <> 0 Then
    '    Text8.Text = SAmt / (SQNTY * FCalVal)
    'End If
    'Text9.Text = BAmt
    'Text5.Text = Val(Text1.Text) - Val(Text4.Text)
    'Text6.Text = Format(Val(BAmt) - Val(SAmt), "0.00")
    'Text10.Text = SAmt
End Sub

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If Fb_Press = 2 Then
    End If
End Sub
Private Sub DataGrid1_GotFocus()
    vcDTP1.Enabled = False
    Text2.Enabled = False
    DataCombo1.Enabled = False
    Combo1.Enabled = False
    DataCombo4.Enabled = False
    Text3.Enabled = False
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Set Rec = Nothing: Set Rec = New ADODB.Recordset: Set Rec = RECGRID.Clone
    BQnty = 0: SQnty = 0: BAMT = 0: SAmt = 0
    Do While Not Rec.EOF
        If Rec!CONTYPE = "B" Then
            BQnty = BQnty + Val(Rec!BQnty & "")
            LBAmt = LBAmt + (Val(Rec!BQnty & "") * Val(Rec!BRate & ""))
            BAMT = BAMT + (Val(Rec!BQnty & "") * Val(Rec!BRate & "")) * FCalval
        Else
            SQnty = SQnty + Val(Rec!BQnty & "")
            LSAmt = LSAmt + (Val(Rec!BQnty & "") * Val(Rec!BRate & ""))
            SAmt = SAmt + (Val(Rec!BQnty & "") * Val(Rec!BRate & "")) * FCalval
        End If
        Rec.MoveNext
    Loop
    Text1.text = BQnty: Text4.text = SQnty
    If BQnty <> 0 Then
        Text7.text = BAMT / (BQnty * FCalval)
    Else
        Text7.text = 0
    End If
    If SQnty <> 0 Then
        Text8.text = SAmt / (SQnty * FCalval)
    Else
        Text8.text = 0
    End If
    Text7.text = Format(Text7.text, "0.00")
    Text8.text = Format(Text8.text, "0.00")
    Text9.text = Format(BAMT, "0.00") ' Bought Amount
    Text5.text = Format(Val(Text1.text) - Val(Text4.text), "0.00")
    Text6.text = Format(Val(BAMT) - Val(SAmt), "0.00")
    LBDiffAmt = LSAmt - LBAmt
    If Val(Text5.text) <> 0 Then
        Text11.text = Format(Val(LBDiffAmt) / Val(Text5.text), "0.00")
    End If
    Text10.text = Format(SAmt, "0.00")
    If KeyCode = 13 And DataGrid1.Col = 5 Then
        BCODE = RECGRID!BCODE
        BNAME = RECGRID!BNAME
        LConType = RECGRID!CONTYPE
        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            If Combo1.ListIndex = Val(1) Then   ''OPENING
                RECGRID!BRate = Val(Text3.text)
            Else                        ''LAST INFORMATION CARIES
                RECGRID!BCODE = BCODE
                RECGRID!BNAME = BNAME
            End If
            RECGRID!CONTYPE = LConType
            RECGRID!DIMPORT = 0
            RECGRID!USERID = LUserId & ""
            RECGRID!CONTIME = Time
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
    ElseIf KeyCode = 118 Then   ''F7 KEY
    '    RNO = InputBox("Enter the row number.", "Sauda")
    '    If Val(RNO) > Val(0) Then
    '        RECGRID.MoveFirst
    '        RECGRID.Find "SRNO=" & Val(RNO) & "", , adSearchForward
    '        If RECGRID.EOF Then
    '            MsgBox "Record not found.", vbCritical, "Error"
    '            RECGRID.MoveFirst
    '        End If
            DataGrid1.Col = 1
            DataGrid1.SetFocus
'        End If
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
'----------
    Set DataGrid1.DataSource = RECGRID: DataGrid1.Refresh
    
    Frame1.Enabled = False
'--------
    vcDTP1.Value = Date
    LDataImport = 0
    DataCombo3.Top = Val(2055): DataCombo3.Left = Val(1080)
    mysql = "SELECT ITEMCODE, ITEMCODE+','+ITEMNAME AS ITEMNAME,LOT FROM ITEMMAST WHERE COMPCODE=" & GCompCode & " ORDER BY ITEMCODE"
    Set Rec = Nothing: Set Rec = New ADODB.Recordset: Rec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec.EOF Then
        Set DataCombo2.RowSource = Rec: DataCombo2.BoundColumn = "ITEMCODE": DataCombo2.ListField = "ITEMNAME"
        QACC_CHANGE = False: Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
        Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND GCODE IN (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
        If Not Rec_Account.EOF Then Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
        Set CtrMRec = Nothing: Set CtrMRec = New ADODB.Recordset
        mysql = "SELECT * FROM CTR_M WHERE COMPCODE=" & GCompCode & " ORDER BY CONSNO"
        CtrMRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    Else
        Call Get_Selection(12)
    End If
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    If QACC_CHANGE Then
        QACC_CHANGE = False: Set Rec_Account = Nothing
        Set Rec_Account = New ADODB.Recordset
        Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND GCODE IN  (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
        If Not Rec_Account.EOF Then
            Set DataCombo3.RowSource = Rec_Account
            DataCombo3.BoundColumn = "AC_CODE"
            DataCombo3.ListField = "NAME"
        Else
            MsgBox "Please Create Customer Account", vbInformation
            Call Get_Selection(12)
        End If
    End If
    If Fb_Press > 0 Then Call Get_Selection(Fb_Press)
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then FrmSauda.Show
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    FLOWDIR = 1
    If Len(Trim(Text2.text)) < 1 Then
        DataCombo1.SetFocus
    Else
        If Not GetCloseRate Then Text2.text = vbNullString: DataCombo1.SetFocus
    End If
End Sub
Private Sub Text3_GotFocus()
    FLOWDIR = 0: Text3.SelLength = Len(Text3.text)
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
    RECGRID!CONTIME = Time
    RECGRID!USERID = LUserId
    RECGRID.Update
    
    LConNo = LConNo + 1
    RECGRID!SrNo = LConNo  'RECGRID.AbsolutePosition
    RECGRID!CONTYPE = "B"
    DataGrid1.Col = 1
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Text3.text = Format(Text3.text, "0.00")
End Sub
Sub Delete_Voucher(LPVou_No As String)
    Cnn.Execute "DELETE FROM VOUCHER WHERE COMPCODE=" & GCompCode & " AND VOU_NO='" & LPVou_No & "'"
    Cnn.Execute "DELETE FROM VCHAMT  WHERE COMPCODE=" & GCompCode & " AND VOU_NO='" & LPVou_No & "'"
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
    Set Rec_Sauda = Nothing
    Set Rec_Sauda = New ADODB.Recordset
    Rec_Sauda.Open "SELECT SAUDAID,EXID,ITEMID,EXCODE,SAUDACODE,SAUDANAME, ITEMCODE,INSTTYPE  FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND MATURITY >=  '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'ORDER BY ITEMCODE,MATURITY", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec_Sauda.EOF Then
        Set DataCombo1.RowSource = Rec_Sauda
        DataCombo1.BoundColumn = "SAUDACODE"
        DataCombo1.ListField = "SAUDANAME"
    End If
End Sub
Function GetCloseRate() As Boolean
     'Set Rec_Sauda = Nothing: Set Rec_Sauda = New ADODB.Recordset
     'MYSQL = "SELECT SAUDAID,EXID,ITEMID,EXCODE,SAUDACODE,SAUDANAME, ITEMCODE,INSTTYPE FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND SAUDACODE='" & Text2.text & "'"
     'Rec_Sauda.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
     Rec_Sauda.Find "saudacode='" & Text2.text & "'"
     If Rec_Sauda.EOF Then
         MsgBox "Invalid Sauda Code.", vbExclamation, "Error"
         GetCloseRate = False
     Else
         GetCloseRate = True
         Set REC_CloRate = Nothing: Set REC_CloRate = New ADODB.Recordset
         REC_CloRate.Open "SELECT CloseRate,DataImport FROM CTR_R WHERE COMPCODE=" & GCompCode & " AND SAUDAID =" & Rec_Sauda!SAUDAID & " AND CONDATE ='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", Cnn, adOpenForwardOnly, adLockReadOnly
         If Not REC_CloRate.EOF Then
            Text3.text = Format(REC_CloRate!CLOSERATE, "0.00")
        End If
         Text2.text = Rec_Sauda!saudacode
         DataCombo1.BoundText = CStr(Text2.text)
         DataCombo2.BoundText = Rec_Sauda!ITEMCODE
         TxtSaudaID.text = Rec_Sauda!SAUDAID
         TxtExID.text = Rec_Sauda!EXID
         TxtItemID.text = Rec_Sauda!itemid
         TxtExCode.text = Rec_Sauda!EXCODE
         LExCode = Rec_Sauda!EXCODE
    End If
End Function
