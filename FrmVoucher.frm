VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form FrmVoucher 
   Caption         =   "Voucher Entry"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18285
   BeginProperty Font 
      Name            =   "Times New Roman"
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
   ScaleHeight     =   9720
   ScaleWidth      =   18285
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtVouNo 
      Height          =   495
      Left            =   18720
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      TabIndex        =   29
      Top             =   2040
      Width           =   18615
      Begin VB.Frame VouFrame 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   6255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   18615
         Begin VB.Frame Frame7 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   735
            Left            =   120
            TabIndex        =   37
            Top             =   5400
            Width           =   18375
            Begin VB.TextBox TxtDebit 
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
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   120
               Width           =   1700
            End
            Begin VB.TextBox TxtCredit 
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
               Left            =   8700
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   120
               Width           =   1700
            End
            Begin VB.TextBox TxtDiff 
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
               Left            =   16200
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   120
               Width           =   1700
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Debit"
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
               TabIndex        =   43
               Top             =   158
               Width           =   1095
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Credit"
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
               Left            =   7440
               TabIndex        =   42
               Top             =   158
               Width           =   1095
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Differnce"
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
               Left            =   15120
               TabIndex        =   41
               Top             =   158
               Width           =   855
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   4935
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   18375
            _ExtentX        =   32411
            _ExtentY        =   8705
            _Version        =   393216
            BackColor       =   -2147483624
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
      End
      Begin VB.Frame PartyFrame 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   5655
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   18375
         Begin VB.Frame Frame10 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   735
            Left            =   120
            TabIndex        =   44
            Top             =   4560
            Width           =   18015
            Begin VB.TextBox TxtPBal 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   16200
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   165
               Width           =   1700
            End
            Begin VB.TextBox TxtPCredit 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   8700
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   165
               Width           =   1700
            End
            Begin VB.TextBox TxtPDebit 
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
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   165
               Width           =   1700
            End
            Begin VB.Label label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Account Balance"
               Height          =   255
               Left            =   14400
               TabIndex        =   50
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Credit"
               Height          =   255
               Left            =   7440
               TabIndex        =   49
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Debit"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   1095
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4335
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   7646
            _Version        =   393216
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
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18615
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   615
         Left            =   15240
         TabIndex        =   55
         Top             =   0
         Width           =   3375
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
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   960
         End
         Begin VB.CommandButton CmdMod 
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
            Height          =   375
            Left            =   1200
            TabIndex        =   5
            Top             =   120
            Width           =   960
         End
         Begin VB.CommandButton CmdCancel 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   120
            Width           =   960
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   500
         Left            =   0
         TabIndex        =   22
         Top             =   1440
         Width           =   18615
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Enteries From Date"
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
            Left            =   12480
            TabIndex        =   57
            Top             =   0
            Width           =   2175
         End
         Begin VB.TextBox TxtBalDrCr 
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
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   120
            Width           =   750
         End
         Begin VB.TextBox TxtBalance 
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
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   120
            Width           =   2000
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   495
            Left            =   14760
            TabIndex        =   23
            Top             =   0
            Width           =   3615
            Begin VB.OptionButton OptDate 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Date  Wise"
               Height          =   255
               Left            =   1920
               TabIndex        =   25
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton OptParty 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Party Wise"
               Height          =   255
               Left            =   240
               TabIndex        =   24
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFC0&
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
            Left            =   8760
            TabIndex        =   53
            Top             =   45
            Width           =   3255
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   51
            Top             =   0
            Width           =   4215
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Left            =   4560
            TabIndex        =   26
            Top             =   158
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   615
         Left            =   0
         TabIndex        =   18
         Top             =   600
         Width           =   18615
         Begin VB.TextBox TxtVchNo 
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
            Left            =   3000
            TabIndex        =   8
            Top             =   120
            Width           =   735
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
            Height          =   495
            Left            =   16920
            TabIndex        =   14
            Top             =   120
            Width           =   1320
         End
         Begin VB.TextBox TxtNarr 
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
            Left            =   12720
            TabIndex        =   13
            Top             =   120
            Width           =   4095
         End
         Begin VB.TextBox TxtAmt 
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
            Left            =   10560
            TabIndex        =   12
            Top             =   120
            Width           =   1365
         End
         Begin VB.TextBox TxtDrCr 
            Alignment       =   2  'Center
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
            Left            =   9240
            MaxLength       =   1
            TabIndex        =   11
            Text            =   "D"
            Top             =   120
            Width           =   500
         End
         Begin MSDataListLib.DataCombo DComboAcc 
            Height          =   420
            Left            =   5760
            TabIndex        =   10
            Top             =   120
            Width           =   3375
            _ExtentX        =   5953
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
         Begin VB.TextBox TxtAcCode 
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
            Left            =   4680
            TabIndex        =   9
            Top             =   120
            Width           =   975
         End
         Begin vcDateTimePicker.vcDTP vcDTP2 
            Height          =   375
            Left            =   600
            TabIndex        =   7
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
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
            Value           =   43146.7227083333
         End
         Begin VB.Label Label17 
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
            TabIndex        =   56
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Vch No"
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
            TabIndex        =   30
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Narr"
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
            Left            =   12120
            TabIndex        =   21
            Top             =   165
            Width           =   495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Amt"
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
            Left            =   9960
            TabIndex        =   20
            Top             =   165
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Ac Code"
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
            TabIndex        =   19
            Top             =   195
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   615
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   15135
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
            Height          =   405
            Left            =   5400
            TabIndex        =   61
            Top             =   120
            Width           =   855
         End
         Begin VB.ComboBox VouTypeCombo 
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
            ItemData        =   "FrmVoucher.frx":0000
            Left            =   3240
            List            =   "FrmVoucher.frx":000D
            TabIndex        =   59
            Top             =   165
            Width           =   1215
         End
         Begin VB.TextBox TxtClBalDrcr 
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
            Left            =   14400
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   165
            Width           =   600
         End
         Begin VB.TextBox TxtClBalCash 
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
            Left            =   12960
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   165
            Width           =   1365
         End
         Begin VB.TextBox TxtCashCode 
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
            Left            =   7560
            TabIndex        =   2
            Top             =   165
            Width           =   975
         End
         Begin MSDataListLib.DataCombo DComboCash 
            Height          =   420
            Left            =   8640
            TabIndex        =   3
            Top             =   165
            Width           =   3375
            _ExtentX        =   5953
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
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   375
            Left            =   600
            TabIndex        =   1
            Top             =   165
            Width           =   1575
            _ExtentX        =   2778
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
            Value           =   43146.7227083333
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Vou No"
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
            Left            =   4560
            TabIndex        =   60
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Vou Type"
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
            TabIndex        =   58
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cl Bal"
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
            Left            =   12240
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash  A/c"
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
            TabIndex        =   17
            Top             =   210
            Width           =   1095
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
            TabIndex        =   16
            Top             =   225
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "FrmVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim AccRec As ADODB.Recordset:      Dim CashRec As ADODB.Recordset
'Dim LFParty As String:              Dim LedgerRec As ADODB.Recordset
'Dim VouRec As ADODB.Recordset:      Dim LFBPress As Byte
'
'Private Sub ComboVouType_GotFocus()
'    Sendkeys "%{DOWN}"
'End Sub
'
'Private Sub Check1_Click()
'If Check1.Value = 0 Then
'    Check1.Caption = "Enteries From Date"
'Else
'    Check1.Caption = "Enteries of Date"
'End If
'Call Update_VGrid
'End Sub
'
'Private Sub CmdAdd_Click()
'    If LenB(TxtCashCode.text) < 1 Then
'        MsgBox "Please Select Cash A/c"
'        TxtCashCode.SetFocus
'    End If
'    CmdAdd.Enabled = False
'    CmdMod.Enabled = False
'    Frame2.Enabled = False
'    TxtVchNo.text = Trim$(CStr(Get_VouNo()))
'    TxtVouNo.text = "C-" & TxtVchNo.text
'    TxtVchNo.Locked = True:
'    LFBPress = 1
'    Frame3.Enabled = True
'
'    Label15.Caption = "Adding New Vouchers"
'    TxtAcCode.SetFocus
'End Sub
'
'Private Sub CmdCancel_Click()
'Frame3.Enabled = False:                 TxtAcCode.text = vbNullString
'DComboCash.BoundText = vbNullString:    TxtCashCode.text = vbNullString
'DComboAcc.BoundText = vbNullString:     vcDTP1.Value = Date
'vcDTP2.Value = Date
'TxtClBalCash.text = vbNullString:       TxtClBalDrcr.text = vbNullString
'TxtVchNo.text = vbNullString:           TxtVouNo.text = vbNullString
'TxtAmt.text = vbNullString:             TxtNarr.text = vbNullString
'TxtBalance.text = vbNullString:         TxtBalDrCr.text = vbNullString
'Label5.Caption = vbNullString:          Label6.Caption = "Date wise Vouchers"
'TxtDebit.text = vbNullString:           TxtCredit.text = vbNullString
'TxtDiff.text = vbNullString:            TxtPCredit.text = vbNullString
'TxtPDebit.text = vbNullString:          TxtPBal.text = vbNullString
'Frame2.Enabled = True
'CmdMod.Enabled = True
'CmdAdd.Enabled = True
'Call Update_VGrid
'vcDTP1.SetFocus
'
'End Sub
'
'Private Sub CmdMod_Click()
'CmdMod.Enabled = False
'CmdAdd.Enabled = False
'Frame3.Enabled = True
'LFBPress = 2
'TxtVchNo.Locked = False
'Label15.Caption = "Modifying Vouchers"
'TxtVchNo.SetFocus
'End Sub
'
'Private Sub CmdSave_Click()
'Dim LVou_No  As String:     Dim LNarr As String:        Dim LAC_CODE As String:     Dim LCash_Code As String
'Dim LVou_Dt As Date:        Dim LAmount  As Double:     Dim LDrCr As String:        Dim LVchNo As Long
'
'LVou_No = TxtVouNo.text:        LCash_Code = TxtCashCode.text
'LAC_CODE = TxtAcCode.text:      LNarr = TxtNarr.text
'LDrCr = TxtDrCr.text:           LAmount = Val(TxtAmt.text)
'LVchNo = Val(TxtVchNo.text)
'If TxtAcCode.text = TxtCashCode.text Then
'    MsgBox " Cash Account and Party Account can Not be Same "
'    TxtAcCode.SetFocus
'    Exit Sub
'
'End If
'If LFBPress = 2 Then
'    MYSQL = "DELETE FROM VOUCHER WHERE COMPCODE =" & GCompCode & " AND VCHNO =" & LVchNo & ""
'    Cnn.Execute MYSQL
'End If
'If LAmount <> 0 Then
'    Call PInsert_Voucher(LVou_No, "CV", vcDTP2.Value, LCash_Code, LNarr, LVchNo, "0", 0)
'    Call PInsert_Vchamt(LVou_No, "CV", vcDTP2.Value, LDrCr, LAC_CODE, LAmount, LNarr, LVchNo, vbNullString, vbNullString)
'    If LDrCr = "D" Then
'        LDrCr = "C"
'    Else
'        LDrCr = "D"
'    End If
'    Call PInsert_Vchamt(LVou_No, "CV", vcDTP2.Value, LDrCr, LCash_Code, LAmount, LNarr, LVchNo, vbNullString, vbNullString)
'End If
'If LFBPress = 1 Then
'    TxtVchNo.text = Trim$(CStr(Get_VouNo()))
'    TxtVouNo.text = "C-" & TxtVchNo.text
'Else
'    TxtVchNo = vbNullString
'    TxtVouNo.text = vbNullString
'End If
'
'TxtNarr.text = vbNullString
'TxtAmt.text = vbNullString
'Call Update_VGrid
'Call Update_PGrid
'vcDTP2.SetFocus
'End Sub
'
'Private Sub DataGrid1_DblClick()
'DataGrid1.Col = 7
'TxtVchNo.text = DataGrid1.text
'CmdMod.Enabled = False
'LFBPress = 2
'Call TxtVchNo_Validate(False)
'
'End Sub
'
'
'Private Sub DataGrid2_DblClick()
'DataGrid2.Col = 6
'TxtVchNo.text = DataGrid2.text
'CmdMod.Enabled = False
'LFBPress = 2
'Call TxtVchNo_Validate(False)
'
'End Sub
'
'Private Sub DComboAcc_GotFocus()
'    Sendkeys "%{DOWN}"
'End Sub
'Private Sub DComboCash_GotFocus()
'    Sendkeys "%{DOWN}"
'End Sub
'Private Sub Form_Load()
'vcDTP1.Value = Date
'vcDTP2.Value = Date
'Frame3.Enabled = False
'OptDate.Value = True
'Label16.Caption = "Date wise Vouchers "
'Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
'MYSQL = "SELECT AC_CODE,NAME,OP_BAL FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
'AccRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'If Not AccRec.EOF Then
'    Set DComboAcc.RowSource = AccRec
'    DComboAcc.BoundColumn = "AC_CODE"
'    DComboAcc.ListField = "NAME"
'Else
'    MsgBox "Create Account First"
'    Unload Me
'End If
'
'Set CashRec = Nothing: Set CashRec = New ADODB.Recordset
'MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND GCODE =10 ORDER BY NAME"
'CashRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'If Not CashRec.EOF Then
'    Set DComboCash.RowSource = CashRec
'    DComboCash.BoundColumn = "AC_CODE"
'    DComboCash.ListField = "NAME"
'Else
'    MsgBox " Cash Account not Created "
'    DComboCash.Enabled = False
'End If
'Call Update_VGrid
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Sendkeys "{TAB}"
'End Sub
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        On Error Resume Next
'        If Me.ActiveControl.NAME = "vcDTP1" Or Me.ActiveControl.NAME = "vcDTP2" Then
'            Sendkeys "{tab}"
'        End If
'    End If
'
'End Sub
'Public Sub Update_VGrid()
'    Dim LOpFlag As Boolean:         Dim TRec As ADODB.Recordset:    Dim LBal As Double:             Dim LDr As String
'    Dim LAMT As Double:             Dim LVou_Type As String:        Dim LVou_Dt As Date:
'    Dim LVTotDebit As Double:       Dim LVTotCredit As Double:      Dim LVTotDiff As Double:
'
'    MYSQL = " SELECT A.AC_CODE,A.NAME, B.VOU_DT,B.DR_CR,B.VOU_NO, B.AMOUNT,B.NARRATION ,B.VOU_TYPE,B.VCHNO FROM "
'    MYSQL = MYSQL & " ACCOUNTM AS A, VCHAMT AS B,VOUCHER AS V WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE = B.COMPCODE And A.AC_CODE = B.AC_CODE"
'    MYSQL = MYSQL & " AND V.COMPCODE =A.COMPCODE AND V.VOU_NO =B.VOU_NO  AND V.VOU_TYPE='CV'"
'    MYSQL = MYSQL & " AND B.VOU_DT "
'    If Check1.Value = 0 Then MYSQL = MYSQL & " >"
'    MYSQL = MYSQL & " ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND v.VOU_TYPE='CV'"
'    MYSQL = MYSQL & " AND  A.AC_CODE <> '" & TxtCashCode.text & "'"
'    MYSQL = MYSQL & " AND V.VOU_NO IN (SELECT DISTINCT VOU_NO FROM VCHAMT WHERE COMPCODE =" & GCompCode & "AND AC_CODE ='" & TxtCashCode.text & "')"
'    MYSQL = MYSQL & " ORDER BY B.VOU_DT,B.VCHNO "
'    Set TRec = Nothing:        Set TRec = New ADODB.Recordset
'    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'    Call SetVouRec
'    LBal = 0
'    Do While Not TRec.EOF
'        LAMT = 0
'        VouRec.AddNew
'        VouRec!Code = TRec!AC_CODE
'        VouRec!NAME = TRec!NAME
'        VouRec!Date = Format(TRec!VOU_DT, "YYYY/MM/DD")
'        VouRec!VTYPE = TRec!VOU_TYPE
'        LAMT = TRec!AMOUNT
'        If TRec!DR_CR = "D" Then
'            LBal = LBal + (LAMT * -1)
'            VouRec!DEBITAMT = Format(LAMT, "0.00")
'            LVTotDebit = LVTotDebit + LAMT
'        Else
'            LBal = LBal + LAMT
'            LVTotCredit = LVTotCredit + LAMT
'            VouRec!CREDITAMT = Format(LAMT, "0.00")
'        End If
'        If LBal > 0 Then
'            LDr = " Cr"
'        Else
'            LDr = " Dr"
'        End If
'        VouRec!NARRATION = Trim$(Left$(TRec!NARRATION, 100))
'        VouRec!VOUNO = TRec!VOU_NO
'        VouRec!VchNo = TRec!VchNo
'        VouRec.Update
'        TRec.MoveNext
'
'    Loop
'    LVTotDiff = LVTotCredit - LVTotDebit
'    TxtDebit.text = Format(LVTotDebit, "#,##0.00")
'    TxtCredit.text = Format(LVTotCredit, "#,##0.00")
'    TxtDiff.text = Format(LVTotDiff, "#,##0.00")
'
'    Set DataGrid2.DataSource = VouRec
'    DataGrid2.ReBind
'    DataGrid2.Refresh
'    DataGrid2.Columns(0).Width = 800:               DataGrid2.Columns(1).Width = 3500
'    DataGrid2.Columns(2).Width = 1500:              DataGrid2.Columns(3).Width = 1500:
'    DataGrid2.Columns(4).Width = 5000:              DataGrid2.Columns(5).Width = 1300:
'    DataGrid2.Columns(6).Width = 800:              DataGrid2.Columns(7).Width = 800
'    DataGrid2.Columns(8).Width = 800
'    DataGrid2.Columns(2).Alignment = dbgRight:      DataGrid2.Columns(3).Alignment = dbgRight
'    DataGrid2.Columns(6).Alignment = dbgRight
'    DataGrid2.Columns(2).NumberFormat = "0.00":     DataGrid2.Columns(3).NumberFormat = "0.00"
'    DataGrid2.Refresh
'End Sub
'
'Sub SetLedgerRec()
'    Set LedgerRec = Nothing
'    Set LedgerRec = New ADODB.Recordset
'    LedgerRec.Fields.Append "Code", adVarChar, 6, adFldIsNullable
'    LedgerRec.Fields.Append "Name", adVarChar, 100, adFldIsNullable
'    LedgerRec.Fields.Append "Date", adDate, , adFldIsNullable
'    LedgerRec.Fields.Append "DebitAmt", adDouble, 20, adFldIsNullable
'    LedgerRec.Fields.Append "CreditAmt", adDouble, , adFldIsNullable
'    LedgerRec.Fields.Append "Balance", adVarChar, 20, adFldIsNullable
'    LedgerRec.Fields.Append "Narration", adVarChar, 100, adFldIsNullable
'    LedgerRec.Fields.Append "VchNo", adDouble, , adFldIsNullable
'    LedgerRec.Fields.Append "VouNo", adVarChar, 20, adFldIsNullable
'    LedgerRec.Fields.Append "VType", adVarChar, 2, adFldIsNullable
'    LedgerRec.Open , , adOpenKeyset, adLockOptimistic
'End Sub
'
'Sub SetVouRec()
'
'    Set VouRec = Nothing
'    Set VouRec = New ADODB.Recordset
'    VouRec.Fields.Append "Code", adVarChar, 6, adFldIsNullable
'    VouRec.Fields.Append "Name", adVarChar, 100, adFldIsNullable
'    VouRec.Fields.Append "DebitAmt", adDouble, , adFldIsNullable
'    VouRec.Fields.Append "CreditAmt", adDouble, , adFldIsNullable
'    VouRec.Fields.Append "Narration", adVarChar, 100, adFldIsNullable
'    VouRec.Fields.Append "Date", adDate, , adFldIsNullable
'    VouRec.Fields.Append "VchNo", adDouble, , adFldIsNullable
'    VouRec.Fields.Append "VouNo", adVarChar, 20, adFldIsNullable
'    VouRec.Fields.Append "VType", adVarChar, 2, adFldIsNullable
'    VouRec.Open , , adOpenKeyset, adLockOptimistic
'End Sub
'
'Private Sub OptParty_Click()
'If OptParty.Value = True Then
'    VouFrame.Visible = False
'    PartyFrame.Visible = True
'    Label16.Caption = "Party Ledger"
'Else
'    VouFrame.Visible = True
'    PartyFrame.Visible = False
'    Label16.Caption = "Date wise Vouchers "
'End If
'End Sub
'
'Private Sub OptDate_Click()
'If OptDate.Value = True Then
'    VouFrame.Visible = True
'    PartyFrame.Visible = False
'    Label16.Caption = "Date wise Vouchers "
'Else
'    VouFrame.Visible = False
'    PartyFrame.Visible = True
'    Label16.Caption = "Party Ledger"
'End If
'
'
'End Sub
'
'Private Sub TxtAcCode_GotFocus()
'    TxtAcCode.SelStart = 0
'    TxtAcCode.SelLength = Len(TxtAcCode.text)
'End Sub
'
'Private Sub TxtAmt_GotFocus()
'    TxtAmt.SelStart = 0
'    TxtAmt.SelLength = Len(TxtAmt.text)
'End Sub
'Private Sub TxtAmt_KeyPress(KeyAscii As Integer)
'    KeyAscii = NUMBERChk(KeyAscii)
'End Sub
'Private Sub TxtAmt_Validate(Cancel As Boolean)
'If LenB(TxtAmt.text) < 1 Then
'    MsgBox "Amount can not be blank "
'    Cancel = True
'Else
'    If Val(TxtAmt.text) < 0 Then
'        MsgBox "Amount can not be less than zero  "
'        Cancel = True
'    Else
'        TxtAmt.text = Format(TxtAmt.text, "0.00")
'    End If
'End If
'
'End Sub
'
'Private Sub TxtCashCode_Validate(Cancel As Boolean)
'Dim LAcCode As String
'If LenB(TxtAcCode.text) = 0 Then
'    DComboCash.SetFocus
'Else
'    LAcCode = Get_AccountMCode(TxtAcCode.text)
'    If LenB(LAcCode) > 1 Then
'        DComboCash.BoundText = LAcCode
'        Call Update_VGrid
'    Else
'        DComboCash.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub TxtVchNo_Validate(Cancel As Boolean)
'Dim LVch_No As Long
'Dim TRec As ADODB.Recordset
'If LenB(TxtVchNo.text) > 0 Then
'    If LFBPress = 2 Then
'        LVch_No = Val(TxtVchNo.text)
'        MYSQL = "SELECT A.AC_CODE,A.DR_CR,A.NARRATION,A.VOU_DT,A.VOU_NO,A.VOU_TYPE,B.CAC_CODE,A.AMOUNT FROM VCHAMT AS A, VOUCHER AS B "
'        MYSQL = MYSQL & " WHERE A.COMPCODE =" & GCompCode & "  AND A.COMPCODE =B.COMPCODE AND A.VCHNO =B.VCHNO"
'        MYSQL = MYSQL & " AND A.VCHNO =" & LVch_No & " AND A.AC_CODE <> B.CAC_CODE "
'        Set TRec = Nothing
'        Set TRec = New ADODB.Recordset
'        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'        If TRec.EOF Then
'            MsgBox " Invalid Voucher No "
'            Cancel = True
'        Else
'            Frame3.Enabled = True
'            TxtAcCode.text = TRec!AC_CODE
'            DComboAcc.BoundText = TRec!AC_CODE
'            TxtDrCr.text = TRec!DR_CR
'            TxtAmt.text = Format(CStr(TRec!AMOUNT), "0.00")
'            TxtNarr.text = TRec!NARRATION
'            vcDTP2.Value = TRec!VOU_DT
'            TxtVouNo.text = TRec!VOU_NO
'            TxtCashCode.text = TRec!CAC_CODE
'            DComboCash.BoundText = TRec!CAC_CODE
'            Call Update_AccBalance
'            Call Update_PGrid
'            Call Update_VGrid
'            TxtAcCode.SetFocus
'        End If
'    End If
'Else
'    MsgBox "Vou No can not be blank "
'    Call CmdCancel_Click
'End If
'End Sub
'
'Private Sub TxtAcCode_Validate(Cancel As Boolean)
'Dim LAcCode As String
'If LenB(TxtAcCode.text) = 0 Then
'    DComboAcc.SetFocus
'Else
'    LAcCode = Get_AccountMCode(TxtAcCode.text)
'    If LenB(LAcCode) > 1 Then
'        DComboAcc.BoundText = LAcCode
'        Call Update_AccBalance
'        Call Update_PGrid
'        TxtDrCr.SetFocus
'    Else
'        DComboAcc.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub DComboACC_Validate(Cancel As Boolean)
'Dim LAcCode As String
'If LenB(DComboAcc.BoundText) = 0 Then
'    MsgBox "Party can not be blank"
'    Cancel = True
'    Sendkeys "%{DOWN}"
'Else
'    LAcCode = Get_AccountMCode(DComboAcc.BoundText)
'    If LenB(LAcCode) > 1 Then
'        TxtAcCode.text = LAcCode
'        Update_AccBalance
'        Call Update_PGrid
'        TxtDrCr.SetFocus
'    Else
'        DComboAcc.SetFocus
'        Cancel = True
'        Sendkeys "%{DOWN}"
'    End If
'End If
'End Sub
'
'Private Sub DComboCash_Validate(Cancel As Boolean)
'Dim LAcCode As String
'If LenB(DComboCash.BoundText) = 0 Then
'    MsgBox "Cash Account can not be blank"
'    Cancel = True
'    Sendkeys "%{DOWN}"
'Else
'    LAcCode = Get_AccountMCode(DComboCash.BoundText)
'    If LenB(LAcCode) > 1 Then
'        TxtCashCode.text = LAcCode
'      Call Update_CashBalance
'      Call Update_VGrid
'      CmdAdd.SetFocus
'    Else
'        DComboCash.SetFocus
'        Cancel = True
'        Sendkeys "%{DOWN}"
'    End If
'End If
'End Sub
'Private Sub Update_CashBalance()
'Dim LClBal As Double
'Dim LDate As Date
'LDate = Date
'    LClBal = Get_ClosingBal(TxtCashCode.text, LDate)
'    TxtClBalCash = Format(Abs(LClBal), "#,##0.00")
'    If LClBal > 0 Then
'        TxtClBalDrcr.text = "Cr"
'    Else
'        TxtClBalDrcr.text = "Dr"
'    End If
'End Sub
'Private Sub Update_AccBalance()
'Dim LClBal As Double
'Dim LDate As Date
'LDate = Date
'    LClBal = Get_ClosingBal(TxtAcCode.text, LDate)
'    TxtBalance = Format(Abs(LClBal), "#,##0.00")
'    If LClBal > 0 Then
'        TxtBalDrCr.text = "Cr"
'    Else
'        TxtBalDrCr.text = "Dr"
'    End If
'End Sub
'Private Sub TxtDrCr_KeyPress(KeyAscii As Integer)
'If Val(KeyAscii) >= 48 And KeyAscii <= 122 Then
'    If Val(KeyAscii) = 66 Or Val(KeyAscii) = 98 Or Val(KeyAscii) = 83 Or Val(KeyAscii) = 115 Then
'    Else
'        If TxtDrCr.text = "D" Then
'            TxtDrCr.text = "C"
'        Else
'            TxtDrCr.text = "D"
'        End If
'    End If
'End If
'If KeyAscii = 32 Then
'    If TxtDrCr.text = "D" Then
'        TxtDrCr.text = "C"
'    Else
'        TxtDrCr.text = "D"
'    End If
'End If
'End Sub
'
'Private Sub TxtDrCr_Validate(Cancel As Boolean)
'If TxtDrCr.text <> "D" Then
'    If TxtDrCr.text <> "C" Then
'        TxtDrCr.text = "D"
'        Cancel = True
'        TxtDrCr.SetFocus
'    End If
'End If
'End Sub
'Public Sub Update_PGrid()
'    Dim LOpFlag As Boolean:         Dim TRec As ADODB.Recordset:    Dim LBal As Double:             Dim LDr As String
'    Dim LAMT As Double:             Dim LVou_Type As String:        Dim LVou_Dt As Date:            Dim LPDebit  As Double
'    Dim LPCredit As Double:         Dim LPDiff  As Double:
'    Dim LPTotDebit As Double:       Dim LPTotCredit As Double:      Dim LPToDiff As Double
'
'    TxtPCredit.text = vbNullString
'    TxtPDebit.text = vbNullString
'    TxtPBal.text = vbNullString
'    LFParty = TxtAcCode.text
'    If LenB(LFParty) > 1 Then
'        MYSQL = " SELECT A.AC_CODE,A.NAME,A.OP_BAL,B.VOU_TYPE,B.VOU_DT,B.DR_CR,B.VOU_NO, B.AMOUNT,B.NARRATION,VCHNO From"
'        MYSQL = MYSQL & " ACCOUNTM AS A, VCHAMT AS  B"
'        MYSQL = MYSQL & " WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE = B.COMPCODE And A.AC_CODE = B.AC_CODE"
'        MYSQL = MYSQL & " AND A.AC_CODE='" & LFParty & "'"
'        MYSQL = MYSQL & " ORDER BY A.NAME,B.VOU_DT,B.VOU_TYPE,B.VOU_NO"
'        Set TRec = Nothing:        Set TRec = New ADODB.Recordset
'        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'        Call SetLedgerRec
'        LBal = 0:        LPDebit = 0
'        LPCredit = 0:    LPDiff = 0
'        Do While Not TRec.EOF
'            LVou_Type = TRec!VOU_TYPE
'            LVou_Dt = TRec!VOU_DT
'            If LOpFlag = False Then
'                If TRec!OP_BAL <> 0 Then
'                    LedgerRec.AddNew
'                    LedgerRec!Code = TRec!AC_CODE
'                    LedgerRec!NAME = TRec!NAME
'                    LedgerRec!Date = Format(GFinBegin, "YYYY/MM/DD")
'                    LedgerRec!VTYPE = "Op"
'                    LBal = TRec!OP_BAL
'                    If LBal > 0 Then
'                        LedgerRec!CREDITAMT = LBal
'                        LDr = " Cr"
'                        LPCredit = LBal
'                    Else
'                        LedgerRec!DEBITAMT = Abs(LBal)
'                        LDr = " Dr"
'                        LPDebit = Abs(LBal)
'                    End If
'                    LedgerRec!Balance = Trim$(CStr(Format(TRec!OP_BAL, "0.00"))) & LDr
'                    LedgerRec!NARRATION = "Op Bal"
'                    LedgerRec!VOUNO = vbNullString
'                    LedgerRec.Update
'                End If
'            End If
'            LOpFlag = True
'            LAMT = 0
'            LedgerRec.AddNew
'            LedgerRec!Code = TRec!AC_CODE
'            LedgerRec!NAME = TRec!NAME
'            LedgerRec!Date = Format(TRec!VOU_DT, "YYYY/MM/DD")
'            LedgerRec!VTYPE = TRec!VOU_TYPE
'            LAMT = TRec!AMOUNT
'            If TRec!DR_CR = "D" Then
'                LBal = LBal + (LAMT * -1)
'                LedgerRec!DEBITAMT = LAMT
'                LPDebit = LPDebit + LAMT
'            Else
'                LBal = LBal + LAMT
'                LedgerRec!CREDITAMT = LAMT
'                LPCredit = LPCredit + LAMT
'            End If
'            If LBal > 0 Then
'                LDr = " Cr"
'            Else
'                LDr = " Dr"
'            End If
'            LedgerRec!Balance = Trim$(CStr(Format(Abs(LBal), "0.00"))) & LDr
'            LedgerRec!NARRATION = Trim$(Left$(TRec!NARRATION, 100))
'            LedgerRec!VOUNO = TRec!VOU_NO
'            LedgerRec!VchNo = TRec!VchNo
'            LedgerRec.Update
'            TRec.MoveNext
'        Loop
'        TxtPCredit.text = Format(LPCredit, "#,##0.00")
'        TxtPDebit.text = Format(LPDebit, "#,##0.00")
'        LPDiff = LPCredit - LPDebit
'        If LPDiff > 0 Then
'            TxtPBal.text = Format(LPDiff, "#,##0.00") & " Cr"
'        Else
'            TxtPBal.text = Format(Abs(LPDiff), "#,##0.00") & " Dr"
'        End If
'        Set DataGrid1.DataSource = LedgerRec
'
'        DataGrid1.ReBind
'        DataGrid1.Refresh
'        DataGrid1.Columns(0).Width = 900:               DataGrid1.Columns(1).Width = 2500
'        DataGrid1.Columns(2).Width = 1500:              DataGrid1.Columns(3).Width = 1500:
'        DataGrid1.Columns(4).Width = 1600:              DataGrid1.Columns(5).Width = 1500:
'        DataGrid1.Columns(6).Width = 4000:              DataGrid1.Columns(7).Width = 1000
'        DataGrid1.Columns(8).Width = 1000:              DataGrid1.Columns(9).Width = 1000:
'        DataGrid1.Columns(9).Alignment = dbgCenter:
'        DataGrid1.Columns(3).Alignment = dbgRight:      DataGrid1.Columns(4).Alignment = dbgRight
'        DataGrid1.Columns(5).Alignment = dbgRight
'        DataGrid1.Columns(3).NumberFormat = "0.00":     DataGrid1.Columns(4).NumberFormat = "0.00"
'    End If
'
'End Sub
'Private Sub vcDTP1_Validate(Cancel As Boolean)
'vcDTP2.Value = vcDTP1.Value
'End Sub
