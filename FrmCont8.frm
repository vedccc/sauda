VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCont8 
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   14715
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   9015
      Left            =   15360
      TabIndex        =   69
      Top             =   0
      Width           =   6255
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   8655
         Left            =   120
         TabIndex        =   70
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   15266
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6135
      Left            =   0
      TabIndex        =   53
      Top             =   2760
      Width           =   15255
      Begin VB.CommandButton Command9 
         Caption         =   "Delete All Trades Below"
         Height          =   375
         Left            =   12360
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton CmdFilter 
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
         TabIndex        =   63
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox TxtDiffAmt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9360
         TabIndex        =   62
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton CmdClearFilter 
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
         Left            =   240
         TabIndex        =   61
         Top             =   120
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5055
         Left            =   6360
         TabIndex        =   54
         Top             =   720
         Visible         =   0   'False
         Width           =   6015
         Begin VB.ComboBox FilterFieldCombo 
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
            ItemData        =   "FrmCont8.frx":0000
            Left            =   1200
            List            =   "FrmCont8.frx":0013
            TabIndex        =   57
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton CmdFilterOK 
            Caption         =   "Ok"
            Height          =   375
            Left            =   3960
            TabIndex        =   56
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5040
            TabIndex        =   55
            Top             =   600
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3615
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   6376
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
         Begin VB.Label Label15 
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
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label20 
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
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   5775
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   65
         Top             =   600
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   5295
         Left            =   120
         TabIndex        =   66
         Top             =   720
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   9340
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   18
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
         Left            =   7320
         TabIndex        =   68
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label14 
         Height          =   375
         Left            =   3600
         TabIndex        =   67
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1575
      Left            =   0
      TabIndex        =   20
      Top             =   1200
      Width           =   15135
      Begin VB.TextBox TxtConRate 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   12600
         TabIndex        =   35
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtConNo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   34
         Top             =   480
         Width           =   855
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
         Height          =   420
         Left            =   14040
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtRate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6405
         TabIndex        =   32
         Top             =   480
         Width           =   1320
      End
      Begin VB.TextBox TxtQty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2325
         TabIndex        =   31
         Top             =   480
         Width           =   1020
      End
      Begin VB.TextBox TxtContype 
         Height          =   375
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "Buy"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtPtyCode 
         Height          =   360
         Left            =   7800
         MaxLength       =   6
         TabIndex        =   29
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox TxtBrokerCode 
         Height          =   375
         Left            =   11400
         MaxLength       =   6
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtLot 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1635
         TabIndex        =   27
         Top             =   480
         Width           =   650
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   490
         Left            =   4680
         TabIndex        =   22
         Top             =   960
         Width           =   4575
         Begin VB.TextBox TxtSettleRate 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   80
            Width           =   1215
         End
         Begin VB.TextBox TxtRefLot 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   960
            TabIndex        =   23
            Top             =   80
            Width           =   1095
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Settle Rate"
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
            Left            =   2160
            TabIndex        =   26
            Top             =   135
            Width           =   975
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot Size"
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
            TabIndex        =   25
            Top             =   133
            Width           =   735
         End
      End
      Begin VB.CheckBox ChkAppBrok 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Apply Brokerage"
         Height          =   375
         Left            =   9600
         TabIndex        =   21
         Top             =   960
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   300
         Left            =   2880
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
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
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSDataListLib.DataCombo DComboTSauda 
         Height          =   360
         Left            =   3435
         TabIndex        =   38
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   8670
         TabIndex        =   39
         Top             =   480
         Width           =   2595
         _ExtentX        =   4577
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
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   360
         Left            =   12600
         TabIndex        =   40
         Top             =   480
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
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
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Con Rate"
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
         Left            =   11520
         TabIndex        =   52
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Con No"
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
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rate"
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
         Left            =   6405
         TabIndex        =   49
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Qty"
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
         Left            =   2325
         TabIndex        =   48
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "B/S"
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
         Left            =   1080
         TabIndex        =   47
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sauda"
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
         Left            =   3435
         TabIndex        =   46
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name"
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
         Left            =   8670
         TabIndex        =   45
         Top             =   120
         Width           =   2595
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Code"
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
         Left            =   7800
         TabIndex        =   44
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Broker Name"
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
         Left            =   12600
         TabIndex        =   43
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Code"
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
         Left            =   11400
         TabIndex        =   42
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lot"
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
         Left            =   1635
         TabIndex        =   41
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   15135
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
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   100
         Width           =   1000
      End
      Begin VB.CommandButton CmdModify 
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
         Height          =   360
         Left            =   1200
         TabIndex        =   16
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         CausesValidation=   0   'False
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
         Left            =   2280
         TabIndex        =   15
         Top             =   120
         Width           =   1000
      End
      Begin VB.TextBox TxtAdminPadd 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4560
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Mark All Trades  as Opening Standing"
         Height          =   285
         Left            =   11040
         TabIndex        =   13
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sauda Software"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         TabIndex        =   19
         Top             =   0
         Width           =   2640
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Ad Paswd"
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
         Left            =   3480
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15135
      Begin VB.CheckBox ChkShowContract 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Show All Contracts"
         Height          =   360
         Left            =   2040
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin vcDateTimePicker.vcDTP DtpCondate 
         Height          =   360
         Left            =   600
         TabIndex        =   0
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   3
         Value           =   41160.4222453704
      End
      Begin MSDataListLib.DataCombo DComboExchange 
         Height          =   360
         Left            =   4680
         TabIndex        =   2
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
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
      Begin MSDataListLib.DataCombo DComboParty 
         Height          =   360
         Left            =   6360
         TabIndex        =   3
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSDataListLib.DataCombo DComboSauda 
         Height          =   360
         Left            =   9480
         TabIndex        =   4
         Top             =   120
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSDataListLib.DataCombo DComboBroker 
         Height          =   360
         Left            =   12840
         TabIndex        =   6
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Exch"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   173
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   173
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda"
         Height          =   255
         Left            =   8760
         TabIndex        =   8
         Top             =   173
         Width           =   615
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Broker"
         Height          =   255
         Left            =   12180
         TabIndex        =   7
         Top             =   173
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCont8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LFExCode As String:                  Dim LFParty As String:              Dim LFBroker As String
Dim LBillExCodes As String:              Dim LBillParties As String:         Dim LBillSaudas As String
Dim LSParties As String:                 Dim LSSaudas As String:             Dim LSExCodes As String
Dim LSPNames As String:                  Dim LSType As String:               Dim LSUserIds As String
Dim LItemCodeDBCombo As String:          Dim LFSauda As String:              Dim LFBPress As Integer
Dim LPDataImport As Byte:                Dim SaveCalled As Boolean:          Dim LBillItems As String

Dim ExRec As ADODB.Recordset:            Dim PartyRec As ADODB.Recordset:    Dim ItemRec As ADODB.Recordset
Dim AllSaudaRec As ADODB.Recordset:      Dim SaudaRec As ADODB.Recordset:    Dim LFPartyRec As ADODB.Recordset
Dim ContRec As ADODB.Recordset:          Dim LFSaudaRec As ADODB.Recordset:  Dim LFBrokerRec As ADODB.Recordset

Private Sub FilterFieldCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub FilterFieldCombo_Validate(Cancel As Boolean)
Dim TREC As ADODB.Recordset
If FilterFieldCombo.ListIndex = 0 Then
    MYSQL = "SELECT DISTINCT A.PARTY FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
    MYSQL = MYSQL & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
    MYSQL = MYSQL & " AND A.PARTY NOT IN (SELECT DISTINCT CONCODE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "')"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE ='" & LFExCode & "'"
    MYSQL = MYSQL & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' "
    If LenB(LSParties) <> 0 Then MYSQL = MYSQL & " AND A.PARTY IN (" & LSParties & ")"
    If LenB(LSPNames) <> 0 Then MYSQL = MYSQL & " AND B.NAME IN (" & LSPNames & ")"
    If LenB(LSSaudas) <> 0 Then MYSQL = MYSQL & " AND A.SAUDA IN (" & LSSaudas & ")"
    If LenB(LSUserIds) <> 0 Then MYSQL = MYSQL & " AND A.USERID IN (" & LSUserIds & ")"
    MYSQL = MYSQL & " ORDER BY Party"
    Set TREC = Nothing:    Set TREC = New ADODB.Recordset
    TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TREC.EOF Then
        ListView1.ListItems.clear
        ListView1.Enabled = True
        Do While Not TREC.EOF
            ListView1.ListItems.Add , , TREC!PARTY
            TREC.MoveNext
        Loop
    End If
    Set TREC = Nothing
ElseIf FilterFieldCombo.ListIndex = 1 Then
    MYSQL = "SELECT DISTINCT B.NAME FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
    MYSQL = MYSQL & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
    MYSQL = MYSQL & " AND A.PARTY NOT IN (SELECT DISTINCT CONCODE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "')"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE ='" & LFExCode & "' "
    MYSQL = MYSQL & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' "
    If LenB(LSParties) <> 0 Then MYSQL = MYSQL & " AND A.PARTY IN (" & LSParties & ")"
    If LenB(LSPNames) <> 0 Then MYSQL = MYSQL & " AND B.NAME IN (" & LSPNames & ")"
    If LenB(LSSaudas) <> 0 Then MYSQL = MYSQL & " AND A.SAUDA IN (" & LSSaudas & ")"
    If LenB(LSUserIds) <> 0 Then MYSQL = MYSQL & " AND A.USERID IN (" & LSUserIds & ")"
    MYSQL = MYSQL & " ORDER BY B.NAME"
    Set TREC = Nothing
    Set TREC = New ADODB.Recordset
    TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TREC.EOF Then
        ListView1.ListItems.clear
        ListView1.Enabled = True
        Do While Not TREC.EOF
            ListView1.ListItems.Add , , TREC!NAME
            TREC.MoveNext
        Loop
    End If
    Set TREC = Nothing
ElseIf FilterFieldCombo.ListIndex = 2 Then
    MYSQL = "SELECT DISTINCT A.SAUDA FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
    MYSQL = MYSQL & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
    MYSQL = MYSQL & " AND A.PARTY NOT IN (SELECT DISTINCT CONCODE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "')"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE ='" & LFExCode & "'"
    MYSQL = MYSQL & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' "
    If LenB(LSParties) <> 0 Then MYSQL = MYSQL & " AND A.PARTY IN (" & LSParties & ")"
    If LenB(LSPNames) <> 0 Then MYSQL = MYSQL & " AND B.NAME IN (" & LSPNames & ")"
    If LenB(LSSaudas) <> 0 Then MYSQL = MYSQL & " AND A.SAUDA IN (" & LSSaudas & ")"
    If LenB(LSUserIds) <> 0 Then MYSQL = MYSQL & " AND A.USERID IN (" & LSUserIds & ")"
    MYSQL = MYSQL & " ORDER BY SAUDA"
    Set TREC = Nothing:    Set TREC = New ADODB.Recordset
    TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TREC.EOF Then
        ListView1.ListItems.clear
        ListView1.Enabled = True
        Do While Not TREC.EOF
            ListView1.ListItems.Add , , TREC!Sauda
            TREC.MoveNext
        Loop
    End If
    Set TREC = Nothing
ElseIf FilterFieldCombo.ListIndex = 3 Then
    ListView1.ListItems.clear
    ListView1.Enabled = True
    ListView1.ListItems.Add , , "Buy"
    ListView1.ListItems.Add , , "Sell"
ElseIf FilterFieldCombo.ListIndex = 4 Then
    MYSQL = "SELECT DISTINCT A.USERID FROM CTR_D A, ACCOUNTD B WHERE A.COMPCODE=" & GCompCode & ""
    MYSQL = MYSQL & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
    MYSQL = MYSQL & " AND A.PARTY NOT IN (SELECT DISTINCT CONCODE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "')"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE ='" & LFExCode & "' "
    MYSQL = MYSQL & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' "
    If LenB(LSParties) <> 0 Then MYSQL = MYSQL & " AND A.PARTY IN (" & LSParties & ")"
    If LenB(LSPNames) <> 0 Then MYSQL = MYSQL & " AND B.NAME IN (" & LSPNames & ")"
    If LenB(LSSaudas) <> 0 Then MYSQL = MYSQL & " AND A.SAUDA IN (" & LSSaudas & ")"
    MYSQL = MYSQL & " ORDER BY A.USERID "
    Set TREC = Nothing:    Set TREC = New ADODB.Recordset
    TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TREC.EOF Then
        ListView1.ListItems.clear
        ListView1.Enabled = True
        Do While Not TREC.EOF
            ListView1.ListItems.Add , , IIf(IsNull(TREC!USERID), "", TREC!USERID)
            TREC.MoveNext
        Loop
    End If
    Set TREC = Nothing
End If
End Sub
Public Sub SHOW_STANDING()
Dim NStandRec As ADODB.Recordset
MYSQL = "SELECT A.Name ,B.Sauda,SUM(CASE B.CONTYPE WHEN 'B' THEN B.QTY * 1 WHEN 'S' THEN B.QTY*-1 END) AS NetQty"
MYSQL = MYSQL & " FROM ACCOUNTD AS A,CTR_D AS B, SAUDAMAST AS S WHERE A.COMPCODE= " & GCompCode & "  AND A.COMPCODE =B.COMPCODE AND A.COMPCODE =S.COMPCODE AND S.MATURITY>='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' AND B.SAUDA=S.SAUDACODE AND A.AC_CODE =B.PARTY AND B.CONDATE <='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND S.EXCODE='" & LFExCode & "'"
MYSQL = MYSQL & " GROUP BY A.NAME,B.SAUDA HAVING SUM(CASE B.CONTYPE WHEN 'B' THEN B.QTY * 1 WHEN 'S' THEN B.QTY*-1 END) <>0 ORDER BY A.NAME,B.SAUDA"
Set NStandRec = Nothing: Set NStandRec = New ADODB.Recordset
NStandRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
If Not NStandRec.EOF Then
    Set DataGrid2.DataSource = NStandRec
    DataGrid2.ReBind
    DataGrid2.Refresh
    DataGrid2.Columns(0).Width = 1800:
    DataGrid2.Columns(1).Width = 1800
    DataGrid2.Columns(2).Width = 800
    DataGrid2.Columns(2).Alignment = dbgRight:
End If
End Sub
Private Sub CmdSave_Click()
    Dim LExCode As String:          Dim LDelFlag As Boolean:        Dim LOConNo As String
    Dim LContime As String:         Dim LCSauda As String:          Dim LCItemCode As String
    Dim LConType As String:         Dim LSInstType As String:       Dim LStatus As String
    Dim LST_Time As String:         Dim MParty As String:           Dim LCLot As Double:
    Dim LCRefLot As Double:         Dim LCBrokLot As Double:        Dim LSCondate As Date
    Dim LConNo As Long:             Dim LClient As String:          Dim LExCont As String
    Dim MSaudaCode As String::      Dim LItemCode As String
    Dim MQty As Double:             Dim MRate As Double:            Dim LCalval As Double
    Dim MConRate As Double:         Dim LSConSno As Long:           Dim LOrdNo As String
    Dim LSOptType  As String:       Dim LSStrike As Double:         Dim LBSParty As String
    Dim LBrokFlag As String
    
    Dim TREC As ADODB.Recordset:    Dim NRec As ADODB.Recordset
    On Error GoTo ERR1
    LDelFlag = False
    DoEvents
    LSCondate = DtpCondate.Value
    If Check1.Value = 1 Then
        MYSQL = "SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & "  AND CONDATE <'" & LSCondate & "'"
        Set TREC = Nothing
        Set TREC = New ADODB.Recordset
        TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TREC.EOF Then
            MsgBox "No Trades Should be there Before Opening Standing"
            Exit Sub
        End If
        Set TREC = Nothing
        
    End If
    
    Frame1.Enabled = False
    'Frame2.Enabled = False
    If LenB(TxtConNo.text) = 0 Then
        MsgBox "Trade No can not be Blank":        Frame1.Enabled = True
        TxtConNo.Locked = False
        TxtConNo.SetFocus:
        Exit Sub
    Else
        LConNo = Val(TxtConNo.text)
    End If
    If LenB(TxtPtyCode.text) = 0 Then
        MsgBox "Party Code can not be Blank":        Frame1.Enabled = True
        TxtPtyCode.SetFocus
        Exit Sub
    Else
        MParty = Get_AccountDCode(TxtPtyCode.text)
        If LenB(MParty) < 1 Then
            MsgBox "Invalid Party Code":            Frame1.Enabled = True
            TxtPtyCode.SetFocus
            Exit Sub
        Else
            LClient = MParty
        End If
    End If
    If LenB(TxtBrokerCode.text) = 0 Then
        MsgBox "Broker A/c can not be Blank":        Frame1.Enabled = True
        TxtBrokerCode.SetFocus
        Exit Sub
    Else
        LExCont = Get_AccountDCode(TxtBrokerCode.text)
        If LenB(LExCont) < 1 Then
            MsgBox "Invalid Broker Code":            Frame1.Enabled = True
            TxtBrokerCode.SetFocus
            Exit Sub
        End If
    End If
    If MParty = LExCont Then
        MsgBox " Buyer and Seller party Can not be same Pls Correct"
        Frame1.Enabled = True:
        TxtBrokerCode.SetFocus:               Exit Sub
    End If
    LCSauda = vbNullString
    If LenB(DComboTSauda.BoundText) = 0 Then
        MsgBox "Sauda Code can not be Blank"
        Frame1.Enabled = True
        
        TxtPtyCode.SetFocus:
        Exit Sub
    Else
        If ChkShowContract.Value = 1 Then
            AllSaudaRec.MoveFirst
            AllSaudaRec.Find "SAUDACODE='" & DComboTSauda.BoundText & "'", , adSearchForward
            LCLot = AllSaudaRec!LOT:                LCRefLot = AllSaudaRec!REFLOT
            LCBrokLot = AllSaudaRec!BROKLOT:        LSInstType = AllSaudaRec!INSTTYPE
            LCItemCode = vbNullString: LCSauda = vbNullString
            LCItemCode = Get_ItemMaster(AllSaudaRec!EXCODE, AllSaudaRec!EX_SYMBOL)
            If LenB(LCItemCode) < 1 Then LCItemCode = Create_TItemMast(AllSaudaRec!ITEMCODE, AllSaudaRec!ITEMName, AllSaudaRec!EX_SYMBOL, AllSaudaRec!LOT, AllSaudaRec!EXCODE)
            If LenB(LCItemCode) < 1 Then
                Frame1.Enabled = True
                MsgBox "Import new Contracts First"
                Exit Sub
            End If
            LCSauda = Get_SaudaMaster(AllSaudaRec!EXCODE, LCItemCode, AllSaudaRec!MATURITY, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
            If LenB(LCSauda) < 1 Then LCSauda = Create_TSaudaMast(LCItemCode, AllSaudaRec!MATURITY, AllSaudaRec!EXCODE, AllSaudaRec!INSTTYPE, AllSaudaRec!OPTTYPE, AllSaudaRec!STRIKEPRICE)
            If LenB(LCSauda) < 1 Then
                MsgBox "Import new Contracts First"
                Frame1.Enabled = True
                Exit Sub
            End If
        Else
            LCSauda = DComboTSauda.BoundText
        End If
        MYSQL = "SELECT A.SAUDACODE,B.ITEMCODE,B.LOT,A.TRADEABLELOT,B.EXCHANGECODE,INSTTYPE,OPTTYPE,STRIKEPRICE FROM SAUDAMAST A,ITEMMAST B WHERE A.COMPCODE =" & GCompCode & "AND A.SAUDACODE ='" & LCSauda & "'AND A.COMPCODE =B.COMPCODE AND A.ITEMCODE =B.ITEMCODE "
        Set NRec = Nothing:        Set NRec = New ADODB.Recordset
        NRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If NRec.EOF Then
            MsgBox "Invalid Sauda Code":
            Frame1.Enabled = True
            DComboTSauda.SetFocus
            Exit Sub
        Else
            MSaudaCode = NRec!SAUDACODE:            LItemCode = Trim(NRec!ITEMCODE)
            LExCode = NRec!EXCHANGECODE:            LSInstType = NRec!INSTTYPE
            LSOptType = NRec!OPTTYPE:               LSStrike = NRec!STRIKEPRICE
            LCalval = NRec!LOT
            If NRec!EXCHANGECODE = "NSE" Then LCalval = NRec!TRADEABLELOT
        End If
        Set NRec = Nothing
    End If
    If Val(TxtQty.text) = 0 Then
        If LFBPress = 1 Then
            MsgBox "Trade Qty can not be Zero "
            Frame1.Enabled = True:
            TxtQty.SetFocus:        Exit Sub
        Else
            If MsgBox("You are about to Delete this Trade. Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
                LDelFlag = True
            Else
                MsgBox "Trade Qty can not be Zero ":
                Frame1.Enabled = True:
                TxtQty.SetFocus:                Exit Sub
            End If
        End If
    Else
        MQty = Round(Val(TxtQty.text), 2)
    End If
    If Val(TxtRate.text) = 0 Then
        MsgBox "Trade Rate can not be Zero ":
        Frame1.Enabled = True
        TxtRate.SetFocus
        Exit Sub
    Else
        MRate = Round(Val(TxtRate.text), 4)
    End If
    MConRate = 0
    If Val(TxtConRate.text) = 0 Then
        MsgBox "Trade Con Rate can not be Zero "
        Frame1.Enabled = True:
        TxtConRate.SetFocus:          Exit Sub
    Else
        MConRate = Round(Val(TxtConRate.text), 4)
    End If
    LSConSno = Get_ConSNo(LSCondate, MSaudaCode, LItemCode, LExCode)
    LOConNo = LConNo:    LContime = Time:    LOrdNo = LTrim$(RTrim$(Str(LConNo)))
    DoEvents
    CNNERR = True
    Cnn.BeginTrans
    If LFBPress = 2 Then
        LConNo = Val(TxtConNo.text):        LOConNo = Trim(Text7.text)
        LOrdNo = Trim(TxtOrdNo.text)
        MYSQL = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONNO=" & Val(TxtConNo.text) & "  AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
        Cnn.Execute MYSQL
    End If
    If LDelFlag = False Then
        If LFBPress = 1 Then
            If GConNoType <> 0 Then
                LConNo = Get_Max_ConNo(LSCondate, LFExCode)
            Else
                LConNo = Get_Max_ConNo(LSCondate, vbNullString)
            End If
            
            LConNo = LConNo + 1
            LOConNo = LConNo
        End If
        If TxtContype.text = "Buy" Then
            LConType = "B"
        Else
            LConType = "S"
        End If
        If TxtContype.text = "Buy" Then
            LConType = "B"
        Else
            LConType = "S"
        End If
        If ChkAppBrok.Value = 1 Then
            LBrokFlag = "Y"
        Else
            LBrokFlag = "N"
        End If
        Call Add_To_Ctr_D2(LConType, LClient, LSConSno, LSCondate, LConNo, MSaudaCode, LItemCode, MParty, MQty, MRate, MConRate, LExCont, LContime, LOrdNo, vbNullString, LOConNo, LExCode, LCalval, LPDataImport, vbNullString, LSInstType, LSOptType, LSStrike, Left$(TxtFileType.text, 1), LBrokFlag)
        If Check1.Value = 1 Then
            MYSQL = MYSQL & "UPDATE CTR_D SET PATTAN ='O',BROKFLAG='N' WHERE COMPCODE =" & GCompCode & " AND CONSNO =" & LSConSno & " And CONNO = " & LConNo & " AND CONATE ='" & Format(LSCondate) & "'"
            Cnn.Execute MYSQL
            MYSQL = MYSQL & "UPDATE CTR_M SET PATTAN ='O' WHERE COMPCODE =" & GCompCode & " AND CONSNO =" & LSConSno & "  AND CONATE ='" & Format(LSCondate) & "'"
            Cnn.Execute MYSQL
        End If
        LBSParty = "'" & MParty & "','" & LExCont & "'"
        'Call Delete_Inv_D(LBSParty, "'" & LEXCODE & "'", "'" & MSaudaCode & "'", DtpCondate.Value)
    End If
    Cnn.CommitTrans
    If LenB(LBillParties) < 1 Then
        LBillParties = "'" & MParty & "','" & LExCont & "'"
    Else
        If InStr(LBillParties, "'" & MParty & "'") < 1 Then LBillParties = LBillParties & ",'" & MParty & "'"
        If InStr(LBillParties, "'" & LExCont & "'") < 1 Then LBillParties = LBillParties & ",'" & LExCont & "'"
    End If
    If LenB(LBillExCodes) < 1 Then
        LBillExCodes = "'" & LExCode & "'"
    Else
        If InStr(LBillExCodes, LExCode) < 1 Then LBillExCodes = LBillExCodes & ",'" & LExCode & "'"
    End If
    If LenB(LBillItems) < 1 Then
        LBillItems = "'" & LItemCode & "'"
    Else
        If InStr(LBillItems, LItemCode) < 1 Then LBillItems = LBillItems & "," & "'" & LItemCode & "'"
    End If
    If LenB(LBillSaudas) < 1 Then
        LBillSaudas = "'" & MSaudaCode & "'"
    Else
        If InStr(LBillSaudas, MSaudaCode) < 1 Then LBillSaudas = LBillSaudas & ",'" & MSaudaCode & "'"
    End If
    CNNERR = False:
    LConNo = LConNo + 1
    Call DATA_GRID_REFRESH
    Call SHOW_STANDING
    TxtLot = vbNullString:          TxtQty.text = vbNullString:
    TxtRate.text = vbNullString:    TxtConRate.text = vbNullString
    Frame1.Enabled = True
    If LFBPress = 2 Then
        TxtConNo.text = vbNullString:               TxtPtyCode.text = vbNullString
        DataCombo2.BoundText = vbNullString:        DComboTSauda.BoundText = vbNullString
        TxtOrdNo = vbNullString:                    TxtConNo.SetFocus
    Else
        TxtConNo.text = LConNo:                     TxtPtyCode.SetFocus
    End If
    SaveCalled = True
    Exit Sub
ERR1:
If err.Number <> 0 Then
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    'Resume
    Frame1.Enabled = True
   If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End If
End Sub
Private Sub CmdAdd_Click()
   ' Call Chk_Brokerage(DtpCondate.Value)
    If GConNoType <> 0 Then
        TxtConNo.text = Trim(Get_Max_ConNo(DtpCondate.Value, LFExCode) + 1)
    Else
        TxtConNo.text = Trim(Get_Max_ConNo(DtpCondate.Value, vbNullString) + 1)
    End If
    
    TxtConNo.Locked = True:                             TxtQty.Locked = False
    TxtPtyCode.text = vbNullString:                     DataCombo2.BoundText = vbNullString
    TxtLot = vbNullString:                              TxtQty.text = vbNullString
    TxtRate.text = vbNullString:                        TxtConRate.text = vbNullString
    CmdModify.Enabled = False:                          Frame2.Enabled = True
    TxtPtyCode.SetFocus:                                CmdAdd.Enabled = False
    CmdCancel.Enabled = True
    DtpCondate.Enabled = False
    LFBPress = 1
    LPDataImport = "0"
    Label12.Caption = "Adding New Trades"
End Sub
Private Sub CmdModify_Click()
    Call Mod_Rec
End Sub
Private Sub CmdCancel_Click()

Call CANCEL_REC
End Sub
Private Sub CmdFilterOK_Click()
Dim I As Integer
If FilterFieldCombo.ListIndex = 0 Then
    LSParties = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LSParties) <> 0 Then LSParties = LSParties & ", "
            LSParties = LSParties & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
ElseIf FilterFieldCombo.ListIndex = 1 Then
    LSPNames = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LSPNames) <> 0 Then LSPNames = LSPNames & ", "
            LSPNames = LSPNames & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
ElseIf FilterFieldCombo.ListIndex = 2 Then
    LSSaudas = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LSSaudas) <> 0 Then LSSaudas = LSSaudas & ", "
            LSSaudas = LSSaudas & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
ElseIf FilterFieldCombo.ListIndex = 3 Then
    LSType = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LSType) <> 0 Then LSType = LSType & ", "
            LSType = LSType & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
ElseIf FilterFieldCombo.ListIndex = 4 Then
    LSUserIds = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LSUserIds) <> 0 Then LSUserIds = LSUserIds & ","
            LSUserIds = LSUserIds & "'" & ListView1.ListItems(I) & "'"
        End If
  Next I
End If
Call DATA_GRID_REFRESH
Frame4.Visible = False
End Sub
Private Sub Command6_Click()
Frame4.Visible = False
DataGrid1.SetFocus
End Sub
Private Sub CmdClearFilter_Click()
LSParties = vbNullString:    LSSaudas = vbNullString:    LSPNames = vbNullString:    LSUserIds = vbNullString
Call DATA_GRID_REFRESH
End Sub
Private Sub CmdFilter_Click()
    If CmdAdd.Enabled = False Then
        If ContRec.RecordCount > 0 Then
            Frame4.Visible = True:                  DoEvents
            ListView1.ListItems.clear:              FilterFieldCombo.SetFocus
        End If
    End If
End Sub
Private Sub Command9_Click()
Dim LDel As Boolean
If ContRec.RecordCount > 0 Then
    If MsgBox("Are You Sure You Want to Delte all Trades of " & DtpCondate.Value & " of " & DComboExchange.BoundText & "", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
        If Not ContRec.EOF Then
            ContRec.MoveFirst
            Do While Not ContRec.EOF
                MYSQL = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & ContRec!EXCODE & "' AND CONNO = " & ContRec!CONNO & " AND CONDATE ='" & Format(ContRec!Condate, "YYYY/MM/DD") & "'"
                Cnn.Execute MYSQL
                ContRec.MoveNext
            Loop
        End If
        Call Delete_Inv_D(vbNullString, vbNullString, vbNullString, DtpCondate.Value)
    End If
    DATA_GRID_REFRESH
End If
End Sub
Private Sub DComboExchange_Validate(Cancel As Boolean)
If LenB(DComboExchange.BoundText) = 0 Then
    LFExCode = vbNullString
    If GConNoType <> 0 Then
        MsgBox "Please Select Exchange"
        Cancel = True
        Exit Sub
    End If
Else
    LFExCode = DComboExchange.BoundText
    ExRec.Filter = adFilterNone
    ExRec.Filter = "EXCODE='" & LFExCode & "'"
End If
    If LFExCode = "LME" Then
        DataCombo5.Visible = True
        vcDTP2.Visible = True
        vcDTP2.Value = DtpCondate.Value + 90
        Set ItemRec = Nothing
        MYSQL = "SELECT ITEMCODE, ITEMNAME,EXCHANGECODE, LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & "  AND EXCHANGECODE ='LME' ORDER BY ITEMCODE "
        Set ItemRec = New ADODB.Recordset
        ItemRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        Set DataCombo5.RowSource = ItemRec:
        DataCombo5.BoundColumn = "ITEMCODE"
        DataCombo5.ListField = "ITEMNAME"
        
    End If
    FillTradeSaudaCombo
    Set LFPartyRec = Nothing
    Set LFPartyRec = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT A.AC_CODE,A.NAME FROM ACCOUNTD AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
    MYSQL = MYSQL & " AND A.AC_CODE =B.PARTY AND  B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
    MYSQL = MYSQL & " ORDER BY A.NAME"
    LFPartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not LFPartyRec.EOF Then
        DComboParty.Enabled = True
        Set DComboParty.RowSource = LFPartyRec
        DComboParty.BoundColumn = "AC_CODE"
        DComboParty.ListField = "NAME"
    Else
        DComboParty.Enabled = False
    End If
    Set LFSaudaRec = Nothing
    Set LFSaudaRec = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT A.SAUDACODE,A.SAUDANAME,A.MATURITY,A.ITEMCODE,A.REFLOT FROM SAUDAMAST AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
    MYSQL = MYSQL & " AND A.SAUDACODE =B.SAUDA AND B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
    MYSQL = MYSQL & " ORDER BY A.ITEMCODE,A.MATURITY"
    LFSaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not LFPartyRec.EOF Then
        DComboSauda.Enabled = True:                 Set DComboSauda.RowSource = LFSaudaRec
        DComboSauda.BoundColumn = "SAUDACODE":      DComboSauda.ListField = "SAUDANAME"
    Else
        DComboSauda.Enabled = False
    End If
Call DATA_GRID_REFRESH
Call SHOW_STANDING
End Sub
Private Sub DataCombo4_Validate(Cancel As Boolean)
Dim NRec As ADODB.Recordset
If LenB(DataCombo4.text) = 0 Then
    MsgBox "Broker A/c can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & DataCombo4.BoundText & "'"
    Set NRec = Nothing:         Set NRec = New ADODB.Recordset
    NRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not NRec.EOF Then
        TxtBrokerCode.text = NRec!AC_CODE
    Else
        DataCombo4.SetFocus
        Cancel = True
        Sendkeys "%{DOWN}"
    End If
    TxtConRate.SetFocus
End If
End Sub
Private Sub DataCombo2_Validate(Cancel As Boolean)
'Dim NRec As ADODB.Recordset
Dim LAcCode As String

If LenB(DataCombo2.text) = 0 Then
    MsgBox "Party can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    LAcCode = Get_AccountDCode(DataCombo2.BoundText)
    If LenB(LAcCode) > 1 Then
        TxtPtyCode.text = LAcCode
        If Frame2.Enabled = False Then
            Frame10.Enabled = True
            DtpCondate.Enabled = True
            DtpCondate.SetFocus
        Else
            DComboTSauda.SetFocus
        End If
    Else
        DataCombo2.SetFocus
        Cancel = True
        Sendkeys "%{DOWN}"
    End If
End If
End Sub
Private Sub DComboTSauda_Validate(Cancel As Boolean)
If LenB(DComboTSauda.text) = 0 Then
    MsgBox "Sauda can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    Call Get_Value
End If
TxtSettleRate.text = Format(SDCLRATE(DComboTSauda.text, DtpCondate.Value, "C"), "0.00")
End Sub
Private Sub DataCombo5_Validate(Cancel As Boolean)
If LenB(DataCombo5.text) = 0 Then
    MsgBox "Itemcode can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    LItemCodeDBCombo = DataCombo5.BoundText
End If
End Sub
Private Sub DComboParty_Validate(Cancel As Boolean)
If LenB(DComboParty.BoundText) <> 0 Then
    LFParty = DComboParty.BoundText
Else
    LFParty = vbNullString
End If
Set LFSaudaRec = Nothing
Set LFSaudaRec = New ADODB.Recordset
MYSQL = "SELECT DISTINCT A.SAUDACODE,A.SAUDANAME,A.MATURITY,A.ITEMCODE,A.REFLOT FROM SAUDAMAST AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
MYSQL = MYSQL & " AND A.SAUDACODE =B.SAUDA AND B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
If LenB(LFParty) <> 0 Then MYSQL = MYSQL & " AND B.PARTY ='" & LFParty & "'"
MYSQL = MYSQL & " ORDER BY A.ITEMCODE,A.MATURITY"
LFSaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
If Not LFPartyRec.EOF Then
    DComboSauda.Enabled = True
    Set DComboSauda.RowSource = LFSaudaRec
    DComboSauda.BoundColumn = "SAUDACODE"
    DComboSauda.ListField = "SAUDANAME"
Else
    DComboSauda.Enabled = False
End If
Call DATA_GRID_REFRESH
End Sub

Private Sub DComboSauda_Validate(Cancel As Boolean)



If LenB(DComboSauda.BoundText) <> 0 Then
    LFSauda = DComboSauda.BoundText
Else
    LFSauda = vbNullString
End If
Set LFBrokerRec = Nothing
Set LFBrokerRec = New ADODB.Recordset
MYSQL = "SELECT DISTINCT A.AC_CODE,A.NAME FROM ACCOUNTD  AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
MYSQL = MYSQL & " AND A.AC_CODE  =B.CONCODE  AND B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
If LenB(LFParty) <> 0 Then MYSQL = MYSQL & " AND B.PARTY ='" & LFParty & "'"
If LenB(LFSauda) <> 0 Then MYSQL = MYSQL & " AND B.SAUDA ='" & LFSauda & "'"
MYSQL = MYSQL & " ORDER BY A.NAME"
LFBrokerRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
If Not LFBrokerRec.EOF Then
    DComboBroker.Enabled = True
    Set DComboBroker.RowSource = LFBrokerRec
    DComboBroker.BoundColumn = "AC_CODE"
    DComboBroker.ListField = "NAME"
Else
    DComboBroker.Enabled = False
End If
Call DATA_GRID_REFRESH
End Sub
Private Sub DComboBroker_Validate(Cancel As Boolean)
If LenB(DComboBroker.BoundText) <> 0 Then
    LFBroker = DComboBroker.BoundText
Else
    LFBroker = vbNullString
End If
Call DATA_GRID_REFRESH
End Sub
Private Sub DataGrid1_DblClick()
Dim LPConNo As Long:         Dim LPSauda As String:      Dim LPConType As String:        Dim TREC As ADODB.Recordset
If GCINNo = "2000" Then
    DataGrid1.Col = 4:              LPSauda = DataGrid1.text
    DataGrid1.Col = 10:             LPConNo = DataGrid1.text
    DataGrid1.Col = 2:              LPConType = DataGrid1.text
    'DataGrid1.Col = 9:              LPDataImport = Trim(DataGrid1.text)
Else
    DataGrid1.Col = 2:              LPSauda = DataGrid1.text
    DataGrid1.Col = 10:             LPConNo = DataGrid1.text
    DataGrid1.Col = 3:              LPConType = DataGrid1.text
End If
Call Mod_Rec
MYSQL = "SELECT CONSNO,CONNO, QTY,RATE,PARTY,CONTYPE,SAUDA,ITEMCODE,EXCODE,CALVAL,CONCODE,BROKAMT,SCGROUP,FILETYPE,BROKFLAG FROM CTR_D "
MYSQL = MYSQL & " WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'AND SAUDA='" & LPSauda & "'AND CONNO=" & LPConNo & ""
Set TREC = Nothing
Set TREC = New ADODB.Recordset
TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
If Not TREC.EOF Then
    Do While Not TREC.EOF
        TxtConNo.text = TREC!CONNO
        ChkAppBrok.Value = IIf(TREC!BROKFLAG = "Y", 1, 0)
        If TREC!CONTYPE = LPConType Then
            TxtPtyCode.text = TREC!PARTY
            DataCombo2.BoundText = TREC!PARTY
            TxtRate.text = Format(TREC!Rate, "0.0000")
        Else
            DataCombo4.BoundText = TREC!PARTY
            TxtConRate.text = Format(TREC!Rate, "0.0000")
        End If
        TxtBrokerCode.text = TREC!CONCODE
        DataCombo4.BoundText = TREC!CONCODE
        DComboTSauda.BoundText = TREC!Sauda
        If LPConType = "B" Then
            TxtContype.text = "Buy"
        Else
            TxtContype.text = "Sel"
        End If
        TxtQty.text = TREC!QTY
        TxtFileType.text = TREC!FILETYPE
        If LenB(LBillExCodes) < 1 Then
            LBillExCodes = "'" & TREC!EXCODE & "'"
        Else
            If InStr(LBillExCodes, TREC!EXCODE) < 1 Then LBillExCodes = LBillExCodes & ",'" & TREC!EXCODE & "'"
        End If
        If LenB(LBillParties) < 1 Then
            If TREC!PARTY <> TREC!CONCODE Then
                LBillParties = "'" & TREC!PARTY & "','" & TREC!CONCODE & "'"
            Else
                LBillParties = "'" & TREC!PARTY & "'"
            End If
        Else
            If InStr(LBillParties, "'" & TREC!PARTY & "'") < 1 Then LBillParties = LBillParties & ",'" & TREC!PARTY & "'"
            If InStr(LBillParties, "'" & TREC!CONCODE & "") < 1 Then LBillParties = LBillParties & ",'" & TREC!CONCODE & "'"
        End If
        If LenB(LBillItems) < 1 Then
            LBillItems = "'" & TREC!ITEMCODE & "'"
        Else
            If InStr(LBillItems, TREC!ITEMCODE) < 1 Then LBillItems = LBillItems & "," & "'" & TREC!ITEMCODE & "'"
        End If
        If LenB(LBillSaudas) < 1 Then
            LBillSaudas = "'" & TREC!Sauda & "'"
        Else
            If InStr(LBillSaudas, TREC!Sauda) < 1 Then LBillSaudas = LBillSaudas & ",'" & TREC!Sauda & "'"
        End If
        TREC.MoveNext
    Loop
    Call Get_Value:
    If Val(TxtRefLot.text) <> 0 Then
        TxtLot = CStr(Val(TxtQty.text) / Val(TxtRefLot.text))
    End If
    
    CmdAdd.Enabled = True:                          CmdModify.Enabled = False
    CmdCancel.Enabled = True:                       DtpCondate.Enabled = False
    DComboExchange.Enabled = False:                 ChkShowContract.Enabled = False
    Frame2.Enabled = True:                          LFBPress = 2
    Label12.Caption = "Modifying Existing Trades"
    TxtConNo.SetFocus
End If
Set TREC = Nothing
End Sub
Private Sub DataGrid3_DblClick()
Dim LPConNo As Long:            Dim LPSauda As String:          Dim LPConType As String:        Dim TREC As ADODB.Recordset
    DataGrid3.Col = 2:          LPSauda = DataGrid3.text
    DataGrid3.Col = 7:          LPConNo = DataGrid3.text
    DataGrid3.Col = 3:          LPConType = DataGrid3.text
    Call Mod_Rec
    MYSQL = "SELECT CONSNO,CONNO, QTY,RATE,PARTY,CONTYPE,SAUDA,ITEMCODE,EXCODE,CONCODE,STATUS,ST_TIME FROM CTR_L WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'AND SAUDA='" & LPSauda & "'AND CONNO=" & LPConNo & ""
    Set TREC = Nothing:    Set TREC = New ADODB.Recordset
    TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TREC.EOF Then
        Do While Not TREC.EOF
            TxtConNo.text = TREC!CONNO:                         TxtPtyCode.text = TREC!PARTY
            DataCombo2.BoundText = TREC!PARTY:                  TxtRate.text = Format(TREC!Rate, "0.0000")
            TxtBrokerCode.text = TREC!CONCODE:                  DataCombo4.BoundText = TREC!CONCODE
            DComboTSauda.BoundText = TREC!Sauda
            If LPConType = "B" Then
                TxtContype.text = "Buy"
            Else
                TxtContype.text = "Sel"
            End If
            TxtQty.text = TREC!QTY
            TREC.MoveNext
        Loop
        Get_Value
        CmdAdd.Enabled = True:              CmdModify.Enabled = False
        CmdCancel.Enabled = True:           DtpCondate.Enabled = False
        DComboExchange.Enabled = False:     ChkShowContract.Enabled = False
        Frame2.Enabled = True:              LFBPress = 2
        Label12.Caption = "Modifying Existing Trades"
        TxtConNo.SetFocus
    End If
    Set TREC = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "DtpCondate" Or Me.ActiveControl.NAME = "vcDTP2" Then
            Sendkeys "{tab}"
        End If
    End If
    If CmdAdd.Enabled = False Then
       If KeyCode = 121 Then Frame4.Visible = True
    End If
End Sub
Private Sub Form_Load()
LPDataImport = 0:           LSParties = vbNullString:   LSSaudas = vbNullString:   LSExCodes = vbNullString:    LSPNames = vbNullString:    LSType = vbNullString
LSUserIds = vbNullString:   LFExCode = vbNullString:    LFParty = vbNullString:     LFSauda = vbNullString
LFBroker = vbNullString
ChkAppBrok.Value = 1
TxtFileType.text = "0"
If GShowLot = "1" Then
    TxtLot.Visible = True
Else
    TxtLot.Visible = False
End If
DtpCondate.Value = Date
Frame2.Visible = True
If GCINNo = "2000" Then
    DataCombo5.Left = 7300:         DComboTSauda.Left = 7300
    Label7.Left = 5280:             Label8.Left = 6080
    Label6.Left = 7300:             TxtContype.Left = 5280
    TxtQty.Left = 6080:             TxtContype.TabIndex = 11
    TxtQty.TabIndex = 12:           DComboTSauda.TabIndex = 13
End If
Set ExRec = Nothing: Set ExRec = New ADODB.Recordset
MYSQL = "SELECT EXCODE,EXNAME,CONTRACTACC FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME"
ExRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
If Not ExRec.EOF Then
    Set DComboExchange.RowSource = ExRec
    DComboExchange.BoundColumn = "EXCODE"
    DComboExchange.ListField = "EXCODE"
End If
If ExRec.RecordCount = 1 Then
    ExRec.MoveFirst
    DComboExchange.BoundText = ExRec!EXCODE
End If
Set PartyRec = Nothing: Set PartyRec = New ADODB.Recordset
MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
PartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
If Not PartyRec.EOF Then
    Set DataCombo2.RowSource = PartyRec
    DataCombo2.BoundColumn = "AC_CODE"
    DataCombo2.ListField = "NAME"
    Set DataCombo4.RowSource = PartyRec
    DataCombo4.BoundColumn = "AC_CODE"
    DataCombo4.ListField = "NAME"
End If

End Sub
Private Sub DComboExchange_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboParty_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboSauda_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboBroker_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo4_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo2_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboTSauda_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo5_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CANCEL_REC
End Sub


Private Sub TxtPtyCode_Validate(Cancel As Boolean)
'Dim NRec As ADODB.Recordset
Dim LAcCode As String
If Frame2.Enabled = True Then
    If LenB(TxtPtyCode.text) = 0 Then
        DataCombo2.SetFocus
    Else
        LAcCode = Get_AccountDCode(TxtPtyCode.text)
        If LenB(LAcCode) > 1 Then
            DataCombo2.BoundText = LAcCode
            DComboTSauda.SetFocus
        Else
            DataCombo2.SetFocus
        End If
    End If
Else
    Frame1.Enabled = True
    Frame10.Enabled = True
    DtpCondate.Enabled = True
    DtpCondate.SetFocus
End If
End Sub
Private Sub TxtAdminPadd_Validate(Cancel As Boolean)
If GRegNo2 = EncryptNEW(TxtAdminPadd.text, 13) Then
    Frame2.Enabled = True
    CmdAdd.Enabled = True
    CmdModify.Enabled = True
Else
    Frame2.Enabled = False
    MsgBox "Invalid Password No Modificatiobn Allowed"
    Cancel = True
End If
End Sub
Private Sub TxtConNo_Validate(Cancel As Boolean)
    Dim NewRec As ADODB.Recordset
    TxtQty.Locked = False
    If LFBPress = 2 Then
        If LenB(TxtConNo.text) = 0 Then
            TxtQty.Locked = True
        Else
            TxtQty.Locked = False
            MYSQL = "SELECT A.PARTY,A.EXCODE,A.ITEMCODE,A.SAUDA,A.CONDATE,A.CONNO,A.CONSNO,A.CONTYPE,A.QTY,A.RATE,B.NAME,A.ROWNO1,A.ORDNO,A.CONTIME,BROKAMT,CONCODE,A.DATAIMPORT,FILETYPE  FROM CTR_D AS A, ACCOUNTD AS B WHERE A.COMPCODE =" & GCompCode & " "
            MYSQL = MYSQL & " AND A.COMPCODE=B.COMPCODE AND A.PARTY=B.AC_CODE AND A.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'  AND A.CONNO=" & Val(TxtConNo.text) & " AND A.PARTY <>A.CONCODE "
            If LenB(LFExCode) > 0 Then MYSQL = MYSQL & " AND A.EXCODE='" & LFExCode & "'"
            Set NewRec = Nothing
            Set NewRec = New ADODB.Recordset
            NewRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not NewRec.EOF Then
                TxtConNo.text = NewRec!CONNO:                TxtPtyCode.text = NewRec!PARTY
                DataCombo2.BoundText = NewRec!PARTY:         DComboTSauda.BoundText = NewRec!Sauda
                TxtQty.text = NewRec!QTY:                    TxtRate.text = Format(NewRec!Rate, "0.0000")
                Text7.text = NewRec!ROWNO1:                  TxtBrokerCode.text = NewRec!CONCODE:
                DataCombo4.BoundText = NewRec!CONCODE
                TxtContime.text = NewRec!CONTIME:            TxtConRate.text = Format(NewRec!BROKAMT, "0.0000")
                TxtOrdNo.text = NewRec!ORDNO
                TxtFileType.text = NewRec!FILETYPE
                If NewRec!CONTYPE = "B" Then
                    TxtContype.text = "Buy"
                Else
                    TxtContype.text = "Sel"
                End If
                If NewRec!DATAIMPORT = True Then
                    LPDataImport = 1
                Else
                    LPDataImport = 0
                End If
                TxtDataImport.text = LPDataImport
                
                If LenB(LBillExCodes) < 1 Then
                    LBillExCodes = "'" & NewRec!EXCODE & "'"
                Else
                    If InStr(LBillExCodes, NewRec!EXCODE) < 1 Then LBillExCodes = LBillExCodes & ",'" & NewRec!EXCODE & "'"
                End If
                If LenB(LBillParties) < 1 Then
                    If NewRec!PARTY <> NewRec!CONCODE Then
                        LBillParties = "'" & NewRec!PARTY & "','" & NewRec!CONCODE & "'"
                    Else
                        LBillParties = "'" & NewRec!PARTY & "'"
                    End If
                Else
                    If InStr(LBillParties, "'" & NewRec!PARTY & "'") < 1 Then LBillParties = LBillParties & ",'" & NewRec!PARTY & "'"
                    If InStr(LBillParties, "'" & NewRec!CONCODE & "") < 1 Then LBillParties = LBillParties & ",'" & NewRec!CONCODE & "'"
                End If
                If LenB(LBillItems) < 1 Then
                    LBillItems = "'" & NewRec!ITEMCODE & "'"
                Else
                    If InStr(LBillItems, NewRec!ITEMCODE) < 1 Then LBillItems = LBillItems & "," & "'" & NewRec!ITEMCODE & "'"
                End If
                If LenB(LBillSaudas) < 1 Then
                    LBillSaudas = "'" & NewRec!Sauda & "'"
                Else
                    If InStr(LBillSaudas, NewRec!Sauda) < 1 Then LBillSaudas = LBillSaudas & ",'" & NewRec!Sauda & "'"
                End If
            Else
                MsgBox " Invalid Trade No"
                Cancel = True
            End If
            Set NewRec = Nothing
        End If
    End If
End Sub
Private Sub TxtContype_KeyPress(KeyAscii As Integer)
If Val(KeyAscii) >= 48 And KeyAscii <= 122 Then
    If Val(KeyAscii) = 66 Or Val(KeyAscii) = 98 Or Val(KeyAscii) = 83 Or Val(KeyAscii) = 115 Then
    Else
        If TxtContype.text = "Buy" Then
            TxtContype.text = "Sel"
        Else
            TxtContype.text = "Buy"
        End If
    End If
End If
If KeyAscii = 32 Then
    If TxtContype.text = "Buy" Then
        TxtContype.text = "Sel"
    Else
        TxtContype.text = "Buy"
    End If
End If
End Sub

Private Sub TxtContype_Validate(Cancel As Boolean)
If TxtContype.text <> "Buy" Then
    If TxtContype.text <> "Sel" Then
        TxtContype.text = "Buy"
        Cancel = True
        TxtContype.SetFocus
    End If
End If
End Sub
Private Sub TxtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtLot_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtQty_Validate(Cancel As Boolean)
    TxtQty.text = Format(TxtQty.text, "0.00")
End Sub
Private Sub TxtRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtRate_Validate(Cancel As Boolean)
    TxtRate.text = Format(TxtRate.text, "0.0000")
    If LFBPress = 1 Then TxtConRate.text = Format(TxtRate.text, "0.0000")
    Call Get_Value
End Sub
Private Sub TxtBrokerCode_GotFocus()
TxtBrokerCode.SelStart = 0
TxtBrokerCode.SelLength = Len(TxtBrokerCode.text)
End Sub
Private Sub TxtBrokerCode_Validate(Cancel As Boolean)
Dim NRec As ADODB.Recordset
If LenB(TxtBrokerCode.text) = 0 Then
    DataCombo4.SetFocus
Else
    MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & TxtBrokerCode.text & "'"
    Set NRec = Nothing
    Set NRec = New ADODB.Recordset
    NRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not NRec.EOF Then
        DataCombo4.BoundText = NRec!AC_CODE
        If NRec!AC_CODE = TxtPtyCode.text Then
            MsgBox "Broker A/c can not be Same As Party A/c"
            TxtConRate.SetFocus
        End If
        DataCombo4.SetFocus
    Else
        DataCombo4.SetFocus
    End If
    Set NRec = Nothing
End If
End Sub
Private Sub TxtConRate_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtConRate_Validate(Cancel As Boolean)
    TxtConRate.text = Format(TxtConRate.text, "0.0000")
End Sub
Private Sub TxtQty_GotFocus()
    TxtQty.SelStart = 0
    TxtQty.SelLength = Len(TxtQty.text)
End Sub

Private Sub TxtLot_GotFocus()
    TxtLot.SelStart = 0
    TxtLot.SelLength = Len(TxtLot.text)
End Sub

Private Sub TxtConRate_GotFocus()
    TxtConRate.SelStart = 0
    TxtConRate.SelLength = Len(TxtConRate.text)
End Sub

Private Sub TxtRate_GotFocus()
    TxtRate.SelStart = 0
    TxtRate.SelLength = Len(TxtRate.text)
End Sub
Private Sub TxtPtyCode_GotFocus()
    TxtPtyCode.SelStart = 0
    TxtPtyCode.SelLength = Len(TxtPtyCode.text)
End Sub
Public Sub DATA_GRID_REFRESH()
Dim LShreeRec As ADODB.Recordset
    If GCINNo = "2000" Then
        MYSQL = "SELECT A.Party,B.Name,A.CONTYPE AS Type ,A.Qty,A.Sauda ,A.Rate,A.UserID,A.ConTime,A.ROWNO1 AS TradeNo,"
        MYSQL = MYSQL & " A.ConNo,A.CONCODE AS Code,C.NAME AS Broker,A.BROKAMT AS ConRate,A.DATAIMPORT,A.EXCODE FROM "
        MYSQL = MYSQL & " CTR_D A, ACCOUNTD B, ACCOUNTD AS C  WHERE A.COMPCODE=" & GCompCode & ""
        MYSQL = MYSQL & " AND A.COMPCODE =C.COMPCODE AND A.CONCODE =C.AC_CODE "
        MYSQL = MYSQL & " AND A.PARTY NOT IN (SELECT DISTINCT CONCODE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' "
        If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE ='" & LFExCode & "'"
        MYSQL = MYSQL & ")"
        MYSQL = MYSQL & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
        If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE ='" & LFExCode & "'"
        MYSQL = MYSQL & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "' ORDER BY ROWNO1 DESC "
        Set ContRec = Nothing
        Set ContRec = New ADODB.Recordset
        ContRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        Set DataGrid1.DataSource = ContRec
        DataGrid1.ReBind
        DataGrid1.Refresh
        
    Else
        '0  Party,         '1  Name,:               '2  Sauda,:               '3  CONTYPE AS Type ,   '4  Qty,
        '5  Rate,          '6  CONCODE AS Code,     '7 NAME AS Broker,        '8 BROKAMT AS ConRate, '9 RATE-A.BROKAMT AS DiffRate,
        '10 conno          '11  contime              12  rowno1                13'dataimport
        
        MYSQL = "SELECT A.Party,B.Name,A.Sauda,A.CONTYPE AS BS ,A.Qty, A.Rate,A.CONCODE AS Code,C.NAME AS Broker,"
        MYSQL = MYSQL & " A.BROKAMT AS ConRate,A.RATE-A.BROKAMT AS DiffRate,A.ConNo,A.ConTime,A.ROWNO1 AS TradeNo,A.DATAIMPORT,A.EXCODE "
        MYSQL = MYSQL & " FROM CTR_D A, ACCOUNTD B, ACCOUNTD AS C  WHERE A.COMPCODE=" & GCompCode & ""
        MYSQL = MYSQL & " AND A.COMPCODE =C.COMPCODE AND A.CONCODE =C.AC_CODE "
        MYSQL = MYSQL & " AND A.COMPCODE =B.COMPCODE AND A.PARTY=B.AC_CODE "
        If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE ='" & LFExCode & "'"
        If LenB(LSParties) <> 0 Then MYSQL = MYSQL & " AND A.Party  IN (" & LSParties & ")"
        If LenB(LSSaudas) <> 0 Then MYSQL = MYSQL & " AND A.SAUDA IN (" & LSSaudas & ")"
        If LenB(LSPNames) <> 0 Then MYSQL = MYSQL & " AND B.NAME IN (" & LSPNames & ")"
        If LenB(LSType) <> 0 Then MYSQL = MYSQL & " AND A.CONTYPE  IN (" & LSType & ")"
        If LenB(LSUserIds) <> 0 Then MYSQL = MYSQL & " AND A.USERID IN (" & LSUserIds & ")"
        If LenB(LFParty) <> 0 Then MYSQL = MYSQL & " AND A.Party  ='" & LFParty & "'"
        If LenB(LFSauda) <> 0 Then MYSQL = MYSQL & " AND A.SAUDA  ='" & LFSauda & "'"
        If LenB(LFBroker) <> 0 Then MYSQL = MYSQL & " AND A.CONCODE   ='" & LFBroker & "'"
        MYSQL = MYSQL & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'AND PERCONT='P' ORDER BY CONNO DESC "
        Set ContRec = Nothing:        Set ContRec = New ADODB.Recordset
        ContRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        Set DataGrid1.DataSource = ContRec
        DataGrid1.ReBind
        DataGrid1.Refresh
    End If
    Call Resize_Grid
    TxtDiffAmt.text = "0.00"
    MYSQL = "SELECT SUM(CASE CONTYPE WHEN 'B' THEN (A.Qty*A.Rate*A.CALVAL) WHEN 'S'THEN (A.Qty*A.Rate*A.CALVAL)*-1 END) AS DIFFAMT FROM CTR_D A "
    MYSQL = MYSQL & " WHERE A.COMPCODE=" & GCompCode & ""
    If LenB(LFBroker) <> 0 Then MYSQL = MYSQL & " AND CONCODE   ='" & LFBroker & "'"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND A.EXCODE  ='" & LFExCode & "'"
    MYSQL = MYSQL & " AND CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    Set LShreeRec = Nothing
    Set LShreeRec = New ADODB.Recordset
    LShreeRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not LShreeRec.EOF Then
        If Not IsNull(LShreeRec!diffaMt) Then TxtDiffAmt.text = Format(Val(LShreeRec!diffaMt), "0.00")
    End If
    Set LShreeRec = Nothing
'End If
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
    If Not ContRec.EOF Then Set DataGrid1.DataSource = ContRec: DataGrid1.ReBind: DataGrid1.Refresh
    Call Resize_Grid
    
Error1:    Exit Sub
End Sub
Private Sub CANCEL_REC()

Dim SREC As ADODB.Recordset:        Dim PREC As ADODB.Recordset
'Dim LBSaudas As String:                     Dim LBParties As String:            Dim LBItems As String
LSSaudas = vbNullString:                    LSParties = vbNullString:
LFParty = vbNullString:                     LFExCode = vbNullString:            LFSauda = vbNullString
LFBroker = vbNullString:                    TxtCalVal.text = vbNullString:      TxtValue.text = vbNullString:
LFExCode = vbNullString:                    Text15.text = vbNullString
TxtSettleRate.text = vbNullString:          Frame2.Enabled = False
TxtFileType.text = "0"
Label12.Caption = "Updateing Bills Please Wait"
GETMAIN.Toolbar1_Buttons(6).Enabled = False
    On Error GoTo ERR1
    CmdAdd.Enabled = True:                      CmdModify.Enabled = True
    CmdCancel.Enabled = False:                  DtpCondate.Enabled = True
    DComboExchange.BoundText = vbNullString:    DataCombo2.BoundText = vbNullString
    DComboTSauda.BoundText = vbNullString:        DataCombo4.BoundText = vbNullString
    DComboParty.BoundText = vbNullString:       DComboSauda.BoundText = vbNullString
    DComboBroker.BoundText = vbNullString:      TxtBrokerCode.text = vbNullString
    TxtConNo.text = vbNullString:               TxtPtyCode.text = vbNullString
    
    DComboExchange.Enabled = True
    If ExRec.RecordCount = 1 Then
        ExRec.MoveFirst
        DComboExchange.BoundText = ExRec!EXCODE
    End If
    ChkShowContract.Enabled = True
    TxtConNo.Locked = False
    Frame2.Enabled = False
    LFBPress = 0:
    
    If SaveCalled = True Then
        Frame1.Enabled = False:
        'Frame2.Enabled = False
        'Frame3.Enabled = False
        Call RATE_TEST(DtpCondate.Value)
        Call Shree_Posting(DateValue(DtpCondate.Value))
        
        CNNERR = True:                 Cnn.BeginTrans
        Call Update_Charges(LBillParties, LBillExCodes, LBillSaudas, LBillItems, DtpCondate.Value, DtpCondate.Value, True)
        GETMAIN.Label1.Caption = "Updating Brokerage Rate Itemwise Complete"
        Cnn.CommitTrans
        Cnn.BeginTrans
        If BILL_GENERATION(DtpCondate.Value, GFinEnd, LBillSaudas, LBillParties, LBillExCodes) Then
            Cnn.CommitTrans
            CNNERR = False
        End If
        Call Chk_Billing
    End If
    LBillParties = vbNullString:    LBillExCodes = vbNullString
    LBillSaudas = vbNullString:     LBillItems = vbNullString
    SaveCalled = False:             Frame1.Enabled = True
    Frame3.Enabled = True:          Frame10.Enabled = True
    Frame2.Enabled = True:          GETMAIN.Toolbar1_Buttons(6).Enabled = True
    DtpCondate.Enabled = True:      DtpCondate.SetFocus
    ChkAppBrok.Value = 1
    Label12.Caption = "Bills Updated Successfully "
    Exit Sub
ERR1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
        'Resume
        Cnn.RollbackTrans: CNNERR = False
        Frame1.Enabled = True
        
    End If
End Sub

Private Sub TxtLot_Validate(Cancel As Boolean)
If Val(TxtLot) > 0 Then TxtQty.text = CStr(Val(TxtLot.text) * Val(TxtRefLot.text))
End Sub
Private Sub DtpCondate_Validate(Cancel As Boolean)
    Dim NRec As ADODB.Recordset
    vcDTP2.MinDate = DtpCondate.Value
    vcDTP2.Value = DtpCondate.Value + 90
    If GRateSlab = 1 Then
        MYSQL = "SELECT COMPCODE FROM CTR_R WHERE COMPCODE  =" & GCompCode & "  AND CONDATE ='" & Format(DtpCondate.Value, "yyyy/mm/dd") & "'"
        Set NRec = Nothing:        Set NRec = New ADODB.Recordset
        NRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not NRec.EOF Then
            Label19.Visible = True:            TxtAdminPadd.Visible = True
            CmdModify.Enabled = False:         CmdAdd.Enabled = False
        Else
            Label19.Visible = False:           TxtAdminPadd.Visible = False
            CmdModify.Enabled = True:          CmdAdd.Enabled = True
        End If
        Set NRec = Nothing
    Else
        Label19.Visible = False:                TxtAdminPadd.Visible = False
    End If
    Set LFPartyRec = Nothing:    Set LFPartyRec = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT A.AC_CODE,A.NAME FROM ACCOUNTD AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
    MYSQL = MYSQL & " AND A.AC_CODE =B.PARTY AND B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
    MYSQL = MYSQL & " ORDER BY A.NAME"
    LFPartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not LFPartyRec.EOF Then
        DComboParty.Enabled = True
        Set DComboParty.RowSource = LFPartyRec
        DComboParty.BoundColumn = "AC_CODE"
        DComboParty.ListField = "NAME"
    Else
        DComboParty.Enabled = False
    End If
    Set LFSaudaRec = Nothing:    Set LFSaudaRec = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT A.SAUDACODE,A.SAUDANAME,A.MATURITY,A.ITEMCODE,A.REFLOT FROM SAUDAMAST AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
    MYSQL = MYSQL & " AND A.SAUDACODE =B.SAUDA AND B.CONDATE ='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    If LenB(LFExCode) <> 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
    MYSQL = MYSQL & " ORDER BY A.ITEMCODE,A.MATURITY"
    LFSaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not LFPartyRec.EOF Then
        DComboSauda.Enabled = True:                 Set DComboSauda.RowSource = LFSaudaRec
        DComboSauda.BoundColumn = "SAUDACODE":      DComboSauda.ListField = "SAUDANAME"
    Else
        DComboSauda.Enabled = False
    End If
    FillTradeSaudaCombo
End Sub
Private Sub vcDTP2_Validate(Cancel As Boolean)
Dim TREC As ADODB.Recordset:        Dim LSaudaCode As String:       Dim LFLAG As Boolean:       Dim LTExCode  As String
    MYSQL = "SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE ='" & DataCombo5.BoundText & "' AND MATURITY='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
    Set TREC = Nothing
    Set TREC = New ADODB.Recordset
    TREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If TREC.EOF Then
        MsgBox "Creating New Contract for " & vcDTP2.Value & " Prompt Date"
        LSaudaCode = DataCombo5.text & " PD " & vcDTP2.Value
        ItemRec.MoveFirst
        ItemRec.Find "ITEMCODE ='" & DataCombo5.BoundText & "'"
        If ItemRec.EOF Then
            MsgBox "Invalid Item"
            Exit Sub
        Else
            LTExCode = ItemRec!EXCHANGECODE
        End If
        Call PInsert_Saudamast(LSaudaCode, LSaudaCode, LItemCodeDBCombo, vcDTP2.Value, 1, 1, 0, "FUT", vbNullString, 0, LTExCode, 1)
        'MYSQL = "EXEC INSERT_SAUDAMAST " & GCompCode  & ",'" & LSaudaCode & "','" & LSaudaCode & "','" & LItemCodeDBCombo & "','" & Format(vcDTP2.Value, "yyyy/MM/dd") & "',1,'FUT','',0,'" & LTExCode & "',1,1"
        'Cnn.Execute MYSQL
        LFLAG = True
    Else
        LSaudaCode = TREC!SAUDACODE
    End If
    Set TREC = Nothing
    If LFLAG = True Then
        MYSQL = "SELECT S.SAUDACODE,S.SAUDANAME,S.TRADEABLELOT,I.LOT,EX.LOTWISE,EX.EXCODE,S.REFLOT FROM SAUDAMAST AS S,ITEMMAST AS I,EXMAST AS EX "
        MYSQL = MYSQL & " WHERE S.COMPCODE =" & GCompCode & " AND S.COMPCODE =I.COMPCODE AND EX.COMPCODE =S.COMPCODE AND EX.EXCODE =I.EXCHANGECODE "
        MYSQL = MYSQL & " AND EX.EXCODE ='" & DComboExchange.BoundText & "'"
        MYSQL = MYSQL & " AND S.ITEMCODE=I.ITEMCODE AND S.MATURITY>='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'ORDER BY S.ITEMCODE,S.INSTTYPE,S.MATURITY"
        Set SaudaRec = Nothing
        Set SaudaRec = New ADODB.Recordset
        SaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        If Not SaudaRec.EOF Then
            Set DComboTSauda.RowSource = SaudaRec
            DComboTSauda.BoundColumn = "SAUDACODE"
            DComboTSauda.ListField = "SAUDANAME"
        End If
    End If
    DComboTSauda.BoundText = LSaudaCode
End Sub
Private Sub Mod_Rec()
    If ContRec.RecordCount > 0 Then
        Frame2.Enabled = True:              CmdAdd.Enabled = False
        CmdModify.Enabled = False:          CmdCancel.Enabled = True
        DtpCondate.Enabled = False:         DComboExchange.Enabled = False
        ChkShowContract.Enabled = False:    TxtConNo.text = vbNullString
        TxtConNo.Locked = False:            TxtQty.Locked = False
        TxtQty.text = vbNullString:         TxtRate.text = vbNullString
        TxtConRate.text = vbNullString:
        LFBPress = 2
        Label12.Caption = "Modifying Existing Trades"
        TxtConNo.SetFocus
    Else
        MsgBox "No Records to Modify "
    End If
End Sub
Private Sub Get_Value()
    If ChkShowContract.Value = 1 Then
        AllSaudaRec.Filter = adFilterNone
        AllSaudaRec.MoveFirst
        AllSaudaRec.Find "SAUDACODE ='" & DComboTSauda.BoundText & "'"
        If AllSaudaRec.EOF Then
            MsgBox "Invalid Contract"
        Else
            TxtCalVal.text = Format(AllSaudaRec!LOT, "0.00")
            Text15.text = Format(Val(TxtCalVal.text) * Val(TxtQty.text), "0.00")
            TxtValue.text = Format(Val(Text15.text) * Val(TxtRate.text), "0.00")
            TxtRefLot.text = Format(AllSaudaRec!REFLOT, "0.00")
        End If
    Else
        SaudaRec.MoveFirst
        SaudaRec.Find "SAUDACODE ='" & DComboTSauda.BoundText & "'"
        If SaudaRec.EOF Then
            MsgBox "Invalid Contract"
            Sendkeys "%{DOWN}"
        Else
            If SaudaRec!EXCODE = "MCX" Or SaudaRec!EXCODE = "NSE" Then
                If SaudaRec!LOTWISE = "Y" Then
                    TxtCalVal.text = Format(SaudaRec!TRADEABLELOT, "0.00")
                Else
                    TxtCalVal.text = Format(SaudaRec!LOT, "0.00")
                End If
            Else
                TxtCalVal.text = Format(SaudaRec!LOT, "0.00")
            End If
            Text15.text = Format(Val(TxtCalVal.text) * Val(TxtQty.text), "0.00")
            TxtValue.text = Format(Val(Text15.text) * Val(TxtRate.text), "0.00")
            TxtRefLot.text = Format(SaudaRec!REFLOT, "0.00")
        End If
    End If
End Sub
Public Sub FillTradeSaudaCombo()
If ChkShowContract.Value = 1 Then
    MYSQL = "SELECT S.SAUDACODE,S.SAUDANAME,S.EX_SYMBOL,S.ITEMCODE,C.ITEMNAME,S.MATURITY,S.EXCODE,S.INSTTYPE,S.OPTTYPE,S.STRIKEPRICE,S.LOT,S.BROKLOT,S.REFLOT,EX.LOTWISE "
    MYSQL = MYSQL & " FROM SCRIPTMASTER AS S,EXMAST AS EX ,CONTRACTMASTER AS C WHERE EX.EXCODE =S.EXCODE AND EX.COMPCODE =" & GCompCode & ""
    MYSQL = MYSQL & " AND C.ITEMCODE=S.ITEMCODE AND C.EXCODE=S.EXCODE AND S.MATURITY>='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    If LenB(LFExCode) > 0 Then MYSQL = MYSQL & " AND EX.EXCODE ='" & LFExCode & "'"
    MYSQL = MYSQL & " ORDER BY S.EXCODE,S.ITEMCODE,S.INSTTYPE,S.MATURITY"
    Set AllSaudaRec = Nothing
    Set AllSaudaRec = New ADODB.Recordset
    AllSaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not AllSaudaRec.EOF Then
        Set DComboTSauda.RowSource = AllSaudaRec:            DComboTSauda.BoundColumn = "SAUDACODE"
        DComboTSauda.ListField = "SAUDANAME"
    End If
Else
    MYSQL = "SELECT S.SAUDACODE,S.SAUDANAME,S.TRADEABLELOT,I.LOT,EX.LOTWISE,EX.EXCODE,S.REFLOT FROM SAUDAMAST AS S,ITEMMAST AS I,EXMAST AS EX "
    MYSQL = MYSQL & " WHERE S.COMPCODE =" & GCompCode & " AND S.COMPCODE =I.COMPCODE AND EX.COMPCODE =S.COMPCODE AND EX.EXCODE =I.EXCHANGECODE "
    MYSQL = MYSQL & " AND S.ITEMCODE=I.ITEMCODE AND S.MATURITY>='" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'"
    If LenB(LFExCode) > 0 Then MYSQL = MYSQL & " AND EX.EXCODE ='" & LFExCode & "'"
    MYSQL = MYSQL & " ORDER BY S.ITEMCODE,S.INSTTYPE,S.MATURITY"
    Set SaudaRec = Nothing
    Set SaudaRec = New ADODB.Recordset
    SaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not SaudaRec.EOF Then
        Set DComboTSauda.RowSource = SaudaRec
        DComboTSauda.BoundColumn = "SAUDACODE"
        DComboTSauda.ListField = "SAUDANAME"
    End If
End If
End Sub

Private Sub Resize_Grid()

If GCINNo = "2000" Then
        DataGrid1.Columns(1).Width = 2500:              DataGrid1.Columns(2).Width = 800
        DataGrid1.Columns(3).Width = 900:               DataGrid1.Columns(4).Width = 3000
        DataGrid1.Columns(5).Width = 1000:              DataGrid1.Columns(6).Width = 1200
        DataGrid1.Columns(7).Width = 1200:              DataGrid1.Columns(8).Width = 1200
        
        DataGrid1.Columns(7).Alignment = dbgRight:      DataGrid1.Columns(8).Alignment = dbgRight
        DataGrid1.Columns(2).Alignment = dbgCenter:     DataGrid1.Columns(9).Alignment = dbgRight
        DataGrid1.Columns(10).Alignment = dbgRight:     DataGrid1.Columns(3).Alignment = dbgCenter
        DataGrid1.Columns(5).NumberFormat = "0.00":     DataGrid1.Columns(5).Alignment = dbgRight:
        DataGrid1.Columns(6).Alignment = dbgRight
    Else
        '0  Party,         '1  Name,:               '2  Sauda,:               '3  CONTYPE AS Type ,   '4  Qty,
        '5  Rate,          '6  CONCODE AS Code,     '7 NAME AS Broker,        '8 BROKAMT AS ConRate, '9 RATE-A.BROKAMT AS DiffRate,
        '10 conno          '11  contime              12  rowno1                13'dataimport
                
        DataGrid1.Columns(0).Width = 900
        DataGrid1.Columns(1).Width = 2500:              DataGrid1.Columns(2).Width = 2500
        DataGrid1.Columns(3).Width = 500:               DataGrid1.Columns(4).Width = 700
        DataGrid1.Columns(5).Width = 1300:              DataGrid1.Columns(6).Width = 800
        DataGrid1.Columns(7).Width = 1800:              DataGrid1.Columns(8).Width = 1300
        DataGrid1.Columns(9).Width = 1000:              DataGrid1.Columns(10).Width = 800
        DataGrid1.Columns(11).Width = 1300:             DataGrid1.Columns(12).Width = 800:
        DataGrid1.Columns(13).Visible = False:          DataGrid1.Columns(3).Alignment = dbgCenter
        DataGrid1.Columns(5).Alignment = dbgRight:      DataGrid1.Columns(6).Alignment = dbgCenter
        DataGrid1.Columns(7).Alignment = dbgLeft:       DataGrid1.Columns(8).Alignment = dbgRight
        DataGrid1.Columns(4).Alignment = dbgRight:      DataGrid1.Columns(9).Alignment = dbgRight
        DataGrid1.Columns(10).Alignment = dbgRight
        DataGrid1.Columns(5).NumberFormat = "0.0000":     DataGrid1.Columns(9).NumberFormat = "0.00"
        DataGrid1.Columns(8).NumberFormat = "0.0000":     DataGrid1.Columns(11).Alignment = dbgLeft
        
        If GSoftwareType = "X" Then
            DataGrid1.Columns(5).NumberFormat = "0.0000"
            DataGrid1.Columns(11).NumberFormat = "0.0000"
        End If
    End If
End Sub

