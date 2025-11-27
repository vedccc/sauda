VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form GETVCH 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11565
   Icon            =   "GETVCH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
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
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame2 
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   855
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   1695
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "Sauda"
            BeginProperty Font 
               Name            =   "Harrington"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Sauda"
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "DB_VOUCHER"
      Height          =   6495
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   11535
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   10935
         Begin VB.TextBox TXT_NARR 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   24
            Top             =   660
            Width           =   9015
         End
         Begin VB.TextBox clo_bal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "clo_bal"
            Top             =   255
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   1560
            TabIndex        =   22
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Closing Balance"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6600
            TabIndex        =   27
            Top             =   315
            Visible         =   0   'False
            Width           =   1455
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Narration"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   660
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Party Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Visible         =   0   'False
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   10935
         Begin VB.TextBox VOU_NO 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "VOU_NO"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Frame pr_frame 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   675
            Visible         =   0   'False
            Width           =   3135
            Begin VB.OptionButton pmt_opn 
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               Caption         =   "Payment"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton Rpt_opn 
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               Caption         =   "Receipt"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   2160
               TabIndex        =   11
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.ComboBox VOU_TYPE 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            ItemData        =   "GETVCH.frx":0442
            Left            =   1560
            List            =   "GETVCH.frx":044F
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   8
            Text            =   "Text11"
            Top             =   720
            Width           =   1455
         End
         Begin vcDateTimePicker.vcDTP DTPicker1 
            Height          =   315
            Left            =   4800
            TabIndex        =   14
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37680.7250462963
         End
         Begin vcDateTimePicker.vcDTP vcDTP4 
            Height          =   315
            Left            =   8520
            TabIndex        =   15
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37680.7250462963
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   3360
            TabIndex        =   20
            Top             =   720
            Width           =   930
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voucher Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Voucher Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   18
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voucher Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7095
            TabIndex        =   17
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   7095
            TabIndex        =   16
            Top             =   720
            Width           =   675
         End
      End
      Begin VB.Frame LineFrame 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   120
         Index           =   0
         Left            =   390
         TabIndex        =   6
         Top             =   1116
         Width           =   10455
      End
      Begin MSDataListLib.DataCombo ACCMBO 
         Bindings        =   "GETVCH.frx":0468
         Height          =   315
         Left            =   480
         TabIndex        =   0
         ToolTipText     =   "1:  ALT + Down Arrow key to open list.    2 :   Enter key to select.    3 :  F3  to create new account."
         Top             =   2280
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc ACCADO 
         Height          =   330
         Left            =   120
         Top             =   19440
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.TextBox clo_dcr 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "clo_dcr"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSDataGridLib.DataGrid VCHGRID 
         Bindings        =   "GETVCH.frx":047D
         Height          =   3255
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Press ~ 1. F5 to Repeat Narration,  2. Enter to Select."
         Top             =   2880
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   16777215
         ColumnHeaders   =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
         TabAction       =   1
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "AC_NAME"
            Caption         =   "Particulars"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "HH:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "DR_CR"
            Caption         =   "D/C"
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
            DataField       =   "AMOUNT"
            Caption         =   "    Amount"
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
         BeginProperty Column03 
            DataField       =   "NARRATION"
            Caption         =   "Narration"
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
            DataField       =   "CHQNO"
            Caption         =   "Chq. No."
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
            DataField       =   "CHQDT"
            Caption         =   "Chq. Date"
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
            DataField       =   "BANK"
            Caption         =   "Bank Name"
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
            DataField       =   "BRANCH"
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
         BeginProperty Column08 
            DataField       =   "CL_BAL"
            Caption         =   "Balance"
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
         BeginProperty Column09 
            DataField       =   "VOU_TYPE"
            Caption         =   "Vou. Type "
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
            DataField       =   "VOU_NO"
            Caption         =   "Vou. No"
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
            DataField       =   "VOU_DT"
            Caption         =   "Vou. Dt."
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
            DataField       =   "AC_CODE"
            Caption         =   "Ac. Code"
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
            DataField       =   "G_CODE"
            Caption         =   "G. Code"
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
            DataField       =   "VOU_ID"
            Caption         =   "VOU_ID"
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
            ScrollBars      =   3
            AllowSizing     =   0   'False
            BeginProperty Column00 
               DividerStyle    =   1
               ColumnWidth     =   3525.166
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               DividerStyle    =   1
               ColumnWidth     =   464.882
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               DividerStyle    =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               DividerStyle    =   1
               ColumnWidth     =   4754.835
            EndProperty
            BeginProperty Column04 
               DividerStyle    =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               DividerStyle    =   1
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column06 
               DividerStyle    =   1
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column07 
               DividerStyle    =   1
               ColumnWidth     =   2550.047
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               DividerStyle    =   1
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column09 
               DividerStyle    =   1
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column10 
               DividerStyle    =   1
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               DividerStyle    =   1
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               DividerStyle    =   1
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               DividerStyle    =   1
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView NARNLIST 
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   2040
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15790335
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Narration"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   5880
         Width           =   615
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   6780
      Left            =   75
      Top             =   1080
      Width           =   11865
   End
End
Attribute VB_Name = "GETVCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecAcc As ADODB.Recordset:              Dim TempVch As ADODB.Recordset
Public MYRS As ADODB.Recordset:             Public MYRS1 As ADODB.Recordset
Public MYRS_VCHAMT As ADODB.Recordset:      Public Fb_Press As Byte
Dim VouRec As ADODB.Recordset
Dim LASTNARR As String:                     Dim WORDAMT1 As String
Public F_Vor_Dt As String:                  Public F_Payrpt As String
Public F_Vou_No As String:                  Public F_VOU_NO_OLD As String
Public VOU_ID As Integer:                   Public G_CODE As Long:
Public Account_Code As String:              Public ACCOUNT_GCODE As Long:
Dim MAMT As Double:                         Dim TotAmt As Double:
Dim DR_TOTAL As Double:                     Dim CR_TOTAL As Double:
Dim Voucher_Date As Date
Dim ContraRec As ADODB.Recordset
Private Sub ACCMBO_DblClick(Area As Integer)
    Call ACCMBO_KeyPress(13)
End Sub
Private Sub ACCMBO_GotFocus()
    MTXT = VchGrid.text
    ACCMBO.Top = Val(VchGrid.Top) + Val(VchGrid.RowTop(VchGrid.Row))
    If VOU_TYPE.ListIndex = 0 Or VOU_TYPE.ListIndex = 1 Then
        MYSQL = "SELECT UPPER(NAME) AS AC_NAME, AC_CODE,GCODE,OP_BAL FROM ACCOUNTM WHERE COMPCODE=" & MC_CODE & " AND ACTIVE = 1 AND AC_CODE NOT IN('" & DataCombo1.BoundText & "') ORDER BY NAME"
    Else
        MYSQL = "SELECT UPPER(NAME) AS AC_NAME, AC_CODE,GCODE,OP_BAL FROM ACCOUNTM WHERE COMPCODE=" & MC_CODE & " AND ACTIVE = 1 ORDER BY NAME "
    End If
    Set ContraRec = Nothing
    Set ContraRec = New ADODB.Recordset
    ContraRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    Set ACCMBO.RowSource = ContraRec
    ACCMBO.BoundColumn = "AC_CODE"
    ACCMBO.ListField = "AC_NAME"
    ACCMBO.Refresh
    ACCMBO.text = MTXT
    If Frame2.Visible = True Then
        Text1.SetFocus
    Else
        Sendkeys "%{DOWN}"
    End If
End Sub
Private Sub ACCMBO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Or KeyCode = 9 Then
        VchGrid.SetFocus
        ACCMBO.Visible = False
    End If
End Sub
Private Sub ACCMBO_KeyPress(KeyAscii As Integer)
Dim ACNAME As String
    If KeyAscii = 13 Then
        If MAc_Name <> ACCMBO.text Then
            ContraRec.MoveFirst
            ContraRec.Find "AC_CODE='" & ACCMBO.BoundText & "'", , adSearchForward
            If Not ContraRec.EOF Then
                ACNAME = ContraRec!AC_NAME
                Account_Code = ContraRec!AC_CODE
                ACCOUNT_GCODE = ContraRec!GCODE
                'MAC_CL_BAL = Val(ContraRec!OP_BAL) + Val(ContraRec!CREDIT) - Val(ContraRec!DEBIT)
                TempVch!AC_NAME = ACNAME
                If Len(TempVch!DR_CR & "") = Val(0) Then        ''JUST IN CASE OF ADD NEW
                    Select Case VOU_TYPE.ListIndex
                        Case 0, 1, 2, 4
                            If Rpt_opn.Value = True Then
                                TempVch!DR_CR = "C"
                            Else
                                TempVch!DR_CR = "D"
                            End If
                        Case 3
                            TempVch!DR_CR = "D"
                    End Select
                End If
                If LenB(VchGrid.Columns(2).text) = 0 Then TempVch!AMOUNT = 0
                TempVch!chqdt = Format(F_Vou_Dt, "yyyy/MM/dd")
                TempVch!CL_BAL = MAC_CL_BAL
                TempVch!VOUTYPE = MVou_Type
                TempVch!VchNo = VOU_NO.text
                TempVch!VCHDT = F_Vou_Dt
                TempVch!AC_CODE = Account_Code & ""
                TempVch!G_CODE = ACCOUNT_GCODE
                TempVch.Update
            End If
            ContraRec.MoveFirst
            VchGrid.Col = 1
            VchGrid.SetFocus
            VchGrid.Refresh
            ACCMBO.Visible = False
        End If
    End If
End Sub
Private Sub Combo1_Click()
    If Mid(Combo1.text, 1, 1) = "S" Then
        DataCombo2.BoundText = Val(DataCombo1.BoundText)
        DataCombo2.Locked = True
    Else
        DataCombo2.Locked = False
    End If
End Sub
Private Sub Combo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo2_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub Command1_Click()
    VchGrid.Col = 0
    VchGrid.SetFocus
    Frame2.Visible = False
End Sub
Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    If Len(Trim(DataCombo1.text)) > Val(0) Then

        RecAcc.MoveFirst
        RecAcc.Find "AC_CODE='" & DataCombo1.BoundText & "'", , adSearchForward

        Call VCH_Number                         ''''SUB ROUTINE FOR NEXT VCH NUMBER
        MGC_VOU = RecAcc!GCODE
        If Val(RecAcc!CL_BAL & "") = Val(0) Then
            clo_bal.text = "0.00"
        ElseIf Val(RecAcc!CL_BAL) < Val(0) Then
            clo_bal.text = Format(Val(RecAcc!CL_BAL) * (-1), "0.00")
            clo_dcr.text = "Dr"
        ElseIf Val(RecAcc!CL_BAL) > Val(0) Then
            clo_bal.text = Format(RecAcc!CL_BAL, "0.00")
            clo_dcr.text = "Cr"
        End If
        If MVou_Type = "SP" Then    ''FOR STORE SUPPLIER
            If Fb_Press = 1 Then DataCombo2.BoundText = Val(DataCombo1.BoundText)
            Frame2.Visible = True
            Text1.SetFocus
            Exit Sub
        End If
    Else
        Cancel = True
    End If
End Sub
Private Sub DataCombo2_GotFocus()
    If DataCombo2.Locked = False Then Sendkeys "%{DOWN}"
End Sub
Private Sub DTPicker1_Validate(Cancel As Boolean)
    If SYSTEMLOCK(DateValue(DTPicker1)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
        Cancel = True
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Private Sub NARNLIST_LostFocus()
    NARNLIST.Visible = False
    If MVou_Type = "CV" Then
        VchGrid.Col = 8
    ElseIf MVou_Type = "BV" Then
        VchGrid.Col = 4
    End If
    VchGrid.SetFocus
End Sub
Private Sub DTPicker1_LostFocus()
    Call VCH_Number             ''''SUBROUTINE CALL FOR VOUCHER NUMBER
    Voucher_Date = DTPicker1.Value
End Sub
Private Sub Form_Paint()
    If GETMAIN.ActiveForm.NAME = Me.NAME Then
        GETMAIN.StatusBar1.Panels(1).text = "Voucher Details"
        If ACCOUNT_CHANGE Then
            Call Items
            ACCOUNT_CHANGE = False
        End If

        If MNarr = True Then
            NARNLIST.ListItems.clear

            VchGrid.Columns(3).Visible = True
            Set MYRS = Nothing
            Set Itmx = Nothing
        Else
            VchGrid.Columns(3).Visible = False
        End If

        If Fb_Press <> 0 Then Call Get_Selection(Fb_Press)

        Call SHOW_VCHTOTAL

        If MFORMAT1 = "Query on Account" Then
            MFORMAT1 = vbNullString
            Call Get_Selection(2) ''MODIFY
            Fb_Press = 2
            Call VOUCHER_ACCESS(F_Vou_No)
            DTPicker1.Enabled = True
            VOU_TYPE.Enabled = False
            pr_frame.Enabled = False
            VchGrid.Enabled = True
            Frame1.Enabled = True

            VOU_TYPE.SetFocus
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CLEAR_SCREEN
    FLAG_QRYACC = False
    Fb_Press = 0
    GETMAIN.StatusBar1.Panels(1).text = vbNullString
    Unload Me
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text11_Validate(Cancel As Boolean)
    If Len(Trim(Text11.text)) < 1 Then Cancel = True
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If Fb_Press = 1 Then
        If MGC_VOU = Val(10) Then
            GRNTYPE = "C"
        Else
            GRNTYPE = "R"
        End If
        Set MYRS = Nothing
        Set MYRS = New ADODB.Recordset
        MYRS.Open "SELECT GRNCODE FROM GRN WHERE COMPCODE=" & MC_CODE & " AND GRNCODE='" & GRNTYPE & Text2.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not MYRS.EOF Then
            Cancel = True
            MsgBox "Duplicate G.R.N. number.", vbExclamation, "Warning"
        End If
        Set MYRS = Nothing
    End If
End Sub
Private Sub Text6_GotFocus()
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.text)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text6_Validate(Cancel As Boolean)
    Text6.text = Format(Text6.text, "0.00")
End Sub
Private Sub Text8_GotFocus()
    Text8.SelLength = Len(Text8.text)
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    Text8.text = Format(Text8.text, "0.00")
End Sub

Private Sub vcDTP5_Validate(Cancel As Boolean)
    If Fb_Press = 1 Then vcDTP2.Value = vcDTP5.Value + Val(30)
End Sub

Private Sub VCHGRID_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        If Fb_Press = 1 Then
            If (UCase(TempVch!DR_CR) = "D") And (VchGrid.Col = 8 Or VchGrid.Col = 4 Or VchGrid.Col = 2) Then
                If Val(MAC_CL_BAL) <= 0 Then
                    TempVch!CL_BAL = -(((Val(MAC_CL_BAL)) * (-1)) + TempVch!AMOUNT)
                    TempVch!CL_BAL = -(TempVch!AMOUNT) + ((Val(MAC_CL_BAL)) * (-1))
                ElseIf Val(MAC_CL_BAL) > 0 Then
                    TempVch!CL_BAL = TempVch!AMOUNT + Val(MAC_CL_BAL)
                    TempVch!CL_BAL = Val(MAC_CL_BAL) - TempVch!AMOUNT
                End If

            ElseIf (UCase(TempVch!DR_CR) = "C") And (VchGrid.Col = 8 Or VchGrid.Col = 4 Or VchGrid.Col = 2) Then
                If Val(MAC_CL_BAL) <= 0 Then
                    TempVch!CL_BAL = -((Val(MAC_CL_BAL) * (-1)) - TempVch!AMOUNT)
                    TempVch!CL_BAL = -((Val(MAC_CL_BAL) * (-1)) - TempVch!AMOUNT)
                ElseIf Val(MAC_CL_BAL) > 0 Then
                    TempVch!CL_BAL = Val(MAC_CL_BAL) - TempVch!AMOUNT
                    TempVch!CL_BAL = Val(MAC_CL_BAL) + TempVch!AMOUNT
                End If

            End If
        End If
        TempVch.Update
    ElseIf ColIndex = 1 Then
        TempVch!DR_CR = UCase(TempVch!DR_CR)
        If Not (UCase(TempVch!DR_CR) = "C" Or UCase(TempVch!DR_CR) = "D") Then
            TempVch.CancelUpdate
        End If
        TempVch.Update
    ElseIf IsDate(TempVch!chqdt) Then
        If DateValue(TempVch!chqdt & "") < GFinBegin Or DateValue(TempVch!chqdt & "") > GFinEnd Then
            MsgBox "Cheque Date must be between FINANCIAL YEAR.", vbExclamation, "Error"
            VchGrid.Col = 4
            TempVch!chqdt = ""
            VchGrid.SetFocus
            Exit Sub
        End If
    End If
    Call SHOW_VCHTOTAL
End Sub
Private Sub VCHGRID_AfterColUpdate(ByVal ColIndex As Integer)   ''TO REPEAT NARRATION
    If Len(Trim(TempVch!NARRATION & "")) > Val(0) Then LASTNARR = TempVch!NARRATION
End Sub
Private Sub VCHGRID_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 Or KeyCode = 9) And VchGrid.Col <> Val(8) Then
        If (ACCOUNT_GCODE = Val(12) Or ACCOUNT_GCODE = Val(13) Or ACCOUNT_GCODE = Val(14)) And (Val(VchGrid.Columns(2).text) > 0 And VchGrid.Col = 2) Then
            MROW = VchGrid.Row

            VchGrid.Row = MROW
        End If

        If VchGrid.Col = 0 And Len(Trim(TempVch!AC_NAME & "")) >= 0 Then
            ACCMBO.Visible = True
            ACCMBO.SetFocus
        End If

    ElseIf Not (KeyCode = 9 And Shift = 1) And VchGrid.Col = 8 Then ''ONLY THROUGH SHIFT+TAB ONE CAN LEAVE THE GRID
        If KeyCode < Val(14) Then
            KeyCode = 0                                             ''DONE BCOS ON LAST COL OF GRID IF ONE HITS TAB FOCUS GOES ON ITEM SELECTION COMBO AND ON THE LOST FOCUS ON DATACOMBO1 SOCNUMBER CHANGES
            Call VchGrid.SetFocus
        End If

    ElseIf VchGrid.Col = 3 And Shift = 0 Then
        I = Len(VchGrid.text & "")
        If Not (KeyCode = 8 Or KeyCode = 13 Or KeyCode = 27 Or KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 46) Then
            If I = Val(100) Then
                MsgBox "Length Overflow.", vbExclamation, "Warning"
                KeyCode = 0
                VchGrid.SetFocus
            End If
        End If
    End If
End Sub
Private Sub VCHGRID_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 9 Or KeyCode = 13) And VchGrid.Col = 8 Then
        If Not TempVch.EOF Then
            TempVch.Update
            MVou_Id = TempVch!VOU_ID
            TempVch.MoveNext
            VchGrid.LeftCol = 0
            VchGrid.Col = 0
        End If
        If TempVch.EOF Then
            TempVch.AddNew
            TempVch!VOU_ID = MVou_Id + 1
        End If
        TempVch.Update
        VchGrid.ReBind
        VchGrid.Refresh
        VchGrid.Col = 0
    ElseIf KeyCode = 13 Then  ''right arrow key
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Form_Load()
    Call ClearFormFn(GETVCH)
    vcDTP1.MinDate = DateValue("01/01/1990")        ''THIS IS BECOS PROVISION FOR AWATIED BILLS FROM LAST YEARS
    vcDTP2.MinDate = DateValue("01/01/1990")
    vcDTP3.MinDate = DateValue("01/01/1990")
    vcDTP4.MinDate = DateValue("01/01/1990")
    vcDTP5.MinDate = DateValue("01/01/1990")
    Voucher_Date = DTPicker1.Value
    Frame2.Top = Val(1950)
    Frame2.Left = Val(720)
    Frame2.ZOrder
    Frame2.Visible = False

    If MNarr = True Then
        Label10.Visible = False
        TXT_NARR.Visible = False
    End If

    Call MakeRec
    Set VouRec = Nothing:    Set VouRec = New ADODB.Recordset
    MYSQL = "SELECT * FROM VOUCHER where COMPCODE=" & MC_CODE & " ORDER BY VOU_NO"
    VouRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    
    MYSQL = "SELECT NAME AS AC_NAME, AC_CODE, GCODE, OP_BAL FROM ACCOUNTM WHERE COMPCODE=" & MC_CODE & " AND ACTIVE = 1   ORDER BY NAME"
    Set ContraRec = Nothing:    Set ContraRec = New ADODB.Recordset
    ContraRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    
    Set VchGrid.DataSource = TempVch
    VchGrid.ReBind
    VchGrid.Refresh
    DTPicker1.Value = Date
    NARNLIST.ZOrder
    VOU_TYPE.ListIndex = 0
    Frame1.Enabled = False
    VchGrid.Columns(8).Locked = True
    VchGrid.Columns(9).Visible = False
    VchGrid.Columns(4).Visible = False
    VchGrid.Columns(5).Visible = False
    VchGrid.Columns(6).Visible = False
    VchGrid.Columns(7).Visible = False
    VchGrid.Columns(9).Visible = False
    VchGrid.Columns(10).Visible = False
    VchGrid.Columns(11).Visible = False
    VchGrid.Columns(12).Visible = False
    VchGrid.Columns(13).Visible = False
    If Not FLAG_QRYACC Then
        Call CLEAR_SCREEN
    End If
    Call Items
    If Not FLAG_QRYACC Then
        MVou_Type = "CV"
        MSEL_OPT = 1
    End If
    ACCOUNT_CHANGE = False
    GROUP_CHANGE = False
    Set MYRS = Nothing
End Sub
Sub SEARCH_RECORD()
    DTPicker1.Enabled = True
    If VouRec.EOF = True Then
        MsgBox "No Entry"
        Exit Sub
    End If
    Fb_Press = 1
    Call Get_Selection(1)
End Sub
Sub add_record()
    Call Get_Selection(1)
    Call CLEAR_SCREEN
    Fb_Press = 1
    Call MakeRec
    Set VchGrid.DataSource = TempVch
    VchGrid.ReBind
    VchGrid.Refresh
    Frame1.Enabled = True
    pr_frame.Enabled = True
    VOU_TYPE.SetFocus
End Sub
Sub Save_Record()
    On Error GoTo ERR1
    Dim MMonth As String
    Dim RES As Byte
    Dim Lvou_No As String
    If DTPicker1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: DTPicker1.SetFocus: Exit Sub
    If DTPicker1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: DTPicker1.SetFocus: Exit Sub
    
    If VOU_TYPE.ListIndex = Val(0) Or VOU_TYPE.ListIndex = Val(1) Then
        If pmt_opn.Value = True Then
            F_Payrpt = "PAYMENT"
        Else
            F_Payrpt = "RECEIPT"
        End If
    Else
        F_Payrpt = "0"
    End If
    TempVch.MoveFirst
    Do While Not TempVch.EOF
        If Len(Trim(TempVch!AC_CODE & "")) < 1 Then
            TempVch.Delete
        End If
        TempVch.MoveNext
    Loop
    TempVch.MoveFirst
    If Not TempVch.EOF Then                     ''''SOME VALIDATIONS
        If VOU_TYPE.ListIndex = Val(2) Then     ''''IN JV := BALANCE AMT MUST BE ZERO
            Call SHOW_VCHTOTAL

            If Val(MAMT) <> Val(0) Then
                MsgBox "Total mismatch. CREDIT AND DEBIT must be same.", vbExclamation, "Error"
                GETMAIN.ProgressBar1.Visible = False
                VchGrid.Col = 0:    VchGrid.SetFocus
                Exit Sub
            End If

        Else                                    ''''IF CASH OR BANK ACCOUNT NOT SELECTED
            If Val(DataCombo1.BoundText) = Val(0) Then
                MsgBox VOU_TYPE.text & " A/c. Not Selected. Select Account before save.", vbExclamation, "Error"
                GETMAIN.ProgressBar1.Visible = False
                DataCombo1.SetFocus
                Exit Sub
            End If
        End If

        If Fb_Press = 1 Then    '' NEW VOU_NO
            MYSQL = "SELECT VOU_NO FROM VOUCHER WHERE COMPCODE=" & MC_CODE & " AND VOU_NO='" & VOU_NO.text & "'"
            Set MYRS = Nothing
            Set MYRS = New ADODB.Recordset
            MYRS.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not MYRS.EOF Then
                If MsgBox("Voucher Number : " & VOU_NO.text & "   already exists." & vbCrLf & vbCrLf & "Do you want to assign a new number ?", vbQuestion + vbYesNo, "Warning") = vbYes Then
                    Do While Not MYRS.EOF
                        
                        VOU_NO.text = Get_VouNo(Mid(VOU_NO.text, 1, 4), Mid(VOU_NO.text, 5, 8))

                        MYSQL = "SELECT VOU_NO FROM VOUCHER WHERE COMPCODE=" & MC_CODE & " AND VOU_NO='" & VOU_NO.text & "'"
                        Set MYRS = Nothing
                        Set MYRS = New ADODB.Recordset
                        MYRS.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
                    Loop
                    F_Vou_No = VOU_NO.text
                    MsgBox "New Voucher Number  :  " & VOU_NO.text, vbInformation, "Message"
                    Beep 500, 250
                Else
                    GETMAIN.ProgressBar1.Visible = False
                    Exit Sub
                End If
            End If
        End If
        GETMAIN.ProgressBar1.Visible = True
        GETMAIN.ProgressBar1.Max = 100
        GETMAIN.ProgressBar1.Value = 20
        GETMAIN.Label1.Caption = vbNullString
        Cnn.BeginTrans
        CNNERR = True
        If Fb_Press = 2 Then Call Delete_Entry
        GETMAIN.ProgressBar1.Value = 30
        If Len(Trim(VOU_NO.text)) < Val(1) Then     ''IF VOUCHER NUMBER IS NOT THERE
            MsgBox "Please save it again", vbInformation, "Message"
            GETMAIN.ProgressBar1.Visible = False
            Exit Sub
        End If
        If Left(VOU_NO.text, 3) = "CSH" Then MVou_Type = "CV"
        If Left(VOU_NO.text, 3) = "BNK" Then MVou_Type = "BV"
        If Left(VOU_NO.text, 4) = "JRNL" Then MVou_Type = "JV"
        MYSQL = "INSERT INTO VOUCHER (COMPCODE,VOU_NO, VOU_TYPE, VOU_DT, VOU_PR, USER_NAME, USER_DATE, USER_TIME, USER_action) VALUES(" & MC_CODE & ",'" & VOU_NO.text & "','" & MVou_Type & "', '" & Format(DTPicker1.Value, "yyyy/MM/dd") & "','PAYMENT','" & USER_ID & "', '" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD')"
        Cnn.Execute MYSQL
        TempVch.MoveFirst
        GETMAIN.ProgressBar1.Value = 50
        Call ENT_SAVE
        Cnn.CommitTrans
        CNNERR = False
        NARNLIST.Visible = False
        GETMAIN.ProgressBar1.Value = 80
        Call Items
        Call CLEAR_SCREEN
        Call Get_Selection(4)
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Max
        GETMAIN.ProgressBar1.Visible = False
    End If
    Exit Sub
ERR1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
    GETMAIN.ProgressBar1.Visible = False
End Sub
Sub Delete_Entry()
    On Error GoTo Error1
    MYSQL = "SELECT AC_CODE, AMOUNT, DR_CR FROM VCHAMT WHERE COMPCODE=" & MC_CODE & " AND VOU_NO = '" & F_VOU_NO_OLD & "'"
    Set MYRS = Nothing
    Set MYRS = New ADODB.Recordset
    MYRS.Open MYSQL, Cnn, 2, adLockReadOnly
    While Not MYRS.EOF
        If MYRS!DR_CR = "D" Then
            If Not MYRS.EOF Then
                Cnn.Execute "UPDATE ACCOUNTM SET DEBIT = DEBIT - " & MYRS!AMOUNT & " WHERE COMPCODE = " & MC_CODE & " AND AC_CODE ='" & MYRS!AC_CODE & "'"
            End If
        ElseIf MYRS!DR_CR = "C" Then
            If Not MYRS.EOF Then
                Cnn.Execute "UPDATE ACCOUNTM SET CREDIT =CREDIT - " & MYRS!AMOUNT & " WHERE COMPCODE = " & MC_CODE & " AND AC_CODE ='" & MYRS!AC_CODE & "'"
            End If
        End If
        MYRS.MoveNext
    Wend
    Set MYRS = Nothing
    Cnn.Execute "DELETE  FROM VCHAMT WHERE COMPCODE=" & MC_CODE & " AND VOU_NO = '" & F_VOU_NO_OLD & "'"
    Cnn.Execute "DELETE  FROM VOUCHER WHERE COMPCODE=" & MC_CODE & " AND VOU_NO = '" & F_VOU_NO_OLD & "'"
    If VOU_TYPE.ListIndex = Val(3) Then
        If MGC_VOU = Val(10) Then
            GRNTYPE = "C"
        Else
            GRNTYPE = "R"
        End If
        Cnn.Execute "DELETE FROM GRN WHERE COMPCODE=" & MC_CODE & " AND GRNCODE='" & GRNTYPE & Text2.text & "'"
        Cnn.Execute "DELETE FROM GRN WHERE COMPCODE=" & MC_CODE & " AND GRNCODE='" & GRNTYPE & Text2.text & "'"
    End If
    If Fb_Press = 3 Then
        TempVch.MoveFirst
        Do While Not TempVch.EOF
            TempVch.Delete
            TempVch.Update
            TempVch.MoveNext
        Loop
        VchGrid.Refresh
        VchGrid.ReBind
        Call CLEAR_SCREEN
        Call Get_Selection(4)
    End If
    Exit Sub
Error1: If err.Number = 0 Then
            Cnn.RollbackTrans
        End If
End Sub
Sub RECORD_ACCESS()
Dim month As String
Dim MON As String
    Call Get_Selection(2)
    Call MakeRec
    VchGrid.Refresh
    VchGrid.ReBind
    
    Frame1.Enabled = True
    VOU_TYPE.Enabled = True
    pmt_opn.Value = True
    Label3.Caption = "Date"
    VOU_NO.Visible = False
    VOU_TYPE.SetFocus
    VOU_TYPE.ListIndex = 0
End Sub
Sub CANCEL_RECORD()
    Call CLEAR_SCREEN
    Call Get_Selection(5)
End Sub
Private Sub NARNLIST_DblClick()
    TempVch!NARRATION = NARNLIST.SelectedItem.text
    TempVch.Update
    NARNLIST.Height = 375
    NARNLIST.Visible = False
    VchGrid.Col = 3
    VchGrid.SetFocus
End Sub
Private Sub NARNLIST_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TempVch!NARRATION = NARNLIST.SelectedItem.text
        TempVch.Update
        NARNLIST.Height = 375
        NARNLIST.Visible = False
        If MVou_Type = "CV" Then
            VchGrid.Col = 8
        ElseIf MVou_Type = "BV" Then
            VchGrid.Col = 4
        End If
        VchGrid.SetFocus
    ElseIf KeyAscii = 27 Then
        NARNLIST.Height = 375
        NARNLIST.Visible = False
        If MVou_Type = "CV" Then
            VchGrid.Col = 3
        ElseIf MVou_Type = "BV" Then
            VchGrid.Col = 3
        End If
        VchGrid.SetFocus
    End If
End Sub
Private Sub SHOW_SCR()  '' IN CASE OF VOUCHER MODIFY & DELETE  CALLED FROM --- VOUNO_COM_LostFocus
    Call MakeRec
    TempVch.Delete

    If MVou_Type = "CV" Or MVou_Type = "BV" Or MVou_Type = "SP" Then
        MYSQL = "SELECT V.VOU_TYPE, V.VOU_NO, V.vouid, V.VOU_DT, V.DR_CR, V.AMOUNT, V.NARRATION, V.CHEQUE_NO, V.CHEQUE_DT, V.BANK_NAME, V.BRANCH, V.AC_CODE, A.GCODE, A.NAME, (A.OP_BAL + A.CREDIT - A.DEBIT) AS CL_BAL FROM VCHAMT AS V, ACCOUNTM AS A WHERE V.COMPCODE=" & MC_CODE & " AND V.COMPCODE =A.COMPCODE AND V.VOU_TYPE = '" & MVou_Type & "' AND V.VOU_NO = '" & F_Vou_No & "' AND A.AC_CODE = V.AC_CODE ORDER BY V.vouid"
    Else
        MYSQL = "SELECT V.VOU_TYPE, V.VOU_NO, V.vouid, V.VOU_DT, V.DR_CR, V.AMOUNT, V.NARRATION, V.CHEQUE_NO, V.CHEQUE_DT, V.BANK_NAME, V.BRANCH, V.AC_CODE, A.GCODE, A.NAME, (A.OP_BAL + A.CREDIT - A.DEBIT) AS CL_BAL FROM VCHAMT AS V, ACCOUNTM AS A WHERE V.COMPCODE=" & MC_CODE & " AND V.COMPCODE =A.COMPCODE AND V.VOU_TYPE = '" & MVou_Type & "' AND V.VOU_NO = '" & F_Vou_No & "' AND A.AC_CODE = V.AC_CODE ORDER BY V.vouid"
    End If
    Set MYRS1 = Nothing
    Set MYRS1 = New ADODB.Recordset
    MYRS1.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    While Not MYRS1.EOF
        If (MVou_Type = "CV" And MYRS1!GCODE = 10 And MYRS1!DR_CR = "D" And F_Payrpt = "RECEIPT") Or (MVou_Type = "CV" And MYRS1!GCODE = 10 And MYRS1!DR_CR = "C" And F_Payrpt = "PAYMENT") Or (MVou_Type = "BV" And MYRS1!GCODE = 11 And MYRS1!DR_CR = "D" And F_Payrpt = "RECEIPT") Or (MVou_Type = "BV" And MYRS1!GCODE = 11 And MYRS1!DR_CR = "C" And F_Payrpt = "PAYMENT") Or (MVou_Type = "SP" And MYRS1!DR_CR = "C") Then
            DataCombo1.BoundText = MYRS1!AC_CODE
            MGC_VOU = MYRS1!GCODE
            TXT_NARR.text = MYRS1!NARRATION & ""
            If (MYRS1!CL_BAL < 0) Then
                clo_bal.text = Format(Val(MYRS1!CL_BAL) * (-1), "0.00")
                clo_dcr.text = "Dr"
            ElseIf (MYRS1!CL_BAL > 0) Then
                clo_bal.text = Format(MYRS1!CL_BAL, "0.00")
                clo_dcr.text = "Cr"
            ElseIf (MYRS1!CL_BAL = 0) Then
                clo_bal.text = "0"
                clo_dcr.text = "Cr"
            End If
        Else
            If MVou_Type = "JV" Then
                TXT_NARR.text = MYRS1!NARRATION & ""
            End If
            TempVch.AddNew
            TempVch!VOUTYPE = MYRS1!VOU_TYPE
            TempVch!VchNo = MYRS1!VOU_NO
            TempVch!VCHDT = MYRS1!VOU_DT
            TempVch!DR_CR = MYRS1!DR_CR
            TempVch!AMOUNT = MYRS1!AMOUNT
            TempVch!NARRATION = MYRS1!NARRATION
            TempVch!CHQNO = MYRS1!CHEQUE_NO
            TempVch!chqdt = MYRS1!CHEQUE_DT
            TempVch!Bank = MYRS1!BANK_NAME
            TempVch!BRANCH = MYRS1!BRANCH
            TempVch!AC_CODE = MYRS1!AC_CODE
            TempVch!G_CODE = MYRS1!GCODE
            TempVch!AC_NAME = UCase(MYRS1!NAME)
            TempVch!CL_BAL = CL_BAL
            TempVch!VOU_ID = MYRS1!VOUID
            TempVch.Update
        End If
        MYRS1.MoveNext
    Wend

    TempVch.MoveFirst

    Set MYRS1 = Nothing
    Set VchGrid.DataSource = TempVch
    VchGrid.ReBind
    VchGrid.Refresh
End Sub
Private Sub ENT_SAVE()
Dim Mdr_Amt, Mcr_Amt As Double
Dim MYSQL1 As String
Dim MMonth As String
Dim CL_BAL As Long

    TempVch.MoveFirst

    If Left(VOU_NO.text, 3) = "CSH" Then MVou_Type = "CV"
    If Left(VOU_NO.text, 3) = "BNK" Then MVou_Type = "BV"
    If Left(VOU_NO.text, 4) = "JRNL" Then MVou_Type = "JV"
    While Not TempVch.EOF
        If Val(TempVch!AMOUNT) > Val(0) Then
            If (UCase(TempVch!DR_CR) = "D" Or UCase(TempVch!DR_CR) = "C") Then
                MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_TYPE, VOU_NO, VOU_DT, DR_CR, AMOUNT, NARRATION, Cheque_No, CHEQUE_DT, BANK_NAME, BRANCH, AC_CODE) "
                MYSQL = MYSQL & " VALUES( " & MC_CODE & ",'" & MVou_Type & "','" & VOU_NO.text & "','" & Format(DTPicker1.Value, "yyyy/MM/dd") & "','" & TempVch!DR_CR & "'," & Val(TempVch!AMOUNT & "") & ",'" & IIf(MNarr = True, (TempVch!NARRATION & ""), TXT_NARR.text) & "','" & (TempVch!CHQNO & "") & "','" & (TempVch!chqdt & "") & "','" & (TempVch!Bank & "") & "','" & (TempVch!BRANCH & "") & "','" & TempVch!AC_CODE & "" & "')"
            End If

            Cnn.Execute MYSQL
        End If
        TempVch.MoveNext
    Wend

    TempVch.MoveFirst
    GETMAIN.ProgressBar1.Value = 60

    Mdr_Amt = 0:    Mcr_Amt = 0

    While Not TempVch.EOF
        If (UCase(TempVch!DR_CR) = "D" Or UCase(TempVch!DR_CR) = "C") Then
            Account_Code = TempVch!AC_CODE

            ACCOUNT_GCODE = TempVch!G_CODE

            If UCase(TempVch!DR_CR) = "D" Then
                MYSQL = "UPDATE ACCOUNTM SET DEBIT = DEBIT +" & Val(TempVch!AMOUNT) & " WHERE COMPCODE = " & MC_CODE & " AND AC_CODE ='" & Account_Code & "'"
                Cnn.Execute MYSQL
                Mdr_Amt = Mdr_Amt + TempVch!AMOUNT

            ElseIf UCase(TempVch!DR_CR) = "C" Then
                MYSQL = "UPDATE ACCOUNTM SET CREDIT = CREDIT + " & Val(TempVch!AMOUNT) & " WHERE COMPCODE = " & MC_CODE & " AND AC_CODE ='" & Account_Code & "'"
                Cnn.Execute MYSQL
                Mcr_Amt = Mcr_Amt + TempVch!AMOUNT

            End If
        End If
        TempVch.MoveNext
    Wend

    TempVch.MoveFirst
    GETMAIN.ProgressBar1.Value = 70

    If MVou_Type = "SP" Then
        SQL = TempVch!NARRATION & " Bill No:" & Text11.text & ", Dt. " & Format(vcDTP4.Value, "yyyy/MM/dd")
    Else
        SQL = TempVch!NARRATION & ""
    End If

    '' CBLIST,PRTYLIST ACCOUNT MODIFICATION (USING GRID TOTAL DEBIT & TOTAL CREDIT AMOUNT) AND INSERTION INTO VCHAMT**********
    If MVou_Type = "CV" Or MVou_Type = "BV" Or MVou_Type = "SP" Then
        If Mdr_Amt > Mcr_Amt Then
            Mcr_Amt = Mdr_Amt - Mcr_Amt
            MYSQL = "UPDATE ACCOUNTM SET CREDIT = CREDIT + " & Mcr_Amt & " WHERE COMPCODE = " & MC_CODE & " AND AC_CODE ='" & Val(DataCombo1.BoundText) & "'"
            Cnn.Execute MYSQL

            MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_TYPE, VOU_NO, VOU_DT, DR_CR, AMOUNT, NARRATION, AC_CODE) VALUES(" & MC_CODE & ",'" & MVou_Type & "','" & F_Vou_No & "','" & Format(DTPicker1.Value, "yyyy/MM/dd") & "','C'," & Mcr_Amt & ",'" & SQL & "" & "'," & Val(DataCombo1.BoundText) & ")"
            Cnn.Execute MYSQL

        ElseIf Mcr_Amt >= Mdr_Amt Then
            Mdr_Amt = Mcr_Amt - Mdr_Amt

            MYSQL = "UPDATE ACCOUNTM SET DEBIT = DEBIT + " & Mdr_Amt & " WHERE COMPCODE = " & MC_CODE & " AND AC_CODE ='" & Val(DataCombo1.BoundText) & "'"
            Cnn.Execute MYSQL

            MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_TYPE, VOU_NO, VOU_DT, DR_CR, AMOUNT, NARRATION, AC_CODE) VALUES( " & MC_CODE & ",'" & MVou_Type & "','" & F_Vou_No & "','" & Format(DTPicker1.Value, "yyyy/MM/dd") & "','D'," & Mdr_Amt & ",'" & SQL & "" & "'," & Val(DataCombo1.BoundText) & ")"
            Cnn.Execute MYSQL

        End If
    End If
    GETMAIN.ProgressBar1.Value = 75
End Sub
Private Sub pmt_opn_LostFocus()
    If LenB(VOU_TYPE.text) <> 0 Then F_Payrpt = "PAYMENT"
    If Fb_Press = Val(1) Then Call DTPicker1_LostFocus
End Sub
Private Sub Rpt_opn_LostFocus()
    F_Payrpt = "RECEIPT"
    Call DTPicker1_LostFocus
End Sub
Private Sub VOU_NO_GotFocus()
    VOU_NO.text = F_Vou_No
End Sub
Private Sub VOU_TYPE_Click()
    Call VOUCHER_TYPE
    Fb_Press = Fb_Press
End Sub
Private Sub VOU_TYPE_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub VOU_TYPE_KeyUp(KeyCode As Integer, Shift As Integer)
    If LenB(VOU_TYPE.text) = 0 Then
        VOU_TYPE.SetFocus
        Exit Sub
    End If
    If (KeyCode = 38) Or (KeyCode = 40) Then
        Call VOUCHER_TYPE
    End If
End Sub
Sub Items()
    If ACCOUNT_CHANGE Then
        Set ContraRec = Nothing
        Set ContraRec = New ADODB.Recordset
        MYSQL = "SELECT NAME AS AC_NAME, AC_CODE, GCODE, OP_BAL, CREDIT, DEBIT FROM ACCOUNTM WHERE COMPCODE=" & MC_CODE & " AND ACTIVE=1  ORDER BY NAME"
        ContraRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        ACCOUNT_CHANGE = False
    End If
    If Not ContraRec.EOF Then
        Set ACCMBO.RowSource = ContraRec
        ACCMBO.ListField = "AC_NAME"
        ACCMBO.BoundColumn = "AC_CODE"
    End If
End Sub
Private Sub VOU_TYPE_LostFocus()
    Call VOUCHER_TYPE
    Fb_Press = Fb_Press
    If LenB(VOU_TYPE.text) = 0 And Fb_Press < 4 Then
        VOU_TYPE.SetFocus
    End If
End Sub
Private Sub VOU_TYPE_Validate(Cancel As Boolean)
On Error GoTo Error1
    If Len(Trim(VOU_TYPE.text)) < Val(1) Then VOU_TYPE.text = VOU_TYPE.list(0)

    F_Payrpt = "PAYMENT"

    Select Case VOU_TYPE.ListIndex
        Case 0                              ''CASH VOUCHER
            If Fb_Press = 1 Then
                If Not CASH_VOUCHER Then    ''FUNCTION CALL
                    VOU_TYPE.ListIndex = -1
                    VOU_TYPE.SetFocus
                Else
                    If RecAcc.EOF Then
                        MsgBox "Cash account not defined yet.", vbInformation
                        Cancel = True
                        Exit Sub
                    End If

                    VchGrid.Columns(4).Visible = False
                    VchGrid.Columns(5).Visible = False
                    VchGrid.Columns(6).Visible = False
                    VchGrid.Columns(7).Visible = False

                    If MNarr = False Then
                        Label10.Visible = True
                        TXT_NARR.Visible = True
                    End If
                End If
            End If
        Case 1                              ''BANK VOUCHER
            If Fb_Press = 1 Then
                If Not Bank_Voucher Then    ''FUNCTION CALL
                    VOU_TYPE.ListIndex = -1
                    VOU_TYPE.SetFocus
                Else
                    If RecAcc.EOF Then
                        MsgBox "Cash account not defined yet.", vbInformation
                        Cancel = True
                        Exit Sub
                    End If

                    VchGrid.Columns(2).AllowSizing = True
                    VchGrid.Columns(2).Width = 1110.008
                    VchGrid.Columns(2).AllowSizing = False
                    VchGrid.Columns(0).AllowSizing = False
                    VchGrid.Columns(4).Visible = True
                    VchGrid.Columns(5).Visible = True
                    VchGrid.Columns(6).Visible = True
                    VchGrid.Columns(7).Visible = True
                    If MNarr = False Then
                        Label10.Visible = True
                        TXT_NARR.Visible = True
                    End If
                End If
            End If
        Case 2                      ''JV
            If Fb_Press = 1 Then
                Label9.Visible = False
                clo_bal.Visible = False
                clo_dcr.Visible = False


                pr_frame.Visible = False
                If MNarr = False Then
                    Label10.Visible = True
                    TXT_NARR.Visible = True
                End If
            End If

            Label1.Visible = False
            VchGrid.Columns(4).Visible = False
            VchGrid.Columns(5).Visible = False
            VchGrid.Columns(6).Visible = False
            VchGrid.Columns(7).Visible = False
    
        Case 3                              ''STORE_PURCHASE
            If Fb_Press = 1 Then
                If Not STORE_PURCHASE Then  ''FUNCTION CALL
                    VOU_TYPE.ListIndex = -1
                    VOU_TYPE.SetFocus
                Else
                    VchGrid.Columns(4).Visible = False
                    VchGrid.Columns(5).Visible = False
                    VchGrid.Columns(6).Visible = False
                    VchGrid.Columns(7).Visible = False

                    If MNarr = False Then
                        Label10.Visible = True
                        TXT_NARR.Visible = True
                    End If
                End If
            End If

    End Select

    Call VCH_Number

    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    Cancel = True
End Sub
Function VOUCHER_ACCESS(Lvou_No As String) As Boolean
    Dim RES As Byte
    On Error GoTo Error1
    Label3.Caption = "Voucher Date"
    DTPicker1.Visible = True
    VOU_NO.Visible = True
    Label6.Caption = "Total"
    If Not VouRec.EOF Then
        VouRec.MoveFirst
    End If
    VouRec.Find "VOU_NO = '" & Lvou_No & "'", , adSearchForward
    If VouRec.EOF Then
        VOUCHER_ACCESS = False
    ElseIf Not VouRec.EOF Then
        VOUCHER_ACCESS = True
        VOU_TYPE.Enabled = False
        pr_frame.Enabled = False
        DTPicker1.Enabled = False
        vcDTP4.Enabled = False
        DataCombo1.Locked = True
        VOU_NO.text = Lvou_No
        F_Vou_No = Lvou_No
        F_VOU_NO_OLD = Lvou_No
        DTPicker1.Value = VouRec!VOU_DT
        F_Vou_Dt = DTPicker1.Value
        Select Case VouRec!VOU_TYPE
            Case "CV"
                VOU_TYPE.ListIndex = 0
                CASH_VOUCHER
                If MNarr = False Then
                    Label10.Visible = True
                    TXT_NARR.Visible = True
                End If
                If VouRec!VOU_PR = "PAYMENT" Then
                    F_Payrpt = "PAYMENT"
                    pmt_opn.Value = True
                    Rpt_opn.Value = False
                ElseIf VouRec!VOU_PR = "RECEIPT" Then
                    F_Payrpt = "RECEIPT"
                    pmt_opn.Value = False
                    Rpt_opn.Value = True
                End If
                pr_frame.Enabled = False
                VchGrid.Columns(4).Visible = False

            Case "BV"
                VOU_TYPE.ListIndex = 1

                Bank_Voucher

                If MNarr = False Then
                    Label10.Visible = True
                    TXT_NARR.Visible = True
                End If
                VchGrid.Columns(2).AllowSizing = True
                VchGrid.Columns(2).Width = 1110.008
                VchGrid.Columns(0).AllowSizing = False
                VchGrid.Columns(2).AllowSizing = False
                VchGrid.Columns(4).Visible = True
                VchGrid.Columns(5).Visible = True
                VchGrid.Columns(6).Visible = True
                VchGrid.Columns(7).Visible = True
                If VouRec!VOU_PR = "PAYMENT" Then
                    F_Payrpt = "PAYMENT"
                    pmt_opn.Value = True
                    Rpt_opn.Value = False
                ElseIf VouRec!VOU_PR = "RECEIPT" Then
                    F_Payrpt = "RECEIPT"
                    pmt_opn.Value = False
                    Rpt_opn.Value = True
                End If
                pr_frame.Enabled = False

            Case "JV"
                VOU_TYPE.ListIndex = 2
                MVou_Type = "JV"
                Label1.Visible = False
                VchGrid.Columns(4).Visible = False
        End Select
        Call SHOW_SCR
    End If
    Call SHOW_VCHTOTAL

    Exit Function
Error1: If err.Number <> 0 And err.Number <> 5 Then
            MsgBox err.Description, vbInformation
        End If
End Function
Sub VOUCHER_TYPE()
    Dim LVou_Type As String
    LVou_Type = MVou_Type
    Select Case VOU_TYPE.ListIndex
        Case 0, 1
            If VOU_TYPE.ListIndex = 0 Then
                MVou_Type = "CV"
                MSEL_OPT = 1
                VchGrid.Columns(8).Width = 1600
                Label1.Visible = True
                Label1.Caption = "Cash Account"
                Label4(15).Visible = False
                Label4(16).Visible = False
                Text11.Visible = False
                vcDTP4.Visible = False
                Frame2.Visible = False

            ElseIf VOU_TYPE.ListIndex = 1 Then
                MVou_Type = "BV"
                MSEL_OPT = 2
                VchGrid.Columns(8).Width = 1600
                Label1.Visible = True
                Label1.Caption = "Bank Account"
                Label4(15).Visible = False
                Label4(16).Visible = False
                Text11.Visible = False
                vcDTP4.Visible = False
                Frame2.Visible = False

            End If
            DataCombo1.Visible = True
        Case 2
            MVou_Type = "JV"
            MSEL_OPT = 3
            pr_frame.Visible = False
            DataCombo1.Visible = False
            Label1.Visible = False
            Label4(15).Visible = False
            Label4(16).Visible = False
            Text11.Visible = False
            vcDTP4.Visible = False
            Frame2.Visible = False
    End Select
    If LVou_Type <> MVou_Type Then TXT_NARR.text = vbNullString
End Sub
Function Bank_Voucher() As Boolean
On Error GoTo Error1
    MYSQL = "SELECT NAME AS AC_NAME, GCODE, AC_CODE, ACTIVE, (OP_BAL + CREDIT - DEBIT) AS CL_BAL FROM ACCOUNTM WHERE COMPCODE=" & MC_CODE & " AND GCODE =11 AND ACTIVE=1 ORDER BY NAME"
    Set RecAcc = Nothing
    Set RecAcc = New ADODB.Recordset
    RecAcc.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
    Set DataCombo1.RowSource = RecAcc
    DataCombo1.ListField = "AC_NAME"
    DataCombo1.BoundColumn = "AC_CODE"
    If Not RecAcc.EOF Then
        DataCombo1.BoundText = RecAcc!AC_CODE
    End If

    Bank_Voucher = True

    VchGrid.Columns(2).AllowSizing = True
    VchGrid.Columns(2).Width = 1110.008
    VchGrid.Columns(0).AllowSizing = False
    VchGrid.Columns(2).AllowSizing = False
    Exit Function
Error1:
    MsgBox err.Description, vbCritical, err.HelpFile
    Bank_Voucher = False
End Function
Function CASH_VOUCHER() As Boolean
    On Error GoTo Error1
    MYSQL = "SELECT NAME AS AC_NAME, GCODE, AC_CODE, ACTIVE, (OP_BAL + CREDIT - DEBIT) AS CL_BAL FROM ACCOUNTM WHERE COMPCODE=" & MC_CODE & " AND GCODE =10 AND ACTIVE=1 ORDER BY NAME"
    Set RecAcc = Nothing
    Set RecAcc = New ADODB.Recordset
    RecAcc.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly

    Set DataCombo1.RowSource = RecAcc
    DataCombo1.ListField = "AC_NAME"
    DataCombo1.BoundColumn = "AC_CODE"
    If Not RecAcc.EOF Then
        DataCombo1.BoundText = RecAcc!AC_CODE
    End If

    CASH_VOUCHER = True

    Exit Function
Error1:
    MsgBox err.Description, vbCritical, err.HelpFile
    CASH_VOUCHER = False
End Function
Sub list()
    VouLst.Show
End Sub
Sub CLEAR_SCREEN()
Dim OBJCONTROL As Control
    ACCOUNT_GCODE = 0: MAMT = 0:     GETMAIN.Label1.Caption = vbNullString: F_VOU_NO_OLD = vbNullString: F_Vou_No = vbNullString
    Fb_Press = 0:      DR_TOTAL = 0: CR_TOTAL = 0: VOU_ID = 0:    VOU_NO.text = vbNullString

    Call ClearFormFn(GETVCH)

    vcDTP1.MinDate = DateValue("01/01/1990")        ''THIS IS BECOS PROVISION FOR AWATIED BILLS FROM LAST YEARS
    vcDTP2.MinDate = DateValue("01/01/1990")
    vcDTP3.MinDate = DateValue("01/01/1990")
    vcDTP4.MinDate = DateValue("01/01/1990")
    vcDTP5.MinDate = DateValue("01/01/1990")

    DTPicker1.Value = Voucher_Date

    VOU_TYPE.ListIndex = 0: DataCombo1.text = vbNullString
    Frame2.Visible = False
    Text2.Enabled = True
    Text11.Enabled = True

    Call MakeRec

    Set VchGrid.DataSource = TempVch
    VchGrid.ReBind
    VchGrid.Refresh

    GETMAIN.ProgressBar1.Visible = False

    VOU_TYPE.Enabled = True
    DTPicker1.Enabled = True
    VOU_TYPE.Enabled = True
    VOU_NO.Enabled = True
    vcDTP1.Enabled = True
    DataCombo1.Locked = False
    DTPicker1.Enabled = True

    VOU_NO.Visible = True
    pr_frame.Enabled = True
    If MVou_Type = "BV" Then
        VchGrid.Columns(4).Visible = False
        VchGrid.Columns(5).Visible = False
        VchGrid.Columns(6).Visible = False
        VchGrid.Columns(7).Visible = False
    End If

    ACCMBO.Visible = False

    NARNLIST.Visible = False
    Label9.Visible = False
    clo_bal.Visible = False
    clo_dcr.Visible = False

    Label9.Visible = False
    clo_bal.Visible = False
    clo_dcr.Visible = False

    VchGrid.LeftCol = 0
    Frame1.Enabled = False

    If FLAG_QRYACC And MFORMAT1 = vbNullString Then
        MFormat = "Query on Account"
        FLAG_QRYACC = False
        Call Get_Selection(12)
        Unload Me
    End If
End Sub
Sub MakeRec()
    Set TempVch = Nothing
    Set TempVch = New ADODB.Recordset
    TempVch.Fields.Append "AC_NAME", adVarChar, 60, adFldIsNullable
    TempVch.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
    TempVch.Fields.Append "AMOUNT", adDouble, , adFldIsNullable
    TempVch.Fields.Append "NARRATION", adVarChar, 100, adFldIsNullable
    TempVch.Fields.Append "CHQNO", adVarChar, 10, adFldIsNullable
    TempVch.Fields.Append "CHQDT", adVarChar, 10, adFldIsNullable
    TempVch.Fields.Append "BANK", adVarChar, 30, adFldIsNullable
    TempVch.Fields.Append "BRANCH", adVarChar, 30, adFldIsNullable
    TempVch.Fields.Append "CL_BAL", adDouble, , adFldIsNullable
    TempVch.Fields.Append "VOUTYPE", adVarChar, 2, adFldIsNullable
    TempVch.Fields.Append "VCHNO", adVarChar, 15, adFldIsNullable
    TempVch.Fields.Append "VCHDT", adDate, , adFldIsNullable
    TempVch.Fields.Append "AC_CODE", adVarChar, 6, adFldIsNullable
    TempVch.Fields.Append "G_CODE", adDouble, , adFldIsNullable
    TempVch.Fields.Append "VOU_ID", adDouble, , adFldIsNullable
    TempVch.Open , , adOpenKeyset, adLockOptimistic
    TempVch.AddNew
    TempVch!VOU_ID = 1
    TempVch.Update
    Set VchGrid.DataSource = TempVch
    VchGrid.ReBind
    VchGrid.Refresh
End Sub
Sub RptRec()
    Set RptData = Nothing
    Set RptData = New ADODB.Recordset
    RptData.Fields.Append "ACC", adVarChar, 60, adFldIsNullable
    RptData.Fields.Append "NARRATION", adVarChar, 1, adFldIsNullable
    RptData.Fields.Append "AMT", adDouble, , adFldIsNullable
    RptData.Fields.Append "CHQ_NO", adVarChar, 10, adFldIsNullable
    RptData.Fields.Append "DT", adVarChar, 10, adFldIsNullable
    RptData.Fields.Append "VOU_NO", adVarChar, 15, adFldIsNullable
    RptData.Open , , adOpenKeyset, adLockOptimistic
End Sub
Sub SHOW_VCHTOTAL()
    Set MRec = Nothing
    Set MRec = New ADODB.Recordset
    Set MRec = TempVch.Clone
    If MRec.EOF Then Exit Sub
    MAMT = 0: MRec.MoveFirst
    Do While Not MRec.EOF
        If UCase(MRec!DR_CR & "") = "C" Then
            MAMT = MAMT + Val(MRec!AMOUNT & "")

        ElseIf UCase(MRec!DR_CR & "") = "D" Then
            MAMT = MAMT - Val(MRec!AMOUNT & "")

        End If
        MAMT = Round(MAMT, 2)

        MRec.MoveNext
    Loop
    Set MRec = Nothing

    If MAMT < Val(0) Then
        GETMAIN.Label1.Caption = "Voucher Total : " & Format(Val(MAMT) * Val(-1), "0.00") & " Dr"
    Else
        GETMAIN.Label1.Caption = "Voucher Total : " & Format(Val(MAMT), "0.00") & " Cr"
    End If
End Sub
Sub VCH_Number()
    Dim MVchno As String
    If (F_Payrpt <> "" And (MVou_Type = "CV" Or MVou_Type = "BV")) Or (MVou_Type = "JV") Or (MVou_Type = "SP") Or (MVou_Type = "PV") Then
        
        MFINBEGEND = FIN_YEAR(DTPicker1.Value)  ''FUNCTION TO FIND OUT FINANCIAL YEAR

        If (Fb_Press = 1) And (MVou_Type = "JV") Then
            F_Vou_No = Get_VouNo("JRNL", "0304")   ''GENERAL ROUTINE FOR NEXT VOUCHER NUMBER

        ElseIf (MVou_Type = "CV" Or MVou_Type = "BV" Or MVou_Type = "SP") And Fb_Press = Val(1) Then
            If MVou_Type = "CV" Then
                RecAcc.MoveFirst
                RecAcc.Find "AC_CODE='" & DataCombo1.BoundText & "'", , adSearchForward

                If Not RecAcc.EOF Then
                    If F_Payrpt = "PAYMENT" Then
                        MVchno = "CASH"
                    ElseIf F_Payrpt = "RECEIPT" Then
                        MVchno = "CSHR"
                    End If
                End If

                F_Vou_No = Get_VouNo(MVchno, "0304") ''GENERAL ROUTINE FOR NEXT VOUCHER NUMBER

            ElseIf MVou_Type = "BV" Then
                If Not RecAcc.EOF Then
                    If F_Payrpt = "PAYMENT" Then
                        MVchno = "BANK"
                    ElseIf F_Payrpt = "RECEIPT" Then
                        MVchno = "BANK"
                    End If
                End If
                F_Vou_No = Get_VouNo(MVchno, "0304")   ''GENERAL ROUTINE FOR NEXT VOUCHER NUMBER
            ElseIf MVou_Type = "SP" Then

                F_Vou_No = Get_VouNo("STPU", "0304")   ''GENERAL ROUTINE FOR NEXT VOUCHER NUMBER

            End If
            If Fb_Press = 1 Then F_Vou_Dt = DTPicker1.Value
        End If
        F_Vou_Dt = DTPicker1.Value
        VchGrid.Enabled = True

        VOU_NO.text = F_Vou_No
    End If
End Sub
