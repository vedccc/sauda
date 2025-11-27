VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form GETCOMP 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Company Setup"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13530
   Icon            =   "GETCOMP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   13530
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame14 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0FF&
      Height          =   1335
      Left            =   5520
      TabIndex        =   92
      Top             =   4320
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         TabIndex        =   93
         Top             =   120
         Width           =   4575
         Begin VB.TextBox TxtAdminPass 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   94
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label27 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Admin Password"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1200
            TabIndex        =   95
            Top             =   120
            Width           =   2295
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
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
      Height          =   855
      Left            =   120
      TabIndex        =   82
      Top             =   0
      Width           =   12735
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         TabIndex        =   83
         Top             =   0
         Width           =   12735
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Company Setup"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   0
            TabIndex        =   84
            Top             =   120
            Width           =   12495
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6690
      Left            =   3000
      TabIndex        =   53
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11800
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Company Details"
      TabPicture(0)   =   "GETCOMP.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Combo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "General Settings"
      TabPicture(1)   =   "GETCOMP.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame13"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame13 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   6015
         Left            =   -74880
         TabIndex        =   78
         Top             =   600
         Width           =   8895
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4095
            Left            =   0
            TabIndex        =   97
            Top             =   1920
            Width           =   8895
            Begin VB.CheckBox Check14 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Interest "
               CausesValidation=   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2880
               TabIndex        =   105
               Top             =   2160
               Width           =   1215
            End
            Begin VB.CheckBox CheckOrder 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Voucher Format New"
               CausesValidation=   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   240
               TabIndex        =   43
               Top             =   2160
               Width           =   2415
            End
            Begin VB.CheckBox Check13 
               BackColor       =   &H00FFFFFF&
               Caption         =   "L Rate"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5400
               TabIndex        =   46
               Top             =   2640
               Width           =   2295
            End
            Begin VB.CheckBox Check7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Data Backup at exit"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2880
               TabIndex        =   45
               Top             =   2640
               Width           =   2415
            End
            Begin VB.CheckBox Check5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Data Import at login"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   240
               TabIndex        =   44
               Top             =   2640
               Width           =   2535
            End
            Begin VB.CheckBox ChkOnlyBrok 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apply Only Brokerage"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5160
               TabIndex        =   102
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CheckBox Check6 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Amount/100"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   240
               TabIndex        =   101
               Top             =   1200
               Width           =   1695
            End
            Begin VB.CheckBox Check10 
               BackColor       =   &H00FFFFFF&
               Caption         =   "ConNo ExCode Wise"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5160
               TabIndex        =   42
               Top             =   1680
               Width           =   2415
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Invoice"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7560
               TabIndex        =   34
               Top             =   360
               Width           =   1335
            End
            Begin VB.CheckBox Check20 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Security Trasaction Tax"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2040
               TabIndex        =   39
               Top             =   1155
               Width           =   2655
            End
            Begin VB.CheckBox Check21 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Sub Brok"
               CausesValidation=   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   240
               TabIndex        =   40
               Top             =   1680
               Width           =   1695
            End
            Begin VB.CheckBox Check8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Settlementwise Invoice No."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2040
               TabIndex        =   41
               Top             =   1680
               Width           =   3135
            End
            Begin VB.Frame Frame17 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame17"
               Height          =   975
               Left            =   120
               TabIndex        =   85
               Top             =   3000
               Width           =   8655
               Begin VB.TextBox TxtGenQuery 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   4560
                  TabIndex        =   50
                  Top             =   600
                  Width           =   855
               End
               Begin VB.TextBox TxtCtr_Type 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   4560
                  TabIndex        =   48
                  Top             =   120
                  Width           =   855
               End
               Begin VB.ComboBox Combo4 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  ItemData        =   "GETCOMP.frx":047A
                  Left            =   6600
                  List            =   "GETCOMP.frx":0484
                  TabIndex        =   51
                  Text            =   "Combo4"
                  Top             =   600
                  Width           =   1575
               End
               Begin vcDateTimePicker.vcDTP vcDTP1 
                  Height          =   375
                  Left            =   6600
                  TabIndex        =   49
                  Top             =   120
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   39796.5153356481
               End
               Begin vcDateTimePicker.vcDTP vcDTP3 
                  Height          =   375
                  Left            =   1680
                  TabIndex        =   47
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   39796.5153356481
               End
               Begin VB.Label Label21 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Settlemt Lock Date"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   103
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.Label Label30 
                  BackColor       =   &H00C0C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "STM Date"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   5520
                  TabIndex        =   100
                  Top             =   180
                  Width           =   975
               End
               Begin VB.Label Label32 
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gen.Query"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Left            =   3480
                  TabIndex        =   88
                  Top             =   645
                  Width           =   1095
               End
               Begin VB.Label Label31 
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cont Entry"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   3480
                  TabIndex        =   87
                  Top             =   165
                  Width           =   1095
               End
               Begin VB.Label Label26 
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Stm Order"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   5520
                  TabIndex        =   86
                  Top             =   660
                  Width           =   975
               End
            End
            Begin VB.CheckBox Check24 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Margin"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7560
               TabIndex        =   38
               Top             =   720
               Width           =   1815
            End
            Begin VB.CheckBox Check23 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Standing Charges"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5160
               TabIndex        =   37
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox Check22 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Value wise Intraday"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5160
               TabIndex        =   33
               Top             =   360
               Width           =   2415
            End
            Begin VB.CheckBox Check19 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Minimum Brokerage"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2040
               TabIndex        =   36
               Top             =   735
               Width           =   2415
            End
            Begin VB.CheckBox Check18 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Stamp Duty"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   35
               Top             =   705
               Width           =   1695
            End
            Begin VB.CheckBox Check17 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Transaction Fees"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2040
               TabIndex        =   32
               Top             =   360
               Width           =   2175
            End
            Begin VB.CheckBox Check16 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Service Tax"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   240
               TabIndex        =   31
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Caption         =   "Brokerage && Tax Structure"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   0
               TabIndex        =   80
               Top             =   0
               Width           =   8895
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Width           =   8895
            Begin VB.CheckBox Check15 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Show Lot Entry"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2160
               TabIndex        =   106
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Show Qty in Decimal"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4560
               TabIndex        =   104
               Top             =   360
               Width           =   2415
            End
            Begin VB.CheckBox Check12 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Daily Bill"
               CausesValidation=   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7080
               TabIndex        =   99
               Top             =   0
               Width           =   1815
            End
            Begin VB.CheckBox ChkRate 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check Rate"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7080
               TabIndex        =   98
               Top             =   1080
               Width           =   1575
            End
            Begin VB.CheckBox ChkShowSTD 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Show Standing"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4560
               TabIndex        =   96
               Top             =   1080
               Width           =   1815
            End
            Begin VB.CheckBox Check40 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Post Margin"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2160
               TabIndex        =   30
               Top             =   1080
               Width           =   2175
            End
            Begin VB.CheckBox Check39 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apply Numeric Codes "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4560
               TabIndex        =   23
               Top             =   0
               Width           =   2415
            End
            Begin VB.CheckBox Check38 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apply Cash MTM"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   29
               Top             =   1080
               Width           =   2295
            End
            Begin VB.CheckBox Check37 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apply Options MTM"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2160
               TabIndex        =   28
               Top             =   360
               Width           =   2295
            End
            Begin VB.CheckBox Check9 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apply Contract No"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2160
               TabIndex        =   25
               Top             =   0
               Width           =   2175
            End
            Begin VB.CheckBox Check11 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Apply SpreadQty"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   2055
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Rounding Off "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   24
               Top             =   0
               Width           =   1575
            End
            Begin VB.CheckBox Check4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Show Lot"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   26
               Top             =   720
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   9255
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            TabIndex        =   75
            Top             =   4560
            Width           =   8775
            Begin VB.TextBox TxtBackup 
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
               Left            =   1440
               TabIndex        =   22
               Top             =   720
               Width           =   7455
            End
            Begin VB.TextBox TxtRptPath 
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
               Left            =   1440
               TabIndex        =   21
               Top             =   120
               Width           =   7455
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Backup Path"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1095
               Left            =   0
               TabIndex        =   77
               Top             =   735
               Width           =   1575
            End
            Begin VB.Label Label10 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Report Path"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   0
               TabIndex        =   76
               Top             =   195
               Width           =   1575
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   120
            TabIndex        =   68
            Top             =   2640
            Width           =   8775
            Begin VB.TextBox TxtSebiRegNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   6480
               MaxLength       =   50
               TabIndex        =   20
               ToolTipText     =   "Enter MPST No."
               Top             =   1440
               Width           =   2295
            End
            Begin VB.TextBox TxtGST 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   1440
               MaxLength       =   50
               TabIndex        =   19
               ToolTipText     =   "Enter MPST No."
               Top             =   1440
               Width           =   3375
            End
            Begin VB.TextBox TxtEmail 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1440
               MaxLength       =   50
               TabIndex        =   13
               Top             =   120
               Width           =   3375
            End
            Begin VB.TextBox TxtPANNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   6480
               MaxLength       =   50
               TabIndex        =   14
               ToolTipText     =   "Enter CST No."
               Top             =   120
               Width           =   2535
            End
            Begin VB.TextBox TxtITSrvNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   6480
               MaxLength       =   50
               TabIndex        =   18
               ToolTipText     =   "Enter MPST No."
               Top             =   1000
               Width           =   2295
            End
            Begin VB.TextBox TxtSrvRegNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   1440
               MaxLength       =   50
               TabIndex        =   15
               ToolTipText     =   "Enter MPST No."
               Top             =   560
               Width           =   3375
            End
            Begin VB.TextBox TxtTitle 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6480
               MaxLength       =   50
               TabIndex        =   16
               Top             =   560
               Width           =   2655
            End
            Begin VB.TextBox TxtCINNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   1440
               MaxLength       =   50
               TabIndex        =   17
               ToolTipText     =   "Enter MPST No."
               Top             =   1000
               Width           =   3375
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "SEBI Reg. No"
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
               Left            =   5040
               TabIndex        =   91
               Top             =   1485
               Width           =   1335
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "GSTIN No"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   89
               Top             =   1485
               Width           =   1335
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Title"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   74
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "SrvTax RegNo"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   73
               Top             =   570
               Width           =   1215
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "IT Srv No"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   72
               Top             =   1035
               Width           =   975
            End
            Begin VB.Label Regno1 
               BackStyle       =   0  'Transparent
               Caption         =   "CIN No"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   71
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PAN No"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   5040
               TabIndex        =   70
               Top             =   165
               Width           =   720
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Mail Id."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   69
               Top             =   180
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            TabIndex        =   63
            Top             =   1680
            Width           =   8775
            Begin VB.TextBox TxtMobile 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6480
               MaxLength       =   50
               TabIndex        =   12
               Top             =   600
               Width           =   2775
            End
            Begin VB.TextBox TxtPhoneR 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   6480
               MaxLength       =   50
               TabIndex        =   10
               ToolTipText     =   "Enter Fax No."
               Top             =   120
               Width           =   2775
            End
            Begin VB.TextBox TxtFax 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1440
               MaxLength       =   50
               TabIndex        =   11
               Top             =   600
               Width           =   3375
            End
            Begin VB.TextBox TxtPhoneO 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   1440
               MaxLength       =   50
               TabIndex        =   9
               ToolTipText     =   "Enter Phone No."
               Top             =   120
               Width           =   3375
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   5040
               TabIndex        =   67
               Top             =   660
               Width           =   615
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   5040
               TabIndex        =   66
               Top             =   180
               Width           =   600
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone(O) "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   0
               TabIndex        =   65
               Top             =   180
               Width           =   1005
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   0
               TabIndex        =   64
               Top             =   660
               Width           =   345
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            TabIndex        =   56
            Top             =   0
            Width           =   9135
            Begin VB.TextBox TxtState 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   6000
               MaxLength       =   20
               TabIndex        =   7
               ToolTipText     =   "Enter City"
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox TxtCompcode 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   120
               Width           =   1575
            End
            Begin VB.TextBox TxtCompName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   405
               Left            =   3120
               MaxLength       =   100
               TabIndex        =   1
               ToolTipText     =   "Enter Company Name"
               Top             =   120
               Width           =   5535
            End
            Begin VB.TextBox TxtCity 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   5
               ToolTipText     =   "Enter City"
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox TxtAdd 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   6000
               MaxLength       =   200
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               ToolTipText     =   "Enter Company Address"
               Top             =   675
               Width           =   3015
            End
            Begin VB.TextBox TxtPin 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   3360
               MaxLength       =   6
               TabIndex        =   6
               ToolTipText     =   "Enter City"
               Top             =   1200
               Width           =   1455
            End
            Begin vcDateTimePicker.vcDTP DTPicker1 
               Height          =   375
               Left            =   1440
               TabIndex        =   2
               Top             =   675
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   37680.5968981481
            End
            Begin vcDateTimePicker.vcDTP DTPicker2 
               Height          =   375
               Left            =   3360
               TabIndex        =   3
               Top             =   675
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   37680.5968981481
            End
            Begin VB.TextBox TxtID 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   7560
               MaxLength       =   50
               TabIndex        =   8
               ToolTipText     =   "Enter City"
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "State          "
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
               Height          =   255
               Index           =   2
               Left            =   5040
               TabIndex        =   90
               Top             =   1260
               Width           =   495
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "ID "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   7320
               TabIndex        =   81
               Top             =   1245
               Width           =   255
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Company "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   62
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "to"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3000
               TabIndex        =   61
               Top             =   765
               Width           =   255
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Fin.Year"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   60
               Top             =   765
               Width           =   1095
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "City             "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   59
               Top             =   1260
               Width           =   495
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pin "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   2760
               TabIndex        =   58
               Top             =   1245
               Width           =   360
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   57
               Top             =   765
               Width           =   855
            End
         End
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "GETCOMP.frx":04A0
         Left            =   1920
         List            =   "GETCOMP.frx":04AA
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   6180
         Width           =   1815
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "GETCOMP.frx":04C6
      Height          =   6780
      Left            =   120
      TabIndex        =   54
      Top             =   830
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11959
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   16761024
      ForeColor       =   0
      ListField       =   "Author"
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
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   495
      Left            =   11280
      TabIndex        =   111
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   495
      Left            =   11280
      TabIndex        =   110
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   495
      Left            =   11280
      TabIndex        =   109
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   11280
      TabIndex        =   108
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   11280
      TabIndex        =   107
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   6780
      Left            =   240
      Top             =   840
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   7140
      Left            =   3120
      Top             =   720
      Width           =   9645
   End
End
Attribute VB_Name = "GETCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte:              Dim old_comp As String:         Public Passwrd As String:
Dim Upd_Brok As Boolean
Dim Upd_Date As Boolean:              Dim MLRec As ADODB.Recordset:   Dim CompRec As ADODB.Recordset
Dim TRec As ADODB.Recordset:          Dim TRec2 As ADODB.Recordset:   Dim LCompCode  As Long:
Dim MYRS As ADODB.Recordset:          'Dim MCnn As ADODB.Connection:
Private shlShell As Shell32.Shell
Private shlFolder As Shell32.FOLDER:  Private Const BIF_RETURNONLYFSDIRS = &H1
Sub add_record()
    Fb_Press = 1
    DataList1.Locked = True
    Call Get_Selection(1)
    SSTab1.Enabled = True:  old_comp = vbNullString:  TxtCompName.SetFocus
    TxtCompName.text = vbNullString
    TxtBackup.text = App.Path
End Sub

Private Sub Check8_Click()
    If Check8.Value Then
        Check8.Caption = "Settlemnet Wise Invoice No"
    Else
        Check8.Caption = "Year Wise Invoice No"
    End If
End Sub

Private Sub CheckOrder_Click()
    Call voucherformat
End Sub
Private Sub voucherformat()
    If CheckOrder.Value Then
        CheckOrder.Caption = "Voucher Format New"
    Else
        CheckOrder.Caption = "Voucher Format Old"
    End If
End Sub
Private Sub Combo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataList1_Click()
    TxtCompName.text = DataList1.text
End Sub
Private Sub DataList1_DblClick()
    If DataList1.Locked Then
    Else
        Call Get_Selection(2)
        Fb_Press = 2
        Call COMPANY_ACCESS
    End If
End Sub
Private Sub DTPicker1_LostFocus()
    If Fb_Press = 1 Then DTPicker2.Value = DTPicker1.Value + 364
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then
        Frame14.Visible = True
        TxtAdminPass.SetFocus
    End If
    If KeyCode = 122 Then
        Upd_Brok = True
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Form_Load()
    Dim CnnString As String
    Check12.Value = 1
    GCompanyName = vbNullString:
    ServerString = MServer
    ServerString = Left$(MServer, (InStr(1, MServer, "database") + 8)) & "MASTER"
    CnnString = ServerString
    
    If MCnn.State = 0 Then ' 0=closed
        Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
        MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
        MCnn.Open
    End If
    Check2.Value = 1
    Call CANCEL_RECORD
    
    Set CompRec = Nothing: Set CompRec = New ADODB.Recordset
    mysql = "SELECT * FROM COMPANY ORDER BY COMPCODE"
    CompRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    If Not CompRec.EOF Then
        Set DataList1.RowSource = CompRec
        DataList1.ListField = "Name"
        DataList1.BoundColumn = "COMPCODE"
    Else
        Fb_Press = 1
        Call Get_Selection(1)
        old_comp = vbNullString: SSTab1.Enabled = True
    End If
End Sub
Private Sub Form_Paint()
    'Me.BackColor = GETMAIN.BackColor: SSTab1.BackColor = GETMAIN.BackColor
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CANCEL_RECORD
    GETMAIN.StatusBar1.Panels(1).text = vbNullString
    Unload Me
End Sub
Sub Save_Record()
    Dim MSet As Double
    Dim TRec As ADODB.Recordset
    On Error GoTo Error1
    If LenB(TxtCompName.text) = 0 Then MsgBox "Company Name required before save.", vbExclamation, "Name Missing": Exit Sub
    ''''    Duplicate Company Name Checking
    If old_comp <> TxtCompName.text Then
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT * FROM COMPANY WHERE NAME='" & TxtCompName.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Duplicate Company Name.", vbInformation, "Error"
            TxtCompName.text = vbNullString: TxtCompName.SetFocus: Set TRec = Nothing
            Exit Sub
        End If
    End If
    If Fb_Press = 1 Then
        
        'If Not CompRec.EOF And GHDNO <> "1713138570" Then
         '   GReportPath = CompRec!Rpt_Path
        'end If

        If Not CompRec.EOF Then
            GReportPath = CompRec!Rpt_Path
        End If
        If CompRec.EOF Then
            GCompCode = 1001
        Else
            CompRec.MoveLast
            LCompCode = Val(CompRec!CompCode) + Val(1)
        End If
        mysql = "INSERT INTO COMPANY (COMPCODE)"
        mysql = mysql & " VALUES (" & LCompCode & " "
        Cnn.Execute mysql
    End If
    
    mysql = "UPDATE COMPANY SET "
    mysql = mysql & " NAME ='" & TxtCompName.text & "'"
    mysql = mysql & " ,ADD1 ='" & TxtAdd.text & "'"
    mysql = mysql & " ,CITY ='" & TxtCity.text & "'"
    mysql = mysql & " ,PIN ='" & TxtPin.text & "'"
    mysql = mysql & " ,STATE ='" & TxtState.text & "'"
    mysql = mysql & " ,PHONEO = '" & TxtPhoneO.text & "'"
    mysql = mysql & " ,PHONER = '" & TxtPhoneR.text & "'"
    mysql = mysql & " ,MOBILE ='" & TxtMobile.text & "'"
    mysql = mysql & " ,FAX ='" & TxtFax.text & "'"
    mysql = mysql & " ,EMAIL = '" & TxtEmail.text & "'"
    mysql = mysql & " ,FINBEGIN = '" & Format(DTPicker1.Value, "yyyy/MM/dd") & "'"
    mysql = mysql & " ,FINEND  = '" & Format(DTPicker2.Value, "yyyy/MM/dd") & "'"
    mysql = mysql & " ,AUTOBROK = " & Check37.Value & ""
    mysql = mysql & " ,RPT_PATH = '" & Trim$(TxtRptPath.text) & "'"
    mysql = mysql & " ,PARTYCODETYPE ='" & Left$(Combo1.text, 1) & "'"
    mysql = mysql & " ,INVOICE  =" & Check1.Value & ""
    mysql = mysql & " ,BILLINGCYCLE ='" & Check12.Value & "'"
    mysql = mysql & " ,QTY_DECIMAL  ='" & Check3.Value & "'"
    'mysql = mysql & " ,CTRTYPE = '" & Left$(TxtCtr_Type.text, 1) & "'"
    mysql = mysql & " ,CTRTYPE = '" & TxtCtr_Type.text & "'"
    mysql = mysql & " ,REGNO1 = '" & TxtSrvRegNo & "'"
    mysql = mysql & " ,ROUNDOFF = '" & Check2.Value & "'"
    mysql = mysql & " ,AMTDIVIDE = '" & Check6.Value & "'"
    mysql = mysql & " ,APPSPREAD = '" & IIf(Check11.Value, "Y", "N") & "'"
    mysql = mysql & " ,SETINVNO  = '" & IIf(Check8.Value, "S", "D") & "'"
    mysql = mysql & " ,DPATH = '" & Trim$(TxtBackup.text) & "'"
    mysql = mysql & " ,PANNO = '" & TxtPANNo.text & "'"
    mysql = mysql & " ,APPCNOTE = " & Check9.Value & ""
    'mysql = mysql & " ,SYSLOCKDT = '" & Format(vcDTP3.Value, "YYYY/MM/DD") & "'"
    mysql = mysql & " ,CHECKRT  ='" & IIf(ChkRate.Value, "Y", "N") & "'"
    mysql = mysql & " ,ACORDER = '" & Left$(Combo4.text, 1) & "'"
    mysql = mysql & " ,TRANFEES ='" & IIf(Check17.Value, "Y", "N") & "'"
    mysql = mysql & " ,STAMPDUTY = '" & IIf(Check18.Value, "Y", "N") & "'"
    mysql = mysql & " ,VALUEWISE = '" & IIf(Check22.Value, "Y", "N") & "'"
    mysql = mysql & " ,STANDING  = '" & IIf(Check23.Value, "Y", "N") & "'"
    mysql = mysql & " ,MINBROKYN = '" & IIf(Check19.Value, "Y", "N") & "'"
    mysql = mysql & " ,STT = '" & IIf(Check20.Value, "Y", "N") & "'"
    mysql = mysql & " ,SUBBROK = '" & IIf(Check21.Value, "Y", "N") & "'"
    mysql = mysql & " ,OrderEntry = '" & IIf(CheckOrder.Value, "Y", "N") & "'"
    mysql = mysql & " ,SRVTAX  = '" & IIf(Check16.Value, "Y", "N") & "'"
    mysql = mysql & " ,ConNoType = " & IIf(Check10.Value, 1, 0)
    mysql = mysql & " ,UNIQCLIENTID ='" & TxtID.text & "'"
    mysql = mysql & " ,MARGIN = '" & IIf(Check24.Value, "Y", "N") & "'"
    mysql = mysql & " ,EQ = '" & IIf(Check38.Value, "1", "0") & "'"
    mysql = mysql & " ,STMDT   ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    mysql = mysql & " ,GENQUERY  = '" & Trim(TxtGenQuery.text) & "'"
    mysql = mysql & " ,SHOWSTD = '" & IIf(ChkShowSTD.Value, "Y", "N") & "'"
    mysql = mysql & " ,CINNO = '" & TxtCINNo.text & "'"
    mysql = mysql & " ,SEBIREGNO ='" & TxtSebiRegNo.text & "'"
    mysql = mysql & " ,GSTIN ='" & TxtGST.text & "'"
    mysql = mysql & " ,SHOWLOT= '" & IIf(Check4.Value, "Y", "N") & "'"
    mysql = mysql & " ,SHOWLOTENTRY= '" & IIf(Check15.Value, "Y", "N") & "'"
    mysql = mysql & " ,ONLYBROK =" & Val(ChkOnlyBrok.Value) & ""
    
    mysql = mysql & " ,FLAGDATAIMPORT='" & IIf(Check5.Value, "Y", "N") & "'"
    mysql = mysql & " ,FLAGDATABACKUP='" & IIf(Check7.Value, "Y", "N") & "'"
    mysql = mysql & " ,FlagLiveRate='" & IIf(Check13.Value, "Y", "N") & "'"
    
    
    
    mysql = mysql & " WHERE COMPCODE =" & LCompCode & ""
    Cnn.Execute mysql
    If Fb_Press = 1 Then
        mysql = "UPDATE SYSCOMP SET FINBEGIN='" & Format(DTPicker1.Value, "yyyy/MM/dd") & "',FINEND='" & Format(DTPicker2.Value, "yyyy/MM/dd") & "',RPT_PATH='" & EncryptNEW(TxtRptPath.text, 13) & "',D_PATH='" & EncryptNEW(TxtBackup.text, 13) & "'WHERE COMPCODE=" & LCompCode & " AND DATABASENAME='" & GDatabaseName & "'"
    Else
        mysql = "UPDATE SYSCOMP SET FINBEGIN='" & Format(DTPicker1.Value, "yyyy/MM/dd") & "',FINEND='" & Format(DTPicker2.Value, "yyyy/MM/dd") & "',RPT_PATH='" & EncryptNEW(TxtRptPath.text, 13) & "',D_PATH='" & EncryptNEW(TxtBackup.text, 13) & "' WHERE COMPCODE=" & LCompCode & " AND  DATABASENAME='" & GDatabaseName & "'" '>>> , SYSLOCKDT='" & Format(vcDTP3.Value, "YYYY/MM/DD") & "'
    End If
    MCnn.Execute mysql
    GCompanyName = TxtCompName.text
    If Upd_Brok = True Then
        mysql = "UPDATE PITBROK SET UPTOSTDT='" & Format(DTPicker2.Value, "yyyy/MM/dd") & "'  WHERE COMPCODE =" & LCompCode & " AND UPTOSTDT ='" & Format(GFinEnd, "yyyy/MM/dd") & "'"
        Cnn.Execute mysql
        mysql = "UPDATE PITSBROK SET UPTOSTDT='" & Format(DTPicker2.Value, "yyyy/MM/dd") & "' WHERE COMPCODE =" & LCompCode & " AND UPTOSTDT ='" & Format(GFinEnd, "yyyy/MM/dd") & "'"
        Cnn.Execute mysql
        mysql = "UPDATE PEXBROK SET UPTOSTDT='" & Format(DTPicker2.Value, "yyyy/MM/dd") & "'  WHERE COMPCODE =" & LCompCode & " AND UPTOSTDT ='" & Format(GFinEnd, "yyyy/MM/dd") & "'"
        Cnn.Execute mysql
        mysql = "UPDATE PEXSBROK SET UPTOSTDT='" & Format(DTPicker2.Value, "yyyy/MM/dd") & "' WHERE COMPCODE =" & LCompCode & " AND UPTOSTDT ='" & Format(GFinEnd, "yyyy/MM/dd") & "'"
        Cnn.Execute mysql
        mysql = "UPDATE EXTAX SET TODT ='" & Format(DTPicker2.Value, "yyyy/MM/dd") & "'       WHERE COMPCODE =" & LCompCode & " AND TODT ='" & Format(GFinEnd, "yyyy/MM/dd") & "'"
        Cnn.Execute mysql
        
        If Fb_Press = 1 Then
            mysql = "DELETE FROM SETTLE WHERE COMPCODE =" & LCompCode & ""
            Cnn.Execute mysql
            Set TRec = Nothing
            Set TRec = New ADODB.Recordset
            mysql = "SELECT DISTINCT EXCODE,EXNAME FROM EXMAST "
            TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            Do While Not TRec.EOF
                mysql = "INSERT INTO EXMAST (COMPCODE,EXCODE,EXNAME)"
                mysql = mysql & " VALUES (" & LCompCode & ",'" & TRec!excode & "','" & TRec!EXNAME & "')"
                Cnn.Execute mysql
                TRec.MoveNext
            Loop
        End If
        mysql = "SELECT *  FROM SETTLE WHERE COMPCODE=" & LCompCode & " AND SETDATE ='" & Format(DTPicker2.Value, "YYYY/MM/DD") & "'"
        Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If TRec.EOF Then
            mysql = "SELECT MAX(SETNO) AS MSET FROM SETTLE WHERE COMPCODE  = " & LCompCode & ""
            Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            If TRec2.EOF Or IsNull(TRec2!MSet) Then
                MSet = 1
            Else
                MSet = TRec2!MSet + 1
            End If
            mysql = "INSERT INTO SETTLE(COMPCODE,SETNO, SETDATE) VALUES(" & LCompCode & "," & MSet & ",'" & Format(DTPicker2.Value, "yyyy/MM/dd") & "')"
            Cnn.Execute mysql
        End If
    End If

    GFinBegin = DTPicker1.Value: GFinEnd = DTPicker2.Value
    MFinBeg = Right$(CStr(GFinBegin), 2): GETMAIN.StatusBar1.Panels(3).text = GFinBegin
    MFinEnd = Right$(CStr(GFinEnd), 2): GETMAIN.StatusBar1.Panels(4).text = GFinEnd
    GCompanyAdd1 = TxtAdd.text: GCCity = TxtCity.text
    GAcCodeType = Mid(Combo1.text, 1, 1): MNarr = True
    GSTDAc = 4: GShree = 1: GTrading = 0: GBrokrage = 3: GTrFeesAc = 2: GNBOTBank = 0
    GETMAIN.TRANS.Enabled = True: GETMAIN.mnuquery.Enabled = True: GETMAIN.report.Enabled = True: GETMAIN.master.Enabled = True: GETMAIN.utilities.Enabled = True
    Call LogIn
    Call CompanySelection(GCompCode)
    Call CANCEL_RECORD
    Set CompRec = Nothing: Set CompRec = New ADODB.Recordset
    mysql = "SELECT * FROM COMPANY ORDER BY COMPCODE"
    CompRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not CompRec.EOF Then
        Set DataList1.RowSource = CompRec
        DataList1.ListField = "Name"
        DataList1.BoundColumn = "COMPCODE"
    End If
    If LenB(GCompanyName) < 1 Then
        MsgBox "Software is shuting down.Please restart software.", vbInformation
        Exit Sub
    Else
        Exit Sub
    End If
Error1:
    MsgBox err.Description, vbCritical, err.HelpFile
End Sub
Sub CANCEL_RECORD()
    TxtCompcode.text = vbNullString:    TxtCompName.text = vbNullString:      TxtAdd.text = vbNullString:               TxtPin.text = vbNullString
    TxtBackup.text = vbNullString:      TxtPhoneR.text = vbNullString:        TxtPhoneO.text = vbNullString:            TxtCINNo.text = vbNullString
    TxtCity.text = vbNullString:        TxtEmail.text = vbNullString:         TxtFax.text = vbNullString:               TxtSrvRegNo.text = vbNullString
    TxtMobile.text = vbNullString:      TxtITSrvNo.text = vbNullString:       TxtAdminPass.text = vbNullString:         TxtID.text = vbNullString
    TxtCtr_Type.text = vbNullString:    TxtGenQuery.text = vbNullString:      TxtState.text = vbNullString:             TxtPANNo.text = vbNullString
    TxtSebiRegNo.text = vbNullString:   TxtGST.text = vbNullString:           TxtRptPath.text = vbNullString
    ChkShowSTD.Value = 0
    Check5.Value = 0
    Check7.Value = 0
    
    Fb_Press = 0: Combo1.Locked = False
    DTPicker1.MinDate = CDate("01/04/1990")
    DTPicker1.MaxDate = CDate("01/04/2990")

    DTPicker2.MinDate = CDate("01/04/1990")
    DTPicker2.MaxDate = CDate("01/04/2990")
    DataList1.Locked = False
    
    Frame3.Enabled = False
    Frame11.Enabled = False
    
    Call CompanySelection(GCompCode)
    Call Get_Selection(10)
    SSTab1.Tab = 0: SSTab1.Enabled = False
End Sub
Sub COMPANY_ACCESS()
    On Error GoTo Error1
    If Val(DataList1.BoundText) > 0 Then
        DataList1.Locked = True
        CompRec.MoveFirst
        CompRec.Find "COMPCODE=" & DataList1.BoundText & "", , adSearchForward
        If Not CompRec.EOF Then
            LCompCode = DataList1.BoundText
            TxtCompcode.text = CompRec!CompCode:                 TxtCompName.text = CompRec!NAME & vbNullString:  old_comp = TxtCompName.text
            TxtRptPath.Enabled = True:                           TxtBackup.Enabled = True
            TxtAdd.text = CompRec!ADD1 & vbNullString:           TxtCity.text = CompRec!City & vbNullString
            TxtPin.text = CompRec!Pin & vbNullString:            TxtState.text = CompRec!State & vbNullString
            TxtEmail.text = CompRec!Email & vbNullString
            TxtPhoneO.text = CompRec!PhoneO & vbNullString:      TxtPhoneR.text = CompRec!PhoneR & vbNullString
            TxtFax.text = CompRec!Fax & vbNullString:            TxtMobile.text = CompRec!Mobile & vbNullString
            TxtSrvRegNo.text = CompRec!Regno1 & vbNullString
            TxtSebiRegNo.text = CompRec!SEBIREGNO & vbNullString
            TxtID.text = IIf(IsNull(CompRec!UNIQCLIENTID), vbNullString, CompRec!UNIQCLIENTID)
            
            'TxtCINNo.text = CompRec!CINNO & vbNullString
            TxtCtr_Type.text = Trim(CompRec!CTRTYPE):                  TxtGenQuery = Trim(CompRec!GENQUERY)
            
            ChkOnlyBrok.Value = 0
            If (CompRec!ONLYBROK & "") = "1" Then
                ChkOnlyBrok.Value = CompRec!ONLYBROK & ""
            End If
            
            vcDTP1.Value = CompRec!STMDT
            TxtGST.text = Trim(IIf(IsNull(CompRec!GSTIN), vbNullString, CompRec!GSTIN))
            DTPicker1.Enabled = True:                            DTPicker2.Enabled = True:
            'vcDTP3.Enabled = True:
            TxtRptPath.Enabled = False
            TxtBackup.Enabled = False:
            Check6.Value = Val(CompRec!AMTDIVIDE)
            
            If CompRec!BILLINGCYCLE = True Then
                Check12.Value = 1
            Else
                Check12.Value = 0
            End If
            If CompRec!QTY_DECIMAL = True Then
                Check3.Value = 1
            Else
                Check3.Value = 0
            End If
            If CompRec!CHECKRT = "Y" Then
                ChkRate.Value = 1
            Else
                ChkRate.Value = 0
            End If
            
            mysql = "SELECT * FROM SYSCOMP WHERE DATABASENAME ='" & GDatabaseName & "' AND COMPCODE =" & LCompCode & ""
            Set MLRec = Nothing
            Set MLRec = New ADODB.Recordset
            MLRec.Open mysql, MCnn, adOpenStatic, adLockReadOnly
            If MLRec.EOF Then
                DTPicker1.Value = DateValue(Date)
                DTPicker2.Value = DateValue(Date)
                TxtRptPath.text = vbNullString
                TxtBackup.text = vbNullString
                'vcDTP3.Value = DateValue(Date)
            Else
                DTPicker1.Value = DateValue(MLRec!finbegin)
                DTPicker2.Value = DateValue(MLRec!finend)
                TxtRptPath.text = DecryptNEW(MLRec!Rpt_Path, 13)
                TxtBackup.text = DecryptNEW(MLRec!D_PATH, 13)
                'vcDTP3.Value = DateValue(MLRec!SYSLOCKDT)
            End If
            If Not IsNull(CompRec!EQ) Then
                If CompRec!EQ = 0 Then
                    Check38.Value = 0
                Else
                    Check38.Value = 1
                End If
            Else
                Check38.Value = 0
            End If
            DTPicker1.Enabled = False
            DTPicker2.Enabled = False
            'vcDTP3.Enabled = False
            TxtPANNo.text = CompRec!PANNO & ""
            If CompRec!PARTYCODETYPE = "N" Then
                Combo1.ListIndex = Val(0)
            Else
                Combo1.ListIndex = Val(1)
            End If
            If CompRec!SETINVNO = "S" Then
                Check8.Value = 1
                Check8.Caption = "Settlement Wise Invoice No"
            Else
                Check8.Value = 0
                Check8.Caption = "Year Wise Invoice No"
            End If
            If CompRec!APPSPREAD = "Y" Then
                Check11.Value = 1
            Else
                Check11.Value = 0
            End If
            
            If CompRec!APPCNOTE = "1" Then
                Check9.Value = 1
            Else
                Check9.Value = 0
            End If
            Check10.Value = CompRec!CONNOTYPE
            
            Check5.Value = 0
            If CompRec!FlagDataImport = "Y" Then
                Check5.Value = 1
            End If
                                    
            Check7.Value = 0
            If CompRec!FlagDataBackup = "Y" Then
                Check7.Value = 1
            End If
            
            Check13.Value = 0
            If CompRec!FlagLiveRate = "Y" Then
                Check13.Value = 1
            End If
                                    
            If Not IsNull(CompRec!autobrok) Then
                If CompRec!autobrok = True Then
                    Check37.Value = 1
                Else
                    Check37.Value = 0
                End If
            Else
                Check37.Value = 0
            End If
            Combo1.Locked = True
            If CompRec!showlot = "Y" Then
                Check4.Value = 1
            Else
                Check4.Value = 0
            End If
            If CompRec!showlotentry = "Y" Then
                Check15.Value = 1
            Else
                Check15.Value = 0
            End If
            Check16.Value = IIf(CompRec!SRVTAX = "Y", 1, 0)
            Check17.Value = IIf(CompRec!TRANFEES = "Y", 1, 0)
            Check18.Value = IIf(CompRec!STAMPDUTY = "Y", 1, 0)
            Check19.Value = IIf(CompRec!MINBROKYN = "Y", 1, 0)
            Check20.Value = IIf(CompRec!STT = "Y", 1, 0)
            Check21.Value = IIf(CompRec!SUBBROK = "Y", 1, 0)
            CheckOrder.Value = IIf(CompRec!OrderEntry = "Y", 1, 0)
            Call voucherformat
            Check22.Value = IIf(CompRec!VALUEWISE = "Y", 1, 0)
            Check23.Value = IIf(CompRec!standing = "Y", 1, 0)
            Check24.Value = IIf(CompRec!MARGIN = "Y", 1, 0)
            ChkShowSTD.Value = IIf(CompRec!SHOWSTD = "Y", 1, 0)
            Check2.Value = IIf(IsNull(CompRec!ROUNDOFF), 1, CompRec!ROUNDOFF)
            If IsNull(CompRec!ACORDER) Then
                Combo4.ListIndex = 0
            Else
                If CompRec!ACORDER = "C" Then
                    Combo4.ListIndex = 1
                Else
                    Combo4.ListIndex = 0
                End If
            End If
        End If
        SSTab1.Enabled = True
        SSTab1.Tab = 0
        TxtCompName.SetFocus
    Else
        MsgBox "Please Select Company", vbInformation
        Call Get_Selection(10)
        Call CANCEL_RECORD
        DataList1.SetFocus
    End If
    Exit Sub
Error1:
   MsgBox err.Description, vbCritical, err.HelpFile
End Sub

Private Sub TxtAdminPass_Validate(Cancel As Boolean)
If TxtAdminPass.text = VAdminPass Then
    TxtRptPath.Enabled = True
    TxtBackup.Enabled = True
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    'vcDTP3.Enabled = True
    Frame3.Enabled = True
    Frame11.Enabled = True
End If
TxtAdminPass.text = vbNullString
End Sub
Private Sub TxtAdminPass_LostFocus()
    Frame14.Visible = False
End Sub




