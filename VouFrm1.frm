VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRVIEWER.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form VouFrm1 
   Caption         =   "Voucher"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame13 
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
      TabIndex        =   43
      Top             =   0
      Width           =   13935
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   13935
         Begin VB.Label Label9 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Voucher Entry"
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
            Height          =   855
            Left            =   0
            TabIndex        =   45
            Top             =   120
            Width           =   14175
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8355
      Left            =   0
      TabIndex        =   21
      Top             =   720
      Width           =   13935
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   14175
         Begin VB.CommandButton CmdNewAccount 
            Caption         =   "Create New A/c"
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
            Left            =   12120
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton CmdImport 
            Caption         =   "Import Vouchers"
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
            Left            =   10200
            TabIndex        =   59
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox TxtVouNo 
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
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   112
            Width           =   3975
         End
         Begin VB.ComboBox ComboVouType 
            Appearance      =   0  'Flat
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
            Height          =   360
            ItemData        =   "VouFrm1.frx":0000
            Left            =   915
            List            =   "VouFrm1.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   705
            Width           =   1350
         End
         Begin VB.Frame pr_frame 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2355
            TabIndex        =   38
            Top             =   720
            Visible         =   0   'False
            Width           =   2535
            Begin VB.OptionButton pmt_opn 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "Payment"
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
               Left            =   0
               TabIndex        =   1
               Top             =   80
               Width           =   1335
            End
            Begin VB.OptionButton Rpt_opn 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "Receipt"
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
               Left            =   1320
               TabIndex        =   2
               Top             =   80
               Width           =   1215
            End
         End
         Begin VB.CommandButton CmdPrintVoucher 
            Caption         =   "Print Voucher"
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
            Left            =   8640
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin vcDateTimePicker.vcDTP DTPicker1 
            Height          =   375
            Left            =   5760
            TabIndex        =   16
            Top             =   105
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
            Value           =   37680.7250462963
         End
         Begin VB.Label LBLNETAMT 
            Alignment       =   1  'Right Justify
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
            Height          =   375
            Left            =   12720
            TabIndex        =   65
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LBLVOUAMT 
            Alignment       =   1  'Right Justify
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
            Height          =   375
            Left            =   11760
            TabIndex        =   64
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label LBLCLOSING 
            Alignment       =   1  'Right Justify
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
            Height          =   375
            Left            =   10320
            TabIndex        =   63
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label LBLNAME 
            BackColor       =   &H00FFFFC0&
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
            Left            =   9360
            TabIndex        =   62
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label LBLPARTY 
            BackColor       =   &H00FFFFC0&
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
            Left            =   8280
            TabIndex        =   61
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label8 
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
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   120
            TabIndex        =   41
            Top             =   765
            Width           =   495
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   5115
            TabIndex        =   40
            Top             =   165
            Width           =   465
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vou No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   120
            TabIndex        =   39
            Top             =   165
            Width           =   705
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   29
         Top             =   1230
         Width           =   14055
         Begin VB.TextBox txtnarration 
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
            Left            =   6840
            TabIndex        =   9
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox Text2 
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
            Left            =   4560
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   4
            Top             =   360
            Width           =   615
         End
         Begin MSDataListLib.DataCombo DComboNarr 
            Bindings        =   "VouFrm1.frx":0026
            Height          =   360
            Left            =   6840
            TabIndex        =   55
            ToolTipText     =   "1:  ALT + Down Arrow key to open list.    2 :   Enter key to select.    3 :  F3  to create new account."
            Top             =   1080
            Visible         =   0   'False
            Width           =   5535
            _ExtentX        =   9763
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
         Begin VB.TextBox txtbranch 
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
            Left            =   6240
            TabIndex        =   13
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtbank 
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
            Left            =   3480
            TabIndex        =   12
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtchqno 
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
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton command1 
            Caption         =   "Add"
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
            Left            =   12480
            TabIndex        =   14
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtamoount 
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
            Height          =   360
            Left            =   5280
            TabIndex        =   8
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Txtdrcr 
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
            Left            =   4560
            MaxLength       =   1
            TabIndex        =   7
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TXT_NARR 
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
            Left            =   5280
            MaxLength       =   1000
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.TextBox AmtTxt 
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
            Left            =   11715
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox TxtClBal 
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
            Left            =   9555
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   360
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DComboCashBank 
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   635
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "VouFrm1.frx":003B
            Height          =   360
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "1:  ALT + Down Arrow key to open list.    2 :   Enter key to select.    3 :  F3  to create new account."
            Top             =   1080
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   635
            _Version        =   393216
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
         Begin vcDateTimePicker.vcDTP chkdt 
            Height          =   375
            Left            =   1800
            TabIndex        =   11
            Top             =   1800
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
            Value           =   37680.7250462963
         End
         Begin VB.TextBox Txtnarrhidden 
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
            Left            =   8400
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txthiddenid 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtCashBankCode 
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Dr/Cr"
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
            TabIndex        =   58
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
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
            Left            =   6240
            TabIndex        =   54
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            Left            =   3480
            TabIndex        =   53
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Chq.Date"
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
            Left            =   1800
            TabIndex        =   52
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Chq.No."
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
            Left            =   120
            TabIndex        =   51
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Narration"
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
            Left            =   6960
            TabIndex        =   50
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   5280
            TabIndex        =   49
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Dr/Cr"
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
            TabIndex        =   48
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
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
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label LblAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
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
            Left            =   120
            TabIndex        =   35
            Top             =   75
            Width           =   2295
         End
         Begin VB.Label LblNarr 
            BackStyle       =   0  'Transparent
            Caption         =   "Narration"
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
            Left            =   5280
            TabIndex        =   34
            Top             =   75
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Total"
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
            Left            =   11640
            TabIndex        =   33
            Top             =   75
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cl Balance"
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
            Left            =   9480
            TabIndex        =   32
            Top             =   75
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   615
         Left            =   0
         TabIndex        =   22
         Top             =   7560
         Width           =   14295
         Begin VB.TextBox TxtDiff 
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
            Height          =   360
            Left            =   11400
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   1700
         End
         Begin VB.TextBox TxtDebit 
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
            Height          =   360
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   120
            Width           =   1700
         End
         Begin VB.TextBox txtCredit 
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
            Height          =   360
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   120
            Width           =   1700
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Debit"
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
            Left            =   120
            TabIndex        =   28
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Credit"
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
            Left            =   5160
            TabIndex        =   27
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Diff."
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
            Left            =   10440
            TabIndex        =   26
            Top             =   180
            Width           =   975
         End
      End
      Begin MSDataListLib.DataCombo AccountCombo 
         Bindings        =   "VouFrm1.frx":0050
         Height          =   360
         Left            =   1200
         TabIndex        =   42
         ToolTipText     =   "1:  ALT + Down Arrow key to open list.    2 :   Enter key to select.    3 :  F3  to create new account."
         Top             =   5640
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
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
      Begin MSDataGridLib.DataGrid VchGrid 
         Bindings        =   "VouFrm1.frx":0065
         Height          =   3960
         Left            =   0
         TabIndex        =   15
         Top             =   3600
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   6985
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   16777215
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
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
         Caption         =   "*******  Voucher Details  *******"
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "AC_CODE"
            Caption         =   "Ac Code"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "HH:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "AC_NAME"
            Caption         =   "Account Name"
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
            DataField       =   "DR_CR"
            Caption         =   "D/C"
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
         BeginProperty Column03 
            DataField       =   "AMOUNT"
            Caption         =   "Amount"
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
         BeginProperty Column04 
            DataField       =   "Narration"
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
         BeginProperty Column05 
            DataField       =   "CHQNO"
            Caption         =   "Chq. No"
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
            DataField       =   "CHQDT"
            Caption         =   "Chq Date"
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
            DataField       =   "BANK"
            Caption         =   "Bank"
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
            DataField       =   "BRANCH"
            Caption         =   "Branch"
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
            DataField       =   "CLBAL"
            Caption         =   "CL Bal"
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
            DataField       =   "AC_CODE1"
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
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   1
               Locked          =   -1  'True
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               DividerStyle    =   1
               Locked          =   -1  'True
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               DividerStyle    =   1
               Locked          =   -1  'True
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               DividerStyle    =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column05 
               DividerStyle    =   1
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               DividerStyle    =   1
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column07 
               DividerStyle    =   1
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               DividerStyle    =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column09 
               DividerStyle    =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
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
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.TextBox TxtPartClBal 
      Height          =   1695
      Left            =   14760
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14640
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport VOULIST 
      Left            =   15120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   255
      Left            =   14880
      TabIndex        =   46
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   8460
      Left            =   120
      Top             =   720
      Width           =   13755
   End
End
Attribute VB_Name = "VouFrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NarrRec As ADODB.Recordset:     Dim TempVch As ADODB.Recordset:
Public AccRec As ADODB.Recordset:   Public GeneralCB As ADODB.Recordset
Public Fb_Press As Byte:            Dim MVou_Type As String:        Public F_Payrpt As String:          Dim LVouNetAmt As Double
Dim VouPrnRec As ADODB.Recordset:   Public F_Vou_Dt As String:      Public F_Vou_No As String:          Public F_VOU_NO_OLD As String
Dim LFileName As String:            Dim LDebitChr As String:        Dim LCreditChr As String
Dim MDt As Date
Dim RowVouId As Integer
Dim Lcashbankacccount As Integer
Private Sub AccountCombo_DblClick(Area As Integer)
    Call AccountCombo_KeyPress(13)
End Sub
Private Sub AccountCombo_GotFocus()
    VchGrid.Col = 0:  AccountCombo.Top = Val(VchGrid.Top) + Val(VchGrid.RowTop(VchGrid.Row)): AccountCombo.text = VchGrid.text
    'Call LSendKeys_Down
    Sendkeys "%{DOWN}"
End Sub
Private Sub AccountCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Or KeyCode = 9 Then
        VchGrid.SetFocus: AccountCombo.Visible = False
    End If
End Sub


Private Sub CmdImport_Click()
 On Error GoTo err1
 Dim Jcnn  As ADODB.Connection:     Dim TxtRec As ADODB.Recordset
 Dim TxtPath As String:             Dim LVDate As Date:     Dim LVAmnt As Double:              Dim LVNarr As String
 Dim LVNo As String:                Dim LAcc1 As String:    Dim LAcc2 As String:               Dim LDrCr1 As String
 Dim LDrCr2 As String:              Dim LTranNo As Long:    Dim LMaxVno As Long:               Dim LVouNo As String
 Dim LFileSource As String:         Dim LNarr As String:    Dim LVAmt As Double
 Dim LMVNo As Long:                  Dim LAccID1 As Long:   Dim LAccID2 As Long
 
 
 CommonDialog1.InitDir = App.Path
 CommonDialog1.ShowOpen
 If CommonDialog1.FileName <> "" Then
    Text1.text = CommonDialog1.FileName
    TxtPath = CommonDialog1.FileTitle
    LFileName = CommonDialog1.FileName
    LFileSource = Left$(LFileName, (Len(LFileName) - Len(TxtPath)) - 1) & ";"
    Set Jcnn = Nothing
    Set Jcnn = New ADODB.Connection
    Jcnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & LFileSource & _
    "Extended Properties=""TEXT;HDR=No;IMEX=1;FMT=Delimited"""
    If Not FileExist(LFileName) Then
        MsgBox LFileName & "  file not found", vbCritical
    Else
        Set TxtRec = Nothing: Set TxtRec = New ADODB.Recordset
        mysql = "Select * from " & TxtPath & " "
        TxtRec.Open mysql, Jcnn, adOpenStatic, adLockReadOnly, adCmdText
    End If
    If Not TxtRec.EOF Then
        MDt = DateValue(DTPicker1.Value)
        mysql = "DELETE FROM VOUCHER WHERE COMPCODE =" & GCompCode & " AND VOU_TYPE='JV' AND IMPORT ='1' "
        mysql = mysql & "AND VOU_DT ='" & Format(MDt, "YYYY/MM/DD") & "'"
        Cnn.Execute mysql
        Do While Not TxtRec.EOF
            LVDate = DateValue(TxtRec!F1):            LAcc1 = TxtRec!F2
            LAcc2 = TxtRec!f3:                        LVAmnt = Val(TxtRec!F4)
            LNarr = TxtRec!F5:                        LTranNo = Val(TxtRec!f6)
            LAcc1 = Get_AccountMCode(LAcc1)
            If LenB(LAcc1) < 1 Then
                MsgBox "Ac Code Does Not Exist" & LAcc1 & ""
                GoTo FLAG_NEXT
            End If
            
            LAcc2 = Get_AccountMCode(LAcc2)
            If LenB(LAcc2) < 1 Then
                MsgBox "Ac Code Does Not Exist" & LAcc2 & ""
                GoTo FLAG_NEXT
            End If
            If LVAmnt > 0 Then
                LDrCr1 = "C":                LDrCr2 = "D"
            Else
                LDrCr1 = "D":                LDrCr2 = "C"
            End If
            LAccID1 = Get_AccID(LAcc1)
            LAccID2 = Get_AccID(LAcc2)
            
            
            LMaxVno = Val(Right$(Get_VouNo("JRNL", GFinYear), 7))
            LVouNo = Get_Next_Vou_No(LMaxVno, "JRNL", Right$(GFinBegin, 2) & Right$(GFinEnd, 2))
            LMVNo = PInsert_Voucher(LVouNo, LVDate, "JV", "P", LTranNo, "ADD", vbNullString, 0, LNarr, vbNullString, "1", 0, 0)
            Call PInsert_Vchamt(LVouNo, "JV", LVDate, LDrCr1, LAcc1, Abs(LVAmnt), vbNullString, LVDate, LNarr, vbNullString, vbNullString, 0, vbNullString, LMVNo, 0, 0, LAccID1)
            Call PInsert_Vchamt(LVouNo, "JV", LVDate, LDrCr2, LAcc2, Abs(LVAmnt), vbNullString, LVDate, LNarr, vbNullString, vbNullString, 0, vbNullString, LMVNo, 0, 0, LAccID1)
FLAG_NEXT:
            TxtRec.MoveNext
        Loop
    End If
End If

FLAG20:
MsgBox "File Import Complete"""
Exit Sub

err1:
    If err.Number <> 0 Then
        MsgBox err.Description
        Exit Sub
    End If
End Sub

Private Sub CmdNewAccount_Click()
    GETACNT.Show
    GETACNT.ZOrder
    GETACNT.add_record
End Sub

Private Sub AccountCombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And LenB(AccountCombo.BoundText) > 0 Then
        AccRec.MoveFirst
        AccRec.Find "AC_CODE='" & AccountCombo.BoundText & "' ", , adSearchForward
        If Not AccRec.EOF Then
            TempVch!AC_CODE = AccRec!AC_CODE
            TempVch!AC_NAME = AccRec!AC_NAME
            Call Get_PartyBal(AccRec!AC_CODE, AccRec!AC_NAME)
            If LenB(TempVch!DR_CR & vbNullString) < 1 Then       ''JUST IN CASE OF ADD NEW
                Select Case ComboVouType.ListIndex
                    Case 0, 1
                        TempVch!DR_CR = LCreditChr
                    Case 2, 3, 4, 5
                        TempVch!DR_CR = LDebitChr
                End Select
            End If
            If LenB(VchGrid.Columns(3).text) < 1 Then TempVch!AMOUNT = 0
            TempVch!chqdt = DTPicker1.Value
            TempVch!CL_BAL = 0:                       TempVch!VOUTYPE = MVou_Type
            TempVch!VchNo = TxtVouNo.text:            TempVch!VCHDT = DTPicker1.Value
            TempVch!G_CODE = AccRec!GCODE:            TempVch.Update
        End If
        AccRec.MoveFirst:        VchGrid.Col = 2
        VchGrid.SetFocus:        VchGrid.Refresh
        AccountCombo.Visible = False
    End If
End Sub
Private Sub CmdPrintVoucher_Click()
    Call VouPrinting
End Sub
Private Sub Command1_Click()
    '>>>add row in grid
    
    Dim TRec As ADODB.Recordset:    Dim LRowNo As Integer:    Dim LCounter As Integer
    On Error GoTo ERR4
    Text2.text = UCase(Text2.text)
    Txtdrcr.text = UCase(Txtdrcr.text)
    If Text2.Visible And Text2.text <> "D" And Text2.text <> "C" Then
        MsgBox "Invalid entry", vbCritical
        Text2.SetFocus
        Exit Sub
    ElseIf Txtdrcr.text <> "D" And Txtdrcr.text <> "C" Then
        MsgBox "Invalid entry", vbCritical
        Txtdrcr.SetFocus
        Exit Sub
    ElseIf Text2.Visible And Text2.text = Txtdrcr.text Then
        MsgBox "Invalid entry", vbCritical
        Txtdrcr.SetFocus
        Exit Sub
    End If
        
    If LenB(DataCombo1.BoundText) > 0 Then
        'If txtnarration.text = "" Then
        '    txtnarration.text = TXT_NARR.text
        'End If
        If (txthiddenid.text = "0" Or txthiddenid.text = "") Then
            RowVouId = RowVouId + 1
            If ComboVouType.ListIndex > 1 Then
                TempVch.AddNew
                    TempVch!VOUTYPE = MVou_Type:
                    TempVch!VchNo = TxtVouNo.text
                    TempVch!VCHDT = DTPicker1.Value
                    TempVch!AMOUNT = txtamoount.text
                    TempVch!NARRATION = TXT_NARR.text 'TXT_NARR.text
                    TempVch!CHQNO = txtchqno.text
                    TempVch!chqdt = Format(chkdt.Value, "dd/MM/yyyy")
                    TempVch!BANK = txtbank.text
                    TempVch!BRANCH = txtbranch.text
                    TempVch!AC_CODE = DComboCashBank.BoundText
                    TempVch!G_CODE = "0"
                    TempVch!AC_NAME = UCase(DComboCashBank.text)
                    TempVch!CL_BAL = 0:
                    TempVch!VOU_ID = RowVouId
                    TempVch!DR_CR = UCase(Text2.text)
                TempVch.Update
            End If
            
            'RowVouId = RowVouId + 1
            TempVch.AddNew
                TempVch!VOUTYPE = MVou_Type:
                TempVch!VchNo = TxtVouNo.text
                TempVch!VCHDT = DTPicker1.Value
                TempVch!AMOUNT = txtamoount.text
                TempVch!NARRATION = txtnarration.text
                TempVch!CHQNO = txtchqno.text
                TempVch!chqdt = Format(chkdt.Value, "dd/MM/yyyy")
                TempVch!BANK = txtbank.text
                TempVch!BRANCH = txtbranch.text
                TempVch!AC_CODE = DataCombo1.BoundText
                TempVch!G_CODE = "0"
                TempVch!AC_NAME = UCase(DataCombo1.text)
                TempVch!CL_BAL = 0:
                TempVch!VOU_ID = RowVouId
                TempVch!DR_CR = UCase(Txtdrcr.text)
            TempVch.Update
        Else
            Dim lrcount As Integer
            lrcount = 0
            TempVch.MoveFirst
            TempVch.Find "VOU_ID='" & txthiddenid.text & "' ", , adSearchForward
            While Not TempVch.EOF
                If (TempVch!VOU_ID = txthiddenid.text) Then
                    If ComboVouType.ListIndex > 1 Then 'jv
                        lrcount = lrcount + 1
                        If lrcount = 1 Then
                            TempVch!AMOUNT = txtamoount.text
                            TempVch!NARRATION = txtnarration.text
                            TempVch!CHQNO = txtchqno.text
                            TempVch!chqdt = Format(chkdt.Value, "dd/MM/yyyy")
                            TempVch!BANK = txtbank.text
                            TempVch!BRANCH = txtbranch.text
                            TempVch!AC_CODE = DComboCashBank.BoundText
                            TempVch!G_CODE = "0"
                            TempVch!AC_NAME = UCase(DComboCashBank.text)
                            TempVch!CL_BAL = 0:
                            TempVch!VOU_ID = txthiddenid.text
                            TempVch!DR_CR = "D"
                        Else
                            TempVch!AMOUNT = txtamoount.text
                            TempVch!NARRATION = txtnarration.text
                            TempVch!CHQNO = txtchqno.text
                            TempVch!chqdt = Format(chkdt.Value, "dd/MM/yyyy")
                            TempVch!BANK = txtbank.text
                            TempVch!BRANCH = txtbranch.text
                            TempVch!AC_CODE = DataCombo1.BoundText
                            TempVch!G_CODE = "0"
                            TempVch!AC_NAME = UCase(DataCombo1.text)
                            TempVch!CL_BAL = 0:
                            TempVch!VOU_ID = txthiddenid.text
                            TempVch!DR_CR = "C"
                        End If
                    Else
                        TempVch!AMOUNT = txtamoount.text
                        TempVch!NARRATION = txtnarration.text
                        TempVch!CHQNO = txtchqno.text
                        TempVch!chqdt = Format(chkdt.Value, "dd/MM/yyyy")
                        TempVch!BANK = txtbank.text
                        TempVch!BRANCH = txtbranch.text
                        TempVch!AC_CODE = DataCombo1.BoundText
                        TempVch!G_CODE = "0"
                        TempVch!AC_NAME = UCase(DataCombo1.text)
                        TempVch!CL_BAL = 0:
                        TempVch!VOU_ID = txthiddenid.text
                        TempVch!DR_CR = UCase(Txtdrcr.text)
                    End If
                    TempVch.Update
                End If
                TempVch.MoveNext
            Wend
        End If
        TempVch.MoveFirst
        Set VchGrid.DataSource = TempVch:
        VchGrid.ReBind:
        VchGrid.Refresh
        Call Show_VchTotal
        DataCombo1.text = ""
        DataCombo1.BoundText = ""
        'Txtdrcr.text = ""
        txtamoount.text = ""
        Txtnarrhidden.text = txtnarration.text
        'txtnarration.text = ""
        txtchqno.text = ""
        'chkdt.text = ""
        txtbank.text = ""
        txtbranch.text = ""
        txthiddenid.text = ""
    End If
    If ComboVouType.ListIndex > 1 Then
        DComboCashBank.SetFocus
    Else
        DataCombo1.SetFocus
    End If
        
    Exit Sub
ERR4:
If err.Number <> 0 Then
    MsgBox err.Description
End If
End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
    Call DataCombo1_KeyPress(13)
End Sub
Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And LenB(DataCombo1.BoundText) > 0 Then
        Select Case ComboVouType.ListIndex
            Case 0, 1
                If pmt_opn.Value = True Then
                    Txtdrcr.text = LDebitChr
                Else
                    Txtdrcr.text = LCreditChr
                End If
            Case 2, 3, 4, 5
                If (Text2.text = "D") Then
                    Txtdrcr.text = "C"
                Else
                    Txtdrcr.text = "D"
                End If
        End Select
        
        Call setscreen
        'Sendkeys "{tab}"
        'txtamoount.SetFocus
    End If
End Sub
Private Sub setscreen()
        If ComboVouType.ListIndex = "1" Then
            txtchqno.Enabled = True
            chkdt.Enabled = True
            txtbank.Enabled = True
            txtbranch.Enabled = True
            'Txtdrcr.Locked = False
            chkdt.Value = Date
        Else
            txtchqno.Enabled = False
            chkdt.Enabled = False
            txtbank.Enabled = False
            txtbranch.Enabled = False
            If ComboVouType.ListIndex >= 2 Then 'JV
                'Txtdrcr.Locked = False
                Text2.Locked = False
            End If
        End If
End Sub
Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And LenB(DataCombo1.BoundText) > 0 Then
        Sendkeys "{tab}"
    End If
End Sub

Private Sub DComboCashBank_GotFocus()
    Sendkeys "%{DOWN}"
    'Call LSendKeys_Down
End Sub

Private Sub DComboCashBank_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And LenB(DComboCashBank.BoundText) > 0 Then
        Select Case ComboVouType.ListIndex
            Case 0, 1
                If pmt_opn.Value = True Then
                    Text2.text = LCreditChr
                Else
                    Text2.text = LDebitChr
                End If
            Case 2, 3, 4, 5
                Text2.text = "D"
        End Select
    End If
End Sub

Private Sub DComboCashBank_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub DComboCashBank_Validate(Cancel As Boolean)
    On Error GoTo Error1
    If (ComboVouType.ListIndex = 0 Or ComboVouType.ListIndex = 1) Then
        If LenB(DComboCashBank.BoundText) > 0 Then
            GeneralCB.MoveFirst: GeneralCB.Find "AC_CODE = '" & DComboCashBank.BoundText & "'", , adSearchForward
            If Not GeneralCB.EOF Then
                TxtCashBankCode = GeneralCB!AC_CODE
                Call Get_Bal
            Else
                MsgBox "Please Select Valid Account"
                Cancel = True
            End If
        Else
            GeneralCB.MoveFirst
            DComboCashBank.text = GeneralCB!AC_NAME
            DComboCashBank.BoundText = GeneralCB!AC_CODE
            TxtCashBankCode = GeneralCB!AC_CODE
            Call Get_Bal
        End If
    End If
Error1:
    If err.Number <> 0 Then
        MsgBox err.Description
        Call CANCEL_RECORD
    End If
End Sub
Private Sub DComboNarr_DblClick(Area As Integer)
Call DComboNarr_KeyPress(13)
End Sub

Private Sub DComboNarr_GotFocus()
    VchGrid.Col = 0
    DComboNarr.Top = Val(VchGrid.Top) + Val(VchGrid.RowTop(VchGrid.Row))
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboNarr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Or KeyCode = 9 Then
    VchGrid.SetFocus: DComboNarr.Visible = False
End If
End Sub
Private Sub DComboNarr_KeyPress(KeyAscii As Integer)
Dim NarrName As String
    If KeyAscii = 13 Then VchGrid.Columns(4).text = DComboNarr.text
    VchGrid.Col = 9
    VchGrid.SetFocus
    VchGrid.Refresh
    DComboNarr.Visible = False
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    Dim MDt As String
    
    If SYSTEMLOCK(DateValue(DTPicker1.Value)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
        Cancel = True
    End If
    If DTPicker1.Value < GFinBegin Then
        MsgBox "Date can not be before financial year begin date.", vbCritical:
        Cancel = True
    End If
    If DTPicker1.Value > GFinEnd Then
        MsgBox "Date can not be beyond financial year end date.", vbCritical
        Cancel = True
    End If
    MDt = Format(Now, "dd/mm/yyyy")
    If DTPicker1.Value > MDt Then
        MsgBox "Voucher Date Is Greater than Current Date"
    End If

    Call Get_NewVouNo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo err1
    'If Frame6.Enabled Then
        If (Me.ActiveControl.NAME = "DTPicker1") Or (Me.ActiveControl.NAME = "chkdt") Then
            If KeyCode = 13 Then Sendkeys "{tab}"
        ElseIf (KeyCode = 113) Then 'F2
           If NarrRec.RecordCount > 0 Then
                DComboNarr.Visible = True
                DComboNarr.SetFocus
           Else
           DComboNarr.Visible = True
           End If
        End If
    'End If
err1:
End Sub

Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    If GETMAIN.ActiveForm.NAME = Me.NAME Then
        Call PERMISSIONS("VCHENT")
        GETMAIN.StatusBar1.Panels(1).text = "Voucher Details"
        Call Items
        If Fb_Press <> 0 Then
            If (GSysLockDt < DTPicker1.Value) Then
                Call Get_Selection(Fb_Press)
            End If
        End If
        Call Show_VchTotal
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        CRViewer1.Visible = False
        Cancel = 1
    Else
        Call CLEAR_SCREEN
        FLAG_QRYACC = False: Fb_Press = 0
        GETMAIN.StatusBar1.Panels(1).text = vbNullString
        Unload Me
    End If
End Sub

Private Sub pmt_opn_Click()
    Call Get_NewVouNo
End Sub
Private Sub pmt_opn_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Rpt_opn_Click()
    Call Get_NewVouNo
End Sub
Private Sub Rpt_opn_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If Text2.text = "C" Or Text2.text = "D" Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub
Private Sub txtamoount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{tab}"
    End If
End Sub
Private Sub txtbank_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub txtbranch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtCashBankCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub TxtCashBankCode_Validate(Cancel As Boolean)
If LenB(TxtCashBankCode.text) > 0 Then
    GeneralCB.MoveFirst: GeneralCB.Find "AC_CODE = '" & TxtCashBankCode.text & "'", , adSearchForward
    If Not GeneralCB.EOF Then
        DComboCashBank.BoundText = GeneralCB!AC_CODE
        Call Get_Bal
        TXT_NARR.SetFocus
    End If
End If
End Sub
Private Sub TXT_NARR_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtchqno_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub Txtdrcr_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txtnarration_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

'Private Sub VCHGRID_AfterColEdit(ByVal ColIndex As Integer)
'    If ColIndex = 2 Then
'        TempVch!DR_CR = UCase(TempVch!DR_CR)
'        If Not (UCase(TempVch!DR_CR) = LCreditChr Or UCase(TempVch!DR_CR) = LDebitChr) Then TempVch.CancelUpdate
'        TempVch.Update
'    ElseIf ColIndex = 0 And Len(Trim(VchGrid.text & vbNullString)) >= 0 Then
'        AccRec.MoveFirst
'        AccRec.Find "AC_CODE='" & VchGrid.text & "' ", , adSearchForward
'        If Not AccRec.EOF Then
'            TempVch!AC_CODE = AccRec!AC_CODE
'            TempVch!AC_NAME = AccRec!AC_NAME
'            Call Get_PartyBal(AccRec!AC_CODE, AccRec!AC_NAME)
'            If Len(TempVch!DR_CR & vbNullString) = Val(0) Then        ''JUST IN CASE OF ADD NEW
'                Select Case ComboVouType.ListIndex
'                    Case 0, 1
'                        If Rpt_opn.Value = True Then
'                            TempVch!DR_CR = LCreditChr
'                        Else
'                            TempVch!DR_CR = LDebitChr
'                        End If
'                    Case 2, 3, 4, 5
'                        TempVch!DR_CR = LCreditChr
'                End Select
'            End If
'            If VchGrid.Columns(3).text = "" Then TempVch!AMOUNT = 0
'            TempVch!chqdt = DTPicker1.Value
'            TempVch!CL_BAL = 0
'            TempVch!VOUTYPE = MVou_Type
'            TempVch!VchNo = TxtVouNo.text
'            TempVch!VCHDT = DTPicker1.Value
'            TempVch!G_CODE = AccRec!GCODE
'            TempVch.Update
'            VchGrid.Col = 2
'        Else
'            TempVch!AC_CODE = vbNullString
'            TempVch!AC_NAME = vbNullString
'        End If
'    End If
'    Call Show_VchTotal
'End Sub
Private Sub VchGrid_DblClick()
                   
    If ComboVouType.ListIndex > 1 Then
        Dim lid As String
        lid = VchGrid.Columns(14).text
        TempVch.MoveFirst
        TempVch.Find "VOU_ID='" & lid & "' ", , adSearchForward
        While Not TempVch.EOF
            If TempVch!VOU_ID = lid Then
                If (TempVch!DR_CR = "D") Then
                    DComboCashBank.BoundText = TempVch!AC_CODE
                    Text2.text = TempVch!DR_CR
                    TXT_NARR.text = TempVch!NARRATION
                Else
                    DataCombo1.BoundText = TempVch!AC_CODE
                    Txtdrcr.text = TempVch!DR_CR
                    txtamoount.text = TempVch!AMOUNT
                    txtnarration.text = TempVch!NARRATION
                    txtchqno.text = TempVch!CHQNO
                    chkdt.Value = TempVch!chqdt
                    txtbank.text = TempVch!BANK
                    txtbranch.text = TempVch!BRANCH
                    TempVch!CL_BAL = 0
                    txthiddenid.text = TempVch!VOU_ID
                End If
            Else
                GoTo exitsub
            End If
            TempVch.MoveNext
        Wend
    Else
        DataCombo1.BoundText = VchGrid.Columns(0).text
        Txtdrcr.text = VchGrid.Columns(2).text
        txtamoount.text = VchGrid.Columns(3).text
        txtnarration.text = VchGrid.Columns(4).text
        txtchqno.text = VchGrid.Columns(5).text
        chkdt.Value = VchGrid.Columns(6).text
        txtbank.text = VchGrid.Columns(7).text
        txtbranch.text = VchGrid.Columns(8).text
        TempVch!CL_BAL = 0:
        txthiddenid.text = VchGrid.Columns(14).text
    End If
exitsub:
    Call setscreen
    DataCombo1.SetFocus
End Sub
'Private Sub VCHGRID_KeyDown(KeyCode As Integer, Shift As Integer)
'    If (KeyCode = 13 Or KeyCode = 9) And VchGrid.Col <> 9 Then
'        If VchGrid.Col = 1 And Len(Trim(TempVch!AC_NAME & vbNullString)) = 0 Then
'            If Frame6.Enabled Then
'                AccountCombo.Visible = True: AccountCombo.SetFocus
'            End If
'        ElseIf VchGrid.Col = 4 And Len(VchGrid.text & vbNullString) < 1 Then
'           If NarrRec.RecordCount > 0 Then
'                DComboNarr.Visible = True
'                DComboNarr.SetFocus
'           End If
'        ElseIf VchGrid.Col = 3 Then
'            If Val(VchGrid.Columns(2).text) > 0 Then
'            End If
'        End If
'    ElseIf Not (KeyCode = 9 And Shift = 1) And VchGrid.Col = 9 Then ''ONLY THROUGH SHIFT+TAB ONE CAN LEAVE THE GRID
'        If KeyCode < Val(14) Then
'            KeyCode = 0                                             ''DONE BCOS ON LAST COL OF GRID IF ONE HITS TAB FOCUS GOES ON ITEM SELECTION COMBO AND ON THE LOST FOCUS ON DComboCashBank SOCNUMBER CHANGES
'            Call VchGrid.SetFocus
'        End If
'    ElseIf VchGrid.Col = 4 And Shift = 0 Then
'        If Not (KeyCode = 8 Or KeyCode = 13 Or KeyCode = 27 Or KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 46) Then
'            If Len(VchGrid.text & vbNullString) >= Val(200) Then
'                MsgBox "Length Overflow.", vbExclamation, "Warning": KeyCode = 0: VchGrid.SetFocus
'            End If
'        End If
'    End If
'End Sub
'Private Sub VCHGRID_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim MVou_Id As Long
'    If (KeyCode = 9 Or KeyCode = 13) And VchGrid.Col = 9 Then
'        If Not TempVch.EOF Then
'            TempVch.Update
'            MVou_Id = TempVch!VOU_ID
'            TempVch.MoveNext
'            VchGrid.LeftCol = 0
'            VchGrid.Col = 0
'        End If
'        If TempVch.EOF Then
'            TempVch.AddNew
'            TempVch!VOU_ID = MVou_Id + 1
'            If pmt_opn.Value = True Then
'                TempVch!DR_CR = "D"
'            Else
'                TempVch!DR_CR = "C"
'            End If
'        End If
'        TempVch.Update
'        VchGrid.ReBind
'        VchGrid.Refresh
'        VchGrid.Col = 0
'    ElseIf KeyCode = 13 Then  ''right arrow key
'        Sendkeys "{TAB}"
'    End If
'End Sub
Private Sub Form_Load()
    If GUniqClientId = "2386CHN" Then
        LDebitChr = "D"
        LCreditChr = "L"
    ElseIf GUniqClientId = "2387MUM" Then
        LDebitChr = "U"
        LCreditChr = "J"
    Else
        LDebitChr = "D"
        LCreditChr = "C"
    End If
    mysql = "SELECT * FROM NARRATIONM WHERE COMPCODE=" & GCompCode & " ORDER BY NARRNAME"
    Set NarrRec = Nothing
    Set NarrRec = New ADODB.Recordset
    NarrRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    DComboNarr.text = vbNullString
    If Not NarrRec.EOF Then
        Set DComboNarr.RowSource = NarrRec
        DComboNarr.ListField = "NARRNAME"
        DComboNarr.BoundColumn = "NARRCODE"
    Else
        DComboNarr.Enabled = False
    End If
    
    Call PERMISSIONS("VCHENT")
    'Call ClearFormFn(VouFrm)
    Call MakeRec
    Set VchGrid.DataSource = TempVch: VchGrid.ReBind: VchGrid.Refresh
    DTPicker1.MaxDate = GFinEnd: DTPicker1.MinDate = GFinBegin
    If Date < GFinBegin Then
        DTPicker1.Value = GFinBegin
    ElseIf Date > GFinEnd Then
        DTPicker1.Value = GFinEnd
    Else
        DTPicker1.Value = Date
    End If
    VchGrid.Columns(9).Locked = True: VchGrid.Columns(10).Visible = False: VchGrid.Columns(5).Visible = False
    VchGrid.Columns(6).Visible = False: VchGrid.Columns(7).Visible = False
    VchGrid.Columns(8).Visible = False: VchGrid.Columns(10).Visible = False
    VchGrid.Columns(11).Visible = False: VchGrid.Columns(12).Visible = False
    VchGrid.Columns(13).Visible = False: VchGrid.Columns(14).Visible = False
    Call Items
    If Not FLAG_QRYACC Then Call CLEAR_SCREEN
End Sub
Sub add_record()
    Call Get_Selection(1)
    Call CLEAR_SCREEN
    Fb_Press = 1
    RowVouId = 0
    Call MakeRec
    Set VchGrid.DataSource = TempVch: VchGrid.ReBind: VchGrid.Refresh
    Frame6.Enabled = True:
    pr_frame.Enabled = True
    Get_NewVouNo
    If ComboVouType.Visible Then
    ComboVouType.SetFocus
    End If
End Sub
Sub Save_Record()

    Dim TRec As ADODB.Recordset:     Dim MMonth As String:    Dim LVOU_NO As String:    Dim LCashCode As String: Dim LMVNo As Long
    
    On Error GoTo err1
    If DTPicker1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: DTPicker1.SetFocus: Exit Sub
    If DTPicker1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: DTPicker1.SetFocus: Exit Sub
    
    If Val(Len(TxtVouNo.text)) < 13 Then MsgBox "Invalid Voucher No", vbInformation: ComboVouType.SetFocus: Exit Sub

    If ComboVouType.ListIndex = Val(0) Or ComboVouType.ListIndex = Val(1) Then
        If pr_frame.Visible Then
            If pmt_opn.Value = True Then
                F_Payrpt = "PAYMENT"
            Else
                F_Payrpt = "RECEIPT"
            End If
        Else
            F_Payrpt = "0"
        End If
    Else
        F_Payrpt = "0"
    End If
    
    If TempVch.RecordCount > 0 Then
        TempVch.MoveFirst
        Do While Not TempVch.EOF
            If LenB(Trim$(TempVch!AC_CODE & vbNullString)) < 1 Then TempVch.Delete
            TempVch.MoveNext
        Loop
    Else
        MsgBox "Invalid Entry", vbInformation: Exit Sub
    End If
    
    TempVch.MoveFirst
    If Not TempVch.EOF Then                     ''''SOME VALIDATIONS
        If ComboVouType.ListIndex = 2 Or ComboVouType.ListIndex = 3 Or ComboVouType.ListIndex = 4 Or ComboVouType.ListIndex = 5 Then     ''''IN JV := BALANCE AMT MUST BE ZERO
            Call Show_VchTotal
            If LVouNetAmt <> 0 Then
                MsgBox "Total mismatch. Credit and Debit must be same.", vbExclamation, "Error"

                VchGrid.Col = 0:   VchGrid.SetFocus: Exit Sub
            End If
        Else                                    ''''IF CASH OR BANK ACCOUNT NOT SELECTED
            If Len(Trim(DComboCashBank.BoundText)) < 1 Then
                MsgBox ComboVouType.text & " A/c. Not Selected. Select Account before save.", vbExclamation, "Error"
                DComboCashBank.SetFocus: Exit Sub
            Else
                LCashCode = Get_AccountMCode(Trim(DComboCashBank.BoundText))
                If LenB(LCashCode) < 1 Then
                    MsgBox "Invalid Cash Ac "
                    DComboCashBank.SetFocus: Exit Sub
                End If
            End If
        End If
        If Fb_Press = 1 Then    '' NEW VOU_NO
            mysql = "SELECT VOU_NO FROM VOUCHER WHERE COMPCODE = " & GCompCode & " AND VOU_NO='" & TxtVouNo.text & "'"
            Set TRec = Nothing: Set TRec = New ADODB.Recordset:
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                MsgBox ("Voucher Number : " & TxtVouNo.text & " already exists. Please restart Voucher Form ")
                Exit Sub
            End If
        End If
        If Fb_Press = 2 Then Call Delete_Entry
        Get_Voucher_Type
        Cnn.BeginTrans
        CNNERR = True
        LMVNo = PInsert_Voucher(TxtVouNo.text, DTPicker1.Value, MVou_Type, F_Payrpt, 0, "ADD", LCashCode, 0, Trim(TXT_NARR.text), vbNullString, "0", 0, 0)
        TempVch.MoveFirst
        Call ENT_SAVE(LMVNo)
        mysql = "DELETE FROM VCHAMT WHERE AMOUNT=0"
        Cnn.Execute mysql
        mysql = "DELETE FROM VOUCHER WHERE VOU_NO NOT IN (SELECT DISTINCT VOU_NO FROM VCHAMT)"
        Cnn.Execute mysql
        
        Cnn.CommitTrans
        CNNERR = False
        If GUniqClientId = "89IND" Then
            If MsgBox("Print Voucher?", vbYesNo) = vbYes Then
                Call VouPrinting
            End If
        End If
        Call Items
        Call CLEAR_SCREEN
        '-----
        Call Get_Selection(4)
    End If
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End Sub
Sub Delete_Entry()
    On Error GoTo Error1
    Cnn.Execute "DELETE FROM VOUCHER WHERE COMPCODE = " & GCompCode & " AND VOU_NO = '" & F_VOU_NO_OLD & "'"
    Cnn.Execute "DELETE FROM VCHAMT WHERE  COMPCODE = " & GCompCode & " AND VOU_NO = '" & F_VOU_NO_OLD & "'"
    If Fb_Press = 3 Then
        TempVch.MoveFirst
        Do While Not TempVch.EOF
            TempVch.Delete: TempVch.Update: TempVch.MoveNext
        Loop
        VchGrid.Refresh: VchGrid.ReBind
        Call CLEAR_SCREEN
        Call Get_Selection(4)
    End If
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False

End Sub
Sub RECORD_ACCESS()
Dim month As String
Dim MON As String
    Call Get_Selection(2)
    Call MakeRec
    VchGrid.Refresh: VchGrid.ReBind
    Frame1.Enabled = True: ComboVouType.Enabled = True: pr_frame.Enabled = True: pmt_opn.Value = True
    Label3.Caption = "Date": TxtVouNo.Visible = False: ComboVouType.SetFocus: ComboVouType.ListIndex = 0
End Sub
Sub CANCEL_RECORD()
    Call CLEAR_SCREEN
    Call Get_Selection(5)
End Sub
Private Sub SHOW_SCR()  '' IN CASE OF VOUCHER MODIFY & DELETE  CALLED FROM --- VOUNO_COM_LostFocus
    Dim TRec As ADODB.Recordset:    Dim LRowNo As Integer:    Dim LCounter As Integer
    On Error GoTo ERR4
    Call MakeRec
    TempVch.Delete
    If MVou_Type = "CV" Or MVou_Type = "BV" Then
        mysql = "SELECT VO.VOU_TYPE,VO.VOU_NO,VO.VOU_DT,VO.CASHCODE,V.DR_CR,V.AMOUNT,V.NARRATION,V.CHEQUE_NO,V.CHEQUE_DT,V.BANK_NAME,V.BRANCH,V.AC_CODE,A.GCODE,A.NAME,ISNULL(V.INVNO,'0') AS 'INVNO' "
        mysql = mysql & " FROM VOUCHER AS VO ,VCHAMT AS V, "
        mysql = mysql & " ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode & " AND  VO.VOU_ID=V.VOU_ID "
        mysql = mysql & " AND V.VOU_TYPE = '" & MVou_Type & "' AND V.VOU_NO = '" & TxtVouNo.text & "' "
        mysql = mysql & " AND A.ACCID  = V.ACCID ORDER BY V.VOUID"
    Else
        mysql = "SELECT V.VOU_TYPE,V.VOU_NO, V.VOU_DT,V.DR_CR,V.AMOUNT,V.NARRATION,V.CHEQUE_NO,V.CHEQUE_DT,V.BANK_NAME,V.BRANCH,V.AC_CODE,A.GCODE,A.NAME,ISNULL(V.INVNO,'0') AS 'INVNO' FROM VCHAMT AS V, "
        mysql = mysql & " ACCOUNTM as A WHERE A.COMPCODE=" & GCompCode & "  AND  V.VOU_TYPE = '" & MVou_Type & "' AND V.VOU_NO = '" & TxtVouNo.text & "'"
        mysql = mysql & " AND A.ACCID = V.ACCID ORDER BY V.VOUID"
    End If
    Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If MVou_Type = "CV" Or MVou_Type = "BV" Then
        TRec.MoveLast
        DComboCashBank.BoundText = TRec!CASHCODE
        TxtCashBankCode.text = TRec!CASHCODE
        TXT_NARR.text = TRec!NARRATION
        Call Get_Bal
    End If
    LRowNo = TRec.RecordCount
    LCounter = 1
    TRec.MoveFirst
    Dim Lvouid As Integer
    Lvouid = 1
    While Not TRec.EOF
        TempVch.AddNew
        TempVch!VOUTYPE = TRec!VOU_TYPE:                        TempVch!VchNo = TRec!VOU_NO
        TempVch!VCHDT = TRec!VOU_DT:                            TempVch!AMOUNT = TRec!AMOUNT
        TempVch!NARRATION = TRec!NARRATION & vbNullString:      TempVch!CHQNO = TRec!CHEQUE_NO & vbNullString
        TempVch!chqdt = Format(TRec!CHEQUE_DT, "dd/MM/yyyy"):   TempVch!BANK = TRec!BANK_NAME & vbNullString
        TempVch!BRANCH = TRec!BRANCH & vbNullString:            TempVch!AC_CODE = TRec!AC_CODE
        TempVch!G_CODE = TRec!GCODE:                            TempVch!AC_NAME = UCase(TRec!NAME)
        TempVch!CL_BAL = 0:
        If MVou_Type = "CV" Or MVou_Type = "BV" Then
            TempVch!VOU_ID = Lvouid
        Else
            TempVch!VOU_ID = TRec!INVNO
        End If
        
        If TRec!DR_CR = "D" Then
            TempVch!DR_CR = LDebitChr
        Else
            TempVch!DR_CR = LCreditChr
        End If
        TempVch.Update
        Lvouid = Lvouid + 1
        TRec.MoveNext
        LCounter = LCounter + 1
        If LCounter = LRowNo And (MVou_Type = "CV" Or MVou_Type = "BV") Then
            If TxtCashBankCode.text = TRec!AC_CODE Then TRec.MoveNext
        End If
    Wend
    TempVch.MoveFirst
    RowVouId = Lvouid
    Set TRec = Nothing: Set VchGrid.DataSource = TempVch: VchGrid.ReBind: VchGrid.Refresh
    Exit Sub
ERR4:
If err.Number <> 0 Then
    MsgBox err.Description
End If
End Sub
Private Sub ENT_SAVE(LDMVnO As Long)
    Dim Mdr_Amt As Double:    Dim Mcr_Amt As Double:    Dim LNetAmt As Double:    Dim LDr As String: Dim LDrCr As String
    Dim LACCID As Long
    TempVch.MoveFirst
    Mdr_Amt = 0: Mcr_Amt = 0
    While Not TempVch.EOF
        If Val(TempVch!AMOUNT & vbNullString) > 0 Then
            If (UCase(TempVch!DR_CR) = LDebitChr Or UCase(TempVch!DR_CR) = LCreditChr) Then
                If UCase(TempVch!DR_CR) = LDebitChr Then
                    LDrCr = "D"
                Else
                    LDrCr = "C"
                End If
                LACCID = Get_AccID(TempVch!AC_CODE)
                Call PInsert_Vchamt(TxtVouNo.text, MVou_Type, DTPicker1.Value, LDrCr, TempVch!AC_CODE, Val(TempVch!AMOUNT), Trim$(TempVch!CHQNO & vbNullString), TempVch!chqdt, (TempVch!NARRATION & vbNullString), (TempVch!BRANCH & vbNullString), (TempVch!BANK & vbNullString), TempVch!VOU_ID, vbNullString, LDMVnO, 0, 0, LACCID)
            End If
            If LDrCr = "D" Then
                Mdr_Amt = Mdr_Amt + Val(TempVch!AMOUNT & vbNullString)
            ElseIf LDrCr = "C" Then
                Mcr_Amt = Mcr_Amt + Val(TempVch!AMOUNT & vbNullString)
            End If
        End If
        TempVch.MoveNext
    Wend
    'CBLIST,PRTYLIST ACCOUNT MODIFICATION (USING GRID TOTAL DEBIT & TOTAL CREDIT AMOUNT) AND INSERTION INTO VCHAMt
    If MVou_Type = "CV" Or MVou_Type = "BV" Then
        If Mdr_Amt > Mcr_Amt Then
            LNetAmt = Mdr_Amt - Mcr_Amt
            LDr = "C"
        ElseIf Mcr_Amt >= Mdr_Amt Then
            LNetAmt = Mcr_Amt - Mdr_Amt
            LDr = "D"
        End If
        LACCID = Get_AccID(TxtCashBankCode.text)
        Call PInsert_Vchamt(TxtVouNo.text, MVou_Type, DTPicker1.Value, LDr, TxtCashBankCode.text, LNetAmt, vbNullString, DTPicker1.Value, TXT_NARR.text, vbNullString, vbNullString, 0, vbNullString, LDMVnO, 0, 0, LACCID)
    End If

End Sub
Private Sub pmt_opn_LostFocus()
    F_Payrpt = "PAYMENT"
End Sub
Private Sub Rpt_opn_LostFocus()
    F_Payrpt = "RECEIPT"
End Sub
Private Sub ComboVouType_Click()
    Call Get_Voucher_Type
    Get_NewVouNo
End Sub
Private Sub ComboVouType_GotFocus()
    Sendkeys "%{DOWN}"
'    LSendKeys_Down
End Sub
Sub Items()
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    mysql = "SELECT NAME AS AC_NAME, AC_CODE, GCODE, OP_BAl FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & "  ORDER BY NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        Set AccountCombo.RowSource = AccRec
        AccountCombo.ListField = "AC_NAME"
        AccountCombo.BoundColumn = "AC_CODE"
        
        Set DataCombo1.RowSource = AccRec
        DataCombo1.ListField = "AC_NAME"
        DataCombo1.BoundColumn = "AC_CODE"
    End If
End Sub
Private Sub ComboVouType_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub ComboVouType_LostFocus()
    Call ComboVouType_Validate(False)
End Sub

Private Sub ComboVouType_Validate(Cancel As Boolean)
    On Error GoTo Error1
    If Len(Trim(ComboVouType.text)) < Val(1) Then ComboVouType.text = ComboVouType.List(0)
    Select Case ComboVouType.ListIndex
        Case 0                              ''CASH VOUCHER
            If Fb_Press = 1 Then
                If Not CASH_VOUCHER Then
                    'ComboVouType.ListIndex = -1:
                    Cancel = True
                    Exit Sub
                End If
                Get_Voucher_Type
            End If
        Case 1                              ''BANK VOUCHER
            If Fb_Press = 1 Then
                If Not Bank_Voucher Then    ''FUNCTION CALL
                   'C 'omboVouType.ListIndex = -1:
                   Cancel = True
                   Exit Sub
                End If
                Get_Voucher_Type
            End If
        Case 2, 3, 4, 5                   ''JV,DN,CN
            Get_Voucher_Type
            Call JV_Voucher
            pr_frame.Enabled = False:
    End Select
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    
End Sub
Function VOUCHER_ACCESS(LVOU_NO As String) As Boolean
    Dim TRec As ADODB.Recordset
            
    On Error GoTo Error1
    CmdPrintVoucher.Visible = True
    Label3.Caption = "Date": DTPicker1.Visible = True: TxtVouNo.Visible = True
    mysql = "SELECT VOU_NO,VOU_DT,VOU_TYPE,VOU_PR FROM VOUCHER  WHERE COMPCODE = " & GCompCode & " AND VOU_NO = '" & LVOU_NO & "'"
    Set TRec = Nothing: Set TRec = New ADODB.Recordset:
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        VOUCHER_ACCESS = True:
        ComboVouType.Enabled = False:
        pr_frame.Enabled = False ':     DComboCashBank.Locked = True
        TxtVouNo.text = LVOU_NO:
        F_Vou_No = LVOU_NO:
        F_VOU_NO_OLD = LVOU_NO
        DTPicker1.Value = TRec!VOU_DT
        
        If (GSysLockDt < DTPicker1.Value) Then
            Call Get_Selection(Fb_Press)
        Else
            Call Get_Selection(5)
            CmdImport.Enabled = False
        End If
    
        F_Vou_Dt = DTPicker1.Value
        F_Payrpt = TRec!VOU_PR
        If LenB(F_Payrpt) > 0 Then
            If Left$(F_Payrpt, 1) = "P" Then
                pmt_opn.Value = True
                Rpt_opn.Value = False
            Else
                pmt_opn.Value = False
                Rpt_opn.Value = True
            End If
        End If
        TXT_NARR.MaxLength = 4000:
        'TXT_NARR.text = IIf(IsNull(TRec!NARRATION), vbNullString, TRec!NARRATION)
        Select Case TRec!VOU_TYPE
            ' Case "CSFZHR", "CSHP"                     ''CASH VOUCHER
            Case "CV"
                ComboVouType.ListIndex = 0
                CASH_VOUCHER
                Get_Voucher_Type
                pr_frame.Enabled = True:
            Case "BV"
                ComboVouType.ListIndex = 1
                Bank_Voucher
                Get_Voucher_Type
                pr_frame.Enabled = True:
            Case "JV"
                ComboVouType.ListIndex = 2:
                Get_Voucher_Type
                Call JV_Voucher
            Case "M"
                ComboVouType.ListIndex = 3:
                Get_Voucher_Type
                Call JV_Voucher
            Case "F"
                ComboVouType.ListIndex = 5:
                Get_Voucher_Type
                Call JV_Voucher
            Case "S"
                MVou_Type = "S": VchGrid.Columns(4).Visible = False: Call JV_Voucher
            Case "H"
                MVou_Type = "H": VchGrid.Columns(4).Visible = False: Call JV_Voucher
            Case "B"
                MVou_Type = "B": VchGrid.Columns(4).Visible = False: Call JV_Voucher
            Case "O"
                MVou_Type = "O": VchGrid.Columns(4).Visible = False: Call JV_Voucher
            Case "K"
                ComboVouType.ListIndex = 4:
                Get_Voucher_Type
                Call JV_Voucher
        End Select
        Call SHOW_SCR
        Call Show_VchTotal
    Else
        VOUCHER_ACCESS = False
    End If
    Exit Function
Error1: If err.Number <> 0 And err.Number <> 5 Then
            
            MsgBox err.Description, vbInformation
        End If
End Function
Private Sub Get_Voucher_Type()
    Dim LVou_Type As String
    DComboCashBank.Enabled = True
    TXT_NARR.Visible = True: LblNarr.Visible = True
    Select Case ComboVouType.ListIndex
        Case 0, 1
            If ComboVouType.ListIndex = 0 Then
                MVou_Type = "CV"
                VchGrid.Columns(4).Width = 5000:        VchGrid.Columns(5).Visible = False:
                VchGrid.Columns(6).Visible = False:     VchGrid.Columns(7).Visible = False:
                VchGrid.Columns(8).Visible = False:
            ElseIf ComboVouType.ListIndex = 1 Then
                MVou_Type = "BV"
                VchGrid.Columns(2).AllowSizing = True:      VchGrid.Columns(2).AllowSizing = False
                VchGrid.Columns(0).AllowSizing = False:     VchGrid.Columns(5).Visible = True:
                VchGrid.Columns(6).Visible = True:          VchGrid.Columns(7).Visible = True
                VchGrid.Columns(8).Visible = True:
            End If
            TxtCashBankCode.Visible = True:
            pr_frame.Enabled = True:
            Label18.Visible = False: Text2.Visible = False
            TXT_NARR.Visible = False: LblNarr.Visible = False:
            If Lcashbankacccount = 1 Then
                DComboCashBank.Enabled = False
                TXT_NARR.MaxLength = 300: Label1.Visible = True: TxtClBal.Visible = True
            Else
                DComboCashBank.Enabled = True
            End If
        Case 2
            Text2.Locked = False
            MVou_Type = "JV":                     pr_frame.Enabled = False
            'DComboCashBank.Visible = False:       TXT_NARR.Visible = False
            'TxtCashBankCode.Visible = False:      LblAccount.Visible = False
            'LblNarr.Visible = False
            VchGrid.Columns(4).Width = 5000:      VchGrid.Columns(5).Visible = False:
            VchGrid.Columns(6).Visible = False:   VchGrid.Columns(7).Visible = False:
            VchGrid.Columns(8).Visible = False
            'Label1.Visible = False
            'TxtClBal.Visible = False
        Case 3
            Text2.Locked = False
            MVou_Type = "M":                      pr_frame.Enabled = False
            'DComboCashBank.Visible = False:       TXT_NARR.Visible = False
            'TxtCashBankCode.Visible = False:      LblAccount.Visible = False
            'LblNarr.Visible = False:              VchGrid.Columns(8).Visible = False
            VchGrid.Columns(4).Width = 5000:      VchGrid.Columns(5).Visible = False:
            VchGrid.Columns(6).Visible = False:   VchGrid.Columns(7).Visible = False:
            'Label1.Visible = False
            'TxtClBal.Visible = False
        Case 4
            Text2.Locked = False
            MVou_Type = "K":                      pr_frame.Enabled = False
            'DComboCashBank.Visible = False:       TXT_NARR.Visible = False
            'TxtCashBankCode.Visible = False:      LblAccount.Visible = False
            'LblNarr.Visible = False:              VchGrid.Columns(8).Visible = False
            VchGrid.Columns(4).Width = 5000:      VchGrid.Columns(5).Visible = False:
            VchGrid.Columns(6).Visible = False:   VchGrid.Columns(7).Visible = False:
            'Label1.Visible = False
            'TxtClBal.Visible = False
        Case 5
            Text2.Locked = False
            MVou_Type = "F":                      pr_frame.Enabled = False
            'DComboCashBank.Visible = False:       TXT_NARR.Visible = False
            'TxtCashBankCode.Visible = False:      LblAccount.Visible = False
            'LblNarr.Visible = False:              VchGrid.Columns(8).Visible = False
            VchGrid.Columns(4).Width = 5000:      VchGrid.Columns(5).Visible = False:
            VchGrid.Columns(6).Visible = False:   VchGrid.Columns(7).Visible = False:
            'Label1.Visible = False
            'TxtClBal.Visible = False
    End Select
End Sub
Function Bank_Voucher() As Boolean
    On Error GoTo Error1
    mysql = "SELECT NAME AS AC_NAME, AC_CODE FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " AND GCODE =11 ORDER BY NAME"
    Set GeneralCB = Nothing: Set GeneralCB = New ADODB.Recordset: GeneralCB.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not GeneralCB.EOF Then
        Set DComboCashBank.RowSource = GeneralCB
        DComboCashBank.ListField = "AC_NAME"
        DComboCashBank.BoundColumn = "AC_CODE"
        Lcashbankacccount = GeneralCB.RecordCount
        If Lcashbankacccount = 1 Then
            DComboCashBank.BoundText = GeneralCB!AC_CODE
            TxtCashBankCode.text = GeneralCB!AC_CODE
        End If
        
        Bank_Voucher = True
        VchGrid.Columns(2).AllowSizing = True
        VchGrid.Columns(0).AllowSizing = False
        VchGrid.Columns(2).AllowSizing = False
        Exit Function
    Else
        MsgBox "Please Create Bank Account with Account Type [Bank Account]  "
        'Call CANCEL_RECORD
        Bank_Voucher = False
        Exit Function
    End If
Error1:
    MsgBox err.Description, vbCritical, err.HelpFile
    
End Function
Function JV_Voucher()
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    mysql = "SELECT NAME AS AC_NAME, AC_CODE, GCODE, OP_BAl FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & "  ORDER BY NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        Set DComboCashBank.RowSource = AccRec
        DComboCashBank.ListField = "AC_NAME"
        DComboCashBank.BoundColumn = "AC_CODE"
        
        Lcashbankacccount = AccRec.RecordCount
        If Lcashbankacccount = 1 Then
            DComboCashBank.BoundText = AccRec!AC_CODE
            TxtCashBankCode.text = AccRec!AC_CODE
        End If
    End If
End Function

Function CASH_VOUCHER() As Boolean
    On Error GoTo Error1
    mysql = "SELECT NAME AS AC_NAME, AC_CODE  FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " AND GCODE =10  ORDER BY NAME"
    Set GeneralCB = Nothing: Set GeneralCB = New ADODB.Recordset: GeneralCB.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not GeneralCB.EOF Then
        Set DComboCashBank.RowSource = GeneralCB: DComboCashBank.ListField = "AC_NAME": DComboCashBank.BoundColumn = "AC_CODE"
        Lcashbankacccount = GeneralCB.RecordCount
        If Lcashbankacccount = 1 Then
            DComboCashBank.BoundText = GeneralCB!AC_CODE
            TxtCashBankCode.text = GeneralCB!AC_CODE
        End If
        CASH_VOUCHER = True
        Exit Function
    Else
        MsgBox "Please Create Cash Account with Account Type [Cash in Hand] "
        'Call CANCEL_RECORD
        CASH_VOUCHER = False
        Exit Function
    End If
Error1:
    MsgBox err.Description, vbCritical, err.HelpFile
    
End Function
Sub CLEAR_SCREEN()
    Dim OBJCONTROL As Control
    ComboVouType.Visible = True:            Label8.Visible = True:  'pr_frame.Visible = True:
    LVouNetAmt = 0:                 F_VOU_NO_OLD = vbNullString:    F_Vou_No = vbNullString:     Fb_Press = 0:
    TxtVouNo.text = vbNullString:       TxtCashBankCode.text = vbNullString
    TxtClBal.text = vbNullString:       TXT_NARR.text = vbNullString
    AmtTxt.text = vbNullString:         TxtDebit.text = vbNullString
    txtCredit.text = vbNullString:      TxtDiff.text = vbNullString
    CmdPrintVoucher.Visible = False:
    If Date < GFinBegin Then
        DTPicker1.Value = GFinBegin
    ElseIf Date > GFinEnd Then
        DTPicker1.Value = GFinEnd
    Else
        DTPicker1.Value = Date
    End If
    txthiddenid.text = ""
    ComboVouType.ListIndex = 0:
    DComboCashBank.text = vbNullString
    DataCombo1.text = vbNullString: Txtdrcr.text = "": txtamoount.text = "": txtnarration.text = "": Text2.text = "": TXT_NARR.text = ""
    txtchqno.text = "": txtbank.text = "": txtbranch.text = ""
    Call MakeRec
    Set VchGrid.DataSource = TempVch: VchGrid.ReBind: VchGrid.Refresh
    ComboVouType.Enabled = True:    DTPicker1.Enabled = True:       ComboVouType.Enabled = True:            TxtVouNo.Enabled = True
    DComboCashBank.Locked = False:  DTPicker1.Enabled = True:       TxtVouNo.Visible = True:                pr_frame.Enabled = True
    If MVou_Type = "BV" Then VchGrid.Columns(4).Visible = False:    VchGrid.Columns(5).Visible = False:     VchGrid.Columns(6).Visible = False:
    VchGrid.Columns(7).Visible = False
    AccountCombo.Visible = False
    
    VchGrid.LeftCol = 0: Frame6.Enabled = False
    If FLAG_QRYACC And LenB(MFormat) < 1 Then
        MFormat = "Query on Account": FLAG_QRYACC = False
        Call Get_Selection(12)
        Unload Me
    End If
End Sub
Sub MakeRec()
    Set TempVch = Nothing
    Set TempVch = New ADODB.Recordset
    TempVch.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
    TempVch.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
    TempVch.Fields.Append "AMOUNT", adDouble, , adFldIsNullable
    TempVch.Fields.Append "NARRATION", adVarChar, 1000, adFldIsNullable
    TempVch.Fields.Append "CHQNO", adVarChar, 15, adFldIsNullable
    TempVch.Fields.Append "CHQDT", adVarChar, 10, adFldIsNullable
    TempVch.Fields.Append "BANK", adVarChar, 25, adFldIsNullable
    TempVch.Fields.Append "BRANCH", adVarChar, 25, adFldIsNullable
    'no need
    TempVch.Fields.Append "CL_BAL", adDouble, , adFldIsNullable
    TempVch.Fields.Append "VOUTYPE", adVarChar, 2, adFldIsNullable
    TempVch.Fields.Append "VCHNO", adVarChar, 18, adFldIsNullable
    TempVch.Fields.Append "VCHDT", adDate, , adFldIsNullable
    TempVch.Fields.Append "AC_CODE", adVarChar, 6, adFldIsNullable
    TempVch.Fields.Append "G_CODE", adDouble, , adFldIsNullable
    TempVch.Fields.Append "VOU_ID", adDouble, , adFldIsNullable
    TempVch.Open , , adOpenKeyset, adLockOptimistic
    
If Fb_Press = 2 Or Fb_Press = 3 Then
    TempVch.AddNew
    If pmt_opn.Value = True Then
        TempVch!DR_CR = "D"
    Else
        TempVch!DR_CR = "C"
    End If
    TempVch!VOU_ID = 1
    TempVch.Update
    Set VchGrid.DataSource = TempVch
    VchGrid.ReBind
    VchGrid.Refresh
End If
End Sub
Sub VouPrn()
    Set VouPrnRec = Nothing
    Set VouPrnRec = New ADODB.Recordset
    
    VouPrnRec.Fields.Append "Voudt", adVarChar, 10, adFldIsNullable
    VouPrnRec.Fields.Append "VouType", adVarChar, 10, adFldIsNullable
    VouPrnRec.Fields.Append "AC_Code", adVarChar, 100, adFldIsNullable
    VouPrnRec.Fields.Append "Account", adVarChar, 100, adFldIsNullable
    VouPrnRec.Fields.Append "Dr_Cr", adVarChar, 2, adFldIsNullable
    VouPrnRec.Fields.Append "Narration", adVarChar, 1000, adFldIsNullable
    VouPrnRec.Fields.Append "Chequeno", adVarChar, 18, adFldIsNullable
    VouPrnRec.Fields.Append "Chequedt", adVarChar, 10, adFldIsNullable
    VouPrnRec.Fields.Append "CLBAL", adDouble, , adFldIsNullable
    VouPrnRec.Fields.Append "GCODE", adDouble, , adFldIsNullable
    VouPrnRec.Fields.Append "Amount", adDouble, , adFldIsNullable
    VouPrnRec.Fields.Append "Vouno", adVarChar, 20, adFldIsNullable
    VouPrnRec.Fields.Append "BANK_NAME", adVarChar, 200, adFldIsNullable
    VouPrnRec.Fields.Append "Branch", adVarChar, 200, adFldIsNullable
    VouPrnRec.Fields.Append "VOUID", adVarChar, 20, adFldIsNullable
    VouPrnRec.Fields.Append "BILLNO", adVarChar, 20, adFldIsNullable
    VouPrnRec.Fields.Append "BILLDATE", adVarChar, 10, adFldIsNullable
    VouPrnRec.Fields.Append "TOTAL_P", adDouble, , adFldIsNullable
    VouPrnRec.Fields.Append "TOTAL_R", adDouble, , adFldIsNullable
    VouPrnRec.Fields.Append "CBAcc", adVarChar, 100, adFldIsNullable
    VouPrnRec.Fields.Append "VNarration", adVarChar, 1000, adFldIsNullable
    VouPrnRec.Fields.Append "WordAmt", adVarChar, 100, adFldIsNullable
    VouPrnRec.Open , , adOpenKeyset, adLockOptimistic
End Sub
Sub Show_VchTotal()
    Dim MRec As ADODB.Recordset:    Dim LDebitAmt As Double:
    Dim LCreditAmt As Double:       Dim LVouAmt As Double
    Dim LPartyNetAmt As Double
    LDebitAmt = 0:    LCreditAmt = 0
    Set MRec = Nothing: Set MRec = New ADODB.Recordset: Set MRec = TempVch.Clone
    LVouNetAmt = 0
    LVouAmt = 0
    LPartyNetAmt = Val(TxtPartClBal.text & vbNullString)
    If MRec.RecordCount > 0 Then MRec.MoveFirst
    Do While Not MRec.EOF
        If MRec!AC_CODE = LBLPARTY.Caption Then
            If UCase(MRec!DR_CR & vbNullString) = LCreditChr Then
                LVouAmt = LVouAmt + Val(MRec!AMOUNT & vbNullString)
            ElseIf UCase(MRec!DR_CR & vbNullString) = LDebitChr Then
                LVouAmt = LVouAmt + (Val(MRec!AMOUNT & vbNullString) * -1)
            End If
        End If
        If UCase(MRec!DR_CR & vbNullString) = LCreditChr Then
            LCreditAmt = LCreditAmt + Val(MRec!AMOUNT & vbNullString)
            LVouNetAmt = LVouNetAmt + Val(MRec!AMOUNT & vbNullString)
        ElseIf UCase(MRec!DR_CR & vbNullString) = LDebitChr Then
            LDebitAmt = LDebitAmt + Val(MRec!AMOUNT & vbNullString)
            LVouNetAmt = LVouNetAmt - Val(MRec!AMOUNT & vbNullString)
        End If
        LVouNetAmt = Round(LVouNetAmt, 2)
        MRec.MoveNext
    Loop
    LPartyNetAmt = LPartyNetAmt + LVouAmt
    If LVouAmt > 0 Then
        LBLVOUAMT.Caption = Format(LVouAmt, "0.00") & " " & "Cr"
        LBLVOUAMT.BackColor = &HFFC0C0
    ElseIf LVouAmt < 0 Then
        LBLVOUAMT.Caption = Format(LVouAmt, "0.00") & " " & "Dr"
        LBLVOUAMT.BackColor = &H8080FF
    Else
        LBLVOUAMT.Caption = Format(LVouAmt, "0.00")
        LBLVOUAMT.BackColor = &HFFFFC0
    End If
    If LPartyNetAmt > 0 Then
        LBLNETAMT.Caption = Format(LPartyNetAmt, "#,##0.00") & " " & "Cr"
        LBLNETAMT.BackColor = &HFFC0C0
    ElseIf LPartyNetAmt < 0 Then
        LBLNETAMT.Caption = Format(LPartyNetAmt, "#,##0.00") & " " & "Dr"
        LBLNETAMT.BackColor = &H8080FF
    Else
        LBLNETAMT.Caption = Format(LPartyNetAmt, "0.00")
        LBLNETAMT.BackColor = &HFFFFC0
    End If
    TxtDebit.text = Format(LDebitAmt, "#,##0.00")
    txtCredit.text = Format(LCreditAmt, "#,##0.00")
    If LVouNetAmt > 0 Then
        TxtDiff.text = Format(LVouNetAmt, "#,##0.00") & " Cr"
    Else
       TxtDiff.text = Format(Abs(LVouNetAmt), "#,##0.00") & " Dr"
    End If
        
    Set MRec = Nothing
    If ComboVouType.ListIndex = 2 Then
        If LVouNetAmt < 0 Then
            AmtTxt.text = Format(Abs(LVouNetAmt), "0.00") & " Dr"
        Else
            AmtTxt.text = Format(Val(LVouNetAmt), "0.00") & " Cr"
        End If
    Else
        If LVouNetAmt < Val(0) Then
            AmtTxt.text = Format(Val(LVouNetAmt) * Val(-1), "0.00") & " Cr"
        Else
            AmtTxt.text = Format(Val(LVouNetAmt), "0.00") & " Dr"
        End If
    End If
End Sub
Sub VouPrinting()
    Dim PartyAdo As ADODB.Recordset:    Dim MYRS As ADODB.Recordset
    Dim AmountCr As Double:    Dim AmountDr As Double
    Dim PrtyAdrs As String:    Dim PrtyCity As String:    Dim LPhone  As String
    Call VouPrn
    AmountDr = 0
    If MVou_Type = "CV" Then
        mysql = " SELECT ACC.NAME,VOU.VOU_NO,VOU.VOU_TYPE,VT.DR_CR,VT.BANK_NAME,VT.BRANCH,VT.CHEQUE_NO,VT.CHEQUE_DT,VT.VOU_DT,VT.AC_CODE,VT.AMOUNT AS AMT,VT.NARRATION FROM VOUCHER AS VOU,VCHAMT AS VT,ACCOUNTM AS ACC "
        mysql = mysql & " WHERE ACC.COMPCODE=" & GCompCode & " AND ACC.COMPCODE=VOU.COMPCODE AND  VOU.VOU_NO='" & TxtVouNo.text & "' AND VOU.VOU_ID=VT.VOU_ID AND VT.AC_CODE=ACC.AC_CODE ORDER BY  VT.DR_CR DESC"
        Set MYRS = Nothing: Set MYRS = New ADODB.Recordset: MYRS.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not MYRS.EOF Then
            While Not MYRS.EOF
                If (MYRS!DR_CR = "D" And UCase(Left$(MYRS!VOU_NO, 4)) = "CASP") Or (MYRS!DR_CR = "C" And UCase(Left$(MYRS!VOU_NO, 4)) = "CASR") Then
                    AmountDr = AmountDr + Val(MYRS!AMT)
                    VouPrnRec.AddNew
                    VouPrnRec.Fields!Account = MYRS!NAME:           VouPrnRec.Fields!BRANCH = vbNullString
                    VouPrnRec.Fields!VOUDT = MYRS!VOU_DT:           VouPrnRec.Fields!VOUNO = MYRS!VOU_NO
                    VouPrnRec.Fields!AMOUNT = MYRS!AMT:             VouPrnRec.Fields!Total_P = 0:
                    VouPrnRec.Fields!Total_R = 0:                   VouPrnRec.Fields!VOUDT = Format(MYRS!VOU_DT, "dd/MM/yyyy")
                    VouPrnRec.Fields!VOUNO = MYRS!VOU_NO:           VouPrnRec.Fields!VNarration = Left$(MYRS!NARRATION, 100):
                    VouPrnRec.Fields!ChequeNo = vbNullString:       VouPrnRec.Fields!CHEQUEDT = vbNullString
                    VouPrnRec.Fields!CBAcc = DComboCashBank.text:   VouPrnRec.Fields!DR_CR = MYRS!DR_CR
                    VouPrnRec.Fields!NARRATION = vbNullString:      VouPrnRec.Fields!WORDAMT = vbNullString
                    VouPrnRec.Update
                ElseIf (MYRS!DR_CR = "D" And UCase(Left$(MYRS!VOU_NO, 2)) = "CP") Or (MYRS!DR_CR = "C" And UCase(Left$(MYRS!VOU_NO, 2)) = "CR") Then
                
                End If
                MYRS.MoveNext
            Wend
        End If
        VOULIST.ReportFileName = App.Path + "\Reports\CashVou.rpt"
    ElseIf MVou_Type = "BV" Then
        AmountDr = 0
        mysql = " SELECT ACC.NAME,VOU.VOU_NO,VOU.VOU_TYPE,VT.NARRATION,VT.DR_CR,VT.BANK_NAME,VT.BRANCH,VT.CHEQUE_NO,VT.CHEQUE_DT,VT.VOU_DT,VT.AC_CODE,VT.AMOUNT AS AMT FROM VOUCHER AS VOU,VCHAMT AS VT,ACCOUNTM AS ACC "
        mysql = mysql & " WHERE ACC.COMPCODE=" & GCompCode & " AND ACC.COMPCODE=VOU.COMPCODE AND  VOU.VOU_NO='" & TxtVouNo.text & "' AND VOU.VOU_ID=VT.VOU_ID AND VT.AC_CODE=ACC.AC_CODE ORDER BY VT.DR_CR"
        Set MYRS = Nothing: Set MYRS = New ADODB.Recordset: MYRS.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not MYRS.EOF Then
            While Not MYRS.EOF
                AmountDr = AmountDr + Val(MYRS!AMT)
                If (MYRS!DR_CR = "D" And UCase(Left$(MYRS!VOU_NO, 4)) = "BANP") Or (MYRS!DR_CR = "C" And UCase(Left$(MYRS!VOU_NO, 4)) = "BANR") Then
                    VouPrnRec.AddNew
                    VouPrnRec.Fields!Account = MYRS!NAME: VouPrnRec.Fields!BRANCH = MYRS!BRANCH
                    VouPrnRec.Fields!AMOUNT = MYRS!AMT: VouPrnRec.Fields!Total_P = 0: VouPrnRec.Fields!Total_R = 0
                    VouPrnRec.Fields!VOUDT = Format(MYRS!VOU_DT, "dd/MM/yyyy"): VouPrnRec.Fields!VOUNO = MYRS!VOU_NO
                    VouPrnRec.Fields!VNarration = MYRS!NARRATION: VouPrnRec.Fields!ChequeNo = MYRS!CHEQUE_NO: VouPrnRec.Fields!CHEQUEDT = Format(MYRS!CHEQUE_DT, "dd/MM/yyyy")
                    VouPrnRec.Fields!CBAcc = DComboCashBank.text: VouPrnRec.Fields!DR_CR = MYRS!DR_CR: VouPrnRec.Fields!NARRATION = ""
                    VouPrnRec.Fields!WORDAMT = vbNullString
                    VouPrnRec.Update
                End If
                MYRS.MoveNext
            Wend
        End If
    Else 'jv
        mysql = " SELECT ACC.NAME,VOU.VOU_NO,VOU.VOU_TYPE,VT.DR_CR,VT.BANK_NAME,VT.BRANCH,VT.CHEQUE_NO,VT.CHEQUE_DT,VT.VOU_DT,VT.AC_CODE,VT.AMOUNT AS AMT,VT.NARRATION FROM VOUCHER AS VOU,VCHAMT AS VT,ACCOUNTM AS ACC "
        mysql = mysql & " WHERE ACC.COMPCODE=" & GCompCode & " AND ACC.COMPCODE=VOU.COMPCODE AND VOU.VOU_NO='" & TxtVouNo.text & "' AND VOU.VOU_ID=VT.VOU_ID AND VT.AC_CODE=ACC.AC_CODE ORDER BY  VT.DR_CR DESC"
        Set MYRS = Nothing: Set MYRS = New ADODB.Recordset: MYRS.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not MYRS.EOF Then
            AmountCr = 0: AmountDr = 0
            While Not MYRS.EOF
                If (MYRS!DR_CR = "D") And (MVou_Type = "JV" Or MVou_Type = "DN") Then
                    If MVou_Type = "DN" Then
                        PrtyAdrs = vbNullString: PrtyCity = vbNullString
                        Set PartyAdo = Nothing
                        Set PartyAdo = New ADODB.Recordset
                        mysql = "SELECT AC_ADD,CITY FROM ACCOUNTD WHERE COMPODE = " & GCompCode & " AND AC_CODE ='" & MYRS!AC_CODE & "'"
                        PartyAdo.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                        If Not PartyAdo.EOF Then PrtyAdrs = PartyAdo!AC_ADD & "": PrtyCity = PartyAdo!City & ""
                        AmountCr = AmountCr + MYRS!AMT
                    Else
                        AmountDr = AmountDr + MYRS!AMT
                    End If
                    VouPrnRec.AddNew
                        VouPrnRec.Fields!Account = MYRS!NAME: VouPrnRec.Fields!BRANCH = MYRS!BRANCH
                        VouPrnRec.Fields!AMOUNT = 0: VouPrnRec.Fields!Total_P = MYRS!AMT: VouPrnRec.Fields!Total_R = 0
                        VouPrnRec.Fields!VOUDT = Format(MYRS!VOU_DT, "dd/MM/yyyy"): VouPrnRec.Fields!VOUNO = MYRS!VOU_NO
                        If MVou_Type = "DN" Then
                            VouPrnRec.Fields!VNarration = Left$(PrtyCity, 100)
                            VouPrnRec.Fields!CBAcc = Left$(PrtyAdrs, 100)
                        Else
                            VouPrnRec.Fields!VNarration = Left$(MYRS!NARRATION, 100)
                            VouPrnRec.Fields!CBAcc = ""
                        End If
                        VouPrnRec.Fields!ChequeNo = "": VouPrnRec.Fields!CHEQUEDT = ""
                        VouPrnRec.Fields!DR_CR = MYRS!DR_CR: VouPrnRec.Fields!NARRATION = TXT_NARR.text
                        VouPrnRec.Fields!WORDAMT = vbNullString
                    VouPrnRec.Update

                ElseIf MYRS!DR_CR = "C" And (MVou_Type = "JV" Or MVou_Type = "CN") Then
                    If MVou_Type = "CN" Then
                        PrtyAdrs = "": PrtyCity = ""
                        Set PartyAdo = Nothing
                        Set PartyAdo = New ADODB.Recordset
                        mysql = "SELECT AC_ADD,CITY FROM ACCOUNTD WHERE COMPCODE  = " & GCompCode & " AND AC_CODE ='" & MYRS!AC_CODE & "'"
                        PartyAdo.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                        If Not PartyAdo.EOF Then PrtyAdrs = PartyAdo!AC_ADD & "": PrtyCity = PartyAdo!City & ""
                        AmountCr = AmountCr + MYRS!AMT
                    Else
                        AmountCr = AmountCr + MYRS!AMT
                    End If
                    VouPrnRec.AddNew
                    VouPrnRec.Fields!Account = MYRS!NAME: VouPrnRec.Fields!BRANCH = MYRS!BRANCH
                    VouPrnRec.Fields!AMOUNT = 0: VouPrnRec.Fields!Total_P = 0: VouPrnRec.Fields!Total_R = MYRS!AMT
                    VouPrnRec.Fields!VOUDT = MYRS!VOU_DT: VouPrnRec.Fields!VOUNO = MYRS!VOU_NO
                    If MVou_Type = "CN" Then
                        VouPrnRec.Fields!VNarration = Left$(PrtyCity, 100)
                        VouPrnRec.Fields!CBAcc = Left$(PrtyAdrs, 100)
                    Else
                        VouPrnRec.Fields!VNarration = Left$(MYRS!NARRATION, 100)
                        VouPrnRec.Fields!CBAcc = ""
                    End If
                    VouPrnRec.Fields!ChequeNo = "": VouPrnRec.Fields!CHEQUEDT = ""
                    VouPrnRec.Fields!DR_CR = MYRS!DR_CR: VouPrnRec.Fields!NARRATION = TXT_NARR.text
                    VouPrnRec.Fields!WORDAMT = vbNullString
                    VouPrnRec.Update
                End If
                MYRS.MoveNext
            Wend
        End If
    End If
    If MVou_Type = "JV" Then
        mysql = "J O U R N A L   V O U C H E R "
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "JVVouPrn.rpt", 1)
    ElseIf MVou_Type = "BV" Then 'dn,cn
        mysql = "B A N K  V O U C H E R "
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "BVVouPrn.rpt", 1)
    ElseIf MVou_Type = "CV" Then 'dn,cn
        mysql = "C A S H  V O U C H E R "
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "CVVouPrn.rpt", 1)
    End If
    Set MYRS = Nothing
    Set MYRS = New ADODB.Recordset
    Set MYRS = VouPrnRec.Clone
    LPhone = "Ph: " & GCompPhoneO & ", " & GCompPhoneR & "," & GCompMobile & ""
    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource MYRS
    RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
    RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'" & mysql & "'"
    RDCREPO.FormulaFields.GetItemByName("OADD1").text = "'" & GCompanyAdd1 & "'"
    RDCREPO.FormulaFields.GetItemByName("OCITY").text = "'" & GCCity & "'"
    RDCREPO.FormulaFields.GetItemByName("OPHONE1").text = "'" & LPhone & "'"
    RDCREPO.FormulaFields.GetItemByName("SRVNO").text = "'" & GSrvRegNo & "'"
    
    CRViewer1.ZOrder
    CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Visible = True
    CRViewer1.ReportSource = RDCREPO
    CRViewer1.Zoom 1
    CRViewer1.ViewReport
End Sub
Public Sub Get_NewVouNo()
    Dim LVchSeries  As String
    If Fb_Press = 1 Then
        If MVou_Type = "JV" Then
            LVchSeries = "JNRL"
        ElseIf MVou_Type = "M" Then
            LVchSeries = "MRGN"
        ElseIf MVou_Type = "F" Then
            LVchSeries = "FIXD"
        ElseIf MVou_Type = "CV" Then
            If pmt_opn.Value Then
                LVchSeries = "CASP"
            Else
                LVchSeries = "CASR"
            End If
        ElseIf MVou_Type = "BV" Then
            If pmt_opn.Value Then
                LVchSeries = "BANP"
            Else
                LVchSeries = "BANR"
            End If
        End If
        TxtVouNo.text = Get_VouNo(LVchSeries, GFinYear)
    End If
End Sub

Private Sub Get_Bal()
Dim LBal As Double
Dim LACCID As Long
LACCID = Get_AccID(TxtCashBankCode.text)
LBal = Get_ClosingBal(LACCID, DTPicker1.Value)
If LBal > 0 Then
    TxtClBal.text = Format(LBal, "#,##0.00") & " Cr"
Else
    TxtClBal.text = Format(Abs(LBal), "#,##0.00") & " Dr"
End If
End Sub

Private Sub Get_PartyBal(lCCode As String, LCName As String)
Dim TRec As ADODB.Recordset
Dim LClBal As Double
Dim LACCID  As Long
LACCID = Get_AccID(lCCode)
LClBal = Get_ClosingBal(LACCID, DTPicker1.Value)

LBLPARTY.Caption = lCCode
LBLNAME.Caption = LCName
TxtPartClBal.text = LClBal
If LClBal > 0 Then
    LBLCLOSING.Caption = Format(Abs(LClBal), "#,##,###0.00") & " Cr"
    LBLCLOSING.BackColor = &HFFC0C0
ElseIf LClBal < 0 Then
    LBLCLOSING.Caption = Format(Abs(LClBal), "#,##,###0.00") & " Dr"
    LBLCLOSING.BackColor = &H8080FF
Else
    LBLCLOSING.Caption = Format(Abs(LClBal), "#,##,###0.00")
    LBLCLOSING.BackColor = &HFFFFC0
End If
Call Show_VchTotal
End Sub


