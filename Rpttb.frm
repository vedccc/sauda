VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form RPTTB 
   BackColor       =   &H00808080&
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Trial Balance"
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
         TabIndex        =   3
         Top             =   120
         Width           =   11655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Trial Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8205
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   11415
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3120
         TabIndex        =   4
         Top             =   4200
         Width           =   6135
         Begin VB.CheckBox ChkCashVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cash Voucher"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox ChkJVVOU 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Journal Voucher"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   120
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox ChkSetVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Settlement  Voucher"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   11
            Top             =   120
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox ChkShareVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Share Voucher"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox ChkShreeVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Shree Voucher"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   480
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox ChkBrokShVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Brok Share Voucher"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   8
            Top             =   480
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox ChkMarginVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Margin Vou"
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
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox ChkFixedVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fixed Margin Vou"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox ChkInterestVou 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Interest Voucher"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   5
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   7455
         Left            =   3000
         TabIndex        =   14
         Top             =   120
         Width           =   6615
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
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
            Left            =   0
            TabIndex        =   39
            Top             =   6600
            Width           =   6135
            Begin VB.CheckBox ChkNCDX 
               BackColor       =   &H00C0E0FF&
               Caption         =   "NCDX"
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
               Left            =   120
               TabIndex        =   44
               Top             =   120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1000
            End
            Begin VB.CheckBox ChkMCX 
               BackColor       =   &H00C0E0FF&
               Caption         =   "MCX"
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
               Left            =   1380
               TabIndex        =   43
               Top             =   120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1000
            End
            Begin VB.CheckBox ChkNSE 
               BackColor       =   &H00C0E0FF&
               Caption         =   "NSE"
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
               Left            =   2640
               TabIndex        =   42
               Top             =   120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1000
            End
            Begin VB.CheckBox ChkEQ 
               BackColor       =   &H00C0E0FF&
               Caption         =   "EQ"
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
               Left            =   3900
               TabIndex        =   41
               Top             =   120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1000
            End
            Begin VB.CheckBox ChkCmx 
               BackColor       =   &H00C0E0FF&
               Caption         =   "CMX"
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
               Left            =   5160
               TabIndex        =   40
               Top             =   120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1000
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   6375
            Begin VB.OptionButton OptGroup 
               BackColor       =   &H0080C0FF&
               Caption         =   "Group wise"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2100
               TabIndex        =   25
               Top             =   120
               Width           =   1695
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H0080C0FF&
               Caption         =   "Alphabatical"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   24
               Top             =   120
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H0080C0FF&
               Caption         =   "Four Coloumn"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4080
               TabIndex        =   23
               Top             =   120
               Width           =   1695
            End
         End
         Begin VB.ComboBox ComboVouType 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            ItemData        =   "Rpttb.frx":0000
            Left            =   1320
            List            =   "Rpttb.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CheckBox ChkVertical 
            BackColor       =   &H0080C0FF&
            Caption         =   "Vertical"
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
            Left            =   4320
            TabIndex        =   20
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CommandButton CANCEL_CMD 
            BackColor       =   &H00C0C0C0&
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
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
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   5880
            Width           =   1095
         End
         Begin VB.CommandButton OK_CMD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&OK"
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
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   5880
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H0080C0FF&
            Caption         =   "With Last  Date  Settlement"
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
            TabIndex        =   17
            Top             =   5400
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox ChkOpBal 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Include Op Bal"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            TabIndex        =   16
            Top             =   5400
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Summary"
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
            Left            =   3000
            TabIndex        =   15
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo Grpdb 
            Height          =   420
            Left            =   1320
            TabIndex        =   26
            Top             =   960
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   741
            _Version        =   393216
            ForeColor       =   16711680
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
         Begin vcDateTimePicker.vcDTP DtpFromDate 
            Height          =   375
            Left            =   1320
            TabIndex        =   27
            Top             =   1440
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
            Value           =   37680.7250462963
         End
         Begin vcDateTimePicker.vcDTP DtpToDate 
            Height          =   375
            Left            =   4200
            TabIndex        =   28
            Top             =   1440
            Visible         =   0   'False
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
            Value           =   37680.7250462963
         End
         Begin MSDataListLib.DataCombo FmlyCombo 
            Height          =   420
            Left            =   1320
            TabIndex        =   29
            Top             =   2400
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   741
            _Version        =   393216
            ForeColor       =   16711680
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
         Begin MSDataListLib.DataCombo HeadCombo 
            Height          =   420
            Left            =   1320
            TabIndex        =   30
            Top             =   3000
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   741
            _Version        =   393216
            ForeColor       =   16711680
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
            Left            =   4200
            TabIndex        =   31
            Top             =   3600
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
            Value           =   37680.7250462963
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Group"
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
            TabIndex        =   38
            Top             =   1043
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Upto Date"
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
            TabIndex        =   37
            Top             =   1500
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Option"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   1995
            Width           =   630
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "To Date"
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
            Left            =   3360
            TabIndex        =   35
            Top             =   1500
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
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
            Top             =   2483
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Party Head"
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
            TabIndex        =   33
            Top             =   3083
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Include Settlemet Entries Till Date"
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
            TabIndex        =   32
            Top             =   3600
            Width           =   3135
         End
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   855
      Left            =   16320
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   870
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
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   8460
      Left            =   120
      Top             =   720
      Width           =   11685
   End
End
Attribute VB_Name = "RPTTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecBal As ADODB.Recordset:      Dim RecAcc As ADODB.Recordset
Dim RecGroup As ADODB.Recordset:    Dim RecRpt As ADODB.Recordset:  Dim LFmlyCode  As String
Dim HeadRec As ADODB.Recordset
Dim FmlyRec As ADODB.Recordset
Sub SOPOTBCCBCCBC()
    Screen.MousePointer = 11: OK_CMD.Enabled = False
    If Mid(MFormat, 1, 13) = "Trial Balance" Or Mid(MFormat, 1, 11) = "MTM SUMMARY" Or MFormat = "Branch wise Trial Balance 3" Then
        Call NEW_TRIALBALANCE(0)
    ElseIf MFormat = "Partywise Outstanding" Then
        Call NEW_Outstanding(0)
    ElseIf MFormat = "Partywise Interest Collection" Then
        Call NEW_Outstanding(1)
    End If
    Screen.MousePointer = 0
    OK_CMD.Enabled = True
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error"
    Screen.MousePointer = 0: OK_CMD.Enabled = True
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check1.Caption = "Summary"
    Else
        Check1.Caption = "Detail"
    End If
End Sub

Private Sub ComboVouType_GotFocus()
    Sendkeys "%{down}"
End Sub


Private Sub DtpFromDate_Change()
vcDTP1.Value = DtpFromDate.Value
End Sub

Private Sub DtpFromDate_Validate(Cancel As Boolean)
vcDTP1.Value = DtpFromDate.Value
End Sub

Private Sub FmlyCombo_GotFocus()
    Sendkeys "%{down}"
End Sub

Private Sub Form_Paint()
MFormat = Label9.Caption
End Sub

Private Sub Grpdb_Validate(Cancel As Boolean)
If LenB(Grpdb.BoundText) > 0 Then LFmlyCode = Grpdb.BoundText
End Sub

Private Sub HeadCombo_GotFocus()
    Sendkeys "%{down}"
End Sub

Private Sub OK_CMD_Click()
    If DtpFromDate.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: DtpFromDate.SetFocus: Exit Sub
    If DtpFromDate.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: DtpFromDate.SetFocus: Exit Sub
    Call SOPOTBCCBCCBC
End Sub
Private Sub CANCEL_CMD_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CRViewer1.ZOrder
    Call Get_Selection(12)
    DtpFromDate.Value = Date
    vcDTP1.Value = Date
    DtpToDate.Value = GFinEnd
    Dim TRec As New ADODB.Recordset
    
    ChkNCDX.Value = 1:                  ChkMCX.Value = 1
    ChkNSE.Value = 1:                   ChkEQ.Value = 1:
    ChkNCDX.Visible = False:            ChkMCX.Visible = False
    ChkNSE.Visible = False:             ChkEQ.Visible = False:
    ChkCmx.Value = 1
    ChkCmx.Visible = False:
    
    mysql = "SELECT EXCODE FROM EXMAST WHERE COMPCODE = " & GCompCode & ""
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            If TRec!excode = "NCDX" Then
                ChkNCDX.Visible = True
            ElseIf TRec!excode = "MCX" Then
                ChkMCX.Visible = True
            ElseIf TRec!excode = "NSE" Then
                ChkNSE.Visible = True
            ElseIf TRec!excode = "EQ" Then
                ChkEQ.Visible = True
            ElseIf TRec!excode = "CMX" Then
                ChkCmx.Visible = True
            End If
            TRec.MoveNext
        Loop
    End If
    Set TRec = Nothing
    
    Label9.Caption = MFormat
    Check1.Visible = False
    If Mid(MFormat, 1, 13) = "Trial Balance" Then
        Set HeadRec = Nothing
        Set HeadRec = New ADODB.Recordset
        mysql = "SELECT HEADCODE,HEADNAME FROM PARTYHEAD WHERE COMPCODE =" & GCompCode & " ORDER BY HEADNAME"
        HeadRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not HeadRec.EOF Then
            HeadCombo.Visible = True
            Label6.Visible = True
            Set HeadCombo.RowSource = HeadRec
            HeadCombo.BoundColumn = "HEADCODE"
            HeadCombo.ListField = "HEADNAME"
        End If
        Set FmlyRec = Nothing
        Set FmlyRec = New ADODB.Recordset
        mysql = "SELECT FMLYID,FMLYCODE,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " ORDER BY FMLYNAME "
        FmlyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If FmlyRec.EOF Then
            FmlyCombo.Visible = False
        Else
            Set FmlyCombo.RowSource = FmlyRec
            FmlyCombo.BoundColumn = "FMLYCODE"
            FmlyCombo.ListField = "FMLYNAME"
        End If
        Set RecGroup = Nothing
        Set RecGroup = New ADODB.Recordset
        mysql = "SELECT CODE, G_NAME FROM AC_GROUP ORDER BY G_NAME"
        RecGroup.Open mysql, Cnn, adLockReadOnly, adLockReadOnly
        If Not RecGroup.EOF Then
            Set Grpdb.RowSource = RecGroup
            Grpdb.ListField = "G_NAME"
            Grpdb.BoundColumn = "CODE"
        End If
        mysql = "SELECT AC.AC_CODE, AC.NAME, AC.OP_BAL, AG.G_NAME, AG.G_CODE , AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG WHERE AC.COMPCODE=" & GCompCode & " AND AC.GCODE=AG.CODE AND AC.GRPCODE = AG.G_CODE ORDER BY AC.NAME, AG.G_NAME"
        Set RecAcc = Nothing: Set RecAcc = New ADODB.Recordset
        RecAcc.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    ElseIf MFormat = "Partywise Outstanding" Then
        OptGroup.Caption = "Amount Wise"
        mysql = "SELECT FMLYID ,FMLYCODE,FMLYNAME FROM ACCFMLY ORDER BY FMLYNAME"
        Set RecAcc = Nothing: Set RecAcc = New ADODB.Recordset
        RecAcc.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not RecAcc.EOF Then
            Set Grpdb.RowSource = RecAcc
            Grpdb.ListField = "FMLYNAME"
            Grpdb.BoundColumn = "FMLYCODE"
        End If
        Label1.Caption = "Branch"
        Label4.Visible = False
        ComboVouType.Visible = False
        ChkVertical.Visible = False
    ElseIf MFormat = "Party Ageing Analysis (Annexure 4)" Then
        Label4.Visible = False
        ComboVouType.Visible = False
        ChkVertical.Visible = False
        Frame3.Visible = False
        Label1.Visible = False
        Grpdb.Visible = False
    End If
    If Mid(MFormat, 1, 11) = "MTM SUMMARY" Then
            DtpFromDate.Visible = True
            DtpToDate.Visible = True
            DtpToDate.Value = Date
            If Date <= GFinEnd Then
                DtpFromDate.Value = GFinBegin
            End If
            Label3.Visible = True
            Label5.Visible = True
            Label3.Caption = "From Date"
            ChkVertical.Value = 0
    End If
    If Mid(MFormat, 1, 18) = "Partywise Interest" Then
            DtpFromDate.Visible = True
            DtpToDate.Visible = True
            DtpToDate.Value = Date
            If Date <= GFinEnd Then
                DtpFromDate.Value = GFinBegin
            End If
            Label3.Visible = True
            Label5.Visible = True
            Label3.Caption = "From Date"
            ChkVertical.Value = 0
            ChkNCDX.Value = 1:                  ChkMCX.Value = 1
            ChkNSE.Value = 1:                   ChkEQ.Value = 1:
            ChkNCDX.Visible = False:            ChkMCX.Visible = False
            ChkNSE.Visible = False:             ChkEQ.Visible = False:
            ChkCmx.Value = 1: ChkOpBal.Visible = False
            ChkOpBal.Value = 0: Check3.Visible = False
            ChkCmx.Visible = False:
            Label1.Visible = False: Grpdb.Visible = False
            Label3.Caption = "From Date"
            Label4.Visible = False: Label2.Visible = False: Label6.Visible = False
            Label7.Visible = False: FmlyCombo.Visible = False: HeadCombo.Visible = False
            ComboVouType.Visible = False: Check1.Visible = False: ChkVertical.Visible = False
            vcDTP1.Visible = False:              Frame7.Visible = False
            ChkCashVou.Value = 0: ChkJVVOU.Value = 0: ChkSetVou.Value = 0
            ChkShareVou.Value = 0: ChkBrokShVou.Value = 0: ChkShreeVou.Value = 0
            ChkMarginVou.Value = 0: ChkFixedVou.Value = 0: ChkInterestVou.Value = 1
    End If
    If Mid(MFormat, 1, 13) = "Trial Balance" Then
        ComboVouType.ListIndex = 1
        Me.Caption = Mid(MFormat, 1, 13)
        Frame1.Visible = True
        If Val(Right$(MFormat, 1)) = Val(1) Then
            DtpFromDate.Visible = False
            Label3.Visible = False
        ElseIf Val(Right$(MFormat, 1)) = 2 Then
            
            DtpFromDate.Visible = True

            Label3.Visible = True
        ElseIf Val(Right$(MFormat, 1)) = 3 Then
            DtpFromDate.Visible = True
            If Date <= GFinEnd Then DtpFromDate.Value = Date
            If Date >= GFinEnd Then DtpFromDate.Value = GFinEnd
            Label3.Visible = True
            ChkVertical.Value = 0
        ElseIf Val(Right$(MFormat, 1)) = 4 Then
            

        End If
        
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    GETMAIN.Label1.Caption = vbNullString
    If CRViewer1.Visible = True Then
        CRViewer1.Visible = False
        Cancel = 1
    Else
        GETMAIN.StatusBar1.Panels(1).text = vbNullString
        CNNERR = False
        Unload Me
    End If
End Sub
Sub NEW_TRIALBALANCE(GrpCode As Long)
    On Error GoTo Error1
    Dim LAcCode As String:          Dim LAcName  As String
    Dim LG_Name As String:          Dim LBal As Double
    Dim LBalance As Double:         Dim LDebitAmt As Double
    Dim LOpBalance As Double:       Dim LCreditAmt As Double
    Dim TRec As ADODB.Recordset
    Dim LGroup As String:           Dim LFromDate As Date
    Dim LToDate As Date:            Dim Balance As Double
    Dim LVouTypes  As String
    Dim LOp_Bal As Double
    
    If MFormat = "MTM SUMMARY" Then
        'ChkJVVOU.Value = 0
        'ChkCashVou.Value = 0
    End If

    LVouTypes = vbNullString
    If ChkCashVou.Value = 1 Then LVouTypes = "'CV','BV'"
    If ChkJVVOU.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'JV'"
    End If
    If ChkSetVou.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'S'"
    End If
    If ChkShareVou.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'H'"
    End If
    If ChkBrokShVou.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'B'"
    End If
    If ChkShreeVou.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'O'"
    End If
    If ChkMarginVou.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'M'"
    End If
    If ChkFixedVou.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'F'"
    End If
    If ChkInterestVou.Value = 1 Then
        If LenB(LVouTypes) > 0 Then LVouTypes = LVouTypes & ","
        LVouTypes = LVouTypes & "'I'"
    End If
    
'    Dim LToDate  As Date:           Dim LFromDate As Date
    If OptGroup.Value = True Then
        mysql = "SELECT AC.AC_CODE,AC.NAME,AC.OP_BAL,AG.G_NAME,AG.G_CODE,AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG "
        mysql = mysql & " WHERE AC.COMPCODE=" & GCompCode & " AND AC.GCODE=AG.CODE AND AC.GRPCODE = AG.G_CODE "
        If LenB(Grpdb.BoundText) > 0 Then mysql = mysql & " AND AC.GCODE = " & Grpdb.BoundText & ""
        If LenB(FmlyCombo.BoundText) > 0 Then mysql = mysql & " AND AC.ACCID IN (SELECT DISTINCT ACCID FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & "AND FMLYCODE ='" & FmlyCombo.BoundText & "')"
        mysql = mysql & " ORDER BY AG.G_NAME,AC.NAME"
    Else
        
        mysql = "SELECT AC.AC_CODE, AC.NAME, AC.OP_BAL, AG.G_NAME, AG.G_CODE , AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG "
        mysql = mysql & " WHERE AC.COMPCODE=" & GCompCode & " AND AC.GCODE=AG.CODE AND AC.GRPCODE = AG.G_CODE "
        If LenB(Grpdb.BoundText) > 0 Then mysql = mysql & " AND AC.GCODE = " & Grpdb.BoundText & ""
        If LenB(HeadCombo.BoundText) > 0 Then mysql = mysql & " AND AC.PTYHEAD = " & HeadCombo.BoundText & ""
        If LenB(FmlyCombo.BoundText) > 0 Then mysql = mysql & " AND AC.ACCID IN (SELECT DISTINCT ACCID FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & "AND FMLYCODE ='" & FmlyCombo.BoundText & "')"
        If GUniqClientId = "JMD-AHM" Or GUniqClientId = "JMD2-AHM" Then
            mysql = mysql & " ORDER BY AC.AC_CODE,AC.NAME, AG.G_NAME"
        Else
            mysql = mysql & " ORDER BY AC.NAME, AG.G_NAME"
        End If
    End If
    Set RecAcc = Nothing: Set RecAcc = New ADODB.Recordset
    RecAcc.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    GETMAIN.ProgressBar1.Max = RecAcc.RecordCount + 2
    GETMAIN.ProgressBar1.Value = 0
    GETMAIN.ProgressBar1.Visible = True
    Balance = 0
    If ChkVertical.Value = 1 Then
        Call RecSet
    Else
        Call RecTRlBal
    End If
    If MFormat = "Trial Balance 1" Then
        If RecAcc.RecordCount > 0 Then RecAcc.MoveFirst
        Do While Not RecAcc.EOF
            If ChkOpBal.Value = 0 Then
                LOp_Bal = 0
            Else
                LOp_Bal = Val(RecAcc!OP_BAL)
            End If
            
            If LOp_Bal <> 0 Then
                With RecRpt
                    .AddNew
                    !Balance = LOp_Bal
                    !AC_CODE = RecAcc!AC_CODE & vbNullString:                    !AC_NAME = RecAcc!NAME & vbNullString
                    !GroupName = RecAcc!g_name & vbNullString:
                    !DEBIT = 0:                 !CREDIT = 0:
                    .Update
                End With
            End If
            GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
            Call PERCENTAGE(GETMAIN.ProgressBar1.Max, GETMAIN.ProgressBar1.Value)
            RecAcc.MoveNext
        Loop
        GETMAIN.PERLBL.Caption = ""
    Else
        If MFormat = "Trial Balance 2" Then
            'LFromDate = DTPicker2.Value
            LToDate = DtpFromDate.Value
        ElseIf MFormat = "Trial Balance 3" Then
            LFromDate = GFinBegin
            'tdt = DtpFromDate.Value
            LToDate = DtpFromDate.Value + 1
        Else
            'fdt = DtpFromDate.Value
            LFromDate = DtpFromDate.Value
            'tdt = DtpToDate.Value
            LToDate = DtpToDate.Value + 1
        End If
        
        Do While Not RecAcc.EOF
            LGroup = RecAcc!g_name
            LDebitAmt = 0: LCreditAmt = 0
            If MFormat = "Trial Balance 2" Then    ''FOR THE PERIOD
                LBalance = Val(0)
            ElseIf MFormat = "MTM SUMMARY" Then
                LBalance = Val(0)
            Else                                    ''UP TO DATE
                If ComboVouType.ListIndex <> 0 Then
                    LBalance = Val(RecAcc!OP_BAL)
                    LOpBalance = Val(RecAcc!OP_BAL)
                Else
                    LBalance = 0
                    LOpBalance = 0
                    
                End If
            End If
            LAcCode = RecAcc!AC_CODE
            LAcName = RecAcc!NAME
            LG_Name = UCase(RecAcc!g_name & vbNullString)
            LBal = 0
            DoEvents
            If MFormat = "MTM SUMMARY" Then
                mysql = "SELECT SUM(Case DR_CR WHEN 'D' THEN  AMOUNT * -1 WHEN 'C' THEN  AMOUNT * 1  END)  AS AMT FROM VCHAMT WHERE COMPCODE= " & GCompCode & "  "
                If ChkNCDX.Value = 0 Then mysql = mysql & " AND BANK_NAME <>'NCDX' "
                If ChkMCX.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'MCX'  "
                If ChkNSE.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'NSE'  "
                If ChkEQ.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'EQ'   "
                If ChkCmx.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'CMX'   "
                mysql = mysql & " AND VOU_TYPE IN (" & LVouTypes & ")"
                mysql = mysql & " AND AC_CODE= '" & LAcCode & "' AND VOU_DT>='" & Format(LFromDate, "yyyy/MM/DD") & "' AND VOU_DT< '" & Format(LToDate, "YYYY/MM/DD") & "'"
                Set TRec = Nothing
                Set TRec = New ADODB.Recordset
                TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                If Not TRec.EOF Then LBal = IIf(IsNull(TRec!AMT), 0, TRec!AMT)
            Else
                If Option3.Value = True Then
                    mysql = "SELECT DR_CR,SUM(AMOUNT)  AS AMT FROM VCHAMT WHERE COMPCODE= " & GCompCode & "  AND AC_CODE= '" & LAcCode & "'"
                    mysql = mysql & " AND VOU_DT>='" & Format(GFinBegin, "yyyy/MM/DD") & "' "
                    If ChkNCDX.Value = 0 Then mysql = mysql & " AND BANK_NAME <>'NCDX' "
                    If ChkMCX.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'MCX'  "
                    If ChkNSE.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'NSE'  "
                    If ChkEQ.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'EQ'   "
                    If ChkCmx.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'CMX'   "
                    mysql = mysql & " AND VOU_TYPE IN (" & LVouTypes & ")"
                    mysql = mysql & " AND VOU_DT< '" & Format(LToDate, "YYYY/MM/DD") & "' GROUP BY DR_CR ORDER BY DR_CR"
                    Set TRec = Nothing
                    Set TRec = New ADODB.Recordset
                    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                    If Not TRec.EOF Then
                        If TRec!DR_CR = "D" Then
                            LDebitAmt = TRec!AMT
                        ElseIf TRec!DR_CR = "C" Then
                            LCreditAmt = TRec!AMT
                        End If
                        TRec.MoveNext
                        If Not TRec.EOF Then
                            If TRec!DR_CR = "D" Then
                                LDebitAmt = TRec!AMT
                            ElseIf TRec!DR_CR = "C" Then
                                LCreditAmt = TRec!AMT
                            End If
                        End If
                    End If
                    LBal = Net_DrCr(LAcCode, LToDate)
                Else
                    If Check3.Value = 1 Then
                        LBal = 0
                        If DateValue(vcDTP1.Value) <> DateValue(LToDate - 1) Then
                            mysql = "SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT*-1 WHEN 'C' THEN AMOUNT END) AS AMT FROM VCHAMT "
                            mysql = mysql & " WHERE COMPCODE= " & GCompCode & "  AND AC_CODE= '" & LAcCode & "'"
                            mysql = mysql & " AND VOU_TYPE IN (" & LVouTypes & ")"
                            mysql = mysql & " AND VOU_DT>='" & Format(GFinBegin, "yyyy/MM/DD") & "' "
                            mysql = mysql & " AND VOU_TYPE NOT IN ('S','H','B','O' )"
                            If ChkNCDX.Value = 0 Then mysql = mysql & " AND BANK_NAME <>'NCDX' "
                            If ChkMCX.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'MCX'  "
                            If ChkNSE.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'NSE'  "
                            If ChkEQ.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'EQ'   "
                            If ChkCmx.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'CMX'   "
                            mysql = mysql & " AND VOU_DT< '" & Format(LToDate, "YYYY/MM/DD") & "'"
                            
                            Set TRec = Nothing
                            Set TRec = New ADODB.Recordset
                            TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                            If Not TRec.EOF Then
                                LBal = IIf(IsNull(TRec!AMT), 0, TRec!AMT)
                            End If
                            
                            mysql = "SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT*-1 WHEN 'C' THEN AMOUNT END) AS AMT FROM VCHAMT "
                            mysql = mysql & " WHERE COMPCODE= " & GCompCode & "  AND AC_CODE= '" & LAcCode & "'"
                            mysql = mysql & " AND VOU_DT>='" & Format(GFinBegin, "yyyy/MM/DD") & "' "
                            If ChkNCDX.Value = 0 Then mysql = mysql & " AND BANK_NAME <>'NCDX' "
                            If ChkMCX.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'MCX'  "
                            If ChkNSE.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'NSE'  "
                            If ChkEQ.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'EQ'   "
                            If ChkCmx.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'CMX'   "
                            mysql = mysql & " AND VOU_TYPE  IN ('S','H','B','O' )"
                            mysql = mysql & " AND VOU_DT<= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
                            Set TRec = Nothing
                            Set TRec = New ADODB.Recordset
                            TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                            If Not TRec.EOF Then
                                LBal = LBal + IIf(IsNull(TRec!AMT), 0, TRec!AMT)
                            End If
                        Else
                            mysql = "SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT*-1 WHEN 'C' THEN AMOUNT END) AS AMT FROM VCHAMT "
                            mysql = mysql & " WHERE COMPCODE= " & GCompCode & "  AND AC_CODE= '" & LAcCode & "'"
                            mysql = mysql & " AND VOU_TYPE IN (" & LVouTypes & ")"
                            If ChkNCDX.Value = 0 Then mysql = mysql & " AND BANK_NAME <>'NCDX' "
                            If ChkMCX.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'MCX'  "
                            If ChkNSE.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'NSE'  "
                            If ChkEQ.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'EQ'   "
                            If ChkCmx.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'CMX'   "
                            mysql = mysql & " AND VOU_DT>='" & Format(GFinBegin, "yyyy/MM/DD") & "' "
                            mysql = mysql & " AND VOU_DT< '" & Format(LToDate, "YYYY/MM/DD") & "'"
                            Set TRec = Nothing
                            Set TRec = New ADODB.Recordset
                            TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                            If Not TRec.EOF Then
                                LBal = IIf(IsNull(TRec!AMT), 0, TRec!AMT)
                            End If
                        End If
                    Else
                        LBal = Net_DrCr(LAcCode, LToDate - 1)
                        mysql = "SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT*-1 WHEN 'C'THEN AMOUNT END) AS AMT  FROM VCHAMT "
                        mysql = mysql & " WHERE COMPCODE =" & GCompCode & "  AND VOU_DT='" & Format(LToDate - 1, "YYYY/MM/DD") & "'"
                        mysql = mysql & " AND VOU_TYPE IN (" & LVouTypes & ")"
                        If ChkNCDX.Value = 0 Then mysql = mysql & " AND BANK_NAME <>'NCDX' "
                        If ChkMCX.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'MCX'  "
                        If ChkNSE.Value = 0 Then mysql = mysql & "  AND BANK_NAME <>'NSE'  "
                        If ChkEQ.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'EQ'   "
                        If ChkCmx.Value = 0 Then mysql = mysql & "   AND BANK_NAME <>'CMX'   "
                        mysql = mysql & " AND VOU_TYPE NOT IN ('S','M') AND AC_CODE ='" & LAcCode & "'"
                        Set TRec = Nothing
                        Set TRec = New ADODB.Recordset
                        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                        If Not TRec.EOF Then
                            If Not IsNull(TRec!AMT) Then
                                LBal = LBal + TRec!AMT
                            End If
                        End If
                    End If
                End If
            End If
            LBalance = LBalance + LBal
            'If GRoundOff = "Y" Then
             '   LBalance = Round(LBalance)
            'Else
                LBalance = Round(LBalance, 2)
            'End If
            If Option3.Value = True Then
                With RecRpt
                    .AddNew
                    !OP_BALANCE = LOpBalance:       !Balance = LBalance
                    !AC_CODE = LAcCode:             !AC_NAME = LAcName
                    !GroupName = LG_Name:           !DEBIT = Val(LDebitAmt):
                    !CREDIT = Val(LCreditAmt):      .Update
                End With
            Else
                If LBalance <> 0 Then
                    If OptGroup.Value = True Then
                        If ChkVertical.Value = 1 Then
                            RecRpt.AddNew
                            RecRpt!GroupName = LG_Name:            RecRpt!AC_NAME = LAcName
                            RecRpt!AC_CODE = LAcCode
                            If LBalance < 0 Then
                                RecRpt!DEBIT = Abs(LBalance):
                                RecRpt!CREDIT = 0
                            Else
                                RecRpt!DEBIT = 0:
                                RecRpt!CREDIT = LBalance:
                            End If
                        Else
                            If LBalance < 0 Then
                                If RecRpt.EOF Then
                                    RecRpt.AddNew
                                    RecRpt!CAC_NAME = vbNullString:    RecRpt!CBALANCE = 0
                                    RecRpt!CREDITAC = vbNullString
                                Else
                                    RecRpt.MoveFirst
                                    RecRpt.Filter = adFilterNone
                                    If OptGroup.Value = True Then
                                        RecRpt.Filter = "GNAME='" & LGroup & "' AND DAC_NAME = ''"
                                    Else
                                        RecRpt.Filter = "DAC_NAME = ''"
                                    End If
                                    If RecRpt.EOF Then
                                        RecRpt.AddNew
                                        RecRpt!CAC_NAME = vbNullString:      RecRpt!CBALANCE = 0
                                        RecRpt!CREDITAC = vbNullString
                                    End If
                                End If
                                RecRpt!GNAME = LG_Name:            RecRpt!DAC_NAME = LAcName
                                RecRpt!DBALANCE = LBalance:        RecRpt!DEBITAC = LAcCode
                                RecRpt.Update
                            Else
                                If RecRpt.EOF Then
                                    RecRpt.AddNew:                      RecRpt!DAC_NAME = vbNullString
                                    RecRpt!DEBITAC = vbNullString:     RecRpt!DBALANCE = 0
                                Else
                                    RecRpt.MoveFirst
                                    RecRpt.Filter = adFilterNone
                                    If OptGroup.Value = True Then
                                        RecRpt.Filter = "GNAME='" & LGroup & "' AND CAC_NAME = ''"
                                    Else
                                        RecRpt.Filter = "CAC_NAME = ''"
                                    End If
                                    If RecRpt.EOF Then
                                        RecRpt.AddNew
                                        RecRpt!DAC_NAME = vbNullString:                            RecRpt!DBALANCE = 0
                                        RecRpt!DEBITAC = vbNullString:
                                    End If
                                End If
                                RecRpt!GNAME = LG_Name:         RecRpt!CAC_NAME = LAcName
                                RecRpt!CBALANCE = LBalance:     RecRpt!CREDITAC = LAcCode
                                RecRpt.Update
                            End If
                        End If
                        RecRpt.Update
                    Else
                        If ChkVertical.Value = 0 Then
                            If LBalance < 0 Then
                                If RecRpt.EOF Then
                                    RecRpt.AddNew
                                    RecRpt!CAC_NAME = vbNullString:    RecRpt!CBALANCE = 0
                                    RecRpt!CREDITAC = vbNullString
                                Else
                                    RecRpt.MoveFirst
                                    RecRpt.Filter = adFilterNone
                                    If OptGroup.Value = True Then
                                        RecRpt.Filter = "GNAME='" & LGroup & "' AND DAC_NAME = ''"
                                    Else
                                        RecRpt.Filter = "DAC_NAME = ''"
                                    End If
                                    If RecRpt.EOF Then
                                        RecRpt.AddNew
                                        RecRpt!CAC_NAME = vbNullString:      RecRpt!CBALANCE = 0
                                        RecRpt!CREDITAC = vbNullString
                                    End If
                                End If
                                RecRpt!GNAME = LG_Name:            RecRpt!DAC_NAME = LAcName
                                RecRpt!DBALANCE = LBalance:        RecRpt!DEBITAC = LAcCode
                                RecRpt.Update
                            Else
                                If RecRpt.EOF Then
                                    RecRpt.AddNew:                                    RecRpt!DAC_NAME = vbNullString
                                    RecRpt!DBALANCE = 0
                                    RecRpt!DEBITAC = vbNullString
                                Else
                                    RecRpt.MoveFirst
                                    RecRpt.Filter = adFilterNone
                                    If OptGroup.Value = True Then
                                        RecRpt.Filter = "GNAME='" & LGroup & "' AND CAC_NAME = ''"
                                    Else
                                        RecRpt.Filter = "CAC_NAME = ''"
                                    End If
                                    If RecRpt.EOF Then
                                        RecRpt.AddNew
                                        RecRpt!DAC_NAME = vbNullString:                            RecRpt!DBALANCE = 0
                                    End If
                                End If
                                RecRpt!GNAME = LG_Name:                                  RecRpt!CAC_NAME = LAcName
                                RecRpt!CBALANCE = LBalance:                              RecRpt!CREDITAC = LAcCode
                                RecRpt.Update
                            End If
                        ElseIf ChkVertical.Value = 1 Then
                            With RecRpt
                                .AddNew
                                !Balance = LBalance:             !AC_CODE = LAcCode
                                !AC_NAME = LAcName:              !GroupName = LG_Name
                                !DEBIT = 0:                      !CREDIT = 0:
                                .Update
                            End With
                        End If
                    End If
                End If
                GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
                Call PERCENTAGE(GETMAIN.ProgressBar1.Max, GETMAIN.ProgressBar1.Value)
            End If
            RecAcc.MoveNext
        Loop
        GETMAIN.PERLBL.Caption = vbNullString
    End If
    RecRpt.Filter = adFilterNone
    If Not RecRpt.EOF Then
        Set RDCREPO = Nothing
        If Option1.Value = True Then
            If ChkVertical.Value = 1 Then
                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "TRIAL1.Rpt", 1) 'alphabetical
            Else
                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "TRIAL21.Rpt", 1) 'alphabetical
            End If
        ElseIf Option3.Value = True Then
                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "TRIAL5.Rpt", 1) 'group wise
        Else
            If ChkVertical.Value = 1 Then
                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "TRIAL2.Rpt", 1) 'group wise
            Else
                If Check1.Value = 1 Then
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "Trial12_SUMMARY.rpt", 1) 'group wise summary
                Else
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "Trial12.rpt", 1) 'group wise detail
                End If
            End If
        End If
        If MFormat = "Trial Balance 1" Then
            RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Opening Trial Balance'"
        ElseIf MFormat = "Trial Balance 2" Then
            RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Trial Balance for " & LFromDate & " To " & " & LToDate & " '"
        ElseIf MFormat = "MTM SUMMARY" Then
            RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'MTM SUMMARY FROM  ' & '" & LFromDate & "' &' To '& '" & LToDate & "'"
        Else
            If LenB(FmlyCombo.BoundText) > 0 Then
                mysql = "'Trial Balance For  " & FmlyCombo.text & " as on " & DtpFromDate.Value & " '"
            Else
                mysql = "'Trial Balance as on " & DtpFromDate.Value & " '"
            End If
            
            RDCREPO.FormulaFields.GetItemByName("TITLE").text = mysql
        End If
        
        RDCREPO.DiscardSavedData
        If OptGroup.Value = True Then
            If ChkVertical.Value = 1 Then
                RecRpt.Sort = "GROUPNAME"
            Else
                RecRpt.Sort = "GNAME"
            End If
        End If
        RDCREPO.Database.SetDataSource RecRpt
        RDCREPO.FormulaFields.GetItemByName("ORG").text = "' " & GCompanyName & "'"
        RDCREPO.FormulaFields.GetItemByName("ADD1").text = "' " & MAdd1 & "'"
        RDCREPO.FormulaFields.GetItemByName("ADD2").text = "' " & GCCity & "'"
'        RDCREPO.FormulaFields.GetItemByName("RPTYPE").text = "'SUMMARY'"
        If Option1.Value = True Then
            CRViewer1.EnableGroupTree = False
        Else
            CRViewer1.EnableGroupTree = True
        End If
        CRViewer1.Width = CInt(GETMAIN.Width - 100)
        CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
        CRViewer1.Top = 0
        CRViewer1.Left = 0
        CRViewer1.ReportSource = RDCREPO
        CRViewer1.Visible = True
        CRViewer1.ViewReport
        GETMAIN.ProgressBar1.Visible = False
        Set RecRpt = Nothing
    Else
        MsgBox String(10, " ") & "No record found." & String(15, " "), vbExclamation, "Message"
        GETMAIN.ProgressBar1.Visible = False
        GETMAIN.PERLBL.Caption = vbNullString
    End If
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error"
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL.Caption = vbNullString
End Sub
Sub RecSet()    ''Sub Routine to Open Recordset Without Table
    Set RecRpt = Nothing
    Set RecRpt = New ADODB.Recordset
    RecRpt.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
    RecRpt.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
    RecRpt.Fields.Append "DEBIT", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "CREDIT", adDouble, adFldIsNullable
    RecRpt.Fields.Append "BALANCE", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "GROUPCODE", adInteger, , adFldIsNullable
    RecRpt.Fields.Append "GROUPNAME", adVarChar, 100, adFldIsNullable
    RecRpt.Open , , adOpenKeyset, adLockOptimistic
End Sub
Sub RecTRlBal()
    Set RecRpt = Nothing: Set RecRpt = New ADODB.Recordset
    RecRpt.Fields.Append "GName", adVarChar, 100, adFldIsNullable
    RecRpt.Fields.Append "DAC_NAME", adVarChar, 100, adFldIsNullable
    RecRpt.Fields.Append "CAC_NAME", adVarChar, 100, adFldIsNullable
    RecRpt.Fields.Append "DBALANCE", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "CBALANCE", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "DebitAc", adVarChar, 15, adFldIsNullable
    RecRpt.Fields.Append "CreditAc", adVarChar, 15, adFldIsNullable
    
    RecRpt.Open , , adOpenKeyset, adLockOptimistic
End Sub

Sub NEW_Outstanding(GrpCode As Long)
    On Error GoTo Error1
    Dim Account_Code As String
    Dim TRec As ADODB.Recordset
    Dim fdt As Date:
    Dim tdt As Date
    Dim MOpBal As Double
    Dim MBal As Double
    Dim LBal As Double
    Dim LName As String
    Call RecBalance
    fdt = GFinBegin
    Dim LToDate As Date
    Dim LFromDate As Date
    LFromDate = DtpFromDate.Value
    LToDate = DtpToDate.Value
    If GrpCode = 1 Then
        mysql = "SELECT A.AC_CODE, A.NAME, SUM(Case DR_CR WHEN 'D' THEN  AMOUNT * -1 WHEN 'C' THEN  AMOUNT * 1  END) as OP_BAL FROM ACCOUNTM AS A , VCHAMT AS V  "
        mysql = mysql & " WHERE A.COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND A.ACCID =V.ACCID AND V.VOU_TYPE = 'I' AND VOU_DT>='" & Format(LFromDate, "yyyy/MM/DD") & "' AND VOU_DT <= '" & Format(LToDate, "YYYY/MM/DD") & "' GROUP BY A.AC_CODE,A.NAME"
    Else
        If LenB(Grpdb.BoundText) > 0 Then
            mysql = "SELECT A.AC_CODE, A.NAME, A.OP_BAL,B.FMLYCODE,C.FMLYNAME FROM ACCOUNTM AS A ,ACCFMLYD AS B,ACCFMLY AS C "
            mysql = mysql & " WHERE A.COMPCODE =" & GCompCode & "  AND AnD C.FMLYID =B.FMLYID  AND A.ACCID =B.ACCID "
            mysql = mysql & " AND B.FMLYCODE ='" & Grpdb.BoundText & "'"
            mysql = mysql & " AND A.ACCID IN (SELECT DISTINCT ACCID FROM CTR_D WHERE COMPCODE =" & GCompCode & ") ORDER BY B.FMLYNAME,A.NAME"
        Else
            mysql = "SELECT A.AC_CODE, A.NAME, A.OP_BAL FROM ACCOUNTM AS A  "
            mysql = mysql & " WHERE A.COMPCODE =" & GCompCode & " "
            mysql = mysql & " AND A.ACCID IN (SELECT DISTINCT ACCID FROM CTR_D WHERE COMPCODE =" & GCompCode & ") ORDER BY A.NAME"
        End If
    End If
    
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            MOpBal = TRec!OP_BAL: MBal = MOpBal
            LName = TRec!NAME
            
            RecBal.AddNew
            RecBal!Balance = Val(MBal)
            If GrpCode <> 1 Then
                LBal = Net_DrCr(TRec!AC_CODE, LToDate)
                MBal = MBal + LBal
                If Grpdb.BoundText <> "" Then
                    RecBal!GNAME = TRec!FmlyNAME
                Else
                    RecBal!GNAME = TRec!FmlyNAME
                End If
                RecBal!Balance = Val(MBal) * -1
            End If
            RecBal!AC_NAME = LName
            
            RecBal.Update
            TRec.MoveNext
        Loop
        If OptGroup.Value = True Then
            RecBal.Sort = "BALANCE"
        End If
        GETMAIN.PERLBL.Caption = vbNullString
        
            Set RDCREPO = Nothing
        If GrpCode = 1 Then
            Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PTYINT.Rpt", 1) 'alphabetical
            RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Party wise Interest Collection' & '" & LFromDate & "' &' To '& '" & LToDate & "'"
        Else
            Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PTYOUT.Rpt", 1) 'alphabetical
            RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Party wise OutStanding Till  " & tdt - 1 & "'"
        End If
        RDCREPO.DiscardSavedData
        RDCREPO.Database.SetDataSource RecBal

        RDCREPO.FormulaFields.GetItemByName("ORG").text = "' " & GCompanyName & "'"
        RDCREPO.FormulaFields.GetItemByName("ADD1").text = "' " & MAdd1 & "'"
        RDCREPO.FormulaFields.GetItemByName("ADD2").text = "' " & GCCity & "'"

        CRViewer1.Width = CInt(GETMAIN.Width - 100)
        CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
        CRViewer1.Top = 0
        CRViewer1.Left = 0

        CRViewer1.ReportSource = RDCREPO
        CRViewer1.Visible = True
        CRViewer1.ViewReport

        GETMAIN.ProgressBar1.Visible = False
        Set RecBal = Nothing
    Else
        MsgBox String(10, " ") & "No record found." & String(15, " "), vbExclamation, "Message"
        GETMAIN.ProgressBar1.Visible = False
        GETMAIN.PERLBL.Caption = vbNullString
    End If
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error"
    GETMAIN.PERLBL.Caption = vbNullString
End Sub
Sub RecBalance()
    Set RecBal = Nothing: Set RecBal = New ADODB.Recordset
    RecBal.Fields.Append "GName", adVarChar, 100, adFldIsNullable
    RecBal.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
    RecBal.Fields.Append "BALANCE", adDouble, , adFldIsNullable
    RecBal.Fields.Append "DAYS1", adDouble, , adFldIsNullable
    RecBal.Fields.Append "DAYS2", adDouble, , adFldIsNullable
    RecBal.Fields.Append "DAYS3", adDouble, , adFldIsNullable
    RecBal.Fields.Append "DAYS4", adDouble, , adFldIsNullable
    RecBal.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub OptGroup_Click()
    If OptGroup.Value Then
        Check1.Visible = True
    End If
End Sub

Private Sub Option1_Click()
    If Option1.Value Then
        Check1.Visible = False
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value Then
        Check1.Visible = False
    End If
End Sub
