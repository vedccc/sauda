VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form GETQYRPT 
   BackColor       =   &H00808080&
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GETQYRPT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
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
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   13335
      Begin VB.Frame Frame4 
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
         ForeColor       =   &H00404080&
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   13095
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "General Ledger"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   12615
         End
      End
   End
   Begin VB.CheckBox ChkSelAll 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select All"
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
      Height          =   255
      Left            =   11880
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   0
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   13095
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
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
         Height          =   495
         Left            =   6600
         TabIndex        =   33
         Top             =   120
         Width           =   4215
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Account wise"
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
            TabIndex        =   35
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
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
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2640
            TabIndex        =   34
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   6495
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6375
         Begin VB.Frame Frame11 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   615
            Left            =   120
            TabIndex        =   52
            Top             =   5040
            Width           =   6135
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
               Height          =   405
               ItemData        =   "GETQYRPT.frx":000C
               Left            =   4560
               List            =   "GETQYRPT.frx":0016
               TabIndex        =   55
               Top             =   120
               Width           =   1455
            End
            Begin VB.CheckBox chktelegram 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Send Telegram"
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
               Left            =   2040
               TabIndex        =   54
               Top             =   120
               Width           =   1695
            End
            Begin VB.CheckBox CreateChk 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Export Report"
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
               TabIndex        =   53
               Top             =   120
               Width           =   1695
            End
         End
         Begin VB.CheckBox ChkInterest 
            BackColor       =   &H0080C0FF&
            Caption         =   "With Interest"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2520
            TabIndex        =   51
            Top             =   6120
            Width           =   1455
         End
         Begin VB.TextBox TxtBranchCode 
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
            Left            =   960
            TabIndex        =   48
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            TabIndex        =   29
            Top             =   2040
            Width           =   6135
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
               Left            =   4080
               TabIndex        =   49
               Top             =   1260
               Value           =   1  'Checked
               Width           =   1935
            End
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
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Width           =   6135
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
                  TabIndex        =   47
                  Top             =   840
                  Width           =   2055
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
                  TabIndex        =   46
                  Top             =   840
                  Width           =   2055
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
                  TabIndex        =   45
                  Top             =   840
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
                  TabIndex        =   44
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   2175
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
                  TabIndex        =   43
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   1695
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
                  TabIndex        =   42
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   1695
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
                  TabIndex        =   41
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   2175
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
                  TabIndex        =   40
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   1815
               End
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
                  TabIndex        =   39
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   1695
               End
            End
            Begin vcDateTimePicker.vcDTP vcDTP1 
               Height          =   375
               Left            =   2520
               TabIndex        =   36
               Top             =   1200
               Width           =   1500
               _ExtentX        =   2646
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
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Include Settle Entries Till"
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
               Top             =   1260
               Width           =   2295
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFC0&
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
            Height          =   375
            Left            =   1560
            TabIndex        =   26
            Top             =   4560
            Width           =   4695
            Begin VB.OptionButton OptShareClient 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Client Wise"
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
               TabIndex        =   28
               Top             =   80
               Width           =   1455
            End
            Begin VB.OptionButton OptShareExchange 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Exhange Wise"
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
               TabIndex        =   27
               Top             =   80
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   3960
            Width           =   6135
            Begin VB.OptionButton OptSaudaWise 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Daily Sauda Wise"
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
               TabIndex        =   25
               Top             =   120
               Width           =   1935
            End
            Begin VB.OptionButton OptDailySettle 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Daily Exch Wise"
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
               Width           =   1815
            End
            Begin VB.OptionButton OptWeekSettle 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Week Wise "
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
               Left            =   2280
               TabIndex        =   23
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.CommandButton OK_CMD 
            BackColor       =   &H00FFFFC0&
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
            Height          =   435
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   6000
            Width           =   1095
         End
         Begin VB.CommandButton CANCEL_CMD 
            BackColor       =   &H00FFFFC0&
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
            Height          =   430
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   6000
            Width           =   975
         End
         Begin VB.CheckBox ChkNewPage 
            BackColor       =   &H0080C0FF&
            Caption         =   "New Page After Each A/c"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   6120
            Width           =   2415
         End
         Begin VB.Frame Frame1 
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
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   6135
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
               TabIndex        =   50
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
               TabIndex        =   10
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
               TabIndex        =   9
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
               TabIndex        =   8
               Top             =   120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1000
            End
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
               TabIndex        =   7
               Top             =   120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1000
            End
         End
         Begin vcDateTimePicker.vcDTP VcDtpFromDate 
            Height          =   375
            Left            =   720
            TabIndex        =   5
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
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
         Begin vcDateTimePicker.vcDTP VcDtpToDate 
            Height          =   375
            Left            =   4800
            TabIndex        =   6
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MSDataListLib.DataCombo DComboFmly 
            Height          =   420
            Left            =   2160
            TabIndex        =   14
            Top             =   1200
            Width           =   4095
            _ExtentX        =   7223
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
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "Voucher Type"
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
            TabIndex        =   32
            Top             =   1680
            Width           =   6135
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "Sharing Entry"
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
            Left            =   0
            TabIndex        =   31
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "Settlement Entry"
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
            TabIndex        =   30
            Top             =   3600
            Width           =   6135
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080C0FF&
            Caption         =   "Family"
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
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1290
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            Height          =   255
            Left            =   4440
            TabIndex        =   12
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "From"
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
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView ListViewAcc 
         Height          =   5775
         Left            =   6600
         TabIndex        =   3
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   14095
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         X1              =   0
         X2              =   3840
         Y1              =   -120
         Y2              =   -120
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7020
      Left            =   75
      Top             =   960
      Width           =   13365
   End
End
Attribute VB_Name = "GETQYRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AccRec As ADODB.Recordset
Dim AccRecDetail As ADODB.Recordset
Dim RecRpt As ADODB.Recordset
Dim FmlyRec As ADODB.Recordset

Private Sub ChkSelAll_Click()
    Dim I As Integer
    For I = 1 To ListViewAcc.ListItems.Count
        If ChkSelAll.Value = 1 Then
            ListViewAcc.ListItems.Item(I).Checked = True
        Else
            ListViewAcc.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub ComboVouType_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub DComboFmly_Validate(Cancel As Boolean)
Dim TRec As ADODB.Recordset
Dim LFmlyID As Long
If LenB(DComboFmly.text) <> 0 Then
    TxtBranchCode.text = DComboFmly.BoundText
    LFmlyID = Get_Fmlyid(TxtBranchCode.text)
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    mysql = "SELECT AC.ACCID,AC.AC_CODE, AC.NAME,  AC.OP_BAL, AG.G_NAME, AG.CODE, AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG, ACCFMLYD AS AD  "
    mysql = mysql & " WHERE AC.COMPCODE=" & GCompCode & "  AND  AC.ACCID =AD.ACCID  AND AC.GCODE=AG.CODE AND AC.GRPCODE=AG.G_CODE"
    mysql = mysql & " AND AD.FMLYID =" & LFmlyID & " ORDER BY AC.NAME"
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not TRec.EOF Then
        ListViewAcc.Visible = True
        ListViewAcc.Checkboxes = True
        ListViewAcc.ListItems.Clear
        Do While Not TRec.EOF
            ListViewAcc.ListItems.Add , , UCase(TRec!NAME)
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , TRec!AC_CODE
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , TRec!ACCID
            TRec.MoveNext
        Loop
    End If
    Set TRec = Nothing
    ListViewAcc.ToolTipText = "Press F2 to Select all AND F3 to UnSELECT."
End If
End Sub

Private Sub Form_Activate()
    MFormat = Label7.Caption
End Sub

Private Sub Form_Click()
MFormat = Label7.Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
    MFormat = Label7.Caption
End Sub


Private Sub Form_Paint()
    MFormat = Label7.Caption
End Sub
Private Sub ListViewAcc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim MBOOL As Boolean
    Dim I As Integer
    If ListViewAcc.SelectedItem.SubItems(2) = "Group" Then    ''TO CHECK/UNCHECK ALL A/C FOR THE SELECTED GROUP
        If ListViewAcc.SelectedItem.Checked = True Then
            MBOOL = True
        Else
            MBOOL = False
        End If

        I = Item.Index + Val(1)
        While Not ListViewAcc.ListItems(I).ListSubItems(2).text = "Group"
            ListViewAcc.ListItems(I).Checked = MBOOL
            I = I + Val(1)
        Wend
    End If
End Sub
Private Sub ListViewAcc_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    For I = 1 To ListViewAcc.ListItems.Count
        If KeyCode = 118 Then
            ListViewAcc.ListItems(I).Checked = True
        ElseIf KeyCode = 119 Then
            ListViewAcc.ListItems(I).Checked = False
        End If
    Next
End Sub
Private Sub OK_CMD_Click()
    If VcDtpFromDate.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: VcDtpFromDate.SetFocus: Exit Sub
    If VcDtpFromDate.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: VcDtpFromDate.SetFocus: Exit Sub
    
    If VcDtpToDate.Visible Then
        If VcDtpToDate.Value < VcDtpFromDate.Value Then MsgBox "Invalid date range.", vbCritical: VcDtpFromDate.SetFocus: Exit Sub
        If VcDtpToDate.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: VcDtpToDate.SetFocus: Exit Sub
        If VcDtpToDate.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: VcDtpToDate.SetFocus: Exit Sub
        If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
        If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before  financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    End If
    If MFormat = "General Ledger" Or MFormat = "Detailed Ledger" Then
        If OptDailySettle.Value = True Then
            Call Ledger_NEW
        Else
            Call Ledger_NEW
        End If
    ElseIf MFormat = "Invoice wise Ledger" Then
        Call INV_WS_GL
    End If
End Sub
Private Sub CANCEL_CMD_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Dim TRec As ADODB.Recordset
    
    Combo1.ListIndex = 0
    CRViewer1.ZOrder
    Call Get_Selection(12)
    TxtBranchCode.text = vbNullString
    Label1.Caption = "To Date"
    If Date > GFinEnd Then
        VcDtpToDate.Value = GFinEnd:
        vcDTP1.Value = GFinEnd
    Else
        VcDtpToDate.Value = Date:
        vcDTP1.Value = Date
    End If
    
    VcDtpFromDate.Value = GFinBegin
    ChkNCDX.Value = 1:                  ChkMCX.Value = 1
    ChkNSE.Value = 1:                   ChkEQ.Value = 1:
    ChkNCDX.Visible = False:            ChkMCX.Visible = False
    ChkNSE.Visible = False:             ChkEQ.Visible = False:
    ChkCmx.Value = 1
    OptDailySettle.Value = True
    
    OptShareExchange.Value = True
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
    If MFormat = "Invoice wise Ledger" Then
        Option3.Visible = False
        Option4.Visible = False
    End If
    Set FmlyRec = Nothing
    Set FmlyRec = New ADODB.Recordset
    mysql = "SELECT FMLYID,FMLYCODE,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " ORDER BY FMLYNAME"
    FmlyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not FmlyRec.EOF Then
        Set DComboFmly.RowSource = FmlyRec
        DComboFmly.ListField = "FmlyName"
        DComboFmly.BoundColumn = "FmlyCode"
    End If
    Set FmlyRec = Nothing
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC.ACCID,AC.AC_CODE,AC.NAME,AC.OP_BAL,AG.G_NAME,AG.CODE,AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG "
    mysql = mysql & " WHERE AC.COMPCODE=" & GCompCode & " AND AC.GCODE=AG.CODE AND AC.GRPCODE=AG.G_CODE ORDER BY AC.NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        ListViewAcc.Visible = True
        ListViewAcc.Checkboxes = True
        ListViewAcc.ListItems.Clear
        Do While Not AccRec.EOF
            ListViewAcc.ListItems.Add , , UCase(AccRec!NAME)
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , AccRec!AC_CODE
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , AccRec!ACCID
            
            AccRec.MoveNext
        Loop
        AccRec.MoveFirst
    End If
    Set AccRecDetail = Nothing
    Set AccRecDetail = New ADODB.Recordset
    mysql = "SELECT ACCID,AC_CODE,NAME,ADD1,CITY,PANNO FROM ACCOUNTD ORDER BY AC_CODE "
    AccRecDetail.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    ListViewAcc.ToolTipText = "Press F2 to Select all and F3 to UnSelect."
    Me.Caption = MFormat:          VcDtpFromDate.Visible = True
    VcDtpToDate.Visible = True:     Label1.Visible = True
    OptWeekSettle.Value = True
    Label4.Visible = True:
    If GDailyBill = False Then
        OptWeekSettle.Value = True
        OptShareClient.Value = True
    
    Else
        OptDailySettle.Value = True
        OptShareExchange.Value = True
    End If
    
    VcDtpFromDate.Value = GFinBegin
    Frame2.Visible = True

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
Sub RecSet()    ''Sub Routine to Open Recordset Without Table
    Set RecRpt = Nothing
    Set RecRpt = New ADODB.Recordset
    RecRpt.Fields.Append "QAC_DT", adDate, , adFldIsNullable
    RecRpt.Fields.Append "DEBIT", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "CREDIT", adDouble, adFldIsNullable
    RecRpt.Fields.Append "BALANCE", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "NARRATION", adVarChar, 500, adFldIsNullable
    RecRpt.Fields.Append "CHEQUE_NO", adVarChar, 10, adFldIsNullable
    RecRpt.Fields.Append "CHEQUE_DT", adVarChar, 20, adFldIsNullable
    RecRpt.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
    RecRpt.Fields.Append "BILL_NO", adVarChar, 11, adFldIsNullable
    RecRpt.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
    RecRpt.Fields.Append "REFRENCE", adVarChar, 30, adFldIsNullable
    RecRpt.Fields.Append "SM_NAME", adVarChar, 200, adFldIsNullable
    RecRpt.Fields.Append "GROUP", adVarChar, 50, adFldIsNullable
    RecRpt.Fields.Append "TYPE", adVarChar, 5, adFldIsNullable
    RecRpt.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
    RecRpt.Fields.Append "OP_BALANCE", adDecimal, , adFldIsNullable
    RecRpt.Fields.Append "INV_NO", adVarChar, 20, adFldIsNullable
    RecRpt.Fields.Append "INV_DT", adDate, , adFldIsNullable
    RecRpt.Fields.Append "BILL_DT", adDate, , adFldIsNullable
    RecRpt.Fields.Append "G_TYPE", adVarChar, 1, adFldIsNullable
    RecRpt.Fields.Append "G_CAT", adVarChar, 30, adFldIsNullable
    RecRpt.Fields.Append "MarginAmt", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "ADDRESS", adVarChar, 200, adFldIsNullable
    RecRpt.Fields.Append "CINNO", adVarChar, 30, adFldIsNullable
    RecRpt.Fields.Append "PANNO", adVarChar, 20, adFldIsNullable
    RecRpt.Fields.Append "INTRATE1", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "INTRATE2", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "PREVBAL", adDouble, , adFldIsNullable
    RecRpt.Fields.Append "INTAMT", adDouble, , adFldIsNullable
    RecRpt.Open , , adOpenKeyset, adLockOptimistic
End Sub
Private Sub Option3_LostFocus()     ''GROUP WISE LEDGER
    Dim LGroupCode As Long
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC.ACCID,AC.AC_CODE, AC.NAME AS NAME, AC.OP_BAL, AG.G_NAME AS G_NAME, AG.CODE FROM ACCOUNTM AS AC, AC_GROUP AS AG "
    mysql = mysql & " WHERE AC.COMPCODE=" & GCompCode & " AND  AC.GRPCODE=AG.CODE ORDER BY AG.G_NAME, AC.NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    AccRec.MoveFirst
    ListViewAcc.Visible = True
    ListViewAcc.Checkboxes = True
    ListViewAcc.ListItems.Clear
    Do While Not AccRec.EOF
        If Not Val(LGroupCode) = Val(AccRec!Code) Then
            ListViewAcc.ListItems.Add , , AccRec!g_name
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , "G"
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , "Group"
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ForeColor = Val(255)
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).Bold = True
            LGroupCode = AccRec!Code
        End If
        ListViewAcc.ListItems.Add , , AccRec!NAME
        ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , AccRec!AC_CODE
        ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , AccRec!ACCID
        AccRec.MoveNext
    Loop
    AccRec.MoveFirst
End Sub
Private Sub Option4_LostFocus()     ''A/C. WISE LEDGER
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC.ACCID,AC.AC_CODE, AC.NAME AS NAME, AC.OP_BAL, UPPER(AG.G_NAME) AS G_NAME, AG.CODE FROM ACCOUNTM AS AC, AC_GROUP AS AG WHERE ac.COMPCODE=" & GCompCode & " AND AC.GRPCODE=AG.CODE ORDER BY AC.NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    If Not AccRec.EOF Then
        AccRec.MoveFirst
        ListViewAcc.Visible = True
        ListViewAcc.Checkboxes = True
        ListViewAcc.ListItems.Clear
        Do While Not AccRec.EOF
            ListViewAcc.ListItems.Add , , AccRec!NAME
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , AccRec!AC_CODE
            ListViewAcc.ListItems(ListViewAcc.ListItems.Count).ListSubItems.Add , , AccRec!ACCID
            AccRec.MoveNext
        Loop
        AccRec.MoveFirst
    End If
End Sub
Sub INV_WS_GL()
Dim I As Integer:       Dim LOp_Bal As Double:      Dim LAddress As String
Dim LCINNo As String:   Dim LPANNO As String:       Dim LNarration As String
Dim LBal As Double:     Dim TRec As ADODB.Recordset
Dim LAC_CODE As String
    
    On Error GoTo Error1
    Call RecSet
    GETMAIN.ProgressBar1.Visible = True
    GETMAIN.ProgressBar1.Max = ListViewAcc.ListItems.Count + 1
    GETMAIN.ProgressBar1.Value = 0

    Screen.MousePointer = 11: OK_CMD.Enabled = False

    For I = 1 To ListViewAcc.ListItems.Count
        If ListViewAcc.ListItems(I).ListSubItems(2).text = "Group" Then I = I + Val(1)
        If ListViewAcc.ListItems(I).Checked = True Then
            LAC_CODE = ListViewAcc.ListItems(I).SubItems(1)
            LOp_Bal = 0
            AccRec.MoveFirst
            AccRec.Find "AC_CODE='" & LAC_CODE & "'", , adSearchForward
            AccRecDetail.MoveFirst
            AccRecDetail.Find "AC_CODE='" & LAC_CODE & "'", , adSearchForward
            If Not AccRecDetail.EOF Then
                LAddress = Trim(AccRecDetail!ADD1 & " " & AccRecDetail!City)
                LCINNo = vbNullString
                LPANNO = Trim(IIf(IsNull(AccRecDetail!PANNO), "", AccRecDetail!PANNO))
            Else
                LAddress = vbNullString: LCINNo = vbNullString: LPANNO = vbNullString
            End If
            If Val(AccRec!OP_BAL & "") <> 0 Then LOp_Bal = Val(AccRec!OP_BAL & "")
            LBal = Net_DrCr(LAC_CODE, VcDtpFromDate.Value)
            LOp_Bal = LOp_Bal + LBal
            If LOp_Bal <> 0 Then
                With RecRpt
                    .AddNew
                    !AC_CODE = AccRec!AC_CODE & "":                    !AC_NAME = AccRec!NAME & ""
                    !Address = LAddress:                               !CINNO = LCINNo
                    !PANNO = LPANNO:                                   !INV_NO = "Opening Balance"
                    !Balance = Val(LOp_Bal)
                    If LOp_Bal < 0 Then
                        !DEBIT = (-1) * Val(LOp_Bal):          !CREDIT = 0
                    Else
                        !DEBIT = 0:                            !CREDIT = Val(LOp_Bal)
                    End If
                    !BILL_NO = ""
                    .Update
                End With
            End If
            mysql = "SELECT ACC.NAME AS AC_NAME,V.SAUDA,VT.VOU_NO AS V_NO,V.VOU_NO,VT.AC_CODE,VT.DR_CR,VT.AMOUNT,VT.NARRATION,VT.CHEQUE_NO,VT.CHEQUE_DT,"
            mysql = mysql & " VT.VOU_DT, V.INVNO FROM VCHAMT AS VT, ACCOUNTM AS ACC,VOUCHER AS V WHERE ACC.COMPCODE=" & GCompCode & " "
            mysql = mysql & " AND VT.VOU_DT >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND VT.VOU_DT <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "' AND VT.ACCID = ACC.ACCID "
            mysql = mysql & " AND VT.AC_CODE ='" & LAC_CODE & "' AND V.VOU_ID = VT.VOU_ID "
            mysql = mysql & " AND V.VOU_TYPE NOT IN ('S','M','B','H') ORDER BY VT.VOU_DT, VT.VOU_NO "
            Set TRec = Nothing
            Set TRec = New ADODB.Recordset
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                TRec.MoveFirst
                While Not TRec.EOF
                    LNarration = TRec!NARRATION & ""
                    If TRec!DR_CR = "D" Then
                        With RecRpt
                            .AddNew
                            !INV_NO = TRec!VOU_NO & "":                          !AC_CODE = AccRec!AC_CODE & ""
                            !AC_NAME = AccRec!NAME & "":                         !Address = LAddress
                            !CINNO = LCINNo:                                     !PANNO = LPANNO
                            !QAC_DT = TRec!VOU_DT
                            If TRec!DR_CR = "D" Then
                                !DEBIT = Val(TRec!AMOUNT & ""):                  !CREDIT = 0
                                LOp_Bal = LOp_Bal - Val(TRec!AMOUNT & "")
                            Else
                                !DEBIT = 0:                                       !CREDIT = Val(TRec!AMOUNT & "")
                                LOp_Bal = LOp_Bal + Val(TRec!AMOUNT & "")
                            End If
                            !NARRATION = LNarration:                           !CHEQUE_NO = TRec!CHEQUE_NO & ""
                            !CHEQUE_DT = TRec!CHEQUE_DT:                       !Balance = Val(LOp_Bal & "")
                            !BILL_NO = TRec!invno & ""
                            'If Len(TRec!INVDATE & "") = Val(10) Then !INV_DT = TRec!INVDATE & ""
                            !SM_NAME = TRec!Sauda & ""
                            .Update
                        End With
                    End If
                    TRec.MoveNext
                Wend
            End If
            '--invoice
            mysql = "SELECT ACC.NAME AS AC_NAME, I.INVNO AS Vou_NO, I.PARTY AS AC_CODE, I.INVDATE AS VOU_DT, SUM(I.BILLAMT) AS AMOUNT FROM INV_D AS I, "
            mysql = mysql & " ACCOUNTM AS ACC WHERE ACC.COMPCODE=" & GCompCode & " AND I.STDATE >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' "
            mysql = mysql & " AND I.STDATE <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "' AND I.ACCID  = ACC.ACCID  AND I.PARTY ='" & LAC_CODE & "'"
            mysql = mysql & " AND I.INVNO > 0 GROUP BY ACC.NAME, I.INVNO,I.PARTY,I.INVDATE "
            Set TRec = Nothing
            Set TRec = New ADODB.Recordset
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                TRec.MoveFirst
                While Not TRec.EOF
                    If TRec!AMOUNT < Val(0) Then
                        LOp_Bal = LOp_Bal - Val(TRec!AMOUNT & "")
                        With RecRpt
                            .AddNew
                            !INV_NO = TRec!VOU_NO & "":                          !AC_CODE = AccRec!AC_CODE & ""
                            !AC_NAME = AccRec!NAME & "":                         !Address = LAddress
                            !CINNO = LCINNo:                                     !PANNO = LPANNO
                            !QAC_DT = TRec!VOU_DT
                            If TRec!AMOUNT < Val(0) Then
                                LOp_Bal = LOp_Bal - Val(TRec!AMOUNT & "")
                                !DEBIT = Val(TRec!AMOUNT & "") * Val(-1)
                                !CREDIT = 0
                            Else
                                LOp_Bal = LOp_Bal + Val(TRec!AMOUNT & "")
                                !DEBIT = 0
                                !CREDIT = Val(TRec!AMOUNT & "")
                            End If
                            !NARRATION = "Invoice":                            !CHEQUE_NO = ""
                            !Balance = Val(LOp_Bal & ""):                      !BILL_NO = ""
                            !SM_NAME = ""
                            .Update
                        End With
                    End If
                    TRec.MoveNext
                Wend
            End If
            Set TRec = Nothing
        End If
    Next
    If RecRpt.RecordCount <> 0 Then
        Set RDCREPO = Nothing
        Set RDCREPO = New CRAXDRT.report
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PRTSTTIW.RPT", 1)
        RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
        RDCREPO.FormulaFields.GetItemByName("ADD1").text = "'For " & CStr(Format(VcDtpFromDate.Value, "yyyy/MM/dd")) & " To " & CStr(Format(VcDtpToDate.Value, "yyyy/MM/dd")) & "'"
        RDCREPO.Database.SetDataSource RecRpt
        CRViewer1.Width = CInt(GETMAIN.Width - 100)
        CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
        CRViewer1.Top = 0
        CRViewer1.Left = 0
        CRViewer1.ReportSource = RDCREPO
        CRViewer1.Visible = True
        CRViewer1.ViewReport
        GETMAIN.ProgressBar1.Visible = False
        GETMAIN.PERLBL = vbNullString
        Set RecRpt = Nothing
    End If
    GETMAIN.PERLBL.Caption = vbNullString
    Screen.MousePointer = 0
    OK_CMD.Enabled = True
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    
    GETMAIN.ProgressBar1.Value = 0: OK_CMD.Enabled = True
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    Screen.MousePointer = 0
End Sub
Public Function Get_Parties() As String
Dim LFParty_Codes As String
Dim I As Integer
LFParty_Codes = vbNullString
For I = 1 To ListViewAcc.ListItems.Count
    If ListViewAcc.ListItems(I).Checked = True Then
        If LenB(LFParty_Codes) <> 0 Then LFParty_Codes = LFParty_Codes & ", "
        LFParty_Codes = LFParty_Codes & "'" & ListViewAcc.ListItems(I).SubItems(1) & "'"
    End If
Next
Get_Parties = LFParty_Codes
End Function
Sub Ledger_NEW()
    Dim LPartyName As String:           Dim LParties As String:             Dim LAddress As String
    Dim LCINNo As String:               Dim LPANNO As String:               Dim LOpBal As Double
    Dim LAC_CODE As String:             Dim LDt1 As Date:                   Dim LDT2 As Date
    Dim MDate As Date:                  Dim LVou_Type As String:            Dim LVouTypes As String
    Dim NewNarr As String:              Dim MBalance As Double:             Dim LEX As String
    Dim LEXCHANGE As String:            Dim LClient As String:              Dim IFlag  As Boolean
    Dim LClName As String:              Dim LEXNAME As String:              Dim VouRec As ADODB.Recordset:
    Dim SettleRec As ADODB.Recordset:   Dim TRec As ADODB.Recordset:        Dim PartyRec As ADODB.Recordset
    Dim LMaxDate As Date:               Dim LVouAmount As Double:           Dim LDrCr As String
    Dim LVouNo As String:               Dim LVouDt As Date:                 Dim LNarration  As String
    Dim LChqNo As String:               Dim LChqDt As Date:                 Dim LIntRate1 As Double
    Dim LIntRate2 As Double:            Dim LADRec As ADODB.Recordset:      Dim LprevBal  As Double
    Dim LNoDays As Integer:             Dim LPrevDate As Date:              Dim LVou_Dt As Date: Dim LBackDate As Date:  Dim LNextDate As Date: Dim LFirstDate As Date
    Dim LIntType  As String:                Dim LOLDbal As Double
    Dim LACCID As Long:                    Dim MIntAmt As Double: Dim mfirst As Integer: Dim TelegramFilePath As String
    Call RecSet
    LParties = Get_Parties
    If LenB(LParties) = 0 Then
        MsgBox "Please Select Account"
        ListViewAcc.SetFocus
        Exit Sub
    End If
    Set TRec = Nothing:    Set TRec = New ADODB.Recordset
    mysql = "SELECT MAX(VOU_DT) AS MDT FROM VOUCHER  WHERE COMPCODE =" & GCompCode & "  "
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        LMaxDate = VcDtpToDate.Value
    Else
        If IsNull(TRec!MDt) Then
            LMaxDate = VcDtpToDate.Value
        Else
            LMaxDate = TRec!MDt
        End If
    End If
    
    Set LADRec = Nothing:     Set LADRec = New ADODB.Recordset
    mysql = "SELECT ACCID,AC_CODE,INTRATE ,SEBIRATE,INTTYPE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE IN (" & LParties & ") ORDER BY NAME"
    LADRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    Set PartyRec = Nothing:    Set PartyRec = New ADODB.Recordset
    mysql = "SELECT ACCID,AC_CODE,NAME,OP_BAL FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE IN (" & LParties & ") ORDER BY NAME"
    PartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    Screen.MousePointer = 11: OK_CMD.Enabled = False
    GETMAIN.ProgressBar1.Visible = True
    GETMAIN.ProgressBar1.Max = PartyRec.RecordCount + 1
    GETMAIN.ProgressBar1.Value = 0
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
    Do While Not PartyRec.EOF
        If CreateChk.Value = 1 Or ChkTelegram.Value = 1 Then
            Call RecSet
        End If
        
        LOpBal = 0:                LAC_CODE = PartyRec!AC_CODE
        LACCID = PartyRec!ACCID
        
        AccRec.MoveFirst
        AccRec.Find "AC_CODE='" & LAC_CODE & "'", , adSearchForward
        AccRecDetail.MoveFirst
        AccRecDetail.Find "AC_CODE='" & LAC_CODE & "'", , adSearchForward
        If Not AccRecDetail.EOF Then
            LAddress = Trim(AccRecDetail!ADD1 & " " & AccRecDetail!City)
            LCINNo = vbNullString
            LPANNO = Trim(IIf(IsNull(AccRecDetail!PANNO), "", AccRecDetail!PANNO))
        Else
            LAddress = vbNullString: LCINNo = vbNullString: LPANNO = vbNullString
        End If
        LPartyName = PartyRec!NAME
        LOpBal = Val(PartyRec!OP_BAL)
        If ChkInterest.Value = 1 Then
            LIntRate1 = 0
            LIntRate2 = 0
            LADRec.MoveFirst
            LADRec.Find "AC_CODE='" & LAC_CODE & "'"
            If Not LADRec.EOF Then
                LIntRate1 = LADRec!INTRATE
                LIntRate2 = LADRec!SEBIRATE
                LIntType = LADRec!INTTYPE
            End If
        End If
        LPrevDate = VcDtpFromDate.Value
        LNoDays = 0
        mysql = "SELECT SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END )AS AMOUNT "
        mysql = mysql & " FROM VOUCHER AS VOU,VCHAMT AS VT WHERE VT.VOU_ID = VOU.VOU_ID "
        mysql = mysql & " AND VT.ACCID =" & LACCID & " "
        mysql = mysql & " AND VOU.VOU_DT < '" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND VOU.VOU_DT >= '" & Format(GFinBegin, "yyyy/MM/dd") & "'"
        mysql = mysql & " AND Vou.VOU_TYPE IN (" & LVouTypes & ")"
        Set TRec = Nothing:     Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        
        If Not TRec.EOF Then LOpBal = LOpBal + IIf(IsNull(TRec!AMOUNT), 0, TRec!AMOUNT)
        If ChkOpBal.Value = 0 Then LOpBal = 0
        Set TRec = Nothing
        If LOpBal <> 0 Then
            With RecRpt
                .AddNew
                !AC_CODE = LAC_CODE:                !AC_NAME = LPartyName
                !Address = LAddress:                !CINNO = LCINNo
                !PANNO = LPANNO:                    !INV_NO = " Op Bal"
                !Balance = Val(LOpBal):             !BILL_NO = vbNullString
                !INTRATE1 = LIntRate1:              !INTRATE2 = LIntRate2
                !G_TYPE = LIntType:                 !MarginAmt = 0
                If LOpBal < 0 Then
                    !DEBIT = Abs(LOpBal)
                    !CREDIT = 0
                Else
                    !DEBIT = 0
                    !CREDIT = Val(LOpBal)
                End If
                !NARRATION = vbNullString:                !CHEQUE_NO = vbNullString
                !CHEQUE_DT = vbNullString:                !REFRENCE = vbNullString
                !SM_NAME = vbNullString:                  !Group = vbNullString
                !Type = vbNullString:                     !DR_CR = vbNullString
                !G_CAT = vbNullString:                    !OP_BALANCE = 0:
                
                !QAC_DT = Format(VcDtpFromDate.Value - 1, "yyyy/MM/dd")
                !INV_DT = Format(VcDtpFromDate.Value, "yyyy/MM/dd")
                !BILL_DT = Format(VcDtpFromDate.Value, "yyyy/MM/dd")
                .Update
            End With
        End If
        
        LFirstDate = "01/01/1900"
        
        mysql = "SELECT V.VOU_NO,VT.AC_CODE,VT.DR_CR,VT.AMOUNT,VT.NARRATION,VT.CHEQUE_NO,VT.CHEQUE_DT,V.VOU_TYPE,V.VOU_DT,V.EXCODE"
        mysql = mysql & " FROM VCHAMT AS VT, VOUCHER AS V WHERE "
        mysql = mysql & " V.VOU_DT >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND V.VOU_DT <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
        mysql = mysql & " AND VT.ACCID =" & LACCID & " AND V.VOU_ID = VT.VOU_ID  "
        mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
        mysql = mysql & " AND V.VOU_TYPE NOT IN ('S','H','B','M') "
        mysql = mysql & " ORDER BY VT.VOU_DT,VT.VOU_TYPE,VT.VOU_NO"
        Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
        VouRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        While Not VouRec.EOF
            MDate = VouRec!VOU_DT:            LVou_Type = VouRec!VOU_TYPE
            
            If LFirstDate = "01/01/1900" Then
                LFirstDate = VouRec!VOU_DT
            End If
            
            DoEvents:                         NewNarr = vbNullString
            With RecRpt
                .AddNew
                If VouRec!DR_CR = "D" Then
                    !DEBIT = Val(VouRec!AMOUNT)
                    !CREDIT = 0
                Else
                    !DEBIT = 0
                    !CREDIT = Val(VouRec!AMOUNT)
                End If
                !INV_NO = VouRec!VOU_NO:                        !AC_CODE = LAC_CODE
                !AC_NAME = LPartyName:                          !Address = LAddress
                !CINNO = LCINNo:                                !PANNO = LPANNO
                !INTRATE1 = LIntRate1:                          !INTRATE2 = LIntRate2
                !G_TYPE = LIntType:                             !QAC_DT = VouRec!VOU_DT
                !NARRATION = VouRec!NARRATION:                   !MarginAmt = 0
                !CHEQUE_NO = VouRec!CHEQUE_NO & vbNullString:   !CHEQUE_DT = Format(VouRec!CHEQUE_DT, "dd/mm/yyyy")
                !Balance = 0:                                   !SM_NAME = NewNarr
                !REFRENCE = vbNullString:                       !Group = vbNullString
                !Type = vbNullString:                           !DR_CR = vbNullString
                !G_CAT = vbNullString:                          !OP_BALANCE = 0:
                !INV_DT = Format(VouRec!VOU_DT, "yyyy/MM/dd"):  !BILL_DT = Format(VouRec!VOU_DT, "yyyy/MM/dd")
                .Update
            End With
            VouRec.MoveNext
        Wend
        'Settlement
        If OptDailySettle.Value = True Then
            mysql = " SELECT V.EXCODE,V.VOU_DT,SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END ) AS AMT "
            mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V WHERE "
            mysql = mysql & " V.VOU_ID =VT.VOU_ID AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE='S' "
            mysql = mysql & " AND V.VOU_DT <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND V.VOU_DT >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND V.VOU_DT <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
            If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
            If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX'  "
            If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE'  "
            If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ'   "
            If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
            mysql = mysql & " GROUP BY V.VOU_DT,V.EXCODE"
            mysql = mysql & " ORDER BY V.VOU_DT,V.EXCODE "
            Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
            VouRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not VouRec.EOF
                MDate = VouRec!VOU_DT:
                If Round(VouRec!AMT, 2) <> 0 Then
                    With RecRpt
                        .AddNew
                        !INV_NO = "AAS " & CStr(MDate):                !AC_CODE = LAC_CODE
                        !AC_NAME = LPartyName:                       !Address = LAddress
                        !CINNO = LCINNo:                             !PANNO = LPANNO
                        !QAC_DT = MDate:                             !INTRATE1 = LIntRate1
                        !INTRATE2 = LIntRate2:                       !G_TYPE = LIntType:
                        If Round(VouRec!AMT, 2) < 0 Then
                            !DEBIT = Abs(VouRec!AMT):                !CREDIT = 0
                        Else
                            !DEBIT = 0:                              !CREDIT = VouRec!AMT
                        End If
                        !NARRATION = VouRec!excode & " Settlement":  !Balance = 0
                        !BILL_NO = vbNullString:                     !SM_NAME = VouRec!excode & " Settlement"
                        !OP_BALANCE = 0
                        .Update
                    End With
                End If
                VouRec.MoveNext
            Loop
            'Sharing
            If OptShareExchange.Value = True Then
                mysql = " SELECT V.EXCODE,V.VOU_DT,V.VOU_TYPE,SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END ) AS AMT "
                mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V WHERE "
                mysql = mysql & " V.VOU_DT >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND V.VOU_DT <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
                mysql = mysql & " AND V.VOU_DT <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
                mysql = mysql & " AND V.VOU_ID =VT.VOU_ID "
                mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
                mysql = mysql & " AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE IN ('H','B') "
                If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
                If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX' "
                If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE' "
                If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ' "
                If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
                mysql = mysql & " GROUP BY V.VOU_DT,V.EXCODE,V.VOU_TYPE "
                mysql = mysql & " ORDER BY V.VOU_DT,V.EXCODE,V.VOU_TYPE "
            Else
                mysql = " SELECT B.NAME AS CLNAME,V.EXCODE ,V.VOU_DT,V.VOU_TYPE,SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END ) AS AMT "
                mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V,ACCOUNTM AS B WHERE "
                mysql = mysql & " V.VOU_DT >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND V.VOU_DT <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
                mysql = mysql & " AND V.VOU_DT <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
                mysql = mysql & " AND B.COMPCODE =V.COMPCODE AND B.AC_CODE =VT.cheque_no AND V.VOU_ID =VT.VOU_ID "
                mysql = mysql & " AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE IN ('H','B') "
                mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
                If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
                If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX' "
                If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE' "
                If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ' "
                If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
                mysql = mysql & " GROUP BY V.VOU_DT,V.VOU_TYPE,V.EXCODE,B.NAME  "
                mysql = mysql & " ORDER BY V.VOU_DT,V.VOU_TYPE,B.NAME,V.EXCODE "
            End If
            Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
            VouRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not VouRec.EOF
                MDate = VouRec!VOU_DT
                If Round(VouRec!AMT, 2) <> 0 Then
                    With RecRpt
                        .AddNew
                        If VouRec!VOU_TYPE = "H" Then
                            !INV_NO = "AH " & VouRec!excode & " " & CStr(MDate):
                        Else
                            !INV_NO = "AB " & VouRec!excode & " " & CStr(MDate):
                        End If
                        !AC_CODE = LAC_CODE
                        !AC_NAME = LPartyName:                       !Address = LAddress
                        !CINNO = LCINNo:                             !PANNO = LPANNO
                        !INTRATE1 = LIntRate1:                       !INTRATE2 = LIntRate2
                        !G_TYPE = LIntType:                          !QAC_DT = MDate
                        If VouRec!AMT < 0 Then
                            !DEBIT = Abs(VouRec!AMT):                !CREDIT = 0
                        Else
                            !DEBIT = 0:                              !CREDIT = VouRec!AMT
                        End If
                        If OptShareExchange.Value = True Then
                            If VouRec!VOU_TYPE = "H" Then
                                !NARRATION = VouRec!excode & " Sharing":
                                !SM_NAME = VouRec!excode & " Sharing"
                            Else
                                !NARRATION = VouRec!excode & " Sub Brokerage":
                                !SM_NAME = VouRec!excode & " Sub Brokerage"
                            End If
                        Else
                            If VouRec!VOU_TYPE = "H" Then
                                !NARRATION = VouRec!CLNAME & " " & VouRec!excode & " Sharing":
                                !SM_NAME = VouRec!CLNAME & " " & VouRec!excode & " Sharing"
                            Else
                                !NARRATION = VouRec!CLNAME & " " & VouRec!excode & " Sub Brokerage":
                                !SM_NAME = VouRec!CLNAME & " " & VouRec!excode & " Sub Brokerage"
                            End If
                        End If
                        !Balance = 0:                                !BILL_NO = vbNullString:
                        !OP_BALANCE = 0
                        .Update
                    End With
                End If
                VouRec.MoveNext
            Loop
            'Margin
            mysql = " SELECT V.EXCODE,V.VOU_DT,VT.DR_CR,SUM(AMOUNT) AS AMT "
            mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V WHERE "
            mysql = mysql & " V.VOU_DT >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND V.VOU_DT <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND V.VOU_ID =VT.VOU_ID AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE='M' "
            If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
            If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX'  "
            If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE'  "
            If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ'   "
            If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
            mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
            mysql = mysql & " GROUP BY V.VOU_DT,V.EXCODE, VT.DR_CR "
            mysql = mysql & " ORDER BY V.VOU_DT,V.EXCODE, VT.DR_CR "
            Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
            VouRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not VouRec.EOF
                MDate = VouRec!VOU_DT:
                With RecRpt
                    .AddNew
                    !INV_NO = "Margin " & CStr(MDate):           !AC_CODE = LAC_CODE
                    !AC_NAME = LPartyName:                       !Address = LAddress
                    !CINNO = LCINNo:                             !PANNO = LPANNO
                    !QAC_DT = MDate:                             !INTRATE1 = LIntRate1
                    !INTRATE2 = LIntRate2:                       !G_TYPE = LIntType:
                    If VouRec!DR_CR = "D" Then
                        !DEBIT = Abs(VouRec!AMT):                !CREDIT = 0
                        !NARRATION = VouRec!excode & " Margin Debited "
                        !SM_NAME = VouRec!excode & " Margin Debited "
                    Else
                        !DEBIT = 0:                              !CREDIT = VouRec!AMT
                        !NARRATION = VouRec!excode & " Margin Credited"
                        !SM_NAME = VouRec!excode & " Margin Credited"
                    End If
                    !Balance = 0
                    !BILL_NO = vbNullString:
                    !OP_BALANCE = 0
                    .Update
                End With
                VouRec.MoveNext
            Loop
        ElseIf OptSaudaWise.Value = True Then ' SAUDA WISE
            mysql = " SELECT V.EXCODE,V.SAUDA,V.VOU_DT,SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END ) AS AMT "
            mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V WHERE "
            mysql = mysql & " V.VOU_ID =VT.VOU_ID   AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE='S' "
            mysql = mysql & " AND V.VOU_DT <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND V.VOU_DT >='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND V.VOU_DT <= '" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
            If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
            If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX'  "
            If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE'  "
            If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ'   "
            If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
            mysql = mysql & " GROUP BY V.VOU_DT,V.EXCODE,V.SAUDA"
            mysql = mysql & " ORDER BY V.VOU_DT,V.EXCODE ,V.SAUDA"
            Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
            VouRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not VouRec.EOF
                MDate = VouRec!VOU_DT:
                If Round(VouRec!AMT, 2) <> 0 Then
                    With RecRpt
                        .AddNew
                        !INV_NO = "AAS " & CStr(MDate):                !AC_CODE = LAC_CODE
                        !AC_NAME = LPartyName:                       !Address = LAddress
                        !CINNO = LCINNo:                             !PANNO = LPANNO
                        !QAC_DT = MDate:                             !INTRATE1 = LIntRate1
                        !INTRATE2 = LIntRate2
                        If Round(VouRec!AMT, 2) < 0 Then
                            !DEBIT = Abs(VouRec!AMT):                !CREDIT = 0
                        Else
                            !DEBIT = 0:                              !CREDIT = VouRec!AMT
                        End If
                        !NARRATION = VouRec!Sauda & " Settlement":  !Balance = 0
                        !BILL_NO = vbNullString:                     !SM_NAME = VouRec!excode & " Settlement"
                        !OP_BALANCE = 0
                        .Update
                    End With
                End If
                VouRec.MoveNext
            Loop
        Else
            'WEEK WISE
            LDt1 = DateValue(VcDtpFromDate.Value)
            Set SettleRec = Nothing
            Set SettleRec = New ADODB.Recordset
            mysql = " SELECT SETDATE FROM SETTLE WHERE COMPCODE =" & GCompCode & ""
            mysql = mysql & " AND SETDATE>='" & Format(VcDtpFromDate.Value, "yyyy/MM/dd") & "' AND SETDATE<='" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " AND SETDATE <='" & Format(LMaxDate, "YYYY/MM/DD") & "'"
            mysql = mysql & " AND SETDATE <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
            mysql = mysql & " ORDER BY SETDATE "
            SettleRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not SettleRec.EOF
                LDT2 = SettleRec!SETDATE
                mysql = " SELECT V.EXCODE,SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END ) AS AMT "
                mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V WHERE "
                mysql = mysql & " V.VOU_ID = VT.VOU_ID AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE='S' "
                mysql = mysql & " AND V.VOU_DT>='" & Format(LDt1, "YYYY/MM/DD") & "' AND V.VOU_DT<='" & Format(LDT2, "YYYY/MM/DD") & "'"
                mysql = mysql & " AND V.VOU_DT  <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
                mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
                If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
                If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX'  "
                If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE'  "
                If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ'   "
                If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
                mysql = mysql & " GROUP BY V.EXCODE "
                mysql = mysql & " ORDER BY V.EXCODE "
                Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
                VouRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                Do While Not VouRec.EOF
                    MDate = LDT2
                    If Round(VouRec!AMT, 2) <> 0 Then
                        With RecRpt
                            .AddNew
                            !INV_NO = "AAS " & CStr(MDate):                !AC_CODE = LAC_CODE
                            !AC_NAME = LPartyName:                       !Address = LAddress
                            !CINNO = LCINNo:                             !PANNO = LPANNO
                            !QAC_DT = MDate:                             !INTRATE1 = LIntRate1
                            !INTRATE2 = LIntRate2:                       !G_TYPE = LIntType:
                            If VouRec!AMT < 0 Then
                                !DEBIT = Abs(VouRec!AMT):                !CREDIT = 0
                            Else
                                !DEBIT = 0:                              !CREDIT = VouRec!AMT
                            End If
                            !NARRATION = VouRec!excode & " Settlement":  !Balance = 0
                            !BILL_NO = vbNullString:                     !SM_NAME = VouRec!excode & " Settlement"
                            !OP_BALANCE = 0
                            .Update
                        End With
                    End If
                    VouRec.MoveNext
                Loop
                LDt1 = LDT2 + 1
                SettleRec.MoveNext
            Loop
            If SettleRec.RecordCount >= 1 Then
                SettleRec.MoveFirst
                
            End If
            
            LDt1 = DateValue(VcDtpFromDate.Value)
            Do While Not SettleRec.EOF
                LDT2 = SettleRec!SETDATE
                If OptShareExchange.Value = True Then
                    mysql = " SELECT V.EXCODE,V.VOU_TYPE,SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END ) AS AMT "
                    mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V WHERE  "
                    mysql = mysql & " V.VOU_DT>='" & Format(LDt1, "YYYY/MM/DD") & "' AND V.VOU_DT<='" & Format(LDT2, "YYYY/MM/DD") & "'"
                    mysql = mysql & " AND V.VOU_DT <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
                    mysql = mysql & " AND V.VOU_ID = VT.VOU_ID "
                    mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
                    mysql = mysql & " AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE IN ('H','B') "
                    If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
                    If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX' "
                    If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE' "
                    If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ' "
                    If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
                    mysql = mysql & " GROUP BY V.EXCODE,V.VOU_TYPE "
                    mysql = mysql & " ORDER BY V.EXCODE,V.VOU_TYPE "
                Else
                    mysql = " SELECT B.NAME AS CLNAME,V.VOU_TYPE,SUM(CASE VT.DR_CR WHEN 'C' THEN VT.AMOUNT  WHEN 'D' THEN VT.AMOUNT*-1 END ) AS AMT "
                    mysql = mysql & " FROM VCHAMT AS VT,VOUCHER AS V,ACCOUNTM AS B WHERE  "
                    mysql = mysql & " V.VOU_DT>='" & Format(LDt1, "YYYY/MM/DD") & "' AND V.VOU_DT<='" & Format(LDT2, "YYYY/MM/DD") & "'"
                    mysql = mysql & " AND V.VOU_DT <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
                    mysql = mysql & " AND B.COMPCODE = V.COMPCODE AND B.AC_CODE =VT.cheque_no AND V.VOU_ID =VT.VOU_ID "
                    mysql = mysql & " AND VT.ACCID =" & LACCID & " AND V.VOU_TYPE IN ('H','B') "
                    If ChkNCDX.Value = 0 Then mysql = mysql & " AND V.EXCODE <>'NCDX' "
                    If ChkMCX.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'MCX' "
                    If ChkNSE.Value = 0 Then mysql = mysql & "  AND V.EXCODE <>'NSE' "
                    If ChkEQ.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'EQ' "
                    If ChkCmx.Value = 0 Then mysql = mysql & "   AND V.EXCODE <>'CMX'   "
                    mysql = mysql & " AND V.VOU_TYPE IN (" & LVouTypes & ")"
                    mysql = mysql & " GROUP BY V.VOU_TYPE,B.NAME  "
                    mysql = mysql & " ORDER BY V.VOU_TYPE,B.NAME "
                End If
                Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
                VouRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                Do While Not VouRec.EOF
                    MDate = LDT2
                    If Round(VouRec!AMT, 2) <> 0 Then
                        With RecRpt
                            .AddNew
                            If VouRec!VOU_TYPE = "H" Then
                                !INV_NO = "AH " & CStr(MDate):
                            Else
                                !INV_NO = "AB " & CStr(MDate):
                            End If
                            !AC_CODE = LAC_CODE
                            !AC_NAME = LPartyName:                       !Address = LAddress
                            !CINNO = LCINNo:                             !PANNO = LPANNO
                            !QAC_DT = MDate:                             !INTRATE1 = LIntRate1
                            !INTRATE2 = LIntRate2:                       !G_TYPE = LIntType:
                            If VouRec!AMT < 0 Then
                                !DEBIT = Abs(VouRec!AMT):                !CREDIT = 0
                            Else
                                !DEBIT = 0:                              !CREDIT = VouRec!AMT
                            End If
                            If OptShareExchange.Value = True Then
                                If VouRec!VOU_TYPE = "H" Then
                                    !NARRATION = VouRec!excode & " Sharing":
                                    !SM_NAME = VouRec!excode & " Sharing"
                                Else
                                    !NARRATION = VouRec!excode & " Sub Brokerage":
                                    !SM_NAME = VouRec!excode & " Sub Brokerage"
                                End If
                            Else
                                If VouRec!VOU_TYPE = "H" Then
                                    !NARRATION = VouRec!CLNAME & " Sharing":
                                    !SM_NAME = VouRec!CLNAME & "  Sharing"
                                Else
                                    !NARRATION = VouRec!CLNAME & " Sub Brokerage":
                                    !SM_NAME = VouRec!CLNAME & "  Sub Brokerage"
                                End If
                            End If
                            !Balance = 0:                                !BILL_NO = vbNullString:
                            !OP_BALANCE = 0
                            .Update
                        End With
                    End If
                    VouRec.MoveNext
                Loop
                LDt1 = LDT2 + 1
                SettleRec.MoveNext
            Loop
        End If
        Set VouRec = Nothing
        'End If
FLAG1:
        If CreateChk.Value = 1 Or ChkTelegram.Value = 1 Then
            If IFlag = True Then IFlag = False
            If RecRpt.RecordCount > 0 Then
                'RecRpt.Sort = "AC_NAME,VOU_DT,VOU_NO"
                RecRpt.Sort = "AC_NAME,QAC_DT,INV_NO"
                Set RDCREPO = Nothing
                Set RDCREPO = New CRAXDRT.report
                If ChkInterest.Value = 1 Then
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PRTSTT-INT.RPT", 1)
                Else
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PRTSTT.RPT", 1)
                End If
                RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
                RDCREPO.FormulaFields.GetItemByName("ADD1").text = "'For " & CStr(VcDtpFromDate.Value) & " To " & CStr(VcDtpToDate.Value) & "'"
                RDCREPO.Database.SetDataSource RecRpt
                If ChkNewPage.Value = Val(1) Then RDCREPO.Areas(3).NewPageBefore = True ''''FOR EACH GROUP STARTS WITH A NEW PAGE
                If Combo1.ListIndex = 0 Then
                    TelegramFilePath = App.Path & "\RPT\LEDGER-" & LPartyName & "-" & Left$(CDate(VcDtpFromDate.Value), 2) & Mid(CDate(VcDtpFromDate.Value), 4, 2) & Right$(CDate(VcDtpFromDate.Value), 4) & "-" & Left$(CDate(VcDtpToDate.Value), 2) & Mid(CDate(VcDtpToDate.Value), 4, 2) & Right$(CDate(VcDtpToDate.Value), 4) & ".PDF"
                    RDCREPO.ExportOptions.DiskFileName = App.Path & "\RPT\LEDGER-" & LPartyName & "-" & Left$(CDate(VcDtpFromDate.Value), 2) & Mid(CDate(VcDtpFromDate.Value), 4, 2) & Right$(CDate(VcDtpFromDate.Value), 4) & "-" & Left$(CDate(VcDtpToDate.Value), 2) & Mid(CDate(VcDtpToDate.Value), 4, 2) & Right$(CDate(VcDtpToDate.Value), 4) & ".PDF"
                    RDCREPO.ExportOptions.FormatType = crEFTPortableDocFormat
                Else
                    TelegramFilePath = App.Path & "\RPT\LEDGER-" & LPartyName & "-" & Left$(CDate(VcDtpFromDate.Value), 2) & Mid(CDate(VcDtpFromDate.Value), 4, 2) & Right$(CDate(VcDtpFromDate.Value), 4) & "-" & Left$(CDate(VcDtpToDate.Value), 2) & Mid(CDate(VcDtpToDate.Value), 4, 2) & Right$(CDate(VcDtpToDate.Value), 4) & ".XLS"
                    RDCREPO.ExportOptions.DiskFileName = App.Path & "\RPT\LEDGER-" & LPartyName & "-" & Left$(CDate(VcDtpFromDate.Value), 2) & Mid(CDate(VcDtpFromDate.Value), 4, 2) & Right$(CDate(VcDtpFromDate.Value), 4) & "-" & Left$(CDate(VcDtpToDate.Value), 2) & Mid(CDate(VcDtpToDate.Value), 4, 2) & Right$(CDate(VcDtpToDate.Value), 4) & ".XLS"
                    RDCREPO.ExportOptions.FormatType = crEFTExcel80
                End If
                RDCREPO.ExportOptions.DestinationType = crEDTDiskFile
                RDCREPO.ExportOptions.PDFExportAllPages = True
                RDCREPO.Export False
            End If
        End If
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        Call PERCENTAGE(GETMAIN.ProgressBar1.Max, GETMAIN.ProgressBar1.Value)
        DoEvents
        
        If ChkTelegram.Value = 1 Then
            Dim RecTelegram As ADODB.Recordset: Dim LUserId As String: Dim LAccessHash As String
            Set RecTelegram = Nothing
            Set RecTelegram = New ADODB.Recordset
            mysql = "SELECT FAX,DIRECTOR FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " and NAME ='" & LPartyName & "'"
            RecTelegram.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not RecTelegram.EOF Then
                LUserId = Trim(RecTelegram!DIRECTOR)
                LAccessHash = RecTelegram!Fax
                'And LenB(LAccessHash) > 0
                If LenB(LUserId) > 0 Then
                    Call Send_Telegram(LUserId, LAccessHash, TelegramFilePath, "")
                Else
                   ' MsgBox "UserId not Defined for " & LPartyName & ""
                End If
            End If
        End If
        
        DoEvents
        PartyRec.MoveNext
    Loop
    If ChkInterest.Value = 1 Then
        If RecRpt.RecordCount > 0 Then
            Dim TRec1 As ADODB.Recordset
            Dim LNarr As String
            Set TRec1 = Nothing
            Set TRec1 = New ADODB.Recordset
            Set TRec1 = RecRpt.Clone
            
            TRec1.Sort = "AC_NAME,QAC_DT,INV_NO"
            TRec1.MoveFirst
            Do While Not TRec1.EOF
                LAC_CODE = TRec1!AC_CODE
                LPartyName = TRec1!AC_NAME
                LPrevDate = VcDtpFromDate.Value
                LIntRate1 = TRec1!INTRATE1
                LIntRate2 = TRec1!INTRATE1
                LIntType = TRec1!G_TYPE
                LNoDays = 0
                LprevBal = 0
                mfirst = 1
                                                
                Do While LAC_CODE = TRec1!AC_CODE
                    LVou_Dt = TRec1!QAC_DT
                    LOLDbal = LprevBal
                    Do While LAC_CODE = TRec1!AC_CODE And LVou_Dt = TRec1!QAC_DT
                        'LVouNo = TRec1!INV_NO
                        'LNarr = TRec1!NARRATION
                        
                        RecRpt.MoveFirst
                        'RecRpt.Filter = " AC_CODE ='" & LAC_CODE & "'  AND QAC_DT='" & Format(LVou_Dt, "YYYY/MM/DD") & "'   AND INV_NO='" & LVouNo & "'  AND NARRATION = '" & LNarr & "' "
                        RecRpt.Filter = " AC_CODE ='" & LAC_CODE & "'  AND QAC_DT='" & Format(LVou_Dt, "YYYY/MM/DD") & "'   "
                        If Not RecRpt.EOF Then
                            RecRpt!OP_BALANCE = 0
                            RecRpt!INTAMT = 0
                            RecRpt!PREVBAL = LprevBal
                            If TRec1!DEBIT <> 0 Then
                                LprevBal = LprevBal + (TRec1!DEBIT * -1)
                            Else
                                LprevBal = LprevBal + (TRec1!CREDIT)
                            End If
                            RecRpt.Update
                        End If

'                        TRec1.MovePrevious  '>>> TRec1.MoveNext 'sACHIN
'                        If TRec1.BOF Then '>>>>If TRec1.EOF Then
'                            LBackDate = VcDtpFromDate.Value   '>>>LNextDate = VcDtpToDate.Value  'sACHIN
'                            Exit Do
'                        Else ''sACHIN
'                            '>>>LNextDate = TRec1!QAC_DT 'sACHIN
'                            LBackDate = TRec1!QAC_DT 'sACHIN
'                        End If
                        TRec1.MoveNext
                        If TRec1.EOF Then Exit Do
                    Loop

                    TRec1.MovePrevious
                    'LVou_Dt = VcDtpFromDate.Value Then
                    If mfirst = 1 Then
                        mfirst = 0
                        LPrevDate = LPrevDate - 1
                    End If
                    
                    If LPrevDate < LVou_Dt Then
                        LNoDays = DateDiff("d", LPrevDate, LVou_Dt)
                        '>>>LNoDays = DateDiff("d", LVou_Dt, LNextDate)  'sACHIN
                        'LNoDays = DateDiff("d", LBackDate, LVou_Dt) 'sACHIN
                        
                        LPrevDate = LVou_Dt
                        RecRpt.MoveFirst
                        RecRpt.Filter = " AC_CODE ='" & LAC_CODE & "'  AND QAC_DT='" & Format(LVou_Dt, "YYYY/MM/DD") & "'   "
                        RecRpt.MoveLast
                        'RecRpt.Filter = " AC_CODE ='" & LAC_CODE & "' AND INV_NO='" & LVouNo & "' AND QAC_DT='" & Format(LVou_Dt, "YYYY/MM/DD") & "'  AND NARRATION ='" & LNarr & "'"
                        If Not RecRpt.EOF Then
                            RecRpt!INTAMT = 0
                            RecRpt!OP_BALANCE = LNoDays
                            If LIntType = "P" And LprevBal < 0 Then
                                MIntAmt = 0
                                If LOLDbal < 0 Then
                                    MIntAmt = (LNoDays - 1) * ((RecRpt!INTRATE1 / 100) / 365) * LOLDbal
                                End If
                                MIntAmt = MIntAmt + (1 * ((RecRpt!INTRATE1 / 100) / 365) * LprevBal)
                                RecRpt!INTAMT = MIntAmt
                            'ElseIf LIntType = "P" And RecRpt!PREVBAL > 0 Then
                             '   RecRpt!INTAMT = LNoDays * ((RecRpt!INTRATE1 / 100) / 365) * LprevBal
                            ElseIf LIntType = "R" And RecRpt!PREVBAL > 0 Then
                                RecRpt!INTAMT = LNoDays * ((RecRpt!INTRATE1 / 100) / 365) * LprevBal
                            ElseIf LIntType = "B" Then
                                RecRpt!INTAMT = LNoDays * ((RecRpt!INTRATE1 / 100) / 365) * LprevBal
                            End If
                            RecRpt.Update
                        End If
                    End If
                    TRec1.MoveNext
                    If TRec1.EOF Then Exit Do
                Loop
                If LVou_Dt <> VcDtpToDate.Value Then
                    
                    TRec1.MoveLast
                    LBackDate = TRec1!QAC_DT
                    
                    RecRpt.AddNew
                    RecRpt!INV_NO = "Closing Bal"
                    RecRpt!AC_CODE = LAC_CODE
                    RecRpt!AC_NAME = LPartyName:
                    RecRpt!Address = vbNullString
                    RecRpt!CINNO = vbNullString:
                    RecRpt!PANNO = vbNullString
                    RecRpt!INTRATE1 = LIntRate1:
                    RecRpt!INTRATE2 = LIntRate2
                    RecRpt!PREVBAL = LprevBal
                    RecRpt!G_TYPE = LIntType:
                    RecRpt!QAC_DT = Format(VcDtpToDate.Value, "YYYY/MM/DD")
                    RecRpt!DEBIT = 0:
                    RecRpt!CREDIT = 0
                    
                    '>>>LNoDays = DateDiff("d", LPrevDate, VcDtpToDate.Value)
                    LNoDays = DateDiff("d", LBackDate, VcDtpToDate.Value)
                    
                    LPrevDate = LVou_Dt
                    RecRpt!OP_BALANCE = LNoDays + 1
                    RecRpt!NARRATION = "Closing Bal"
                    RecRpt!SM_NAME = vbNullString
                    RecRpt!Balance = 0:
                    RecRpt!BILL_NO = vbNullString:
                    'RecRpt!INTAMT = 0   '(LNoDays + 1) * ((RecRpt!INTRATE1 / 100) / 365) * RecRpt!PREVBAL 'sACHIN
                    If RecRpt!PREVBAL < 0 Then
                        MIntAmt = (LNoDays) * ((RecRpt!INTRATE1 / 100) / 365) * RecRpt!PREVBAL  'sACHIN
                        RecRpt!INTAMT = (LNoDays) * ((RecRpt!INTRATE1 / 100) / 365) * RecRpt!PREVBAL  'sACHIN
                        RecRpt.Update
                    End If
                End If
                
'                '>> calculate inerest on opening amount  ' sACHIN
'                If Not RecRpt.EOF Then
'                        LAC_CODE = " Op Bal"
'                        RecRpt.MoveFirst
'                        RecRpt.Filter = " INV_NO ='" & LAC_CODE & "'  "
'
'                        If Not RecRpt.EOF Then
'                            RecRpt.MoveLast
'                            LNoDays = DateDiff("d", VcDtpFromDate.Value - 1, LFirstDate)  ''sACHIN
'                            RecRpt!INTAMT = LNoDays * ((RecRpt!INTRATE1 / 100) / 365) * LOpBal
'                            RecRpt.Update
'                            RecRpt.MoveFirst
'                        End If
'                End If
                
                
                'If TRec1.EOF Then Exit Do
                Exit Do
            Loop
        End If
    End If
    
    
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    If CreateChk.Value = 0 And ChkTelegram.Value = 0 Then
        RecRpt.Filter = adFilterNone
        If RecRpt.RecordCount > 0 Then
            RecRpt.Sort = "AC_NAME,QAC_DT,INV_NO"
            Set RDCREPO = Nothing
            Set RDCREPO = New CRAXDRT.report
            'Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PRTSTT.RPT", 1)
            If ChkInterest.Value = 1 Then
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PRTSTT-INT.RPT", 1)
                Else
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PRTSTT.RPT", 1)
                End If
            RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
            RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Ledger From  " & CStr(VcDtpFromDate.Value) & " To " & CStr(VcDtpToDate.Value) & "'"
            RDCREPO.Database.SetDataSource RecRpt
            If ChkNewPage.Value = Val(1) Then RDCREPO.Areas(3).NewPageBefore = True ''''FOR EACH GROUP STARTS WITH A NEW PAGE
            CRViewer1.Width = CInt(GETMAIN.Width - 100)
            CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
            CRViewer1.Top = 0
            CRViewer1.Left = 0
            CRViewer1.ReportSource = RDCREPO
            CRViewer1.Visible = True
            CRViewer1.ViewReport
        Else
            MsgBox "Record does not exists."
        End If
    Else
        If ChkTelegram.Value = 1 Then
            MsgBox "Ledger Telegram Succesfully"
        Else
            MsgBox "Ledger Exported Succesfully"
        End If
    End If
        GETMAIN.ProgressBar1.Visible = False
            
        Set RecRpt = Nothing
        GETMAIN.PERLBL.Caption = vbNullString
        GETMAIN.ProgressBar1.Value = 0
        Screen.MousePointer = 0
        OK_CMD.Enabled = True
        Set RecRpt = Nothing
        Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    GETMAIN.ProgressBar1.Value = 0: OK_CMD.Enabled = True
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    Screen.MousePointer = 0
End Sub
Private Sub TxtBranchCode_Validate(Cancel As Boolean)
If LenB(TxtBranchCode.text) > 0 Then
    If FmlyRec.RecordCount > 0 Then FmlyRec.MoveFirst
    FmlyRec.Find "FMLYCODE='" & TxtBranchCode.text & "'"
    If Not FmlyRec.EOF Then
        DComboFmly.BoundText = TxtBranchCode
    Else
        TxtBranchCode.text = vbNullString
    End If
End If
End Sub
