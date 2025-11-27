VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A6BDE5D5-8F7A-11D1-9C65-4CA605C10E27}#5.0#0"; "ACTIVEGUI.OCX"
Begin VB.MDIForm GETMAIN 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SAUDA"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   -4005
   ClientWidth     =   8760
   Icon            =   "Getmain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13800
      Top             =   7440
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":1B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":20D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":21F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":2312
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":2432
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":2552
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":2672
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Getmain.frx":2792
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1659
            MinWidth        =   1659
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1660
            MinWidth        =   1660
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9253
            MinWidth        =   9253
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1499
            MinWidth        =   1499
            TextSave        =   "7:04 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveGUICtl.ActiveToolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      Negotiate       =   -1  'True
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   741
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   390
         Left            =   6960
         TabIndex        =   2
         Top             =   45
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   688
         _Version        =   393216
         ForeColor       =   16711680
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin ActiveGUICtl.ActiveSeparator ActiveSeparator1 
         Height          =   510
         Index           =   0
         Left            =   2200
         Top             =   0
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   900
         Size            =   1
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Add New Record"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":478E
         BackColor       =   -2147483638
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Modify an Existing Record"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":48A0
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Delete"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":49B2
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   4
         Left            =   1650
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Save"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":4AC4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   5
         Left            =   2295
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":4BD6
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   6
         Left            =   2800
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Close"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":4CE8
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   7
         Left            =   3435
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Search"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":4DFA
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveButton Toolbar1_Buttons 
         Height          =   375
         Index           =   8
         Left            =   3945
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "View List"
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Picture         =   "Getmain.frx":4F0C
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         Object.ToolTipText     =   ""
      End
      Begin ActiveGUICtl.ActiveSeparator ActiveSeparator1 
         Height          =   510
         Index           =   1
         Left            =   3360
         Top             =   0
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   900
         Size            =   1
      End
      Begin VB.Label PERLBL 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         TabIndex        =   12
         Top             =   45
         Width           =   6375
      End
   End
   Begin VB.PictureBox PictureMenu 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   8760
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   8760
   End
   Begin VB.Menu Database 
      Caption         =   "&Action       "
      Begin VB.Menu addrec 
         Caption         =   "&New"
         Shortcut        =   {F1}
      End
      Begin VB.Menu modifyrec 
         Caption         =   "&Edit"
         Shortcut        =   {F2}
      End
      Begin VB.Menu deleterec 
         Caption         =   "&Delete"
         Shortcut        =   {F3}
      End
      Begin VB.Menu saverec 
         Caption         =   "Sa&ve"
         Shortcut        =   {F4}
      End
      Begin VB.Menu SLS01 
         Caption         =   "-"
      End
      Begin VB.Menu cancle 
         Caption         =   "&Cancel"
         Shortcut        =   ^Z
      End
      Begin VB.Menu close 
         Caption         =   "Cl&ose"
         Shortcut        =   ^Q
      End
      Begin VB.Menu SLS02 
         Caption         =   "-"
      End
      Begin VB.Menu serchrec 
         Caption         =   "&Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu listrec 
         Caption         =   "&List"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu TRANS 
      Caption         =   "&Entries      "
      Begin VB.Menu CONTRACTENTRY 
         Caption         =   "&Contract Entry (1)"
         Visible         =   0   'False
      End
      Begin VB.Menu CONTRACTENTRY7 
         Caption         =   "&Contract Entry (7)"
         Visible         =   0   'False
      End
      Begin VB.Menu SLT01 
         Caption         =   "-"
      End
      Begin VB.Menu CLOSERATE 
         Caption         =   "&Closing Rate Entry"
      End
      Begin VB.Menu mnuptyclose 
         Caption         =   "Party Wise Closing Rate"
      End
      Begin VB.Menu dcm 
         Caption         =   "Daily Contract Modify"
      End
      Begin VB.Menu mnutrdconfirm 
         Caption         =   "Daily Trade Confirmation"
      End
      Begin VB.Menu CORRDIV 
         Caption         =   "Correction/Dividend Entry"
      End
      Begin VB.Menu SLT03 
         Caption         =   "-"
      End
      Begin VB.Menu VCHENT 
         Caption         =   "&Voucher Entry"
      End
      Begin VB.Menu mnutrfbal 
         Caption         =   "Transfer Account  Balance"
      End
      Begin VB.Menu SLT04 
         Caption         =   "-"
      End
      Begin VB.Menu MENUINVGEN 
         Caption         =   "Invoice Generation"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuinvset 
         Caption         =   "Invoice Generation Settlement Wise"
      End
      Begin VB.Menu SLE05 
         Caption         =   "-"
      End
      Begin VB.Menu mnudata 
         Caption         =   "&Data Import"
      End
   End
   Begin VB.Menu mnuquery 
      Caption         =   "&Query     "
      Begin VB.Menu mnuqaccount 
         Caption         =   "Query On Account"
      End
      Begin VB.Menu mnuQTB 
         Caption         =   "&Query on Trial Balance"
      End
      Begin VB.Menu GenQry 
         Caption         =   "&General Query"
      End
      Begin VB.Menu mnuqstand 
         Caption         =   "Query On Standing"
      End
      Begin VB.Menu mnuqtrade 
         Caption         =   "Query on Trade"
      End
      Begin VB.Menu mnuopening 
         Caption         =   "Opening Balance"
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu QonBlst 
         Caption         =   "Query on Bill List (Live Rate)"
      End
      Begin VB.Menu QrtonBill 
         Caption         =   "Query on Bill"
      End
      Begin VB.Menu QTradeFinder 
         Caption         =   "Query Trade Finder"
      End
   End
   Begin VB.Menu report 
      Caption         =   "&Reports     "
      Begin VB.Menu CONTREPORTS 
         Caption         =   "&Contract Reports"
         Begin VB.Menu mnucnote 
            Caption         =   "Contract Note"
         End
         Begin VB.Menu CONTRACTREG 
            Caption         =   "Contract Register"
         End
         Begin VB.Menu mnutrdreg 
            Caption         =   "Trade Register"
         End
         Begin VB.Menu DATEWSCONTLIST 
            Caption         =   "Date wise Contract List"
         End
         Begin VB.Menu mnutrurnover 
            Caption         =   "TurnOver Report"
         End
         Begin VB.Menu mnuannual 
            Caption         =   "Annual Global Transaction Statement"
         End
         Begin VB.Menu SPE100 
            Caption         =   "-"
         End
         Begin VB.Menu SAUDAWSSTND 
            Caption         =   "Standing Statement"
         End
         Begin VB.Menu MST 
            Caption         =   "Maturity wise Standing Report"
         End
         Begin VB.Menu mnudtst 
            Caption         =   "Datewise Standing Report"
         End
         Begin VB.Menu spe151 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrade 
            Caption         =   "Trade File Genration"
         End
         Begin VB.Menu mnuTradeDelete 
            Caption         =   "Trade Deletion"
         End
         Begin VB.Menu spr152 
            Caption         =   "-"
         End
         Begin VB.Menu mnudeposit 
            Caption         =   "ODIN Deposit Upload"
         End
         Begin VB.Menu mnumarginfile 
            Caption         =   "Daily Margin File Upload"
         End
         Begin VB.Menu CrURpt 
            Caption         =   "Credit Utilisation Report"
         End
         Begin VB.Menu OrdRegRPT 
            Caption         =   "Order Register"
         End
         Begin VB.Menu contractlogreport 
            Caption         =   "Contract Log"
         End
      End
      Begin VB.Menu mnuBill 
         Caption         =   "Billwise Reports"
         Begin VB.Menu mnubillacstt 
            Caption         =   "Billwise Account Statement"
         End
         Begin VB.Menu mnuacsttsumm 
            Caption         =   "Billwise Account Statement Summary"
         End
         Begin VB.Menu mnuBillsumm 
            Caption         =   "Billwise Bill Summary"
         End
      End
      Begin VB.Menu mnusepp27 
         Caption         =   "-"
      End
      Begin VB.Menu mnusetrpt 
         Caption         =   "Settlment Reports"
         Begin VB.Menu mnumargin 
            Caption         =   "Margin Reports"
            Begin VB.Menu mnumrgnrt 
               Caption         =   "Margin Rate List"
            End
            Begin VB.Menu mnuline2 
               Caption         =   "-"
            End
            Begin VB.Menu mnudclmgn 
               Caption         =   "Daily Client Wise Margin"
            End
            Begin VB.Menu dtmrgnrpt 
               Caption         =   "Date Wise Margin Report"
            End
            Begin VB.Menu Marsry 
               Caption         =   "&Margin Report"
            End
            Begin VB.Menu munmarginfiledaily 
               Caption         =   "Daily Margin File"
            End
         End
         Begin VB.Menu swmrt 
            Caption         =   "Missing Closing Rate"
         End
         Begin VB.Menu mnuinv 
            Caption         =   "Invoice Reports"
            Begin VB.Menu MenuInvPrint 
               Caption         =   "I&nvoice Printing"
            End
            Begin VB.Menu mnuline3 
               Caption         =   "-"
            End
            Begin VB.Menu mnuINVLIST 
               Caption         =   "&Invoice List"
            End
            Begin VB.Menu mnublist 
               Caption         =   "Invoice GST Report"
            End
         End
         Begin VB.Menu mnuaccstt 
            Caption         =   "Account Statement"
            Begin VB.Menu MENUACCSTT 
               Caption         =   "&Account Statement"
            End
            Begin VB.Menu mnuaccsmry 
               Caption         =   "Account Statement Summary"
            End
            Begin VB.Menu mnuactivity 
               Caption         =   "Activity Summary"
            End
            Begin VB.Menu mnudaily 
               Caption         =   "Daily Activity Statement"
            End
            Begin VB.Menu mnuNewStm 
               Caption         =   "Billwise Summary"
            End
            Begin VB.Menu mnuline4 
               Caption         =   "-"
            End
            Begin VB.Menu mnubillsmry 
               Caption         =   "Bill Summary"
            End
            Begin VB.Menu mnudass 
               Caption         =   "Bill Summary With Sharing"
            End
            Begin VB.Menu mnusetbsmry 
               Caption         =   "Branchwise Account Statement Summary"
            End
         End
         Begin VB.Menu mnubrbal 
            Caption         =   "Branchwise Balances"
         End
         Begin VB.Menu mnubrshare 
            Caption         =   "Branchwise Sharing Report"
         End
         Begin VB.Menu mnubrokerage 
            Caption         =   "Brokerage Reports"
            Begin VB.Menu mnubrbrok 
               Caption         =   "Branch Wise Brokerage Report"
            End
            Begin VB.Menu RPTBROKSMRY 
               Caption         =   "B&rokerage Summary"
            End
            Begin VB.Menu mnulin45 
               Caption         =   "-"
            End
            Begin VB.Menu BrokLst 
               Caption         =   "B&rokerage List"
            End
            Begin VB.Menu mnusubshare 
               Caption         =   "Sub Brokerage && Sharing List"
            End
         End
         Begin VB.Menu mnusaudasmry 
            Caption         =   "Sauda Summary"
            Visible         =   0   'False
         End
         Begin VB.Menu SAUDASUMRY 
            Caption         =   "&Generate Reports"
         End
         Begin VB.Menu RtLst 
            Caption         =   "&Rate List"
         End
         Begin VB.Menu Sdatrf 
            Caption         =   "Sauda &Transfer (Party to Party)"
         End
         Begin VB.Menu Sdatrf_dt 
            Caption         =   "Sauda &Transfer (Date to Date)"
         End
         Begin VB.Menu RPTSLotChange 
            Caption         =   "Sauda Lot Change"
         End
      End
      Begin VB.Menu SLRC1 
         Caption         =   "-"
      End
      Begin VB.Menu voucherm 
         Caption         =   "Voucher"
         Begin VB.Menu voulist 
            Caption         =   "&Voucher List"
         End
         Begin VB.Menu voucherLog 
            Caption         =   "Voucher &Log"
         End
      End
      Begin VB.Menu SLR04 
         Caption         =   "-"
      End
      Begin VB.Menu mnudaybook 
         Caption         =   "Day Books"
         Begin VB.Menu cashbook_f2 
            Caption         =   "&Cash Book"
         End
         Begin VB.Menu munilike25 
            Caption         =   "-"
         End
         Begin VB.Menu bankbook_f2 
            Caption         =   "&Bank Book"
         End
      End
      Begin VB.Menu SLR03 
         Caption         =   "Party Ledger"
         Index           =   0
         Begin VB.Menu genled 
            Caption         =   "&General Ledger"
         End
         Begin VB.Menu MNUInvWsLedg 
            Caption         =   "&Invoice wise Ledger"
         End
         Begin VB.Menu mnuoutstanding 
            Caption         =   "Partywise Outstanding"
         End
         Begin VB.Menu partyage 
            Caption         =   "Party Ageing Analysis (annexure 4)"
         End
         Begin VB.Menu mnuinterest 
            Caption         =   "Partywise Interest Collection"
         End
         Begin VB.Menu mnulplp 
            Caption         =   "-"
         End
         Begin VB.Menu mnidetldg 
            Caption         =   "Detailed Ledger"
         End
      End
      Begin VB.Menu SLRC01 
         Caption         =   "-"
      End
      Begin VB.Menu chqreg 
         Caption         =   "Che&que Register"
      End
      Begin VB.Menu clnrpt 
         Caption         =   "C&ollection Report"
         Visible         =   0   'False
         Begin VB.Menu ccrpt 
            Caption         =   "&Cash Receipts"
            Visible         =   0   'False
         End
         Begin VB.Menu bcrpt 
            Caption         =   "&Bank Receipts"
         End
         Begin VB.Menu cbcrpt 
            Caption         =   "C&ash/Bank Receipts"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu udtb 
         Caption         =   "&Trial Balance"
      End
      Begin VB.Menu bwtbal 
         Caption         =   "Branch wise Trial balance"
         Visible         =   0   'False
      End
      Begin VB.Menu mnutrialdt 
         Caption         =   "MTM Summary"
      End
      Begin VB.Menu PANDLMENU 
         Caption         =   "&Profit && Loss"
      End
      Begin VB.Menu balancesht 
         Caption         =   "B&alance Sheet"
      End
   End
   Begin VB.Menu master 
      Caption         =   "&Setup       "
      Begin VB.Menu COMPSETUP 
         Caption         =   "&Company Setup"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu ACTYPE 
         Caption         =   "&Group Setup"
      End
      Begin VB.Menu ACCOUNTHEAD 
         Caption         =   "Account Setup"
      End
      Begin VB.Menu FmlyStup 
         Caption         =   "Branch Setup"
      End
      Begin VB.Menu SLM01 
         Caption         =   "-"
      End
      Begin VB.Menu ITEMSETUP 
         Caption         =   "&Item Setup"
      End
      Begin VB.Menu SAUDAMAST 
         Caption         =   "&Sauda Master Setup"
      End
      Begin VB.Menu SETMASTER 
         Caption         =   "&Settlement MasterSetup"
      End
      Begin VB.Menu mnunarr 
         Caption         =   "Narration Setup"
      End
      Begin VB.Menu MniFileFormat 
         Caption         =   "Trade File Format"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubrokerageset 
         Caption         =   "Brokerage"
         Begin VB.Menu mnuexbrok 
            Caption         =   "Exchange Wise Brokerage Setup"
         End
      End
      Begin VB.Menu mnuslab21 
         Caption         =   "Sub Brokerage and Sharing "
         Begin VB.Menu mnuexsbrok 
            Caption         =   "Exchange Wise Sub Brokerage and Sharing Setup"
         End
      End
      Begin VB.Menu Exstp 
         Caption         =   "&Exchange Setup"
      End
      Begin VB.Menu UrSetup 
         Caption         =   "&User Setup"
      End
      Begin VB.Menu mnuweb 
         Caption         =   "&Web User Master"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu utilities 
      Caption         =   "&Tools       "
      Begin VB.Menu PRINTER 
         Caption         =   "&Printer"
         Begin VB.Menu PRINTTOGGLE 
            Caption         =   "Print Toggle On"
            Checked         =   -1  'True
            Shortcut        =   ^{F12}
         End
         Begin VB.Menu SUP01 
            Caption         =   "-"
         End
         Begin VB.Menu PRINTERSETUP 
            Caption         =   "Printer Setup"
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu rwb 
         Caption         =   "&Correcting Books"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu comsel 
         Caption         =   "&Change Company"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu lofoff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu sp4 
         Caption         =   "-"
      End
      Begin VB.Menu DBkp 
         Caption         =   "DataBase Backup"
      End
      Begin VB.Menu Reindex 
         Caption         =   "ReIndexing"
      End
      Begin VB.Menu Notepad 
         Caption         =   "&Notepad"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MnuCalc 
         Caption         =   "&Calculator"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnutele 
         Caption         =   "Telegram Api"
      End
      Begin VB.Menu loginoff 
         Caption         =   "Login/Logoff Detail"
         Visible         =   0   'False
      End
      Begin VB.Menu SysLockDate 
         Caption         =   "Settlement Lock Date"
      End
      Begin VB.Menu YrUpdate 
         Caption         =   "Year Updation"
      End
      Begin VB.Menu packupdata 
         Caption         =   "Packup Data"
      End
   End
   Begin VB.Menu mnuwind 
      Caption         =   "&Window       "
      WindowList      =   -1  'True
   End
   Begin VB.Menu Help 
      Caption         =   "&Help        "
      Visible         =   0   'False
      Begin VB.Menu Contents 
         Caption         =   "Contents..."
      End
   End
   Begin VB.Menu exit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "GETMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public old_frm As String:   Public New_Frm As String:   Public Fb_Press As Byte
Public Sub AccountHead_Click()
    Call Menu_Check("ACCOUNTHEAD", "M")
    Call PERMISSIONS("ACCOUNTHEAD")
    GETACNT.Show
End Sub
Public Sub ACTYPE_Click()
    Call Menu_Check("ACTYPE", "M")
    Call PERMISSIONS("ACTYPE")
    AccGrp.Show
    Call Get_Selection(10)
End Sub
Private Sub addrec_Click()
    Call Toolbar1_Buttons_Click(1)
End Sub
Public Sub balancesht_Click()
    Call Menu_Check("balancesht", "R")
    MFormat = "Balance Sheet"
    RPTPANDL.Show
End Sub
Public Sub bankbook_f1_Click()
    MFormat = "Bank Book Regular Format"
    GETACRPT.Show
End Sub
Public Sub bankbook_f2_Click()
    Call Menu_Check("bankbook_f2", "R")
    MFormat = "Bank Book Ledger Format"
    GETACRPT.Show
End Sub
Public Sub bcrpt_Click()
    MFormat = "Bank Collection"
    GETQYRPT.Show
End Sub
Public Sub BrokLst_Click()
    Call Menu_Check("BrokLst", "R")
    MFormat = "Brokerage List"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub bwtbal_Click()
    Call Menu_Check("udtb", "R")
    MFormat = "Branch wise Trial Balance 3"
    RPTTB.Show
End Sub
Private Sub cancle_Click()
    Call Toolbar1_Buttons_Click(5)
End Sub
Public Sub cashbook_f1_Click()
    MFormat = "Cash Book Regular Format"
    GETACRPT.Show
End Sub
Public Sub cashbook_f2_Click()
    Call Menu_Check("cashbook_f2", "R")
    MFormat = "Cash Book Ledger Format"
    GETACRPT.Show
End Sub
Public Sub cbcrpt_Click()
    MFormat = "Cash/Bank Collection"
    GETQYRPT.Show
End Sub
Public Sub ccrpt_Click()
    MFormat = "Cash Collection"
    GETQYRPT.Show
End Sub
Public Sub chqreg_Click()
    Call Menu_Check("chqreg", "R")
    MFormat = "Cheque Register"
    GETACRPT.Show
End Sub
Private Sub close_Click()
    Fb_Press = CByte(0)
    Call CLOSE_RECORD
End Sub
Public Sub CLOSERATE_Click()
    Call Menu_Check("CLOSERATE", "T")
    Call PERMISSIONS("CLOSERATE")
    FrmCloseRate.Show
    Call Get_Selection(10)
End Sub
Public Sub COMPSETUP_Click()
    Call Menu_Check("COMPSETUP", "M")
    Call PERMISSIONS("COMPSETUP")
    GETCOMP.Show
    Call Get_Selection(10)
End Sub
Public Sub comsel_Click()
    LogOffBool = True
    Set SelComp_Ado = Nothing
    Set SelComp_Ado = New ADODB.Recordset
    mysql = "SELECT COMPCODE,NAME,ACORDER FROM COMPANY ORDER BY COMPANY.COMPCODE"
    SelComp_Ado.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    SELCOMP.Show
End Sub
Public Sub Contents_Click()
    CommonDialog1.HelpCommand = cdlHelpHelpOnHelp
    CommonDialog1.ShowHelp
End Sub
Public Sub CONTRACTENTRY_Click()
    Call Menu_Check("CONTRACTENTRY", "T")
    Call PERMISSIONS("CONTRACTENTRY")
    Call Get_Selection(10)
       
    Dim tempGCtrType As String
    tempGCtrType = ";" + GCtrType + ";"
           
    If InStr(1, tempGCtrType, ";1;") > 0 Then
        GETCont.Show
    ElseIf InStr(1, tempGCtrType, ";3;") > 0 Then
        SINGLECONTRACT.Show
    ElseIf InStr(1, tempGCtrType, ";4;") > 0 Then
        FRM_NEW_SINGLE_ENTRY.Show
    ElseIf InStr(1, tempGCtrType, ";5;") > 0 Then
        GET_CONTBS.Show
    ElseIf InStr(1, tempGCtrType, ";6;") > 0 Then
        Call Get_Selection(12)
        Frm_GetContSauda.Show
    'ElseIf InStr(1, tempGCtrType, ";8;") > 0 Then
    '    Call Get_Selection(12)
    '    FrmCont5.Show
    ElseIf InStr(1, tempGCtrType, ";2;") > 0 Then
        CTRBUYSELL.Show
    End If
End Sub

Public Sub CONTRACTENTRY7_Click()
    Call Menu_Check("CONTRACTENTRY7", "T")
    Call PERMISSIONS("CONTRACTENTRY7")
    Call Get_Selection(10)
       
    Dim tempGCtrType As String
    tempGCtrType = ";" + GCtrType + ";"
           
    If InStr(1, tempGCtrType, ";7;") > 0 Then
        Call Get_Selection(12)
        frmcont4.Show
    End If
End Sub

Private Sub contractlogreport_Click()
    Call Menu_Check("contractlogreport", "R")
    MFormat = "Contract Log"
    Contractlog.Show
End Sub

Public Sub CONTRACTREG_Click()
    Call Menu_Check("CONTRACTREG", "R")
    MFormat = "Contract Register"
    Dim ContRegFrm As New NRPTCONA
    ContRegFrm.Show
End Sub

Private Sub CORRDIV_Click()
    Call Menu_Check("CORRDIV", "T")
    Call PERMISSIONS("CORRDIV")
    FrmCrDe.Show
    Call Get_Selection(10)
End Sub

Private Sub CrURpt_Click()
    Call Menu_Check("CrURpt", "R")
    MFormat = "Credit Utilisaction Report"
    CreditUtifrm.Show
End Sub
Private Sub DataCombo1_DblClick(Area As Integer)
    Call DATA_ACCESS
End Sub
Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err1
    If KeyCode = 27 Then
        Call Get_Selection(10)
        GETMAIN.ActiveForm.Fb_Press = 0
        DataCombo1.Visible = False
    End If
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    DataCombo1.Visible = False
End Sub
Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call DATA_ACCESS
End Sub
Public Sub DATEWSCONTLIST_Click()
    Call Menu_Check("DATEWSCONTLIST", "R")
    MFormat = "Date wise Contract"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub DBkp_Click()
    Call Menu_Check("DBkp", "U")
    Call DataBackUp
End Sub
Private Sub dcm_Click()
    Call PrtUpfrm.Show
End Sub
Private Sub DELETErec_Click()
    Call Toolbar1_Buttons_Click(3)
End Sub
Private Sub dtmrgnrpt_Click()
    Call Menu_Check("dtmrgnrpt", "R")
    MFormat = "Date wise Margin Report"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Private Sub exit_Click()
    Unload Me
End Sub
Public Sub ftptb_Click()
    MFormat = "Trial Balance 2"
    GETQYRPT.Show
End Sub
Public Sub Exstp_Click()
    Call Menu_Check("Exstp", "M")
    ExMFrm.Show
End Sub
Public Sub FmlyStup_Click()
    
    Call Menu_Check("FmlyStup", "M")
    Call PERMISSIONS("FmlyStup")
    FmlyFrm.Show
    Call Get_Selection(10)
End Sub
Public Sub genled_Click()
    Call Menu_Check("genled", "R")
    MFormat = "General Ledger"
    GETQYRPT.Show
End Sub
Public Sub GenQry_Click()
    Call Menu_Check("GenQry", "Q")
    rmdyfrm.Show
End Sub
Public Sub ITEMSETUP_Click()
    Call Menu_Check("ITEMSETUP", "M")
    Call PERMISSIONS("ITEMSETUP")
    ITEMMAST.Show
    Call Get_Selection(10)
End Sub
Private Sub listrec_Click()
    Call Toolbar1_Buttons_Click(8)
End Sub
Private Sub lofoff_Click()
    'LogOffBool = True
    'Call LogOff
    DoEvents
    GETMAIN.MousePointer = 0
    StatusBar1.Panels(2).text = "Database backup in process..."
    Call DataBackUp_OnLogOff("Y")
    DoEvents
    StatusBar1.Panels(2).text = ""
End Sub
Public Sub loginoff_Click()
    MFormat = "Login / Logoff Detail"
    GETQYRPT.Show
End Sub
Public Sub Marsry_Click()
    Call Menu_Check("Marsry", "R")
    MFormat = "Margin Report"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
   
    GETMAIN.MousePointer = 0
    DoEvents
    
    'GETMAIN.StatusBar1.Panels(2).text = "Database backup in process..."
    'GETMAIN.PictureMenu.Visible = True
    DoEvents
    If FlagDataBackup = "Y" Then
        If MsgBox("Do you want to take database backup?" & String(10, " "), vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
        Call DataBackUp_OnLogOff("Y")
        End If
    End If

    DoEvents
   ' StatusBar1.Panels(2).text = ""
    GETMAIN.PictureMenu.Visible = False
    If LogOffBool = False Then
        If MsgBox(String(5, " ") & "Exit Application ?" & String(10, " "), vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
            Cancel = 0
            If CNNERR = True Then
                Cnn.RollbackTrans
            End If
            If Cnn.State = 1 Then
                Cnn.close
            End If
        Else
            'MenuOpt.Show
            MenuOptfrm.Show
            Cancel = 1
        End If
    End If
End Sub
Public Sub MENUACCSTT_Click()
    Call Menu_Check("MENUACCSTT", "R")
    MFormat = "Account Statement"
    Dim AcsttFrm As New NRPTCONA
    AcsttFrm.Show
End Sub
Public Sub MenuInvPrint_Click()
    Call Menu_Check("MenuInvPrint", "R")
    MFormat = "Invoice Printing"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Private Sub mnidetldg_Click()
    Call Menu_Check("mnidetldg", "R")
    MFormat = "Detailed Ledger"
    GETQYRPT.Show
End Sub

Public Sub mnuaccsmry_Click()
    Call Menu_Check("mnuaccsmry", "R")
    MFormat = "Account Statement Summary"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Private Sub mnuacsttsumm_Click()
    Call Menu_Check("mnuacsttsumm", "R")
    MFormat = "Billwise Account Statement Summary"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub mnuactivity_Click()
    Call Menu_Check("mnuactivity", "R")
    MFormat = "Activity Summary"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub mnuannual_Click()
Call Menu_Check("mnuannual", "R")
    MFormat = "Annual Global Transaction Statement"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub mnubillacstt_Click()
    Call Menu_Check("mnubillacstt", "R")
    MFormat = "Billwise Account Statement"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnubillsmry_Click()
    Call Menu_Check("mnubillsmry", "R")
    MFormat = "Bill Summary"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnublist_Click()
    Call Menu_Check("mnublist", "R")
    MFormat = "Invoice GST Report"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnubrbrok_Click()
    Call Menu_Check("mnubrbrok", "R")
    MFormat = "Branch Wise Brokerage Report"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Public Sub mnubrshare_Click()
    Call Menu_Check("mnubrshare", "R")
    MFormat = "Branch Wise Sharing Report"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Public Sub MnuCalc_Click()
    Call CALC_EXE
End Sub
Public Sub mnucnote_Click()
    Call Menu_Check("mnucnote", "R")
    MFormat = "Contract Note"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnudaily_Click()
Call Menu_Check("mnudaily", "R")
    MFormat = "Daily Activity Statement"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnudass_Click()
    Call Menu_Check("mnudass", "R")
    MFormat = "Bill Summary With Sharing"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnudata_Click()
    Call Menu_Check("mnudata", "T")
    frmdata.Show
End Sub
Public Sub mnudclmgn_Click()
    Call Menu_Check("mnudclmgn", "R")
    MFormat = "Daily Client Wise Margin"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Public Sub mnuexbrok_Click()
    Call Menu_Check("mnuexbrok", "M")
    Call PERMISSIONS("mnuexbrok")
    FrmExBrok2.Show
End Sub

Public Sub mnuexsbrok_Click()
    Call PERMISSIONS("mnuexsbrok")
    Call Menu_Check("mnuexsbrok", "S")
    FrmExSBrok.Show
End Sub

Public Sub mnuNewStm_Click()
    Call Menu_Check("mnuNewStm", "R")
    MFormat = "Billwise Summary"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub mnuptyclose_Click()
    Call Menu_Check("mnuptyclose", "T")
    Call PERMISSIONS("mnuptyclose")
    FrmPtyClose.Show
    Call Get_Selection(10)
End Sub

Private Sub mnuqaccount_Click()
FRMACCOUNT.Show
End Sub

Private Sub mnuqstand_Click()
    'GRptViewType = "Query on Standing"
    'Call Menu_Check("mnuqstand", "Q")
    'Call PERMISSIONS("mnuqstand")
    'RptView.Show
End Sub

Private Sub mnuqtrade_Click()
    'GRptViewType = "Query on Trade"
    'Call Menu_Check("mnuqtrade", "Q")
    'Call PERMISSIONS("mnuqtrade")
    'RptView.Show
End Sub
Private Sub mnusaudalot_Click()
    Call Menu_Check("mnusaudalot", "S")
    Call PERMISSIONS("mnusaudalot")
    'frmsaudalot.Show
End Sub

Private Sub mnutele_Click()
Dim LFilepath As String
Dim LFilePath2 As String
    LFilePath2 = App.Path & "\TELEGRAMAPI\ "
    LFilepath = App.Path & "\TELEGRAMAPI\" & "python " & "TELE_API.PY"
    'Shell LFilepath, vbNormalFocus
    'Shell "C:\Windows\System32\cmd.exe /c" & LFilepath & "", vbNormalFocus
    
    'Shell "C:\Windows\System32\cmd.exe /C", vbNormalFocus
    'Shell "cd " & CurDir & ""
    'MsgBox CurDir
    Shell "C:\Windows\System32\cmd.exe /c" & LFilepath & "", vbMaximizedFocus
   ' ShellExecute 0, "runas", LFilepath, "/admin", lfilepath2, vbNormalFocus
    'If InStr(Command, "/admin") = 0 Then
    '    ShellExecute 0, "runas", LFilepath, Command & "/admin", vbNullString, vbNormalFocus
    'Else
    '   ShellExecute 0, "runas", LFilepath, Command, vbNullString, vbNormalFocus
    'End If
    Exit Sub
End Sub

Private Sub mnuTradeDelete_Click()
Call Menu_Check("mnuTradeDelete", "R")
    MFormat = "Trade Deletion"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub


Private Sub mnutrfbal_Click()
Call Menu_Check("mnutrfbal", "E")
    Call PERMISSIONS("mnutrfbal")
    FrmTrfBal.Show
End Sub

Public Sub munmarginfiledaily_click()
    Call Menu_Check("munmarginfiledaily", "R")
    MFormat = "Daily Margin File"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnudeposit_Click()
    Call Menu_Check("mnudeposit", "R")
    MFormat = "ODIN Deposit Upload File Generation"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Private Sub mnumarginfile_Click()
    Call Menu_Check("mnumarginfile", "R")
    MFormat = "Daily Margin File Upload"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnudtst_Click()
    Call Menu_Check("mnudtst", "R")
    MFormat = "Date wise Standing"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Public Sub mnuINVLIST_Click()
    Call Menu_Check("mnuINVLIST", "R")
    MFormat = "Invoice List"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub mnuinvset_Click()
Call Menu_Check("MNUINVSET", "T")
    Call PERMISSIONS("MNUINVSET")
    MFormat = "Invoice Generation Settlement Wise"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub MNUInvWsLedg_Click()
    Call Menu_Check("MNUInvWsLedg", "R")
    MFormat = "Invoice wise Ledger"
    GETQYRPT.Show
End Sub
Private Sub mnumrgnrt_Click()
    Call Menu_Check("mnumrgnrt", "R")
    MFormat = "Margin Rate List"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Private Sub mnunarr_Click()
    Call Menu_Check("MnuNarr", "M")
    Call PERMISSIONS("MnuNarr")
    FrmNarr.Show
    Call Get_Selection(10)

End Sub
Private Sub mnuopening_Click()
Call Menu_Check("mnuopening", "Q")
    FrmOpBAl.Show
End Sub

Public Sub mnuoutstanding_Click()
    Call Menu_Check("mnuoutstanding", "R")
    MFormat = "Partywise Outstanding"
    RPTTB.Show
End Sub

Public Sub mnuinterest_Click()
    Call Menu_Check("mnuinterest", "R")
    MFormat = "Partywise Interest Collection"
    RPTTB.Show
End Sub

Public Sub mnuQTB_Click()
    Call Menu_Check("mnuQTB", "Q")
    QRYTB.Show
End Sub
Public Sub mnusubshare_Click()
    Call Menu_Check("mnusubshare", "R")
    MFormat = "Sub Brokerage && Sharing List"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnutrade_Click()
    Call Menu_Check("mnutrade", "R")
    MFormat = "Trade File Generation"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub mnutrdconfirm_Click()
    Call Menu_Check("mnutrdconfirm", "R")
    FRM_TRDCONFIRM.Show
    'frm_trdmatch.Show
    Call Get_Selection(12)
End Sub
Public Sub mnutrdreg_Click()
    Call Menu_Check("mnutrdreg", "R")
    MFormat = "Trade Register"
    
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Public Sub mnutrialdt_Click()
Call Menu_Check("mnutrialdt", "R")
    MFormat = "MTM SUMMARY"
    RPTTB.Show
End Sub
Public Sub mnutrurnover_Click()
    Call Menu_Check("mnutrurnover", "R")
    MFormat = "Turnover Report"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub MST_Click()
    Call Menu_Check("MST", "R")
    MFormat = "Maturitywise Standing Report"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub Notepad_Click()
    Call NOTEPAD_EXE
End Sub
Public Sub opntb_Click()
    MFormat = "Trial Balance 1"
    GETQYRPT.Show
End Sub

Private Sub OrdRegRPT_Click()
    Call Menu_Check("OrdRegRpt", "R")
    MFormat = "Order Register Report"
    OrderRegister.Show
End Sub

Public Sub PackUpData_Click()
    Call Menu_Check("packupdata", "U")
    GETMAIN.Caption = "Please Wait"
    MFormat = "PackUp Data"
    YrUpdt.Show
    
End Sub

Public Sub PANDLMENU_Click()
    Call Menu_Check("PANDLMENU", "R")
    MFormat = "Profit & Loss"
    RPTPANDL.Show
End Sub
Private Sub partyage_Click()
    Call Menu_Check("partyage", "R")
    MFormat = "Party Ageing Analysis (Annexure 4)"
    RPTTB.Show
End Sub

Private Sub PRINTERSETUP_Click()
On Error Resume Next
    RDCREPO.PRINTERSETUP (GETMAIN.hwnd)
End Sub
Private Sub PRINTTOGGLE_Click()
    If PRINTTOGGLE.Checked = True Then
        PRINTTOGGLE.Checked = False
        PRINTTOGGLE.Caption = "Print Toggle Off"
        MsgBox "Print Toggle is OFF ( Window Base Printing ).", vbInformation, "Message"
    Else
        PRINTTOGGLE.Checked = True
        PRINTTOGGLE.Caption = "Print Toggle On"
        MsgBox "Print Toggle is ON ( DOS Base Printing ).", vbInformation, "Message"
    End If
End Sub

Private Sub QrtonBill_Click()
    Call Menu_Check("QrtonBill", "Q")
    MFormat = "Query on Bill"
    frmINVD.Show
End Sub

Private Sub QTradeFinder_Click()
    Call Menu_Check("QTradeFinder", "Q")
    MFormat = "Query Trade Finder"
    FrmTradeFinder.Show
End Sub

Public Sub Reindex_Click()
    Call Menu_Check("Reindex", "U")
    Me.MousePointer = 11
    On Error Resume Next
    Cnn.Execute "EXECUTE ReIndexTable"
    Me.MousePointer = 0
    MsgBox "Reindexing completed."
End Sub
Public Sub RPTBROKSMRY_Click()
    Call Menu_Check("RPTBROKSMRY", "R")
    MFormat = "Brokerage Summary"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub RPTSLotChange_Click()
    Call Menu_Check("Sdatrf", "R")
    MFormat = "Sauda Lot Change"
    frmSaudaLotChange.Show
End Sub
Public Sub RtLst_Click()
    Call Menu_Check("RtLst", "R")
    MFormat = "Rate List"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Public Sub rwb_Click()
    Call Menu_Check("rwb", "U")
    GETMAIN.Caption = "Please Wait"
    RWBook.Show
    'Call CORECTING_BOOKS
    GETMAIN.Caption = GCompanyName
End Sub
Private Sub MDIForm_Load()
    Set ADOUser = Nothing
    Set ADOUser = New ADODB.Recordset
    mysql = "SELECT * FROM USER_RIGHTS WHERE USER_NAME='" & GUserName & "'"
    ADOUser.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
End Sub
Private Sub modifyrec_Click()
    Call Toolbar1_Buttons_Click(2)
End Sub
Public Sub SAUDAMAST_Click()
    Call Menu_Check("SAUDAMAST", "M")
    Call PERMISSIONS("SAUDAMAST")
    FrmSauda.Show
    Call Get_Selection(10)
End Sub
Public Sub SAUDAWSSTND_Click()
    Call Menu_Check("SAUDAWSSTND", "R")
    MFormat = "Standing Statement"
    'MFormat = "Sauda wise Standing"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub
Private Sub saverec_Click()
    Call Toolbar1_Buttons_Click(4)
End Sub
Private Sub Sdatrf_Click()
    Call Menu_Check("Sdatrf", "R")
    MFormat = "Sauda Transfer"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub Sdatrf_dt_Click()
    Call Menu_Check("Sdatrf_dt", "R")
    MFormat = "Sauda Transfer DD"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub serchrec_Click()
On Error GoTo Error1
    Call Toolbar1_Buttons_Click(7)
Error1: If err.Number <> 0 Then
        End If
End Sub
Public Sub SETMASTER_Click()
    Call Menu_Check("SETMASTER", "M")
    'If GSoft <> 1000 Then
        SETLEMNT.Show
    'End If
    Call Get_Selection(12)
End Sub
Private Sub Singlecont_Click()
Call Menu_Check("Singlecont", "T")
SINGLECONTRACT.Show
Call Get_Selection(10)
End Sub

Public Sub swmrt_Click()
    Call Menu_Check("swmrt", "R")
    MFormat = "Sauda wise missing rate"
    Dim newfrm As New NRPTCONA
    newfrm.Show
End Sub

Private Sub SysLockDate_Click()
    GETMAIN.Caption = "Please Wait"
    MFormat = "System Lock Date"
    frmSysLDt.Show
End Sub
Private Sub Timer1_Timer()
    LTimeCount = LTimeCount + 1
End Sub
Private Sub Toolbar1_Buttons_Click(Index As Integer)
    Call ButtonClick(Index)
End Sub

Public Sub udtb_Click()
    Call Menu_Check("udtb", "R")
    Call PERMISSIONS("udtb")
    MFormat = "Trial Balance 3"
    RPTTB.Show
End Sub
Public Sub UrSetup_Click()
    Call Menu_Check("UrSetup", "M")
    Call PERMISSIONS("UrSetup")
    USetup.Show
    Call Get_Selection(10)
End Sub
Public Sub VCHENT_Click()
    Call Menu_Check("VCHENT", "T")
    Call PERMISSIONS("VCHENT")
    If GVoucherFormat = "Y" Then 'new
        VouFrmNew.Show
    Else
        VouFrm.Show 'old
    End If
    'FrmVoucher.Show
    Call Get_Selection(10)
End Sub
Sub SEARCH_RECORD()
    On Error GoTo Error1
    Call Get_Selection(Fb_Press)
    old_frm = CStr(GETMAIN.ActiveForm.NAME)
    Sendkeys "%{DOWN}"
Error1: If err.Number <> 0 Then
            MsgBox err.Description, vbInformation, err.Number
        End If
End Sub

Private Sub voucherLog_Click()
    Call Menu_Check("voucherLog", "R")
    MFormat = "Voucher Log"
    Contractlog.Show
End Sub

Public Sub voulist_Click()
    Call Menu_Check("voulist", "R")
    MFormat = "Voucher List"
    GETACRPT.Show
End Sub
Sub add_record()
    On Error GoTo Error1
    
    Select Case GETMAIN.ActiveForm.NAME
    Case "GETACNT"
        GETACNT.add_record
    Case "FrmPtyClose"
        FrmPtyClose.Add_Rec
        FrmPtyClose.Fb_Press = 1
    Case "CTRBUYSELL"
        CTRBUYSELL.Fb_Press = 1
        CTRBUYSELL.Add_Rec
    Case "GET_CONTBS"
        GET_CONTBS.Fb_Press = 1
        GET_CONTBS.Add_Rec
    Case "FrmNarr"
        FrmNarr.Add_Rec
    Case "GETCOMP"
        GETCOMP.add_record
    Case "FRM_NEW_SINGLE_ENTRY"
        FRM_NEW_SINGLE_ENTRY.Fb_Press = 1
        FRM_NEW_SINGLE_ENTRY.Add_Rec
    Case "FrmCloseRate"
        FrmCloseRate.Add_Rec
    Case "ITEMMAST"
        ITEMMAST.Add_Rec
    Case "FmlyFrm"
        FmlyFrm.Add_Rec
    Case "ExMFrm"
        ExMFrm.Add_Rec
    Case "FrmSauda"
        FrmSauda.Add_Rec
    Case "AccGrp"
        AccGrp.add_record
    Case "VouFrm"
        VouFrm.add_record
    Case "VouFrmNew"
        VouFrmNew.add_record
    Case "GETCont"
        GETCont.Fb_Press = 1
        GETCont.Add_Rec
    Case "NRPTCONA"
        NRPTCONA.Fb_Press = 1
        NRPTCONA.Add_Rec
        NRPTCONA.Label1 = "Upto Date"
    Case "FrmExBrok"
        FrmExBrok2.ADD_NEW
    Case "FrmExSBrok"
        FrmExSBrok.ADD_NEW
    Case "USetup"
        USetup.ADD_NEW_RECORD
    End Select
Error1: If err.Number <> 0 Then
        End If
End Sub
Sub MODIFY_RECORD()
On Error GoTo Error1
    Call Get_Selection(2)
    Select Case GETMAIN.ActiveForm.NAME
    Case "AccGrp"
        AccGrp.MODIFY_REC
    Case "FrmPtyClose"
        FrmPtyClose.Fb_Press = 2
        FrmPtyClose.MODIFY_REC
    Case "CTRBUYSELL"
        CTRBUYSELL.Fb_Press = 2
        CTRBUYSELL.Add_Rec
    Case "GET_CONTBS"
        GET_CONTBS.Fb_Press = 2
        GET_CONTBS.Add_Rec
    Case "VouFrm"
        FRM_VCH.Show
        FRM_VCH.Fb_Press = 2
    Case "VouFrmNew"
        FRM_VCH.Show
        FRM_VCH.Fb_Press = 2
    Case "FrmCloseRate"
        FrmCloseRate.Fb_Press = 2
        FrmCloseRate.MODIFY_REC
    Case "FRM_NEW_SINGLE_ENTRY"
        FRM_NEW_SINGLE_ENTRY.Fb_Press = 2
        FRM_NEW_SINGLE_ENTRY.Add_Rec
    Case "GETCont"
        GETCont.Fb_Press = 2
        GETCont.Add_Rec
    Case "FrmExBrok2"
        FrmExBrok2.ADD_NEW
    Case "GETCOMP"
        GETCOMP.COMPANY_ACCESS
    Case "GETACNT"
        GETACNT.ACCOUNT_ACCESS
    Case "ITEMMAST"
        ITEMMAST.MODIFY_REC
    Case "FmlyFrm"
        FmlyFrm.MODIFY_REC
    Case "FrmSauda"
        FrmSauda.MODIFY_REC
    Case "ExMFrm"
        ExMFrm.MODIFY_REC
    Case "USetup"
        USetup.DATA_ACCESS
    Case "FrmExSBrok"
        FrmExSBrok.ADD_NEW
    End Select
    old_frm = CStr(GETMAIN.ActiveForm.NAME)
Error1: If err.Number <> 0 Then
        End If
End Sub
Sub Delete_Record()
On Error GoTo Error1
    Call Get_Selection(3)
    If GETMAIN.ActiveForm.NAME = "GETACNT" Then
        GETACNT.Delete_Record
    ElseIf GETMAIN.ActiveForm.NAME = "FrmPtyClose" Then
        FrmPtyClose.DELETE_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FrmNarr" Then
        FrmNarr.Delete_Record
    ElseIf GETMAIN.ActiveForm.NAME = "VouFrm" Then
        FRM_VCH.Show
        FRM_VCH.Fb_Press = 3
    ElseIf GETMAIN.ActiveForm.NAME = "VouFrmNew" Then
        FRM_VCH.Show
        FRM_VCH.Fb_Press = 3
    ElseIf GETMAIN.ActiveForm.NAME = "GETCont" Then
        GETCont.Fb_Press = 3
        GETCont.Add_Rec
    'ElseIf GETMAIN.ActiveForm.NAME = "SINGLECONTRACT" Then
    '    SINGLECONTRACT.Fb_Press = 3
     '   SINGLECONTRACT.ADD_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FRM_NEW_SINGLE_ENTRY" Then
        FRM_NEW_SINGLE_ENTRY.Fb_Press = 3
        FRM_NEW_SINGLE_ENTRY.Add_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "CTRBUYSELL" Then
        CTRBUYSELL.Fb_Press = 3
        CTRBUYSELL.Add_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "GET_CONTBS" Then
        GET_CONTBS.Fb_Press = 3
        GET_CONTBS.Add_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "AccGrp" Then
        AccGrp.Delete_Record
    ElseIf GETMAIN.ActiveForm.NAME = "NRPTCONA" Then
        GETMAIN.ActiveForm.Fb_Press = 3
        GETMAIN.ActiveForm.Label1 = "From Date"
        GETMAIN.ActiveForm.Add_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "FrmCloseRate" Then
        FrmCloseRate.DELETE_REC
    End If
    old_frm = CStr(GETMAIN.ActiveForm.NAME)
Error1: If err.Number <> 0 Then
End If
End Sub
Sub Save_Record()
On Error GoTo Error1
    GETMAIN.MousePointer = 11
    If GETMAIN.ActiveForm.NAME = "GETACNT" Then
        GETACNT.Save_Record
    ElseIf GETMAIN.ActiveForm.NAME = "FrmPtyClose" Then
        FrmPtyClose.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "PrtUpfrm" Then
        PrtUpfrm.UpdateParty
    ElseIf GETMAIN.ActiveForm.NAME = "FrmNarr" Then
        FrmNarr.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "GETCOMP" Then
        GETCOMP.Save_Record
    ElseIf GETMAIN.ActiveForm.NAME = "FrmCloseRate" Then
        FrmCloseRate.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "ITEMMAST" Then
        ITEMMAST.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "FmlyFrm" Then
        FmlyFrm.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "ExMFrm" Then
        ExMFrm.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "FrmSauda" Then
        FrmSauda.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "AccGrp" Then
        AccGrp.Save_Record
    ElseIf GETMAIN.ActiveForm.NAME = "VouFrm" Then
        VouFrm.Save_Record
    ElseIf GETMAIN.ActiveForm.NAME = "VouFrmNew" Then
        VouFrmNew.Save_Record
    ElseIf GETMAIN.ActiveForm.NAME = "GETCont" Then
        GETCont.Save_Rec
        
    ElseIf GETMAIN.ActiveForm.NAME = "FRM_NEW_SINGLE_ENTRY" Then
        FRM_NEW_SINGLE_ENTRY.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "CTRBUYSELL" Then
        CTRBUYSELL.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "GET_CONTBS" Then
        GET_CONTBS.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "FrmExBrok2" Then
        FrmExBrok2.Save_Rec
    ElseIf GETMAIN.ActiveForm.NAME = "USetup" Then
        USetup.Save_Record
    ElseIf GETMAIN.ActiveForm.NAME = "FrmExSBrok" Then
        FrmExSBrok.Save_Rec
    End If

Error1: If err.Number <> 0 Then
            MsgBox err.Description, vbCritical, err.Number
            'Resume
        End If

    GETMAIN.MousePointer = 0
    StatusBar1.Panels(2).text = "Ready"
End Sub
Sub CANCEL_RECORD()
On Error GoTo Error1
    GETMAIN.DataCombo1.Visible = False
    Call Get_Selection(10)
    If GETMAIN.ActiveForm.NAME = "GETACNT" Then
        GETACNT.CANCEL_RECORD
    ElseIf GETMAIN.ActiveForm.NAME = "GETCOMP" Then
        GETCOMP.CANCEL_RECORD
    ElseIf GETMAIN.ActiveForm.NAME = "FrmCloseRate" Then
        FrmCloseRate.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "ITEMMAST" Then
        ITEMMAST.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FmlyFrm" Then
        FmlyFrm.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "ExMFrm" Then
        ExMFrm.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FrmSauda" Then
        FrmSauda.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "AccGrp" Then
        AccGrp.CANCEL_RECORD
    ElseIf GETMAIN.ActiveForm.NAME = "VouFrm" Then
        VouFrm.CANCEL_RECORD
    ElseIf GETMAIN.ActiveForm.NAME = "VouFrmNew" Then
        VouFrmNew.CANCEL_RECORD
    ElseIf GETMAIN.ActiveForm.NAME = "GETCont" Then
        GETCont.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FRM_NEW_SINGLE_ENTRY" Then
        FRM_NEW_SINGLE_ENTRY.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "CTRBUYSELL" Then
           CTRBUYSELL.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "GET_CONTBS" Then
           GET_CONTBS.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FrmExBrok2" Then
        FrmExBrok2.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FrmExSBrok" Then
        FrmExSBrok.CANCEL_REC
    ElseIf GETMAIN.ActiveForm.NAME = "USetup" Then
        USetup.CANCEL_RECORD
    ElseIf GETMAIN.ActiveForm.NAME = "FrmPtyClose" Then
        FrmPtyClose.CANCEL_REC
    End If
    StatusBar1.Panels(2).text = "Ready"
Error1: If err.Number <> 0 Then
        End If
End Sub
Sub CLOSE_RECORD()
On Error GoTo Error1
    Dim LActiveForm As Form
    Set LActiveForm = Me.ActiveForm
   Fb_Press = CByte(0)
   GETMAIN.DataCombo1.Visible = False
   GETMAIN.Caption = "Please Wait"
   Unload LActiveForm
   If Not LenB(GETMAIN.ActiveForm.NAME) <> 0 Then Call Get_Selection(10)
   If GETMAIN.ActiveForm.NAME = "MenuOptfrm" Then Call Get_Selection(12)
Error1: If err.Number = 91 Then
            Call Get_Selection(12)
            'MenuOpt.Show
            'MenuOpt.WindowState = 2
            MenuOptfrm.Show
            MenuOptfrm.WindowState = 2
        End If
        old_frm = vbNullString
        GETMAIN.Caption = GCompanyName
        StatusBar1.Panels(2).text = "Ready"
End Sub
Sub SEARCHDB()
On Error GoTo Error1
    Call Get_Selection(0)
Error1: If err.Number <> 0 Then
            MsgBox err.Description, vbInformation, "Error"
        End If
End Sub
Sub DATA_ACCESS()
    If Fb_Press = 2 Then GETMAIN.ActiveForm.Fb_Press = CByte(2)
    If GETMAIN.ActiveForm.NAME = "GETACNT" Then
        Call GETACNT.ACCOUNT_ACCESS

    ElseIf GETMAIN.ActiveForm.NAME = "GETCOMP" Then
        GETCOMP.COMPANY_ACCESS
    ElseIf GETMAIN.ActiveForm.NAME = "ITEMMAST" Then
        ITEMMAST.MODIFY_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FmlyFrm" Then
        FmlyFrm.MODIFY_REC
    ElseIf GETMAIN.ActiveForm.NAME = "ExMFrm" Then
        ExMFrm.MODIFY_REC
    ElseIf GETMAIN.ActiveForm.NAME = "FrmSauda" Then
        FrmSauda.MODIFY_REC
    ElseIf GETMAIN.ActiveForm.NAME = "USetup" Then
        USetup.MUsername = DataCombo1.text
        USetup.DATA_ACCESS
    End If
    Fb_Press = 0
    GETMAIN.DataCombo1.Visible = False
End Sub
Public Sub YrUpdate_Click()
    Call Menu_Check("YrUpdate", "U")
    GETMAIN.Caption = "Please Wait"
    MFormat = "Year Updation"
    YrUpdt.Show
End Sub
Sub ButtonClick(LIndex As Integer)
On Error GoTo Error1
    Select Case LIndex
        Case 1
            Call add_record
        Case 2
            New_Frm = CStr(GETMAIN.ActiveForm.NAME)
            GETMAIN.ActiveForm.Fb_Press = 2
            Fb_Press = CByte(2)
            If GETMAIN.ActiveForm.NAME = "GETContA" Or GETMAIN.ActiveForm.NAME = "GET_CONTBS" Or GETMAIN.ActiveForm.NAME = "CTRBUYSELL" Or GETMAIN.ActiveForm.NAME = "FrmExSBrok" _
             Or GETMAIN.ActiveForm.NAME = "GETCont" Or GETMAIN.ActiveForm.NAME = "SINGLECONTRACT" Or GETMAIN.ActiveForm.NAME = "FRM_NEW_SINGLE_ENTRY" _
             Or GETMAIN.ActiveForm.NAME = "VouFrm" Or GETMAIN.ActiveForm.NAME = "VouFrmNew" Or GETMAIN.ActiveForm.NAME = "AccGrp" Or GETMAIN.ActiveForm.NAME = "FrmCloseRate" Or GETMAIN.ActiveForm.NAME = "frmbrok" _
             Or GETMAIN.ActiveForm.NAME = "FrmExBrok2" Or GETMAIN.ActiveForm.NAME = "GETCOMP" Or GETMAIN.ActiveForm.NAME = "GETACNT" Or GETMAIN.ActiveForm.NAME = "ITEMMAST" _
             Or GETMAIN.ActiveForm.NAME = "FmlyFrm" Or GETMAIN.ActiveForm.NAME = "FrmSauda" Or GETMAIN.ActiveForm.NAME = "ExMFrm" Or GETMAIN.ActiveForm.NAME = "USetup" _
             Or GETMAIN.ActiveForm.NAME = "FRMNARR" Or GETMAIN.ActiveForm.NAME = "webuser" _
             Or GETMAIN.ActiveForm.NAME = "FrmPtyClose" Then
                   GETMAIN.ActiveForm.Fb_Press = CByte(2)
                   Call MODIFY_RECORD

            Else
                New_Frm = CStr(GETMAIN.ActiveForm.NAME)
                Call SEARCH_RECORD
            End If
        Case 3
            New_Frm = CStr(GETMAIN.ActiveForm.NAME)
            Fb_Press = CByte(3)
            GETMAIN.ActiveForm.Fb_Press = CByte(3)
            If GETMAIN.ActiveForm.NAME = "NRPTCONA" Or GETMAIN.ActiveForm.NAME = "VouFrm" Or GETMAIN.ActiveForm.NAME = "VouFrmNew" Or GETMAIN.ActiveForm.NAME = "GET_CONTBS" Or GETMAIN.ActiveForm.NAME = "CTRBUYSELL" Or GETMAIN.ActiveForm.NAME = "GETCont" Or GETMAIN.ActiveForm.NAME = "SINGLECONTRACT" Or GETMAIN.ActiveForm.NAME = "FRM_NEW_SINGLE_ENTRY" Or GETMAIN.ActiveForm.NAME = "FrmCloseRate" Or GETMAIN.ActiveForm.NAME = "frmbrok" Or GETMAIN.ActiveForm.NAME = "FrmExBrok2" Or GETMAIN.ActiveForm.NAME = "AccGrp" Or GETMAIN.ActiveForm.NAME = "FrmNarr" Or GETMAIN.ActiveForm.NAME = "FrmPtyClose" Then
                Call Get_Selection(3)
                GETMAIN.ActiveForm.Fb_Press = CByte(3)
                Call Delete_Record
            ElseIf GETMAIN.ActiveForm.NAME = "GETACNT" Or GETMAIN.ActiveForm.NAME = "ITEMMAST" Or GETMAIN.ActiveForm.NAME = "FmlyFrm" Or GETMAIN.ActiveForm.NAME = "FrmSauda" Or GETMAIN.ActiveForm.NAME = "FrmExSBrok" Or GETMAIN.ActiveForm.NAME = "ExMFrm" Or GETMAIN.ActiveForm.NAME = "USetup" Then
                Call MODIFY_RECORD
            Else
                New_Frm = CStr(GETMAIN.ActiveForm.NAME)
                Call SEARCH_RECORD
            End If
        Case 4
            Fb_Press = CByte(0)
            Call Save_Record
        Case 5
            Fb_Press = CByte(0)
            Call CANCEL_RECORD
        Case 6
            Fb_Press = CByte(0)
            Call CLOSE_RECORD
        Case 8  'LIST
            Call Get_Selection(12)
            If GETMAIN.ActiveForm.NAME = "GETACNT" Then
                Call GETACNT.List_Rec
            ElseIf GETMAIN.ActiveForm.NAME = "ITEMMAST" Then
                Call ITEMMAST.LIST_ITEM
            ElseIf GETMAIN.ActiveForm.NAME = "FrmSauda" Then
                Call FrmSauda.List_Sauda
            ElseIf GETMAIN.ActiveForm.NAME = "SETLEMNT" Then
                Call SETLEMNT.LIST_ITEM
            End If
        Case 9
            GETMAIN.Caption = GCompanyName
        End Select
        
Error1: If err.Number <> 0 Then

        MsgBox err.Description
        End If
End Sub
Public Sub QonBlst_Click()
    Call Menu_Check("QonBlst", "Q")
    MFormat = "Query on Bill List"
    FrmQueryBill1.Show
End Sub
