VERSION 5.00
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form AccSummryGrid 
   Caption         =   "Account Summary Grid view"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   18960
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   15495
      Begin VB.CommandButton ExCmd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Exchange"
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton PartyCmd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Parties"
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton SaudaCmd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Contracts"
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton FmlyCmd 
         BackColor       =   &H00FFFFC0&
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
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton OkCmd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Ok"
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
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CancelCmd 
         BackColor       =   &H00FFFFC0&
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
         Left            =   14400
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   495
         Left            =   10440
         TabIndex        =   20
         Top             =   120
         Width           =   2895
         Begin VB.CheckBox ChkFut 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Future"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   23
            Top             =   80
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox ChkOpt 
            BackColor       =   &H00FFC0C0&
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
            Height          =   300
            Left            =   1080
            TabIndex        =   22
            Top             =   80
            Width           =   975
         End
         Begin VB.CheckBox ChkCsh 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2040
            TabIndex        =   21
            Top             =   80
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   495
         Left            =   10560
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   4695
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF80&
            Caption         =   "DateWise"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   1750
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Contract Wise"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2760
            TabIndex        =   18
            Top             =   120
            Width           =   1750
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Net Rate"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3600
            TabIndex        =   17
            Top             =   120
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   495
         Left            =   5040
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   5415
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "UnConfirmed"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1740
            TabIndex        =   15
            Top             =   80
            Width           =   1800
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Confirmed"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   14
            Top             =   80
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3600
            TabIndex        =   13
            Top             =   80
            Width           =   1675
         End
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "AccSummryGrid.frx":0000
         Left            =   645
         List            =   "AccSummryGrid.frx":000A
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "With Sharing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5721
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "AccSummryGrid.frx":001F
         Left            =   2880
         List            =   "AccSummryGrid.frx":0021
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox OptMTMChk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Opt MTM "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13635
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox CshMTMchk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Cash MTM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11442
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "With Sharing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   9615
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "With Standing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3774
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox CreateChk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Create PDF"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton BrokerCmd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Broker"
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "With Summary"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1827
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Trade Conf"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7548
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   600
         TabIndex        =   30
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
         Value           =   42079.7973611111
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   2520
         TabIndex        =   31
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
         Value           =   42079.7973611111
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
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
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
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
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   180
         Width           =   255
      End
   End
End
Attribute VB_Name = "AccSummryGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AllExcodes As Boolean
Dim LSExCodes As String
Dim AllSaudas As Boolean
Dim AllParties As Boolean
Dim AllFmly As Boolean


Private Sub Form_Load()
    Dim LNo As Integer
    Dim TRec As ADODB.Recordset
    Me.Caption = MFormat:       Combo1.ListIndex = 0:           AllExcodes = True:          AllFmly = True
    AllBrokers = True:          AllParties = True:              AllSaudas = True:
    'ExChk.Visible = False:
    'FmlyChk.Visible = False:
    
'    PartyChk.Visible = False:       SaudaChk.Visible = False:   Label4.Caption = MFormat
    vcDTP1.Value = Date:
    'FilterChk = False:
    vcDTP2.MaxDate = GFinEnd:
    'TxtReport.Visible = False
    'TxtReport.text = App.Path & "\RPT\"
    LCURR = "N"
    vcDTP2.Value = GFinEnd:     vcDTP1.MinDate = GFinBegin:
    'vcDTP3.Value = GFinBegin:
    'vcDTP4.Value = DateValue(GFinEnd + 365)
    'Label4.Caption = MFormat:
    Option1.Value = True
    
    If Date >= GFinEnd Then
        vcDTP1.Value = GFinBegin
    Else
        vcDTP1.Value = Date
    End If
    
    If Date <= GFinEnd Then vcDTP2.Value = Date
    
    'Call Get_Selection(12)
    
    ChkFut.Value = 1:    ChkOpt.Value = 1:    ChkCsh.Value = 1
    ChkFut.Visible = False: ChkOpt.Visible = False: ChkCsh.Visible = False
    
    MYSQL = "SELECT DISTINCT INSTTYPE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " "
    LNo = 0
    Set TRec = Nothing:    Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            If TRec!INSTTYPE = "FUT" Then
                ChkFut.Visible = True
                LNo = LNo + 1
            ElseIf TRec!INSTTYPE = "OPT" Then
                ChkOpt.Visible = True
                LNo = LNo + 1
            ElseIf TRec!INSTTYPE = "CSH" Then
                ChkCsh.Visible = True
                LNo = LNo + 1
            End If
            TRec.MoveNext
        Loop
    End If
    Set TRec = Nothing
End Sub

Private Sub OkCmd_Click()

    If vcDTP1.Value > vcDTP2.Value Then
        MsgBox " Invalid Date To Date Should br greater than or equal to FromDate  "
        vcDTP2.SetFocus
        Exit Sub
    End If

    Call ACC_STT_SMRY

End Sub

Sub ACC_STT_SMRY()

    Dim TRec As ADODB.Recordset
    
    LSParties = Get_Parties:
    LSSaudas = Get_Saudas:
    LSExCodes = Get_ExCodes:
    LSInst = Get_Inst
            
    GFBroktype = "N"
    MYSQL = "SELECT TOP 1 BROKTYPE FROM PEXBROK WHERE COMPCODE =" & GCompCode & " AND MBROKTYPE='F'"
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then GFBroktype = "Y"
    
    Set PartyRec = Nothing
    Set PartyRec = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT A.ACCID,A.AC_CODE,A.NAME,A.OP_BAL,B.ADD1,B.CITY,B.PANNO,B.PIN,B.PHONEO,B.PHONER,B.FAX,B.MOBILE,B.PIN,B.SRTAXAPP,B.CTTTYPE,B.RISKMTYPE,B.APPLYON,"
    MYSQL = MYSQL & " S.SAUDACODE,S.SAUDAID,B.PARTYTYPE,B.EMAIL,B.GSTIN,B.STATE,B.STATECODE,B.OPTCUTBROK,B.FUTCUTBROK,"
    MYSQL = MYSQL & " S.INSTTYPE,I.ITEMCODE,S.STRIKEPRICE,S.OPTTYPE,S.MATURITY,I.LOT,S.TRADEABLELOT,I.RISKMAPP,S.BROKLOT,I.SCGROUP,"
    MYSQL = MYSQL & " B.SEBITYPE,S.REFLOT,B.CGST,B.SGST,B.IGST,B.UTT ,EX.EXCODE,EX.LOTWISE,EX.CONTRACTACC ,EX.EXID ,I.ITEMID"
    MYSQL = MYSQL & " FROM ACCOUNTM AS A, ACCOUNTD AS B, CTR_D AS C, ITEMMAST AS I, SAUDAMAST AS S , EXMAST EX"
    MYSQL = MYSQL & " WHERE A.COMPCODE =" & GCompCode & " AND  A.ACCID =B.ACCID "
    MYSQL = MYSQL & " AND S.ITEMID =I.ITEMID AND S.MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    MYSQL = MYSQL & " AND A.ACCID =C.ACCID  AND C.SAUDAID=S.SAUDAID AND  EX.EXID =S.EXID "
    If AllFmly = False And LenB(LSFmlyIDs) > 0 And AllParties = True Then
        MYSQL = MYSQL & " AND A.ACCID IN (SELECT ACCID  FROM ACCFMLYD WHERE FMLYID IN (" & LSFmlyIDs & "))"
    End If
    If AllParties = False And LenB(LSParties) > 0 Then MYSQL = MYSQL & " AND  A.AC_CODE  IN (" & LSParties & ") "
    If AllExcodes = False And LenB(LSExCodes) > 0 Then MYSQL = MYSQL & " AND I.EXID  IN (" & LSExCodes & ")"
    If AllSaudas = False Then MYSQL = MYSQL & " AND S.SAUDAID IN (" & LSSaudas & ") "
    If AllSaudas = True And AllInst = False Then MYSQL = MYSQL & " AND  S.INSTTYPE  IN  (" & LSInst & ") "
    MYSQL = MYSQL & " ORDER BY A.NAME,EX.EXCODE,I.ITEMCODE,S.INSTTYPE,S.MATURITY"
            
'    '--------------
'    PartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'    If PartyRec.EOF Then
'        MsgBox "No Records.", vbCritical: PartyCmd.SetFocus: Exit Sub
'    End If
'    Call AccRecNew
'    Call NewGenerateStatement1
'    '--------------
    
    
    'NEW process applied - 05 FEB 2021 - start
        Call AccRecNew_MODULE
        Set GlobalPartyRec = Nothing
        Set GlobalPartyRec = New ADODB.Recordset
        GlobalPartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        If GlobalPartyRec.EOF Then
            MsgBox "No Records.", vbCritical: PartyCmd.SetFocus: Exit Sub
        End If
        Call New_Generate_Statementl_MODULE(vcDTP2.Value, Option4.Value, Option5.Value, Check10.Value, vcDTP1.Value, Check4.Value, OptMTMChk.Value, CshMTMchk.Value, Option1.Value, Check5.Value, Check9.Value, CreateChk.Value, AllExcodes, LSExCodes)
        'GlobalAccRecSet
    'NEW process applied - 05 FEB 2021 - end
    
        
'    AccRecSet.Filter = adFilterNone
'    If CreateChk.Value = 1 Then
'        MsgBox "Files Exported Successfully"
'        Me.MousePointer = 0
'        GETMAIN.ProgressBar1.Visible = False
'        OkCmd.Enabled = True
'        Exit Sub
'    Else
'        If AccRecSet.EOF Then
'            MsgBox "Record not found"
'        Else
'            AccRecSet.MoveFirst
'            Set TRec = Nothing
'            Set TRec = New ADODB.Recordset
'            Set TRec = AccRecSet.Clone
'            If MFormat = "Sauda Summary" Then
'                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "ACSTSMRY.RPT", 1)
'                MYSQL = "'Sauda Summary From " & vcDTP1.Value & " To " & vcDTP2.Value & "'"
'            ElseIf MFormat = "Branchwise Account Statement Summary" Then
'                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "BRACSTTSMR.RPT", 1)
'                MYSQL = "'Branch wise Account Statement Summary From " & vcDTP1.Value & " To " & vcDTP2.Value & "'"
'            Else
'                If Check4.Value = 1 Then
'                    If Combo2.ListIndex = 1 Then
'                        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "AcsttSmr-Summ-SHARE.RPT", 1)
'                        MYSQL = "'Account Statement Summary From " & vcDTP1.Value & " To " & vcDTP2.Value & "'"
'                    Else
'                        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "AcsttSmr-SHARE.RPT", 1)
'                        MYSQL = "'Account Statement Summary From " & vcDTP1.Value & " To " & vcDTP2.Value & "'"
'
'                    End If
'                ElseIf Combo2.ListIndex = 1 Then
'                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "AcsttSmrSumm.RPT", 1)
'                    MYSQL = "'Account Statement Summary From " & vcDTP1.Value & " To " & vcDTP2.Value & "'"
'                Else
'                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "AcsttSmr.rpt", 1)
'                    MYSQL = "'Account Statement Summary From " & vcDTP1.Value & " To " & vcDTP2.Value & "'"
'                End If
'            End If
'            RDCREPO.DiscardSavedData
'            RDCREPO.Database.SetDataSource TRec
'            RDCREPO.FormulaFields.GetItemByName("COMPANY").text = "'" & GCompanyName & "'"
'            RDCREPO.FormulaFields.GetItemByName("TITLE").text = MYSQL
'            RDCREPO.FormulaFields.GetItemByName("OADD1").text = "'" & GCompanyAdd1 & "'"
'            RDCREPO.FormulaFields.GetItemByName("OADD2").text = "'" & GCompanyAdd2 & "'"
'            RDCREPO.FormulaFields.GetItemByName("OCITY").text = "'" & GCCity & "'"
'            CRViewer1.ZOrder
'            CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
'            CRViewer1.Visible = True
'            CRViewer1.ReportSource = RDCREPO
'            CRViewer1.Zoom 1
'            CRViewer1.ViewReport
'        End If
'        Me.MousePointer = 0
'        GETMAIN.ProgressBar1.Visible = False
'        OkCmd.Enabled = True
'        Exit Sub
'    End If
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number

    Me.MousePointer = 0: OkCmd.Enabled = True
    GETMAIN.ProgressBar1.Visible = False
End Sub
Public Function Get_Saudas() As String
    Dim LFSaudas   As String: Dim LCount As Integer: Dim I As Integer
    
    LFSaudas = vbNullString
    LCount = SaudaList.ListItems.Count
    If AllSaudas = False Then
        For I = 1 To SaudaList.ListItems.Count
            If SaudaList.ListItems(I).Checked = True Then
                LCount = LCount - 1
                If LenB(LFSaudas) <> 0 Then LFSaudas = LFSaudas & ", "
                LFSaudas = LFSaudas & SaudaList.ListItems(I).SubItems(2) & ""
            End If
        Next
        If LCount <> 0 Then
            If LCount <> SaudaList.ListItems.Count Then AllSaudas = False
        Else
            AllSaudas = True
        End If
    End If
    If FilterChk = True Then AllSaudas = False
    Get_Saudas = LFSaudas
End Function

Public Function Get_Inst() As String
    LSInst = vbNullString
    AllInst = True
    OkCmd.Enabled = True
    If ChkFut.Visible = True Then
        If ChkFut.Value = 1 Then
            LSInst = "'FUT'"
        Else
            AllInst = False
        End If
    End If
    If ChkOpt.Visible = True Then
        If ChkOpt.Value = 1 Then
            If Len(LSInst) < 1 Then
                LSInst = "'OPT'"
            Else
                LSInst = LSInst & ",'OPT'"
            End If
        Else
            AllInst = False
        End If
    End If
    If ChkCsh.Visible = True Then
        If ChkCsh.Value = 1 Then
            If Len(LSInst) < 1 Then
                LSInst = "'CSH'"
            Else
                LSInst = LSInst & ",'CSH'"
            End If
        Else
            AllInst = False
        End If
    End If
    If LenB(LSInst) = 0 Then
        MsgBox "Please Select at Least One Instrument"
        OkCmd.Enabled = False
    End If
    Get_Inst = LSInst
End Function


Public Function Get_ExCodes() As String
    Dim LFExCode  As String:    Dim LCount As Integer:  Dim I As Integer
    LFExCode = vbNullString
    If AllExcodes = False Then
        LCount = ExList.ListItems.Count
        For I = 1 To ExList.ListItems.Count
            If ExList.ListItems(I).Checked = True Then
                LCount = LCount - 1
                If LenB(LFExCode) <> 0 Then LFExCode = LFExCode & ", "
                LFExCode = LFExCode & ExList.ListItems(I).ListSubItems(2) & ""
                If MFormat = "Contract Note" Or MFormat = "Daily Margin File Upload" Then LFExCode = "" & ExList.ListItems(I).ListSubItems(2) & ""
            End If
        Next
        If MFormat = "Contract Note" Or MFormat = "Daily Margin File Upload" Then
            If LCount <> ExList.ListItems.Count - 1 And ExList.ListItems.Count <> 0 Then MsgBox "Please Select Only One Exchange"
        End If
    
        If LCount <> 0 Then
            If LCount <> ExList.ListItems.Count Then AllExcodes = False
        Else
            AllExcodes = True
        End If
        If MFormat = "Contract Note" Or MFormat = "Daily Margin File Upload" Or MFormat = "Turnover Report" Then
            If ExRec.RecordCount = 1 Then
                ExRec.MoveFirst
                LFExCode = "" & ExRec!EXID & ""
            End If
        End If
    End If
    If LenB(LFExCode) < 1 Then AllExcodes = True
    Get_ExCodes = LFExCode
End Function


Public Function Get_Parties() As String
    Dim LFParties   As String
    Dim LCount As Integer
    Dim I As Integer
    LSFmlyIDs = Get_FmlyCodes
    LFParties = vbNullString
    LCount = PartyList.ListItems.Count
    If AllParties = False Then
        For I = 1 To PartyList.ListItems.Count
            If PartyList.ListItems(I).Checked = True Then
                LCount = LCount - 1
                If LenB(LFParties) <> 0 Then LFParties = LFParties & ", "
                LFParties = LFParties & "'" & PartyList.ListItems(I).ListSubItems(1) & "'"
            End If
        Next
        If LCount = PartyList.ListItems.Count Then
            AllParties = True
        End If
    End If
    Get_Parties = LFParties
End Function

