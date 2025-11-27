VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RPTPANDL 
   BackColor       =   &H00808080&
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5895
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   12615
      Begin VB.CommandButton CANCEL_CMD 
         BackColor       =   &H80000000&
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton OK_CMD 
         BackColor       =   &H80000000&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   5655
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   3240
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
         End
         Begin vcDateTimePicker.vcDTP DTPicker1 
            Height          =   330
            Left            =   2640
            TabIndex        =   6
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37680.7250462963
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "As on Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   743
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00404080&
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
      TabIndex        =   1
      Top             =   0
      Width           =   12615
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Balance Sheet"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   12375
      End
   End
   Begin MSAdodcLib.Adodc CDN_CONTRA_ADO 
      Height          =   375
      Left            =   17040
      Top             =   4800
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "CDN_CONTRA_ADO"
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
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   14400
      TabIndex        =   0
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
End
Attribute VB_Name = "RPTPANDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOp_Bal As Double
Dim RptRec As ADODB.Recordset
Dim RecVchamt As ADODB.Recordset
Dim AccRec As ADODB.Recordset
Sub NEW_TRIALBALANCE()
    On Error GoTo Error1
    Set RDCREPO = Nothing
    Dim LBalance  As Double
    Dim LTillDate As Date
    Dim Account_Code As String
    Dim Account_Name As String
    Dim LGrpName  As String
    GETMAIN.ProgressBar1.Max = AccRec.RecordCount + 2
    GETMAIN.ProgressBar1.Value = 0
    GETMAIN.ProgressBar1.Visible = True
    GETMAIN.PERLBL.Visible = True
    LBalance = 0
    Call RecSet
    OK_CMD.Enabled = False
    

    LTillDate = DTPicker1.Value
    mysql = "SELECT VT.DR_CR, VT.AMOUNT AS AMOUNT, VT.AC_CODE FROM VCHAMT AS VT WHERE VT.COMPCODE=" & GCompCode & " "
    mysql = mysql & " AND  (VT.VOU_DT >= '" & Format(GFinBegin, "yyyy/MM/dd") & "' AND VT.VOU_DT <= '" & Format(LTillDate, "yyyy/MM/dd") & "') "
    mysql = mysql & " AND VT.VOU_TYPE NOT IN  ('M') "
    Set RecVchamt = Nothing
    Set RecVchamt = New ADODB.Recordset
    RecVchamt.Open mysql, Cnn, adOpenKeyset, adLockReadOnly

    If AccRec.RecordCount > 0 Then AccRec.MoveFirst
    Do While Not AccRec.EOF
        LBalance = Val(AccRec!OP_BAL & "")
        Account_Code = AccRec!ac_code & ""
        Account_Name = UCase(AccRec!NAME & "")
        LGrpName = UCase(AccRec!g_name & "")
        RecVchamt.Filter = "AC_CODE='" & AccRec!ac_code & "'"
        Do While Not RecVchamt.EOF
            If UCase(RecVchamt!DR_CR) = "D" Then
                LBalance = LBalance - Val(RecVchamt!AMOUNT & "")
            ElseIf UCase(RecVchamt!DR_CR) = "C" Then
                LBalance = LBalance + Val(RecVchamt!AMOUNT & "")
            End If
            RecVchamt.MoveNext
        Loop
        If Val(Round(LBalance, 2)) <> 0 Then
            With RptRec
                .AddNew
                !Balance = Round(LBalance, 2)
                !ac_code = Account_Code
                !AC_NAME = Account_Name
                !NARRATION = LGrpName
                !QAC_DT = Date: !DEBIT = 0:      !CREDIT = 0:   !CHEQUE_NO = AccRec!Type & "": !CHEQUE_DT = ""
                !BILL_NO = "":  !REFRENCE = "":  !SM_NAME = "": !Group = "":     !Type = ""
                !DR_CR = "":    !OP_BALANCE = 0: !INV_NO = "":  !INV_DT = Date:  !BILL_DT = Date
                !G_TYPE = "":   !G_CAT = ""
                .Update
            End With
        End If
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        Call PERCENTAGE(GETMAIN.ProgressBar1.Max, GETMAIN.ProgressBar1.Value)
        AccRec.MoveNext
    Loop
    GETMAIN.PERLBL.Caption = vbNullString
    If RptRec.RecordCount > 0 Then
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PLSTT.Rpt", 1) 'group wise
        RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Profit & Loss Statement as on " & LTillDate & "'"
        RDCREPO.FormulaFields.GetItemByName("ORG").text = "' " & GCompanyName & "'"
        RDCREPO.FormulaFields.GetItemByName("ADD1").text = "' " & MAdd1 & "'"
        RDCREPO.FormulaFields.GetItemByName("ADD2").text = "' " & GCCity & "'"
        RDCREPO.DiscardSavedData
        RDCREPO.Database.SetDataSource RptRec
        CRViewer1.ZOrder
        CRViewer1.EnableGroupTree = False
        CRViewer1.Width = CInt(GETMAIN.Width - 100)
        CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
        CRViewer1.Top = 0
        CRViewer1.Left = 0
        CRViewer1.ReportSource = RDCREPO
        CRViewer1.Visible = True
        CRViewer1.ViewReport
    End If
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    Set RptRec = Nothing
    OK_CMD.Enabled = True
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error"
    GETMAIN.ProgressBar1.Visible = False: OK_CMD.Enabled = True
    GETMAIN.PERLBL.Caption = vbNullString
End Sub
Private Sub CANCEL_CMD_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC.AC_CODE, AC.NAME, AC.OP_BAL, AG.G_NAME, AG.CODE, AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG WHERE ac.COMPCODE=" & GCompCode & " AND AC.GCODE = AG.CODE AND AG.TYPE IN ('I', 'E') ORDER BY AC.NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Me.Caption = MFormat
    DTPicker1.Value = Date
End Sub
Sub RecSet()    ''Sub Routine to Open Recordset Without Table
    Set RptRec = Nothing
    Set RptRec = New ADODB.Recordset
    RptRec.Fields.Append "QAC_DT", adDate, , adFldIsNullable
    RptRec.Fields.Append "DEBIT", adDouble, , adFldIsNullable
    RptRec.Fields.Append "CREDIT", adDouble, adFldIsNullable
    RptRec.Fields.Append "BALANCE", adDouble, , adFldIsNullable
    RptRec.Fields.Append "NARRATION", adVarChar, 100, adFldIsNullable
    RptRec.Fields.Append "CHEQUE_NO", adVarChar, 10, adFldIsNullable
    RptRec.Fields.Append "CHEQUE_DT", adVarChar, 10, adFldIsNullable
    RptRec.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
    RptRec.Fields.Append "BILL_NO", adVarChar, 11, adFldIsNullable
    RptRec.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
    RptRec.Fields.Append "REFRENCE", adVarChar, 30, adFldIsNullable
    RptRec.Fields.Append "SM_NAME", adVarChar, 30, adFldIsNullable
    RptRec.Fields.Append "GROUP", adVarChar, 50, adFldIsNullable
    RptRec.Fields.Append "TYPE", adVarChar, 5, adFldIsNullable
    RptRec.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
    RptRec.Fields.Append "OP_BALANCE", adDecimal, , adFldIsNullable
    RptRec.Fields.Append "INV_NO", adVarChar, 15, adFldIsNullable
    RptRec.Fields.Append "INV_DT", adDate, , adFldIsNullable
    RptRec.Fields.Append "BILL_DT", adDate, , adFldIsNullable
    RptRec.Fields.Append "G_TYPE", adVarChar, 1, adFldIsNullable
    RptRec.Fields.Append "G_CAT", adVarChar, 30, adFldIsNullable

    RptRec.Open , , adOpenKeyset, adLockOptimistic
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        CRViewer1.Visible = False
        Cancel = 1
    End If
End Sub

Private Sub OK_CMD_Click()
    If DTPicker1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: DTPicker1.SetFocus: Exit Sub
    If DTPicker1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: DTPicker1.SetFocus: Exit Sub
    If MFormat = "Profit & Loss" Then
        Call NEW_TRIALBALANCE
    ElseIf MFormat = "Balance Sheet" Then
        If Check1.Value = 1 Then
            Call Balance_Sheet
        Else
            Call B_SheetNew
        End If
    End If
End Sub
Sub Balance_Sheet()
    On Error GoTo Error1
    Set RDCREPO = Nothing
    Dim Account_Code As String
    Dim Account_Name  As String
    Dim LGrpName  As String
    Dim LBalance As Double
    Dim LTillDate As Date
    Dim PAC_CODE  As String
    Dim LOp_Bal As Double
    GETMAIN.ProgressBar1.Value = 0
    GETMAIN.ProgressBar1.Visible = True
    GETMAIN.PERLBL.Visible = True
    LBalance = 0
    Call RecSet
    OK_CMD.Enabled = False

    LTillDate = DTPicker1.Value
    mysql = "SELECT AC.AC_CODE, AC.NAME,  AC.OP_BAL, AG.G_NAME, AG.CODE, AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG WHERE ac.COMPCODE=" & GCompCode & " AND AC.GCODE=AG.CODE AND AG.TYPE IN ('I', 'E') ORDER BY AC.NAME"
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    PAC_CODE = "'"
    LOp_Bal = 0
    Do While Not AccRec.EOF
        PAC_CODE = PAC_CODE & AccRec!ac_code & "'"

        LOp_Bal = Val(LOp_Bal) + Val(AccRec!OP_BAL)
        AccRec.MoveNext
        If Not AccRec.EOF Then
            PAC_CODE = PAC_CODE & ", '"
        End If
    Loop
    If PAC_CODE = "'" Then
    Else
    mysql = "SELECT VT.DR_CR, SUM(VT.AMOUNT) AS AMOUNT FROM VCHAMT AS VT WHERE VT.COMPCODE=" & GCompCode & " AND VT.VOU_DT >= '" & Format(GFinBegin, "yyyy/MM/dd") & "' AND VT.VOU_DT <= '" & Format(LTillDate, "yyyy/MM/dd") & "' AND VT.AC_CODE IN (" & PAC_CODE & ") AND vt.vou_type not in ('M') GROUP BY VT.DR_CR ORDER BY VT.DR_CR"
    Set RecVchamt = Nothing
    Set RecVchamt = New ADODB.Recordset
    RecVchamt.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    LBalance = 0
    Do While Not RecVchamt.EOF
        If RecVchamt!DR_CR = "D" Then
            LOp_Bal = LOp_Bal - RecVchamt!AMOUNT
        Else
            LOp_Bal = LOp_Bal + RecVchamt!AMOUNT
        End If
        RecVchamt.MoveNext
    Loop
    With RptRec
        .AddNew
        !Balance = Round(LOp_Bal, 2)
        !ac_code = ""
        !AC_NAME = IIf(LOp_Bal < 0, " NET LOSS", " NET PROFIT")
        !NARRATION = "": !QAC_DT = Date
        If LOp_Bal < 0 Then
            !DEBIT = Abs(LOp_Bal)
        Else
            !CREDIT = LOp_Bal
        End If
        
        .Update
    End With
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC.AC_CODE, AC.NAME,  AC.OP_BAL, AG.G_NAME, AG.CODE, AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG WHERE ac.COMPCODE=" & GCompCode & " AND AC.GCODE=AG.CODE AND AG.TYPE IN ('A', 'L') ORDER BY AC.NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    GETMAIN.ProgressBar1.Max = AccRec.RecordCount + 2
    AccRec.MoveFirst
    mysql = "SELECT VT.DR_CR, VT.AMOUNT AS AMOUNT, VT.AC_CODE FROM VCHAMT AS VT WHERE VT.COMPCODE=" & GCompCode & " AND (VT.VOU_DT >= '" & Format(GFinBegin, "yyyy/MM/dd") & "' AND VT.VOU_DT <= '" & Format(LTillDate, "yyyy/MM/dd") & "') AND vt.vou_type not in ('M')"
    Set RecVchamt = Nothing
    Set RecVchamt = New ADODB.Recordset
    RecVchamt.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    Do While Not AccRec.EOF
        LBalance = Val(AccRec!OP_BAL & "")
        Account_Code = AccRec!ac_code & ""
        Account_Name = UCase(AccRec!NAME & "")
        LGrpName = UCase(AccRec!g_name & "")
        RecVchamt.Filter = "AC_CODE='" & AccRec!ac_code & "'"
        Do While Not RecVchamt.EOF
            If UCase(RecVchamt!DR_CR) = "D" Then
                LBalance = LBalance - Val(RecVchamt!AMOUNT & "")
            ElseIf UCase(RecVchamt!DR_CR) = "C" Then
                LBalance = LBalance + Val(RecVchamt!AMOUNT & "")
            End If
            RecVchamt.MoveNext
        Loop
        If Val(Round(LBalance, 2)) <> 0 Then
            With RptRec
                .AddNew
                !Balance = Round(LBalance, 2)
                !ac_code = Account_Code
                !AC_NAME = Account_Name
                !NARRATION = LGrpName
                .Update
            End With
        End If
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        Call PERCENTAGE(GETMAIN.ProgressBar1.Max, GETMAIN.ProgressBar1.Value)

        AccRec.MoveNext
    Loop
    GETMAIN.PERLBL.Caption = vbNullString

    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "PLSTT.Rpt", 1)

    RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Balance Sheet as on " & LTillDate & "'"

    RDCREPO.FormulaFields.GetItemByName("ORG").text = "' " & GCompanyName & "'"
    RDCREPO.FormulaFields.GetItemByName("ADD1").text = "' " & MAdd1 & "'"
    RDCREPO.FormulaFields.GetItemByName("ADD2").text = "' " & GCCity & "'"

    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource RptRec

    CRViewer1.ZOrder
    CRViewer1.EnableGroupTree = False
    CRViewer1.Width = CInt(GETMAIN.Width - 100)
    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Top = 0
    CRViewer1.Left = 0

    CRViewer1.ReportSource = RDCREPO
    CRViewer1.Visible = True
    CRViewer1.ViewReport
    End If
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    Set RptRec = Nothing
    OK_CMD.Enabled = True
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error"
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL.Caption = vbNullString: OK_CMD.Enabled = True
End Sub
Sub B_SheetNew()
    On Error GoTo Error1
    Set RDCREPO = Nothing
    Dim LAcCode As String:          Dim LAcName  As String
    Dim LBalance As Double:         Dim LOpBal As Double
    Dim LPAcCodes  As String:       Dim LGrpName As String
    Dim LTillDate As Date
    
    GETMAIN.ProgressBar1.Value = 0
    GETMAIN.ProgressBar1.Visible = True
    GETMAIN.PERLBL.Visible = True
    LBalance = 0
    LPAcCodes = vbNullString
    Call RecTRlBal
    OK_CMD.Enabled = False
    LTillDate = DTPicker1.Value
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC.AC_CODE, AC.NAME, AC.OP_BAL, AG.G_NAME, AG.CODE, AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG "
    mysql = mysql & " WHERE ac.COMPCODE=" & GCompCode & " AND AC.GCODE=AG.CODE AND AG.TYPE IN ('I', 'E') ORDER BY AC.NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    LOp_Bal = 0
    Do While Not AccRec.EOF
        If LenB(LPAcCodes) > 1 Then
            LPAcCodes = LPAcCodes & ",'" & AccRec!ac_code & "'"
        Else
            LPAcCodes = "'" & AccRec!ac_code & "'"
        End If
        LOpBal = LOpBal + Val(AccRec!OP_BAL)
        AccRec.MoveNext
    Loop
    If LenB(LPAcCodes) > 1 Then
        mysql = "SELECT SUM(CASE  DR_CR  WHEN 'C' THEN VT.AMOUNT WHEN 'D' THEN VT.AMOUNT*-1 END ) AS LAMT FROM VCHAMT AS VT WHERE VT.COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND VT.VOU_DT >= '" & Format(GFinBegin, "yyyy/MM/dd") & "' AND VT.VOU_DT <= '" & Format(LTillDate, "yyyy/MM/dd") & "' "
        mysql = mysql & " AND VT.AC_CODE IN (" & LPAcCodes & ") AND VT.VOU_TYPE NOT IN ('M') "
        Set RecVchamt = Nothing
        Set RecVchamt = New ADODB.Recordset
        RecVchamt.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        LBalance = 0
        If Not RecVchamt.EOF Then
            If Not IsNull(RecVchamt!LAMT) Then
                    LOpBal = LOpBal + RecVchamt!LAMT
            End If
        End If
        With RptRec
            .AddNew
            If LOpBal < 0 Then
                !DAC_NAME = "NET LOSS"
                !DBALANCE = Abs(LOpBal)
                !CAC_NAME = vbNullString
                !CBALANCE = 0
            Else
                !CAC_NAME = "NET PROFIT"
                !CBALANCE = Abs(LOpBal)
                !DAC_NAME = vbNullString
                !DBALANCE = 0
            End If
            .Update
        End With
    End If
    mysql = "SELECT AC.AC_CODE,AC.NAME,AC.OP_BAL,AG.G_NAME,AG.CODE,AG.TYPE FROM ACCOUNTM AS AC, AC_GROUP AS AG WHERE AC.COMPCODE =" & GCompCode & " "
    mysql = mysql & " AND AC.GCODE=AG.CODE AND AG.TYPE IN ('A', 'L') ORDER BY AC.NAME"
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    AccRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    GETMAIN.ProgressBar1.Max = AccRec.RecordCount + 2
    AccRec.MoveFirst
    mysql = "SELECT AC_CODE, SUM(CASE DR_CR WHEN 'C' THEN AMOUNT WHEN 'D' THEN  AMOUNT*-1 END)  AS AMT FROM VCHAMT  WHERE COMPCODE=" & GCompCode & ""
    mysql = mysql & " AND VOU_DT >= '" & Format(GFinBegin, "yyyy/MM/dd") & "' AND  VOU_DT <= '" & Format(LTillDate, "yyyy/MM/dd") & "' "
    mysql = mysql & " AND VOU_TYPE NOT IN ('M')"
    mysql = mysql & " GROUP BY AC_CODE"
    mysql = mysql & " ORDER BY AC_CODE"
    Set RecVchamt = Nothing
    Set RecVchamt = New ADODB.Recordset
    RecVchamt.Open mysql, Cnn, adOpenKeyset, adLockReadOnly

    Do While Not AccRec.EOF
        LBalance = Val(AccRec!OP_BAL & vbNullString)
        LAcCode = AccRec!ac_code & vbNullString
        LAcName = UCase(AccRec!NAME & "")
        LGrpName = UCase(AccRec!g_name & vbNullString)
        RecVchamt.Filter = "AC_CODE='" & AccRec!ac_code & "'"
        If Not RecVchamt.EOF Then
            If UCase(RecVchamt!AMT) < 0 Then
                LBalance = LBalance - Abs(Val(RecVchamt!AMT))
            Else
                LBalance = LBalance + Val(RecVchamt!AMT)
            End If
            RecVchamt.MoveNext
        End If
        If Val(Round(LBalance, 2)) <> 0 Then
            With RptRec
                If LBalance < 0 Then
                    If RptRec.RecordCount <> 0 Then
                        RptRec.MoveFirst
                        RptRec.Find "DAC_NAME = ''", , adSearchForward
                    End If
                    If RptRec.EOF Then
                        .AddNew
                        !DAC_NAME = LAcName
                        !DBALANCE = Abs(LBalance)
                        !CAC_NAME = vbNullString
                        !CBALANCE = 0
                    Else
                        !DAC_NAME = LAcName
                        !DBALANCE = Abs(LBalance)
                    End If
                Else
                    If RptRec.RecordCount <> 0 Then
                        RptRec.MoveFirst
                        RptRec.Find "CAC_NAME = ''", , adSearchForward
                    End If
                    If RptRec.EOF Then
                        .AddNew
                        !CAC_NAME = LAcName
                        !CBALANCE = LBalance
                        !DAC_NAME = vbNullString
                        !DBALANCE = 0
                    Else
                        !CAC_NAME = LAcName
                        !CBALANCE = LBalance
                    End If
                End If
                .Update
            End With
        End If
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        Call PERCENTAGE(GETMAIN.ProgressBar1.Max, GETMAIN.ProgressBar1.Value)
        AccRec.MoveNext
    Loop
    GETMAIN.PERLBL.Caption = vbNullString
    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "BS21.Rpt", 1)
    RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'Balance Sheet as on " & LTillDate & "'"
    RDCREPO.FormulaFields.GetItemByName("ORG").text = "' " & GCompanyName & "'"
    RDCREPO.FormulaFields.GetItemByName("ADD1").text = "' " & MAdd1 & "'"
    RDCREPO.FormulaFields.GetItemByName("ADD2").text = "' " & GCCity & "'"
    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource RptRec
    CRViewer1.ZOrder
    CRViewer1.EnableGroupTree = False
    CRViewer1.Width = CInt(GETMAIN.Width - 100)
    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.ReportSource = RDCREPO
    CRViewer1.Visible = True
    CRViewer1.ViewReport
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    Set RptRec = Nothing
    OK_CMD.Enabled = True
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error"
    
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL.Caption = vbNullString: OK_CMD.Enabled = True
End Sub

Sub RecTRlBal()
    Set RptRec = Nothing: Set RptRec = New ADODB.Recordset
    RptRec.Fields.Append "GName", adVarChar, 100, adFldIsNullable
    RptRec.Fields.Append "DAC_NAME", adVarChar, 100, adFldIsNullable
    RptRec.Fields.Append "CAC_NAME", adVarChar, 100, adFldIsNullable
    RptRec.Fields.Append "DBALANCE", adDouble, , adFldIsNullable
    RptRec.Fields.Append "CBALANCE", adDouble, , adFldIsNullable
    RptRec.Open , , adOpenKeyset, adLockOptimistic
End Sub

