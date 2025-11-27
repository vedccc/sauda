VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form FrmTrfBal 
   Caption         =   "Transfer Account Balance"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   19005
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11040
      TabIndex        =   20
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8535
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   8895
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Close Standing Position"
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
         Left            =   3240
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Create Single Voucher"
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
         Left            =   2640
         TabIndex        =   16
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox ChkVouFile 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Create Voucher File"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
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
         Height          =   405
         Left            =   6960
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show From A/c"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Transfer"
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
         Left            =   5040
         TabIndex        =   8
         Top             =   1200
         Width           =   3495
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8493
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Debit Bal"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Credit Bal"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "AccID"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   420
         Left            =   5040
         TabIndex        =   2
         Top             =   120
         Width           =   3495
         _ExtentX        =   6165
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
         Height          =   405
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   43676.4908564815
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   405
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   43676.4908564815
      End
      Begin MSDataListLib.DataCombo BranchCombo 
         Height          =   420
         Left            =   5040
         TabIndex        =   4
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net"
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
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   8160
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   7680
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bal As on Date"
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
         TabIndex        =   12
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Transfer From A/c"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   8535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer To A/c"
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
         TabIndex        =   10
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vou Date"
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Transfer Account Balances"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "FrmTrfBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AccRec As ADODB.Recordset:  Dim LVou_Dt As Date
Dim LToAcc As String:           Dim LAccs As String
Dim LBalDate As Date:           Dim AccFmlyRec As ADODB.Recordset

Private Sub BranchCombo_Validate(Cancel As Boolean)
If LenB(BranchCombo.BoundText) > 0 Then
    If AccFmlyRec.RecordCount > 0 Then
        AccFmlyRec.MoveFirst
        AccFmlyRec.Find "FMLYCODE='" & BranchCombo.BoundText & "'"
        If AccFmlyRec.EOF Then
            BranchCombo.BoundText = vbNullString
            MsgBox "Invalid Branch"
            Cancel = True
        End If
    Else
        BranchCombo.BoundText = vbNullString
        MsgBox "No Branch Available"
        Cancel = True
    End If
End If

End Sub

Private Sub Check1_Click()

Text1.text = 0
Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        If Check1.Value = 1 Then
            ListView1.ListItems.Item(I).Checked = True
            Text1.text = Text1.text + ListView1.ListItems(I).SubItems(3)
            Text1.text = Text1.text + ListView1.ListItems(I).SubItems(2)
        Else
            ListView1.ListItems.Item(I).Checked = False
        End If
    Next I
    DoEvents
End Sub

Private Sub Command1_Click()
Dim LToAcc As String:           Dim LToAccName As String
Dim I As Integer:               Dim LAcCode As String
Dim LAcName As String:          Dim LBal As Double
Dim LPtyDrCr As String:         Dim LCDrCr As String
Dim LMaxVouNo As Long:          Dim LVchNo  As String
Dim TRec As ADODB.Recordset:    Dim LMVNo As Long
Dim LNarr As String:            Dim LFileFlag As Boolean
Dim LFileSystemObject As Scripting.FileSystemObject
Dim LFile As TextStream:        Dim TxtPath As String
Dim TradeLine As String:        Dim LCount As Integer
Dim LChk_Vou As Boolean
Dim LToAccID As Long
Dim LACCID As Long



If Check3.Value = 1 Then
    Dim LSCondate As Date:          Dim MCount As Long
    Dim LSSaudaCode As String:      Dim LSItemCode  As String
    Dim LSItemid  As Integer:       Dim LSExCode  As String
    Dim LSSaudaID As Long:          Dim LSEXID As Integer
    Dim LSConSno As Long:           Dim LQTY As Double
    Dim LRATE As Double:            Dim LSPtyContype As String
    Dim LLOT As Double
    
    LToAcc = DataCombo1.BoundText
    If LenB(LToAcc) < 1 Then
        MsgBox " Please Select To Party"
        DataCombo1.SetFocus
        Exit Sub
    End If
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & LToAcc & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        MsgBox " Please Select Valid To Party"
        DataCombo1.SetFocus
        Exit Sub
    Else
        LToAccName = TRec!NAME
    End If
    LSCondate = vcDTP1.Value
    MCount = Get_Max_ConNo(LSCondate, 0)
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            LAcName = vbNullString
            LAcCode = ListView1.ListItems(I).text
            Set TRec = Nothing
            Set TRec = New ADODB.Recordset
            mysql = "SELECT EXID,EXCODE,ITEMID,ITEMCODE,SAUDAID,SAUDA,CALVAL, SUM(CASE CONTYPE WHEN 'B' THEN QTY WHEN 'S' THEN QTY *-1 END ) AS NETQTY "
            mysql = mysql & " FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & LAcCode & "'"
            mysql = mysql & "  GROUP BY EXID,EXCODE,ITEMID,ITEMCODE,SAUDAID,SAUDA,CALVAL "
            mysql = mysql & "  HAVING SUM(CASE CONTYPE WHEN 'B' THEN QTY WHEN 'S' THEN QTY *-1 END ) <> 0 "
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not TRec.EOF
                
                If Round(TRec!NETQTY, 2) <> 0 Then
                    LQTY = TRec!NETQTY
                    If LQTY > 0 Then
                        LSPtyContype = "S"
                        LQTY = Abs(LQTY)
                    Else
                        LSPtyContype = "B"
                        LQTY = Abs(LQTY)
                    End If
                    MCount = MCount + 1
                    LSSaudaCode = TRec!Sauda
                    LSItemCode = TRec!excode
                    LSExCode = TRec!excode
                    LSSaudaID = TRec!SAUDAID
                    LSItemid = TRec!itemid
                    LSEXID = TRec!exid
                    LLOT = TRec!CALVAL
                    LRATE = SDCLRATE(LSSaudaID, LSCondate, "C")
                    GETMAIN.Label1.Caption = " Trade  " & MCount & ""
                    DoEvents
                    If LQTY <> 0 And LRATE <> 0 Then
                        LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSSaudaID, LSItemid, LSEXID)
                        Call Add_To_Ctr_D(LSPtyContype, LAcCode, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, LAcCode, LQTY, LRATE, LToAcc, vbNullString, vbNullString, vbNullString, Str(MCount), LSExCode, LLOT, 1, vbNullString, "FUT", vbNullString, 0, "0", "Y", LSEXID, LSItemid, LSSaudaID)
                    Else
                        MsgBox "closing rate not found " & LSSaudaCode & ""
                    End If
                End If
                TRec.MoveNext
            Loop
        End If
    Next I
    Cnn.BeginTrans
    CNNERR = True
    Call Update_Charges(vbNullString, vbNullString, vbNullString, vbNullString, LSCondate, GFinEnd, False)
    DoEvents
    Cnn.CommitTrans
    CNNERR = False
    Cnn.BeginTrans
    If BILL_GENERATION(LSCondate, GFinEnd, vbNullString, vbNullString, vbNullString) Then
        Cnn.CommitTrans
        CNNERR = False
    Else
        Cnn.RollbackTrans
        CNNERR = False
    End If

    
    'Call Chk_Billing
    CNNERR = False
    MsgBox "Task complete "
Else
    LChk_Vou = False
    LToAcc = DataCombo1.BoundText
    If LenB(LToAcc) < 1 Then
        MsgBox " Please Select To Party"
        DataCombo1.SetFocus
        Exit Sub
    End If
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = "SELECT ACCID,AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & LToAcc & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        MsgBox " Please Select Valid To Party"
        DataCombo1.SetFocus
        Exit Sub
    Else
        LToAccName = TRec!NAME
        LToAccID = TRec!ACCID
    End If
    LMaxVouNo = Val(Right$(Get_VouNo("JNRL", GFinYear), 7))
    LFileFlag = False
        
    LCount = Val(GEmail & vbNullString)
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            LAcName = ListView1.ListItems(I).SubItems(1)
            LAcCode = ListView1.ListItems(I).text
            LACCID = ListView1.ListItems(I).SubItems(4)
            
            'Set TRec = Nothing
            'Set TRec = New ADODB.Recordset
            'MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & LAcCode & "'"
            'TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            'If Not TRec.EOF Then
            '    LAcName = TRec!NAME
            'End If
            
            LBal = Get_ClosingBal(LACCID, vcDTP2.Value)
            If LBal <> 0 Then
                If LFileFlag = False And ChkVouFile.Value = 1 Then
                    TxtPath = App.Path & "\VOUCHER\VOU" & CStr(Left$(vcDTP1.Value, 2)) & CStr(Mid(vcDTP1.Value, 4, 2)) & Right$(CStr(Year(vcDTP1.Value)), 4) & ".CSV"
                    Set LFileSystemObject = CreateObject("Scripting.FileSystemObject")
                    Set LFile = LFileSystemObject.CreateTextFile(TxtPath, True)
                    LFileFlag = True
                End If
                If LBal > 0 Then
                    LPtyDrCr = "D"
                    LCDrCr = "C"
                ElseIf LBal < 0 Then
                    LPtyDrCr = "C"
                    LCDrCr = "D"
                End If
                If Check2.Value = 1 And LChk_Vou = False Then
                    LChk_Vou = True
                    LMaxVouNo = LMaxVouNo + 1
                    LVchNo = Get_Next_Vou_No(LMaxVouNo, "JNRL", Right$(GFinBegin, 2) & Right$(GFinEnd, 2))
                    LMVNo = PInsert_Voucher(LVchNo, vcDTP1.Value, "JV", "P", 0, "ADD", vbNullString, 0, "Transfer Bal", vbNullString, "0", 0, 0)
                End If
                
                LNarr = "Bal Trf to " & LToAccName & ""
                Call PInsert_Vchamt(LVchNo, "JV", vcDTP1.Value, LPtyDrCr, LAcCode, Abs(LBal), vbNullString, vcDTP1.Value, LNarr, vbNullString, vbNullString, 0, vbNullString, LMVNo, 0, 0, LACCID)
                'If ChkVouFile.Value = 1 Then
                '    TradeLine = vcDTP1.Value & "," & Left(vcDTP1.Value, 2) & Left(MonthName(month(vcDTP1.Value)), 3) & Year(vcDTP1.Value) & "-" & CStr(LMaxVouNo) & ",JV," & LAcCode & "," & LAcName & "," & LPtyDrCr & "," & Abs(LBal) & "," & LNarr
                '    LFile.WriteLine (TradeLine)
                'End If
                LNarr = "Bal Trf from " & LAcName & ""
                Call PInsert_Vchamt(LVchNo, "JV", vcDTP1.Value, LCDrCr, LToAcc, Abs(LBal), vbNullString, vcDTP1.Value, LNarr, vbNullString, vbNullString, 0, vbNullString, LMVNo, 0, 0, LToAccID)
                If ChkVouFile.Value = 1 Then
                    LNarr = "Bal Trf from " & LAcName & " TO " & LToAccName
                    LCount = LCount + 1
                    If LPtyDrCr = "C" Then
                        TradeLine = vcDTP1.Value & "," & LAcCode & "," & LToAcc & "," & Abs(LBal) & "," & LNarr & "," & LCount & "," & LAcName & "," & LToAccName
                    Else
                        TradeLine = vcDTP1.Value & "," & LToAcc & "," & LAcCode & "," & Abs(LBal) & "," & LNarr & "," & LCount & "," & LAcName & "," & LToAccName
                    End If
                    LFile.WriteLine (TradeLine)
                End If
            End If
        End If
    Next
    If LFileFlag = True Then LFile.close
    MsgBox " Transfer Complete"
    Call Fill_List
End If
End Sub
Private Sub Command2_Click()
Call Fill_List
End Sub
Private Sub Form_Load()
vcDTP1.Value = Date
vcDTP2.Value = Date
Check2.Value = 1
Set AccRec = Nothing
Set AccRec = New ADODB.Recordset
mysql = "SELECT ACCID,AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " ORDER BY NAME  "
AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
AccRec.MoveFirst
Set DataCombo1.RowSource = AccRec
    DataCombo1.BoundColumn = "AC_CODE"
    DataCombo1.ListField = "NAME"
Set AccFmlyRec = Nothing
Set AccFmlyRec = New ADODB.Recordset
mysql = "SELECT FMLYID,FMLYCODE ,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " ORDER BY FMLYNAME  "
AccFmlyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not AccFmlyRec.EOF Then
 Set BranchCombo.RowSource = AccFmlyRec
    BranchCombo.ListField = "FMLYNAME"
    BranchCombo.BoundColumn = "FMLYCODE"
End If

End Sub
Private Sub Fill_List()
Dim LBal As Double
ListView1.ListItems.Clear
Set AccRec = Nothing
Set AccRec = New ADODB.Recordset
mysql = "SELECT ACCID,AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " "
If LenB(BranchCombo.BoundText) > 0 Then mysql = mysql & " AND ACCID IN (SELECT ACCID FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND FMLYCODE ='" & BranchCombo.BoundText & "')"
mysql = mysql & " ORDER BY NAME  "
AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If AccRec.EOF Then
    MsgBox "No Accounts for selected Values"
    Exit Sub
End If

AccRec.MoveFirst
Do While Not AccRec.EOF
    
    LBal = Round(Get_ClosingBal(AccRec!ACCID, vcDTP2.Value), 2)
    If LBal <> 0 Then
        ListView1.ListItems.Add , , AccRec!AC_CODE
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , AccRec!NAME
        If LBal > 0 Then
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ""
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(LBal, "0.00")
        ElseIf LBal < 0 Then
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(Abs(LBal), "0.00")
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ""
        End If
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Str(AccRec!ACCID)
    End If
    AccRec.MoveNext
Loop
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub vcDTP1_Validate(Cancel As Boolean)
 ListView1.ListItems.Clear
End Sub
Private Sub vcDTP2_Validate(Cancel As Boolean)
 ListView1.ListItems.Clear
End Sub
