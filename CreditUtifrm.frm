VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form CreditUtifrm 
   Caption         =   "Credit Utilisaction"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   9750
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   2760
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   9135
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   8415
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   1920
         TabIndex        =   15
         Top             =   8040
         Width           =   3375
         Begin VB.OptionButton OptWeekM 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Week Mtm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   17
            Top             =   65
            Width           =   1695
         End
         Begin VB.OptionButton OptLedg 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Ledger"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   65
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   1920
         TabIndex        =   8
         Top             =   7545
         Width           =   3375
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00FFC0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   65
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptCCL 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cross Credit Limit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Top             =   65
            Width           =   2175
         End
      End
      Begin VB.CheckBox PartyChk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select All"
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
         Left            =   6840
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton CancelCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7560
         Width           =   975
      End
      Begin VB.CommandButton OkCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7560
         Width           =   975
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   120
         Width           =   1500
         _ExtentX        =   2646
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
         Value           =   42079.7973611111
      End
      Begin MSComctlLib.ListView PartyList 
         Height          =   6615
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select members, F5 to select non members."
         Top             =   840
         Visible         =   0   'False
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11668
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parties"
            Object.Width           =   7673
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "PartyType"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Op_Bal"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OpBal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "FmlyId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SRVTAXAPP"
            Object.Width           =   0
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
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
         Value           =   42079.7973611111
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Type"
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
         Left            =   360
         TabIndex        =   14
         Top             =   8160
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Generate"
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
         Left            =   360
         TabIndex        =   11
         Top             =   7725
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Parties"
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
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "As on date"
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
         Left            =   480
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Credit Utilisation Report"
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
      TabIndex        =   12
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "CreditUtifrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AccRec As ADODB.Recordset
Dim RecRpt As ADODB.Recordset
Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Date > GFinEnd Then
        vcDTP2.Value = GFinEnd:
    Else
        vcDTP2.Value = Date
    End If
    
    Fill_Parties
End Sub
Private Sub Form_Unload(Cancel As Integer)
 If CRViewer1.Visible = True Then
        CRViewer1.Visible = False
        Cancel = 1
    Else
        Unload Me
    End If
End Sub

Private Sub OkCmd_Click()
    On Error GoTo err1
    
    Dim Ltype As String
    Dim LParties As String
    Dim Weekdy As Integer
    Dim VStkVal As Double
    
    LParties = Get_Parties
    
    Ltype = "1"
    If OptAll.Value Then
        Ltype = "0" 'All
    End If
   
    VStkVal = 0
     
    Dim MRec As ADODB.Recordset
    Dim MRecSV As ADODB.Recordset
        
    If OptWeekM.Value Then  ' report type ledger or weekly
        Weekdy = Weekday(Date)
        
        If Weekdy = 1 Then 'Sunday
            vcDTP1.Value = Date - 6
            vcDTP2.Value = vcDTP2.Value - 1
        ElseIf Weekdy = 2 Then 'monday
            vcDTP1.Value = Date
        ElseIf Weekdy = 3 Then
            vcDTP1.Value = Date - 1
        ElseIf Weekdy = 4 Then
            vcDTP1.Value = Date - 2
        ElseIf Weekdy = 5 Then
            vcDTP1.Value = Date - 3
        ElseIf Weekdy = 6 Then
            vcDTP1.Value = Date - 4
        ElseIf Weekdy = 7 Then
            vcDTP1.Value = Date - 5
        End If
        
        
        mysql = " EXEC Credit_Utilisaction '" & LParties & "','" & Format(vcDTP2.Value, "YYYY/MM/DD") & "','" & Ltype & "','" & Format(vcDTP1.Value, "YYYY/MM/DD") & "','W','" & GCompCode & "'"
    Else
        mysql = " EXEC Credit_Utilisaction '" & LParties & "','" & Format(vcDTP2.Value, "YYYY/MM/DD") & "','" & Ltype & "','','','" & GCompCode & "'"
    End If
    
    Set MRec = Nothing
    Set MRec = New ADODB.Recordset
    MRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not MRec.EOF Then
                        
            Me.MousePointer = Val(11): OkCmd.Enabled = False
            
            Set RecRpt = Nothing
            Set RecRpt = New ADODB.Recordset
            RecRpt.Fields.Append "NAME", adVarChar, 200, adFldIsNullable
            RecRpt.Fields.Append "LEDGERBALANCE", adDouble, , adFldIsNullable
            RecRpt.Fields.Append "LBDC", adVarChar, 15, adFldIsNullable
            RecRpt.Fields.Append "CrLimit", adDouble, , adFldIsNullable
            RecRpt.Fields.Append "CrLDC", adVarChar, 15, adFldIsNullable
            RecRpt.Fields.Append "Balance", adDouble, , adFldIsNullable
            RecRpt.Fields.Append "StockVal", adDouble, , adFldIsNullable
            RecRpt.Open , , adOpenKeyset, adLockBatchOptimistic
                                                                   
            Do While Not MRec.EOF
            
                'Check Stock Value with standing report
                VStkVal = 0
                mysql = "select party,contype,sum(qty) as 'stkqty' From ctr_d where compcode = '" & GCompCode & "' and condate <= '" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' and party = '" & MRec!AC_CODE & "' group by party,contype order by contype"
                Set MRecSV = Nothing
                Set MRecSV = New ADODB.Recordset
                MRecSV.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                While Not MRecSV.EOF
                    If MRecSV!CONTYPE = "B" Then
                        VStkVal = MRecSV!stkqty
                    Else
                        VStkVal = VStkVal - MRecSV!stkqty
                    End If
                    MRecSV.MoveNext
                Wend

                RecRpt.AddNew
                RecRpt!NAME = MRec!NAME
                RecRpt!LEDGERBALANCE = MRec!LEDGERBALANCE
                RecRpt!LBDC = MRec!LBDC
                RecRpt!CrLimit = MRec!CrLimit
                RecRpt!CrLDC = MRec!CrLDC
                RecRpt!Balance = MRec!Balance
                RecRpt!StockVal = VStkVal
                RecRpt.Update
                MRec.MoveNext
            Loop
            Set MRec = Nothing
        
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        Set MRec = RecRpt.Clone
        
        If OptWeekM.Value Then  ' balance type  weekly
            mysql = "'Credit Utilisation (weekly) as on " & vcDTP2.Value & "'"
        Else
            mysql = "'Credit Utilisation (Ledger) as on " & vcDTP2.Value & "'"
        End If
                
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "CrUtilisaction.rpt", 1)
                
        RDCREPO.DiscardSavedData
        RDCREPO.Database.SetDataSource MRec
        RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
        RDCREPO.FormulaFields.GetItemByName("TITLE").text = mysql
        CRViewer1.ZOrder
        CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1100)
        CRViewer1.Visible = True
        CRViewer1.ReportSource = RDCREPO
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Me.MousePointer = 0
        OkCmd.Enabled = True
        Exit Sub
    Else
        MsgBox "No Records Found", vbInformation
    End If
    
err1:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    End If
    Me.MousePointer = 0
    OkCmd.Enabled = True
    
End Sub

Private Sub Partychk_Click()
    Dim I As Integer
    For I = 1 To PartyList.ListItems.Count
        If PartyChk.Value = 1 Then
            PartyList.ListItems.Item(I).Checked = True
        Else
            PartyList.ListItems.Item(I).Checked = False
        End If
    Next I

End Sub
Sub Fill_Parties()
    PartyList.Visible = True
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC.ACCID,AC.AC_CODE,AC.NAME,AC.OP_BAL,AG.G_NAME,AG.CODE,AG.TYPE FROM ACCOUNTM AS AC, ACCOUNTD AS AD, AC_GROUP AS AG "
    mysql = mysql & " WHERE AC.COMPCODE=" & GCompCode & "  and ac.ac_code = ad.ac_code and AC.GCODE=AG.CODE AND AC.GRPCODE=AG.G_CODE AND (AC.GCODE=12 or ac.gcode = 14) AND ISNULL(AD.CRLIMIT,0)<>0 ORDER BY AC.NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        PartyList.Checkboxes = True
        PartyList.ListItems.Clear
        Do While Not AccRec.EOF
            PartyList.ListItems.Add , , UCase(AccRec!NAME)
            PartyList.ListItems(PartyList.ListItems.Count).ListSubItems.Add , , AccRec!AC_CODE
            PartyList.ListItems(PartyList.ListItems.Count).ListSubItems.Add , , AccRec!ACCID
            
            AccRec.MoveNext
        Loop
        AccRec.MoveFirst
    End If
End Sub
Public Function Get_Parties() As String
    Dim LFParty_Codes As String
    Dim I As Integer
    LFParty_Codes = vbNullString
    For I = 1 To PartyList.ListItems.Count
        If PartyList.ListItems(I).Checked = True Then
            'If LenB(LFParty_Codes) <> 0 Then LFParty_Codes = LFParty_Codes & "^"
            LFParty_Codes = LFParty_Codes & PartyList.ListItems(I).SubItems(1) & "^"
        End If
    Next
    Get_Parties = LFParty_Codes
End Function
