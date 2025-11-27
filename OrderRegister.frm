VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRVIEWER.DLL"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form OrderRegister 
   Caption         =   "Order Register"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   11370
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   9135
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   8415
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
         TabIndex        =   4
         Top             =   8520
         Width           =   975
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
         TabIndex        =   5
         Top             =   8520
         Width           =   975
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
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   3360
         TabIndex        =   1
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
         Height          =   7215
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select members, F5 to select non members."
         Top             =   1200
         Visible         =   0   'False
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   12726
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
         TabIndex        =   0
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
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
         TabIndex        =   10
         Top             =   180
         Width           =   1215
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
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   1200
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
   Begin VB.Label Label27 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Order Register"
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
      TabIndex        =   6
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "OrderRegister"
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

    vcDTP1.Value = Date
    If Date > GFinEnd Then
        vcDTP2.Value = GFinEnd
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
    
    LParties = Get_Parties
         
    Dim MRec As ADODB.Recordset
        
    MYSQL = " EXEC ORDER_REGISTER '" & GCompCode & "','" & LParties & "','" & Format(vcDTP1.Value, "YYYY/MM/DD") & "','" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
           
    Set MRec = Nothing
    Set MRec = New ADODB.Recordset
    MRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not MRec.EOF Then
                        
            Me.MousePointer = Val(11): OkCmd.Enabled = False
            
            Set RecRpt = Nothing
            Set RecRpt = New ADODB.Recordset
            RecRpt.Fields.Append "ORDDT", adDate, , adFldIsNullable
            RecRpt.Fields.Append "ORDNO", adDouble, , adFldIsNullable
            RecRpt.Fields.Append "SAUDA", adVarChar, 200, adFldIsNullable
            RecRpt.Fields.Append "BS", adVarChar, 6, adFldIsNullable
            RecRpt.Fields.Append "QTY", adDouble, , adFldIsNullable
            RecRpt.Fields.Append "RATE", adDouble, , adFldIsNullable
            RecRpt.Fields.Append "STATUS", adVarChar, 20, adFldIsNullable
            
            RecRpt.Open , , adOpenKeyset, adLockBatchOptimistic
                                                                   
            Do While Not MRec.EOF
                RecRpt.AddNew
                    RecRpt!ORDDT = MRec!ORDDT
                    RecRpt!ORDNO = MRec!ORDNO
                    RecRpt!Sauda = MRec!Sauda
                    RecRpt!BS = MRec!BS
                    RecRpt!QTY = MRec!QTY
                    RecRpt!Rate = MRec!Rate
                    RecRpt!Status = MRec!Status
                RecRpt.Update
                MRec.MoveNext
            Loop
            Set MRec = Nothing
        
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        Set MRec = RecRpt.Clone
              
        MYSQL = "'Order Register " & vcDTP1.Value & " to " & vcDTP2.Value & "'"
      
                
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "OrderRegister.rpt", 1)
                
        RDCREPO.DiscardSavedData
        RDCREPO.Database.SetDataSource MRec
        RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
        RDCREPO.FormulaFields.GetItemByName("TITLE").text = MYSQL
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
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    
   ' MYSQL = " EXEC Get_PartyCtr_d " & GCompCode & ",'" & Format(DtpCondate.Value, "YYYY/MM/DD") & "'," & LFExID & ""
    
    'MYSQL = "SELECT AC.ACCID,AC.AC_CODE,AC.NAME,AC.OP_BAL,AG.G_NAME,AG.CODE,AG.TYPE FROM ACCOUNTM AS AC, ACCOUNTD AS AD, AC_GROUP AS AG "
    'MYSQL = MYSQL & " WHERE AC.COMPCODE=" & GCompCode & "  and ac.ac_code = ad.ac_code and AC.GCODE=AG.CODE AND AC.GRPCODE=AG.G_CODE AND AC.GCODE=12 AND ISNULL(AD.CRLIMIT,0)<>0 ORDER BY AC.NAME"
    
    MYSQL = "SELECT AC.ACCID,AC.AC_CODE,AC.NAME,AC.OP_BAL FROM ACCOUNTM AS AC, ACCOUNTD AS AD "
    MYSQL = MYSQL & " WHERE AC.COMPCODE=" & GCompCode & "  and ac.compcode = ad.compcode and ac.ac_code = ad.ac_code AND AC.AC_CODE in (select PARTY from ord_m) ORDER BY AC.NAME"
    
    AccRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        PartyList.Visible = True
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

