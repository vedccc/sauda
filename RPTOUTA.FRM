VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form RPTOUTA 
   Caption         =   "Standing Report"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sauda wise Standing"
      TabPicture(0)   =   "RPTOUTA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "vcDTP1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   3720
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "RPTOUTA.frx":001C
         Left            =   5760
         List            =   "RPTOUTA.frx":0029
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   660
         Width           =   1935
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37866.8704282407
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   1920
         TabIndex        =   5
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sauda"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   6
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4440
         TabIndex        =   3
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Up to Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   1125
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   0
      TabIndex        =   9
      Top             =   0
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
Attribute VB_Name = "RPTOUTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC As New ADODB.Recordset
Private Sub Command1_Click()
    On Error GoTo ERR1
    Dim RECRPT As ADODB.Recordset

    Dim SAUDAS As String

    Me.MousePointer = Val(11)
    SAUDAS = ""

    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            SAUDAS = SAUDAS & "'"
            SAUDAS = SAUDAS & ListView1.ListItems(I).ListSubItems(1)
            SAUDAS = SAUDAS & "'"
        End If

        If I < ListView1.ListItems.Count Then
            If ListView1.ListItems(I + 1).Checked = True And Len(SAUDAS) > Val(0) Then
                SAUDAS = SAUDAS & ", "
            End If
        End If
    
    Next

    If Len(Trim(SAUDAS)) < 1 Then
        MsgBox "Select sauda to get the report.", vbInformation, "Message"
        Me.MousePointer = Val(0)
        Exit Sub
    End If

    If Combo1.ListIndex = Val(0) Then
        PARTYTYPE = "'M'"

    ElseIf Combo1.ListIndex = Val(0) Then
        PARTYTYPE = "'N'"

    Else
        PARTYTYPE = "'M', 'N', 'B'"
    End If

    MYSQL = "SELECT Sum(CD.QTY) AS SumOfQTY, CD.CONTYPE, SU.SAUDANAME, P.NAME FROM CTR_D AS CD, SAUDAMAST AS SU, ACCOUNTD AS P, PITBROK AS BR WHERE CD.SAUDA=SU.SAUDACODE AND CD.PARTY=P.AC_CODE AND SU.ITEMCODE=BR.ITEMCODE AND BR.AC_CODE=CD.PARTY AND CD.CONDATE <=DATEVALUE('" & Format(vcDTP1.Value, "dd/MM/yyyy") & "') AND BR.PARTYTYPE IN (" & PARTYTYPE & ") AND CD.SAUDA IN (" & SAUDAS & ") GROUP BY CD.CONTYPE, SU.SAUDANAME, P.NAME ORDER BY SU.SAUDANAME, P.NAME"
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open MYSQL, CNN, adOpenKeyset, adLockReadOnly

    Set RECRPT = Nothing
    Set RECRPT = New ADODB.Recordset
    RECRPT.Fields.Append "SAUDA", adVarChar, 35, adFldIsNullable
    RECRPT.Fields.Append "SQNTY", adDouble, , adFldIsNullable
    RECRPT.Fields.Append "SPARTY", adVarChar, 35, adFldIsNullable
    RECRPT.Fields.Append "BQNTY", adDouble, , adFldIsNullable
    RECRPT.Fields.Append "BPARTY", adVarChar, 35, adFldIsNullable
    RECRPT.Open , , adOpenKeyset, adLockBatchOptimistic

    Do While Not REC.EOF
        MSAUDA = REC!SAUDANAME
        Do While MSAUDA = REC!SAUDANAME
            MPARTY = REC!Name
            MQNTY = 0
            Do While MPARTY = REC!Name
                If REC!CONTYPE = "B" Then
                    MQNTY = MQNTY + REC!SUMOFQTY
                Else
                    MQNTY = MQNTY - REC!SUMOFQTY
                End If
                
                REC.MoveNext
                If REC.EOF Then GoTo FLAG1
            Loop
'------
FLAG1:
            If Not RECRPT.EOF Then RECRPT.MoveFirst

            If MQNTY > 0 Then
                RECRPT.Filter = "SAUDA='" & MSAUDA & "'"
                RECRPT.Find "BPARTY=''", , adSearchForward
                If RECRPT.EOF Then
                    RECRPT.AddNew
                    MADDNEW = "1"
                Else
                    MADDNEW = ""
                End If
                If MADDNEW = "1" Then
                    RECRPT!SQNTY = Val(0)
                    RECRPT!SPARTY = ""
                End If
                RECRPT!SAUDA = MSAUDA
                RECRPT!BQNTY = Abs(Val(MQNTY))
                RECRPT!BPARTY = MPARTY
                RECRPT.Update
    
            ElseIf MQNTY < 0 Then
                RECRPT.Filter = "SAUDA='" & MSAUDA & "'"
                RECRPT.Find "SPARTY=''", , adSearchForward
                If RECRPT.EOF Then
                    RECRPT.AddNew
                    MADDNEW = "1"
                Else
                    MADDNEW = ""
                End If
                If MADDNEW = "1" Then
                    RECRPT!BQNTY = Val(0)
                    RECRPT!BPARTY = ""
                End If
                RECRPT!SAUDA = MSAUDA
                RECRPT!SQNTY = Abs(Val(MQNTY))
                RECRPT!SPARTY = MPARTY
                RECRPT.Update
            End If
'-----
            If REC.EOF Then GoTo FLAG2
        Loop
FLAG2:
    Loop

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    Set REC = RECRPT.Clone

    Set RDCREPO = RDCAPP.OpenReport(RPT_PATH & "MOUT.rpt", 1)

    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource REC

    RDCREPO.FormulaFields.GetItemByName("COMPANY").Text = "'" & MC_NAME & "'"
    MYSQL = "'Sauda Wise Standing upto " & Format(vcDTP1.Value, "dd/MM/yyyy") & "'"
    RDCREPO.FormulaFields.GetItemByName("TITLE").Text = MYSQL

    CRViewer1.ZOrder

    CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)

    CRViewer1.Visible = True
    CRViewer1.ReportSource = RDCREPO
    CRViewer1.Zoom 1
    CRViewer1.ViewReport
    
    Me.MousePointer = 0
    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
    Me.MousePointer = Val(0)
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Call ClearFormFn(RPTOUTA)

    Combo1.ListIndex = Val(2)

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open "SELECT * FROM SAUDAMAST ORDER BY SAUDANAME", CNN, adOpenForwardOnly, adLockReadOnly
    Do While Not REC.EOF
        ListView1.ListItems.ADD , , REC!SAUDANAME
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , REC!SAUDACODE
        REC.MoveNext
    Loop
End Sub

Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    SSTab1.BackColor = GETMAIN.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        CRViewer1.Visible = False
        Cancel = 1
    End If
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    For I = 1 To ListView1.ListItems.Count
        If KeyCode = 113 Then
            ListView1.ListItems(I).Checked = True
        ElseIf KeyCode = 114 Then
            ListView1.ListItems(I).Checked = False
        End If
    Next
End Sub
