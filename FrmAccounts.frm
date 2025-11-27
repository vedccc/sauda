VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMACCOUNT 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   7335
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   18495
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   2775
         Left            =   7440
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   3375
         Begin MSComctlLib.ListView ExListView 
            Height          =   2460
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   4339
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388736
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CODE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   18255
         _ExtentX        =   32200
         _ExtentY        =   12515
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   18495
      Begin VB.CommandButton CmdEx 
         Caption         =   "Select Exchnage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12240
         TabIndex        =   18
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox TxtBranchCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CmdShw 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   9120
         TabIndex        =   13
         Top             =   120
         Width           =   3015
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Brokerage"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   4
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo BranchCombo 
         Height          =   390
         Left            =   2880
         TabIndex        =   6
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   688
         _Version        =   393216
         Text            =   ""
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
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   44095.417037037
      End
      Begin VB.Label Label5 
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
         Left            =   12120
         TabIndex        =   17
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   803
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Starting From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   12
         Top             =   60
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Accounts From Trade Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   2655
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Query on Account"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   18495
   End
End
Attribute VB_Name = "FRMACCOUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LAccRec As ADODB.Recordset
Dim RECGRID As ADODB.Recordset
Dim LBranchRec As ADODB.Recordset
Dim LExIDS As String
Dim ExRec As ADODB.Recordset

Private Sub BranchCombo_Validate(Cancel As Boolean)
If LenB(BranchCombo.BoundText) > 0 Then
    If LBranchRec.RecordCount > 0 Then LBranchRec.MoveFirst
    LBranchRec.Find "FMLYCODE='" & BranchCombo.BoundText & "'"
    If Not LBranchRec.EOF Then
        TxtBranchCode.text = LBranchRec!FMLYCODE
    Else
        TxtBranchCode.text = vbNullString
        BranchCombo.BoundText = vbNullString
    End If
End If
End Sub

Private Sub CmdEx_Click()
If Frame5.Visible = True Then
    Frame5.Visible = False
Else
    Frame5.Visible = True
End If
End Sub

Private Sub CmdSave_Click()
Dim TRec As ADODB.Recordset:    Dim LOp_Bal As Double:      Dim LAC_CODE As String
Dim LExCode As String:          Dim LBrokType As String:    Dim LBrokRate As Double
Dim LUptostDt As Date:          Dim LInstType As String:    Dim LExID As Integer
Dim LName As String:            Dim LSParties As String:    Dim LSExIDS As String
Dim LACCID As Long

LSParties = vbNullString
LSExIDS = vbNullString


Label5.Caption = "Updating Data "
DoEvents
On Error GoTo err1
If Option1.Value = True Then
    If RECGRID.RecordCount > 0 Then
        Cnn.BeginTrans
        RECGRID.MoveFirst
        Do While Not RECGRID.EOF
            LACCID = RECGRID!ACCID
            LAC_CODE = RECGRID!AC_CODE
            LName = RECGRID!NAME
            LOp_Bal = Val(RECGRID!OP_BAL & vbNullString)
            Set TRec = Nothing
            Set TRec = New ADODB.Recordset
            mysql = "SELECT AC_CODE FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE <>'" & LAC_CODE & "'AND NAME ='" & LName & "'"
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                MsgBox " Name already Exist for Account Code " & TRec!AC_CODE & ""
            Else
                mysql = "UPDATE ACCOUNTM SET NAME ='" & UCase(LName) & " ' WHERE ACCID =" & LACCID & " "
                Cnn.Execute mysql
                mysql = "UPDATE ACCOUNTD SET NAME ='" & UCase(LName) & " ' WHERE  ACCID =" & LACCID & " "
                Cnn.Execute mysql
            End If
            mysql = "UPDATE ACCOUNTM SET OP_BAL  =" & LOp_Bal & "  WHERE  ACCID  =" & LACCID & " "
            Cnn.Execute mysql
            RECGRID.MoveNext
        Loop
        Cnn.CommitTrans
        Label5.Caption = "Data Update Complete"
        DoEvents
        Set DataGrid1.DataSource = Nothing
    End If
Else
    If RECGRID.RecordCount > 0 Then
        Cnn.BeginTrans
        RECGRID.MoveFirst
        Do While Not RECGRID.EOF
            LAC_CODE = RECGRID!AC_CODE
            LExCode = RECGRID!EXCODE
            LExID = RECGRID!EXID
            LInstType = RECGRID!INSTTYPE
            LBrokType = RECGRID!broktype
            LBrokRate = RECGRID!brokrate
            LUptostDt = RECGRID!UPTOSTDT
            LACCID = RECGRID!ACCID
            If InStr(LSParties, LAC_CODE) = 0 Then
                If LenB(LSParties) > 0 Then LSParties = LSParties & ","
                LSParties = LSParties & "'" & LAC_CODE & "'"
            End If
            If InStr(LSExIDS, Str(LExID)) = 0 Then
                If LenB(LSExIDS) > 0 Then LSExIDS = LSExIDS & ","
                LSExIDS = LSExIDS & Str(LExID)
            End If
            
            mysql = "DELETE FROM PEXBROK WHERE EXID=" & LExID & " AND ACCID =" & LACCID & " AND "
            mysql = mysql & " UPTOSTDT ='" & Format(LUptostDt, "YYYY/MM/DD") & "' AND INSTTYPE ='" & LInstType & "'"
            Cnn.Execute mysql
            Call PInsert_PExBrok(LAC_CODE, LExCode, LBrokType, LBrokRate, 0, 0, 0, "P", 0, 0, "P", 0, 0, "V", 0, LUptostDt, LInstType, 0, 0, LExID, LACCID)
            RECGRID.MoveNext
        Loop
        Cnn.CommitTrans
        Label5.Caption = "Data Update Complete"
        
        Cnn.BeginTrans: CNNERR = True
        Call Update_Charges(LSParties, LSExIDS, vbNullString, vbNullString, GFinBegin, GFinEnd, False)
        Cnn.CommitTrans: CNNERR = False
        Cnn.BeginTrans: CNNERR = True
        If BILL_GENERATION(GFinBegin, GFinEnd, vbNullString, LSParties, LSExIDS) Then
            Cnn.CommitTrans: CNNERR = False
        Else
            Cnn.RollbackTrans: CNNERR = False
        End If
        'Call Chk_Billing
        MsgBox "Brokerage Succesfully Updated "
        GETMAIN.ProgressBar1.Visible = False
        
        DoEvents
        Set DataGrid1.DataSource = Nothing
    End If
End If
Exit Sub
err1:
 MsgBox err.Description
 If CNNERR = True Then
    
    Cnn.RollbackTrans
 End If
End Sub



Private Sub RecSet()
Set RECGRID = Nothing
Set RECGRID = New ADODB.Recordset
RECGRID.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
RECGRID.Fields.Append "NAME", adVarChar, 100, adFldIsNullable
RECGRID.Fields.Append "OP_BAL", adDouble, , adFldIsNullable
RECGRID.Fields.Append "ACCID", adInteger, , adFldIsNullable
RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub CmdShw_Click()

If Option1.Value = True Then
    Call Update_Name
Else
    Call Update_BROK
End If
End Sub

    
'Private Sub ()

Private Sub DataGrid1_Click()

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim LGridRow As Integer
Dim LGridCol As Integer
Dim GridColVal  As String
If KeyCode = 118 Then   'F7
    LGridRow = DataGrid1.Row
    LGridCol = DataGrid1.Col
    If DataGrid1.Col = 4 Then 'BROKTYPE
        GridColVal = RECGRID!broktype
        RECGRID.MoveFirst
        While Not RECGRID.EOF
            RECGRID!broktype = UCase(GridColVal)
            RECGRID.MoveNext
        Wend
    ElseIf DataGrid1.Col = 5 Then 'BROKRATE
        GridColVal = RECGRID!brokrate
        RECGRID.MoveFirst
        While Not RECGRID.EOF
            RECGRID!brokrate = Val(GridColVal)
            RECGRID.MoveNext
        Wend
    End If
End If
End Sub

'End Sub

Private Sub Form_Load()
Option1.Value = True
vcDTP1.Value = Date
Set LBranchRec = Nothing
Set LBranchRec = New ADODB.Recordset
mysql = "SELECT FMLYID,FMLYCODE,FMLYNAME FROM ACCFMLY ORDER BY FMLYNAME "
LBranchRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
If Not LBranchRec.EOF Then
    Set BranchCombo.RowSource = LBranchRec
    BranchCombo.ListField = "FMLYNAME"
    BranchCombo.BoundColumn = "FMLYCODE"
End If

Set ExRec = Nothing
Set ExRec = New ADODB.Recordset
mysql = "SELECT EXCODE,EXID FROM EXMAST where COMPCODE=" & GCompCode & "  ORDER BY EXCODE "
ExRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
ExListView.ListItems.Clear
If Not ExRec.EOF Then
    If ExRec.RecordCount > 1 Then
        ExListView.Visible = False
        Do While Not ExRec.EOF
            ExListView.ListItems.Add , , ExRec!EXCODE
            ExListView.ListItems(ExListView.ListItems.Count).ListSubItems.Add , , ExRec!EXID
            ExRec.MoveNext
        Loop
        ExListView.Visible = True
    Else
        ExListView.Enabled = False:
    End If
End If
End Sub

Private Sub TxtBranchCode_Validate(Cancel As Boolean)
    If LBranchRec.RecordCount > 0 Then LBranchRec.MoveFirst
    LBranchRec.Find "FMLYCODE='" & TxtBranchCode.text & "'"
    If Not LBranchRec.EOF Then
        BranchCombo.BoundText = TxtBranchCode.text
    Else
        TxtBranchCode.text = vbNullString
        BranchCombo.BoundText = vbNullString
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub


Private Sub RecSet2()
Set RECGRID = Nothing
Set RECGRID = New ADODB.Recordset
RECGRID.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
RECGRID.Fields.Append "NAME", adVarChar, 100, adFldIsNullable
RECGRID.Fields.Append "OP_BAL", adDouble, , adFldIsNullable
RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub Update_Name()
Call RecSet
Set LAccRec = Nothing
Set LAccRec = New ADODB.Recordset
mysql = "SELECT A.ACCID,A.AC_CODE,A.NAME,A.OP_BAL FROM ACCOUNTM AS A "
mysql = mysql & "WHERE A.COMPCODE =" & GCompCode & "  "
mysql = mysql & " AND A.ACCID NOT IN (SELECT DISTINCT ACCID FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE <'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
mysql = mysql & " AND A.ACCID IN (SELECT DISTINCT ACCID FROM CTR_D WHERE CONDATE >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
If LenB(TxtName.text) > 0 Then mysql = mysql & " AND A.ACCID IN (SELECT ACCID FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND NAME LIKE '" & TxtName.text & "%' )"
If LenB(TxtBranchCode.text) > 0 Then mysql = mysql & " AND A.ACCID IN (SELECT ACCID FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND FMLYCODE='" & TxtBranchCode.text & "')"
mysql = mysql & " ORDER BY A.NAME "
LAccRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
Do While Not LAccRec.EOF
    RECGRID.AddNew
    RECGRID!AC_CODE = LAccRec!AC_CODE
    RECGRID!NAME = LAccRec!NAME
    RECGRID!OP_BAL = LAccRec!OP_BAL
    RECGRID!ACCID = LAccRec!ACCID
    RECGRID.Update
    LAccRec.MoveNext
Loop
Set DataGrid1.DataSource = RECGRID
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 5000
DataGrid1.Columns(2).Width = 2000
DataGrid1.Columns(2).Alignment = dbgRight
DataGrid1.Columns(2).NumberFormat = "0.00"

DataGrid1.ReBind
DataGrid1.Refresh
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 5000
DataGrid1.Columns(2).Width = 2000
DataGrid1.Columns(2).Alignment = dbgRight
DataGrid1.Columns(2).NumberFormat = "0.00"


End Sub

Private Sub BrokRecSet()
Set RECGRID = Nothing
Set RECGRID = New ADODB.Recordset
RECGRID.Fields.Append "Ac_Code", adVarChar, 15, adFldIsNullable
RECGRID.Fields.Append "Name", adVarChar, 100, adFldIsNullable
RECGRID.Fields.Append "ExCode", adVarChar, 10, adFldIsNullable
RECGRID.Fields.Append "EXID", adInteger, 10, adFldIsNullable
RECGRID.Fields.Append "BrokTYpe", adVarChar, 1, adFldIsNullable
RECGRID.Fields.Append "BrokRate", adDouble, 1, adFldIsNullable
RECGRID.Fields.Append "UPTOSTDT", adDate, , adFldIsNullable
RECGRID.Fields.Append "INSTTYPE", adVarChar, 3, adFldIsNullable
RECGRID.Fields.Append "ACCID", adInteger, , adFldIsNullable

RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub


Private Sub Update_BROK()
Call BrokRecSet
Set LAccRec = Nothing
Set LAccRec = New ADODB.Recordset
Call Get_ExIDs
Label5.Caption = "Getting Data"
DoEvents
mysql = "SELECT A.ACCID,A.AC_CODE,A.NAME,B.EXCODE,B.EXID,B.BROKTYPE,B.BROKRATE,B.UPTOSTDT,B.INSTTYPE FROM ACCOUNTD AS A ,PEXBROK AS B "
mysql = mysql & " WHERE A.COMPCODE =" & GCompCode & "  AND B.ACCID  = A.ACCID AND B.UPTOSTDT ='" & Format(GFinEnd, "YYYY/MM/DD") & "' AND B.INSTTYPE ='FUT' "
If LenB(LExIDS) > 0 Then mysql = mysql & " AND B.EXID IN (" & LExIDS & ")"
mysql = mysql & " AND A.ACCID NOT IN (SELECT DISTINCT ACCID FROM CTR_D WHERE CONDATE <'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
mysql = mysql & " AND A.ACCID  IN (SELECT DISTINCT ACCID FROM CTR_D WHERE CONDATE >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
If LenB(TxtName.text) > 0 Then mysql = mysql & " AND A.ACCID IN (SELECT ACCID FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND NAME LIKE '" & TxtName.text & "%' )"
If LenB(TxtBranchCode.text) > 0 Then mysql = mysql & " AND A.ACCID IN (SELECT ACCID FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND FMLYCODE='" & TxtBranchCode.text & "')"
mysql = mysql & " ORDER BY A.NAME, B.EXID "
LAccRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly

Do While Not LAccRec.EOF
    RECGRID.AddNew
    RECGRID!AC_CODE = LAccRec!AC_CODE
    RECGRID!NAME = LAccRec!NAME
    RECGRID!EXCODE = LAccRec!EXCODE
    RECGRID!EXID = LAccRec!EXID
    RECGRID!UPTOSTDT = LAccRec!UPTOSTDT
    RECGRID!brokrate = LAccRec!brokrate
    RECGRID!broktype = LAccRec!broktype
    RECGRID!INSTTYPE = LAccRec!INSTTYPE
    RECGRID!ACCID = LAccRec!ACCID
    RECGRID.Update
    LAccRec.MoveNext
Loop
Label5.Caption = "Ready"
DoEvents

Set DataGrid1.DataSource = RECGRID
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Locked = True
DataGrid1.Columns(1).Width = 4000
DataGrid1.Columns(2).Locked = True
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 1000
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(7).Visible = False
DataGrid1.Columns(5).Alignment = dbgRight
DataGrid1.Columns(5).NumberFormat = "0.00000"

DataGrid1.ReBind
DataGrid1.Refresh
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 4500
DataGrid1.Columns(2).Width = 1200
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(4).Width = 1200
DataGrid1.Columns(5).Width = 1200
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(7).Visible = False
DataGrid1.Columns(5).Alignment = dbgRight
DataGrid1.Columns(5).NumberFormat = "0.00000"

End Sub


Private Sub Get_ExIDs()
    Dim I As Integer
    LExIDS = vbNullString
    For I = 1 To ExListView.ListItems.Count
        If ExListView.ListItems(I).Checked = True Then
            If LenB(LExIDS) <> 0 Then LExIDS = LExIDS & ", "
            LExIDS = LExIDS & ExListView.ListItems(I).ListSubItems(1) & ""
        End If
    Next
End Sub

