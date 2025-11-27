VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form FrmTrfBalAcc 
   Caption         =   "Transfer Account Balance"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19005
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   19005
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   8775
      Begin VB.CommandButton Command2 
         Caption         =   "Show A/c List"
         Height          =   495
         Left            =   5040
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Transfer"
         Height          =   495
         Left            =   6960
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6975
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   12303
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bal As on Date"
         Height          =   255
         Left            =   120
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   1320
         Width           =   8535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer To A/c"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vou Date"
         Height          =   255
         Left            =   120
         TabIndex        =   6
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
      TabIndex        =   10
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "FrmTrfBalAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AccRec As ADODB.Recordset
Dim LVou_Dt As Date
Dim LToAcc As String
Dim LAccs As String
Dim LBalDate As Date

Private Sub Command1_Click()
Dim LToAcc As String:           Dim LToAccName As String
Dim I As Integer:               Dim LAcCode As String
Dim LAcName As String:          Dim LBal As Double
Dim LPtyDrCr As String:         Dim LCDrCr As String
Dim LNarr As String
Dim LVOU_ID As Long

Dim LVou_No As String
Dim LVNo As Long
Dim LACCID As Long
Dim TRec As ADODB.Recordset
Dim LToAccID As Long
LToAcc = DataCombo1.BoundText
If LenB(LToAcc) < 1 Then
    MsgBox " Please Select To Party"
    DataCombo1.SetFocus
    Exit Sub
End If
Set TRec = Nothing
Set TRec = New ADODB.Recordset
MYSQL = "SELECT ACCID,AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & LToAcc & "'"
TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
If TRec.EOF Then
    MsgBox " Please Select Valid To Party"
    DataCombo1.SetFocus
    Exit Sub
Else
    LToAccName = TRec!Name
    LToAccID = TRec!ACCID
End If

For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(I).Checked = True Then
        LAcName = vbNullString
        LAcCode = ListView1.ListItems(I).text
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        MYSQL = "SELECT ACCID,AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & LAcCode & "'"
        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            LAcName = TRec!Name
            LACCID = TRec!ACCID
        End If
        LBal = Get_ClosingBal(LAcCode, vcDTP2.Value)
        If LBal <> 0 Then
            If LBal > 0 Then
                LPtyDrCr = "D"
                LCDrCr = "C"
            ElseIf LBal < 0 Then
                LPtyDrCr = "C"
                LCDrCr = "D"
            End If
            LVOU_ID = PInsert_Voucher("JV", vcDTP1.Value, vbNullString, "Balance Transfer", LVNo, "3", 0, vbNullString)
            'LNarr = "Balance Transfer to " & LToAcc & " " & LToAccName & ""
            LNarr = "Balance Transfer to " & LToAccName & ""
            Call PInsert_VchAmt("JV", vcDTP1.Value, LPtyDrCr, LAcCode, Abs(LBal), LNarr, LVNo, vbNullString, vbNullString, LVOU_ID, LACCID)
            'LNarr = "Balance Transfer from " & LAcCode & " " & LAcName & ""
            LNarr = "Balance Transfer from " & LAcName & ""
            Call PInsert_VchAmt("JV", vcDTP1.Value, LCDrCr, LToAcc, Abs(LBal), LNarr, LVNo, vbNullString, vbNullString, LVOU_ID, LToAccID)
            'Call PInsert_VchAmt(LVou_No, "JV", vcDTP1.Value, LCDrCr, LToAcc, Abs(LBal), vbNullString, vcDTP1.Value, LNarr, vbNullString, vbNullString, 0, vbNullString)
        End If
    End If
Next
MsgBox " Transfer Complete"
Call Fill_List


End Sub

Private Sub Command2_Click()
Call Fill_List
End Sub

Private Sub Form_Load()
vcDTP1.Value = Date
vcDTP2.Value = Date
Set AccRec = Nothing
Set AccRec = New ADODB.Recordset
MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " ORDER BY NAME  "
AccRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
Call Fill_List
AccRec.MoveFirst
Set DataCombo1.RowSource = AccRec
    DataCombo1.BoundColumn = "AC_CODE"
    DataCombo1.ListField = "NAME"

End Sub

Private Sub Fill_List()
Dim LBal As Double
ListView1.ListItems.clear
AccRec.MoveFirst
Do While Not AccRec.EOF
    LBal = Get_ClosingBal(AccRec!AC_CODE, vcDTP2.Value)
    If LBal <> 0 Then
        ListView1.ListItems.Add , , AccRec!AC_CODE
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , AccRec!Name
        If LBal > 0 Then
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ""
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(LBal, "0.00")
        ElseIf LBal < 0 Then
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(Abs(LBal), "0.00")
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ""
        End If
    End If
    AccRec.MoveNext
Loop
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub


