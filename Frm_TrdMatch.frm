VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form frm_trdmatch 
   Caption         =   "Trade Match"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   3255
      Left            =   -240
      TabIndex        =   21
      Top             =   6240
      Width           =   17775
      Begin MSComctlLib.ListView ListView3 
         Height          =   2535
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ConDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "OConno"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         Caption         =   "Matched Trades"
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
         TabIndex        =   23
         Top             =   120
         Width           =   17655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   17535
      Begin VB.CommandButton Command2 
         Caption         =   "Un Match"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8040
         TabIndex        =   9
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Match"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8040
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   3375
         Left            =   9600
         TabIndex        =   20
         Top             =   840
         Width           =   7935
         Begin MSComctlLib.ListView ListView2 
            Height          =   3135
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5530
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   255
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Condate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Rate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "OConno"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   3375
         Left            =   0
         TabIndex        =   17
         Top             =   840
         Width           =   7935
         Begin MSComctlLib.ListView ListView1 
            Height          =   3135
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   5530
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483646
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Condate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Rate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "OConno"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         Caption         =   "Un Matched Trades"
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
         TabIndex        =   24
         Top             =   0
         Width           =   17535
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Sell Trade"
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
         Left            =   9600
         TabIndex        =   19
         Top             =   360
         Width           =   7935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Buy Trade"
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
         TabIndex        =   18
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   17535
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   17295
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   420
            Left            =   4800
            TabIndex        =   2
            Top             =   120
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   741
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton Command1 
            Caption         =   "List"
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
            Left            =   14640
            TabIndex        =   4
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
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
            Height          =   405
            Left            =   15960
            TabIndex        =   5
            Top             =   120
            Width           =   1095
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   405
            Left            =   1680
            TabIndex        =   1
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
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
            Value           =   40439.5330902778
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   420
            Left            =   9360
            TabIndex        =   3
            Top             =   120
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   741
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
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
            Left            =   8640
            TabIndex        =   26
            Top             =   203
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Party"
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
            Left            =   3960
            TabIndex        =   25
            Top             =   143
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Upto Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   195
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
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
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17535
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5640
         TabIndex        =   10
         Top             =   240
         Width           =   4455
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Trade Match"
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
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   4215
         End
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frm_trdmatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LFromDate As Date:              Dim LToDate As Date:    Dim LParty As String
Dim LSAUDA As String:               Dim LEXCODE As String:  Dim LItemCode As String
Dim LBConNo As Double:              Dim LSConno As Double:  Dim LBCondate As Date
Dim LPartyCode As String
Dim LSaudaCode As String
Dim LScondate As Date
Dim PartyRec As ADODB.Recordset:    Dim SaudaRec As ADODB.Recordset
Dim BuyTradeRec As ADODB.Recordset: Dim SellTradeRec As ADODB.Recordset
Dim MatchTradeRec As ADODB.Recordset:

Private Sub Command1_Click()
    LPartyCode = DataCombo3.BoundText
    LSaudaCode = DataCombo1.BoundText
    MYSQL = "SELECT CONNO,CONDATE,CONTYPE,QTY,RATE,ROWNO1 FROM CTR_D WHERE COMPCODE =" & MC_CODE & " "
    MYSQL = MYSQL & " AND PARTY ='" & LPartyCode & "' AND SAUDA ='" & LSaudaCode & "' "
    MYSQL = MYSQL & " AND CONDATE <='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY CONDATE,CONNO"
    Set LTrdRec = Nothing
    Set LTrdRec = New ADODB.Recordset
    LTrdRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    Do While Not LTrdRec.EOF
        If LTrdRec!CONTYPE = "B" Then
            ListView1.ListItems.Add , , CStr(LTrdRec!Condate)
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LTrdRec!QTY
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LTrdRec!Rate
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LTrdRec!ROWNO1
        Else
            ListView2.ListItems.Add , , CStr(LTrdRec!Condate)
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , LTrdRec!QTY
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , LTrdRec!Rate
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , LTrdRec!ROWNO1
        End If
        'If LTrdRec!CONFIRM = 1 Then
        '    ListView1.ListItems(ListView1.ListItems.Count).Checked = True
        'End If
        LTrdRec.MoveNext
    Loop
    MsgBox ""
End Sub


Private Sub DataCombo1_Validate(Cancel As Boolean)
If LenB(DataCombo1.BoundText) = 0 Then
    MsgBox "Please Select Sauda"
    Cancel = True
End If
End Sub

Private Sub DataCombo3_Validate(Cancel As Boolean)
If LenB(DataCombo3.BoundText) = "" Then
    MsgBox "Please Select Party"
    Cancel = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub Form_Load()
vcDTP1.Value = Date
Set PartyRec = Nothing
Set PartyRec = New ADODB.Recordset
MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & MC_CODE & " ORDER BY NAME"
PartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
If Not PartyRec.EOF Then
    PartyRec.MoveFirst
    Set DataCombo3.RowSource = PartyRec
    DataCombo3.ListField = "NAME"
    DataCombo3.BoundColumn = "AC_CODE"
    DataCombo3.Refresh
Else
    MsgBox "No Parties"
End If

Set SaudaRec = Nothing
Set SaudaRec = New ADODB.Recordset
MYSQL = "SELECT SAUDACODE,SAUDANAME,EXCODE, ITEMCODE,MATURITY FROM SAUDAMAST WHERE COMPCODE =" & MC_CODE & " ORDER BY EXCODE,ITEMCODE,MATURITY"
SaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
 If Not SaudaRec.EOF Then
    Set DataCombo1.RowSource = SaudaRec
    DataCombo1.ListField = "SAUDANAME"
    DataCombo1.BoundColumn = "SAUDACODE"
    DataCombo1.Refresh
Else
    MsgBox "No Contracts"
End If

End Sub
