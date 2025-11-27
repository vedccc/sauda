VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_TRDCONFIRM 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   12165
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3255
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   16335
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   15120
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   13215
         Begin MSDataListLib.DataCombo ExCombo 
            Height          =   360
            Left            =   4920
            TabIndex        =   3
            Top             =   120
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "FRM_TRDCONFIRM.frx":0000
            Left            =   7440
            List            =   "FRM_TRDCONFIRM.frx":000D
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   11880
            TabIndex        =   7
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   10560
            TabIndex        =   6
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Show Trades"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9240
            TabIndex        =   5
            Top             =   120
            Width           =   1215
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   360
            Left            =   960
            TabIndex        =   1
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   40439.5330902778
         End
         Begin vcDateTimePicker.vcDTP vcDTP2 
            Height          =   360
            Left            =   2760
            TabIndex        =   2
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   40439.5330902778
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "ExCode"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   4200
            TabIndex        =   24
            Top             =   173
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "To Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   2400
            TabIndex        =   17
            Top             =   173
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "From Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   173
            Width           =   855
         End
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   225
         Left            =   3480
         TabIndex        =   9
         Top             =   750
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   6600
         TabIndex        =   11
         Top             =   750
         Width           =   1095
      End
      Begin MSComctlLib.ListView ItemLst 
         Height          =   2100
         Left            =   4680
         TabIndex        =   10
         Top             =   1080
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   3704
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   8388736
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sauda Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SaudaCode"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView PartList 
         Height          =   2100
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   3704
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388736
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Party Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   735
         Left            =   8880
         TabIndex        =   25
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   750
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   19
         Top             =   750
         Width           =   960
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Updating Contracts Please Wait"
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
         Left            =   7800
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   5655
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   17535
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   16695
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Daily Trade Confirmation"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   16455
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   17520
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM_TRDCONFIRM.frx":002E
            Key             =   "down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM_TRDCONFIRM.frx":0480
            Key             =   "up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   16335
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   7646
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   -2147483644
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Party"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sauda"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "TradeDate"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Buy Qty"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Buy Rate"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Sell Qty"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Sell Rate"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Trade No"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "AC_CODE"
            Object.Width           =   2028
         EndProperty
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7860
      Left            =   0
      Top             =   720
      Width           =   16605
   End
End
Attribute VB_Name = "FRM_TRDCONFIRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExRec As ADODB.Recordset
Dim LParties As String
Dim LSExCode As String

Private Sub Check1_Click()
    Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        If Check1.Value = 1 Then
            ListView1.ListItems.Item(I).Checked = True
        Else
            ListView1.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub Check2_Click()
    Dim I As Integer
    For I = 1 To PartList.ListItems.Count
        If Check2.Value = 1 Then
            PartList.ListItems(I).Checked = True
        Else
            PartList.ListItems(I).Checked = False
        End If
    Next I
Call ItemList
End Sub
Private Sub Check4_Click()
    Dim I As Integer
    For I = 1 To ItemLst.ListItems.Count
        If Check4.Value = 1 Then
            ItemLst.ListItems(I).Checked = True
        Else
            ItemLst.ListItems(I).Checked = False
        End If
    Next I
End Sub
Private Sub Command1_Click()
Dim LPartyRec As ADODB.Recordset:   Dim LOpRec As ADODB.Recordset
Dim LTrdRec As ADODB.Recordset:     Dim LRateRec As ADODB.Recordset
Dim lOpQty As Double:               Dim LOpAmt As Double
Dim LBAmt As Double:                Dim LSAmt As Double
Dim LCloseAmt As Double:            Dim LBrokAmt As Double
Dim LGrossAmt As Double:            Dim LNetAmt As Double
Dim LCloseQty As Double:            Dim LCloseRate  As Double
Dim lOpRate  As Double:             Dim LSaudas As String
Dim LDiffAmt As Double:             Dim LPartyCode As String
Dim I As Integer:                   Dim LPartyName As String
Dim LSaudaCode As String:           Dim LCalval As Double
Dim MCount As Long

ListView1.ListItems.Clear
LParties = vbNullString
For I = 1 To PartList.ListItems.Count
    If PartList.ListItems(I).Checked = True Then
        If LenB(LParties) > 1 Then LParties = LParties & ", "
        LParties = LParties & "'" & PartList.ListItems(I).SubItems(1) & "'"
    End If
Next
LSaudas = vbNullString
For I = 1 To ItemLst.ListItems.Count
    If ItemLst.ListItems(I).Checked = True Then
        If LenB(LSaudas) > 1 Then LSaudas = LSaudas & ", "
        LSaudas = LSaudas & "'" & ItemLst.ListItems(I).SubItems(1) & "'"
    End If
Next
If LenB(LParties) < 1 Then
    MsgBox "PLease Select Parties"
    PartList.SetFocus
    Exit Sub
End If

If LenB(LSaudas) < 1 Then
    MsgBox "PLease Select Sauda "
    ItemLst.SetFocus
    Exit Sub
End If

'ListView1.ColumnHeaders(2).Icon = "up"
MYSQL = "SELECT DISTINCT A.AC_CODE,A.NAME,C.EXCHANGECODE,C.ITEMCODE,C.LOT,B.SAUDA,D.MATURITY, D.TRADEABLELOT "
MYSQL = MYSQL & " FROM ACCOUNTD AS A, CTR_D AS B, ITEMMAST AS C , SAUDAMAST AS D  WHERE A.COMPCODE =" & GCompCode & " "
MYSQL = MYSQL & " AND A.COMPCODE = B.COMPCODE AND A.COMPCODE = C.COMPCODE AND A.COMPCODE = D.COMPCODE AND A.ACCID = B.ACCID  "
MYSQL = MYSQL & " AND C.ITEMCODE =D.ITEMCODE AND D.SAUDAID = B.SAUDAID "
MYSQL = MYSQL & " AND B.CONDATE >= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND B.CONDATE <= '" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' AND A.AC_CODE IN (" & LParties & ") "
MYSQL = MYSQL & " AND B.SAUDA IN (" & LSaudas & ")"
MYSQL = MYSQL & " ORDER BY A.NAME,C.EXCHANGECODE,C.ITEMCODE,D.MATURITY"
Set LPartyRec = Nothing
Set LPartyRec = New ADODB.Recordset
LPartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
ListView1.Visible = False
If LPartyRec.EOF Then MsgBox "No Trade Found for the selected Parties And Contracts "
Label6.Visible = True
Label6.Caption = vbNullString
Do While Not LPartyRec.EOF
    Label6.Caption = " Please Wait Filling Trades"
    LOpAmt = 0: LBAmt = 0: LSAmt = 0: LCloseAmt = 0: LDiffAmt = 0: lOpQty = 0: LCloseQty = 0
    LPartyCode = LPartyRec!AC_CODE
    LPartyName = LPartyRec!NAME
    LSaudaCode = LPartyRec!Sauda
    If LPartyRec!EXCHANGECODE = "NSE" Then
        LCalval = LPartyRec!TRADEABLELOT
    Else
        LCalval = LPartyRec!LOT
    End If
    LCloseQty = lOpQty
    MYSQL = "SELECT CONNO,CONDATE,CONTYPE,QTY,RATE,BROKRATE,CONFIRM FROM CTR_D WHERE COMPCODE =" & GCompCode & " "
    MYSQL = MYSQL & " AND PARTY ='" & LPartyCode & "' AND SAUDA ='" & LSaudaCode & "' AND CONDATE >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    MYSQL = MYSQL & " AND CONDATE <='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' ORDER BY CONDATE,CONNO"
    Set LTrdRec = Nothing
    Set LTrdRec = New ADODB.Recordset
    LTrdRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    Do While Not LTrdRec.EOF
        MCount = MCount + 1
        
        DoEvents
        If LTrdRec!CONTYPE = "B" Then
            LCloseQty = LCloseQty + LTrdRec!QTY
            LBAmt = LBAmt + LTrdRec!QTY * LTrdRec!Rate * LCalval
            ListView1.ListItems.Add , , LPartyName
            ListView1.ListItems.Item(1).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LSaudaCode
            ListView1.ListItems.Item(MCount).ListSubItems(1).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , CStr(LTrdRec!Condate)
            ListView1.ListItems.Item(MCount).ListSubItems(2).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LTrdRec!QTY
            ListView1.ListItems.Item(MCount).ListSubItems(3).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(LTrdRec!Rate, "0.00")
            ListView1.ListItems.Item(MCount).ListSubItems(4).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , " "
            ListView1.ListItems.Item(MCount).ListSubItems(5).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , " "
            ListView1.ListItems.Item(MCount).ListSubItems(6).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LTrdRec!CONNO
            ListView1.ListItems.Item(MCount).ListSubItems(7).ForeColor = vbBlue
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LPartyCode
            ListView1.ListItems.Item(MCount).ListSubItems(8).ForeColor = vbBlue
        Else
            LCloseQty = LCloseQty - LTrdRec!QTY
            LSAmt = LSAmt + LTrdRec!QTY * LTrdRec!Rate * LCalval
            ListView1.ListItems.Add , , LPartyName
            ListView1.ListItems.Item(1).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LSaudaCode
            ListView1.ListItems.Item(MCount).ListSubItems(1).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , CStr(LTrdRec!Condate)
            ListView1.ListItems.Item(MCount).ListSubItems(2).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , " "
            ListView1.ListItems.Item(MCount).ListSubItems(3).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , " "
            ListView1.ListItems.Item(MCount).ListSubItems(4).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LTrdRec!QTY
            ListView1.ListItems.Item(MCount).ListSubItems(5).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(LTrdRec!Rate, "0.00")
            ListView1.ListItems.Item(MCount).ListSubItems(6).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LTrdRec!CONNO
            ListView1.ListItems.Item(MCount).ListSubItems(7).ForeColor = vbRed
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LPartyCode
            ListView1.ListItems.Item(MCount).ListSubItems(8).ForeColor = vbRed
            
        End If
        If LTrdRec!CONFIRM = 1 Then
            ListView1.ListItems(ListView1.ListItems.Count).Checked = True
        End If
        
        LTrdRec.MoveNext
    Loop
    LSAmt = LSAmt * -1
'    If LCloseQty <> 0 Then
'        MYSQL = "SELECT CLOSERATE FROM CTR_R WHERE COMPCODE =" & GCompCode  & " AND SAUDA='" & LSaudaCode & "' AND CONDATE='" & Format(vcDTP2.Value, "yyyy/mm/dd") & "'"
'        Set LRateRec = Nothing
'        Set LRateRec = New ADODB.Recordset
'        LRateRec.Open MYSQL, cnn, adOpenStatic, adLockReadOnly
'        If Not LRateRec.EOF Then
'            LCloseRate = LRateRec!CLOSERATE
'        Else
'            LCloseRate = 0
'        End If
'        LCloseQty = LCloseQty * -1
'        LCloseAmt = LCloseQty * LCloseRate * LCalval
'        If LCloseQty > 0 Then
'            MCOUNT = MCOUNT + 1
'            ListView1.ListItems.ADD , , LPartyName
'            ListView1.ListItems.Item(1).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , LSaudaCode
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(1).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , CStr(vcDTP2.Value)
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(2).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , LCloseQty
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(3).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , Format(LCloseRate, "0.00")
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(4).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(5).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(6).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , "Closing"
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(7).ForeColor = vbMagenta
'        Else
'            MCOUNT = MCOUNT + 1
'            ListView1.ListItems.ADD , , LPartyName
'            ListView1.ListItems.Item(1).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , LSaudaCode
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(1).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , CStr(vcDTP2.Value)
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(2).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(3).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(4).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , Abs(LCloseQty)
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(5).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , Format(LCloseRate, "0.00")
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(6).ForeColor = vbMagenta
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , "Closing"
'            ListView1.ListItems.Item(MCOUNT).ListSubItems(7).ForeColor = vbMagenta
'        End If
'    End If
'    LDiffAmt = LOpAmt + LBAmt + LSAmt + LCloseAmt
'    If LDiffAmt + LOpAmt + LBAmt + LSAmt + LCloseAmt <> 0 Then
'        MCOUNT = MCOUNT + 1
'        ListView1.ListItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , Format(LDiffAmt, "0.00")
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , " "
'    End If
    
    LPartyRec.MoveNext
Loop
Command2.Enabled = True
ListView1.Visible = True
Label6.Caption = vbNullString
Label6.Visible = False
End Sub
Private Sub Command2_Click()
Dim I As Integer
Frame2.Enabled = False
Label6.Visible = True
    For I = 1 To ListView1.ListItems.Count
        DoEvents
        If ListView1.ListItems(I).Checked = True Then
            If ListView1.ListItems(I).SubItems(3) <> " " Then
                MYSQL = "UPDATE CTR_D SET CONFIRM =1 WHERE COMPCODE =" & GCompCode & "  AND PARTY ='" & ListView1.ListItems(I).SubItems(8) & "' "
                MYSQL = MYSQL & " AND CONNO =" & ListView1.ListItems(I).SubItems(7) & "  AND CONDATE='" & Format(DateValue(ListView1.ListItems(I).SubItems(2)), "YYYY/MM/DD") & "'"
                MYSQL = MYSQL & " AND SAUDA ='" & ListView1.ListItems(I).SubItems(1) & "' AND CONTYPE='B'"
                Cnn.Execute MYSQL
            Else
                MYSQL = "UPDATE CTR_D SET CONFIRM =1 WHERE COMPCODE =" & GCompCode & "  AND PARTY ='" & ListView1.ListItems(I).SubItems(8) & "' "
                MYSQL = MYSQL & " AND CONNO =" & ListView1.ListItems(I).SubItems(7) & "  AND CONDATE='" & Format(DateValue(ListView1.ListItems(I).SubItems(2)), "YYYY/MM/DD") & "' "
                MYSQL = MYSQL & " AND SAUDA ='" & ListView1.ListItems(I).SubItems(1) & "' AND CONTYPE='S'"
                Cnn.Execute MYSQL
            End If
        Else
            If ListView1.ListItems(I).SubItems(3) <> " " Then
                MYSQL = "UPDATE CTR_D SET CONFIRM =0 WHERE COMPCODE =" & GCompCode & "  AND PARTY ='" & ListView1.ListItems(I).SubItems(8) & "'"
                MYSQL = MYSQL & " AND CONNO =" & ListView1.ListItems(I).SubItems(7) & "  AND CONDATE='" & Format(DateValue(ListView1.ListItems(I).SubItems(2)), "YYYY/MM/DD") & "'"
                MYSQL = MYSQL & " AND SAUDA ='" & ListView1.ListItems(I).SubItems(1) & "' AND CONTYPE='B'"
                Cnn.Execute MYSQL
            Else
                MYSQL = "UPDATE CTR_D SET CONFIRM =0 WHERE COMPCODE =" & GCompCode & "  AND PARTY ='" & ListView1.ListItems(I).SubItems(8) & "' "
                MYSQL = MYSQL & " AND CONNO =" & ListView1.ListItems(I).SubItems(7) & "  AND CONDATE='" & Format(DateValue(ListView1.ListItems(I).SubItems(2)), "YYYY/MM/DD") & "' "
                MYSQL = MYSQL & " AND SAUDA ='" & ListView1.ListItems(I).SubItems(1) & "' AND CONTYPE='S'"
                Cnn.Execute MYSQL
                
            End If
            
        End If
        Label6.Caption = " Updating Contracts Pease Wait    " & ListView1.ListItems(I).text
    Next
Call Command3_Click
ListView1.ListItems.Clear
Call PartyList
DoEvents
Call ItemList
Check1.Value = 0
Check2.Value = 0
Check4.Value = 0
Label6.Caption = ""
Label6.Visible = False
PartList.Enabled = True

Frame2.Enabled = True

    
End Sub

Private Sub Command3_Click()
Command1.Enabled = True
Command2.Enabled = False
Check2.Enabled = True
Check4.Enabled = True
PartList.Enabled = True
ItemLst.Enabled = True
vcDTP1.Enabled = True
vcDTP2.Enabled = True
Combo1.Enabled = True
End Sub
Private Sub ExCombo_GotFocus()

Sendkeys "%{DOWN}"
End Sub

Private Sub ExCombo_Validate(Cancel As Boolean)

If LenB(ExCombo.BoundText) > 0 Then
    LSExCode = ExCombo.BoundText
    ExRec.MoveFirst
    ExRec.Find "EXCODE='" & LSExCode & "'"
    If ExRec.EOF Then
        MsgBox "Invalid Exchange"
        Cancel = True
    End If
End If
    
Call PartyList
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "vcDTP1" Or Me.ActiveControl.NAME = "vcDTP2" Then
            Sendkeys "{tab}"
        End If
    End If
End Sub

Private Sub Form_Load()
LSExCode = vbNullString
Call Get_Selection(12)
Set ExRec = Nothing
Set ExRec = New ADODB.Recordset
MYSQL = "SELECT EXCODE FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER  BY EXCODE"
ExRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
Set ExCombo.RowSource = ExRec
ExCombo.ListField = "EXCODE"
ExCombo.BoundColumn = "EXCODE"
Command2.Enabled = False
Combo1.ListIndex = 0
vcDTP1.Value = Date
vcDTP2.Value = Date
Call PartyList
End Sub
Sub PartyList()
    Dim TRec  As ADODB.Recordset
    PartList.ListItems.Clear
    ItemLst.ListItems.Clear
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " "
    MYSQL = MYSQL & " AND AC_CODE IN (SELECT DISTINCT PARTY FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND Condate<='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' "
    If LenB(LSExCode) > 0 Then MYSQL = MYSQL & "AND EXCODE ='" & LSExCode & "'"
    MYSQL = MYSQL & " AND SAUDA IN (SELECT DISTINCT SAUDACODE FROM SAUDAMAST WHERE COMPCODE = " & GCompCode & " AND MATURITY>='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "')) ORDER BY NAME "
    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        PartList.Visible = False
        Do While Not TRec.EOF
            PartList.ListItems.Add , , TRec!NAME
            PartList.ListItems(PartList.ListItems.Count).ListSubItems.Add , , TRec!AC_CODE
            TRec.MoveNext
        Loop
        PartList.Visible = True
        PartList.TabStop = True
    End If
    Call ItemList
End Sub
Sub ItemList()
    ItemLst.ListItems.Clear
    LParties = vbNullString
    Dim I As Integer
    Dim TRec  As ADODB.Recordset
    For I = 1 To PartList.ListItems.Count
        If PartList.ListItems(I).Checked = True Then
            If Len(LParties) > 1 Then
                LParties = LParties & ", "
            End If
            LParties = LParties & "'"
            LParties = LParties & PartList.ListItems(I).SubItems(1)
            LParties = LParties & "'"
        End If
    Next
    If LenB(LParties) > 1 Then
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        MYSQL = "SELECT DISTINCT SAUDACODE,SAUDANAME FROM SAUDAMAST AS S ,ITEMMAST AS I WHERE S.COMPCODE =" & GCompCode & " AND S.COMPCODE=I.COMPCODE "
        If LenB(LSExCode) > 0 Then MYSQL = MYSQL & " AND EXCODE ='" & LSExCode & "'"
        MYSQL = MYSQL & " AND S.SAUDACODE IN (SELECT DISTINCT SAUDA FROM CTR_D WHERE COMPCODE =" & GCompCode & " "
        MYSQL = MYSQL & " AND PARTY IN (" & LParties & ") AND CONDATE>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND CONDATE<='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "') ORDER BY SAUDANAME "
        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            ItemLst.Visible = False
            Do While Not TRec.EOF
                ItemLst.ListItems.Add , , TRec!SAUDANAME
                ItemLst.ListItems(ItemLst.ListItems.Count).ListSubItems.Add , , TRec!SAUDACODE
                TRec.MoveNext
            Loop
            ItemLst.Visible = True
            ItemLst.TabStop = True
        End If
    End If
End Sub


Private Sub PartList_Click()
Call ItemList
End Sub
Private Sub PartList_Validate(Cancel As Boolean)
Call ItemList
End Sub
Private Sub vcDTP1_Validate(Cancel As Boolean)
Call PartyList
End Sub
Private Sub vcDTP2_Validate(Cancel As Boolean)
Call PartyList
End Sub
Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim I As Integer
    For I = 1 To ListView1.ColumnHeaders.Count
        If ListView1.ColumnHeaders(I).Index <> CurrentHeader Then
            ListView1.ColumnHeaders(I).Icon = Empty
        End If
    Next
End Sub

Sub OLD()
'MYSQL = "SELECT A.PARTY,A.CONNO,A.CONSNO,A.CONDATE,A.CONTYPE,A.QTY,A.RATE,C.NAME,B.SAUDANAME,C.NAME,A.CONFIRM FROM CTR_D AS A, SAUDAMAST AS B,ACCOUNTD AS C WHERE A.COMPCODE=B.COMPCODE  AND A.COMPCODE=C.COMPCODE AND A.PARTY=C.AC_CODE AND A.SAUDA=B.SAUDACODE AND A.PARTY IN (" & LParties & ") AND A.SAUDA IN (" & LSaudas & ") "
'
'MYSQL = MYSQL & "AND CONDATE >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND CONDATE <='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
'If Combo1.ListIndex = 1 Then
'    MYSQL = MYSQL & " AND A.CONFIRM  =1 "
'ElseIf Combo1.ListIndex = 2 Then
'    MYSQL = MYSQL & " AND A.CONFIRM  =0 "
'End If
'vcDTP1.Enabled = False
'vcDTP2.Enabled = False
'Combo1.Enabled = False
'PartList.Enabled = False
'ItemLst.Enabled = False
'Label6.Visible = True
'ListView1.Visible = False
'Check2.Enabled = False
'Check4.Enabled = False
'MYSQL = MYSQL & "ORDER BY A.CONDATE,A.PARTY,A.SAUDA,A.CONNO"
'
'Dim LBUYQTY, LBUYRATE, LSELLQTY, LSELLRATE
'LBUYQTY = 0: LBUYRATE = 0: LSELLQTY = 0: LSELLRATE = 0: MCount = 0
'Set Rec = Nothing
'Set Rec = New ADODB.Recordset
'Rec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'If Not Rec.EOF Then
'    ListView1.Visible = False
'    Do While Not Rec.EOF
'        DoEvents
'        Label6.Caption = "Loading Trade List " & Rec!Condate & ""
'        MCount = MCount + 1
'        If Rec!CONTYPE = "B" Then
'            ListView1.ListItems.Add , , Rec!Condate
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!SAUDANAME
'            ListView1.ListItems.Item(MCount).ListSubItems(1).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!NAME
'            ListView1.ListItems.Item(MCount).ListSubItems(2).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!QTY
'            ListView1.ListItems.Item(MCount).ListSubItems(3).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(Rec!Rate, "0.00")
'            ListView1.ListItems.Item(MCount).ListSubItems(4).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LSELLQTY
'            ListView1.ListItems.Item(MCount).ListSubItems(5).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(LSELLRATE, "0.00")
'            ListView1.ListItems.Item(MCount).ListSubItems(6).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!PARTY
'            ListView1.ListItems.Item(MCount).ListSubItems(7).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!CONTYPE
'            ListView1.ListItems.Item(MCount).ListSubItems(8).ForeColor = vbBlue
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!CONNO
'            ListView1.ListItems.Item(MCount).ListSubItems(9).ForeColor = vbBlue
'        Else
'            ListView1.ListItems.Add , , Rec!Condate
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!SAUDANAME
'            ListView1.ListItems.Item(MCount).ListSubItems(1).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!NAME
'            ListView1.ListItems.Item(MCount).ListSubItems(2).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , LBUYQTY
'            ListView1.ListItems.Item(MCount).ListSubItems(3).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(LBUYRATE, "0.00")
'            ListView1.ListItems.Item(MCount).ListSubItems(4).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!QTY
'            ListView1.ListItems.Item(MCount).ListSubItems(5).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(Rec!Rate, "0.00")
'            ListView1.ListItems.Item(MCount).ListSubItems(6).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!PARTY
'            ListView1.ListItems.Item(MCount).ListSubItems(7).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!CONTYPE
'            ListView1.ListItems.Item(MCount).ListSubItems(8).ForeColor = vbRed
'            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Rec!CONNO
'            ListView1.ListItems.Item(MCount).ListSubItems(9).ForeColor = vbRed
'        End If
'        If Rec!CONFIRM = 1 Then
'            ListView1.ListItems(ListView1.ListItems.Count).Checked = True
'        End If
'        Rec.MoveNext
'    Loop
'    ListView1.Visible = True
'    Command1.Enabled = False
'    Command2.Enabled = True
'Else
'    Call Command3_Click
'    MsgBox "No Records Found"
'End If

End Sub
