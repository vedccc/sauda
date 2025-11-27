VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form RptView 
   BackColor       =   &H00004080&
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21090
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10350
   ScaleWidth      =   21090
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   8655
      Left            =   5160
      TabIndex        =   5
      Top             =   1080
      Width           =   15135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   8415
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   14843
         _Version        =   393216
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   23
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4935
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Select All"
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   43020.5154050926
      End
      Begin MSComctlLib.ListView PartyList 
         Height          =   7335
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   12938
         View            =   3
         MultiSelect     =   -1  'True
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
            Text            =   "Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   4410
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   43020.5154050926
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20175
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   19935
      End
   End
End
Attribute VB_Name = "RptView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AccRec  As ADODB.Recordset
Dim LToday As Date
Dim LSParties As String
Dim RecSet As ADODB.Recordset
Private Sub Check1_Click()

Dim I As Integer
    For I = 1 To PartyList.ListItems.Count
        If Check1.Value = 1 Then
            PartyList.ListItems.Item(I).Checked = True
        Else
            PartyList.ListItems.Item(I).Checked = False
        End If
    Next I

End Sub

Private Sub Form_Load()
vcDTP1.Value = Date
vcDTP2.Value = Date
If GRptViewType = "Query on Standing" Then
    Label2.Caption = "Upto Date"
    Label3.Visible = False
    vcDTP2.Visible = False
ElseIf GRptViewType = "Query on Trade" Then
    Label2.Caption = "From Date"
    Label3.Caption = "To Date"
    Label3.Visible = True
    vcDTP2.Visible = True
End If
Label1.Caption = GRptViewType
Call Fill_Parties

End Sub


Public Function Get_Parties() As String
Dim LFParties   As String
Dim LCount As Integer
LFParties = ""
LCount = PartyList.ListItems.Count
For I = 1 To PartyList.ListItems.Count
    If PartyList.ListItems(I).Checked = True Then
        LCount = LCount - 1
        If LFParties <> "" Then LFParties = LFParties & ", "
        LFParties = LFParties & "'" & PartyList.ListItems(I).ListSubItems(1) & "'"
    End If
Next
Get_Parties = LFParties
End Function

Private Sub PartyList_Click()
If GRptViewType = "Query on Standing" Then
    Call Get_Standing
ElseIf GRptViewType = "Query on Trade" Then
    Call Get_Trade
End If
End Sub

Public Sub Set_Rec()
    If GRptViewType = "Query on Standing" Then
        Set RecSet = Nothing
        Set RecSet = New ADODB.Recordset
        RecSet.Fields.Append "Party", adVarChar, 100, adFldIsNullable
        RecSet.Fields.Append "Sauda", adVarChar, 50, adFldIsNullable
        RecSet.Fields.Append "BuyQty", adDouble, , adFldIsNullable
        RecSet.Fields.Append "SellQty", adDouble, , adFldIsNullable
        RecSet.Open , , adOpenKeyset, adLockOptimistic
    ElseIf GRptViewType = "Query on Trade" Then
        Set RecSet = Nothing
        Set RecSet = New ADODB.Recordset
        RecSet.Fields.Append "Condate", adVarChar, 100, adFldIsNullable
        RecSet.Fields.Append "Party", adVarChar, 100, adFldIsNullable
        RecSet.Fields.Append "Sauda", adVarChar, 50, adFldIsNullable
        RecSet.Fields.Append "TradeNo", adDouble, , adFldIsNullable
        RecSet.Fields.Append "BQty", adDouble, , adFldIsNullable
        RecSet.Fields.Append "BRate", adDouble, , adFldIsNullable
        RecSet.Fields.Append "SQty", adDouble, , adFldIsNullable
        RecSet.Fields.Append "SRate", adDouble, , adFldIsNullable
        RecSet.Fields.Append "ConTime", adVarChar, 30, adFldIsNullable
        RecSet.Open , , adOpenKeyset, adLockOptimistic
    End If
    
End Sub


Public Sub Fill_Parties()
Set AccRec = Nothing
Set AccRec = New ADODB.Recordset
MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " "
MYSQL = MYSQL & " AND AC_CODE IN (SELECT DISTINCT PARTY FROM CTR_D WHERE COMPCODE =" & GCompCode & " "
MYSQL = MYSQL & " AND SAUDA IN (SELECT DISTINCT SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'))"
MYSQL = MYSQL & " ORDER BY NAME  "
AccRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
PartyList.ListItems.clear
PartyList.Visible = False
Do While Not AccRec.EOF
    PartyList.ListItems.Add , , AccRec!NAME
    PartyList.ListItems(PartyList.ListItems.Count).ListSubItems.Add , , AccRec!AC_CODE
    AccRec.MoveNext
Loop
PartyList.Visible = True

End Sub

Private Sub Get_Standing()
Dim LStdRec As ADODB.Recordset
Dim LLastdate As Date
Dim TRec As ADODB.Recordset
LSParties = Get_Parties
Set DataGrid1.DataSource = Nothing
If LSParties <> "" Then
    MYSQL = "SELECT A.Name,B.Sauda,SUM(CASE CONTYPE  WHEN 'B' THEN QTY ELSE QTY*-1  END) AS StdQty From CTR_D AS B, ACCOUNTD AS A"
    MYSQL = MYSQL & " WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE AND A.AC_CODE =B.PARTY AND B.PARTY IN (" & LSParties & ") "
    MYSQL = MYSQL & " AND SAUDA IN (SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY>'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
    MYSQL = MYSQL & " AND B.CONDATE < ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    MYSQL = MYSQL & " GROUP BY  A.NAME,B.SAUDA HAVING SUM(CASE CONTYPE  WHEN 'B' THEN QTY ELSE QTY * -1  END)<>0 "
    MYSQL = MYSQL & " ORDER BY A.NAME,B.SAUDA"
    Set LStdRec = Nothing
    Set LStdRec = New ADODB.Recordset
    LStdRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    Call Set_Rec
    If Not LStdRec.EOF Then
        Do While Not LStdRec.EOF
            RecSet.AddNew
            RecSet!Party = LStdRec!NAME
            RecSet!Sauda = LStdRec!Sauda
            If LStdRec!stdqty > 0 Then
                RecSet!BuyQty = LStdRec!stdqty
            Else
                RecSet!SellQty = Abs(LStdRec!stdqty)
            End If
            RecSet.Update
            LStdRec.MoveNext
        Loop
        Set DataGrid1.DataSource = RecSet
        DataGrid1.Columns(0).Width = 4000
        DataGrid1.Columns(1).Width = 5000
        DataGrid1.Columns(2).Width = 1000
        DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(2).Alignment = dbgRight
        DataGrid1.Columns(3).Alignment = dbgRight
    End If
End If
End Sub
Private Sub Get_Trade()
Dim LTradeRec As ADODB.Recordset
Dim TRec As ADODB.Recordset
LSParties = Get_Parties
Set DataGrid1.DataSource = Nothing
If LSParties <> "" Then
    MYSQL = "SELECT B.CONDATE,A.Name,B.Sauda,B.CONNO,B.CONTYPE,B.QTY,B.RATE,B.CONTIME From CTR_D AS B, ACCOUNTD AS A"
    MYSQL = MYSQL & " WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE AND A.AC_CODE =B.PARTY AND B.PARTY IN (" & LSParties & ") "
    MYSQL = MYSQL & " AND SAUDA IN (SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY>'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
    MYSQL = MYSQL & " AND B.CONDATE >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND B.CONDATE <='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
    MYSQL = MYSQL & " ORDER BY A.NAME,b.CONDATE,B.EXCODE,B.SAUDA,B.CONNO"
    Set LTradeRec = Nothing
    Set LTradeRec = New ADODB.Recordset
    LTradeRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    Call Set_Rec
    If Not LTradeRec.EOF Then
        Do While Not LTradeRec.EOF
            RecSet.AddNew
            RecSet!TRADENO = LTradeRec!CONNO
            RecSet!Condate = LTradeRec!Condate
            RecSet!Party = LTradeRec!NAME
            RecSet!Sauda = LTradeRec!Sauda
            RecSet!CONTIME = LTradeRec!CONTIME
            If LTradeRec!CONTYPE = "B" Then
                RecSet!BQty = LTradeRec!QTY
                RecSet!BRate = LTradeRec!Rate
            Else
                RecSet!SQty = Abs(LTradeRec!QTY)
                RecSet!SRate = Abs(LTradeRec!Rate)
            End If
            RecSet.Update
            LTradeRec.MoveNext
        Loop
        Set DataGrid1.DataSource = RecSet
        DataGrid1.Columns(0).Width = 1200 ' CONDATE
        DataGrid1.Columns(1).Width = 3000 ' PARTY
        DataGrid1.Columns(2).Width = 3500 'SAUDA
        DataGrid1.Columns(3).Width = 1200 'CONNO
        DataGrid1.Columns(4).Width = 1000 'BUYQTY
        DataGrid1.Columns(5).Width = 1000 'BUYRATE
        DataGrid1.Columns(6).Width = 1000 'SELLQTY
        DataGrid1.Columns(7).Width = 1000 'SELLRATE
        DataGrid1.Columns(8).Width = 1400 'SELLRATE
        DataGrid1.Columns(3).Alignment = dbgRight
        DataGrid1.Columns(4).Alignment = dbgRight
        DataGrid1.Columns(5).Alignment = dbgRight
        DataGrid1.Columns(6).Alignment = dbgRight
        DataGrid1.Columns(7).Alignment = dbgRight
        DataGrid1.Columns(5).NumberFormat = "0.00"
        DataGrid1.Columns(7).NumberFormat = "0.00"
    End If
        
End If

End Sub
