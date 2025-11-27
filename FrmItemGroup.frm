VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmItemGroup 
   Caption         =   "Form1"
   ClientHeight    =   11460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11460
   ScaleWidth      =   19770
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   4095
      Begin MSDataListLib.DataList DataList1 
         Height          =   6960
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   12277
         _Version        =   393216
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   4080
      TabIndex        =   9
      Top             =   1320
      Width           =   10575
      Begin MSDataListLib.DataCombo ExCombo 
         Height          =   420
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   9975
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ItemCode"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ExCode"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange"
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
         Left            =   480
         TabIndex        =   10
         Top             =   203
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   10575
      Begin VB.TextBox TxtGroupId 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   0
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox TxtGroupName 
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
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   1
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name"
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
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14655
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Item Group Setup"
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
         TabIndex        =   5
         Top             =   120
         Width           =   14415
      End
   End
End
Attribute VB_Name = "FrmItemGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemRec As ADODB.Recordset
Dim ItemGroupRec As ADODB.Recordset
Dim ExRec As ADODB.Recordset
Dim LGroupID As Long
Dim LGroupName As String
Public Fb_Press As Byte

Sub Add_Rec()
    Dim TRec As ADODB.Recordset
    
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    MYSQL = "SELECT MAX(GROUPID) AS MNO FROM ITEMGROUP WHERE COMPCODE =" & GCompCode & ""
    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        If Not IsNull(TRec!MNo) Then
            TxtGroupId.text = Str(TRec!MNo)
        Else
            TxtGroupId.text = "1"
        End If
        TxtGroupId.text = "1"
    End If
    Fb_Press = 1:
    
    TxtGroupName.text = vbNullString
    ExCombo.BoundText = vbNullString
    Call Get_Selection(1)
    DataList1.Locked = True
    Frame1.Enabled = True:
    Frame2.Enabled = True:
    TxtGroupName.SetFocus
End Sub

Private Sub Form_Load()
Set ItemRec = Nothing
Set ItemRec = New ADODB.Recordset
MYSQL = "SELECT ITEMCODE,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXCHANGECODE,ITEMCODE"
ItemRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
Set ItemGroupRec = Nothing
Set ItemGroupRec = New ADODB.Recordset
MYSQL = "SELECT DISTINCT GROUPID,GROUPNAME FROM ITEMGROUP WHERE COMPCODE =" & GCompCode & " ORDER BY GROUPID "
ItemGroupRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
Set ExRec = Nothing
Set ExRec = New ADODB.Recordset
MYSQL = "SELECT EXCODE FRoM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE"
ExRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
Set ExCombo.RowSource = ExRec
ExCombo.BoundColumn = "EXCODE"
ExCombo.ListField = "EXCODE"
If Not ItemGroupRec.EOF Then
        Set DataList1.RowSource = ItemGroupRec
        DataList1.ListField = "GROUPNAME"
        DataList1.BoundColumn = "GROUPID"
    End If
    ListView1.ListItems.Clear
    Do While Not ItemRec.EOF
        ListView1.ListItems.Add , , ItemRec!ITEMCODE
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ItemRec!EXCHANGECODE
        ItemRec.MoveNext
    Loop
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
End Sub
Sub CANCEL_REC()
    TxtGroupId.text = vbNullString
    TxtGroupName.text = vbNullString
    ExCombo.BoundText = vbNullString
    
    Set ItemGroupRec = Nothing: Set ItemGroupRec = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT GROUPID,GROUPNAME FROM ITEMGROUP WHERE COMPCODE =" & GCompCode & " ORDER BY GROUPID "
    ItemGroupRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not ItemGroupRec.EOF Then Set DataList1.RowSource = ItemGroupRec: DataList1.ListField = "GROUPNAME": DataList1.BoundColumn = "GROUPID"
    Fb_Press = 0
    Call Get_Selection(10)
    DataList1.Locked = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = False
    Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub



Sub Save_Rec()
    Dim TRec As ADODB.Recordset
    Dim LGroupID  As Long
    Dim LGroupName As String
    Dim LItemCode As String
    Dim LExCode As String
    Dim I As Integer
    On Error GoTo ERR1
    LGroupID = Val(TxtGroupId.text)
    LGroupName = Trim(TxtGroupName.text)
    If LenB(LGroupName) < 1 Then MsgBox "Group Name required before saving record.", vbCritical, "Error": TxtGroupName.SetFocus: Exit Sub
    CNNERR = True
    Cnn.BeginTrans
    If Fb_Press = 1 Then
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT GROUPNAME FROM ITEMGROUP WHERE COMPCODE =" & GCompCode & " AND GROUPNAME ='" & LGroupName & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then MsgBox "Duplicate Item Group Name ", vbExclamation, "Warning": TxtGroupName.SetFocus: Exit Sub
    Else
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT GROUPNAME  FROM ITEMGROUP WHERE COMPCODE=" & GCompCode & " AND GROUPNAME ='" & LGroupName & "' AND GROUPID <>" & LGroupID & "", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then MsgBox "Group Name already exists.", vbExclamation, "Warning": TxtGroupName.SetFocus: Exit Sub
    End If
    MYSQL = "DELETE FROM ITEMGROUP WHERE COMPCODE =" & GCompCode & " AND GROUPID  =" & LGroupID & ""
    Cnn.Execute MYSQL
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            LItemCode = ListView1.ListItems(I).ListSubItems(1)
            MYSQL = "INSERT INTO ITEMGROUP (COMPCODE,GROUPID,GROUPNAME,EXCODE,ITEMCODE) VALUES "
            MYSQL = MYSQL & "(" & GCompCode & "," & LGroupID & ",'" & LGroupName & "','" & LExCode & "','" & LItemCode & "')"
            Cnn.Execute MYSQL
        End If
    Next
    Call CANCEL_REC
    GETMAIN.bwtbal.Visible = True
    Cnn.CommitTrans
    CNNERR = False
    Exit Sub
ERR1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    Screen.MousePointer = 0:
    'Resume
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End Sub

Sub MODIFY_REC()
    Dim I As Integer
    Dim TRec As ADODB.Recordset
    ExCombo.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    
    If Trim(TxtGroupId.text) <> "" Then
        DataList1.Locked = True
        MYSQL = "SELECT * FROM ITEMGROUP WHERE COMPCODE =" & GCompCode & " AND GROUPID =" & Val(TxtGroupId.text) & ""
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        If Not TRec.EOF Then
            TxtGroupName = TRec!GroupName
            Frame1.Enabled = True: TxtGroupName.SetFocus
            Do While Not TRec.EOF
                For I = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(I).text = TRec!ITEMCODE Then
                        ListView1.ListItems(I).Checked = True
                        Exit For
                    End If
                Next
                TRec.MoveNext
            Loop
        End If
    Else
        MsgBox "Please Select Item Group ", vbCritical
        Call CANCEL_REC
        DataList1.Locked = False
        Frame1.Enabled = False
        Frame2.Enabled = False
        DataList1.SetFocus
    End If
End Sub



Private Sub DataList1_Click()
    TxtGroupId.text = DataList1.BoundText
    TxtGroupName.text = DataList1.text
End Sub
Private Sub DataList1_DblClick()
    If DataList1.Locked Then
    Else
        Call Get_Selection(2)
        Fb_Press = 2
        Call MODIFY_REC
    End If
End Sub
Private Sub DataList1_KeyPress(KeyAscii As Integer)
    If DataList1.Locked Then
    Else
        If KeyAscii = 13 Then
            Call Get_Selection(2)
            Fb_Press = 2
            Call MODIFY_REC
        End If
    End If
End Sub

