VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AccGrp 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13440
      TabIndex        =   11
      Top             =   0
      Width           =   1695
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000013&
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
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13095
      Begin VB.Frame Frame3 
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
         Height          =   615
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   13095
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Account Group Setup"
            BeginProperty Font 
               Name            =   "Verdana"
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
            TabIndex        =   10
            Top             =   120
            Width           =   12855
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
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
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   1440
      Width           =   8175
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   2
         Top             =   180
         Width           =   6375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Effect in Gross Profit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1290
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AccGrp.frx":0000
            Key             =   "GRP"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   7230
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   8415
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
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
         ForeColor       =   &H00FF8080&
         Height          =   5895
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   8175
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5535
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   9763
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   617
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
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
            OLEDragMode     =   1
         End
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   7080
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12488
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   8388736
      ListField       =   ""
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G r o u p   L i s t"
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
      Left            =   13680
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   12
      Height          =   7500
      Left            =   4320
      Top             =   720
      Width           =   8700
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C000C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   12
      Height          =   7500
      Left            =   0
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "AccGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fb_Press As Byte:    Public Code As Long:    Public G_CODE As Long:  Dim GroupRec As ADODB.Recordset
Dim indrag As Boolean
'Flag that signals a Drag Drop operation.
Dim NodX As Object ' Item that is being dragged.
Sub CLOSE_RECORD()
    Call Get_Selection(6)
    Unload Me
End Sub
Private Sub DataList1_Click()
    Text1.text = DataList1.text
End Sub
Private Sub Form_Activate()
    Call TRNODE
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Private Sub Form_Load()
    Call PERMISSIONS("AccGroup")
    Call CLEAR_SCREEN
    Set GroupRec = Nothing
    Set GroupRec = New ADODB.Recordset
    MYSQL = "SELECT G_NAME,CODE FROM AC_GROUP ORDER BY G_NAME"
    GroupRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not GroupRec.EOF Then
        Set DataList1.RowSource = GroupRec
        DataList1.ListField = "G_NAME"
        DataList1.BoundColumn = "CODE"
    End If
End Sub
Private Sub Form_Paint()
    If GETMAIN.ActiveForm.NAME = Me.NAME Then
        Call PERMISSIONS("AccGroup")
        Call Get_Selection(Fb_Press)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CANCEL_RECORD
    GETMAIN.StatusBar1.Panels(1).text = vbNullString: Set GeneralRst = Nothing: Unload Me
End Sub
Sub Save_Record()
    On Error GoTo err1
    Dim TRec As ADODB.Recordset
    Cnn.BeginTrans
    CNNERR = True
    If Fb_Press = Val(2) Then
        Cnn.Execute "UPDATE AC_GROUP SET G_NAME='" & Text1.text & "', GPEFFECT=" & Check1.Value & " WHERE  CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
    Else
        MYSQL = "SELECT MAX(CODE) AS MAX_NO FROM AC_GROUP"
        Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        mcode = Val(TRec!MAX_NO & vbNullString) + 1
        MYSQL = "SELECT MAX(G_CODE) AS MAX_NO FROM AC_GROUP"
        Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        G_CODE = Val(TRec!MAX_NO & vbNullString) + 1
        MYSQL = "SELECT * FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
        Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        Cnn.Execute "INSERT INTO AC_GROUP( CODE, G_CODE, G_NAME, PCODE, TYPE, GPEFFECT, SEQ, TAMT) VALUES(" & mcode & "," & G_CODE & ",'" & Text1.text & "'," & TRec!Code & ",'" & TRec!Type & "'," & Check1.Value & "," & mcode & ", 0)"
    End If
    Cnn.CommitTrans: CNNERR = False
    Call CANCEL_RECORD
    Call TRNODE
    Screen.MousePointer = 0: Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End Sub
Sub add_record()
    On Error GoTo Error1
    TreeView1.Enabled = True: Fb_Press = 1
    Call Get_Selection(1)
    RootName = TreeView1.SelectedItem.Root:
    TreeView1.Enabled = False:
    Text1.text = vbNullString:
    Frame2.Enabled = True:
    Text1.SetFocus
Error1: If err.Number <> 0 Then
        If InStr(err.Description, "Object variable") Then MsgBox "Please SELECT group.", vbInformation
End If
End Sub
Sub CANCEL_RECORD()
    Call Get_Selection(5)
    Call CLEAR_SCREEN
End Sub
Sub CLEAR_SCREEN()
    G_CODE = 0: Fb_Press = 0: Text1.text = "": TreeView1.Enabled = True: Frame2.Enabled = False
End Sub
Sub MODIFY_REC()
    On Error GoTo Error1
    Fb_Press = Val(2)
    Call Get_Selection(2)
    If Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) = Val(0) Then    ''NOT PRIMARY GROUP
        MsgBox "Primary Group, You can not modify this Group.", vbCritical, "Error"
        Get_Selection (10)
        Exit Sub
    Else
        RootName = TreeView1.SelectedItem.Root
    End If
    MYSQL = "SELECT * FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key))
    Set GeneralRst = Nothing
    Set GeneralRst = New ADODB.Recordset
    GeneralRst.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly

    Text1.text = GeneralRst!g_name
    If RootName = "EXPENCES" Then Check1.Value = IIf(GeneralRst!GPEFFECT = 1, 1, 0)
    TreeView1.Enabled = False
    Frame2.Enabled = True
    Text1.SetFocus
Error1: If err.Number <> 0 Then
            If err.Number = 91 Then
                MsgBox "Please SELECT group", vbCritical
                CANCEL_RECORD
            End If
        End If
End Sub
Sub Delete_Record()
    If MsgBox("You want to Remove this Group?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        If TreeView1.SelectedItem.Children = Val(0) Then
            MYSQL = "SELECT * FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
            Set GeneralRst = Nothing: Set GeneralRst = New ADODB.Recordset: GeneralRst.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            If GeneralRst!KEYGROUP & "" = "P" Then
                MsgBox "Parent Group can not be DELETE.", vbCritical, "Warning"
                Call CANCEL_RECORD
                Exit Sub
            End If

            MYSQL = "SELECT AC_CODE FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND GRPCODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
            Set GeneralRst = Nothing: Set GeneralRst = New ADODB.Recordset: GeneralRst.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRst.EOF Then
                MsgBox "Group can not be DELETE.Child group exists.", vbCritical, "Error"
                Call CANCEL_RECORD
                Exit Sub
            Else
                If MsgBox("Confirm DELETE?", vbYesNo) = vbYes Then
                    MYSQL = "DELETE FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
                    Cnn.Execute MYSQL
                Else
                    Exit Sub
                End If
            End If
        Else
            MsgBox "Group can not be Delete. Child group exists.", vbCritical, "Warning"
            Call CANCEL_RECORD
            Exit Sub
        End If
    End If
    Call TRNODE
End Sub
Sub TRNODE()
    Dim SNODE As Node
    Dim PTR, LVL As Byte
    TreeView1.Nodes.Clear
    TreeView1.LabelEdit = False
    TreeView1.LineStyle = tvwRootLines
    
    With TreeView1
        .Sorted = True
        .LabelEdit = False
        .LineStyle = tvwRootLines
    End With

    MYSQL = "SELECT * FROM AC_GROUP ORDER BY SEQ"
    Set GeneralRst = Nothing
    Set GeneralRst = New ADODB.Recordset
    GeneralRst.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    Do While Not GeneralRst.EOF
        If GeneralRst!PCODE = Val(0) Then
            Set SNODE = TreeView1.Nodes.Add(, tvwChild, "A" & CStr(GeneralRst!Code), GeneralRst!g_name)
        Else
            Set SNODE = TreeView1.Nodes.Add("A" & CStr(GeneralRst!PCODE), tvwChild, "A" & GeneralRst!Code, CStr(GeneralRst!g_name))
        End If
        GeneralRst.MoveNext
    Loop
    TreeView1.Nodes.Item("A29").Expanded = True: TreeView1.Nodes.Item("A29").ForeColor = &HFF0000
    TreeView1.Nodes.Item("A30").Expanded = True: TreeView1.Nodes.Item("A30").ForeColor = &HFF0000
    TreeView1.Nodes.Item("A31").Expanded = True: TreeView1.Nodes.Item("A31").ForeColor = &HFF0000
    TreeView1.Nodes.Item("A32").Expanded = True: TreeView1.Nodes.Item("A32").ForeColor = &HFF0000
    TreeView1.Nodes.Item("A29").Selected = True
    TreeView1.SetFocus
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &H0&
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &HC00000
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Save_Record
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    Text1.text = Trim(Text1.text)
End Sub
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Set NodX = TreeView1.SelectedItem  ' Set the item being dragged.
End Sub
