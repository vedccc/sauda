VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form GETACTYP 
   BackColor       =   &H80000001&
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11595
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   7320
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
            Picture         =   "GETACTYP.frx":0000
            Key             =   "GRP"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Effect in Gross Profit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Group Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   555
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   2535
         Left            =   120
         Top             =   120
         Width           =   4455
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11880
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   1235
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11670
   End
End
Attribute VB_Name = "GETACTYP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fb_press As Byte
Dim REC As ADODB.Recordset
Dim MYRS As ADODB.Recordset
Public Code As Long
Public G_CODE As Long
Dim indrag As Boolean ' Flag that signals a Drag Drop operation.
Dim nodX As Object ' Item that is being dragged.
Sub list()
    GP_LIST.ReportSource = crptReport
    GP_LIST.ReportFileName = RPT_PATH + "GROUP.RPT"
    GP_LIST.WindowState = crptMaximized
    GP_LIST.Action = 2
End Sub
Sub CLOSE_RECORD()
    Call Get_Selection(6)
    Unload Me
End Sub
Private Sub Form_Activate()
    Call TRNODE
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Private Sub Form_Load()
    Call CLEAR_SCREEN
    Shape2.Height = 1575
    Frame2.Height = 1815
End Sub
Private Sub Form_Paint()
    If GETMAIN.ActiveForm.Name = Me.Name Then
        GETACTYP.BackColor = BACK_COLOR
        Call Get_Selection(fb_press)
    End If
End Sub
Private Sub Form_Unload(cancel As Integer)
    Call CANCEL_RECORD
    GETMAIN.StatusBar1.Panels(1).Text = " "
    Unload Me
End Sub
Sub save_record()
    On Error GoTo ERR1
    CNN.BeginTrans
    CNNERR = True

    If fb_press = Val(2) Then
        MYSQL = "UPDATE AC_GROUP SET G_NAME='" & Text1.Text & "', GPEFFECT=" & Check1.Value & " WHERE compcode = " & MC_CODE & " and CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
        CNN.Execute MYSQL
    
    Else
        MYSQL = "SELECT MAX(CODE) AS MAX_NO FROM AC_GROUP"
        Set MYRS = Nothing
        Set MYRS = New ADODB.Recordset
        MYRS.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
        MCODE = Val(MYRS!MAX_NO & "") + Val(1)

        MYSQL = "SELECT * FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
        Set MYRS = Nothing
        Set MYRS = New ADODB.Recordset
        MYRS.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly

        MYSQL = "INSERT INTO AC_GROUP(CODE, G_CODE, G_NAME, PCODE, TYPE, GPEFFECT, SEQ, TAMT) VALUES(" & MCODE & "," & MYRS!G_CODE & ",'" & Text1.Text & "'," & MYRS!Code & ",'" & MYRS!Type & "'," & Check1.Value & "," & MCODE & ",0)"
        CNN.Execute MYSQL
    End If

    CNN.CommitTrans
    CNNERR = False

    Call CANCEL_RECORD

    Call TRNODE
        
    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
    If CNNERR = True Then
        CNN.RollbackTrans
        CNNERR = False
    End If
End Sub
Sub ADD_RECORD()
    TreeView1.Enabled = True
    fb_press = 1
    Call Get_Selection(1)

    ROOTNAME = TreeView1.SelectedItem.Root

    If ROOTNAME = "EXPENCES" Or ROOTNAME = "INCOME" Then
        Shape2.Height = 2725
        Frame2.Height = 3015
    Else
        Shape2.Height = 1575
        Frame2.Height = 1815
    End If
    
    TreeView1.Enabled = False

    Text1.Text = ""
    Frame2.Visible = True
    Text1.SetFocus
End Sub
Sub CANCEL_RECORD()
    Call Get_Selection(5)
    Call CLEAR_SCREEN
End Sub
Sub CLEAR_SCREEN()
    G_CODE = 0: fb_press = 0: Text1.Text = ""
    TreeView1.Enabled = True
    Frame2.Visible = False
End Sub
Sub MODIFY_REC()
    fb_press = Val(2)

    Call Get_Selection(2)

    If Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) = Val(0) Then    ''NOT PRIMARY GROUP
        MsgBox "Primary Group, You can not modify this Group.", vbCritical, "Error"
        Get_Selection (10)
        Exit Sub
    Else
        ROOTNAME = TreeView1.SelectedItem.Root
    End If

    MYSQL = "SELECT * FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key))
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly

    Text1.Text = REC!G_NAME
    If ROOTNAME = "EXPENCES" Then
        Check1.Value = IIf(REC!GPEFFECT = 1, 1, 0)
        Shape2.Height = 2725
        Frame2.Height = 3015
    Else
        Shape2.Height = 1575
        Frame2.Height = 1815
    End If

    TreeView1.Enabled = False
    Frame2.Visible = True
    Text1.SetFocus
End Sub
Sub Delete_Record()
    If MsgBox("You want to Remove this Group?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        If Val(TreeView1.SelectedItem.Key) > Val(0) Then   ''NOT MAIN GROUP
            If TreeView1.SelectedItem.Children = Val(0) Then
                MYSQL = "SELECT * FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
                Set MYRS = Nothing
                Set MYRS = Nothing
                Set MYRS = New ADODB.Recordset
                MYRS.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
                If MYRS!KEYGROUP & "" = "P" Then
                    MsgBox "SORRY : Key group, You can not Remove Group.", vbCritical, "Warning"
                    Call CANCEL_RECORD
                    Exit Sub
                End If

                MYSQL = "SELECT * FROM ACCOUNT WHERE compcode=" & MC_CODE & " AND GRPCODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
                Set MYRS = Nothing
                Set MYRS = Nothing
                Set MYRS = New ADODB.Recordset
                MYRS.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
                If Not MYRS.EOF Then
                    MsgBox "SORRY : Some Account has been defined on this Group.", vbCritical, "Error"
                    Call CANCEL_RECORD
                    Exit Sub
                Else
                    MYSQL = "DELETE FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
                    CNN.Execute MYSQL
                End If
            Else
                MsgBox "SORRY : Some Child Node Exists!.", vbCritical, "Warning"
                Call CANCEL_RECORD
                Exit Sub
            End If
        End If
    Else
        Call CANCEL_RECORD
    End If

    Call TRNODE
End Sub
Sub TRNODE()
    Dim SNODE As Node
    Dim PTR, LVL As Byte
    TreeView1.Nodes.clear
    TreeView1.LabelEdit = False
    TreeView1.LineStyle = tvwRootLines

    With TreeView1
        .Sorted = True
        .LabelEdit = False
        .LineStyle = tvwRootLines
    End With

    MYSQL = "SELECT * FROM AC_GROUP ORDER BY SEQ"
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    Do While Not REC.EOF
        If REC!PCODE = Val(0) Then
            Set SNODE = TreeView1.Nodes.ADD(, tvwChild, "A" & CStr(REC!Code), StrConv(REC!G_NAME, vbProperCase))
            SNODE.BackColor = RGB(200, 200, 200)
        Else
            Set SNODE = TreeView1.Nodes.ADD("A" & CStr(REC!PCODE), tvwChild, "A" & REC!Code, StrConv(REC!G_NAME, vbProperCase))
            SNODE.BackColor = RGB(100, 200, 255)
        End If
        REC.MoveNext
    Loop
    TreeView1.SetFocus
End Sub
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Set nodX = TreeView1.SelectedItem  ' Set the item being dragged.
End Sub
Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
   If Button = vbLeftButton Then ' Signal a Drag operation.
      indrag = True
      TreeView1.DragIcon = TreeView1.SelectedItem.CreateDragImage
      TreeView1.Drag vbBeginDrag ' Drag operation.
   End If
End Sub
Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If TreeView1.DropHighlight Is Nothing Then
        Set TreeView1.DropHighlight = Nothing
        indrag = False
        Exit Sub
    Else
        If DELPERM = True Then
            If CStr(TreeView1.SelectedItem.Key) = CStr(TreeView1.DropHighlight.Key) Then Exit Sub
    
            If Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) = Val(0) Then
                MsgBox "Main Group, You can not Move it.", vbCritical, "Error"
                Exit Sub
            End If
            
            Set REC = Nothing
            Set REC = New ADODB.Recordset
            MYSQL = "SELECT * FROM AC_GROUP WHERE CODE=" & Mid(TreeView1.DropHighlight.Key, 2, Len(TreeView1.DropHighlight.Key)) & ""
            REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
            Set REC.ActiveConnection = Nothing

            MYSQL = "UPDATE AC_GROUP SET SEQ=SEQ + 1 WHERE compcode = " & MC_CODE & " and SEQ >" & REC!SEQ & ""
            CNN.Execute MYSQL

            MYSQL = "UPDATE AC_GROUP SET SEQ= " & REC!SEQ + Val(1) & ", PCODE=" & Mid(TreeView1.DropHighlight.Key, 2, Len(TreeView1.DropHighlight.Key)) & " WHERE compcode = " & MC_CODE & " and CODE=" & Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key)) & ""
            CNN.Execute MYSQL

            Set TreeView1.DropHighlight = Nothing
            indrag = False
            Call TRNODE
            TreeView1.SetFocus
        End If
    End If
End Sub
Private Sub TreeView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If DELPERM = True Then
        If indrag = True Then
           Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
        End If
    End If
End Sub
