VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form webuser 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Web User Master"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   17790
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   17790
   ShowInTaskbar   =   0   'False
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
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   1815
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   3240
      TabIndex        =   12
      Top             =   2040
      Width           =   9375
      Begin VB.CheckBox REPOCHK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   8
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CheckBox acc_chk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   6
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CheckBox fmlychk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   6000
         Width           =   1215
      End
      Begin MSComctlLib.ListView FMLYLIST 
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Press : F7 to select all, F8 to unselect."
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   9340
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
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   6862
         EndProperty
      End
      Begin MSComctlLib.ListView PERMLIST 
         Height          =   5295
         Left            =   6480
         TabIndex        =   7
         ToolTipText     =   "Press : F7 to select all, F8 to unselect."
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   9340
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
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   6722
         EndProperty
      End
      Begin MSComctlLib.ListView ACC_LIST 
         Height          =   5295
         Left            =   3000
         TabIndex        =   20
         ToolTipText     =   "Press : F7 to select all, F8 to unselect."
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   9340
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
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   6862
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts"
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
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permissions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   14
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Family"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7395
      Left            =   3165
      TabIndex        =   9
      Top             =   1170
      Width           =   9615
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   75
         TabIndex        =   16
         Top             =   120
         Width           =   9375
         Begin VB.TextBox ucodetxt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   1641
            MaxLength       =   6
            TabIndex        =   1
            Top             =   60
            Width           =   975
         End
         Begin VB.TextBox unametxt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3798
            MaxLength       =   20
            TabIndex        =   2
            Top             =   60
            Width           =   1575
         End
         Begin VB.TextBox upwdtxt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   6960
            MaxLength       =   6
            PasswordChar    =   "$"
            TabIndex        =   3
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Code"
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
            Left            =   240
            TabIndex        =   19
            Top             =   127
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   2895
            TabIndex        =   18
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   5655
            TabIndex        =   17
            Top             =   120
            Width           =   900
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Web User Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3015
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   7350
      Left            =   75
      TabIndex        =   0
      Top             =   1170
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   12965
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   -2147483633
      ListField       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7695
      Left            =   0
      Top             =   1080
      Width           =   2805
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7695
      Left            =   3000
      Top             =   1080
      Width           =   10005
   End
End
Attribute VB_Name = "webuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fb_Press As Byte
Dim GeneralRec  As ADODB.Recordset
Dim Reportrec As ADODB.Recordset
Dim ACCOUNTS As String
Dim REPORTS As String
Dim TEMPR As String
Dim FLAGE As Boolean
Sub add_record()
    Fb_Press = 1: Frame1.Enabled = True
    Call Get_Selection(Fb_Press)
    DataList1.Locked = True
    ucodetxt.Locked = False
    ucodetxt.text = ""
    unametxt.text = ""
    upwdtxt.text = ""
    ucodetxt.SetFocus
    Call ReportList
    Call FMLYLST
End Sub
Sub CANCEL_RECORD()
    Call ClearFormFn(webuser)
    FmlyList.ListItems.clear
    ACC_LIST.ListItems.clear
    PERMLIST.ListItems.clear
    DataList1.Locked = False
    DataList1.SetFocus
    Frame1.Enabled = False
    Call Get_Selection(10)
End Sub
Sub Save_Record()
If Trim(ucodetxt.text) = "" Then
    MsgBox "User Code Required Before Save.Please Enter User Code", vbCritical
    Exit Sub
End If
If Trim(unametxt.text) = "" Then
    MsgBox "User Name Required Before Save.Please Enter User Name", vbCritical
    Exit Sub
End If
If Trim(upwdtxt.text) = "" Then
    MsgBox "User Password Required Before Save.Please Enter User Password", vbCritical
    Exit Sub
End If
Call ACCOUNTCHK
If LenB(ACCOUNTS) = 0 Then
    MsgBox "Please Select Account.", vbCritical
    Exit Sub
End If
If Fb_Press = 2 Then
    Cnn.Execute "DELETE FROM WEBUSER WHERE COMPCODE= " & MC_CODE & " AND USERCODE='" & DataList1.BoundText & "'"
    Cnn.Execute "DELETE FROM USER_COMPANY WHERE COMPCODE= " & MC_CODE & " AND USER_NAME='" & DataList1.BoundText & "'"
    Cnn.Execute "DELETE FROM USERPERM WHERE COMPCODE= " & MC_CODE & " AND USERCODE='" & DataList1.BoundText & "'"
End If
    Cnn.Execute "INSERT INTO WEBUSER (COMPCODE,USERCODE,USERNAME,PWD)VALUES (" & MC_CODE & ",'" & Trim(ucodetxt.text) & "','" & Trim(unametxt.text) & "','" & Trim(upwdtxt.text) & "') "
    Cnn.Execute "INSERT INTO USER_COMPANY (COMPCODE,USER_NAME)VALUES (" & MC_CODE & ",'" & Trim(ucodetxt.text) & "') "
    For I = 1 To ACC_LIST.ListItems.Count
        If ACC_LIST.ListItems(I).Checked = True Then
            For J = 1 To PERMLIST.ListItems.Count
                If PERMLIST.ListItems(J).Checked = True Then
                 
                 Cnn.Execute "INSERT INTO USERPERM (COMPCODE,USERCODE,AC_CODE,RPERM)VALUES (" & MC_CODE & ",'" & Trim(ucodetxt.text) & "','" & ACC_LIST.ListItems(I).text & "','" & PERMLIST.ListItems(J).text & "') "
                End If
            Next
            Cnn.Execute "UPDATE ACCOUNTD SET  APPREP=1 WHERE COMPCODE=" & MC_CODE & " AND AC_CODE='" & ACC_LIST.ListItems(I).text & "'"
        End If
    Next
    Call CANCEL_RECORD
    Call LISTITEM
End Sub
Sub Report_chk()
 REPORTS = vbNullString
        For I = 1 To PERMLIST.ListItems.Count
            If PERMLIST.ListItems(I).Checked = True Then REPORTS = REPORTS & PERMLIST.ListItems(I).text
        Next
End Sub
Sub ACCOUNTCHK()
    ACCOUNTS = vbNullString
    For I = 1 To ACC_LIST.ListItems.Count
        If ACC_LIST.ListItems(I).Checked = True Then
            If LenB(ACCOUNTS) > 1 Then ACCOUNTS = ACCOUNTS & ", "
            ACCOUNTS = ACCOUNTS & "'" & ACC_LIST.ListItems(I).text & "'"
        End If
    Next
End Sub
Sub LISTITEM()
Set list = Nothing
Set list = New ADODB.Recordset
list.Open "SELECT * FROM WEBUSER WHERE COMPCODE=" & MC_CODE & " ORDER BY USERNAME", Cnn, adLockReadOnly
If Not list.EOF Then
    Set DataList1.RowSource = list
    DataList1.BoundColumn = "USERCODE"
    DataList1.ListField = "USERNAME"
End If
End Sub
Sub FMLYLST()
Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
    MYSQL = "SELECT FMLYNAME,FMLYCODE FROM ACCFMLY   WHERE COMPCODE =" & MC_CODE & " ORDER BY FMLYNAME "
    GeneralRec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
    If Not GeneralRec.EOF Then
        FmlyList.ListItems.clear
        FmlyList.Enabled = True: FmlyChk.Enabled = True
        Do While Not GeneralRec.EOF
            FmlyList.ListItems.Add , , GeneralRec!FMLYCODE
            FmlyList.ListItems(FmlyList.ListItems.Count).ListSubItems.Add , , GeneralRec!FMLYNAME
            FmlyList.ListItems(FmlyList.ListItems.Count).ListSubItems.Add , , GeneralRec!FMLYCODE
            GeneralRec.MoveNext
        Loop
        If GeneralRec.RecordCount = 1 Then
            FmlyChk.Value = 1: FmlyList.TabStop = False: FmlyChk.TabStop = False
        Else
            FmlyList.TabStop = True: FmlyChk.TabStop = True
        End If
    Else
        FmlyList.Enabled = False: FmlyChk.Enabled = False
    End If
    If Not Fb_Press = 1 Then
        Call FMLYCHQ
    End If
End Sub
Sub FMLYCHQ()
MYSQL = "SELECT DISTINCT F.FMLYNAME,F.FMLYCODE FROM ACCFMLYD AS F WHERE F.COMPCODE =" & MC_CODE & " AND F.PARTY IN(" & ACCOUNTS & ") ORDER BY FMLYNAME "
Set fmrec = Nothing
Set fmrec = New ADODB.Recordset
fmrec.Open MYSQL, Cnn, adLockReadOnly
If Not fmrec.EOF Then
    fmrec.MoveFirst
    Do While Not fmrec.EOF
        For I = 1 To FmlyList.ListItems.Count
             If FmlyList.ListItems(I).text = fmrec!FMLYCODE Then
                FmlyList.ListItems(I).Checked = True
            End If
        Next
        fmrec.MoveNext
    Loop
End If
End Sub
Sub ReportList()
Dim NRec As ADODB.Recordset
Set Reportrec = Nothing: Set Reportrec = New ADODB.Recordset
MYSQL = "SELECT RCODE,RNAME FROM REPORTM  ORDER BY RNAME "
Reportrec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
If Not Reportrec.EOF Then
    PERMLIST.ListItems.clear
    PERMLIST.Enabled = True: FmlyChk.Enabled = True
    I = 0
    Do While Not Reportrec.EOF
        I = I + 1
        PERMLIST.ListItems.Add , , Reportrec!RCODE
        PERMLIST.ListItems(PERMLIST.ListItems.Count).ListSubItems.Add , , Reportrec!RNAME
        PERMLIST.ListItems(PERMLIST.ListItems.Count).ListSubItems.Add , , Reportrec!RCODE
        If Fb_Press <> 1 Then
            MYSQL = "SELECT * FROM USERPERM WHERE COMPCODE =" & MC_CODE & " AND USERCODE ='" & DataList1.BoundText & "' AND RPERM ='" & Reportrec!RCODE & "'"
            Set NRec = Nothing
            Set NRec = New ADODB.Recordset
            NRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
            If NRec.EOF Then
                PERMLIST.ListItems(I).Checked = False
            Else
                PERMLIST.ListItems(I).Checked = True
            End If
        End If
        Reportrec.MoveNext
    Loop
    If Reportrec.RecordCount = 1 Then
        REPOCHK.Value = 1: PERMLIST.TabStop = False: REPOCHK.TabStop = False
    Else
        PERMLIST.TabStop = True: 'permchk.TabStop = True
    End If
Else
    PERMLIST.Enabled = False
End If
End Sub
Private Sub Acc_Chk_Click()
    For I = 1 To ACC_LIST.ListItems.Count
        If acc_chk.Value = 1 Then
            ACC_LIST.ListItems.Item(I).Checked = True
        Else
            ACC_LIST.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub Acc_Chk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub ACC_LIST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
If KeyCode = 118 Then          'F7:MAKE ALL CHECKED
    acc_chk.Value = 1
    Call Acc_Chk_Click
ElseIf KeyCode = 119 Then      'F8:MAKE ALL UNCHECKED
    acc_chk.Value = 0
    Call Acc_Chk_Click
End If
End Sub
Private Sub DataList1_Click()
If Not DataList1.Locked Then
    unametxt.text = DataList1.text
    ucodetxt.text = DataList1.BoundText
End If
End Sub
Private Sub DataList1_DblClick()
If Not DataList1.Locked Then
    Fb_Press = 2
    Call Get_Selection(2)
    Call USER_ACCESS
End If
End Sub
Private Sub DataList1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call DataList1_DblClick
End If
End Sub
Private Sub fmlychk_Click()
    For I = 1 To FmlyList.ListItems.Count
        If FmlyChk.Value = 1 Then
            FmlyList.ListItems.Item(I).Checked = True
        Else
            FmlyList.ListItems.Item(I).Checked = False
        End If
    Next I
    Call AccountList
End Sub
Sub USER_ACCESS()
If Not Fb_Press = 1 Then
    If DataList1.BoundText = "" Then
        MsgBox "Please Select Account to Modify.", vbCritical
        Call CANCEL_RECORD
        DataList1.SetFocus
        Exit Sub
    End If
    MYSQL = "SELECT USERCODE,USERNAME,PWD FROM WEBUSER WHERE COMPCODE=" & MC_CODE & " AND USERCODE='" & DataList1.BoundText & "'"
    Set USERAC = Nothing
    Set USERAC = New ADODB.Recordset
    USERAC.Open MYSQL, Cnn, adLockReadOnly
    If Not USERAC.EOF Then
        ucodetxt.text = USERAC!USERCODE
        unametxt.text = USERAC!USERNAME
        upwdtxt.text = USERAC!PWD
    End If
    ACSQL = "SELECT DISTINCT A.AC_CODE,A.NAME FROM USERPERM AS U,ACCOUNTM AS A WHERE U.COMPCODE=" & MC_CODE & " AND U.COMPCODE=A.COMPCODE AND A.AC_CODE=U.AC_CODE AND  U.USERCODE='" & DataList1.BoundText & "'"
    Set ACDATA = Nothing
    Set ACDATA = New ADODB.Recordset
    ACDATA.Open ACSQL, Cnn, adLockReadOnly
    If Not ACDATA.EOF Then
        ACC_LIST.Enabled = True
        ACDATA.MoveFirst
        Do While Not ACDATA.EOF
            ACC_LIST.ListItems.Add , , ACDATA!AC_CODE
            ACC_LIST.ListItems(ACC_LIST.ListItems.Count).ListSubItems.Add , , ACDATA!Name
            ACC_LIST.ListItems(ACC_LIST.ListItems.Count).Checked = True
            ACDATA.MoveNext
        Loop
        ACC_LIST.TabStop = True: acc_chk.TabStop = True
    Else
        ACC_LIST.Enabled = True
    End If
    Call ReportList
    Call ACCOUNTCHK
    Call FMLYLST
    DataList1.Locked = True
    ucodetxt.Locked = True
    Frame1.Enabled = True
    unametxt.SetFocus
End If
If Fb_Press = 3 Then
    Call Delete_Record
End If
End Sub
Sub Delete_Record()
If MsgBox(String(5, " ") & "You are about to DELETE this record. Are you sure . ." & String(10, " "), vbYesNo + vbCritical + vbDefaultButton1, "Confirmation") = vbYes Then
    Cnn.Execute "DELETE FROM WEBUSER WHERE COMPCODE= " & MC_CODE & " AND USERCODE='" & DataList1.BoundText & "'"
    Cnn.Execute "DELETE FROM USERPERM WHERE COMPCODE= " & MC_CODE & " AND USERCODE='" & DataList1.BoundText & "'"
    Call LISTITEM
    Call CANCEL_RECORD
Else
    Call CANCEL_RECORD
End If

End Sub
Sub AccountList()
    If FmlyList.Enabled Then
        LFmlyCode = ""
        For I = 1 To FmlyList.ListItems.Count
            If FmlyList.ListItems(I).Checked = True Then
                If Len(LFmlyCode) > 1 Then
                    LFmlyCode = LFmlyCode & ", "
                End If
                    LFmlyCode = LFmlyCode & "'"
                    LFmlyCode = LFmlyCode & FmlyList.ListItems(I).text
                    LFmlyCode = LFmlyCode & "'"
            End If
        Next
            ACC_LIST.ListItems.clear
        If LFmlyCode <> "" Then
            MYSQL = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTD AS A,ACCFMLYD AS F WHERE A.COMPCODE = " & MC_CODE & " AND F.COMPCODE= A.COMPCODE AND A.AC_CODE = F.PARTY AND F.FMLYCODE IN(" & LFmlyCode & ") ORDER BY A.NAME "
            Set Rec = Nothing
            Set Rec = New ADODB.Recordset
            Rec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not Rec.EOF Then
                ACC_LIST.Enabled = True
                Do While Not Rec.EOF
                    ACC_LIST.ListItems.Add , , Rec!AC_CODE
                    ACC_LIST.ListItems(ACC_LIST.ListItems.Count).ListSubItems.Add , , Rec!Name
                    Rec.MoveNext
                Loop
                ACC_LIST.TabStop = True: acc_chk.TabStop = True:
            Else
                ACC_LIST.Enabled = False
            End If
        End If
    End If
    If Fb_Press <> 1 Then
        Call ACCHQ
    End If
End Sub
Sub ACCHQ()
ACSQL = "SELECT A.AC_CODE,A.NAME  FROM USERPERM AS U,ACCOUNTM AS A WHERE U.COMPCODE=" & MC_CODE & " AND U.COMPCODE=A.COMPCODE AND A.AC_CODE=U.AC_CODE AND  U.USERCODE='" & DataList1.BoundText & "'"
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    Rec.Open ACSQL, Cnn, adLockReadOnly
    If Not Rec.EOF Then
        Rec.MoveFirst
        Do While Not Rec.EOF
            For I = 1 To ACC_LIST.ListItems.Count
                If ACC_LIST.ListItems(I).text = Rec!AC_CODE Then
                    ACC_LIST.ListItems(I).Checked = True
                End If
            Next
           Rec.MoveNext
        Loop
    End If
End Sub
Private Sub fmlychk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub FMLYLIST_Click()
Call AccountList
End Sub
Private Sub FMLYLIST_ItemCheck(ByVal Item As MSComctlLib.LISTITEM)
Call AccountList
End Sub
Private Sub FMLYLIST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
If KeyCode = 118 Then          'F7:MAKE ALL CHECKED
    FmlyChk.Value = 1
    Call fmlychk_Click
ElseIf KeyCode = 119 Then      'F8:MAKE ALL UNCHECKED
    FmlyChk.Value = 0
    Call fmlychk_Click
End If
End Sub
Private Sub Form_Load()
Call LISTITEM
Call Get_Selection(10)
End Sub
Private Sub PERMLIST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
If KeyCode = 118 Then          'F7:MAKE ALL CHECKED
    REPOCHK.Value = 1
    Call REPOCHK_Click
ElseIf KeyCode = 119 Then      'F8:MAKE ALL UNCHECKED
    REPOCHK.Value = 0
    Call REPOCHK_Click
End If
End Sub
Private Sub REPOCHK_Click()
For I = 1 To PERMLIST.ListItems.Count
    If REPOCHK.Value = 1 Then
        PERMLIST.ListItems.Item(I).Checked = True
    Else
        PERMLIST.ListItems.Item(I).Checked = False
    End If
Next I
End Sub
Private Sub REPOCHK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub ucodetxt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub ucodetxt_Validate(Cancel As Boolean)
    If Trim(ucodetxt.text) = "" Then
        MsgBox "User Code not To Be Null.Please Enter User Code.", vbCritical
        Cancel = True
    End If
    Cnn.Execute MYSQL
    MYSQL = "SELECT * FROM WEBUSER WHERE COMPCODE=" & MC_CODE & " AND usercode ='" & ucodetxt.text & "'"
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    Rec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not Rec.EOF Then
        MsgBox "Duplicate User"
        Cancel = True
    End If
End Sub
Private Sub unametxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub unametxt_Validate(Cancel As Boolean)
    If Trim(unametxt.text) = "" Then
        MsgBox "User Name not To Be Null.Please Enter User Name.", vbCritical
        Cancel = True
    End If
End Sub
Private Sub upwdtxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub upwdtxt_Validate(Cancel As Boolean)
    If Trim(upwdtxt.text) = "" Then
        MsgBox "Password Required Before Save.Please Enter Password.", vbCritical
        Cancel = True
    End If
End Sub
