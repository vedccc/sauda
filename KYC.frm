VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form kycfrm 
   BackColor       =   &H00FFC0C0&
   Caption         =   "KYC Form"
   ClientHeight    =   11595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11595
   ScaleWidth      =   18795
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   12255
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   11895
         Begin VB.TextBox Text2 
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
            Height          =   405
            Left            =   9720
            TabIndex        =   14
            Text            =   "Text2"
            Top             =   6600
            Width           =   1455
         End
         Begin VB.TextBox Text1 
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
            Height          =   405
            Left            =   6960
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   6600
            Width           =   1455
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   6255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   11033
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Op Bal"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   6255
            Left            =   6000
            TabIndex        =   11
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   11033
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "OP BAL"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
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
            Left            =   8880
            TabIndex        =   15
            Top             =   6675
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Debit"
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
            Left            =   6360
            TabIndex        =   12
            Top             =   6675
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   11895
         Begin VB.TextBox TxtHeadCode 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   720
            MaxLength       =   50
            TabIndex        =   8
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox TxtHeadName 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3000
            MaxLength       =   50
            TabIndex        =   4
            Top             =   120
            Width           =   4215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
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
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   203
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Head Name"
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
            Height          =   255
            Left            =   1800
            TabIndex        =   5
            Top             =   210
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16335
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Party Head Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16095
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   8460
      Left            =   165
      TabIndex        =   6
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   14923
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   -2147483633
      ForeColor       =   4194304
      ListField       =   ""
      BoundColumn     =   ""
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000040&
      BorderWidth     =   12
      Height          =   8820
      Left            =   120
      Top             =   720
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   8820
      Left            =   3960
      Top             =   720
      Width           =   12285
   End
End
Attribute VB_Name = "kycfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rec As ADODB.Recordset
Public Fb_Press As Byte
Dim old_Code As String
Dim LHEADId As Integer
Dim old_HEAD As String
Public Code As Long
Public G_CODE As Long
Dim AccRec As ADODB.Recordset
Dim mcode As Long
Dim indrag As Boolean ' Flag that signals a Drag Drop operation.
Dim NodX As Object ' Item that is being dragged.
Dim old_Family As String
Dim PartyHeadRec As ADODB.Recordset
'Dim old_Code  As String
Sub Add_Rec()
    Fb_Press = 1: old_Family = vbNullString: old_Code = vbNullString
    Call Get_Selection(1)
    TxtHeadName.text = vbNullString
    DataList1.Locked = True
    Frame1.Enabled = True: TxtHeadName.SetFocus
End Sub
Sub CANCEL_REC()
    Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = False
    Next
    ListView2.ListItems.clear
    
End Sub
Private Sub DataList1_Click()
    TxtHeadCode.text = DataList1.BoundText
    TxtHeadName.text = DataList1.text
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


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Form_Load()
    Call CANCEL_RECORD
    Call PERMISSIONS("kycfrm")
    Call CLEAR_SCREEN
    Set PartyHeadRec = Nothing
    Set PartyHeadRec = New ADODB.Recordset
    MYSQL = "SELECT * FROM PARTYHEAD WHERE COMPCODE =" & GCompCode & " ORDER BY HEADNAME"
    PartyHeadRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not PartyHeadRec.EOF Then
        Set DataList1.RowSource = PartyHeadRec
        DataList1.ListField = "HEADName"
        DataList1.BoundColumn = "HEADCode"
    End If
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    MYSQL = "SELECT AC_CODE,NAME,OP_BAL FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " ORDER BY NAME  "
    AccRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        ListView1.ListItems.clear
        Do While Not AccRec.EOF
            ListView1.ListItems.Add , , AccRec!AC_CODE
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , AccRec!NAME
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Format(AccRec!OP_BAL, "0.00")
            
            AccRec.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 96
End Sub
Sub add_record()
    Dim TRec As ADODB.Recordset
    Dim mcode As Long
    On Error GoTo Error1
    Fb_Press = 1
    Call Get_Selection(1)
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT MAX(HeadCode) AS MNCODE FROM Partyhead  WHERE COMPCODE =" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
    mcode = Val(Rec!MNCODE & vbNullString) + 1
    TxtHeadCode.text = Str(mcode)
    TxtHeadName.text = vbNullString:
    Frame2.Enabled = True:
    TxtHeadName.SetFocus
Error1: If err.Number <> 0 Then
        If InStr(err.Description, "Object variable") Then MsgBox "Please Select Party Head", vbInformation
End If
End Sub
Sub CANCEL_RECORD()
    
    
    TxtHeadCode.text = vbNullString
    TxtHeadName.text = vbNullString
    
    Call CLEAR_SCREEN

    Fb_Press = 0
    Call Get_Selection(10)
    DataList1.Locked = False
End Sub
Sub CLEAR_SCREEN()
    G_CODE = 0: Fb_Press = 0
    TxtHeadName.text = vbNullString


End Sub
Sub Save_Rec()
    On Error GoTo ERR1
    Dim TRec As ADODB.Recordset
    Dim I As Integer
    Cnn.BeginTrans
    CNNERR = True
    If LenB(TxtHeadName.text) < 1 Then MsgBox "Party Head required before saving record.", vbCritical, "Error": TxtHeadName.SetFocus: Exit Sub
    If Fb_Press = 2 Then
        Cnn.Execute "UPDATE PARTYHEAD SET HEADNAME='" & TxtHeadName.text & "' WHERE  HEADCODE=" & Val(TxtHeadCode.text) & ""
    Else
        Set Rec = Nothing: Set Rec = New ADODB.Recordset
        Rec.Open "SELECT HeadName FROM PartyHead WHERE COMPCODE =" & GCompCode & " AND HEADNAME  ='" & TxtHeadName.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not Rec.EOF Then MsgBox "Duplicate Party Head Name  ", vbExclamation, "Warning": TxtHeadName.SetFocus: Cnn.RollbackTrans: CNNERR = False: Exit Sub
        MYSQL = " INSERT INTO PARTYHEAD(COMPCODE,HEADCODE,HEADNAME)"
        MYSQL = MYSQL & " VALUES (" & GCompCode & "," & Val(TxtHeadCode.text) & ",'" & TxtHeadName.text & "')"
        Cnn.Execute MYSQL
    End If
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            MYSQL = "UPDATE ACCOUNTM SET PTYHEAD =" & Val(TxtHeadCode.text) & " WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & ListView1.ListItems(I).text & "'"
            Cnn.Execute MYSQL
        End If
            
        Next
    
    Cnn.CommitTrans: CNNERR = False
    PartyHeadRec.Requery
    Set DataList1.RowSource = PartyHeadRec: DataList1.ListField = "HEADName": DataList1.BoundColumn = "HEADCode"
    Call CANCEL_RECORD
    'Call lblcancel_Click
    Screen.MousePointer = 0: Exit Sub
ERR1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
        'Resume
        Cnn.RollbackTrans: CNNERR = False
    End If
End Sub

Sub MODIFY_REC()
    On Error GoTo Error1
    Fb_Press = Val(2)
    Dim TRec As ADODB.Recordset
    Dim LDebit As Double
    Dim LCredit As Double
    Dim I As Integer
    Call Get_Selection(2)
    MYSQL = "SELECT * FROM PARTYHEAD WHERE COMPCODE =" & GCompCode & " AND HEADCODE=" & Val(TxtHeadCode.text) & " "
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
    TxtHeadName.text = TRec!HEADNAME
    
    MYSQL = "SELECT AC_CODE,NAME,OP_BAL FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND PTYHEAD=" & Val(TxtHeadCode.text) & "ORDER BY NAME "
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    ListView2.ListItems.clear
    If Not TRec.EOF Then
        Do While Not TRec.EOF
            ListView2.ListItems.Add , , TRec!AC_CODE
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , TRec!NAME
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , Format(TRec!OP_BAL, "0.00")
            ListView2.ListItems(ListView2.ListItems.Count).Checked = True
            If TRec!OP_BAL > 0 Then
                LCredit = LCredit + TRec!OP_BAL
            Else
                LDebit = LDebit + Abs(TRec!OP_BAL)
            End If
            
            For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).text = TRec!AC_CODE Then
                    ListView1.ListItems(I).Checked = True
                    Exit For
                End If
            Next
            
            TRec.MoveNext
        Loop
    End If
    
    Text1.text = Format(LDebit, "0.00")
    Text2.text = Format(LCredit, "0.00")
    Frame2.Enabled = True
    TxtHeadName.SetFocus

Error1: If err.Number <> 0 Then
        If err.Number = 91 Then
            MsgBox "Please Select Party Head", vbCritical
            CANCEL_RECORD
        End If
End If
End Sub
Sub Delete_Record()
    Dim TRec As ADODB.Recordset
    If MsgBox("You want to Remove this Party Head?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        MYSQL = "DELETE FROM PARTYHEAD WHERE COMPCODE =" & GCompCode & " AND HEADCODE =" & Val(TxtHeadCode.text) & ""
        Cnn.Execute MYSQL
        PartyHeadRec.Requery
        Set DataList1.RowSource = PartyHeadRec: DataList1.ListField = "HEADName": DataList1.BoundColumn = "HEADCode"
        Call CANCEL_RECORD
    Else
        Exit Sub
    End If
End Sub
