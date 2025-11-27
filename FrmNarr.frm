VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmNarr 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Narration Setup"
   ClientHeight    =   11415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11415
   ScaleWidth      =   18750
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Narration Master Setup"
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
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   11655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   8175
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   8055
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   6495
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Narration"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.TextBox NARRCODETEXT 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   6000
      Left            =   45
      TabIndex        =   0
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   10583
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   -2147483633
      ForeColor       =   4194304
      ListField       =   ""
      BoundColumn     =   ""
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
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   6300
      Left            =   3840
      Top             =   720
      Width           =   7965
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   6300
      Left            =   0
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "FrmNarr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fb_Press As Byte
Dim old_Code As String
Dim LNarrId As Integer
Dim old_NArr As String
Dim LSLABCODE  As String
Dim NarrRec As ADODB.Recordset
Sub Add_Rec()
    Fb_Press = 1: old_Code = ""
    Call Get_Selection(1)
    Text1.text = ""
    DataList1.Locked = True
    Frame1.Enabled = True: Text1.SetFocus
End Sub
Sub Save_Rec()
    If Trim(Text1.text) = "" Then MsgBox "Narration required before saving record.", vbCritical, "Error": FmlyCodeTxt.SetFocus: Exit Sub
    If Fb_Press = 1 Then
        Set Rec = Nothing: Set Rec = New ADODB.Recordset
        Rec.Open "SELECT NARRNAME FROM NARRATIONM WHERE COMPCODE =" & GCompCode & " AND NARRNAME ='" & Text1.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not Rec.EOF Then MsgBox "Duplicate Narration ", vbExclamation, "Warning":  Exit Sub
        'Set Rec = Nothing: Set Rec = New ADODB.Recordset
        'Rec.Open "SELECT MAX(NARRCODE) AS MNCODE FROM NARRATIONM WHERE COMPCODE =" & GCompCode  & "", cnn, adOpenForwardOnly, adLockReadOnly
        MYSQL = "INSERT INTO NARRATIONM (COMPCODE,NARRNAME)"
'      changes by rohit
        MYSQL = MYSQL & " VALUES (" & GCompCode & ",'" & Trim(Text1.text) & "')"
        Cnn.Execute MYSQL
        
        
    End If
'    changes by rohit
    If Fb_Press = 2 Then
        Set Rec = Nothing: Set Rec = New ADODB.Recordset
        Rec.Open "SELECT * FROM NARRATIONM WHERE COMPCODE =" & GCompCode & " and NARRNAME<>'" & old_NArr & "' AND NARRNAME='" & Text1.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not Rec.EOF Then MsgBox "Duplicate Narration ", vbExclamation, "Warning":  Exit Sub
    
        MYSQL = "UPDATE NARRATIONM SET "
        MYSQL = MYSQL & " NARRNAME ='" & Text1.text & "' WHERE COMPCODE =" & GCompCode & " and NARRCODE=" & old_Code & ""
        Cnn.Execute MYSQL
    End If
    MYSQL = "SELECT * FROM NARRATIONM WHERE COMPCODE =" & GCompCode & ""
    Set NarrRec = Nothing: Set NarrRec = New ADODB.Recordset
    NarrRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not NarrRec.EOF Then Set DataList1.RowSource = NarrRec: DataList1.ListField = "NARRName": DataList1.BoundColumn = "NARRCode"
    Call CANCEL_REC
    GETMAIN.bwtbal.Visible = True
End Sub
Sub CANCEL_REC()
    Text1.text = vbNullString
    Fb_Press = 0
    Call Get_Selection(10)
    DataList1.Locked = False
    Frame1.Enabled = False
End Sub
Sub MODIFY_REC()
    If Trim(NARRCODETEXT.text) <> "" Then
        DataList1.Locked = True
        NarrRec.MoveFirst
        NarrRec.Find "NARRCODE='" & NARRCODETEXT.text & "'", , adSearchForward
        If Not NarrRec.EOF Then
            NARRCODETEXT.text = NarrRec!NARRCode & ""
            Text1.text = NarrRec!NarrName & ""
            old_Code = NARRCODETEXT.text & ""
            Frame1.Enabled = True
            Text1.SetFocus
            old_NArr = Text1.text
        End If
        If Fb_Press = 3 Then
            If MsgBox("You are about to Delete one record. Confirm Delete ?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                Set Rec = Nothing: Set Rec = New ADODB.Recordset
                MYSQL = "SELECT NARRATION FROM VCHAMT WHERE COMPCODE =" & GCompCode & " AND NARRATION  =" & LNARRCODE & ""
                Rec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
                If Not Rec.EOF Then
                    MsgBox "Transaction exists with " & Rec!NAME, vbExclamation, "Error"
                    Call CANCEL_REC
                    Exit Sub
                Else
                    MYSQL = "DELETE FROM NARRATIONM WHERE COMPCODE =" & GCompCode & " AND NARRCODE ='" & LNARRCODE & "'"
                    Cnn.Execute MYSQL
                    MYSQL = "SELECT NARRCode FROM NARRATIONM  WHERE COMPCODE =" & GCompCode & ""
                    Set Rec = Nothing: Set Rec = New ADODB.Recordset: Rec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
                    If Not Rec.EOF Then
                        GETMAIN.bwtbal.Visible = True
                    Else
                        GETMAIN.bwtbal.Visible = False
                    End If
                End If
            End If
            Call CANCEL_REC
            DataList1.Locked = False
            DataList1.SetFocus
        End If
    Else
        MsgBox "Please Select Narration.", vbCritical
        Call CANCEL_REC
        DataList1.Locked = False
        DataList1.SetFocus
    End If
End Sub


Private Sub DataList1_Click()
    NARRCODETEXT.text = DataList1.BoundText
    Text1.text = DataList1.text
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
    Call CANCEL_REC
    Set NarrRec = Nothing: Set NarrRec = New ADODB.Recordset
    MYSQL = "SELECT NARRCODE,NARRNAME FROM NARRATIONM WHERE COMPCODE =" & GCompCode & " ORDER BY NARRNAME"
    NarrRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not NarrRec.EOF Then
        Set DataList1.RowSource = NarrRec
        DataList1.ListField = "NARRName"
        DataList1.BoundColumn = "NARRCode"
        
    End If
    
    
    
    
End Sub
Private Sub Form_Paint()
'    Me.BackColor = GETMAIN.BackColor
    'Frame1.BackColor = Me.BackColor
End Sub
Sub LIST_ITEM()
    Screen.MousePointer = 11
    Call Get_Selection(12)
    MYSQL = "SELECT NARR.NARRCode as ITEMCODE,NARR.NARRName as ITEMNAME FROM NARRATIONM as NARR WHERE NARR.COMPCODE =" & GCompCode & " "
    Set Rec = Nothing: Set Rec = New ADODB.Recordset: Rec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "RptFList.RPT", 1)
    RDCREPO.DiscardSavedData: RDCREPO.Database.SetDataSource Rec
    CRViewer1.Width = CInt(GETMAIN.Width - 100)
    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Top = 0: CRViewer1.Left = 0
    CRViewer1.Visible = True: CRViewer1.ReportSource = RDCREPO
    CRViewer1.ViewReport: Set RPT = Nothing: Screen.MousePointer = 0
End Sub
Sub Delete_Record()
Dim X As Boolean
    X = MsgBox("Are you Sure you Want to Delete This Record", vbYesNo)
    If X = True Then
        Cnn.Execute "DELETE FROM NARRATIONM WHERE COMPCODE =" & GCompCode & " AND NARRNAME ='" & Text1.text & "'"
    End If
    Set NarrRec = Nothing: Set NarrRec = New ADODB.Recordset
    MYSQL = "SELECT NARRCODE,NARRNAME FROM NARRATIONM WHERE COMPCODE =" & GCompCode & " ORDER BY NARRNAME"
    NarrRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    Call CANCEL_REC
End Sub


