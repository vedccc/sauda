VERSION 5.00
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form RWBook 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Correcting Books"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   795
   ClientWidth     =   17925
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   17925
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Update Order No And Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11880
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "OpenIng Standing"
      ForeColor       =   &H8000000C&
      Height          =   2655
      Left            =   11640
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CheckBox Check12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Update Opening Balances to Zero"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin vcDateTimePicker.vcDTP vcDTP3 
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   39809.5258680556
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin vcDateTimePicker.vcDTP vcDTP4 
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   39809.5258680556
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Delete Trades"
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
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Pass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1620
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Correcting Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   3600
      TabIndex        =   0
      Top             =   2760
      Width           =   5895
      Begin VB.CheckBox Check14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Post Bills "
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
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   3120
         Width           =   2295
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Delete Contract Master"
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
         Height          =   315
         Left            =   2520
         TabIndex        =   27
         Top             =   2070
         Width           =   3135
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update CTR_R "
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
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Delete Trades"
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
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Upd SaudaID"
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
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   930
         Width           =   2295
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Upd Contract No"
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
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update Margin"
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
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   39450.5345486111
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Update ContractTime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   4920
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update Interest"
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
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1500
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update Brokerage Rate"
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
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   930
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update  Bills"
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
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2070
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update Shree Account"
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
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Top             =   1500
         Width           =   3135
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   39450.5345486111
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   8400
      TabIndex        =   29
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Correcting Books"
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
      Left            =   3600
      TabIndex        =   22
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   4980
      Left            =   3480
      Top             =   2280
      Width           =   6210
   End
End
Attribute VB_Name = "RWBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MENDDATE As Date
Dim TxtRec As ADODB.Recordset
Dim MFirstConDate As Date
Sub CORECTING_BOOKS()
    Dim ldate As String:    Dim TRec As ADODB.Recordset:    Dim LDT As Date:    Dim LDT2 As Date
    Screen.MousePointer = 11
    On Error GoTo err1
    LDT = DateValue("01/04/2014")
    LDT2 = DateValue("02/04/2014")
    If Check1.Value = 1 Then ' SHREE POSTING
        Cnn.BeginTrans
        Call Shree_Posting(GFinBegin)
        Cnn.CommitTrans
        CNNERR = False
    End If
    If Check11.Value = 1 Then
        Frame2.Visible = True
        vcDTP3.Visible = True
    End If
    If Check10.Value = 1 Then
        'Call Update_SaudaLot
    End If
    If Check10.Value = 1 Then
        Dim LMdate As Date
        Dim MDT2 As Date
        Dim LFileName As String
        Dim REC_EXSAUDA As ADODB.Recordset
        Set REC_EXSAUDA = Nothing: Set REC_EXSAUDA = New ADODB.Recordset
        mysql = "SELECT CONSNO,CONDATE FROM CTR_M ORDER BY CONDATE"
        REC_EXSAUDA.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Do While Not REC_EXSAUDA.EOF
            mysql = "UPDATE CTR_D SET CONDATE='" & Format(REC_EXSAUDA!Condate, "YYYY/MM/DD") & "' WHERE CONSNO=" & REC_EXSAUDA!CONSNO & ""
            Cnn.Execute mysql
            REC_EXSAUDA.MoveNext
        Loop
    End If
   'FUNCTION TO GENERATE BILL DATES ARE FIN'CIAL YEAR
    
    If Check13.Value = 1 Then
        Call UPDATE_CTR_R
        'Call Chk_Brok(GFinBegin)
    End If
    If Check3.Value = 1 Then
        Call Update_Charges(vbNullString, vbNullString, vbNullString, vbNullString, GFinBegin, GFinEnd, False)
        'Call UpdateBrokRateType(vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    End If
    Dim LName As String
    If Check7.Value = 1 Then
        Call Update_Contract_No(vbNullString)
        'MYSQL = "SELECT A.VOU_NO,B.AC_CODE,B.NAME FROM VCHAMT AS A, ACCOUNTD AS B WHERE A.AC_CODE=B.AC_CODE AND A.VOU_TYPE='JV' AND A.VOU_NO IN "
        'MYSQL = MYSQL & " ( SELECT DISTINCT VOU_NO FROM VCHAMT WHERE AC_CODE IN ('1788',1802,'1809','1828'))"
        'MYSQL = MYSQL & " AND A.AC_CODE NOT IN ('1788',1802,'1809','1828') ORDER BY A.VOU_NO"
        'Set TREC = Nothing
        'Set TREC = New ADODB.Recordset
        'TREC.Open MYSQL, cnn, adOpenStatic, adLockReadOnly
        'Do While Not TREC.EOF
        '    LName = TREC!Name & " " & TREC!AC_CODE & ""
        '    MYSQL = "UPDATE VCHAMT SET NARRATION ='" & LName & "' WHERE VOU_NO ='" & TREC!VOU_NO & "'AND AC_CODE <>'" & TREC!AC_CODE & "'"
        '    cnn.Execute MYSQL
       '     TREC.MoveNext
        'Loop
    End If
    If Check9.Value = 1 Then
        Call UpdateMargin(vbNullString, vbNullString, GFinBegin, GFinEnd, vbNullString)
    End If
    If Check2.Value = 1 Then
        Cnn.BeginTrans: CNNERR = True
        'Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
        'MYSQL = "SELECT MIN(Condate) as MinConDate FROM CTR_M WHERE  COMPCODE = " & GCompCode & " "
        'GeneralRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        'MFirstConDate = IIf(IsNull(GeneralRec!MinConDate), GFinBegin, GeneralRec!MinConDate)
        'Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
        'MYSQL = "SELECT MAX(Condate) as MaxConDate FROM CTR_M WHERE COMPCODE = " & GCompCode & " "
        'GeneralRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        'MENDDATE = IIf(IsNull(GeneralRec!MaxConDate), GFinBegin, GeneralRec!MaxConDate)
        'Call Delete_Inv_D(vbNullString, vbNullString, vbNullString, GFinBegin)
        If BILL_GENERATION(GFinBegin, GFinEnd, vbNullString, vbNullString, vbNullString) Then
            Cnn.CommitTrans: CNNERR = False
        Else
            Cnn.RollbackTrans: CNNERR = False
        End If
        'Call Chk_Billing
    End If
    'Update Brokerage
    If Check6.Value = 1 Then
     '   Call UpdateMargin("", "", CStr(GFinBegin), CStr(GFinEnd))
        'Call UPDATE_PARTY_TYPE
        Call Update_SaudaID
    End If
    If Check5.Value = 1 Then  'personal account
        Dim TRec2 As ADODB.Recordset
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        mysql = "SELECT * FROM PARTYMULTI WHERE COMPCODE =" & GCompCode & " ORDER BY PARTY,ITEMCODE"
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        Dim LMULTI As Double
        Dim LParty   As String
        Dim LItemCode As String
    
        Do While Not TRec.EOF
            LMULTI = TRec!Rate
            LParty = TRec!PARTY
            LItemCode = TRec!ITEMCODE
    
            mysql = "SELECT CONSNO, CONNO , CONDATE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PARTY ='" & LParty & "'"
            If LenB(LItemCode) < 1 Then
                    mysql = mysql & " AND ITEMCODE NOT IN ('NIFTY','BANKNIFTY')"
            Else
                mysql = mysql & " AND ITEMCODE='" & LItemCode & "'"
            End If
            mysql = mysql & " ORDER BY CONDATE,CONSNO,CONNO"
            Set TRec2 = Nothing
            Set TRec2 = New ADODB.Recordset
            TRec2.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not TRec2.EOF
                mysql = "update CTR_D SET MULTI=" & LMULTI & " WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & TRec2!CONSNO & " AND CONNO =" & TRec2!CONNO & ""
                Cnn.Execute mysql
                TRec2.MoveNext
            Loop
            TRec.MoveNext
        Loop

        Cnn.BeginTrans: CNNERR = True
            Cnn.Execute "TRUNCATE TABLE CONTRACTMASTER"
            Cnn.Execute "TRUNCATE TABLE SCRIPTMASTER"
        Cnn.CommitTrans: CNNERR = False
    End If
    If Check4.Value Then
        Command1.Enabled = False
        Call CalInt(GFinBegin)
        Command1.Enabled = True
    End If
    If Check14.Value = 1 Then
        Call Post_Bills
    End If
    Screen.MousePointer = 0
    GETMAIN.Label1.Caption = ""
    MsgBox "Task Over Successfully.", vbExclamation
    Exit Sub
err1:

    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    Screen.MousePointer = 0
    Command1.Enabled = True
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End Sub

Private Sub Check11_Click()
If Check11.Value = 1 Then
    Frame2.Visible = True
Else
    Frame2.Visible = False
End If
End Sub







Private Sub Check8_Click()
If Check8.Value = 1 Then
    vcDTP1.Visible = True
    vcDTP2.Visible = True
Else
    vcDTP1.Visible = False
    vcDTP2.Visible = False
End If
End Sub

Private Sub Command1_Click()
   ' MYSQL = "EXEC Delete_Items"
   ' Cnn.Execute MYSQL
    Command1.Enabled = False
    'MYSQL = "exec delete_items"
    'cnn.Execute MYSQL
    Frame1.Enabled = False
    'Call Send_Telegram
    
    Call CORECTING_BOOKS
    Frame1.Enabled = True
    Command1.Enabled = True
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Command3_Click()
Call Delete_Trd
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Form_Load()
vcDTP1.Value = GFinBegin
vcDTP2.Value = GFinEnd
vcDTP3.Value = Date
vcDTP4.Value = Date
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Sub UPDATE_PARTY_TYPE()
Cnn.BeginTrans: CNNERR = True
'Cnn.Execute "UPDATE CTR_D SET PERCONT  = 'N' WHERE COMPCODE  = " & GCompCode  & ""
'MYSQL = "SELECT AC_CODE ,PERSONNELAC FROM ACCOUNTD WHERE COMPCODE  = " & GCompCode  & " AND PERSONNELAC='Y'"
'Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset: GeneralRec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
'Do While Not GeneralRec.EOF
'    MYSQL = "SELECT CONSNO ,CONNO ,CONDATE FROM CTR_D WHERE COMPCODE  = " & GCompCode  & " AND PARTY  = '" & GeneralRec!ac_code & "' ORDER BY CONDATE"
'    Set Generals
'REC1 = Nothing: Set GeneralRec1 = New ADODB.Recordset: GeneralRec1.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
'    Do While Not GeneralRec1.EOF
'        Cnn.Execute "UPDATE CTR_D SET PERCONT ='Y' WHERE COMPCODE = " & GCompCode  & " AND CONSNO =" & GeneralRec1!ConSNo & "  AND CONNO  = " & GeneralRec1!conno & " AND CONDATE ='" & Format(GeneralRec1!Condate, "yyyy/MM/dd") & "'"
'        GETMAIN.Label1.Caption = "UPDATEING OF DATE " & GeneralRec1!Condate & ""
'        GeneralRec1.MoveNext
'        DoEvents
'    Loop
 '   GeneralRec.MoveNext
'Loop
Cnn.CommitTrans: CNNERR = False
End Sub
Sub UPDATE_ORDERNO()
 Do While Not TxtRec.EOF
    Dim LLDATE As Date
    LOrdNo = TxtRec!f25
    LLDATE = TxtRec!F24
    LTRD = TxtRec!F1
    mysql = "UPDATE  CTR_D SET ORDNO ='" & LOrdNo & "' WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(LLDATE, "YYYY/MM/DD") & "' AND CONNO =" & LTRD & ""
    Cnn.Execute mysql
    TxtRec.MoveNext
 Loop
End Sub
Function FileExist(sTestFile As String) As Boolean
   Dim lSize As Long
   On Error Resume Next
   lSize = -1
   lSize = FileLen(sTestFile)
   If lSize > -1 Then
      FileExist = True
   Else
      FileExist = False
   End If
End Function
Private Sub Delete_Trd()
On Error GoTo err1
'If Text1.text = "2803" Then
    'Cnn.BeginTrans
    'mysql = "DELETE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONDATE >='" & Format(vcDTP3.Value, "YYYY/MM/DD") & "'AND CONDATE <='" & Format(vcDTP4.Value, "YYYY/MM/DD") & "'"
    'Cnn.Execute mysql
    'Cnn.CommitTrans
    'MsgBox "Trades Deleted from " & vcDTP3.Value & " to " & vcDTP4.Value
'
'Else
'    MsgBox "Invalid Admin Password"
'End If


Exit Sub
err1:
    MsgBox err.Description
    Cnn.RollbackTrans
End Sub

Public Sub Update_SaudaID()
Dim TRec As ADODB.Recordset:    Dim LSauda As String
Dim LSaudaID As Long:           Dim LCompRec  As ADODB.Recordset
Dim LCompCode  As Integer:      Dim LExID As Integer
Dim LExCode As String:          Dim LItemID  As Integer
Dim LItemCode  As String

Set LCompRec = Nothing
Set LCompRec = New ADODB.Recordset
mysql = "SELECT COMPCODE FROM COMPANY ORDER BY COMPCODE"
LCompRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
Do While Not LCompRec.EOF
    LCompCode = LCompRec!CompCode
    mysql = "SELECT EXCODE,EXID FROM EXMAST WHERE COMPCODE =" & LCompCode & ""
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    Do While Not TRec.EOF
        LExCode = TRec!excode
        LExID = TRec!EXID
        DoEvents
        GETMAIN.Label1.Caption = "Upd " & LExCode
        mysql = "UPDATE ITEMMAST SET EXID =" & LExID & " WHERE COMPCODE =" & LCompCode & " AND EXCHANGECODE ='" & LExCode & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE SAUDAMAST SET EXID =" & LExID & " WHERE COMPCODE =" & LCompCode & " AND EXCODE ='" & LExCode & "'"
        Cnn.Execute mysql
        mysql = "UPDATE PEXBROK SET EXID =" & LExID & " WHERE COMPCODE =" & LCompCode & " AND EXCODE ='" & LExCode & "'"
        Cnn.Execute mysql
        mysql = "UPDATE PEXSBROK  SET EXID =" & LExID & " WHERE COMPCODE =" & LCompCode & " AND EXCODE ='" & LExCode & "'"
        Cnn.Execute mysql
        mysql = "UPDATE PITBROK  SET EXID =" & LExID & " WHERE COMPCODE =" & LCompCode & " AND EXCODE ='" & LExCode & "'"
        Cnn.Execute mysql
        mysql = "UPDATE PITSBROK  SET EXID =" & LExID & " WHERE COMPCODE =" & LCompCode & " AND EXCODE ='" & LExCode & "'"
        Cnn.Execute mysql
        mysql = "UPDATE ACCT_EX   SET EXID =" & LExID & " WHERE COMPCODE =" & LCompCode & " AND EXCODE ='" & LExCode & "'"
        Cnn.Execute mysql
        
        TRec.MoveNext
        
    Loop
    mysql = "SELECT SAUDACODE,SAUDAID ,EXID ,ITEMID FROM SAUDAMAST WHERE COMPCODE =" & LCompCode & ""
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly

    Do While Not TRec.EOF
        LSauda = TRec!saudacode
        LSaudaID = TRec!SAUDAID
        LExID = TRec!EXID
        LItemID = TRec!itemid
        DoEvents
        GETMAIN.Label1.Caption = "Upd " & LSauda
        
        mysql = "UPDATE CTR_D SET SAUDAID =" & LSaudaID & ",EXID=" & LExID & ",ITEMID =" & LItemID & " WHERE COMPCODE =" & LCompCode & " AND SAUDA ='" & LSauda & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_M SET SAUDAID =" & LSaudaID & ", EXID=" & LExID & ",ITEMID =" & LItemID & " WHERE COMPCODE =" & LCompCode & " AND SAUDA ='" & LSauda & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_R SET SAUDAID =" & LSaudaID & ", EXID=" & LExID & " ,ITEMID =" & LItemID & " WHERE COMPCODE =" & LCompCode & " AND SAUDA ='" & LSauda & "'"
        Cnn.Execute mysql
    
        mysql = "UPDATE DLYMGN SET SAUDAID =" & LSaudaID & ", EXID=" & LExID & ",ITEMID =" & LItemID & " WHERE COMPCODE =" & LCompCode & " AND SAUDA ='" & LSauda & "' "
        Cnn.Execute mysql
    
        mysql = "UPDATE DMARGIN SET SAUDAID =" & LSaudaID & ", EXID=" & LExID & ",ITEMID =" & LItemID & " WHERE COMPCODE =" & LCompCode & " AND SAUDA ='" & LSauda & "'"
        Cnn.Execute mysql
    
        TRec.MoveNext
    Loop
    mysql = "SELECT ITEMCODE,ITEMID,EXID,LOT FROM ITEMMAST WHERE COMPCODE =" & LCompCode & ""
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly

    Do While Not TRec.EOF
        LItemCode = TRec!ITEMCODE
        LItemID = TRec!itemid
        LExID = TRec!EXID
        LCalval = TRec!lot
        
        DoEvents
        GETMAIN.Label1.Caption = "Upd " & LSauda
        
        mysql = "UPDATE SAUDAMAST SET ITEMID  =" & LItemID & " ,EXID = " & LExID & " WHERE COMPCODE =" & LCompCode & " AND ITEMCODE  ='" & LItemCode & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE PITBROK SET ITEMID  =" & LItemID & " ,EXID = " & LExID & " WHERE COMPCODE =" & LCompCode & " AND ITEMCODE  ='" & LItemCode & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE PITSBROK SET ITEMID  =" & LItemID & " ,EXID = " & LExID & " WHERE COMPCODE =" & LCompCode & " AND ITEMCODE  ='" & LItemCode & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_D SET CALVAL  =" & LCalval & " WHERE COMPCODE =" & LCompCode & " AND EXCODE <>'NSE' AND ITEMCODE  ='" & LItemCode & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_D SET CALVAL  =" & LCalval & " WHERE COMPCODE =" & LCompCode & " AND EXCODE <>'NSE' AND ITEMCODE  ='" & LItemCode & "'"
        Cnn.Execute mysql

        mysql = "UPDATE CTR_D SET CALVAL = LOT FROM CTR_D JOIN ITEMMAST ON CTR_D.ITEMID = ITEMMAST.ITEMID AND ITEMMAST.EXID IN (SELECT EXID FROM EXMAST WHERE LOTWISE = 'N' AND EXCODE = 'NSE')"
        Cnn.Execute mysql
       
        
        TRec.MoveNext
    Loop
    Dim LAC_CODE As String
        Dim LACCID As Long
        
    mysql = "SELECT AC_CODE,ACCID FROM ACCOUNTM  WHERE COMPCODE =" & LCompCode & ""
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly

    Do While Not TRec.EOF
        LAC_CODE = TRec!AC_CODE
        LACCID = TRec!ACCID
        
        DoEvents
        GETMAIN.Label1.Caption = "Upd Account " & LAC_CODE
        
        mysql = "UPDATE ACCOUNTD SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND AC_CODE ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE PEXBROK SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND AC_CODE ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE PITBROK SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND AC_CODE ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE PEXSBROK SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND PARTY  ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE PITSBROK SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND PARTY  ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE ACCT_EX SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND AC_CODE  ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE ACCFMLYD  SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND PARTY  ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_D SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND PARTY  ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE INV_D  SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND PARTY  ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE INV_D1 SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND PARTY  ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE VCHAMT SET ACCID  =" & LACCID & " WHERE COMPCODE =" & LCompCode & " AND AC_CODE ='" & LAC_CODE & "'"
        Cnn.Execute mysql
        
        TRec.MoveNext
    Loop
    
    
    LCompRec.MoveNext
Loop

End Sub


Private Sub Post_Bills()
On Error GoTo err1
    CNNERR = True
    Cnn.BeginTrans
    GETMAIN.Label1.Caption = "Deleteing Old Vouchers "
    DoEvents
    mysql = "DELETE FROM VOUCHER WHERE COMPCODE =" & GCompCode & " AND VOU_TYPE IN ('S','B','H')"
    Cnn.Execute mysql
    mysql = "UPDATE INV_D SET VOU_ID =0, POST =0 WHERE COMPCODE = " & GCompCode & " "
    Cnn.Execute mysql
    mysql = "UPDATE INV_D1 SET SVOU_ID =0,BVOU_ID=0, POST =0 WHERE COMPCODE = " & GCompCode & " "
    Cnn.Execute mysql
    
    'Call Post_Inv_d(vbNullString, vbNullString, vbNullString, GFinBegin)
    'If LBilling = True Then
        If GDailyBill = False Then
            Call Post_Inv_d3(vbNullString, vbNullString, vbNullString, GFinBegin)
        Else
            Call Post_Inv_d(vbNullString, vbNullString, vbNullString, GFinBegin)
        End If
    'End If
    Cnn.CommitTrans
    CNNERR = False
    GETMAIN.Label1.Caption = vbNullString
    DoEvents
    Exit Sub
err1:
    MsgBox err.Description
    If CNNERR = True Then
        Cnn.RollbackTrans
    End If

End Sub

