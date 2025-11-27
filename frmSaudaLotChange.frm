VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSaudaLotChange 
   Caption         =   "Sauda Lot Change"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1800
      Top             =   6840
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   13455
      Begin VB.TextBox txtnewrate 
         Alignment       =   1  'Right Justify
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
         Left            =   8280
         TabIndex        =   5
         Text            =   " "
         Top             =   1320
         Width           =   1380
      End
      Begin VB.TextBox txtoldrate 
         Alignment       =   1  'Right Justify
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
         Left            =   8280
         TabIndex        =   3
         Text            =   " "
         Top             =   840
         Width           =   1380
      End
      Begin VB.CommandButton OkCmd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Ok"
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton CancelCmd 
         BackColor       =   &H00FFFFC0&
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
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2760
         Width           =   855
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   120
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   42079.7973611111
      End
      Begin MSDataListLib.DataCombo DoldSauda 
         Height          =   360
         Left            =   3840
         TabIndex        =   2
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DNewSauda 
         Height          =   360
         Left            =   3840
         TabIndex        =   4
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DContraAcc 
         Height          =   360
         Left            =   3840
         TabIndex        =   6
         Top             =   1920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contra Acc."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7200
         TabIndex        =   15
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   14
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "New Sauda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Old Sauda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404080&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13530
      Begin VB.Label Label27 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sauda Lot Change"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   13455
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
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
Attribute VB_Name = "frmSaudaLotChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RECgeneral As ADODB.Recordset
Dim RECstanding As ADODB.Recordset
Sub fillOldcombo()
    mysql = "select saudaID, saudaname from saudamast with(nolock) where COMPCODE=" & GCompCode & " and maturity >='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' ORDER BY saudaname"
    Set RECgeneral = Nothing: Set RECgeneral = New ADODB.Recordset: RECgeneral.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RECgeneral.EOF Then
        Set DoldSauda.RowSource = RECgeneral
        DoldSauda.ListField = "saudaname"
        DoldSauda.BoundColumn = "saudaID"
    End If
End Sub
Sub fillnewcombo()
    If DoldSauda.BoundText <> "" Then
        mysql = "select saudaID, saudaname from saudamast with(nolock) where COMPCODE=" & GCompCode & " and maturity >='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'  "
        mysql = mysql & "and ITEMID IN (SELECT Z.ITEMID FROM saudamast z with(nolock) WHERE saudaID = '" & DoldSauda.BoundText & "')   "
        mysql = mysql & "and saudaID NOT IN ('" & DoldSauda.BoundText & "')  "
        mysql = mysql & "ORDER BY saudaname"
        Set RECgeneral = Nothing: Set RECgeneral = New ADODB.Recordset: RECgeneral.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not RECgeneral.EOF Then
            Set DNewSauda.RowSource = RECgeneral
            DNewSauda.ListField = "saudaname"
            DNewSauda.BoundColumn = "saudaID"
        End If
    End If
End Sub
Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub DoldSauda_Validate(Cancel As Boolean)
    Call fillnewcombo
End Sub
Private Sub Form_Load()

    vcDTP2.Value = Date
    
    mysql = "select ac_code, name from AccountD with(nolock) where COMPCODE=" & GCompCode & " and gcode=12  "
    mysql = mysql & "ORDER BY name"
    Set RECgeneral = Nothing: Set RECgeneral = New ADODB.Recordset: RECgeneral.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RECgeneral.EOF Then
        Set DContraAcc.RowSource = RECgeneral
        DContraAcc.ListField = "name"
        DContraAcc.BoundColumn = "ac_code"
    End If
End Sub

Private Sub OkCmd_Click()

On Error GoTo err1
    If DoldSauda.BoundText = "" Then
        MsgBox "Invalid old sauda selection!!!", vbCritical
        DoldSauda.SetFocus
    ElseIf DNewSauda.BoundText = "" Then
        MsgBox "Invalid new sauda selection!!!", vbCritical
        DNewSauda.SetFocus
    ElseIf DContraAcc.BoundText = "" Then
        MsgBox "Invalid contra account selection!!!", vbCritical
        DContraAcc.SetFocus
    ElseIf Not IsNumeric(txtoldrate.text) Then
        MsgBox "Invalid old rate!!!", vbCritical
        txtoldrate.text = ""
        txtoldrate.SetFocus
    ElseIf Not IsNumeric(txtnewrate.text) Then
        MsgBox "Invalid new rate!!!", vbCritical
        txtnewrate.text = ""
        txtnewrate.SetFocus
    Else
        '>>> get standing
        mysql = "select d.party,a.name, d.clqty from inv_d D with(nolock) , accountd a with(nolock) where saudaid = '" & DoldSauda.BoundText & "' and stdate = '" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'  and clqty <> 0   and d.accid=a.accid"
        Set RECstanding = Nothing
        Set RECstanding = New ADODB.Recordset
        RECstanding.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not RECstanding.EOF Then
            GETMAIN.ProgressBar1.Visible = True
            GETMAIN.ProgressBar1.Max = (RECstanding.RecordCount * 2) + 2
            GETMAIN.ProgressBar1.Value = 0
        
            Dim MCount As Long
            Dim LSConSno As String
            Dim NewSCode As String
            Dim NewSitemCode As String
            Dim NewExCode As String
            Dim NewIteId As Integer
            Dim NewExid As Integer
            Dim NewINSTTYPE As String
            Dim NewOPTTYPE As String
            Dim NewSTRIKEPRICE As Double
            
            Dim VContypeNew As String
            Dim LLOT As Double
            Dim OldSCode As String
            Dim OldSitemCode As String
            Dim OldExCode As String
            Dim OldIteId As Integer
            Dim OldExid As Integer
            Dim OldINSTTYPE As String
            Dim OldOPTTYPE As String
            Dim OldSTRIKEPRICE As Double
            
            'OLD
            OldSCode = ""
            OldSitemCode = ""
            OldExCode = ""
            OldIteId = 0
            OldExid = 0
            OldINSTTYPE = ""
            OldOPTTYPE = ""
            OldSTRIKEPRICE = 0
            mysql = "select saudacode, itemcode, excode, itemid, exid from saudamast with(nolock) where COMPCODE=" & GCompCode & " and  saudaID = '" & DoldSauda.BoundText & "' "
            Set RECgeneral = Nothing: Set RECgeneral = New ADODB.Recordset: RECgeneral.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            If Not RECgeneral.EOF Then
                OldSCode = RECgeneral!saudacode
                OldSitemCode = RECgeneral!ITEMCODE
                OldIteId = RECgeneral!itemid
                OldExCode = RECgeneral!excode
                OldExid = RECgeneral!EXID
            End If
            'NEW
            NewSCode = ""
            NewSitemCode = ""
            NewExCode = ""
            NewIteId = 0
            NewExid = 0
            NewINSTTYPE = ""
            NewOPTTYPE = ""
            NewSTRIKEPRICE = 0
            mysql = "select saudacode, itemcode, excode, itemid, exid,INSTTYPE, STRIKEPRICE,OPTTYPE from saudamast with(nolock) where COMPCODE=" & GCompCode & " and  saudaID = '" & DNewSauda.BoundText & "' "
            Set RECgeneral = Nothing: Set RECgeneral = New ADODB.Recordset: RECgeneral.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            If Not RECgeneral.EOF Then
                    NewSCode = RECgeneral!saudacode
                    NewSitemCode = RECgeneral!ITEMCODE
                    NewIteId = RECgeneral!itemid
                    NewExCode = RECgeneral!excode
                    NewExid = RECgeneral!EXID
                    NewINSTTYPE = RECgeneral!INSTTYPE
                    NewSTRIKEPRICE = RECgeneral!STRIKEPRICE
                    NewOPTTYPE = RECgeneral!OPTTYPE
            End If
                                    
Cnn.BeginTrans
CNNERR = True
            GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
            'OLD INSERT
            mysql = "select consno from ctr_m with(nolock) where CONdate ='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' and saudaid = '" & DoldSauda.BoundText & "' "
            Set RECgeneral = Nothing: Set RECgeneral = New ADODB.Recordset: RECgeneral.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            If RECgeneral.EOF Then
                LSConSno = Get_ConSNo(vcDTP2.Value, OldSCode, OldSitemCode, OldExCode, DoldSauda.BoundText, OldIteId, OldExid)
                mysql = "EXEC INSERT_CTR_M " & GCompCode & "," & LSConSno & ",'" & Format(vcDTP2.Value, "YYYY/MM/DD") & "','" & OldSCode & "','" & OldSitemCode & "','C','" & OldExCode & "'," & OldExid & "," & OldIteId & "," & DoldSauda.BoundText & ""
                Cnn.Execute mysql
            Else
                LSConSno = RECgeneral!CONSNO
            End If
            While Not RECstanding.EOF
                VContype = "B"
                If RECstanding!CLQTY > 0 Then
                    VContype = "S"
                End If
                MCount = Get_Max_ConNo(vcDTP2.Value, OldExid)
                MCount = MCount + 1
                LLOT = Get_LotSize(OldIteId, DoldSauda.BoundText, OldExid)
                Call Add_To_Ctr_D(VContype, RECstanding!PARTY, LSConSno, Format(vcDTP2.Value, "YYYY/MM/DD"), MCount, OldSCode, OldSitemCode, RECstanding!PARTY, Abs(RECstanding!CLQTY), txtoldrate.text, DContraAcc.BoundText, "00:00", "", "", MCount, OldExCode, LLOT, 1, "", OldINSTTYPE, OldOPTTYPE, OldSTRIKEPRICE, "0", "N", OldExid, OldIteId, DoldSauda.BoundText)
                GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
                RECstanding.MoveNext
            Wend
            
            'NEW INSERT
            RECstanding.MoveFirst
            
            mysql = "select consno from ctr_m with(nolock) where CONdate = '" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' and saudaid = '" & DNewSauda.BoundText & "' "
            Set RECgeneral = Nothing: Set RECgeneral = New ADODB.Recordset: RECgeneral.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            If RECgeneral.EOF Then
                LSConSno = Get_ConSNo(vcDTP2.Value, NewSCode, NewSitemCode, NewExCode, DNewSauda.BoundText, NewIteId, NewExid)
                mysql = "EXEC INSERT_CTR_M " & GCompCode & "," & LSConSno & ",'" & Format(vcDTP2.Value, "YYYY/MM/DD") & "','" & NewSCode & "','" & NewSitemCode & "','C','" & NewExCode & "'," & NewExid & "," & NewIteId & "," & DNewSauda.BoundText & ""
                Cnn.Execute mysql
            Else
                LSConSno = RECgeneral!CONSNO
            End If
            While Not RECstanding.EOF
                VContypeNew = "B"
                If RECstanding!CLQTY < 0 Then
                    VContypeNew = "S"
                End If
                MCount = Get_Max_ConNo(vcDTP2.Value, NewExid)
                MCount = MCount + 1
                LLOT = Get_LotSize(NewIteId, DNewSauda.BoundText, NewExid)
                
                Call Add_To_Ctr_D(VContypeNew, RECstanding!PARTY, LSConSno, Format(vcDTP2.Value, "YYYY/MM/DD"), MCount, NewSCode, NewSitemCode, RECstanding!PARTY, Abs(RECstanding!CLQTY), txtnewrate.text, DContraAcc.BoundText, "00:00", "", "", MCount, NewExCode, LLOT, 1, "", NewINSTTYPE, NewOPTTYPE, NewSTRIKEPRICE, "0", "N", NewExid, NewIteId, DNewSauda.BoundText)
                GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
                RECstanding.MoveNext
            Wend
Cnn.CommitTrans
CNNERR = False
        End If
    End If
    
err1:
If err.Number <> 0 Then
    MsgBox err.Description, vbInformation
    If CNNERR = True Then
        Cnn.RollbackTrans
        CNNERR = False
    End If
End If
GETMAIN.ProgressBar1.Visible = False
End Sub

Private Sub vcDTP2_Validate(Cancel As Boolean)
    Call fillOldcombo
End Sub
