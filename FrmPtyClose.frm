VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPtyClose 
   Caption         =   "Party Wise Closing Rate"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   10935
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7245
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   17055
      Begin MSDataListLib.DataCombo DComboSauda 
         Height          =   420
         Left            =   1320
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711680
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6225
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   10980
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   21
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "SAUDACODE"
            Caption         =   "Sauda Code"
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
            DataField       =   "CLOSING"
            Caption         =   "Settle Rate"
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
         BeginProperty Column02 
            DataField       =   "OPEN"
            Caption         =   "Open"
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
         BeginProperty Column03 
            DataField       =   "LOW"
            Caption         =   "Low"
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
         BeginProperty Column04 
            DataField       =   "HIGH"
            Caption         =   "High"
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
         BeginProperty Column05 
            DataField       =   "CLOSE"
            Caption         =   "Close"
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   4004.788
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2294.929
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   17055
         Begin VB.TextBox TxtParty 
            Height          =   420
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   420
            Left            =   840
            TabIndex        =   1
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MonthForeColor  =   0
            Value           =   37905.9259606482
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   390
            Left            =   5520
            TabIndex        =   3
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            Text            =   ""
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
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Party"
            ForeColor       =   &H00000040&
            Height          =   270
            Left            =   2880
            TabIndex        =   10
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            ForeColor       =   &H00000040&
            Height          =   270
            Left            =   240
            TabIndex        =   9
            Top             =   285
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6570
         Left            =   0
         TabIndex        =   7
         Top             =   840
         Width           =   17055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17055
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Party Wise Closing Rate Entry"
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
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   17055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1200
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "temp"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmPtyClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LSSaudas  As String:              Public Fb_Press As Byte:            Dim Rec As ADODB.Recordset
Dim RECGRID As ADODB.Recordset:       Dim RecSauda As ADODB.Recordset:    Dim ExRec As ADODB.Recordset
Dim LSExCode As String:               Dim ldate As Date
Dim LACCID As Integer
Dim LSaudaID As Integer
Dim PartyRec As ADODB.Recordset
Sub Add_Rec()
        Frame1.Enabled = True:        Fb_Press = 1
        Label1.Visible = True:        vcDTP1.Visible = True
        Call Get_Selection(1):
        vcDTP1.Enabled = True:
        vcDTP1.SetFocus:              DataGrid1.LeftCol = 0
End Sub
Sub CANCEL_REC()
On Error GoTo err1
    Fb_Press = 0:                           Call Get_Selection(10)
    vcDTP1.Enabled = True:
    Frame1.Enabled = False:
    TxtParty.text = "": DataCombo1.text = ""
    DComboSauda.Visible = False
    Set PartyRec = Nothing: Set PartyRec = New ADODB.Recordset
    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
    PartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not PartyRec.EOF Then
        Set DataCombo1.RowSource = PartyRec
        DataCombo1.BoundColumn = "AC_CODE"
        DataCombo1.ListField = "NAME"
    End If
    
    Set RecSauda = Nothing:    Set RecSauda = New ADODB.Recordset
    mysql = "SELECT I.EXID,I.ITEMID,I.EXCHANGECODE,S.SAUDAID,S.SAUDACODE,S.SAUDANAME,I.ITEMCODE,S.MATURITY FROM SAUDAMAST AS S,ITEMMAST AS I "
    mysql = mysql & " WHERE S.COMPCODE=" & GCompCode & " AND S.COMPCODE =I.COMPCODE AND S.ITEMID =I.ITEMID "
    mysql = mysql & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'ORDER BY I.ITEMCODE,MATURITY "
    RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not RecSauda.EOF Then
        Set DComboSauda.RowSource = RecSauda
        DComboSauda.ListField = "SAUDANAME"
        DComboSauda.BoundColumn = "SAUDACODE"
    End If
    DComboSauda.Visible = False
    Call RecSet
    Exit Sub
err1:
If err.Number <> 0 Then
    MsgBox err.Description
End If
End Sub
Sub MODIFY_REC()
    Fb_Press = 2
    Frame1.Enabled = True
    Call Get_Selection(2)
    Label1.Visible = True: vcDTP1.Visible = True: vcDTP1.Enabled = True: Frame1.Enabled = True
    If vcDTP1.Enabled Then vcDTP1.SetFocus
End Sub
Sub Save_Rec()
    On Error GoTo err1
    Dim LFromDt As Date
    Dim MSauda As String
    Dim mparty As String
    Dim VSauda As String
    Dim LExIDS As String
    Dim LExID As Integer
    Dim LBillSaudas As String
    Dim trecs As New ADODB.Recordset
    LBillSaudas = ""
    
    If LenB(TxtParty.text) = 0 Then
        MsgBox "Please select party!!!", vbCritical
        TxtParty.SetFocus
        Exit Sub
    End If
    
    VSauda = ""
    Cnn.BeginTrans
    CNNERR = True
    
    RECGRID.MoveFirst
    LFromDt = DateValue(vcDTP1.Value)
    
    mparty = TxtParty.text
    LACCID = Get_AccID(mparty)
    mysql = "DELETE FROM CTR_RP WHERE COMPCODE=" & GCompCode & " AND party ='" & mparty & "' AND CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
    Cnn.Execute mysql
    
    Do While Not RECGRID.EOF
        MSauda = RECGRID!saudacode
        LSaudaID = Get_SaudaID(MSauda)
        
        Set trecs = Nothing:    Set trecs = New ADODB.Recordset
        mysql = "SELECT EXID FROM SAUDAMAST WHERE SAUDAID = '" & LSaudaID & "'"
        trecs.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not trecs.EOF Then
            LExID = trecs!EXID
        End If
        
        If Not InStr(LExIDS, LExID) Then
            If LenB(LExIDS) > 0 Then LExIDS = LExIDS & ","
            LExIDS = LExIDS & str(LExID)
        End If
        
        If VSauda = "" Then
            VSauda = "^" & MSauda & "^"
            If LenB(LBillSaudas) < 1 Then
                LBillSaudas = Trim(str(LSaudaID))
            Else
                If LStr_Exists(LBillSaudas, str(LSaudaID)) = False Then LBillSaudas = LBillSaudas & "," & Trim(str(LSaudaID))
            End If
        Else
            If InStr(1, VSauda, "^" & MSauda & "^") > 0 Then
                If RECGRID!CLOSING > 0 Then
                    Cnn.RollbackTrans: CNNERR = False:
                    MsgBox "Duplicate Sauda " & MSauda & " !!!", vbCritical
                    Exit Sub
                End If
            Else
                VSauda = VSauda & MSauda & "^"
            End If
            If LenB(LBillSaudas) < 1 Then
                LBillSaudas = Trim(str(LSaudaID))
            Else
                If LStr_Exists(LBillSaudas, str(LSaudaID)) = False Then LBillSaudas = LBillSaudas & "," & Trim(str(LSaudaID))
            End If
        End If
                
        If MSauda <> "" Then
            
            If Val(RECGRID!CLOSING) > 0 Then
                mysql = "INSERT INTO CTR_RP (COMPCODE,CONDATE,PARTY,SAUDA,SETTLERATE,ACCID,SAUDAID,[OPEN],[CLOSE],LOW,HIGH)"
                mysql = mysql & " VALUES (" & GCompCode & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "','" & mparty & "','" & MSauda & "'," & Val(RECGRID!CLOSING & "") & ",'" & LACCID & "','" & LSaudaID & "','" & (RECGRID!Open) & "','" & (RECGRID!close) & "','" & (RECGRID!LOW) & "','" & (RECGRID!HIGH) & "')"
                Cnn.Execute mysql
            End If
        End If
        RECGRID.MoveNext
    Loop
    
    Cnn.CommitTrans: CNNERR = False
    Cnn.BeginTrans
    'Call Delete_Inv_D(vbNullString, vbNullString, vbNullString, LFromDt)
    mparty = "'" & mparty & "'"
    
    If BILL_GENERATION(CDate(LFromDt), CDate(GFinEnd), "", mparty, LExIDS) Then
        Cnn.CommitTrans: CNNERR = False
    Else
        Cnn.RollbackTrans: CNNERR = False:
        Call CANCEL_REC: Exit Sub
    End If
    
    'Call Chk_Billing
    Call CANCEL_REC
    err.Number = 0
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number

    If CNNERR = True Then
       Cnn.RollbackTrans: CNNERR = False
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub DataCombo1_Validate(Cancel As Boolean)

    Dim LAcCode As String
    If LenB(DataCombo1.text) = 0 Then
        MsgBox "Party can not be blank"
        Cancel = True
        Sendkeys "%{DOWN}"
    Else
        LAcCode = Get_AccountDCode(DataCombo1.BoundText)
        If LenB(LAcCode) > 1 Then
            TxtParty.text = LAcCode
'            If Frame2.Enabled = False Then
'                Frame10.Enabled = True
'                DtpCondate.Enabled = True
'                DtpCondate.SetFocus
'            ElseIf FrameOpt.Visible = True Then
'                TxtOptType.SetFocus
'            Else
              ' DComboSauda.SetFocus
'            End If
            Fill_Grid
        Else
            DataCombo1.SetFocus
            Cancel = True
            Sendkeys "%{DOWN}"
        End If
    End If

End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    Dim TRec As ADODB.Recordset
    If ColIndex = 0 Then
'        If IsDate(DataGrid1.Columns(0).text) Then    'RECGRID!CONTDATE
'            If DateValue(DataGrid1.Columns(0).text) < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical:  Exit Sub
'            If DateValue(DataGrid1.Columns(0).text) > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
'        Else
'            MsgBox "Please Enter Date.", vbCritical: Exit Sub
'        End If
        
        If LenB(RECGRID!saudacode & "") = 0 Then
            'MsgBox "Please Select Sauda ", vbCritical: Exit Sub
            DComboSauda.Visible = True
            DComboSauda.SetFocus
        End If
'        mysql = "SELECT SAUDA FROM CTR_R WHERE COMPCODE=" & GCompCode & " AND CONDATE='" & Format(RECGRID!CONTDATE, "yyyy/MM/dd") & "' AND SAUDA='" & RECGRID!SAUDACODE & "'"
'        Set TRec = Nothing
'        Set TRec = New ADODB.Recordset
'        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
'        If Not TRec.EOF Then
'            MsgBox "Closing Rate already exists for " & RECGRID!SAUDACODE & ".", vbExclamation
'            DataGrid1.Col = 1
'            RECGRID!SAUDACODE = vbNullString
'            RECGRID!SAUDANAME = vbNullString
'            RECGRID!SAUDAID = 0
'        End If
        ''TO FIND WHETHER THE RECORD EXISTS IN RECORDSET
'        Set TRec = Nothing
'        Set TRec = New ADODB.Recordset
'        Set TRec = RECGRID.Clone
'        TRec.Filter = "CONTDATE='" & RECGRID!CONTDATE & "'"
'        TRec.MoveFirst
'        TRec.Find "SAUDACODE='" & RECGRID!SAUDACODE & "'", , adSearchForward
'        If Not TRec.EOF Then
'            ''THE ORIGINAL RECORD ALWAYS FINDS IN RECORDSET SO FIND NEXT TIME
'            TRec.MoveNext
'            TRec.Find "SAUDACODE='" & RECGRID!SAUDACODE & "'", , adSearchForward
'            If Not TRec.EOF Then
'                MsgBox "Closing Rate already exists for " & RECGRID!SAUDACODE & " in above records.", vbExclamation
'                DataGrid1.Col = 1
'                RECGRID!SAUDACODE = vbNullString
'                RECGRID!SAUDANAME = vbNullString
'                RECGRID!SAUDAID = 0
'            End If
'        End If
'        Set TRec = Nothing
    End If
End Sub

'Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
'If DataGrid1.Col = 3 Then
'    If KeyAscii = 13 Then
'        If Not RECGRID.EOF Then
'            RECGRID.MoveNext
'            If RECGRID.EOF Then RECGRID.MovePrevious
'            DataGrid1.Col = 3
'        End If
'    End If
'End If
'
'End Sub
'
'Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim LGridRow As Integer
'Dim LGridCol As Integer
'Dim GridColVal As Double
'If KeyCode = 118 Then   'F7
'    LGridRow = DataGrid1.Row
'    LGridCol = DataGrid1.Col
'    If DataGrid1.Col = 3 Then 'BROKTYPE
'        GridColVal = RECGRID!CLOSING
'        RECGRID.MoveFirst
'        While Not RECGRID.EOF
'            RECGRID!CLOSING = Val(GridColVal)
'            RECGRID.MoveNext
'        Wend
'    End If
'End If
'
'
'End Sub
Private Sub DataGrid1_GotFocus()
On Error Resume Next
    vcDTP1.Enabled = False
    If DataGrid1.Col = 0 And LenB(RECGRID!saudacode & "") < 1 Then
        DComboSauda.Visible = True
        DComboSauda.SetFocus
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LSDate As Date
    If (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 0 And LenB(RECGRID!saudacode & "") < 1) Then
        DComboSauda.Visible = True
        DComboSauda.SetFocus
'    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = Val(0) And LenB(RECGRID!SAUDACODE & "") > Val(1)) Then
'        Sendkeys "{TAB}"
    ElseIf KeyCode = 13 And DataGrid1.Col = 5 Then
        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            RECGRID!saudacode = vbNullString
            RECGRID!CLOSING = 0
            RECGRID!Open = 0
            RECGRID!LOW = 0
            RECGRID!HIGH = 0
            RECGRID!close = 0
            RECGRID.Update
        End If
        DataGrid1.Col = 0
    End If
End Sub

Private Sub DComboSauda_GotFocus()
    If DataGrid1.Row >= 1 Then
        DComboSauda.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
        Sendkeys "%{DOWN}"
   End If
    
End Sub
Private Sub DComboSauda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If LenB(DComboSauda.BoundText) = 0 Then
            MsgBox "Please Select Sauda.", vbCritical
        Else
            
             RECGRID!saudacode = DComboSauda.BoundText
            
            
'            If LenB(RECGRID!CLOSING & "") = 0 Then
'                RECGRID!CLOSING = 0
'            End If
'            If LenB(RECGRID!Open & "") = 0 Then
'                RECGRID!Open = 0
'            End If
'            If LenB(RECGRID!LOW & "") = 0 Then
'                RECGRID!LOW = 0
'            End If
'            If LenB(RECGRID!HIGH & "") = 0 Then
'                RECGRID!HIGH = 0
'            End If
'            If LenB(RECGRID!close & "") = 0 Then
'                RECGRID!close = 0
'            End If
'            RECGRID!SAUDANAME = DComboSauda.text
            
'            RecSauda.MoveFirst
'            RecSauda.Find "SAUDACODE='" & DComboSauda.BoundText & "'", , adSearchForward
'            If Not RecSauda.EOF Then
'                RECGRID!ITEMCODE = RecSauda!ITEMCODE
'                RECGRID!EXCODE = RecSauda!EXCODE
'                RECGRID!SAUDAID = RecSauda!SAUDAID
'            Else
'                RECGRID!ITEMCODE = vbNullString
'            End If
            Call DataGrid1_AfterColEdit(0)
            DComboSauda.Visible = False
            DataGrid1.Col = 0
            DataGrid1.SetFocus
        End If
    ElseIf KeyCode = 27 Then
        DComboSauda.Visible = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Fb_Press <> 0 Then
        If Me.ActiveControl.NAME = "vcDTP1" Then If KeyCode = 13 Then Sendkeys "{tab}"
        If Me.ActiveControl.NAME = "DComboExchnage" Then
            If KeyCode = 13 Then Sendkeys "{tab}"
        End If
    End If
End Sub
Private Sub Form_Load()
    If Date <= GFinEnd Then
        vcDTP1.Value = Date
    Else
        vcDTP1.Value = DateValue(GFinEnd)
    End If
    
    vcDTP1.MaxDate = GFinEnd:    vcDTP1.MinDate = GFinBegin
    ldate = vcDTP1.Value:
            
    Call CANCEL_REC:             Call Get_Selection(10)
    
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SAUDACODE", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "CLOSING", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "OPEN", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LOW", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "HIGH", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "CLOSE", adDouble, , adFldIsNullable
        
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
    RECGRID.AddNew
    RECGRID.Update
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.ReBind
    DataGrid1.Refresh
End Sub

Private Sub TxtParty_GotFocus()
    TxtParty.SelStart = 0
    TxtParty.SelLength = Len(TxtParty.text)
    If DComboSauda.Visible Then
        DComboSauda.SetFocus
    End If
End Sub

Private Sub TxtParty_Validate(Cancel As Boolean)

Dim LAcCode As String
'If Frame2.Enabled = True Then
    If LenB(TxtParty.text) = 0 Then
        DataCombo1.SetFocus
    Else
        LAcCode = Get_AccountDCode(TxtParty.text)
        If LenB(LAcCode) > 1 Then
            DataCombo1.BoundText = LAcCode
'            DComboTSauda.SetFocus
        Else
            DataCombo1.SetFocus
        End If
    End If
'Else
'    Frame1.Enabled = True
'    Frame10.Enabled = True
'    DtpCondate.Enabled = True
'    DtpCondate.SetFocus
'End If

End Sub

Private Sub vcDTP1_Validate(Cancel As Boolean)
    If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    ldate = vcDTP1.Value
'    Set RecSauda = Nothing:    Set RecSauda = New ADODB.Recordset
'    RecSauda.Open "SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND  MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'", Cnn, adOpenKeyset, adLockReadOnly
'    If Not RecSauda.EOF Then
'        Fill_Grid
'    Else
'        MsgBox "Please Create Sauda, vbInformation"
'        If Frame1.Enabled = True Then
'            vcDTP1.SetFocus
'        End If
'        'Call CANCEL_REC
'    End If
End Sub
Sub DELETE_REC()
    Fb_Press = 3
    Frame1.Enabled = True
    Call Get_Selection(3)
    Label1.Visible = True: vcDTP1.Visible = True: vcDTP1.Enabled = True: Frame1.Enabled = True
    If vcDTP1.Enabled Then vcDTP1.SetFocus
'    If RecSauda.RecordCount > 0 Then
'        Fb_Press = 3
'        Call Get_Selection(3)
'        DataGrid1.Columns(0).Locked = True
'        vcDTP1.Enabled = True
'        Label1.Visible = True: vcDTP1.Visible = True: Frame1.Enabled = True: vcDTP1.SetFocus
'    Else
'        MsgBox "Please Select Sauda.", vbCritical
'        Call CANCEL_REC
'    End If
End Sub

Sub Fill_Grid()
    Dim TRec As ADODB.Recordset
    Dim TRec2 As ADODB.Recordset
    Dim mparty As String
    Call RecSet
    If SYSTEMLOCK(DateValue(vcDTP1.Value)) Then
        MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
    Else
        If Fb_Press = 1 Or Fb_Press = 2 Or Fb_Press = 3 Then
            mysql = "SELECT C.PARTY,C.SAUDA,C.SETTLERATE,C.[OPEN],C.[CLOSE],C.LOW,C.HIGH  FROM CTR_RP AS C,ACCOUNTD AS A "
            mysql = mysql & " WHERE A.COMPCODE=" & GCompCode & " AND  A.ACCID =C.ACCID  AND C.CONDATE ='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' and PARTY='" & TxtParty.text & "' "
            mysql = mysql & " ORDER BY A.NAME"
            Set TRec2 = Nothing
            Set TRec2 = New ADODB.Recordset
            TRec2.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    
            If Not TRec2.EOF Then
                RECGRID.Delete
                Do While Not TRec2.EOF
                    RECGRID.AddNew
                    RECGRID!saudacode = TRec2!Sauda
                    RECGRID!CLOSING = TRec2!SETTLERATE
                    RECGRID!Open = TRec2!Open
                    RECGRID!LOW = TRec2!LOW
                    RECGRID!HIGH = TRec2!HIGH
                    RECGRID!close = TRec2!close
                    RECGRID.Update
                    TRec2.MoveNext
                Loop
                TRec2.MoveFirst: RECGRID.MoveFirst
                DataGrid1.Row = TRec2.RecordCount - 1: DataGrid1.Col = 0: DataGrid1.SetFocus
            ElseIf Fb_Press = 1 Or Fb_Press = 2 Then
                RECGRID.Delete
                RECGRID.AddNew
                RECGRID!saudacode = ""
                RECGRID!CLOSING = 0
                RECGRID!Open = 0
                RECGRID!LOW = 0
                RECGRID!HIGH = 0
                RECGRID!close = 0
                RECGRID.Update
            End If
            If Fb_Press = 3 Then
                If MsgBox("Confirm DELETE?", vbYesNo) = vbYes Then
                    Cnn.BeginTrans
                    mysql = "DELETE FROM CTR_RP WHERE COMPCODE=" & GCompCode & " AND  CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND PARTY='" & TxtParty.text & "' "
                    Cnn.Execute mysql
                    mparty = "'" & TxtParty.text & "'"
                    If BILL_GENERATION(CDate(vcDTP1.Value), CDate(GFinEnd), vbNullString, mparty, vbNullString) Then
                        Cnn.CommitTrans
                        CNNERR = False
                    Else
                        Cnn.RollbackTrans
                        CNNERR = False
                    End If
                End If
                Call CANCEL_REC
            End If
'        ElseIf Fb_Press = 3 Then
'            mysql = "SELECT C.PARTY,A.NAME,C.SAUDA,C.SETTLERATE  FROM CTR_RP AS C,ACCOUNTD AS D "
'            mysql = mysql & " WHERE A.COMPCODE=" & GCompCode & " AND A.COMPCODE =C.COMPCODE AND A.ACCID =C.ACCID  AND C.CONDATE ='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
'            mysql = mysql & " ORDER BY A.NAME"
'            Set TRec = Nothing
'            Set TRec = New ADODB.Recordset
'            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
'            If Not TRec.EOF Then
'                RECGRID.Delete
'                Do While Not TRec.EOF
'                    RECGRID.AddNew
'                    RECGRID!PARTY = TRec!PARTY
'                    RECGRID!PNAME = TRec!NAME
'                    RECGRID!SAUDACODE = TRec!Sauda
'                    RECGRID!CLOSING = TRec!SETTLERATE
'                    RECGRID.Update
'                    TRec.MoveNext
'                Loop
'                If RECGRID.RecordCount > 0 Then RECGRID.MoveFirst
'                DataGrid1.Refresh
'                DataGrid1.Col = 1: DataGrid1.Col = 0: DataGrid1.SetFocus
'            End If
'            If MsgBox("Confirm DELETE?", vbYesNo) = vbYes Then
'                Cnn.BeginTrans
'                mysql = "DELETE FROM CTR_RP WHERE COMPCODE=" & GCompCode & " AND  CONDATE='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
'                Cnn.Execute mysql
'                'Call Delete_Inv_D(vbNullString, vbNullString, vbNullString, vcDTP1.Value)
'                If BILL_GENERATION(CDate(vcDTP1.Value), CDate(GFinEnd), vbNullString, vbNullString, vbNullString) Then
'                    Cnn.CommitTrans
'                    CNNERR = False
'                Else
'                    Cnn.RollbackTrans
'                    CNNERR = False
'                End If
'                'Call Chk_Billing
'            End If
'            Call CANCEL_REC
        End If
    End If
 End Sub
