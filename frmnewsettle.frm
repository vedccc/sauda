VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmnewsettle 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Exchange Wise Settlement Master"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   9885
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   1695
      Begin VB.Label Label1 
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
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
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
      TabIndex        =   8
      Top             =   0
      Width           =   6255
      Begin VB.Frame Frame4 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   3255
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Settlement Master"
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
         BackColor       =   &H00C0C000&
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   9
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Exchange Wise Settlement Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   5295
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   11033
         _Version        =   393216
         AllowArrows     =   -1  'True
         BackColor       =   14938291
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "SRNO"
            Caption         =   "Set.No."
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
            DataField       =   "SETDATE"
            Caption         =   "Set.Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "SDay"
            Caption         =   "WeekDay"
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
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1800
            EndProperty
         EndProperty
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   1095
      Left            =   10560
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   9960
      Top             =   2520
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   1296
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   14880
      Top             =   4440
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   1085
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   6255
   End
End
Attribute VB_Name = "frmnewsettle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RECGRID As ADODB.Recordset
Dim LastDate As Date
Dim EXHREC As ADODB.Recordset
Private Sub Command1_Click()
    On Error GoTo ERR1
        
    RECGRID.MoveFirst
    LastDate = RECGRID!SETDATE
    RECGRID.MoveNext
    Do While Not RECGRID.EOF
        If IsNull(RECGRID!SETDATE) Then
        Else
            If DateValue(LastDate) > DateValue(RECGRID!SETDATE) Then
               ' MsgBox "Check date.Invalid date found.", vbCritical
               ' DataGrid1.Col = 1: DataGrid1.SetFocus
               ' Exit Sub
            ElseIf DateValue(LastDate) > DateValue(RECGRID!SETDATE) Then
            ElseIf DateValue(LastDate) > DateValue(RECGRID!SETDATE) Then
            ElseIf DateValue(RECGRID!SETDATE) > DateValue(GFinEnd) Then
                MsgBox "Settlement date can not be more than financial year end date", vbCritical
                DataGrid1.Col = 1: DataGrid1.SetFocus
                Exit Sub
            ElseIf DateValue(RECGRID!SETDATE) < DateValue(GFinBegin) Then
                MsgBox "Settlement date can not be less than financial year begin date", vbCritical
                DataGrid1.Col = 1: DataGrid1.SetFocus
                Exit Sub
            Else
                LastDate = RECGRID!SETDATE
            End If
            
        End If
        RECGRID.MoveNext
    Loop
    
    Command1.Enabled = False
    Cnn.BeginTrans
    CNNERR = True
    If GSoft <> 1000 Then
        Cnn.Execute "DELETE FROM SETTLE WHERE COMPCODE =" & GCompCode & ""
    Else
        Cnn.Execute "DELETE FROM SETTLE WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & DataCombo1.BoundText & "'"
    End If
    RECGRID.MoveFirst
    Do While Not RECGRID.EOF
        If Len(Trim(RECGRID!SETDATE & "")) > Val(1) Then
            If GSoft <> 1000 Then
                MYSQL = "INSERT INTO SETTLE(COMPCODE,SETNO, SETDATE) VALUES(" & GCompCode & "," & RECGRID!SrNo & ", '" & Format(RECGRID!SETDATE, "yyyy/MM/dd") & "')"
            Else
                MYSQL = "INSERT INTO SETTLE(COMPCODE,SETNO, SETDATE,EXCODE) VALUES(" & GCompCode & "," & RECGRID!SrNo & ",'" & Format(RECGRID!SETDATE, "yyyy/MM/dd") & "','" & DataCombo1.BoundText & "')"
            End If
            Cnn.Execute MYSQL
        End If
        RECGRID.MoveNext
    Loop

    Cnn.CommitTrans
    CNNERR = False
    Command1.Enabled = True
    RECGRID.MoveFirst
    Exit Sub
ERR1:
    MsgBox err.Description, vbCritical
    Cnn.RollbackTrans
    CNNERR = False: Command1.Enabled = True
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
    'Call LSendKeys_Down
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub DataCombo1_Validate(Cancel As Boolean)
If Trim(DataCombo1.BoundText) <> "" Then
    
    RECGRID.AddNew
    RECGRID!SrNo = RECGRID.AbsolutePosition
    RECGRID.Update

    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    
    Rec.Open "SELECT * FROM SETTLE WHERE COMPCODE =" & GCompCode & " AND EXCODE='" & DataCombo1.BoundText & "' ORDER BY SETNO", Cnn, adOpenForwardOnly, adLockReadOnly

    If Not Rec.EOF Then RECGRID.Delete

    Do While Not Rec.EOF
        RECGRID.AddNew
        RECGRID!SrNo = Rec!SETNO
        RECGRID!SETDATE = Rec!SETDATE
        RECGRID!SDay = WeekdayName(Weekday(Rec!SETDATE))
        RECGRID.Update
        Rec.MoveNext
    Loop

    Set DataGrid1.DataSource = RECGRID
    DataGrid1.ReBind
    DataGrid1.Refresh
Else
    MsgBox "Please Select Exchange", vbInformation
    Cancel = True
    DataCombo1.SetFocus
    Sendkeys "%{DOWN}"
'    Call LSendKeys_Down
End If

End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    Set Rec = RECGRID.Clone
    
    Rec.MoveFirst
    If Not IsNull(RECGRID!SETDATE) Then
        Rec.Find "SETDATE=" & RECGRID!SETDATE & "", , adSearchForward
        If Rec.AbsolutePosition <> RECGRID.AbsolutePosition Then
            MsgBox "Duplicate date found.", vbInformation
            RECGRID.CancelUpdate
            Exit Sub
        End If
        RECGRID!SDay = WeekdayName(Weekday(Rec!SETDATE))
    Else
'        MsgBox "Please enter date", vbCritical
    End If
    Set Rec = Nothing
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And DataGrid1.Columns(1).text <> "" Then
            'If Len(Trim(RECGRID!SETDATE & "")) < Val(1) Then Exit Sub
            DoEvents
            MDt = IIf(IsDate(DataGrid1.Columns(1).text), DateValue(DataGrid1.Columns(1).text), GFinBegin) + 7
            If Not RECGRID.EOF Then
               RECGRID.MoveNext
            End If
            If RECGRID.EOF And MDt <= GFinEnd Then
                RECGRID.AddNew
                RECGRID!SrNo = RECGRID.AbsolutePosition
                RECGRID!SETDATE = MDt
                RECGRID!SDay = WeekdayName(Weekday(MDt))
                RECGRID.Update
            End If
            DataGrid1.Col = 1
    End If
End Sub
Private Sub Form_Load()
    Call Get_Selection(12)
    
    Set EXHREC = Nothing
    Set EXHREC = New ADODB.Recordset
    MYSQL = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE=" & GCompCode & ""
    EXHREC.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    Set DataCombo1.RowSource = EXHREC
    DataCombo1.ListField = "EXNAME"
    DataCombo1.BoundColumn = "EXCODE"
    
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "SETDATE", adDate, , adFldIsNullable
    RECGRID.Fields.Append "SDay", adVarChar, 15, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_Paint()
    GETMAIN.listrec.Enabled = True: GETMAIN.Toolbar1_Buttons(8).Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        Call Get_Selection(10)
        CRViewer1.Visible = False
        Cancel = 1
    Else
        GETMAIN.StatusBar1.Panels(1).text = vbNullString
        Unload Me
    End If
End Sub
Sub LIST_ITEM()
    Screen.MousePointer = 11

    Call Get_Selection(12)

    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "SETDATE", adDate, , adFldIsNullable
    RECGRID.Fields.Append "SDay", adVarChar, 15, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic

    MYSQL = "SELECT SETNO, SETDATE FROM SETTLE WHERE COMPCODE=" & GCompCode & " ORDER BY SETNO"
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    Rec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec.EOF Then
        While Not Rec.EOF
            RECGRID.AddNew
                RECGRID!SrNo = Rec!SETNO
                RECGRID!SETDATE = Rec!SETDATE
                RECGRID!SDay = WeekdayName(Weekday(Rec!SETDATE))
            RECGRID.Update
            Rec.MoveNext
        Wend
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "RptSLst.RPT", 1)
    
        RDCREPO.DiscardSavedData
        RDCREPO.Database.SetDataSource RECGRID
        RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
        CRViewer1.Width = CInt(GETMAIN.Width - 100)
        CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
        CRViewer1.Top = 0: CRViewer1.Left = 0
        
    
        CRViewer1.Visible = True
        CRViewer1.ReportSource = RDCREPO
    
        CRViewer1.ViewReport
    Else
        MsgBox "Record does not exists.", vbInformation
    End If
    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label17.ForeColor = &H0&
End Sub
Private Sub Label17_Click()
    Call LIST_ITEM
End Sub
Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label17.ForeColor = &HC00000
End Sub

