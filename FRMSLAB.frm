VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSLAB 
   Caption         =   "RATE SLAB"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9045
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   6240
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   315
      Left            =   7920
      TabIndex        =   9
      Top             =   5880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   6240
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   7920
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5760
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   5760
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5055
         Left            =   4320
         TabIndex        =   4
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8916
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Slabno"
            Caption         =   "Slab No"
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
            DataField       =   "uptorate"
            Caption         =   "Upto Rate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "brokrate"
            Caption         =   "Brok Rate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "uptostdt"
            Caption         =   "Upto Set Date"
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
            DataField       =   "SNO"
            Caption         =   "SNO"
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
            DataField       =   "NEW"
            Caption         =   "NEW"
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
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   495
            Left            =   2520
            TabIndex        =   2
            Top             =   5760
            Width           =   1335
         End
         Begin MSComctlLib.ListView ItemLst 
            Height          =   5340
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   9419
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
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
               Object.Width           =   5185
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Item List"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Rate Slab List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FRMSLAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LSAUDA As String
Dim parties As String
Dim Items As String
Dim ListIt As ListItem
Dim GridColVal As String
Dim CountRow As Integer
Dim SearchRow As Integer
Public FlagBrok As Boolean
Dim LSettlementDt As String
Public REC As ADODB.Recordset
Dim RECGRID As ADODB.Recordset
Dim TempRec As ADODB.Recordset
Dim flag As Boolean
Public fb_press As Byte
Dim ADDMODE As Boolean

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 And DataGrid1.Col = 4 Then
        LSNO = RECGRID!SLABNO
        RECGRID.MoveNext    ''ADDING NEW ROW
        If RECGRID.EOF Then
            RECGRID.AddNew
            RECGRID.Fields("SLABNO") = LSNO + 1
            RECGRID.Fields("BROKRATE") = 0
            RECGRID.Fields("UPTORATE") = 0
            RECGRID.Fields("UPTOSTDT") = LSettlementDt
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            RECGRID.Update
         End If
        DataGrid1.LeftCol = 0: DataGrid1.Col = 0
ElseIf KeyCode = 13 And DataGrid1.Col = 3 Then
    DataCombo3.Visible = True: DataCombo3.SetFocus
End If
If KeyCode = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
FlagBrok = False
   ' LblEdit.Visible = False: LblCancel.Visible = False: LblSave.Visible = False
    'Last Settlement Date
    LSettlementDt = "": Set REC = Nothing: Set REC = New ADODB.Recordset
    REC.Open "SELECT max(setdate) as MaxSettleDate FROM Settle WHERE COMPCODE = " & MC_CODE & "", CNN, adOpenKeyset, adLockReadOnly
    If Not REC.EOF Then LSettlementDt = REC!MaxSettleDate
    
    DataGrid1.Enabled = False
    Call Get_Selection(13)
    Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
    GeneralRec.Open "SELECT ITEMCODE, ITEMNAME FROM ITEMMAST WHERE COMPCODE=" & MC_CODE & " ORDER BY ITEMNAME", CNN, adOpenKeyset, adLockReadOnly
    If Not GeneralRec.EOF Then
        Set DataCombo2.RowSource = GeneralRec
        DataCombo2.ListField = "ITEMNAME"
        DataCombo2.BoundColumn = "ITEMCODE"
        While Not GeneralRec.EOF
            Set ListIt = ItemLst.ListItems.ADD(, , GeneralRec!ITEMName)
            ListIt.SubItems(1) = GeneralRec!ItemCode
            GeneralRec.MoveNext
        Wend
    End If
    Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
    GeneralRec.Open "SELECT * FROM SETTLE WHERE COMPCODE=" & MC_CODE & " ORDER BY SETDATE", CNN, adOpenKeyset, adLockReadOnly
    If Not GeneralRec.EOF Then
        Set DataCombo3.RowSource = GeneralRec
            DataCombo3.ListField = "SETDATE"
            DataCombo3.BoundColumn = "SETNO"
            Set DataCombo4.RowSource = GeneralRec
            DataCombo4.ListField = "SETDATE"
            DataCombo4.BoundColumn = "SETNO"
        'End If
    End If
    Frame1.Enabled = False
    Call CANCEL_REC

End Sub
Sub CANCEL_REC()
    fb_press = 0
    'For I = 1 To PartyLst.ListItems.Count
    '    PartyLst.ListItems.Item(I).Checked = False
    'Next I
    For I = 1 To ItemLst.ListItems.Count
        ItemLst.ListItems.Item(I).Checked = False
    Next I
    Call RECSET
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Enabled = False
    DataCombo1.Enabled = True:
    'ItemDbComb.Enabled = True:
    'CommAND2.Enabled = True
    Frame1.Enabled = False
    Combo1.Visible = False
    'Combo2.Visible = False
    'Combo3.Visible = False
    'Combo4.Visible = False:
    DataCombo2.Visible = False
    DataCombo3.Visible = False
    ' ItemDbComb.Text = ""
    'DataCombo1.Text = ""
    Frame2.Enabled = False
    'Frame3.Enabled = False
    Call Get_Selection(13)
End Sub

Sub RECSET()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SLABNO", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "UPTORATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "UPTOSTDT", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "NEW", adDouble, , adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub


Sub ADD_NEW()
    Frame1.Enabled = True
    Frame2.Enabled = True
    'Frame3.Enabled = True
    Call Get_Selection(1)
    ItemLst.SetFocus

End Sub
Public Sub CommAND1_Click()
    Items = ""
    For j = 1 To ItemLst.ListItems.Count
        If ItemLst.ListItems(j).Checked = True Then
            Items = Items & "'"
            Items = Items & ItemLst.ListItems(j).SubItems(1)
            Items = Items & "'"
            Exit For
        End If
        If j < ItemLst.ListItems.Count Then
            If ItemLst.ListItems(j + 1).Checked = True And Len(Items) > Val(0) Then
                Items = Items & ", "
            End If
        End If
    Next
    
    CountRow = -1
    Call RECSET
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh

    Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
    MYSQL = "SELECT A.ITEMCODE , B.ITEMNAME, A.SLABNO,A.UPTORATE, A.BROKRATE,  A.UPTOSTDT FROM RATESLAB AS A, ITEMMAST AS B WHERE A.COMPCODE=" & MC_CODE & " AND A.COMPCODE=B.COMPCODE AND A.ITEMCODE=B.ITEMCODE "
    If Items = "" Then
    Else
        MYSQL = MYSQL & "AND B.ITEMCODE = " & Items & ""
    End If
    MYSQL = MYSQL & "ORDER BY A.SLABNO,A.UPTOSTDT "
    GeneralRec.Open MYSQL, CNN, adOpenStatic, adLockReadOnly
    If Not GeneralRec.EOF Then
        DataGrid1.Enabled = True
        Do While Not GeneralRec.EOF
            RECGRID.AddNew
            'RECGRID.Fields("ITEMCODE") = GeneralRec!ItemCode
            'RECGRID.Fields("ITEMNAME") = GeneralRec!ITEMName
            RECGRID.Fields("SLABNO") = GeneralRec!SLABNO
            LSNO = GeneralRec!SLABNO
            RECGRID.Fields("BROKRATE") = GeneralRec!BROKRATE
            RECGRID.Fields("UPTORATE") = GeneralRec!UPTORATE
            If IsNull(GeneralRec!UPTOSTDT) Then
                RECGRID.Fields("UPTOSTDT") = LSettlementDt
            Else
                If GeneralRec!UPTOSTDT = "" Then
                    RECGRID.Fields("UPTOSTDT") = LSettlementDt
                ElseIf DateValue(GeneralRec!UPTOSTDT) = DateValue("01/01/1900") Then
                    RECGRID.Fields("UPTOSTDT") = LSettlementDt
                Else
                    RECGRID.Fields("UPTOSTDT") = GeneralRec!UPTOSTDT
                End If
            End If
            
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            'RECGRID.Fields("PARTY") = GeneralRec!Name & ""
            'RECGRID.Fields("PARTYCODE") = GeneralRec!AC_CODE & ""
            RECGRID.Update
            GeneralRec.MoveNext
        Loop
        RECGRID.AddNew
        RECGRID.Fields("SLABNO") = LSNO + 1
        RECGRID.Fields("BROKRATE") = 0
        RECGRID.Fields("UPTORATE") = 0
        RECGRID.Fields("UPTOSTDT") = LSettlementDt
        CountRow = CountRow + 1
        RECGRID.Fields("New") = CountRow
        RECGRID.Update
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        DataCombo1.Enabled = False
        Command1.Enabled = False
        RECGRID.MoveFirst: DataGrid1.SetFocus
        DataGrid1.LeftCol = 0: DataGrid1.Col = 0
       ' Label3.Visible = True
    Else
       ' Label3.Visible = False
        MsgBox "Record does not exists.", vbExclamation
    End If
End Sub
Sub SAVE_REC()
    On Error GoTo ERR1
    If RECGRID.RecordCount > 0 Then
        Set TempRec = RECGRID.Clone
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If IsDate(TempRec!UPTOSTDT) Then
            Else
                MsgBox "Invalid settlement date ", vbCritical
                Exit Sub
            End If
            TempRec.MoveNext
        Loop
    End If
    CNN.BeginTrans: CNNERR = True
    MYSQL = "DELETE FROM RATESLAB WHERE COMPCODE=" & MC_CODE & " AND ITEMCODE = " & Items & ""
    CNN.Execute MYSQL
    TempRec.MoveFirst
    Do While Not TempRec.EOF
        'If IsNull(TempRec!ItemCode) Then
        'Else
         '   If Trim(TempRec!ItemCode) = "" Then
          '  Else
                If Val(TempRec!UPTORATE) <> 0 Then
                    MYSQL = "INSERT INTO RATESLAB(COMPCODE,SLABNO,ITEMCODE, UPTORATE , BROKRATE, uptostdt) "
                    MYSQL = MYSQL & " VALUES(" & MC_CODE & "," & TempRec!SLABNO & ", " & Items & ", " & Val(TempRec!UPTORATE) & "" & ", " & Val(TempRec!BROKRATE & "") & ",'" & Format(TempRec!UPTOSTDT, "yyyy/MM/dd") & "')"
                    CNN.Execute MYSQL
                End If
           ' End If
        'End If
        TempRec.MoveNext
    Loop
    
    LSAUDA = ""
    Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
    MYSQL = "SELECT distinct SaudaCode FROM Saudamast where COMPCODE = " & MC_CODE & " AND itemcode = (" & Items & ") "
    GeneralRec.Open MYSQL, CNN, adOpenStatic, adLockReadOnly
    While Not GeneralRec.EOF
        If LSAUDA = "" Then
            LSAUDA = "'" & GeneralRec!SAUDACODE & "'"
        Else
            LSAUDA = LSAUDA & ",'" & GeneralRec!SAUDACODE & "'"
        End If
        GeneralRec.MoveNext
    Wend
    
    Dim LBranchClient As String
    LBranchClient = "Client"
    'If Check2.Value = 1 Then LBranchClient = "Branch"
    Call UpdateBrokRateType(False, True, parties, Items, , , , , LBranchClient)
  '  Call UpdateStanding(parties, LSAUDA, "", "")
    CNN.CommitTrans: CNNERR = False
    
    'If Check2.Value = 1 Then
    'Else
        'in case of client
        Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
        MYSQL = "SELECT min(Condate) as MinConDate FROM ctr_M where COMPCODE = " & MC_CODE & " AND ITEMCODE=" & Items & ""
        GeneralRec.Open MYSQL, CNN, adOpenStatic, adLockReadOnly
        MFirstConDate = IIf(IsNull(GeneralRec!MinConDate), MFIN_BEG, GeneralRec!MinConDate)
        CNN.BeginTrans: CNNERR = True
        If BILL_GENERATION(MFirstConDate, MFIN_END, LSAUDA, parties) Then
            CNN.CommitTrans: CNNERR = False
        Else
            CNN.RollbackTrans: CNNERR = False
        End If
    'End If
    
    Call CANCEL_REC
    'Call lblcancel_Click
    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
    If CNNERR = True Then CNN.RollbackTrans: CNNERR = False
End Sub



Private Sub DataCombo3_GotFocus()
    DataGrid1.LeftCol = 3
    DataCombo3.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    DataCombo3.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    DataCombo3.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    DataCombo3.Text = RECGRID!UPTOSTDT & ""
    SendKeys "%{DOWN}"
End Sub

Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        DataCombo3.Visible = False
    ElseIf KeyCode = 13 Then
        'check item wise duplicate settlement ****************
        LGridRow = DataGrid1.Row: LGridCol = DataGrid1.Col: GridColVal = RECGRID!ItemCode: SearchRow = RECGRID!New
        RECGRID.MoveFirst
        Do While Not RECGRID.EOF
            If RECGRID!ItemCode = GridColVal Then
                If SearchRow = RECGRID!New Then
                Else
                If RECGRID!UPTOSTDT = DataCombo3.Text Then
                    MsgBox "Duplicate settlement  date found.", vbCritical:
                    RECGRID.MoveFirst: RECGRID.Find "new =" & SearchRow & "", , adSearchForward
                    DataCombo3.SetFocus: Exit Sub
                End If
                    End If
                Else
                        Exit Do
                End If
                RECGRID.MoveNext
            Loop
            RECGRID.MoveFirst: RECGRID.Find "new =" & SearchRow & "", , adSearchForward
            If KeyCode = 13 Then
                RECGRID!UPTOSTDT = DataCombo3.Text
                LGridRow = SearchRow
            End If
            Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataCombo1.Enabled = False
            RECGRID.MoveFirst: DataGrid1.SetFocus
            DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol + 1: DataCombo3.Visible = False: DataGrid1.SetFocus
    End If
End Sub

