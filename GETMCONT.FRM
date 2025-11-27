VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form GETCont 
   BackColor       =   &H80000003&
   Caption         =   "Contract Entry"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   7680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Height          =   7095
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   11655
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   6480
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   6480
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   6480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   6480
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   360
         Left            =   7200
         TabIndex        =   14
         Top             =   2055
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4575
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8070
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   25
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "SRNO"
            Caption         =   "Sr. No."
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
            DataField       =   "BCODE"
            Caption         =   "Buyer Code"
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
            DataField       =   "BNAME"
            Caption         =   "Buyer Name"
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
            DataField       =   "BQNTY"
            Caption         =   "Qnty."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "BRATE"
            Caption         =   "Rate"
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
         BeginProperty Column05 
            DataField       =   "SCODE"
            Caption         =   "Seller Code"
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
         BeginProperty Column06 
            DataField       =   "SNAME"
            Caption         =   "Seller Name"
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
         BeginProperty Column07 
            DataField       =   "SQNTY"
            Caption         =   "Qnty."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "SRATE"
            Caption         =   "Rate"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Height          =   135
         Left            =   165
         TabIndex        =   12
         Top             =   1560
         Width           =   11295
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "Getmcont.frx":0000
         Left            =   9240
         List            =   "Getmcont.frx":000A
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   495
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   9240
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1095
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   4560
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   487
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   5520
         TabIndex        =   2
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   360
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37860.8625462963
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   2280
         TabIndex        =   4
         Top             =   1095
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   "DataCombo2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shree"
         Height          =   225
         Index           =   7
         Left            =   960
         TabIndex        =   22
         Top             =   6555
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Difference"
         Height          =   225
         Index           =   6
         Left            =   5880
         TabIndex        =   19
         Top             =   6548
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sold"
         Height          =   225
         Index           =   5
         Left            =   8520
         TabIndex        =   18
         Top             =   6548
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bought "
         Height          =   225
         Index           =   0
         Left            =   3240
         TabIndex        =   17
         Top             =   6548
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   225
         Index           =   18
         Left            =   8160
         TabIndex        =   11
         Top             =   555
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Rate"
         Height          =   225
         Index           =   4
         Left            =   8160
         TabIndex        =   10
         Top             =   1163
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item  Name"
         Height          =   225
         Index           =   3
         Left            =   1080
         TabIndex        =   9
         Top             =   1163
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda Code"
         Height          =   225
         Index           =   2
         Left            =   3600
         TabIndex        =   8
         Top             =   555
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract Date"
         Height          =   225
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   555
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   6735
         Left            =   120
         Top             =   240
         Width           =   11415
      End
   End
End
Attribute VB_Name = "GETCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FLOWDIR As Byte
Dim VCHNO As String
Dim GRIDPOS As Byte
Public FB_PRESS As Byte
Dim REC As ADODB.Recordset
Dim RECGRID As ADODB.Recordset
Dim REC_SAUDA As ADODB.Recordset
Dim REC_ACCOUNT As ADODB.Recordset
Sub ADD_REC()
    Frame1.Enabled = True
    Call Get_Selection(1)
    vcDTP1.SetFocus
End Sub
Sub SAVE_REC()
On Error GoTo ERR1
    CNN.BeginTrans
    CNNERR = True

    If FB_PRESS = 1 Then
        VCHNO = VOUCHER_NUMBER("CONT", FIN_YEAR(vcDTP1.Value))

        If Not Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveLast
            CONSNO = Adodc1.Recordset!CONSNO
        End If
''''        Adodc1.Recordset.AddNew
''''        Adodc1.Recordset!CONSNO = Val(CONSNO) + Val(1)
        CONSNO = Val(CONSNO) + Val(1)

    Else
        CONSNO = Adodc1.Recordset!CONSNO
        VCHNO = Adodc1.Recordset!VOU_NO & ""
        If Len(Trim(VCHNO)) < Val(1) Then
            VCHNO = VOUCHER_NUMBER("CONT", FIN_YEAR(vcDTP1.Value))
        End If
    End If

    Call DELETE_VOUCHER(VCHNO)

    If FB_PRESS <> 1 Then
        CNN.Execute "DELETE FROM CTR_D WHERE CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
        CNN.Execute "DELETE FROM CTR_R WHERE CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
        CNN.Execute "DELETE FROM CTR_M WHERE CONSNO=" & Val(Adodc1.Recordset!CONSNO) & ""
    End If

    MYSQL = "INSERT INTO CTR_M(CONSNO, CONDATE, SAUDA, ITEMCODE, CLOSERATE, VOU_NO, PATTAN) VALUES(" & CONSNO & ", DATEVALUE('" & Format(vcDTP1.Value, "dd/MM/yyyy") & "'), '" & Text2.Text & "', '" & DataCombo2.BoundText & "', " & Val(Text3.Text) & ", '" & VCHNO & "', '" & Mid(Combo1.Text, 1, 1) & "')"
    CNN.Execute MYSQL

'    With Adodc1.Recordset
'        !CONDATE = Format(vcDTP1.Value, "dd/MM/yyyy")
'        !SAUDA = Text2.Text
'        !ITEMCODE = DataCombo2.BoundText
'        !CLOSERATE = Val(Text3.Text)
'        !VOU_NO = VCHNO
'        !PATTAN = Mid(Combo1.Text, 1, 1)     ''C FOR CONTRACT, J FOR JHOTA
'    End With

    Dim BOOLAC As String * 1
    Dim RC As ADODB.Recordset

    RECGRID.MoveFirst
    Do While Not RECGRID.EOF
        ''RECORDSET RC IS CHECKING WHETHER THE PARTY IS PERSONNEL OR NOT
        BOOLAC = "N"
        MYSQL = "SELECT PERSONNELAC FROM PITBROK WHERE AC_CODE='" & RECGRID!BCODE & "'"
        Set RC = New ADODB.Recordset
        RC.Open MYSQL, CNN, adOpenKeyset, adLockReadOnly
        If RC!PERSONNELAC = "Y" Then
            MYSQL = "SELECT PERSONNELAC FROM PITBROK WHERE AC_CODE='" & RECGRID!scode & "'"
            Set RC = New ADODB.Recordset
            RC.Open MYSQL, CNN, adOpenKeyset, adLockReadOnly
            If RC!PERSONNELAC = "Y" Then
                BOOLAC = "Y"
            End If
        End If

        If Len(RECGRID!BNAME & "") > Val(0) And Len(RECGRID!SNAME & "") > Val(0) Then   ''WHEN BUYER AND SELLER BOTH ARE THERE
            If RECGRID!BQNTY > Val(0) And RECGRID!BRate > Val(0) Then                   ''QNTY AND RATE REQUIRED
               MYSQL = "INSERT INTO CTR_D (CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT) VALUES(" & Val(CONSNO) & ", DATEVALUE('" & Format(vcDTP1.Value, "dd/MM/yyyy") & "')," & Val(RECGRID.AbsolutePosition) & ",'" & Text2.Text & "', '" & DataCombo2.BoundText & "', '" & RECGRID!BCODE & "', " & Val(RECGRID!BQNTY) & "," & Val(RECGRID!BRate) & ",'B', '" & BOOLAC & "')"
               CNN.Execute MYSQL
            End If

            If RECGRID!SQNTY > Val(0) And RECGRID!SRate > Val(0) Then               ''QNTY AND RATE REQUIRED
               MYSQL = "INSERT INTO CTR_D (CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT) VALUES(" & Val(CONSNO) & ", DATEVALUE('" & Format(vcDTP1.Value, "dd/MM/yyyy") & "')," & Val(RECGRID.AbsolutePosition) & ",'" & Text2.Text & "', '" & DataCombo2.BoundText & "', '" & RECGRID!scode & "', " & Val(RECGRID!SQNTY) & "," & Val(RECGRID!SRate) & ",'S', '" & BOOLAC & "')"
               CNN.Execute MYSQL
             End If
        End If
        RECGRID.MoveNext
    Loop

    If Val(Text3.Text) > Val(0) Then
        MYSQL = "INSERT INTO CTR_R(CONSNO, SAUDA, CONDATE, PATTAN, OPRATE, HGRATE, LOWRATE, CLOSERATE) VALUES(" & CONSNO & ",'" & DataCombo1.BoundText & "', DATEVALUE('" & vcDTP1.Value & "'),'" & Mid(Combo1.Text, 1, 1) & "', 0, 0, 0, " & Val(Text3.Text) & ")"
        CNN.Execute MYSQL
    End If

    MYSQL = "INSERT INTO VOUCHER(VOU_NO, VOU_DT, VOU_TYPE, VOU_PR, BILLNO, BILLDT, USER_NAME, USER_DATE, USER_TIME, USER_ACTION) VALUES('" & VCHNO & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','O','','" & Text2.Text & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & USER_ID & "','" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD')"
    CNN.Execute MYSQL

    If Val(Text6.Text) <> Val(0) Then
        MAMOUNT = Abs(Val(Text6.Text))
        If Val(Text6.Text) < Val(0) Then
            MCR = "C"
            MDR = "D"
            SQL = "DEBIT=DEBIT+"
            SQL1 = "CREDIT=CREDIT+"
        Else
            MCR = "D"
            MDR = "C"
            SQL = "CREDIT=CREDIT+"
            SQL1 = "DEBIT=DEBIT+"
        End If

        MNARATION = "Shree for : " & Text2.Text & ", " & Format(vcDTP1.Value, "dd/MM/yyyy")

        ''SHREE POSTING
        MYSQL = "INSERT INTO VCHAMT(VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES('" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MDR & "'," & GSHREE & "," & Val(MAMOUNT) & ",'" & MNARATION & "')"
        CNN.Execute MYSQL

        MYSQL = "UPDATE ACCOUNTM SET " & SQL & Val(MAMOUNT) & " WHERE AC_CODE='" & GSHREE & "'"
        CNN.Execute MYSQL

        ''TRADING AC POSTING
        MYSQL = "INSERT INTO VCHAMT(VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES('" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MCR & "'," & GTRADING & "," & Val(MAMOUNT) & ",'" & MNARATION & "')"
        CNN.Execute MYSQL

        MYSQL = "UPDATE ACCOUNTM SET " & SQL1 & Val(MAMOUNT) & " WHERE AC_CODE='" & GTRADING & "'"
        CNN.Execute MYSQL
    End If

    CNN.CommitTrans
    CNNERR = False

    Adodc1.Refresh

''''    Adodc1.Recordset.Update

 ''BILL GENERATION
 ''TO FIND FROMDATE
    'MYSQL = "SELECT MAX(SETDATE) FROM SETTLE WHERE SeTDATE < DATEVALUE('" & Format(vcDTP1.Value, "dd/MM/yyyy") & "')"
    'Set REC = Nothing
    'Set REC = New ADODB.Recordset
    'REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    'MFROMDATE = REC.Fields(0) + Val(1)
    'MFROMDATE = Format(vcDTP1.Value, "dd/MM/yyyy")

  ''TO FIND TODATE
    'MYSQL = "SELECT MAX(CONDATE) FROM CTR_M WHERE SAUDA='" & Text2.Text & "'"
    'Set REC = Nothing
    'Set REC = New ADODB.Recordset
    'REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    'If Len(Trim(REC.Fields(0) & "")) > Val(10) Then
    '    MTODATE = REC.Fields(0)
    'Else
    '    MTODATE = MFROMDATE
    'End If

    'CNN.BeginTrans
    'CNNERR = False

    'Call BILL_GENERATION(CDate(MFROMDATE), CDate(MTODATE), Text2.Text)

    'CNN.CommitTrans
    'CNNERR = False

    Call CANCEL_REC

    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
    If CNNERR = True Then
       CNN.RollbackTrans
        CNNERR = False
        Adodc1.Recordset.CancelUpdate
    End If
End Sub
Sub CANCEL_REC()
    Call RECSET

    Set DataGrid1.DataSource = RECGRID
    DataGrid1.Refresh

    Call ClearFormFn(GETCont)
    Call Get_Selection(10)
    Frame1.Enabled = False
End Sub
Sub MODIFY_REC()
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    MYSQL = "SELECT * FROM CTR_M WHERE CONDATE=DATEVALUE('" & Format(vcDTP1.Value, "dd/MM/yyyy") & "') AND SAUDA='" & DataCombo1.BoundText & "' AND PATTAN='" & Mid(Combo1.Text, 1, 1) & "'"
    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    
    If REC.EOF Then
        If FB_PRESS = 2 Then
            MsgBox "Transaction does not exists for the selected creteria?", vbExclamation
            Call CANCEL_REC
        End If
        Exit Sub
    End If
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find "CONSNO=" & Val(REC!CONSNO & "") & "", , adSearchForward

    If FB_PRESS = 1 Then FB_PRESS = 2

    With Adodc1.Recordset
        vcDTP1.Value = !CONDATE
        DataCombo1.BoundText = !SAUDA
        Text2.Text = !SAUDA
        DataCombo2.BoundText = !ITEMCODE
        If !PATTAN = "C" Then
            Combo1.ListIndex = Val(0)
        Else
            Combo1.ListIndex = Val(1)
        End If
        Text3.Text = Format(!CLOSERATE, "0.00")
    End With

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    MYSQL = "SELECT CTR_D.*, A.NAME AS NAME FROM CTR_D, ACCOUNTD AS A WHERE CTR_D.PARTY=A.AC_CODE AND CTR_D.CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " ORDER BY CONNO, CONTYPE"
    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly

    Call RECSET
    RECGRID.Delete

    Do While Not REC.EOF
        RECGRID.AddNew
        RECGRID!SRNO = RECGRID.AbsolutePosition
        RECGRID!BCODE = REC!PARTY & ""
        RECGRID!BNAME = REC!Name
        RECGRID!BQNTY = REC!QTY
        RECGRID!BRate = REC!Rate

        REC.MoveNext
        RECGRID!scode = REC!PARTY & ""
        RECGRID!SNAME = REC!Name
        RECGRID!SQNTY = REC!QTY
        RECGRID!SRate = REC!Rate
        RECGRID.Update

        REC.MoveNext
    Loop
    
    Set DataGrid1.DataSource = RECGRID

    Call DataGrid1_AfterColEdit(0)

    If FB_PRESS = 3 Then
        If MsgBox("You are about to delete this record. Confirm Delete?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            On Error GoTo ERR1
            CNN.BeginTrans
            CNNERR = True

            MYSQL = "DELETE FROM CTR_D WHERE CONSNO=" & Adodc1.Recordset!CONSNO & ""
            CNN.Execute MYSQL

            MYSQL = "DELETE FROM CTR_R WHERE CONSNO=" & Adodc1.Recordset!CONSNO & ""
            CNN.Execute MYSQL

            CNN.CommitTrans
            CNNERR = False

            Call DELETE_VOUCHER(Adodc1.Recordset!VOU_NO & "")

            MYSQL = "DELETE FROM CTR_M WHERE CONSNO=" & Adodc1.Recordset!CONSNO & ""
            CNN.Execute MYSQL

            Adodc1.Refresh

            ''REGENERATING SETTLEMENTS
              MFROMDATE = Format(vcDTP1.Value, "dd/MM/yyyy")

            ''TO FIND TODATE
              MYSQL = "SELECT MAX(CONDATE) FROM CTR_M WHERE SAUDA='" & Text2.Text & "'"
              Set REC = Nothing
              Set REC = New ADODB.Recordset
              REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
              If Len(Trim(REC.Fields(0) & "")) > Val(10) Then
                  MTODATE = REC.Fields(0)
              Else
                  MTODATE = MFROMDATE
              End If

              CNN.BeginTrans
              CNNERR = False

              Call BILL_GENERATION(CDate(MFROMDATE), CDate(MTODATE), Text2.Text)

              CNN.CommitTrans
              CNNERR = False
            ''REGENERATING UPTO HERE

ERR1:
            If Err.Number <> 0 Then
                MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
            End If

            If CNNERR = True Then
                CNN.RollbackTrans
                CNNERR = False
            End If
            Call CANCEL_REC
        End If
    End If
End Sub
Private Sub Combo1_GotFocus()
    If FLOWDIR = 1 Then
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT * FROM CTR_M WHERE SAUDA='" & DataCombo1.BoundText & "'", CNN, adOpenForwardOnly, adLockReadOnly
        If REC.EOF Then SendKeys "%{DOWN}"
    Else
       Text2.SetFocus
    End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open "SELECT * FROM CTR_M WHERE PATTAN='O' AND SAUDA='" & DataCombo1.BoundText & "'", CNN, adOpenForwardOnly, adLockReadOnly
    If Not REC.EOF Then
        If Format(vcDTP1.Value, "dd/MM/yyyy") < REC!CONDATE Then
            MsgBox "Opening for this SAUDA has been already entered on " & Format(REC!CONDATE, "dd/MM/yyyy"), vbExclamation, "Warning"
            vcDTP1.Value = Date
            Cancel = True
            Exit Sub
        End If
    Else
        If Not REC.EOF Then
            If REC!CONDATE > Format(vcDTP1.Value, "dd/MM/yyyy") Then
                MsgBox "Opening for this SAUDA has been already entered on " & Format(REC!CONDATE, "dd/MM/yyyy"), vbExclamation, "Warning"
                vcDTP1.Value = Date
                Exit Sub
            End If
        End If
    End If

    If Combo1.ListIndex = Val(0) Then   ''CONTRACT
        DataGrid1.Columns(7).Locked = True
    Else                                ''OPENING
        DataGrid1.Columns(7).Locked = False
    End If

    Call MODIFY_REC
End Sub
Private Sub DataCombo1_GotFocus()
''''    SendKeys "%{DOWN}"
    If Len(Trim(Text2.Text)) > 0 Then Combo1.SetFocus
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    REC_SAUDA.MoveFirst
    REC_SAUDA.Find "SAUDACODE='" & DataCombo1.BoundText & "'", , adSearchForward
    If REC_SAUDA.EOF Then
        Cancel = True
    Else
        Text2.Text = REC_SAUDA!SAUDACODE
        DataCombo1.BoundText = Text2.Text
        DataCombo2.BoundText = REC_SAUDA!ITEMCODE
        Combo1.SetFocus
    End If
End Sub
Private Sub DataCombo2_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo3_GotFocus()
    If DataGrid1.Col = 1 Then
        DataCombo3.Left = Val(1080)
        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    Else
        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
        DataCombo3.Left = Val(7200)
    End If
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If DataGrid1.Col = 1 Then
            RECGRID!BCODE = DataCombo3.BoundText
            RECGRID!BNAME = DataCombo3.Text
            DataGrid1.Col = 2

        ElseIf DataGrid1.Col = 5 Then
            RECGRID!scode = DataCombo3.BoundText
            RECGRID!SNAME = DataCombo3.Text
            DataGrid1.Col = 6

        End If
        DataGrid1.SetFocus
        DataCombo3.Visible = False

    ElseIf KeyCode = 27 Then
        DataGrid1.SetFocus
        DataCombo3.Visible = False
    End If
End Sub
Private Sub DataCombo3_Validate(Cancel As Boolean)
    If DataCombo3.Visible = True Then Cancel = True
End Sub
Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)
    If ColIndex = Val(1) Or ColIndex = Val(5) Then
        REC_ACCOUNT.MoveFirst
        REC_ACCOUNT.Find "AC_CODE='" & DataGrid1.Text & "'", , adSearchForward
        If Not REC_ACCOUNT.EOF Then
            If ColIndex = Val(1) Then
                DataGrid1.Col = 2
                RECGRID!BCODE = REC_ACCOUNT!AC_CODE
                RECGRID!BNAME = REC_ACCOUNT!Name
            Else
                If Combo1.ListIndex = Val(0) Then
                    DataGrid1.Col = 7
                Else
                    DataGrid1.Col = 6
                End If
                RECGRID!scode = REC_ACCOUNT!AC_CODE
                RECGRID!SNAME = REC_ACCOUNT!Name
            End If
        Else
            MsgBox "Invalid Party code!", vbExclamation, "Warning"
            If ColIndex = Val(1) Then
                RECGRID!BCODE = ""
                RECGRID!BNAME = ""
            Else
                RECGRID!scode = ""
                RECGRID!SNAME = ""
            End If
            DataGrid1.SetFocus
        End If

    ElseIf ColIndex = 3 Or ColIndex = 4 Then
        ''IF CONTRACT THEN ONLY CHANGE OCCURS
        If Combo1.ListIndex = Val(0) Then
            RECGRID!SQNTY = RECGRID!BQNTY
            If Val(RECGRID!SRate & "") = Val(0) Then RECGRID!SRate = RECGRID!BRate
        End If
    End If
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    Set REC = RECGRID.Clone
    BQNTY = 0: SQNTY = 0
    BAMT = 0: SAMT = 0
    Do While Not REC.EOF
        BQNTY = BQNTY + Val(REC!BQNTY & "")
        BAMT = BAMT + (Val(REC!BQNTY & "") * Val(REC!BRate & ""))
        SQNTY = SQNTY + Val(REC!SQNTY & "")
        SAMT = SAMT + (Val(REC!SQNTY & "") * Val(REC!SRate & ""))
        REC.MoveNext
    Loop
    Text1.Text = BQNTY
    Text4.Text = SQNTY
    Text5.Text = Val(Text1.Text) - Val(Text4.Text)
    Text6.Text = Format(Val(BAMT) - Val(SAMT), "0.00")
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And DataGrid1.Col = 8 Then
        BCODE = RECGRID!BCODE
        BNAME = RECGRID!BNAME
        scode = RECGRID!scode
        SNAME = RECGRID!SNAME

        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            If Combo1.ListIndex = Val(1) Then   ''OPENING
                RECGRID!BRate = Val(Text3.Text)
                RECGRID!SRate = Val(Text3.Text)
            Else                        ''LAST INFORMATION CARIES
                RECGRID!BCODE = BCODE
                RECGRID!BNAME = BNAME
                RECGRID!scode = scode
                RECGRID!SNAME = SNAME
            End If
            RECGRID.Update
            RECGRID!SRNO = RECGRID.AbsolutePosition
        End If
        DataGrid1.LeftCol = 0
        DataGrid1.Col = 0

    ElseIf KeyCode = 115 Then   ''F4 KEY
        RNO = InputBox("Enter the row number.", "Ankan")
        If Val(RNO) > Val(0) Then
            RECGRID.MoveFirst
            RECGRID.Find "SRNO=" & Val(RNO) & "", , adSearchForward
            If RECGRID.EOF Then
                MsgBox "Record not found.", vbCritical, "Error"
                RECGRID.MoveFirst
            End If
            DataGrid1.Col = 1
            DataGrid1.SetFocus
        End If

    ElseIf KeyCode = 46 And Shift = 2 Then
        RECGRID.Delete
        If RECGRID.RecordCount = 0 Then
            RECGRID.AddNew
            RECGRID!SRNO = RECGRID.RecordCount
            If Combo1.ListIndex = Val(1) Then
                RECGRID!BRate = Val(Text3.Text)
                RECGRID!SRate = Val(Text3.Text)
            End If

            RECGRID.Update
        End If
        Call DataGrid1_AfterColEdit(0)

    ElseIf KeyCode = 13 And (DataGrid1.Col = 1 Or DataGrid1.Col = 5) Then
        If Len(Trim(DataGrid1.Text)) < 1 Then
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        End If
    ElseIf KeyCode = 27 Then
        KeyCode = 0
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
    Call CANCEL_REC
    
    DataCombo3.Top = Val(2055)
    DataCombo3.Left = Val(1080)
    
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open "SELECT ITEMCODE, ITEMCODE+', '+STRING(6-LEN(ITEMCODE),' ')+ITEMNAME AS ITEMNAME FROM ITEMMAST ORDER BY ITEMCODE", CNN, adOpenKeyset, adLockReadOnly

    Set DataCombo2.RowSource = REC
    DataCombo2.BoundColumn = "ITEMCODE"
    DataCombo2.ListField = "ITEMNAME"

    Set REC_SAUDA = Nothing
    Set REC_SAUDA = New ADODB.Recordset
    REC_SAUDA.Open "SELECT * FROM SAUDAMAST ORDER BY SAUDANAME", CNN, adOpenKeyset, adLockReadOnly
    Set DataCombo1.RowSource = REC_SAUDA
    DataCombo1.BoundColumn = "SAUDACODE"
    DataCombo1.ListField = "SAUDANAME"

    Set REC_ACCOUNT = Nothing
    Set REC_ACCOUNT = New ADODB.Recordset
    REC_ACCOUNT.Open "SELECT AC_CODE, NAME FROM ACCOUNTD ORDER BY AC_CODE", CNN, adOpenKeyset, adLockReadOnly

    Set DataCombo3.RowSource = REC_ACCOUNT
    DataCombo3.BoundColumn = "AC_CODE"
    DataCombo3.ListField = "NAME"

    Adodc1.ConnectionString = CNN
    Adodc1.RecordSource = "SELECT * FROM CTR_M ORDER BY CONSNO"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.ReBind
    DataGrid1.Refresh
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    FLOWDIR = 1
    If Len(Trim(Text2.Text)) < 1 Then
        DataCombo1.SetFocus
    Else
        REC_SAUDA.MoveFirst
        REC_SAUDA.Find "SAUDACODE='" & Text2.Text & "'", , adSearchForward
        If REC_SAUDA.EOF Then
            MsgBox "Invalid sauda code.", vbExclamation, "Error"
            Cancel = True
        Else
            Text2.Text = REC_SAUDA!SAUDACODE
            DataCombo1.BoundText = Text2.Text
            DataCombo2.BoundText = REC_SAUDA!ITEMCODE
'            Combo1.SetFocus
        End If
    End If
End Sub
Private Sub Text3_GotFocus()
    FLOWDIR = 0
    Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Sub RECSET()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "BCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "BNAME", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "BQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "SNAME", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "SQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SRATE", adDouble, , adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    RECGRID.AddNew
    RECGRID.Update
    RECGRID!SRNO = RECGRID.AbsolutePosition
    DataGrid1.Col = 1
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Text3.Text = Format(Text3.Text, "0.00")
    If Combo1.ListIndex = Val(1) Then
        RECGRID!BRate = Val(Text3.Text)
        RECGRID!SRate = Val(Text3.Text)
    End If
End Sub
Sub DELETE_VOUCHER(VOU_NO As String)
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open "SELECT * FROM VCHAMT WHERE VOU_NO='" & VOU_NO & "'", CNN, adOpenForwardOnly, adLockReadOnly
    Do While Not REC.EOF
        If REC!DR_CR = "D" Then
            MYSQL = "UPDATE ACCOUNTM SET DEBIT =DEBIT -" & REC!AMOUNT & " WHERE AC_CODE='" & REC!AC_CODE & "'"
        Else
            MYSQL = "UPDATE ACCOUNTM SET CREDIT=CREDIT-" & REC!AMOUNT & " WHERE AC_CODE='" & REC!AC_CODE & "'"
        End If
        CNN.Execute MYSQL
        
        REC.MoveNext
    Loop

    Set REC = Nothing

    CNN.Execute "DELETE FROM VCHAMT  WHERE VOU_NO='" & VOU_NO & "'"
    CNN.Execute "DELETE FROM VOUCHER WHERE VOU_NO='" & VOU_NO & "'"
End Sub
