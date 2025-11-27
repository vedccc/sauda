VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmOpBAl 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Opening Balance"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
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
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   14295
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   14175
         Begin VB.Label Label7 
            BackColor       =   &H00FF8080&
            Caption         =   "Opening Balances"
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
            TabIndex        =   14
            Top             =   0
            Width           =   14655
         End
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   1680
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   8280
      Width           =   14415
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9840
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7440
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   10
         Top             =   203
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   203
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Diff"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   8
         Top             =   203
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Totals"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   203
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   9840
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   14295
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7335
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   12938
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label5 
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
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmOpBAl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REC_OP As ADODB.Recordset
Dim RECGRID As ADODB.Recordset
Private Sub Command1_Click()
Dim LCode As String
Dim LAmount  As Double
If Not RECGRID.EOF Then
    RECGRID.MoveFirst
    Do While Not RECGRID.EOF
        LCode = RECGRID!AC_CODE
        If RECGRID!DEBIT_AMT <> 0 Then
            LAmount = RECGRID!DEBIT_AMT * -1
        Else
            LAmount = RECGRID!CREDIT_AMT
        End If
        mysql = "UPDATE ACCOUNTM SET OP_BAL =" & LAmount & " WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & LCode & "'"
        Cnn.Execute mysql
        RECGRID.MoveNext
    Loop
    MsgBox "Opening Balances Updated"
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
    

End Sub

Private Sub Form_Load()
Call Get_Selection(12)
mysql = "SELECT AC_CODE,NAME,OP_BAL,G_NAME AS GRPNAME FROM ACCOUNTM AS AC , AC_GROUP AS AG Where AC.COMPCODE =" & GCompCode & "And AC.GCODE = AG.Code ORDER BY NAME"
Set REC_OP = Nothing
Set REC_OP = New ADODB.Recordset
REC_OP.Open mysql, Cnn, adOpenStatic, adLockReadOnly
Call RecSet
Dim TotDebit As Double
Dim TotCredit  As Double
TotDebit = 0
TotCredit = 0
If Not REC_OP.EOF Then
    Do While Not REC_OP.EOF
        RECGRID.AddNew
        RECGRID!AC_CODE = REC_OP!AC_CODE
        RECGRID!AC_NAME = REC_OP!NAME
        RECGRID!GroupName = REC_OP!GRPNAME
        If REC_OP!OP_BAL < 0 Then
            RECGRID!DEBIT_AMT = Abs(REC_OP!OP_BAL)
            RECGRID!CREDIT_AMT = 0
            TotDebit = TotDebit + Abs(REC_OP!OP_BAL)
        Else
            RECGRID!DEBIT_AMT = 0
            RECGRID!CREDIT_AMT = Abs(REC_OP!OP_BAL)
            TotCredit = TotCredit + Abs(REC_OP!OP_BAL)
        End If
        REC_OP.MoveNext
    Loop
    RECGRID.MoveFirst
    Text1.text = Format(TotDebit, "0.00")
    Text2.text = Format(TotCredit, "0.00")
    If Val(Text1) > Val(Text2) Then
        Text3.text = Val(Text1.text) - Val(Text2.text)
        Text3.text = Format(Text3.text, "0.00")
    Else
        Text3.text = Val(Text2.text) - Val(Text1.text)
        Text3.text = Format(Text3.text, "0.00")
    End If
        
End If
Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh:
DataGrid1.Columns(0).Width = 1800
DataGrid1.Columns(1).Width = 3500
DataGrid1.Columns(2).Width = 2000
DataGrid1.Columns(3).Width = 2000
DataGrid1.Columns(4).Width = 3500
DataGrid1.Columns(2).Alignment = dbgRight
DataGrid1.Columns(3).Alignment = dbgRight
DataGrid1.Columns(2).NumberFormat = "0.00"
DataGrid1.Columns(3).NumberFormat = "0.00"
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(1).Locked = True
DataGrid1.Columns(4).Locked = True
End Sub

Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "Debit_Amt", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "Credit_Amt", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "GroupName", adVarChar, 100, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub
Public Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    Dim Rec As ADODB.Recordset
    Dim DrAmt As Double
    Dim CrAmt As Double
    Set Rec = Nothing: Set Rec = New ADODB.Recordset: Set Rec = RECGRID.Clone
    DrAmt = 0: CrAmt = 0
    Do While Not Rec.EOF
        DrAmt = DrAmt + Val(Rec!DEBIT_AMT & "")
        CrAmt = CrAmt + Val(Rec!CREDIT_AMT & "")
        Rec.MoveNext
    Loop
    Text1.text = Format(DrAmt, "0.00")
    Text2.text = Format(CrAmt, "0.00")
    If Val(Text1) > Val(Text2) Then
        Text3.text = Val(Text1.text) - Val(Text2.text)
        Text3.text = Format(Text3.text, "0.00")
    Else
        Text3.text = Val(Text2.text) - Val(Text1.text)
        Text3.text = Format(Text3.text, "0.00")
    End If
End Sub
