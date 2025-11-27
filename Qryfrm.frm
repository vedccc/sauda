VERSION 5.00
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form QRYTB 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17925
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Qryfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   17925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   14535
      Begin VB.Label Label27 
         BackColor       =   &H00FF8080&
         Caption         =   "Query on Trial Balance"
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
         TabIndex        =   25
         Top             =   120
         Width           =   14295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   14535
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000000&
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   1100
      End
      Begin VB.TextBox TxtName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         TabIndex        =   22
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox TxtAcCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   112
         Width           =   1100
      End
      Begin VB.CheckBox ChkSettle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With Settlement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   15
         Top             =   112
         Width           =   2175
      End
      Begin VB.CheckBox ChkMargin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With Margin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   14
         Top             =   112
         Width           =   1695
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   105
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   43158.8571990741
      End
      Begin VB.CommandButton CANCEL_CMD 
         BackColor       =   &H80000000&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   112
         Width           =   1100
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   21
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   240
         TabIndex        =   19
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   18
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "As On Date"
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
         Left            =   240
         TabIndex        =   12
         Top             =   172
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
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
      TabIndex        =   3
      Top             =   9480
      Width           =   14535
      Begin VB.TextBox TxtDiff 
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
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   165
         Width           =   1800
      End
      Begin VB.TextBox TxtCredit 
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
         Left            =   8940
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   165
         Width           =   1800
      End
      Begin VB.TextBox TxtDebit 
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   1800
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Trial Total"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   17
         Top             =   165
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Diff"
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
         Left            =   11280
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Credit"
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
         Left            =   7680
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Debit"
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
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid TranGrid 
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13361
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   21
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
   Begin MSDataGridLib.DataGrid TrialGrid 
      Height          =   7575
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   14535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   12
      Height          =   8580
      Left            =   0
      Top             =   1680
      Width           =   14445
   End
End
Attribute VB_Name = "QRYTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RECGRID As ADODB.Recordset:         Dim AccRec As ADODB.Recordset
Dim LTranRow As Integer:                 Dim GRIDROW As Integer
Dim TrialRec As ADODB.Recordset
Dim TranRec As ADODB.Recordset
Public Account_Code As String
Sub Query_Trial()
    
    Dim LAMT As Double:                 Dim LAC_CODE As String:      Dim MDt As Date:            Dim LDebit_Amt As Double:
    Dim LTotDebit As Double:            Dim LTotCredit As Double:    Dim LNetDiff As Double:     Dim LCredit_Amt As Double:
    Dim LCl_Balance As Double:          Dim LOp_Balance As Double:   Dim TRec As ADODB.Recordset
    
    LTotDebit = 0:    LTotCredit = 0:    LNetDiff = 0:    TrialGrid.Visible = False
    If AccRec.RecordCount > 0 Then AccRec.MoveFirst
    MDt = vcDTP1.Value
    Call TrialRecSet
    AccRec.Filter = adFilterNone
    If LenB(TxtAcCode.text) > 0 Then AccRec.Filter = "AC_CODE LIKE '" & TxtAcCode.text & "%'"
    If LenB(TxtName.text) > 0 Then AccRec.Filter = "NAME LIKE '" & TxtName.text & "%'"
    
    Do While Not AccRec.EOF
        LDebit_Amt = 0: LCredit_Amt = 0
        LAC_CODE = AccRec!AC_CODE
        If ChkMargin = 0 And ChkSettle.Value = 1 Then
            LAMT = Net_DrCr(LAC_CODE, MDt)
        Else
            If ChkSettle.Value = 1 Then
                mysql = "SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT*-1 WHEN 'C' THEN  AMOUNT END ) AS TAMT  FROM VCHAMT WHERE  COMPCODE =" & GCompCode & ""
                mysql = mysql & "  AND AC_CODE ='" & LAC_CODE & "' AND VOU_DT <='" & Format(MDt, "YYYY/MM/DD") & "'"
                Set TRec = Nothing
                Set TRec = New ADODB.Recordset
                TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then LAMT = IIf(IsNull(TRec!TAMT), 0, TRec!TAMT)
            Else
                If ChkMargin.Value = 0 Then
                    mysql = "SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT*-1 WHEN 'C' THEN AMOUNT END ) AS TAMT  FROM VCHAMT WHERE  COMPCODE =" & GCompCode & ""
                    mysql = mysql & "  AND AC_CODE ='" & LAC_CODE & "' AND VOU_DT <='" & Format(MDt, "YYYY/MM/DD") & "'AND VOU_TYPE NOT IN ('S','M')"
                    Set TRec = Nothing
                    Set TRec = New ADODB.Recordset
                    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                    If Not TRec.EOF Then LAMT = IIf(IsNull(TRec!TAMT), 0, TRec!TAMT)
                Else
                    mysql = "SELECT SUM(CASE DR_CR WHEN 'D' THEN AMOUNT*-1 WHEN 'C' THEN AMOUNT END ) AS TAMT  FROM VCHAMT WHERE  COMPCODE =" & GCompCode & ""
                    mysql = mysql & "  AND AC_CODE ='" & LAC_CODE & "' AND VOU_DT <='" & Format(MDt, "YYYY/MM/DD") & "' AND VOU_TYPE <>'S' "
                    Set TRec = Nothing
                    Set TRec = New ADODB.Recordset
                    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                    If Not TRec.EOF Then LAMT = IIf(IsNull(TRec!TAMT), 0, TRec!TAMT)
                End If
            
            End If
        End If
        LCl_Balance = AccRec!OP_BAL + LAMT
        TrialRec.AddNew
        TrialRec!AC_CODE = AccRec!AC_CODE
        TrialRec!NAME = AccRec!NAME
        If LCl_Balance > 0 Then
            TrialRec!CREDITAMT = LCl_Balance
            LTotCredit = LTotCredit + LCl_Balance
        ElseIf LCl_Balance < 0 Then
            TrialRec!DEBITAMT = Abs(LCl_Balance)
            LTotDebit = LTotDebit + Abs(LCl_Balance)
        End If
        TrialRec.Update
        AccRec.MoveNext
    Loop
    LNetDiff = LTotCredit - LTotDebit
    If TrialRec.RecordCount > 0 Then TrialRec.MoveFirst
    Set TrialGrid.DataSource = TrialRec
    TrialGrid.ReBind
    TrialGrid.Refresh
    Label5.Caption = "Trial Total"
    TxtDebit.text = Format(LTotDebit, "#,##0.00")
    txtCredit.text = Format(LTotCredit, "#,##0.00")
    If LNetDiff >= 0 Then
        TxtDiff.text = Format(LNetDiff, "#,##0.00") & " Cr"
    Else
        TxtDiff.text = Format(Abs(LNetDiff), "#,##0.00") & " Dr"
    End If
    TrialGrid.Columns(0).Width = 1000
    TrialGrid.Columns(1).Width = 5000
    TrialGrid.Columns(2).Width = 2000
    TrialGrid.Columns(3).Width = 2000
    TrialGrid.Columns(2).Alignment = dbgRight
    TrialGrid.Columns(3).Alignment = dbgRight
    TrialGrid.Columns(2).NumberFormat = "#,##0.00"
    TrialGrid.Columns(3).NumberFormat = "#,##0.00"
    Me.MousePointer = 0
    TrialGrid.Visible = True
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    Me.MousePointer = 0
    Call Get_Selection(10)
    GETMAIN.ProgressBar1.Value = 0
    GETMAIN.ProgressBar1.Visible = False
End Sub

Private Sub Command1_Click()
Query_Trial
End Sub

Private Sub Command2_Click()
Call Query_Trial
End Sub

'Private Sub DataGrid1_DblClick()
'    DataGrid1.Visible = False
'    If TranGrid.Visible = True Then TranGrid.SetFocus
'End Sub
'Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Or KeyCode = 27 Then
'        TranGrid.Visible = True
'        If TranGrid.Visible = True Then TranGrid.SetFocus
'        DataGrid1.Visible = False
'    End If
'End Sub
'Private Sub DataGrid1_Validate(Cancel As Boolean)
'    If DataGrid1.Visible = True Then
'        Cancel = True
'    Else
'        TranGrid.SetFocus
'    End If
'End Sub
Private Sub Form_Load()
    Call Get_Selection(12)
    Account_Code = 0
    vcDTP1.Value = Date
    ChkMargin.Value = 0
    ChkSettle.Value = 1
    
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC_CODE,NAME,OP_BAL, UPPER(NAME) AS AC_NAME FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " ORDER BY NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    TranGrid.Visible = False
    Call Query_Trial
End Sub
Private Sub Form_Paint()
    If GETMAIN.ActiveForm.NAME = Me.NAME Then
        GETMAIN.StatusBar1.Panels(1).text = MFormat
        If TranGrid.Visible = True Then
            TranGrid.SetFocus
        ElseIf TrialGrid.Visible = True Then
            TrialGrid.SetFocus
        End If
    End If
    Call Get_Selection(12)
End Sub
Private Sub CANCEL_CMD_Click()
TxtAcCode.text = vbNullString
TxtName.text = vbNullString
    If TranGrid.Visible Then
        TranGrid.Visible = False
        Call Query_Trial
        TrialGrid.Visible = True
        CANCEL_CMD.Caption = "&Cancel"
    Else
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    GETMAIN.StatusBar1.Panels(1).text = vbNullString
    GETMAIN.Label1.Caption = vbNullString
    Unload Me
End Sub
Sub TRANSACTIONS()
'    Dim LDebit_Amt As Double:       Dim LCredit_Amt As Double:     Dim LOpening_Balance As Double
'    Dim LGridRow As Integer
'    On Error GoTo Error1
'    Call CLEAR_TRANGRID
'    TRAN_GRID.Visible = True
'
'    LGridRow = 0
'
'    LDebit_Amt = 0:         LCredit_Amt = 0:
'    LOpening_Balance = 0:   LGridRow = 0
'
'    If (RIGHT$(MONTH_GRID.text, 2)) = "Cr" Then   'CR
'        LOpening_Balance = CDbl(LEFT$(MONTH_GRID, Len(MONTH_GRID) - 2))
'        LGridRow = GRIDROW + 1
'        TRAN_GRID.Row = GRIDROW
'
'        TRAN_GRID.Col = 0:  TRAN_GRID.text = vbNullString
'        TRAN_GRID.Col = 1:  TRAN_GRID.text = vbNullString
'        TRAN_GRID.Col = 2:  TRAN_GRID.text = vbNullString
'        TRAN_GRID.Col = 4:  TRAN_GRID.text = Format(LOpening_Balance, "0.00")
'        TRAN_GRID.Col = 5:  TRAN_GRID.text = Format(LOpening_Balance, "0.00")
'        TRAN_GRID.Col = 6:  TRAN_GRID.text = "Opening Balance"
'
'        CLOSING_BALANCE = Format(OPENING_BALANCE, "0.00")
'
'    ElseIf (RIGHT$(MONTH_GRID, 2)) = "Dr" Then 'DR
'        LOpening_Balance = CDbl(LEFT$(MONTH_GRID, Len(MONTH_GRID) - 2))
'        GRIDROW = GRIDROW + 1
'        TRAN_GRID.Row = GRIDROW
'
'        TRAN_GRID.Col = 0:  TRAN_GRID.text = vbNullString
'        TRAN_GRID.Col = 1:  TRAN_GRID.text = vbNullString
'        TRAN_GRID.Col = 2:  TRAN_GRID.text = vbNullString
'        TRAN_GRID.Col = 3:  TRAN_GRID.text = Format(LOpening_Balance, "0.00")
'        TRAN_GRID.Col = 4:  TRAN_GRID.text = vbNullString
'        TRAN_GRID.Col = 5:  TRAN_GRID.text = Format(LOpening_Balance, "0.00")
'        TRAN_GRID.Col = 6:  TRAN_GRID.text = "Opening Balance"
'
'        LClosing_Balance = LOpening_Balance * (-1)
'    Else
'        LClosing_Balance = 0
'    End If
'    MYSQL = "SELECT VT.VOU_DT, VT.VOU_NO, VT.VOU_TYPE, VT.DR_CR, VT.AMOUNT, VT.CHEQUE_NO, VT.CHEQUE_DT, ACC.NAME AS AC_NAME, VT.NARRATION FROM VCHAMT AS VT, VOUCHER AS V, ACCOUNTM AS ACC WHERE V.COMPCODE=" & GCompCode  & " AND V.COMPCODE=VT.COMPCODE AND VT.COMPCODE=ACC.COMPCODE AND V.VOU_NO = VT.VOU_NO AND MONTH(VT.VOU_DT) = " & MNTH & " AND YEAR(VT.VOU_DT) = " & YER & "  And VT.AC_CODE =ACC.AC_CODE  AND VT.AC_CODE='" & Account_Code & "' ORDER BY VT.VOU_DT, V.VOU_NO"
'    Set MYRS = Nothing
'    Set MYRS = New ADODB.Recordset
'    MYRS.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'    TRAN_GRID.RowHeight(0) = 350
'
'    While Not MYRS.EOF
'        TRAN_GRID.Rows = TRAN_GRID.Rows + 1
'        TRAN_GRID.RowHeight(TRAN_GRID.Row) = 350
'        If MYRS!DR_CR = "D" Then
'            LDebit_Amt = MYRS!AMOUNT
'            LCredit_Amt = 0
'        ElseIf MYRS!DR_CR = "C" Then
'            LDebit_Amt = 0
'            LCredit_Amt = MYRS!AMOUNT
'        End If
'        LClosing_Balance = LClosing_Balance + LCredit_Amt - LDebit_Amt
'        If LClosing_Balance < 0 Then
'            CL_BAL = CStr(Format(Round(LClosing_Balance * (-1), 2), "0.00")) + " Dr"
'        ElseIf LClosing_Balance > 0 Then
'            CL_BAL = CStr(Format(Round(LClosing_Balance, 2), "0.00")) + " Cr"
'        Else
'            CL_BAL = "0.00 Cr"
'        End If
'
'        If Not IsNull(MYRS!VOU_NO) Then VOUNO = MYRS!VOU_NO
'
'        GRIDROW = GRIDROW + 1
'        TRAN_GRID.Row = GRIDROW
'
'        TRAN_GRID.Col = 0
'        TRAN_GRID.text = IIf(IsNull(MYRS!VOU_DT), vbNullString, MYRS!VOU_DT)
'
'        TRAN_GRID.Col = 1
'        TRAN_GRID.text = VOUNO
'
'        TRAN_GRID.Col = 2
'        TRAN_GRID.text = MYRS!CHEQUE_NO & vbNullString
'
'        TRAN_GRID.Col = 3
'        If Val(LDebit_Amt) <> Val(0) Then TRAN_GRID.text = Format(LDebit_Amt, "0.00")
'
'        TRAN_GRID.Col = 4
'        If Val(LCredit_Amt) <> Val(0) Then TRAN_GRID.text = Format(LCredit_Amt, "0.00")
'
'        TRAN_GRID.Col = 5
'        TRAN_GRID.text = CL_BAL
'
'        TRAN_GRID.Col = 6
'        TRAN_GRID.text = Trim(MYRS!NARRATION)
'
'        TRAN_GRID.Col = 7
'        TRAN_GRID.text = IIf(IsNull(MYRS!VOU_TYPE), "", MYRS!VOU_TYPE)
'
'        TRAN_GRID.Col = 8
'        TRAN_GRID.text = MYRS!VOU_NO
'
'        MYRS.MoveNext
'        TRAN_GRID.Row = TRAN_GRID.Rows - 1
'    Wend
'    TRAN_GRID.RowHeight(TRAN_GRID.Row) = 350
'    TRAN_GRID.Rows = TRAN_GRID.Rows + 1
'    Set MYRS = Nothing
'    CANCEL_CMD.Caption = "&Back"
'    TRAN_GRID.Row = 1
'    TRAN_GRID.Col = 0
'    TRAN_GRID.SetFocus
'Error1: If err.Number <> 0 Then
'       End If
End Sub
Sub ALL_TRANS()
    Dim LDebit_Amt As Double:       Dim LCredit_Amt As Double
    Dim LOp_Bal As Double:          Dim LCl_Bal As Double
    Dim LGridRow As Integer:        Dim TRec As ADODB.Recordset
    Dim LStrCl_Bal As String:       Dim LVouNo As String
    Dim LPDebit As Double:          Dim LPCredit As Double
    Dim LPDiff As Double:
    'Call CLEAR_TRANGRID
    
    TranGrid.Visible = True
    TrialGrid.Col = 1
    Label6.Caption = TrialGrid.text
    TrialGrid.Visible = False
    
    LDebit_Amt = 0:    LCredit_Amt = 0:    LOp_Bal = 0
    LPDebit = 0
    LPCredit = 0
    LPDiff = 0
    LGridRow = 0: LCl_Bal = 0
    'MONTH_GRID.Col = 1
    Dim LAC_CODE As String
    Call TranRecSet
    TrialGrid.Col = 0
    LAC_CODE = TrialGrid.text
    If AccRec.RecordCount > 0 Then
        TranGrid.Visible = False: Me.MousePointer = 11
        AccRec.MoveFirst
        AccRec.Find "AC_CODE='" & LAC_CODE & "'", , adSearchForward
        If AccRec!OP_BAL > 0 Then  'CR
            LOp_Bal = AccRec!OP_BAL
            TranRec.AddNew
            TranRec!VOUDATE = "Opening"
            TranRec!VOUNO = "Opening"
            TranRec!CREDITAMT = LOp_Bal
            LPCredit = LOp_Bal
            TranRec!Balance = Format(LOp_Bal, "#,##0.00)") & " Cr"
            TranRec!NARRATION = " Opening Balance "
            TranRec.Update
            LCl_Bal = LOp_Bal
        ElseIf AccRec!OP_BAL < 0 Then 'DR
            LOp_Bal = Abs(AccRec!OP_BAL)
            LPDebit = LOp_Bal
            TranRec.AddNew
            TranRec!VOUDATE = "Opening"
            TranRec!VOUNO = "Opening"
            TranRec!DEBITAMT = Abs(LOp_Bal)
            TranRec!Balance = Format(Abs(LOp_Bal), "#,##0.00)") & " Dr"
            TranRec!NARRATION = " Opening Balance "
            TranRec.Update
            LCl_Bal = LOp_Bal * (-1)
        End If
        mysql = "SELECT VT.VOU_DT, VT.VOU_NO, VT.VOU_TYPE, VT.DR_CR, VT.AMOUNT, ACC.NAME AS AC_NAME, VT.NARRATION "
        mysql = mysql & " FROM VCHAMT AS VT, VOUCHER AS V, ACCOUNTM AS ACC WHERE ACC.COMPCODE = " & GCompCode & " "
        mysql = mysql & " AND V.COMPCODE=ACC.COMPCODE AND V.VOU_ID = VT.VOU_ID "
        If ChkMargin.Value = 0 And ChkSettle.Value = 1 Then mysql = mysql & " AND V.VOU_TYPE  <>'M' "
        If ChkMargin.Value = 1 And ChkSettle.Value = 0 Then mysql = mysql & " AND V.VOU_TYPE  <>'S' "
        If ChkMargin.Value = 0 And ChkSettle.Value = 0 Then mysql = mysql & " AND V.VOU_TYPE  NOT IN ('S','M') "
        mysql = mysql & " AND V.VOU_DT<='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
        mysql = mysql & " AND VT.AC_CODE =ACC.AC_CODE  AND VT.AC_CODE='" & AccRec!AC_CODE & "'  ORDER BY VT.VOU_DT, V.VOU_NO"
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        While Not TRec.EOF
            TranRec.AddNew
            TranRec!VOUNO = TRec!VOU_NO
            TranRec!VOUDATE = CStr(TRec!VOU_DT)
            TranRec!VOUTYPE = TRec!VOU_TYPE
            TranRec!NARRATION = Left$(TRec!NARRATION, 100)
            If TRec!DR_CR = "D" Then
                LDebit_Amt = TRec!AMOUNT
                LCredit_Amt = 0
                LPDebit = LPDebit + LDebit_Amt
            ElseIf TRec!DR_CR = "C" Then
                LDebit_Amt = 0
                LCredit_Amt = TRec!AMOUNT
                LPCredit = LPCredit + LCredit_Amt
            End If
            LCl_Bal = LCl_Bal + LCredit_Amt - LDebit_Amt
            If LCl_Bal < 0 Then
                LStrCl_Bal = CStr(Format(Abs(LCl_Bal), "#,##0.00")) & "  Dr"
            ElseIf LCl_Bal > 0 Then
                LStrCl_Bal = CStr(Format(LCl_Bal, "#,##0.00")) & " Cr"
            Else
                LStrCl_Bal = "0.00 Cr"
            End If
            If Val(LDebit_Amt) <> 0 Then TranRec!DEBITAMT = LDebit_Amt
            If Val(LCredit_Amt) <> 0 Then TranRec!CREDITAMT = LCredit_Amt
            TranRec!Balance = LStrCl_Bal
            TRec.MoveNext
        Wend
        If TranRec.RecordCount > 0 Then TranRec.MoveFirst
            Set TranGrid.DataSource = TranRec
            TranGrid.ReBind
            TranGrid.Refresh
    
            TranGrid.Columns(0).Width = 1200
            TranGrid.Columns(1).Width = 2200
            TranGrid.Columns(2).Width = 1600
            TranGrid.Columns(3).Width = 1600
            TranGrid.Columns(4).Width = 1800
            TranGrid.Columns(5).Width = 4000
            TranGrid.Columns(2).Alignment = dbgRight
            TranGrid.Columns(3).Alignment = dbgRight
            TranGrid.Columns(4).Alignment = dbgRight
            TranGrid.Columns(2).NumberFormat = "#,##0.00"
            TranGrid.Columns(3).NumberFormat = "#,##0.00"
            Me.MousePointer = 0
            Set TRec = Nothing
            CANCEL_CMD.Caption = "&Back"
            TranGrid.Visible = True: Me.MousePointer = 0
            TranGrid.SetFocus
        End If
        Label5.Caption = "Party Total"
        LPDiff = LPCredit - LPDebit
        
        TxtDebit.text = Format(LPDebit, "#,##0.00")
        txtCredit.text = Format(LPCredit, "#,##0.00")
        If LPDiff >= 0 Then
            TxtDiff.text = Format(LPDiff, "#,##0.00") & " Cr"
        Else
            TxtDiff.text = Format(Abs(LPDiff), "#,##0.00") & " Dr"
        End If
End Sub
Private Sub Text1_Change()

End Sub

Private Sub TranGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        Call TRANGRID_DblClick
    ElseIf KeyCode = 27 Then
        TranGrid.Visible = False
        Call Query_Trial
        TrialGrid.Visible = True
    End If
End Sub
Private Sub TrialGrid_DblClick()
Call ALL_TRANS
End Sub

Private Sub TrialGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call ALL_TRANS
    ElseIf KeyCode = 27 Then
        Unload Me
    End If

End Sub

'Sub VOUCHER_DETAIL()
'    Dim TRec As ADODB.Recordset
'    If ModiPerm = True Then DataGrid1.ToolTipText = "Press F8 to modify the voucher."''''

    'TranGrid.Col = 1
   '
   ' MYSQL = "SELECT UPPER(A.NAME) AS AC_NAME, VT.DR_CR, VT.AMOUNT, VT.NARRATION "
   '' MYSQL = MYSQL & " FROM VCHAMT AS VT, ACCOUNTM AS A WHERE VT.COMPCODE=" & GCompCode & " "
   ' MYSQL = MYSQL & " AND VT.COMPCODE=A.COMPCODE AND VT.VOU_NO='" & TranGrid.text & "' AND A.AC_CODE=VT.AC_CODE ORDER BY VT.vouid"
   ' Set TRec = Nothing
   ' Set TRec = New ADODB.Recordset
   ' TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly'''

    'Call RecSet
        
    'Do While Not TRec.EOF
    '    RECGRID.AddNew
    '    RECGRID!AC_NAME = TRec!AC_NAME
    '    RECGRID!DR_CR = TRec!DR_CR
    '    RECGRID!AMOUNT = Val(TRec!AMOUNT)
    '    RECGRID!NARRATION = TRec!NARRATION & vbNullString
    '    RECGRID.Update
    '
        'TRec.MoveNext
    'Loop
'
'    Set DataGrid1.DataSource = RECGRID
 '   DataGrid1.ReBind
 '   DataGrid1.Refresh
 '   DataGrid1.Visible = True
 '
  '  DataGrid1.ZOrder
  '  DataGrid1.SetFocus
'End Sub
Sub ENTRIES()
'Dim mtype As String
    'MFormate1 = MFormate
    'TranGrid.Row = LTranRow
    'TranGrid.Col = 7
    '''CASH / BANK / JOURNAL VOUCHER
    'If (TranGrid.text = "BV" Or TranGrid.text = "CV" Or TranGrid.text = "JV" Or TranGrid.text = "SP") Then
    '    ''CHECKS SYSTEM LOCKING DATE
    '    'If DateValue((TranGrid.TextMatrix(TranGrid.Row, 0)) < GSysLockDt) Then
    '    '    MsgBox "Sorry you can't modify this voucher. Date lock!", vbExclamation, "Locking Date : " & GSysLockDt & ""
        'End If
   '     'Set MYRS = Nothing
   ' Else
   '     Call VOUCHER_DETAIL
   ' End If
End Sub
Private Sub TRANGRID_DblClick()

    Dim MVou_Type As String
    Dim LVOU_NO As String
    Dim LVou_Dt As Date
    TranGrid.Col = 6
    MVou_Type = TranGrid.text
    TranGrid.Col = 1
    LVOU_NO = TranGrid.text
    
    TranGrid.Col = 0  ' vooudt
    LVou_Dt = CDate(TranGrid.text)
    
    If (GSysLockDt >= LVou_Dt) Then
        If MsgBox("System is locked till date " & GSysLockDt & vbNewLine & "Do you still want to view document?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    
    Call PERMISSIONS("VCHENT")
    If ModiPerm = True Then
    
        VouFrm.Show
        VouFrm.ComboVouType.Visible = True: VouFrm.pr_frame.Visible = True: VouFrm.Label8.Visible = True
        VouFrm.Fb_Press = 2
        
        If (GSysLockDt < LVou_Dt) Then
            Call Get_Selection(2)
        End If
        
        VouFrm.VOUCHER_ACCESS (LVOU_NO)
        VouFrm.Frame6.Enabled = True
        If VouFrm.TXT_NARR.Visible = True Then VouFrm.TXT_NARR.SetFocus
    Else
        MsgBox "Modification Rights not available"
        Exit Sub
    End If
 
End Sub
Private Sub TRAN_GRID_GotFocus()
    TranGrid.LeftCol = 0
    Sendkeys "{HOME}"
    If TranGrid.Visible = True Then TranGrid.SetFocus
End Sub
Private Sub TRAN_GRID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Val(27) Then TrialGrid.Visible = True: TrialGrid.SetFocus: TranGrid.Visible = False
End Sub
Private Sub TRAN_GRID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call TRANGRID_DblClick
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "AC_NAME", adVarChar, 65, adFldIsNullable
    RECGRID.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "AMOUNT", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "NARRATION", adVarChar, 100, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub

Sub TrialRecSet()
    Set TrialRec = Nothing
    Set TrialRec = New ADODB.Recordset
    TrialRec.Fields.Append "Ac_Code", adVarChar, 15, adFldIsNullable
    TrialRec.Fields.Append "Name", adVarChar, 100, adFldIsNullable
    TrialRec.Fields.Append "DebitAmt", adDouble, , adFldIsNullable
    TrialRec.Fields.Append "CreditAmt", adDouble, , adFldIsNullable
    TrialRec.Open , , adOpenKeyset, adLockOptimistic
End Sub

Sub TranRecSet()
    Set TranRec = Nothing
    Set TranRec = New ADODB.Recordset
    TranRec.Fields.Append "VouDate", adVarChar, 10, adFldIsNullable
    TranRec.Fields.Append "VouNo", adVarChar, 20, adFldIsNullable
    TranRec.Fields.Append "DebitAmt", adDouble, , adFldIsNullable
    TranRec.Fields.Append "CreditAmt", adDouble, , adFldIsNullable
    TranRec.Fields.Append "Balance", adVarChar, 25, adFldIsNullable
    TranRec.Fields.Append "Narration", adVarChar, 100, adFldIsNullable
    TranRec.Fields.Append "VouType", adVarChar, 2, adFldIsNullable
    TranRec.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub TrialGrid_KeyPress(KeyAscii As Integer)
 Dim LChar As String
    LChar = UCase(Chr(KeyAscii))
    If KeyAscii = 13 Then
       
    Else
        If Not TrialRec.EOF Then
            If Left$(TrialRec!NAME, 1) = LChar Then
                TrialRec.MoveNext
            ElseIf LChar > Left$(TrialRec!NAME, 1) Then
                Do While Not TrialRec.EOF
                    If Left$(TrialRec!NAME, 1) <> LChar Then
                        TrialRec.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
            Else
                TrialRec.MoveFirst
                Do While Not TrialRec.EOF
                    If Left$(TrialRec!NAME, 1) <> LChar Then
                        TrialRec.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If TrialRec.EOF Then TrialRec.MoveFirst
    End If


End Sub
