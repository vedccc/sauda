VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmTradeFinder 
   Caption         =   "Trade Finder"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11415
   ScaleWidth      =   22065
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   12255
      Left            =   0
      TabIndex        =   30
      Top             =   1440
      Width           =   18855
      Begin TabDlg.SSTab SSTab1 
         Height          =   12015
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   18615
         _ExtentX        =   32835
         _ExtentY        =   21193
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Detail"
         TabPicture(0)   =   "FrmTradeFinder.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSFdetail"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid MSFdetail 
            Height          =   11415
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Visible         =   0   'False
            Width           =   18495
            _ExtentX        =   32623
            _ExtentY        =   20135
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   30
            BackColor       =   16777215
            ForeColor       =   4194368
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
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18855
      Begin VB.TextBox txttype 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         MaxLength       =   4
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtratet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14880
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtratef 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txttimet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txttimef 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton RadioMN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MCX + NSE Live Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   18840
         TabIndex        =   17
         Top             =   840
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton RadioNSE 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NSE Live Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   18840
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "Back"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16560
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdEx 
         Caption         =   "Select Exchnage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18000
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdUpd 
         Caption         =   "Show Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16560
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Stop Timer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   18000
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton RadioMCX 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MCX Live Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   18840
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   41754.4959722222
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   8760
         TabIndex        =   6
         Top             =   780
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   8760
         TabIndex        =   5
         Top             =   180
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
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
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   41754.4959722222
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   12360
         TabIndex        =   36
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Left            =   12360
         TabIndex        =   35
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   14520
         TabIndex        =   34
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   3360
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   5520
         TabIndex        =   32
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
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
         Left            =   3360
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Finder"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange"
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
         Left            =   7680
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda"
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
         Left            =   7680
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   5520
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
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
         Left            =   22200
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By"
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
         Left            =   21240
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   24000
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   23040
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   24000
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   22920
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmTradeFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LExIDS As String:               Dim I As Integer
Dim LFAccID As Long:                Dim LFSaudaID As String
Dim ExRec As ADODB.Recordset:       Dim BillSumRec As ADODB.Recordset
Dim LAccRec As ADODB.Recordset:     Dim LSaudaRec As ADODB.Recordset
Dim LUSERRec As ADODB.Recordset:    Dim LMemberRec As ADODB.Recordset
Dim LREC As ADODB.Recordset
Dim Gridflag As Boolean
Dim Gridbackflag As Boolean
Dim RadioOpt As String
Private Sub CmdUpd_Click()
'    If txttimef.text = "" Then
'        txttimef.text = "00:00"
'    End If
'    If txttimet.text = "" Then
'        txttimet.text = "00:00"
'    End If
    
    If Not IsNumeric(txtratef.text) And txtratef.text <> "" Then
        MsgBox "Invalid rate value!!!", vbCritical
        txtratef.text = ""
        txtratef.SetFocus
    ElseIf Not IsNumeric(txtratet.text) And txtratet.text <> "" Then
        MsgBox "Invalid rate value!!!", vbCritical
        txtratet.text = ""
        txtratet.SetFocus
    ElseIf (Not Mid(txttimef.text, 3, 1) = ":" Or Len(txttimef.text) <> "5") And txttimef.text <> "" Then
        MsgBox "Invalid time!!!, time format should be 00:00", vbCritical
        txttimef.text = "00:00"
        txttimef.SetFocus
    ElseIf (Not Mid(txttimet.text, 3, 1) = ":" Or Len(txttimet.text) <> "5") And txttimet.text <> "" Then
        MsgBox "Invalid time!!!, time format should be 00:00", vbCritical
        txttimet.text = "00:00"
        txttimet.SetFocus
    Else
        Call Get_TRADE
    End If
    
End Sub
Private Sub Get_TRADE()
    Me.MousePointer = 11
    DoEvents
    Call FLEX_GRID_REFRESH
    SSTab1.Tab = 0
    Me.MousePointer = 0
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    LFAccID = 0
    If LenB(DataCombo1.BoundText) > 0 Then
        LFAccID = Val(DataCombo1.BoundText)
        Set LSaudaRec = Nothing
        Set LSaudaRec = New ADODB.Recordset
        mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST AS A where  COMPCODE=" & GCompCode & " AND A.EXID = '" & LFAccID & "' ORDER BY SAUDACODE  "
        LSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not LSaudaRec.EOF Then
            Set DataCombo2.RowSource = LSaudaRec
            DataCombo2.BoundColumn = "SAUDAID"
            DataCombo2.ListField = "SAUDACODE"
        End If
    Else
        MsgBox "Invalid Exchange, Please select Again"
        Cancel = True
    End If
End Sub
Private Sub DataCombo2_Validate(Cancel As Boolean)
    LFSaudaID = 0
    If LenB(DataCombo2.BoundText) > 0 Then
        LFSaudaID = Val(DataCombo2.BoundText)
    Else
        MsgBox "Invalid Sauda, Please select Again"
        Cancel = True
    End If
End Sub
Private Sub Form_Load()
    vcDTP2.Value = Date
    vcDTP1.Value = Date
    
    LFAccID = 0
    LFSaudaID = 0
    txttype.text = "Buy"
        
    mysql = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE=" & GCompCode & "ORDER BY EXNAME"
    Set ExRec = Nothing
    Set ExRec = New ADODB.Recordset
    ExRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not ExRec.EOF Then
        Set DataCombo1.RowSource = ExRec: DataCombo1.BoundColumn = "EXCODE": DataCombo1.ListField = "EXCODE"
    End If
    If ExRec.RecordCount = 1 Then
        DataCombo1.BoundText = ExRec!excode
        DataCombo1.Enabled = False
        mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST  AS A where COMPCODE=" & GCompCode & " AND excode='" & ExRec!excode & "' ORDER BY SAUDACODE  "
    Else
        mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST  AS A WHERE COMPCODE=" & GCompCode & " ORDER BY SAUDACODE  "
    End If
        
    Set LSaudaRec = Nothing
    Set LSaudaRec = New ADODB.Recordset
    LSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LSaudaRec.EOF Then
        Set DataCombo2.RowSource = LSaudaRec
        DataCombo2.BoundColumn = "SAUDAID"
        DataCombo2.ListField = "SAUDACODE"
    End If
End Sub

Private Sub SetRec()
Set BillSumRec = Nothing
Set BillSumRec = New ADODB.Recordset
    BillSumRec.Fields.Append "PARTY", adVarChar, 15, adFldIsNullable
    BillSumRec.Fields.Append "NAME", adVarChar, 100, adFldIsNullable
    'BillSumRec.Fields.Append "OPBAL", adDouble, , adFldIsNullable
    'BillSumRec.Fields.Append "EXCODE", adVarChar, 10, adFldIsNullable
    BillSumRec.Fields.Append "SAUDA", adVarChar, 50, adFldIsNullable
    'BillSumRec.Fields.Append "EXID", adVarChar, 10, adFldIsNullable
    'BillSumRec.Fields.Append "SAUDAID", adVarChar, 50, adFldIsNullable
    'BillSumRec.Fields.Append "OPQTY", adDouble, , adFldIsNullable
    'BillSumRec.Fields.Append "OPAMT", adDouble, , adFldIsNullable
    BillSumRec.Fields.Append "BUYQTY", adDouble, , adFldIsNullable
    BillSumRec.Fields.Append "BUYAMT", adDouble, , adFldIsNullable
    BillSumRec.Fields.Append "SELLQTY", adDouble, , adFldIsNullable
    BillSumRec.Fields.Append "SELLAMT", adDouble, , adFldIsNullable
    BillSumRec.Fields.Append "CLOSEQTY", adDouble, , adFldIsNullable
    BillSumRec.Fields.Append "CLOSEAMT", adDouble, , adFldIsNullable
    BillSumRec.Fields.Append "BROKAMT", adDouble, , adFldIsNullable

End Sub

Private Sub FLEX_GRID_REFRESH()

    Dim TFRec  As ADODB.Recordset: Dim TOTGross As Double: Dim TOTBROKAMT As Double: Dim TOTBILLAMT As Double
    TOTGross = 0:    TOTBROKAMT = 0:    TOTBILLAMT = 0
    
    DoEvents
    Set TFRec = Nothing: Set TFRec = Nothing: Set TFRec = New ADODB.Recordset
    
    If txtratef.text = "" Then
        txtratef.text = "0"
    End If
    If txtratet.text = "" Then
        txtratet.text = "0"
    End If
    
    mysql = "EXEC TradeFinder '" & GCompCode & "','" & Format(vcDTP1.Value, "YYYY/MM/DD") & "','" & Format(vcDTP2.Value, "YYYY/MM/DD") & "','" & txttimef.text & "','" & txttimet.text & "','" & DataCombo1.BoundText & "','" & DataCombo2.BoundText & "','" & Left(Trim(txttype.text), 1) & "','" & txtratef.text & "','" & txtratet.text & "'  "
    TFRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                       
    MSFdetail.Visible = True
    
    If Not TFRec.EOF Then
    
        Dim Vgridrow As Integer
        Dim Vgridcol As Integer

        '>>> Set grid columns
        DoEvents
        
        MSFdetail.Row = 0
        MSFdetail.Col = 0: MSFdetail.ColWidth(0) = TextWidth("XXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 0) = "Party Code"
        MSFdetail.Col = 1: MSFdetail.ColWidth(1) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 1) = "Party name"
        MSFdetail.Col = 2: MSFdetail.ColWidth(2) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 2) = "Qty."
        MSFdetail.Col = 3: MSFdetail.ColWidth(3) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 3) = "Rate"
        MSFdetail.Col = 4: MSFdetail.ColWidth(4) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 4) = "Time"
        
        MSFdetail.Rows = 1
        Vgridrow = 0
        DoEvents
    
        While Not TFRec.EOF
            DoEvents
            Vgridrow = Vgridrow + 1
                                    
            MSFdetail.Rows = MSFdetail.Rows + 1
            MSFdetail.TextMatrix(Vgridrow, 0) = TFRec!PARTY
            MSFdetail.TextMatrix(Vgridrow, 1) = TFRec!NAME
            MSFdetail.TextMatrix(Vgridrow, 2) = TFRec!QTY
            MSFdetail.TextMatrix(Vgridrow, 3) = TFRec!Rate
            MSFdetail.TextMatrix(Vgridrow, 4) = TFRec!contime
                        
            TFRec.MoveNext
        Wend
    Else
        '>>> Set grid columns
        DoEvents
        
        MSFdetail.Row = 0
        MSFdetail.Col = 0: MSFdetail.ColWidth(0) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 0) = ""
        
        MSFdetail.Rows = 1
        Vgridrow = 0
        DoEvents
    
            DoEvents
            Vgridrow = Vgridrow + 1
                                    
            MSFdetail.Rows = MSFdetail.Rows + 1
            MSFdetail.TextMatrix(Vgridrow, 0) = "No record found!!!"
            
    End If '>>>     If Not TFRec.EOF Then
    
End Sub

Private Sub txttype_KeyPress(KeyAscii As Integer)
    If Val(KeyAscii) >= 48 And KeyAscii <= 122 Then
        If Val(KeyAscii) = 66 Or Val(KeyAscii) = 98 Or Val(KeyAscii) = 83 Or Val(KeyAscii) = 115 Then
        Else
            If Trim(txttype.text) = "Buy" Then
                txttype.text = "Sel"
            Else
                txttype.text = "Buy"
            End If
        End If
    End If
    If KeyAscii = 32 Then
        If Trim(txttype.text) = "Buy" Then
            txttype.text = "Sel"
        Else
            txttype.text = "Buy"
        End If
    End If
    If KeyAscii = 43 Then txttype.text = "Buy"
    If KeyAscii = 45 Then txttype.text = "Sel"
End Sub
