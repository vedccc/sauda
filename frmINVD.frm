VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmINVD 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4335
      Left            =   9360
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
      Begin MSComctlLib.ListView ExListView 
         Height          =   4020
         Left            =   90
         TabIndex        =   4
         Top             =   120
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   7091
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388736
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   882
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   12255
      Left            =   0
      TabIndex        =   26
      Top             =   1440
      Width           =   18855
      Begin TabDlg.SSTab SSTab1 
         Height          =   12015
         Left            =   120
         TabIndex        =   27
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
         TabPicture(0)   =   "frmINVD.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSFdetail"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid MSFdetail 
            Height          =   11415
            Left            =   120
            TabIndex        =   8
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
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
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
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   13920
         TabIndex        =   7
         Top             =   120
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
         Left            =   7320
         TabIndex        =   3
         Top             =   360
         Width           =   1815
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
         Left            =   13920
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
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
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
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
         TabIndex        =   9
         Top             =   840
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2415
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   420
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
         Left            =   10200
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
         Left            =   10200
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
         Left            =   3720
         TabIndex        =   1
         Top             =   420
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
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   24
         Top             =   600
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
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   19
         Top             =   720
         Width           =   4575
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
         Left            =   5280
         TabIndex        =   18
         Top             =   480
         Width           =   375
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
         Left            =   9480
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
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
         Left            =   9480
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bill List"
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
         TabIndex        =   15
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmINVD"
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
Private Sub CmdBack_Click()
    Frame5.Visible = False
End Sub
Private Sub CmdUpd_Click()
    Call Get_INVD_Bill
End Sub
Private Sub Get_INVD_Bill()
    Me.MousePointer = 11
    Frame5.Visible = False
    DoEvents
    Call FLEX_GRID_REFRESH
    SSTab1.Tab = 0
    Me.MousePointer = 0
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    LFAccID = 0
    If LenB(DataCombo1.BoundText) > 0 Then
        LFAccID = Val(DataCombo1.BoundText)
        If LFAccID = 0 Then
            MsgBox "Imvalid Party  Please select Again"
            Cancel = True
        Else
            Set LSaudaRec = Nothing
            Set LSaudaRec = New ADODB.Recordset
            
            If LenB(LExIDS) > 1 Then
                mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST AS A where A.EXID IN (" & LExIDS & ")  AND A.SAUDAID in (select distinct SAUDAID from inv_d where ACCID = '" & DataCombo1.BoundText & "') ORDER BY SAUDACODE  "
            Else
                mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST AS A where A.SAUDAID in (select distinct SAUDAID from inv_d where ACCID = '" & DataCombo1.BoundText & "') ORDER BY SAUDACODE  "
            End If
            
            LSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            If Not LSaudaRec.EOF Then
                Set DataCombo2.RowSource = LSaudaRec
                DataCombo2.BoundColumn = "SAUDAID"
                DataCombo2.ListField = "SAUDACODE"
            End If
        End If
        
    End If
End Sub
Private Sub DataCombo2_Validate(Cancel As Boolean)
    LFSaudaID = 0
    If LenB(DataCombo2.BoundText) > 0 Then
        LFSaudaID = Val(DataCombo2.BoundText)
    End If
End Sub

Private Sub Form_Load()
    vcDTP2.Value = Date
    vcDTP1.Value = Date
    LFAccID = 0
    LFSaudaID = 0
    Set ExRec = Nothing
    Set ExRec = New ADODB.Recordset
    
    mysql = "SELECT EXCODE,EXID FROM EXMAST where COMPCODE=" & GCompCode & "  ORDER BY EXCODE "
    ExRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    ExListView.ListItems.Clear
    If Not ExRec.EOF Then
        'If ExRec.RecordCount > 1 Then
            Me.MousePointer = 11:
            ExListView.Visible = False
            Do While Not ExRec.EOF
                ExListView.ListItems.Add , , ExRec!excode
                ExListView.ListItems(ExListView.ListItems.Count).ListSubItems.Add , , ExRec!EXID
                ExRec.MoveNext
            Loop
        'Else
        '    ExListView.Enabled = False:
        'End If
        Me.MousePointer = 0
        ExListView.Visible = True
    End If
    
    Set LAccRec = New ADODB.Recordset
    mysql = " SELECT DISTINCT A.ACCID, A.NAME FROM ACCOUNTD AS A ORDER BY NAME "
    LAccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LAccRec.EOF Then
        Set DataCombo1.RowSource = LAccRec
        DataCombo1.BoundColumn = "ACCID"
        DataCombo1.ListField = "NAME"
    End If
    Set LSaudaRec = Nothing
    Set LSaudaRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST  AS A ORDER BY SAUDACODE  "
    LSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LSaudaRec.EOF Then
        Set DataCombo2.RowSource = LSaudaRec
        DataCombo2.BoundColumn = "SAUDAID"
        DataCombo2.ListField = "SAUDACODE"
    End If
End Sub
Private Sub CmdEx_Click()
    If Frame5.Visible = True Then
        Frame5.Visible = False
        CmdEx.Caption = "Select Exchange"
    Else
        Frame5.Visible = True
        CmdEx.Caption = "Hide List"
    End If
    Call Get_ExIDs
End Sub
Private Sub Get_ExIDs()
    Dim I As Integer
    LExIDS = vbNullString
    For I = 1 To ExListView.ListItems.Count
        If ExListView.ListItems(I).Checked = True Then
            If LenB(LExIDS) <> 0 Then LExIDS = LExIDS & ", "
            LExIDS = LExIDS & ExListView.ListItems(I).ListSubItems(1) & ""
        End If
    Next
    Set LSaudaRec = Nothing
    Set LSaudaRec = New ADODB.Recordset
    If LenB(LExIDS) > 1 Then
        mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST  AS A WHERE A.EXID IN (" & LExIDS & ") ORDER BY SAUDACODE  "
    Else
        mysql = "SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST  AS A ORDER BY SAUDACODE  "
    End If
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
    
'    BillSumRec.Fields.Append "LPARTY", adVarChar, 6, adFldIsNullable
'    BillSumRec.Fields.Append "LNAME", adVarChar, 100, adFldIsNullable
'    BillSumRec.Fields.Append "LAMT", adDouble, , adFldIsNullable
'    BillSumRec.Fields.Append "RPARTY", adVarChar, 6, adFldIsNullable
'    BillSumRec.Fields.Append "RNAME", adVarChar, 100, adFldIsNullable
'    BillSumRec.Fields.Append "RAMT", adDouble, , adFldIsNullable
 '   BillSumRec.Open , , adOpenKeyset, adLockOptimistic
End Sub
'Private Sub Fill_Opening()
'Dim TRec As ADODB.Recordset
'Dim LLastBillDate  As Date
'Get_ExIDs
'Set TRec = Nothing
'Set TRec = New ADODB.Recordset
'MYSQL = "SELECT MAX(STDATE) as MDT FROM INV_D WHERE COMPCODE =" & GCompCode & " AND STDATE <'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
'If LenB(LExIDS) > 1 Then MYSQL = " AND EXID IN (" & LExIDS & ")"
'TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'If Not TRec.EOF Then
'    If Not IsNull(TRec!MDt) Then
'        LLastBillDate = TRec!MDt
'    Else
 '       Exit Sub
 '   End If
 '   Set TRec = Nothing
 ''   Set TRec = New ADODB.Recordset
 '   MYSQL = "SELECT PARTY,EXID,SAUDAID,CLQTY,CLRATE,CALVAL FROM INV_D WHERE COMPCODE =" & GCompCode & " AND CLQTY<>0 "
 '   If LenB(LExIDS) > 1 Then MYSQL = " AND EXID IN (" & LExIDS & ")"
 '   MYSQL = MYSQL & " AND STDATE ='" & Format(LLastBillDate, "YYYY/MM/DD") & "' "
 '   TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
 '   Do While Not TRec.EOF
 '       MYSQL = "EXEC INSERT_BILLSUMOP '" & TRec!PARTY & "'," & TRec!EXID & "," & TRec!SAUDAID & "," & (TRec!CLQTY * -1) & "," & ((TRec!CLQTY * -1) * TRec!ClRate * TRec!Calval) & "," & TRec!Calval & ""
 '       Cnn.Execute MYSQL
 '       TRec.MoveNext
    'Loop
'End If
'End Sub
'Private Sub FLEX_GRID_REFRESH_REPORTFORMAT()
'
'    Dim LSParties As String
'    Dim LSSaudas As String
'    Dim LSExCodes As String
'    Dim LSInst As String
'    Dim Check4 As Integer
'    Dim LSFmlyIDs As String
'    Dim AllParties As Boolean
'    Dim AllFmly As Boolean
'    Dim AllSaudas As Boolean
'    Dim AllExcodes As Boolean
'    Dim AllInst As Boolean
'    Gridflag = True
'    LSParties = ""
'    LSSaudas = ""
'    LSExCodes = ""
'    LSInst = "FUT"
'    Check4 = 0
'    LSFmlyIDs = ""
'    AllParties = True
'    AllFmly = True
'    AllSaudas = True
'    AllExcodes = True
'    AllInst = True
'
'    AllParties = True
'    If DataCombo1.BoundText <> "" Then
'        Set LREC = New ADODB.Recordset
'        mysql = " SELECT A.AC_CODE FROM ACCOUNTD AS A WHERE A.ACCID='" & LFAccID & "'"
'        LREC.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'        If Not LREC.EOF Then
'            LSParties = LREC!AC_CODE
'        End If
'        AllParties = False
'    End If
'    AllSaudas = True
'    If DataCombo2.BoundText <> "" Then
'        LSSaudas = LFSaudaID
'        AllSaudas = False
'    End If
'
'
'    Call Bill_summary_share(LSParties, LSSaudas, LSExCodes, LSInst, vcDTP1.Value, vcDTP2.Value, Check4, LSFmlyIDs, AllParties, AllFmly, AllSaudas, AllExcodes, AllInst, 0, "")
'
'    If Not GlobalRecRpt.EOF Then
'
'        Dim Vgridrow As Integer
'        Dim Vgridcol As Integer
'
'        Dim LDIFFER As Double:        Dim LBrokAmt As Double:        Dim LTranAmt As Double:        Dim LStdAmt As Double:        Dim LBIllAmt As Double
'        Dim RDIFFER As Double:        Dim RBrokAmt As Double:        Dim RTranAmt As Double:        Dim RStdAmt As Double:        Dim RBIllAmt As Double
'        LDIFFER = 0:        LBrokAmt = 0:        LTranAmt = 0:        LStdAmt = 0:        LBIllAmt = 0
'        RDIFFER = 0:        RBrokAmt = 0:        RTranAmt = 0:        RStdAmt = 0:        RBIllAmt = 0
'
'        MSFRPT.Visible = True
'        'DataGrid1.Visible = False
'        '>>> Set grid columns
'        DoEvents
'        MSFRPT.Row = 0
'        MSFRPT.Col = 0: MSFRPT.ColWidth(0) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 0) = "Party": MSFRPT.CellAlignment = 1
'        MSFRPT.Col = 1: MSFRPT.ColWidth(1) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 1) = "Gross MTM"
'        MSFRPT.Col = 2: MSFRPT.ColWidth(2) = TextWidth("XXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 2) = "Brok."
'        MSFRPT.Col = 3: MSFRPT.ColWidth(3) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 3) = "MTM Share"
'        MSFRPT.Col = 4: MSFRPT.ColWidth(4) = TextWidth("XXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 4) = "Brok Share"
'        MSFRPT.Col = 5: MSFRPT.ColWidth(5) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 5) = "Credit Amount"
'        MSFRPT.Col = 6: MSFRPT.ColWidth(6) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 6) = "Party"
'        MSFRPT.Col = 7: MSFRPT.ColWidth(7) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 7) = "Gross MTM"
'        MSFRPT.Col = 8: MSFRPT.ColWidth(8) = TextWidth("XXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 8) = "Brok."
'        MSFRPT.Col = 9: MSFRPT.ColWidth(9) = TextWidth("XXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 9) = "MTM Share"
'        MSFRPT.Col = 10: MSFRPT.ColWidth(10) = TextWidth("XXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 10) = "Brok Share."
'        MSFRPT.Col = 11: MSFRPT.ColWidth(11) = TextWidth("XXXXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 11) = "Debit Amount"
'
'        MSFRPT.Rows = 1
'        Vgridrow = 0
'        DoEvents
'        While Not GlobalRecRpt.EOF
'            DoEvents
'            Vgridrow = Vgridrow + 1
'            MSFRPT.Rows = MSFRPT.Rows + 1
'
'            MSFRPT.Row = Vgridrow
'            MSFRPT.Col = 0
'            MSFRPT.CellAlignment = 1
'
'            MSFRPT.Col = 6
'            MSFRPT.CellAlignment = 1
'
'            MSFRPT.CellAlignment = 1
'            If (InStr(1, GlobalRecRpt!LName, "ZZZZ") > 0) Then
'                MSFRPT.TextMatrix(Vgridrow, 0) = ""
'                MSFRPT.TextMatrix(Vgridrow, 1) = ""
'                MSFRPT.TextMatrix(Vgridrow, 2) = ""
'                MSFRPT.TextMatrix(Vgridrow, 3) = ""
'                MSFRPT.TextMatrix(Vgridrow, 4) = ""
'                MSFRPT.TextMatrix(Vgridrow, 5) = ""
'            Else
'                MSFRPT.TextMatrix(Vgridrow, 0) = GlobalRecRpt!LName
'                MSFRPT.TextMatrix(Vgridrow, 1) = Format(GlobalRecRpt!LDIFFER, "0.00")
'                If GlobalRecRpt!LDIFFER > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 1
'                    MSFRPT.CellBackColor = vbBlue
'                ElseIf GlobalRecRpt!LDIFFER < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 1
'                    MSFRPT.CellBackColor = vbRed
'                End If
'
'                MSFRPT.TextMatrix(Vgridrow, 2) = Format(GlobalRecRpt!LBrokAmt, "0.00")
'                If GlobalRecRpt!LBrokAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 2
'                    MSFRPT.CellBackColor = vbBlue
'                ElseIf GlobalRecRpt!LBrokAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 2
'                    MSFRPT.CellBackColor = vbRed
'                End If
'
'                MSFRPT.TextMatrix(Vgridrow, 3) = Format(GlobalRecRpt!LTranAmt, "0.00")
'                If GlobalRecRpt!LTranAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 3
'                    MSFRPT.CellBackColor = vbBlue
'                ElseIf GlobalRecRpt!LTranAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 3
'                    MSFRPT.CellBackColor = vbRed
'                End If
'
'                MSFRPT.TextMatrix(Vgridrow, 4) = Format(GlobalRecRpt!LStdAmt, "0.00")
'                If GlobalRecRpt!LStdAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 4
'                    MSFRPT.CellBackColor = vbBlue
'                ElseIf GlobalRecRpt!LStdAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 4
'                    MSFRPT.CellBackColor = vbRed
'                End If
'
'                MSFRPT.TextMatrix(Vgridrow, 5) = Format(GlobalRecRpt!LBIllAmt, "0.00")
'                If GlobalRecRpt!LBIllAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 5
'                    MSFRPT.CellBackColor = vbBlue
'                ElseIf GlobalRecRpt!LBIllAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 5
'                    MSFRPT.CellBackColor = vbRed
'                End If
'
'                LDIFFER = LDIFFER + GlobalRecRpt!LDIFFER
'                LBrokAmt = LBrokAmt + GlobalRecRpt!LBrokAmt
'                LTranAmt = LTranAmt + GlobalRecRpt!LTranAmt
'                LStdAmt = LStdAmt + GlobalRecRpt!LStdAmt
'                LBIllAmt = LBIllAmt + GlobalRecRpt!LBIllAmt
'            End If
'
'            If (InStr(1, GlobalRecRpt!RNAME, "ZZZZ") > 0) Then
'                MSFRPT.TextMatrix(Vgridrow, 6) = ""
'                MSFRPT.TextMatrix(Vgridrow, 7) = ""
'                MSFRPT.TextMatrix(Vgridrow, 8) = ""
'                MSFRPT.TextMatrix(Vgridrow, 9) = ""
'                MSFRPT.TextMatrix(Vgridrow, 10) = ""
'                MSFRPT.TextMatrix(Vgridrow, 11) = ""
'            Else
'                MSFRPT.TextMatrix(Vgridrow, 6) = GlobalRecRpt!RNAME
'                MSFRPT.TextMatrix(Vgridrow, 7) = Format(GlobalRecRpt!RDIFFER, "0.00")
'                If GlobalRecRpt!RDIFFER < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 7
'                    MSFRPT.CellBackColor = vbRed
'                ElseIf GlobalRecRpt!RDIFFER > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 7
'                    MSFRPT.CellBackColor = vbBlue
'                End If
'                MSFRPT.TextMatrix(Vgridrow, 8) = Format(GlobalRecRpt!RBrokAmt, "0.00")
'                If GlobalRecRpt!RBrokAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 8
'                    MSFRPT.CellBackColor = vbRed
'                ElseIf GlobalRecRpt!RBrokAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 8
'                    MSFRPT.CellBackColor = vbBlue
'                End If
'                MSFRPT.TextMatrix(Vgridrow, 9) = Format(GlobalRecRpt!RTranAmt, "0.00")
'                If GlobalRecRpt!RTranAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 9
'                    MSFRPT.CellBackColor = vbRed
'                ElseIf GlobalRecRpt!RTranAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 9
'                    MSFRPT.CellBackColor = vbBlue
'                End If
'                MSFRPT.TextMatrix(Vgridrow, 10) = Format(GlobalRecRpt!RStdAmt, "0.00")
'                If GlobalRecRpt!RStdAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 10
'                    MSFRPT.CellBackColor = vbRed
'                ElseIf GlobalRecRpt!RStdAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 10
'                    MSFRPT.CellBackColor = vbBlue
'                End If
'                MSFRPT.TextMatrix(Vgridrow, 11) = Format(GlobalRecRpt!RBIllAmt, "0.00")
'                If GlobalRecRpt!RBIllAmt < 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 11
'                    MSFRPT.CellBackColor = vbRed
'                ElseIf GlobalRecRpt!RBIllAmt > 0 Then
'                    MSFRPT.Row = Vgridrow
'                    MSFRPT.Col = 11
'                    MSFRPT.CellBackColor = vbBlue
'                End If
'                RDIFFER = RDIFFER + GlobalRecRpt!RDIFFER
'                RBrokAmt = RBrokAmt + GlobalRecRpt!RBrokAmt
'                RTranAmt = RTranAmt + GlobalRecRpt!RTranAmt
'                RStdAmt = RStdAmt + GlobalRecRpt!RStdAmt
'                RBIllAmt = RBIllAmt + GlobalRecRpt!RBIllAmt
'            End If
'
'
''            TOTGrossMTM = TOTGrossMTM + Format(GlobalRecRpt!GrossMTM, "0.00")
''            TOTBROKAMT = TOTBROKAMT + Format(GlobalRecRpt!BROKAMT, "0.00")
''            TOTBILLAMT = TOTBILLAMT + Format(GlobalRecRpt!BILLAMT, "0.00")
''
''            TOTM2MSHARE = TOTM2MSHARE + Format(GlobalRecRpt!M2MSHARE, "0.00")
''            TOTBROKSHARE = TOTBROKSHARE + Format(GlobalRecRpt!BROKSHARE, "0.00")
''            TOTNETAMT = TOTNETAMT + Format(GlobalRecRpt!M2MSHARE, "0.00") + Format(GlobalRecRpt!BROKSHARE, "0.00") + Format(GlobalRecRpt!BILLAMT, "0.00")
'
'            DoEvents
'            GlobalRecRpt.MoveNext
'        Wend
'        DoEvents
'        Vgridrow = Vgridrow + 1
'        MSFRPT.Rows = MSFRPT.Rows + 1
'        MSFRPT.Row = Vgridrow
'
'
'
'        MSFRPT.TextMatrix(Vgridrow, 0) = "Total": MSFRPT.Col = 0: MSFRPT.CellBackColor = vbBlue
'        MSFRPT.TextMatrix(Vgridrow, 1) = Format(LDIFFER, "0.00"): MSFRPT.Col = 1:
'        If LDIFFER < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 2) = Format(LBrokAmt, "0.00"): MSFRPT.Col = 2:
'        If LBrokAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 3) = Format(LTranAmt, "0.00"): MSFRPT.Col = 3:
'        If LTranAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'        MSFRPT.TextMatrix(Vgridrow, 4) = Format(LStdAmt, "0.00"): MSFRPT.Col = 4:
'        If LStdAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 5) = Format(LBIllAmt, "0.00"): MSFRPT.Col = 5:
'        If LBIllAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 6) = "Total": MSFRPT.Col = 6: MSFRPT.CellBackColor = vbBlue
'        MSFRPT.TextMatrix(Vgridrow, 7) = Format(RDIFFER, "0.00"): MSFRPT.Col = 7:
'        If RDIFFER < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 8) = Format(RBrokAmt, "0.00"): MSFRPT.Col = 8:
'        If RBrokAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 9) = Format(RTranAmt, "0.00"): MSFRPT.Col = 9:
'        If RTranAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 10) = Format(RStdAmt, "0.00"): MSFRPT.Col = 10:
'        If RStdAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'        MSFRPT.TextMatrix(Vgridrow, 11) = Format(RBIllAmt, "0.00"): MSFRPT.Col = 11:
'        If RBIllAmt < 0 Then
'            MSFRPT.CellBackColor = vbRed
'        Else
'            MSFRPT.CellBackColor = vbBlue
'        End If
'
'
'    End If
'    Gridflag = False
'End Sub
Private Sub FLEX_GRID_REFRESH()

    Dim BillRec  As ADODB.Recordset: Dim TOTGross As Double: Dim TOTBROKAMT As Double: Dim TOTBILLAMT As Double
    TOTGross = 0:    TOTBROKAMT = 0:    TOTBILLAMT = 0
    
    DoEvents
    Set BillRec = Nothing: Set BillRec = Nothing: Set BillRec = New ADODB.Recordset
    mysql = "SELECT STDATE,Sauda,OPQTY,OPRATE,CLQTY,CLRATE,calval AS 'LOT',USDINR,diffamt AS GROSS,BrokAMT,Billamt "
    mysql = mysql & "from inv_d "
    mysql = mysql & "WHERE stdate >= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND stdate <= '" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' "
    If DataCombo1.BoundText <> "" Then
        mysql = mysql & "AND ACCID='" & DataCombo1.BoundText & "' "
    End If
    If DataCombo2.BoundText <> "" Then
        mysql = mysql & "AND SAUDAID='" & DataCombo2.BoundText & "' "
    End If
    mysql = mysql & "ORDER BY stdate"
    BillRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
               
    If Not BillRec.EOF Then
        Dim Vgridrow As Integer
        Dim Vgridcol As Integer
        
        MSFdetail.Visible = True
        '>>> Set grid columns
        DoEvents
        MSFdetail.Row = 0
        MSFdetail.Col = 0: MSFdetail.ColWidth(0) = TextWidth("XXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 0) = "Date"
        MSFdetail.Col = 1: MSFdetail.ColWidth(1) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 1) = "Sauda"
        MSFdetail.Col = 2: MSFdetail.ColWidth(2) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 2) = "OP.Qty."
        MSFdetail.Col = 3: MSFdetail.ColWidth(3) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 3) = "Op.Rate"
        MSFdetail.Col = 4: MSFdetail.ColWidth(4) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 4) = "Cl.Qty."
        MSFdetail.Col = 5: MSFdetail.ColWidth(5) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 5) = "Cl. Rate"
        MSFdetail.Col = 6: MSFdetail.ColWidth(6) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 6) = "Lot"
        MSFdetail.Col = 7: MSFdetail.ColWidth(7) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 7) = "USD/INR"
        MSFdetail.Col = 8: MSFdetail.ColWidth(8) = TextWidth("XXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 8) = "Gross"
        MSFdetail.Col = 9: MSFdetail.ColWidth(9) = TextWidth("XXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 9) = "Brok.Amount"
        MSFdetail.Col = 10: MSFdetail.ColWidth(10) = TextWidth("XXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 10) = "Bill Amount"
        
        MSFdetail.Rows = 1
        Vgridrow = 0
        DoEvents
        While Not BillRec.EOF
            DoEvents
            Vgridrow = Vgridrow + 1
            
            MSFdetail.Rows = MSFdetail.Rows + 1
            MSFdetail.TextMatrix(Vgridrow, 0) = BillRec!STDATE
            MSFdetail.TextMatrix(Vgridrow, 1) = BillRec!Sauda
            MSFdetail.TextMatrix(Vgridrow, 2) = BillRec!OpQty
            MSFdetail.TextMatrix(Vgridrow, 3) = BillRec!OPRATE
            MSFdetail.TextMatrix(Vgridrow, 4) = BillRec!CLQTY
            MSFdetail.TextMatrix(Vgridrow, 5) = BillRec!ClRate
            MSFdetail.TextMatrix(Vgridrow, 6) = BillRec!lot 'Format(BillRec!BuyRATE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 7) = BillRec!USDINR
            MSFdetail.TextMatrix(Vgridrow, 8) = Format(BillRec!GROSS, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 9) = Format(BillRec!BROKAMT, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 10) = Format(BillRec!Billamt, "0.00")
                        
            TOTGross = TOTGross + Format(BillRec!GROSS, "0.00")
            TOTBROKAMT = TOTBROKAMT + Format(BillRec!BROKAMT, "0.00")
            TOTBILLAMT = TOTBILLAMT + Format(BillRec!Billamt, "0.00")
            
            DoEvents
            BillRec.MoveNext
        Wend
         
        If Vgridrow > 0 Then
            Vgridrow = Vgridrow + 1
            MSFdetail.Rows = MSFdetail.Rows + 1
            MSFdetail.TextMatrix(Vgridrow, 0) = ""
            MSFdetail.TextMatrix(Vgridrow, 1) = "Total Amount"
            MSFdetail.Row = Vgridrow: MSFdetail.Col = 1: MSFdetail.CellBackColor = vbYellow
            MSFdetail.TextMatrix(Vgridrow, 2) = ""
            MSFdetail.TextMatrix(Vgridrow, 3) = ""
            MSFdetail.TextMatrix(Vgridrow, 4) = ""
            MSFdetail.TextMatrix(Vgridrow, 5) = ""
            MSFdetail.TextMatrix(Vgridrow, 6) = ""
            MSFdetail.TextMatrix(Vgridrow, 7) = ""
            MSFdetail.TextMatrix(Vgridrow, 8) = Format(TOTGross, "0.00")
            MSFdetail.Row = Vgridrow: MSFdetail.Col = 8: MSFdetail.CellBackColor = vbYellow
            MSFdetail.TextMatrix(Vgridrow, 9) = Format(TOTBROKAMT, "0.00")
            MSFdetail.Row = Vgridrow: MSFdetail.Col = 9: MSFdetail.CellBackColor = vbYellow
            MSFdetail.TextMatrix(Vgridrow, 10) = Format(TOTBILLAMT, "0.00")
            MSFdetail.Row = Vgridrow: MSFdetail.Col = 10: MSFdetail.CellBackColor = vbYellow
        End If
    
    End If '>>>     If Not BillRec.EOF Then
    
End Sub

