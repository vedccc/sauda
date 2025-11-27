VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form rmdyfrm 
   BackColor       =   &H80000000&
   Caption         =   "Data View Sheet"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16110
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
   ScaleWidth      =   16110
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
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
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   7455
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   7335
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select All"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   0
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select All"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   5040
            TabIndex        =   15
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exhange List"
            ForeColor       =   &H00400000&
            Height          =   270
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item List"
            ForeColor       =   &H00400000&
            Height          =   270
            Left            =   3960
            TabIndex        =   17
            Top             =   0
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView ItemLst 
         Height          =   2580
         Left            =   4080
         TabIndex        =   0
         Top             =   600
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   8388736
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Name"
            Object.Width           =   5185
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select item specific."
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1835
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Exchange"
            Object.Width           =   6651
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3255
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   6855
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF0000&
         Height          =   390
         ItemData        =   "rmdyfrm.frx":0000
         Left            =   120
         List            =   "rmdyfrm.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2520
         Width           =   4815
      End
      Begin VB.TextBox TxtAdminPass 
         ForeColor       =   &H00FF0000&
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   345
         Left            =   4125
         TabIndex        =   8
         Top             =   840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton ShowCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Show"
         Height          =   465
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton SearchCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         Height          =   465
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   105
         Width           =   975
      End
      Begin VB.CommandButton RefreshCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         Height          =   465
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   105
         Width           =   1095
      End
      Begin VB.CommandButton ClearCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear"
         Height          =   465
         Left            =   1410
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   105
         Width           =   975
      End
      Begin VB.CommandButton ExitCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         Height          =   465
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   105
         Width           =   855
      End
      Begin VB.CommandButton ListCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "List"
         Height          =   465
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   450
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37680
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   450
         Left            =   2100
         TabIndex        =   12
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37680
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Pass"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1890
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483634
      ForeColor       =   4194368
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
         MarqueeStyle    =   5
         ScrollBars      =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   14295
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FF8080&
      Caption         =   "Data View Sheet"
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
      TabIndex        =   23
      Top             =   0
      Width           =   14415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   9300
      Left            =   120
      Top             =   720
      Width           =   14250
   End
End
Attribute VB_Name = "rmdyfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RECGRID As ADODB.Recordset:     Dim DatabasePath As String:     Dim MainRec As ADODB.Recordset:
Dim SEARCHTXT As String:            Dim SORTORDER As String
Dim listrec As ADODB.Recordset
Private Sub Check4_Click()
Dim I As Integer
For I = 1 To ItemLst.ListItems.Count

    If Check4.Value = 1 Then
        ItemLst.ListItems.Item(I).Checked = True
    Else
        ItemLst.ListItems.Item(I).Checked = False
    End If
Next I
End Sub
Private Sub Check5_Click()
Dim I As Integer
For I = 1 To ListView1.ListItems.Count
    If Check5.Value = 1 Then
        ListView1.ListItems.Item(I).Checked = True
    Else
        ListView1.ListItems.Item(I).Checked = False
    End If
Next I
Call ListView1_Click
End Sub
Private Sub Combo1_Click()
    If Combo1.ListIndex = 2 Then
        ListCmd.Top = 225: ClearCmd.Top = 225: RefreshCmd.Top = 225: SearchCmd.Top = 225: ExitCmd.Top = 225
        ShowCmd.Visible = True:  ShowCmd.Top = 850: vcDTP1.Visible = True: vcDTP2.Visible = True: vcDTP1.Value = Date: vcDTP2.Value = Date: Check1.Visible = True: 'Label4.Visible = True
    ElseIf Combo1.ListIndex = 3 Then
        ShowCmd.Visible = False: vcDTP1.Visible = True: vcDTP2.Visible = True: Check1.Visible = False
        ListCmd.Top = 225: ClearCmd.Top = 225: RefreshCmd.Top = 225: SearchCmd.Top = 225: ExitCmd.Top = 225
    ElseIf Combo1.ListIndex = 4 Then
        ShowCmd.Visible = False: vcDTP1.Visible = True: vcDTP2.Visible = True: Check1.Visible = False
        ListCmd.Top = 225: ClearCmd.Top = 225: RefreshCmd.Top = 225: SearchCmd.Top = 225: ExitCmd.Top = 225
    Else
        ShowCmd.Visible = False: vcDTP1.Visible = False: vcDTP2.Visible = False: Check1.Visible = False
        ListCmd.Top = 350: ClearCmd.Top = 350: RefreshCmd.Top = 350: SearchCmd.Top = 350: ExitCmd.Top = 350
    End If
End Sub
'Dim ConnVar As New ADODB.Connection
Private Sub Combo1_GotFocus()
    If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0
End Sub
Private Sub Combo2_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "vcDTP1" Or Me.ActiveControl.NAME = "vcDTP2" Then Sendkeys "{tab}"
    End If
End Sub
Private Sub ListCmd_Click()
    'fill grid with data ***
    Dim Items As String
    Dim LConSno As Long
    Dim GeneralRec As ADODB.Recordset
    Dim LCount As Integer
    Dim LItems As String
    Dim LExCodes As String
    
    If LenB(Combo1.text) = 0 Then MsgBox "Please Select Query Type.", vbCritical: Combo1.SetFocus: Exit Sub
    Label2.Caption = vbNullString: Me.MousePointer = 11
    Set DataGrid1.DataSource = Nothing
    Items = vbNullString
    LExCodes = vbNullString
    LItems = Get_Items
    LExCodes = Get_ExCodes
    If Combo1.ListIndex = 0 Then 'Account
        Set MainRec = Nothing: Set MainRec = New ADODB.Recordset
        mysql = "SELECT AGP.G_NAME AS GROUPNAME,ACC.AC_CODE AS ACCOUNTCODE,AD.PIN AS Client,AD.FMLYID AS UserId ,ACC.NAME AS AccountName,ACC.Op_Bal,ACC.Active  as Status "
        mysql = mysql & "FROM ACCOUNTM AS ACC,AC_GROUP AS AGP,ACCOUNTD AS AD  WHERE ACC.COMPCODE  = " & GCompCode & " AND ACC.COMPCODE=AD.COMPCODE AND ACC.AC_CODE=AD.AC_CODE AND ACC.GCODE=AGP.CODE ORDER BY ACC.NAME ASC"
        MainRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not MainRec.EOF Then MainRec.MoveFirst: Set DataGrid1.DataSource = MainRec: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus
        DataGrid1.Columns(0).Width = 2000
        DataGrid1.Columns(1).Width = 1200
        DataGrid1.Columns(2).Width = 1200
        DataGrid1.Columns(3).Width = 1200
        DataGrid1.Columns(4).Width = 3000
        DataGrid1.Columns(5).Width = 1500
    ElseIf Combo1.ListIndex = 1 Then 'Customer/Supplier
        Set MainRec = Nothing: Set MainRec = New ADODB.Recordset
        mysql = "SELECT AGP.G_NAME AS GroupName,ACC.Ac_Code AS ACCOUNTCODE ,ACC.NAME as Account,ACC.Op_Bal,ACC.Active,PTY.NAME as Account,PTY.ADD1 AS Address,PTY.CITY as City,PTY.PIN AS PinCode,PTY.EMail,PTY.PhoneO,PTY.PhoneR,PTY.Mobile "    ',Fmly.FmlyName
        mysql = mysql & "FROM ACCOUNTM AS ACC,AccountD AS PTY,AC_GROUP AS AGP  WHERE ACC.COMPCODE = " & GCompCode & " AND ACC.COMPCODE=PTY.COMPCODE AND ACC.AC_CODE=PTY.AC_CODE AND ACC.GCODE=AGP.CODE ORDER BY ACC.NAME ASC"     'AND PTY.COMPCODE=Fmly.COMPCODE AND PTY.FmlyId=Fmly.FmlyId    -- ,AccFmly as Fmly
        MainRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not MainRec.EOF Then MainRec.MoveFirst: Set DataGrid1.DataSource = MainRec: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus
    ElseIf Combo1.ListIndex = 3 Then 'MARGIN RATES
        Set MainRec = Nothing: Set MainRec = New ADODB.Recordset
        mysql = "SELECT I.EXCHANGECODE AS ExCode ,I.ItemCode,s.SaudaCode,D.Condate,(IMRATE+EMRATE+AMRATE+ADDLONG+SPCASHMGLONG+ELMLONG+DELIVERY) AS TotBuyMRate,"
        mysql = mysql & " (IMRATE+EMRATE+AMRATE+ADDSHORT+SPCASHMGSHORT+ELMSHORT+DELIVERY) AS TotSellMRate, IMRATE AS [I Margin],EMRATE AS [E Margin],AMRATE AS [Add Margin],ADDLong, ADDShort,"
        mysql = mysql & " SPCASHMGLONG AS [Cash Long],SPCASHMGSHORT AS [Cash Short],ELMLong ,ELMShort,Delivery "
        mysql = mysql & " FROM ITEMMAST AS I , SAUDAMAST AS S, DLYMGN AS D WHERE I.ITEMID =S.ITEMID AND  D.SAUDAID=S.SAUDAID"
        If LenB(LItems) > 1 Then mysql = mysql & " AND I.ITEMCODE IN (" & LItems & ")"
        mysql = mysql & " AND D.CONDATE>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND  D.CONDATE<='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' ORDER BY D.CONDATE ,I.EXCHANGECODE,I.ITEMCODE,S.MATURITY"
        MainRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        
        If Not MainRec.EOF Then MainRec.MoveFirst: Set DataGrid1.DataSource = MainRec: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus
        DataGrid1.Columns(0).Width = 1000
        DataGrid1.Columns(1).Width = 1500
        DataGrid1.Columns(2).Width = 3000
        DataGrid1.Columns(3).Width = 1200
        DataGrid1.Columns(4).Width = 1500
        DataGrid1.Columns(5).Width = 1500
        DataGrid1.Columns(6).Width = 1000
        DataGrid1.Columns(7).Width = 1000
        DataGrid1.Columns(8).Width = 1000
        DataGrid1.Columns(9).Width = 1000
        DataGrid1.Columns(10).Width = 1000
        DataGrid1.Columns(11).Width = 1000
        DataGrid1.Columns(12).Width = 1000
        
        DataGrid1.Columns(4).Alignment = dbgRight
        DataGrid1.Columns(5).Alignment = dbgRight
        DataGrid1.Columns(6).Alignment = dbgRight
        DataGrid1.Columns(7).Alignment = dbgRight
        DataGrid1.Columns(8).Alignment = dbgRight
        DataGrid1.Columns(9).Alignment = dbgRight
        DataGrid1.Columns(10).Alignment = dbgRight
        DataGrid1.Columns(11).Alignment = dbgRight
        DataGrid1.Columns(12).Alignment = dbgRight
        
        DataGrid1.Columns(4).NumberFormat = "0.00"
        DataGrid1.Columns(5).NumberFormat = "0.00"
        DataGrid1.Columns(6).NumberFormat = "0.00"
        DataGrid1.Columns(7).NumberFormat = "0.00"
        DataGrid1.Columns(8).NumberFormat = "0.00"
        DataGrid1.Columns(9).NumberFormat = "0.00"
        DataGrid1.Columns(10).NumberFormat = "0.00"
        DataGrid1.Columns(11).NumberFormat = "0.00"
        DataGrid1.Columns(12).NumberFormat = "0.00"
        
    ElseIf Combo1.ListIndex = 4 Then 'CLOSING RATES
        Set MainRec = Nothing: Set MainRec = New ADODB.Recordset
        mysql = "SELECT D.ConDate,I.ItemCode , S.SaudaCode , S.Maturity ,D.CloseRate,D.OPRate AS OpenRate ,D.HGRATE AS HighRate ,D.LOWRATE AS LowRate "
        mysql = mysql & " FROM ItemMast AS I , SAUDAMAST AS S, CTR_R AS D WHERE I.COMPCODE=" & GCompCode & " AND  D.SAUDAID=S.SAUDAID AND I.ITEMID =S.ITEMID"
        If LenB(LItems) > 1 Then mysql = mysql & " AND I.ITEMCODE IN (" & Items & ")"
        mysql = mysql & " AND D.CONDATE>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND  D.CONDATE<='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' ORDER BY D.CONDATE ,I.ITEMNAME, S.SAUDANAME, S.MATURITY"
        MainRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not MainRec.EOF Then MainRec.MoveFirst: Set DataGrid1.DataSource = MainRec: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus
        
        DataGrid1.Columns(0).Width = 1200
        DataGrid1.Columns(1).Width = 1500
        DataGrid1.Columns(2).Width = 3000
        DataGrid1.Columns(3).Width = 1200
        DataGrid1.Columns(4).Width = 1500
        DataGrid1.Columns(5).Width = 1500
        DataGrid1.Columns(6).Width = 1000
        DataGrid1.Columns(4).Alignment = dbgRight
        DataGrid1.Columns(5).Alignment = dbgRight
        DataGrid1.Columns(6).Alignment = dbgRight
        DataGrid1.Columns(7).Alignment = dbgRight
        DataGrid1.Columns(4).NumberFormat = "0.00"
        DataGrid1.Columns(5).NumberFormat = "0.00"
        DataGrid1.Columns(6).NumberFormat = "0.00"
        
    ElseIf Combo1.ListIndex = 6 Then 'BROK
        mysql = "SELECT A.AC_CODE,B.NAME,A.EXCODE,A.INSTTYPE,CASE A.BROKTYPE WHEN 'P' THEN 'Percentage Wise 'when 'O' THEN 'Opening Sauda'when 'T' THEN 'Transaction Wise' else 'Other Type'end as BrokType,"
        mysql = mysql & "  A.BROKRATE,A.UPTOSTDT FROM PEXBROK AS A,ACCOUNTD AS B WHERE B.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE "
        If LenB(LExCodes) > 0 Then mysql = mysql & " AND A.EXCODE IN (" & LExCodes & ")"
        mysql = mysql & " AND A.AC_CODE =B.AC_CODE order by B.NAME,A.EXCODE,A.INSTTYPE,A.UPTOSTDT"
        Set MainRec = Nothing: Set MainRec = New ADODB.Recordset
        MainRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not MainRec.EOF Then MainRec.MoveFirst: Set DataGrid1.DataSource = MainRec: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus
        DataGrid1.Columns(0).Width = 1200
        DataGrid1.Columns(1).Width = 350
        DataGrid1.Columns(2).Width = 1500
        DataGrid1.Columns(3).Width = 900
        DataGrid1.Columns(4).Width = 2000
        DataGrid1.Columns(5).Width = 1500
        DataGrid1.Columns(5).Alignment = dbgRight
        DataGrid1.Columns(5).NumberFormat = "0.0000"
        
    ElseIf Combo1.ListIndex = 10 Then 'BROK
        mysql = "SELECT DISTINCT CONDATE,SAUDA,SAUDAID FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND BILLNO = 0 AND PATTAN ='C' ORDER BY CONDATE, SAUDA"
        Set MainRec = Nothing: Set MainRec = New ADODB.Recordset
        MainRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not MainRec.EOF Then MainRec.MoveFirst: Set DataGrid1.DataSource = MainRec: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus
        DataGrid1.Columns(0).Width = 1500
        DataGrid1.Columns(1).Width = 3500
        
    ElseIf Combo1.ListIndex = 2 Then  'Contract
        If vcDTP1.Value > vcDTP2.Value Then MsgBox "Invalid date range", vbCritical: vcDTP1.SetFocus: Exit Sub
        Call RecSet
        If Check1.Value = 1 Then
            Set listrec = Nothing: Set listrec = New ADODB.Recordset
            mysql = "SELECT DISTINCT CTR_M.CONDATE,CTR_M.SAUDA,CTR_M.ITEMCODE,CTR_M.CONSNO,CTR_M.PATTAN AS MPATTAN,CTR_D.EXCODE "
            mysql = mysql & " FROM CTR_M,CTR_D WHERE CTR_M.COMPCODE=" & GCompCode & " AND  CTR_M.COMPCODE =CTR_D.COMPCODE AND CTR_M.ConSno = CTR_D.CONSNO "
            mysql = mysql & " AND CTR_M.CONDATE >= '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND CTR_M.CONDATE <= '" & Format(vcDTP2.Value, "yyyy/MM/dd") & "'"
            If LenB(LExCodes) > 0 Then mysql = mysql & " AND CTR_D.EXCODE IN (" & LExCodes & ")"
            If LenB(LItems) > 0 Then mysql = mysql & " AND CTR_D.ITEMCODE IN (" & LItems & ")"
            mysql = mysql & " ORDER BY CTR_M.CONDATE,CTR_D.EXCODE,CTR_M.SAUDA"
            listrec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            If Not listrec.EOF Then
                Set DataGrid1.DataSource = listrec: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0: DataGrid1.Col = 0: DataGrid1.SetFocus:
                DataGrid1.Columns(0).Width = 1200:
                DataGrid1.Columns(1).Width = 3500
                DataGrid1.Columns(2).Width = 2500
                DataGrid1.Columns(3).Width = 1000
                DataGrid1.Columns(3).Alignment = dbgRight
                DataGrid1.Columns(4).Width = 1000
                DataGrid1.Columns(4).Alignment = dbgCenter
                DataGrid1.Columns(5).Width = 1000
            End If
        Else
            Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
            mysql = "SELECT CTR_M.CONSNO,CTR_M.CONDATE,CTR_M.SAUDA,CTR_M.ITEMCODE,CTR_M.PATTAN AS MPATTAN,CTR_D.*, A.NAME AS NAME "
            mysql = mysql & "FROM CTR_M,CTR_D, ACCOUNTD AS A WHERE CTR_M.COMPCODE =" & GCompCode & " AND CTR_M.COMPCODE = CTR_D.COMPCODE AND CTR_M.CONSNO = CTR_D.CONSNO AND CTR_D.COMPCODE = A.COMPCODE AND CTR_D.PARTY = A.AC_CODE  AND  CTR_M.CONDATE >= '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND  CTR_M.CONDATE <= '" & Format(vcDTP2.Value, "yyyy/MM/dd") & "' "
            mysql = mysql & " AND CTR_D.ITEMCODE IN (" & LItems & ")"
            mysql = mysql & " ORDER BY CTR_M.CONSNO,CTR_D.CONNO,CTR_D.CONTYPE"
            GeneralRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                Do While Not GeneralRec.EOF
                    LConSno = GeneralRec!CONSNO
                    RECGRID.AddNew
                    RECGRID!Condate = GeneralRec!Condate
                    RECGRID!CONSNO = GeneralRec!CONSNO
                    RECGRID!Sauda = GeneralRec!Sauda
                    RECGRID!Item = GeneralRec!ITEMCODE
                    RECGRID!CLOSERATE = 0
                    RECGRID!PATTAN = GeneralRec!MPATTAN
                    RECGRID!TRADENO = GeneralRec!CONNO
                    RECGRID!BCODE = GeneralRec!PARTY & ""
                    RECGRID!BNAME = GeneralRec!NAME
                    RECGRID!BQnty = GeneralRec!QTY
                    RECGRID!BRate = GeneralRec!Rate
                    RECGRID!DIMPORT = IIf(GeneralRec!DATAIMPORT & "", 1, 0)
                    RECGRID!contime = IIf(IsNull(GeneralRec!contime), Time, GeneralRec!contime)
                    RECGRID!CLCODE = IIf(IsNull(GeneralRec!CLCODE), "", GeneralRec!CLCODE)
                    GeneralRec.MoveNext
                    RECGRID!scode = GeneralRec!PARTY & ""
                    RECGRID!SNAME = GeneralRec!NAME
                    RECGRID!SQnty = GeneralRec!QTY
                    RECGRID!SRate = GeneralRec!Rate
                    RECGRID.Update
                    GeneralRec.MoveNext
                Loop
                If Not RECGRID.EOF Then
                    RECGRID.MoveFirst: Set DataGrid1.DataSource = RECGRID:
                    DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Row = 0:
                    DataGrid1.Col = 0: DataGrid1.SetFocus
                    DataGrid1.Columns(0).Width = 1000:
                    DataGrid1.Columns(1).Width = 3000
                    DataGrid1.Columns(2).Width = 2000
                    DataGrid1.Columns(3).Width = 1000
                    DataGrid1.Columns(3).Alignment = dbgRight
                    DataGrid1.Columns(4).Width = 1000
                    DataGrid1.Columns(4).Alignment = dbgRight
                    DataGrid1.Columns(4).NumberFormat = "0.00"
                    DataGrid1.Columns(5).Width = 2000
                    DataGrid1.Columns(6).Width = 1000
                    DataGrid1.Columns(6).Alignment = dbgRight
                    DataGrid1.Columns(7).Width = 1000
                    DataGrid1.Columns(7).Alignment = dbgRight
                    DataGrid1.Columns(7).NumberFormat = "0.00"
                    DataGrid1.Columns(8).Width = 1000
                    DataGrid1.Columns(8).Alignment = dbgCenter
                    DataGrid1.Columns(9).Width = 1000
                    ShowCmd.Enabled = True
                End If
            End If
         End If
    End If
    Me.MousePointer = 0
End Sub
Private Sub ExitCmd_Click()
    Unload Me
End Sub
Private Sub ClearCmd_Click()
    Set DataGrid1.DataSource = Nothing
    SEARCHTXT = ""
End Sub
Private Sub RefreshCmd_Click()
    On Error GoTo Error1
    Me.MousePointer = 11
    If Combo1.ListIndex = 2 Then
        If RECGRID.RecordCount > 0 Then RECGRID.Requery: Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    End If
Error1:
    Label2.Caption = "": SEARCHTXT = ""
    Me.MousePointer = 0
End Sub
Private Sub SearchCmd_Click()
    On Error GoTo Error1
    MainRec.MoveFirst
    SEARCHTXT = InputBox("Enter  " & DataGrid1.Columns.Item(DataGrid1.Col).Caption, "Search", , 6350, 1300)
    Call FINDREC
Error1:     Exit Sub
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    'SORTING ***
    On Error GoTo Error1
    DataGrid1.MarqueeStyle = dbgHighlightCell
    DoEvents
    If Left$(Label2.Caption, 1) = "A" Then
        SORTORDER = DataGrid1.Columns.Item(ColIndex).DataField & "  DESC"
        If Combo1.ListIndex = 2 Then
            RECGRID.Sort = ("" & SORTORDER & "")
        Else
            MainRec.Sort = ("" & SORTORDER & "")
        End If
        Label2.Caption = "Desc. order on " & DataGrid1.Columns.Item(ColIndex).Caption
    Else
        SORTORDER = DataGrid1.Columns.Item(ColIndex).DataField & "  ASC"
        If Combo1.ListIndex = 2 Then
            RECGRID.Sort = ("" & SORTORDER & "")
        Else
            MainRec.Sort = ("" & SORTORDER & "")
        End If
        Label2.Caption = "Asc. order on " & DataGrid1.Columns.Item(ColIndex).Caption
    End If
    DoEvents
    If Not MainRec.EOF Then Set DataGrid1.DataSource = MainRec: DataGrid1.ReBind: DataGrid1.Refresh
Error1:    Exit Sub
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"
    If KeyCode = 114 Then
        If Not MainRec.EOF Then
            MainRec.MoveNext
        Else
            MainRec.MoveFirst
        End If
        Call FINDREC
    End If
End Sub
Private Sub Form_Load()
    Dim TRec As ADODB.Recordset
    Combo1.ListIndex = 2
    vcDTP1.MinDate = GFinBegin: vcDTP2.MinDate = GFinBegin
    vcDTP1.MaxDate = GFinEnd: vcDTP2.MaxDate = GFinEnd
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    mysql = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXNAME "
    TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    ListView1.Enabled = True: Check5.Enabled = True
    If Not TRec.EOF Then
        Me.MousePointer = 11
        ListView1.ListItems.Clear
        Do While Not TRec.EOF
            ListView1.ListItems.Add , , TRec!excode
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , TRec!EXNAME
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , TRec!excode
            TRec.MoveNext
        Loop
        Me.MousePointer = 0
        
        If TRec.RecordCount = 1 Then
            Check5.Value = 1: ListView1.TabStop = False: Check5.TabStop = False
            Call Check5_Click
        Else
            ListView1.TabStop = True: Check5.TabStop = True
        End If
    Else
        Check5.Enabled = False
    End If
End Sub
Private Sub Frame5_Click()
    GETMAIN.CommonDialog1.InitDir = "c:\a"
    GETMAIN.CommonDialog1.FileName = "testing.bmp"
    GETMAIN.CommonDialog1.ShowSave
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Private Sub ShowCmd_Click()
    On Error GoTo Error1
    Dim NRec1 As ADODB.Recordset
    If Combo1.ListIndex = 2 Then
        Me.MousePointer = 11
        If IsDate(DataGrid1.Columns(0).text) Then
            If SYSTEMLOCK(DateValue(DataGrid1.Columns(0).text)) Then
                MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
            Else
                If GGenQuery = "3" Then
                    
                ElseIf GGenQuery = "4" Then
                    FRM_NEW_SINGLE_ENTRY.Show
                    FRM_NEW_SINGLE_ENTRY.Fb_Press = 2
                    Call FRM_NEW_SINGLE_ENTRY.MODIFY_REC(DataGrid1.Columns(0).text, DataGrid1.Columns(1).text, DataGrid1.Columns(3).text, DataGrid1.Columns(5).text)
                    FRM_NEW_SINGLE_ENTRY.Frame1.Enabled = True: FRM_NEW_SINGLE_ENTRY.DataGrid1.SetFocus
                ElseIf GGenQuery = "2" Then
                    CTRBUYSELL.Show
                    CTRBUYSELL.Fb_Press = 2
                    Call CTRBUYSELL.MODIFY_REC(DataGrid1.Columns(0).text, DataGrid1.Columns(1).text, DataGrid1.Columns(5).text)
                    CTRBUYSELL.Frame1.Enabled = True: CTRBUYSELL.DataGrid1.SetFocus
                Else
                    Call PERMISSIONS("CONTRACTENTRY")
                    If ModiPerm = True Then
                        GETCont.Show
                        GETCont.Fb_Press = 2
                        If Check1.Value = 1 Then
                            GETCont.MODIFY_REC DataGrid1.Columns(0).text, DataGrid1.Columns(1).text, DataGrid1.Columns(5).text
                        Else
                            GETCont.MODIFY_REC DataGrid1.Columns(0).text, DataGrid1.Columns(1).text, DataGrid1.Columns(12).text
                        End If
                    Else
                        MsgBox "Modification Rights not Available"
                        Exit Sub
                    End If
                End If
            End If
        End If
        Me.MousePointer = 0
    End If
Error1:
    
    'MsgBox Err.Description
End Sub
Sub FINDREC()
    On Error GoTo Error1
    If SEARCHTXT = "" Then
    Else
        SEARCHTXT = "%" & SEARCHTXT & "%"
        MainRec.Find ("" & DataGrid1.Columns.Item(DataGrid1.Col).DataField & " LIKE '" & SEARCHTXT & "'"), , adSearchForward
        If Not MainRec.EOF Then
            DataGrid1.Row = MainRec.AbsolutePosition - 1
        Else
            MsgBox "Record not found"
            DataGrid1.Row = 0
        End If
    End If
    DataGrid1.Col = DataGrid1.Col: DataGrid1.SetFocus
Error1:    Exit Sub
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    If Check1.Value = 1 Then
        RECGRID.Fields.Append "ConDate", adDate, , adFldIsNullable
        RECGRID.Fields.Append "Sauda", adVarChar, 50, adFldIsNullable
        RECGRID.Fields.Append "CloseRate", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "Item", adVarChar, 20, adFldIsNullable
        RECGRID.Fields.Append "ConSno", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "Pattan", adVarChar, 1, adFldIsNullable
        RECGRID.Fields.Append "EXCODE", adVarChar, 10, adFldIsNullable
    Else
        RECGRID.Fields.Append "ConDate", adDate, , adFldIsNullable
        RECGRID.Fields.Append "Sauda", adVarChar, 50, adFldIsNullable
        RECGRID.Fields.Append "BNAME", adVarChar, 150, adFldIsNullable
        RECGRID.Fields.Append "BQNTY", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "BRATE", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "SNAME", adVarChar, 150, adFldIsNullable
        RECGRID.Fields.Append "SQNTY", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "SRATE", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "TradeNo", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "Contime", adVarChar, 15, adFldIsNullable
        RECGRID.Fields.Append "CloseRate", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "DImport", adInteger, , adFldIsNullable
        RECGRID.Fields.Append "Pattan", adVarChar, 1, adFldIsNullable
        RECGRID.Fields.Append "BCODE", adVarChar, 15, adFldIsNullable
        RECGRID.Fields.Append "SCODE", adVarChar, 15, adFldIsNullable
        RECGRID.Fields.Append "ConSno", adDouble, , adFldIsNullable
        RECGRID.Fields.Append "ITEM", adVarChar, 20, adFldIsNullable
        RECGRID.Fields.Append "CLCODE", adVarChar, 20, adFldIsNullable
    End If
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
End Sub
Private Sub ListView1_Click()
    Dim RecSauda As ADODB.Recordset
    Dim LSExCodes As String
    Dim I As Integer
    Dim ListIt As ListItem
    LSExCodes = vbNullString
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            If LenB(LSExCodes) <> 0 Then LSExCodes = LSExCodes & ", "
            LSExCodes = LSExCodes & "'" & ListView1.ListItems(I).ListSubItems(2) & "'"
        End If
  Next I
  ItemLst.ListItems.Clear
  If LenB(LSExCodes) = 0 Then Me.MousePointer = 0: Exit Sub
    mysql = "SELECT ITEMCODE,ITEMNAME,LOT,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND EXCHANGECODE in (" & LSExCodes & ")  ORDER BY EXCHANGECODE,ITEMNAME"
    Set RecSauda = Nothing: Set RecSauda = New ADODB.Recordset: RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    While Not RecSauda.EOF
        Set ListIt = ItemLst.ListItems.Add(, , RecSauda!ITEMName)
        ListIt.SubItems(1) = RecSauda!ITEMCODE
        RecSauda.MoveNext
    Wend
    Set RecSauda = Nothing
End Sub
Public Function Get_Items() As String
    Dim LCount As Integer
    Dim LItems As String
    Dim J As Integer
    LItems = vbNullString
    LCount = ItemLst.ListItems.Count
    For J = 1 To ItemLst.ListItems.Count
        If ItemLst.ListItems(J).Checked = True Then
            LCount = LCount - 1
            If LenB(LItems) <> 0 Then LItems = LItems & ", "
            LItems = LItems & "'" & ItemLst.ListItems(J).SubItems(1) & "'"
        End If
    Next
    If LCount = 0 Then LItems = vbNullString
    Get_Items = LItems
End Function
Public Function Get_ExCodes() As String
    Dim LCount As Integer
    Dim LExCodes As String
    Dim J As Integer
    LExCodes = vbNullString
    LCount = ListView1.ListItems.Count
    For J = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(J).Checked = True Then
            LCount = LCount - 1
            If LenB(LExCodes) <> 0 Then LExCodes = LExCodes & ", "
            LExCodes = LExCodes & "'" & ListView1.ListItems(J) & "'"
        End If
    Next
    If LCount = 0 Then LExCodes = vbNullString
    Get_ExCodes = LExCodes
End Function

Private Sub vcDTP1_Validate(Cancel As Boolean)
    Label3.Visible = False:                TxtAdminPass.Visible = False
    ShowCmd.Enabled = True
End Sub
