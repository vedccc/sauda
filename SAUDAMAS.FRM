VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form SAUDAmast 
   BackColor       =   &H80000000&
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   495
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   3735
      Left            =   2843
      TabIndex        =   4
      Top             =   1725
      Width           =   6975
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37861.9121759259
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   1770
         Width           =   4575
         _ExtentX        =   8070
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
      Begin VB.TextBox Text1 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   885
         Width           =   1215
      End
      Begin VB.TextBox Text2 
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
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   885
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maturity"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   8
         Top             =   2715
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   1845
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauda Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   3375
         Left            =   120
         Top             =   240
         Width           =   6735
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4920
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
Attribute VB_Name = "SAUDAmast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fb_press As Byte
Dim REC As ADODB.Recordset
Sub ADD_REC()
    fb_press = 1
    Call Get_Selection(1)
    Frame1.Enabled = True
    Text1.SetFocus
End Sub
Sub SAVE_REC()
    If Len(Trim(Text1.Text)) < 1 Then
        MsgBox "Sauda code required before saving record.", vbCritical, "Error"
        Exit Sub
    End If

    If fb_press = 1 Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!SAUDACODE = Text1.Text
    End If
    Adodc1.Recordset!CompCode = MC_CODE
    Adodc1.Recordset!SAUDANAME = Text2.Text
    Adodc1.Recordset!ItemCode = DataCombo1.BoundText
    Adodc1.Recordset!MATURITY = Format(vcDTP1.Value, "dd/MM/yyyy")
    Adodc1.Recordset.Update

    Call CANCEL_REC
End Sub
Sub CANCEL_REC()
    Call ClearFormFn(SAUDAMAST)
    vcDTP1.MinDate = CDate("01/01/1900")
    vcDTP1.MaxDate = CDate("01/01/2900")
    DataCombo1.Locked = False
    Text1.Enabled = True
    fb_press = 0
    Call Get_Selection(10)
    Frame1.Enabled = False
End Sub
Sub MODIFY_REC()
    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find "SAUDACODE='" & GETMAIN.DataCombo1.BoundText & "'", , adSearchForward
    If Not Adodc1.Recordset.EOF Then
        Text1.Text = Adodc1.Recordset!SAUDACODE
        Text1.Enabled = False
        Text2.Text = Adodc1.Recordset!SAUDANAME
        DataCombo1.BoundText = Adodc1.Recordset!ItemCode
        vcDTP1.Value = Adodc1.Recordset!MATURITY
        Frame1.Enabled = True

        Set REC = Nothing
        Set REC = New ADODB.Recordset
        MYSQL = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND SAUDA='" & Text1.Text & "'"
        REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
        If Not REC.EOF Then
            DataCombo1.Locked = True
        Else
            DataCombo1.Locked = False
        End If
        
        Text2.SetFocus
    End If

    If fb_press = 3 Then    ''FOR DELETE
        If MsgBox("You are about to delete this record. Confirm Delete?", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
            Set REC = Nothing
            Set REC = New ADODB.Recordset
            MYSQL = "SELECT * FROM CTR_M WHERE compcode=" & MC_CODE & " AND SAUDA='" & Text1.Text & "'"
            REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
            If Not REC.EOF Then
                MsgBox "Transaction Exists Can't Delete Sauda.", vbExclamation, "Error"
            Else
                Adodc1.Recordset.Delete
            End If
        End If
        Call CANCEL_REC
    End If
    Else
        MsgBox "Record does not exists.", vbCritical
        Call CANCEL_REC
    End If
End Sub
Private Sub DataCombo1_GotFocus()
    If fb_press = 1 Then SendKeys "%{DOWN}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
    Call CANCEL_REC

    Adodc1.ConnectionString = CNN
    Adodc1.RecordSource = "SELECT * FROM SAUDAMAST where compcode=" & MC_CODE & " ORDER BY SAUDACODE"
    Adodc1.Refresh

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    If SWTYPE = "SQL" Then
        REC.Open "SELECT ITEMCODE, ITEMCODE+','+ITEMNAME AS ITEMNAME FROM ITEMMAST WHERE compcode=" & MC_CODE & "  ORDER BY ITEMCODE", CNN, adOpenKeyset, adLockReadOnly
    Else
        REC.Open "SELECT ITEMCODE, ITEMCODE+','+ITEMNAME AS ITEMNAME FROM ITEMMAST WHERE compcode=" & MC_CODE & "  ORDER BY ITEMCODE", CNN, adOpenKeyset, adLockReadOnly
    End If
    Set DataCombo1.RowSource = REC
    DataCombo1.BoundColumn = "ITEMCODE"
    DataCombo1.ListField = "ITEMNAME"

    Call CANCEL_REC
End Sub
Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    Frame1.BackColor = Me.BackColor
End Sub
Private Sub Text1_Validate(cancel As Boolean)
    If fb_press = 1 Then
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT * FROM SAUDAMAST WHERE compcode=" & MC_CODE & " AND SAUDACODE='" & Text1.Text & "'", CNN, adOpenForwardOnly, adLockReadOnly
        If Not REC.EOF Then
            MsgBox "Sauda code already exists.", vbExclamation, "Warning"
            cancel = True
        End If
    End If
End Sub
Sub List_Sauda()
    Screen.MousePointer = 11

    Call Get_Selection(12)

    MYSQL = "SELECT a.saudacode, A.saudaname, B.ITEMCODE, B.ITEMNAME, maturity FROM ITEMMAST as b,SAUDAMAST as a where A.compcode=" & MC_CODE & " AND A.compcode=B.compcode AND a.itemcode=b.itemcode ORDER BY B.ITEMNAME"
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open MYSQL, CNN, adOpenKeyset, adLockReadOnly

    Set RDCREPO = RDCAPP.OpenReport(RPT_PATH & "ItemSauda.RPT", 1)

    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource REC

'    RDCREPO.FormulaFields.GetItemByName("TITLE").Text = "'" & MC_NAME & "'"

    CRViewer1.Width = CInt(GETMAIN.Width - 100)
    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Top = 0
    CRViewer1.Left = 0

    CRViewer1.Visible = True
    CRViewer1.ReportSource = RDCREPO

    CRViewer1.ViewReport

    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub
