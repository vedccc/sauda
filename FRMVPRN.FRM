VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRVIEWER.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form frmvprn 
   Caption         =   "Voucher Printing"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Voucher Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   4560
         Width           =   945
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   4560
         Width           =   945
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Voucher No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Voucher Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Narration"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   38061.9461689815
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   38061.945150463
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmvprn.frx":0000
         Left            =   6240
         List            =   "frmvprn.frx":000D
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   480
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "VoucherType"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
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
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmvprn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MREC As ADODB.Recordset
Dim mvtype As String
Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex = -1 Then
        MsgBox "Please Select Voucher Type"
    End If
    Select Case Combo1.ListIndex
        Case 0
            mvtype = "CV"
        Case 1
            mvtype = "BV"
        Case 2
            mvtype = "JV"
    End Select

End Sub
Private Sub Command1_Click()
    On Error GoTo ERR1
    Dim vous As String, I As Integer
    vous = ""

    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
           
            vous = vous & "'"
            vous = vous & ListView1.ListItems(I).Text
            vous = vous & "'"
            If I < ListView1.ListItems.Count Then
                If ListView1.ListItems(I + 1).Checked = True Then
                    vous = vous & ", "
                End If
            End If
        End If
        DoEvents
    Next
    Set MREC = Nothing
    Set MREC = New ADODB.Recordset
    msql = "SELECT V.VOU_DT, V.VOU_TYPE, A.AC_CODE, A.NAME, VT.DR_CR, VT.NARRATION, VT.CHEQUE_NO, VT.CHEQUE_DT, VT.AMOUNT AS CL_BAL, A.GCODE, VT.AMOUNT, V.VOU_NO, VT.BANK_NAME, VT.BRANCH, VT.VOU_ID, VT.AC_CODE FROM ACCOUNTM AS A, VCHAMT AS VT, VOUCHER AS V WHERE V.VOU_NO=VT.VOU_NO AND VT.AC_CODE=A.AC_CODE AND VT.VOU_DT >= DATEVALUE('" & vcDTP1.Value & "') AND VT.VOU_DT <=datevalue('" & vcDTP2.Value & "') AND V.VOU_TYPE ='" & VOU_TYPE & "' AND V.VOU_NO IN (" & vous & ") ORDER BY Vt.VOU_DT, Vt.VOU_NO, Vt.VOU_ID"
    MREC.Open msql, CNN, adOpenForwardOnly, adLockReadOnly
    Set RDCREPO = RDCAPP.OpenReport(RPT_PATH & "vOUprn.rpt", 1)
    MYSQL = "JOURNAL VOUCHER "

    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource MREC

    RDCREPO.FormulaFields.GetItemByName("ORG").Text = "'" & MC_NAME & "'"
    RDCREPO.FormulaFields.GetItemByName("TITLE").Text = MYSQL

    CRViewer1.ZOrder

    CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)

    CRViewer1.Visible = True
    CRViewer1.ReportSource = RDCREPO
    CRViewer1.Zoom 1
    CRViewer1.ViewReport

    Me.MousePointer = 0
    GETMAIN.ProgressBar1.Visible = False

    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
    Me.MousePointer = Val(0)
    GETMAIN.ProgressBar1.Visible = False
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub DataCombo1_GotFocus()
    SendKeys "%{down}"
End Sub

Private Sub Form_Load()
    Dim msql As String
    Set MREC = Nothing
    Set MREC = New ADODB.Recordset
    msql = "Select ac_code, Name from Accountm order by Name"
    MREC.Open msql, CNN, adOpenStatic, adLockReadOnly
    Set DataCombo1.RowSource = MREC
    DataCombo1.BoundColumn = "ac_code"
    DataCombo1.ListField = "Name"
    DataCombo1.Text = ""
    Combo1.ListIndex = 0
    ListView1.ListItems.clear
End Sub
Private Sub vcDTP2_Validate(Cancel As Boolean)
    ListView1.ListItems.clear
    Set MREC = Nothing
    Set MREC = New ADODB.Recordset
    msql = "SELECT V.VOU_NO, V.VOU_DT, VT.AMOUNT, Vt.NARRATION FROM VOUCHER AS V, VCHAMT AS VT WHERE V.VOU_NO=VT.VOU_NO AND V.VOU_DT >= DATEVALUE('" & vcDTP1.Value & "')  AND V.VOU_DT <= DATEVALUE('" & vcDTP2.Value & "') AND VT.AC_CODE='" & DataCombo1.BoundText & "' AND V.VOU_TYPE='" & mvtype & "' ORDER BY v.VOU_NO, V.VOU_DT"
    MREC.Open msql, CNN, adOpenForwardOnly, adLockReadOnly
    Do While Not MREC.EOF
        ListView1.ListItems.ADD , , MREC!VOU_NO
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , MREC!VOU_DT
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , MREC!NARRATION
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , MREC!AMOUNT
        MREC.MoveNext
    Loop
End Sub
