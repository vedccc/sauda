VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form GETACRPT 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   2370
   ClientWidth     =   11880
   Icon            =   "Acreport.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00008080&
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
      Height          =   855
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11775
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   11175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   11415
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3855
         Left            =   2400
         TabIndex        =   13
         Top             =   120
         Width           =   6615
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   855
            Left            =   960
            TabIndex        =   26
            Top             =   600
            Width           =   4575
            Begin VB.CheckBox ChkCash 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Cash Voucher"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   30
               Top             =   120
               Width           =   1695
            End
            Begin VB.CheckBox ChkJV 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Journal Voucher"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2400
               TabIndex        =   29
               Top             =   120
               Width           =   1815
            End
            Begin VB.CheckBox ChkSet 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Settle Voucher"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   28
               Top             =   480
               Width           =   1695
            End
            Begin VB.CheckBox ChkShare 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Share Voucher"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2400
               TabIndex        =   27
               Top             =   480
               Width           =   1815
            End
         End
         Begin VB.CommandButton CANCEL_CMD 
            BackColor       =   &H00C0E0FF&
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2760
            TabIndex        =   23
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton OK_CMD 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&OK"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   960
            TabIndex        =   22
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   15
            Top             =   2100
            Visible         =   0   'False
            Width           =   4935
         End
         Begin vcDateTimePicker.vcDTP FROM_DT 
            Height          =   375
            Left            =   960
            TabIndex        =   14
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37680.7250462963
         End
         Begin vcDateTimePicker.vcDTP TO_DT 
            Height          =   375
            Left            =   3960
            TabIndex        =   16
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37680.7250462963
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   420
            Left            =   960
            TabIndex        =   17
            Top             =   2580
            Visible         =   0   'False
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   741
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   3480
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2670
            Visible         =   0   'False
            Width           =   975
         End
      End
   End
   Begin VB.Frame delpen_fram 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4200
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
      Begin vcDateTimePicker.vcDTP DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37680.7250462963
      End
      Begin vcDateTimePicker.vcDTP DTPicker2 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37680.7250462963
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFF80&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1575
         Left            =   120
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Frame Inv_fram 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3720
      TabIndex        =   0
      Top             =   2160
      Width           =   4935
      Begin VB.TextBox INVNO_TO_TXT 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox INVNO_FROM_TXT 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFF80&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   2055
         Left            =   120
         Top             =   120
         Width           =   4695
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   870
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
   Begin VB.Line Line8 
      BorderColor     =   &H00000040&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   5820
      Left            =   75
      Top             =   960
      Width           =   11700
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "GETACRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MYRS As ADODB.Recordset
Public MYRS_CASH As ADODB.Recordset

Public inv_type As String:          Public LOC_CODE As String: Public VOU_TYPE As String
Dim BANK_BRANCH As String:          Public WORDAMT1 As String: Public NARR_CHNO As String
Dim Rec As ADODB.Recordset:         Dim AccRec As ADODB.Recordset
Public DATE1 As Date:               Public COUNT_LOOP As Long

Public TOTAL As Double:             Public DR_AMT As Double:   Public CR_AMT As Double
Public OP_BAL As Double:            Public CL_BAL As Double:   Public AMOUNT1 As Double
Public REPT_AMT As Double:          Public PAMT_AMT As Double: Public NEXT_OP_BAL As Double
Dim MVCHSERIES As String:           Dim BOOKREC As ADODB.Recordset

Private Sub ACC_LIST_GotFocus()
    DataCombo1.text = vbNullString
End Sub
Private Sub Check1_Click()
   If Check1.Value = 1 Then
        Check1.Caption = "Day Total Required"
   Else
        Check1.Caption = "Day Total Not Required"
   End If
End Sub
Private Sub CANCEL_CMD_Click()
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    Screen.MousePointer = 0
    Unload Me
End Sub
Private Sub DataCombo1_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub OK_CMD_Click()
    If FROM_DT.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: FROM_DT.SetFocus: Exit Sub
    If FROM_DT.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: FROM_DT.SetFocus: Exit Sub
    
    If TO_DT.Visible Then
        If TO_DT.Value < FROM_DT.Value Then MsgBox "Invalid date range.", vbCritical: FROM_DT.SetFocus: Exit Sub
        If TO_DT.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: TO_DT.SetFocus: Exit Sub
        If TO_DT.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: TO_DT.SetFocus: Exit Sub
    End If
    If MFormat = "Cheque Register" Or MFormat = "Voucher List" Or MFormat = "Cash Book Regular Format" Or MFormat = "Cash Book Ledger Format" Or MFormat = "Bank Book Regular Format" Or MFormat = "Bank Book Ledger Format" Then
        Call CRVLCBBB
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        CRViewer1.Visible = False
        Cancel = 1
    Else
        GETMAIN.StatusBar1.Panels(1).text = vbNullString
        If CNNERR = True Then
            Cnn.RollbackTrans
            CNNERR = False
        End If
        Unload Me
    End If
End Sub
Sub CRVLCBBB()      ''''CASH RECEPT VOUCHER LIST CASH BOOK BANK BOOK.
'On Error GoTo Error1
    Dim LAMT As Double:         Dim MAc_Code As String:    Dim LOp_Bal As Double:      Dim MDrCr  As String
    Dim LAmount As Double:      Dim LExCode As String:     Dim LVou_Type As String:    Dim LVou_Dt As Date
    
    If FROM_DT.Value < DateValue(GFinBegin) Or TO_DT.Value > DateValue(GFinEnd) Then
        MsgBox "Invalid Date " + CStr(FROM_DT.Value), vbInformation, "Error"
        FROM_DT.Value = Date
        Screen.MousePointer = 0
        FROM_DT.SetFocus
        Exit Sub
    End If
    OK_CMD.Enabled = False
    Screen.MousePointer = 11
    If MFormat = "Voucher List" Then       ''Voucher List Starts From Here
        Dim ToDt As Variant
        Dim FrDt As Variant
        Call RecSet

        If ChkCash.Value = 1 Then VOU_TYPE = "'CV'"
        
        If ChkJV.Value = 1 Then
            If LenB(VOU_TYPE) > 0 Then VOU_TYPE = VOU_TYPE & ","
            VOU_TYPE = VOU_TYPE & "'JV'"
        End If
        If ChkSet.Value = 1 Then
            If LenB(VOU_TYPE) > 0 Then VOU_TYPE = VOU_TYPE & ","
            VOU_TYPE = VOU_TYPE & "'S'"
        End If
        If ChkShare.Value = 1 Then
            If LenB(VOU_TYPE) > 0 Then VOU_TYPE = VOU_TYPE & ","
            VOU_TYPE = VOU_TYPE & "'H'"
        End If
        Set Rec = Nothing
        Set Rec = New ADODB.Recordset
        mysql = "SELECT V.VOU_DT,V.VOU_TYPE,A.AC_CODE,A.NAME,VT.DR_CR,VT.NARRATION,VT.AMOUNT AS CL_BAL,"
        mysql = mysql & " A.GCODE, VT.AMOUNT, V.VOU_NO, VT.VOUID, VT.AC_CODE FROM ACCOUNTM AS A, "
        mysql = mysql & " VCHAMT AS VT, VOUCHER AS V WHERE V.COMPCODE=" & GCompCode & "  "
        mysql = mysql & " AND V.COMPCODE=A.COMPCODE AND V.VOU_ID=VT.VOU_ID  AND VT.AC_CODE=A.AC_CODE AND "
        mysql = mysql & " VT.VOU_DT >= '" & Format(FROM_DT.Value, "yyyy/MM/dd") & "' AND VT.VOU_DT <= '" & Format(TO_DT.Value, "yyyy/MM/dd") & "'"
        mysql = mysql & " AND V.VOU_TYPE IN (" & VOU_TYPE & ")"
        mysql = mysql & " ORDER BY Vt.VOU_DT, Vt.VOU_NO ,VT.DR_CR,VT.VOUID"
        Set MYRS = Nothing
        Set MYRS = New ADODB.Recordset
        MYRS.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not MYRS.EOF Then
            Set RDCREPO = Nothing
            If GETMAIN.PRINTTOGGLE.Checked = True Then
                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "VCHLST_D.RPT", 1)
            Else
                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "VCHLST_W.RPT", 1)
            End If

            RDCREPO.DiscardSavedData
            RDCREPO.Database.SetDataSource MYRS

            'RDCREPO.FormulaFields.GetItemByName("TITLE").text = "'" & ACC_LIST.text & "' & ' Voucher List'"
            RDCREPO.FormulaFields.GetItemByName("ORG").text = "'" & GCompanyName & "'"
            RDCREPO.FormulaFields.GetItemByName("ADD1").text = "'" & MAdd1 & "'"
            RDCREPO.FormulaFields.GetItemByName("ADD2").text = "'" & GCCity & "'"
            CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
            CRViewer1.ReportSource = RDCREPO
            CRViewer1.Visible = True
            CRViewer1.ViewReport
        Else
            MsgBox "Transaction does not exists ", vbInformation, "Error"
        End If

    Else        ''CASH/BANK BOOK
        DR_AMT = 0: CR_AMT = 0: COUNT_LOOP = 0: PAMT_AMT = 0: REPT_AMT = 0: LOp_Bal = 0
        Call BBOOKREC
        AccRec.MoveFirst
        AccRec.Find "NAME='" & DataCombo1.BoundText & "'", , adSearchForward
        If Not AccRec.EOF Then
            MAc_Code = AccRec!ac_code & ""
            LOp_Bal = Val(AccRec!OP_BAL & "")
        End If
        LAMT = 0
        LAMT = Net_DrCr(CStr(MAc_Code), CStr(FROM_DT.Value))
        LOp_Bal = LOp_Bal + LAMT
        If LOp_Bal <> 0 Then
            LOp_Bal = LOp_Bal * -1
            BOOKREC.AddNew
            BOOKREC!VOU_NO = "Opening"
            BOOKREC!VOU_TYPE = vbNullString
            BOOKREC!VOU_DT = Format(FROM_DT.Value, "yyyy/MM/dd")
            BOOKREC!AC_NAME = ""
            If LOp_Bal > 0 Then
                BOOKREC!AMOUNT = LOp_Bal
                BOOKREC!DR_CR = "C"
            Else
                BOOKREC!AMOUNT = LOp_Bal
                BOOKREC!DR_CR = "D"
            End If
            
            BOOKREC!OP_BAL = LOp_Bal
            BOOKREC!CL_BAL = LOp_Bal
            BOOKREC!NARRATION = NARR_CHNO
            BOOKREC!CHQ_NO = vbNullString
            BOOKREC!CHQ_DT = vbNullString
            BOOKREC.Update
            NEXT_OP_BAL = LOp_Bal
        End If
        
        Set MYRS = Nothing
        mysql = "SELECT A.NAME AS ACC, A.AC_CODE, VOU.VOU_DT AS DT, VOU.VOU_TYPE AS V_TYPE, VOU.VOU_PR AS PR, VOU.VOU_NO, VT.NARRATION,"
        mysql = mysql & " VT.VOU_TYPE, VT.VOU_NO AS VOU_NO, VT.VOU_DT, VT.DR_CR AS DR_CR, VT.AC_CODE AS ACODE, VT.AMOUNT AS AMT, VT.CHEQUE_NO, "
        mysql = mysql & " VT.CHEQUE_DT FROM ACCOUNTM AS A, VOUCHER AS VOU, VCHAMT AS VT WHERE A.COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND  VT.VOU_ID IN (SELECT VOU_ID FROM VOUCHER WHERE COMPCODE =" & GCompCode & " AND CASHCODE ='" & MAc_Code & "' AND VOU_TYPE<>'M'  AND VOU_DT >= '" & Format(FROM_DT.Value, "yyyy/MM/dd") & "' AND VOU_DT <= '" & Format(TO_DT.Value, "yyyy/MM/dd") & "') "
        mysql = mysql & " AND VT.ACCID = A.ACCID AND VT.AC_CODE <> '" & MAc_Code & "' AND VT.VOU_ID=VOU.VOU_ID "
        mysql = mysql & " ORDER BY VT.VOU_DT, VT.VOUID"
        Set MYRS = Nothing
        Set MYRS = New ADODB.Recordset
        MYRS.Open mysql, Cnn, adOpenStatic, adLockReadOnly

        If Not MYRS.EOF Then
            GETMAIN.ProgressBar1.Value = 0
            GETMAIN.ProgressBar1.Visible = True
            GETMAIN.ProgressBar1.Max = MYRS.RecordCount + Val(1)

            DATE1 = MYRS!DT
            NEXT_OP_BAL = LOp_Bal
            'If Check2.Value = 0 Then
                While Not MYRS.EOF
                    If UCase(MYRS!DR_CR) = "C" Then
                        MDrCr = "C"
                    Else
                        MDrCr = "D"
                    End If
                    BOOKREC.AddNew
                    BOOKREC!VOU_NO = MYRS!VOU_NO
                    BOOKREC!DR_CR = MDrCr
                    BOOKREC!VOU_TYPE = MYRS!VOU_TYPE
                    BOOKREC!VOU_DT = Format(MYRS!DT, "yyyy/MM/dd")
                    BOOKREC!AC_NAME = MYRS!ac_code + " " + MYRS!ACC
                    BOOKREC!AMOUNT = IIf(MDrCr = "D", MYRS!AMT * -1, MYRS!AMT)
                    BOOKREC!OP_BAL = NEXT_OP_BAL
                    BOOKREC!CL_BAL = NEXT_OP_BAL + IIf(MDrCr = "D", MYRS!AMT * -1, MYRS!AMT)
                    BOOKREC!NARRATION = MYRS!NARRATION
                    BOOKREC!CHQ_NO = vbNullString
                    BOOKREC!CHQ_DT = vbNullString
                    BOOKREC.Update
                    NEXT_OP_BAL = NEXT_OP_BAL + IIf(MDrCr = "D", MYRS!AMT * -1, MYRS!AMT)
                    MYRS.MoveNext
                Wend
                'Call ENTRY
            'Else
'                Dim Settle As Boolean
'                Do While Not MYRS.EOF
'                    Settle = False
'                    LExCode = MYRS!BANK_NAME
'                    LVou_Dt = MYRS!VOU_DT
'                    LVou_Type = MYRS!VOU_TYPE
'                    LAmount = 0
'                    Do While LExCode = MYRS!BANK_NAME And LVou_Dt = MYRS!VOU_DT And MYRS!VOU_TYPE = "S"
'                        Settle = True
'                        If UCase(MYRS!DR_CR) = "C" Then
'                            LAmount = LAmount + (MYRS!AMT)
'                        Else
'                            LAmount = LAmount + (MYRS!AMT) * -1
'                        End If
'                        MYRS.MoveNext
'                        If MYRS.EOF Then Exit Do
'                    Loop
'                    If Settle = True Then
'                        If LAmount >= 0 Then
'                            MDrCr = "C"
'                        Else
'                            MDrCr = "D"
'                        End If
'                        BOOKREC.AddNew
'                        BOOKREC!VOU_NO = "SETTLMENT"
'                        BOOKREC!DR_CR = MDrCr
'                        BOOKREC!VOU_TYPE = LVou_Type
'                        BOOKREC!VOU_DT = Format(LVou_Dt, "yyyy/MM/dd")
'                        BOOKREC!AC_NAME = LExCode & " Settlement"
'                        BOOKREC!AMOUNT = LAmount
'                        BOOKREC!OP_BAL = NEXT_OP_BAL
'                        BOOKREC!CL_BAL = NEXT_OP_BAL + LAmount
'                        BOOKREC!NARRATION = NARR_CHNO
'                        BOOKREC!CHQ_NO = vbNullString
'                        BOOKREC!CHQ_DT = vbNullString
'                        BOOKREC.Update
'                        NEXT_OP_BAL = NEXT_OP_BAL + LAmount
'                    Else
'                        If MYRS.EOF Then Exit Do
'                        BOOKREC.AddNew
'                        BOOKREC!VOU_NO = MYRS!VOU_NO
'                        BOOKREC!DR_CR = MYRS!DR_CR
'                        BOOKREC!VOU_TYPE = MYRS!VOU_TYPE
'                        BOOKREC!VOU_DT = Format(MYRS!DT, "yyyy/MM/dd")
'                        BOOKREC!AC_NAME = MYRS!ACC
'                        BOOKREC!AMOUNT = IIf(MYRS!DR_CR = "D", MYRS!AMT * -1, MYRS!AMT)
'                        BOOKREC!OP_BAL = NEXT_OP_BAL
'                        BOOKREC!CL_BAL = NEXT_OP_BAL + IIf(MYRS!DR_CR = "D", MYRS!AMT * -1, MYRS!AMT)
'                        BOOKREC!NARRATION = MYRS!NARRATION
'                        BOOKREC!CHQ_NO = vbNullString
'                        BOOKREC!CHQ_DT = vbNullString
'                        BOOKREC.Update
'                        NEXT_OP_BAL = NEXT_OP_BAL + IIf(MYRS!DR_CR = "D", MYRS!AMT * -1, MYRS!AMT)
'                        MYRS.MoveNext
'                    End If
'                    If MYRS.EOF Then Exit Do
'                Loop
'            End If
            GETMAIN.PERLBL.Visible = False
            Set MYRS = Nothing
            Set RDCREPO = Nothing
            Set RDCREPO = New CRAXDRT.report
            If (MFormat = "Cash Book Ledger Format") Then
                mysql = "SELECT * FROM TEMPBOOKS WHERE COMPCODE =" & GCompCode & " "
                
                
                Set RDCREPO = RDCAPP.OpenReport(GReportPath & "CASHbook_DW.RPT", 1)
            ElseIf (MFormat = "Bank Book Ledger Format") Then
               mysql = "SELECT * FROM TEMPBOOKS "
                If Check1.Value = 1 Then
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "BANK2_W.RPT", 1)
                Else
                    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "BANKBOOK_dW.RPT", 1)
                End If
            End If

            Set MYRS = Nothing
            Set MYRS = New ADODB.Recordset
            Set MYRS = BOOKREC.Clone
        
            RDCREPO.DiscardSavedData
            RDCREPO.Database.SetDataSource MYRS

            'RDCREPO.FormulaFields.GetItemByName("ACCNAME").text = "' " & Mid(MFormate, 1, 9) & " for " & ACC_LIST.text & "'"
            RDCREPO.FormulaFields.GetItemByName("ORG").text = "' " & GCompanyName & "'"
            RDCREPO.FormulaFields.GetItemByName("ADD1").text = "' For " & CStr(FROM_DT.Value) & " to " & CStr(TO_DT.Value) & " '"

            CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)

            CRViewer1.ReportSource = RDCREPO
            CRViewer1.Visible = True
            CRViewer1.ViewReport

            GETMAIN.ProgressBar1.Visible = False
            GETMAIN.PERLBL = vbNullString
            Set Rec = Nothing

         Else
            MsgBox "Transaction does not exists!", vbInformation, "Error"
            If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
        End If
    End If

    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    OK_CMD.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
Error1:
    MsgBox err.Description, vbCritical, err.HelpFile
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
    GETMAIN.ProgressBar1.Visible = False
    GETMAIN.PERLBL = vbNullString
    Screen.MousePointer = 0: OK_CMD.Enabled = True
End Sub
Private Sub Form_Load()
    Label2.Caption = vbNullString
    Label6.Caption = MFormat

    Call Get_Selection(12)

    TO_DT.MinDate = GFinBegin:      TO_DT.MaxDate = GFinEnd
    FROM_DT.MinDate = GFinBegin:    FROM_DT.MaxDate = GFinEnd
    DTPicker1.MinDate = GFinBegin:  DTPicker1.MaxDate = GFinEnd
    DTPicker2.MinDate = GFinBegin:  DTPicker2.MaxDate = GFinEnd

    CRViewer1.ZOrder

    CNNERR = False
    If MFormat = "Voucher List" Then
        Frame1.Visible = True
        Check1.Visible = False
        Inv_fram.Visible = False
        delpen_fram.Visible = False
        FROM_DT.Value = GFinBegin
        TO_DT.Value = Date
        Exit Sub

    Else 'Cash book , Bank book
        Frame3.Visible = False
        Label1.Visible = False
        Label5.Visible = True
        
        If MFormat = "Cash Book Regular Format" Or MFormat = "Cash Book Ledger Format" Then
            mysql = "SELECT AC_CODE,NAME,OP_BAL FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND GCODE = 10"
            Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
            AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            Check1.Caption = "Day Total Required"
            Check1.Value = 0
        ElseIf MFormat = "Bank Book Regular Format" Or MFormat = "Bank Book Ledger Format" Then
            mysql = "SELECT AC_CODE,NAME,OP_BAL FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " AND GCODE = 11"
            Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
            AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            Check1.Caption = "Day Total Not Required"
            Check1.Visible = True
        End If
        If Not AccRec.EOF Then AccRec.MoveFirst
        DataCombo1.Visible = True
        Set DataCombo1.RowSource = AccRec
        DataCombo1.ListField = "NAME"
        DataCombo1.ListField = "NAME"
        DataCombo1.Visible = True
        DataCombo1.text = AccRec!NAME
        Frame1.Visible = True
        Inv_fram.Visible = False
        delpen_fram.Visible = False
        FROM_DT.Value = GFinBegin
        TO_DT.Value = Date
    End If
    
    Me.Caption = Mid(MFormat, 1, 9)
End Sub
Sub ENTRY()
    Dim FrDt As Variant
    Dim CHEQUE_NO As String
    Dim CHEQUE_DT As String
    Dim I As Integer
    Dim MDrCr As String
    Dim MacName  As String
    For I = 1 To COUNT_LOOP
        MYRS.MovePrevious
    Next I
    CL_BAL = NEXT_OP_BAL - REPT_AMT + PAMT_AMT
    For I = 1 To COUNT_LOOP
        MDrCr = MYRS!DR_CR
        MacName = MYRS!ACC & vbNullString

        Set Rec = Nothing
        Set Rec = New ADODB.Recordset
        Set Rec = MYRS.Clone
        Rec.Filter = "VOU_NO='" & MYRS!VOU_NO & "'"
        Rec.MoveLast
        MacName = Rec!ACC & vbNullString

        NARR_CHNO = vbNullString
        If MYRS!V_TYPE = "S" Or MYRS!V_TYPE = "P" Then
            NARR_CHNO = MYRS!NARRATION & vbNullString
        Else
            NARR_CHNO = MYRS!NARRATION & vbNullString
        End If
        CHEQUE_NO = MYRS!CHEQUE_NO & vbNullString
        CHEQUE_DT = MYRS!CHEQUE_DT & vbNullString

        If MFormat = "Cash Book Regular Format" Or MFormat = "Cash Book Ledger Format" Then
            If Check1.Value = 0 Then
                If MYRS!V_TYPE = "S" Then
                    Set MYRS_CASH = Nothing
                    Set MYRS_CASH = New ADODB.Recordset
                    MYRS_CASH.Open "SELECT * FROM TEMPBOOKS WHERE COMPCODE=" & GCompCode & " AND DT='" & Format(MYRS!DT, "yyyy/MM/dd") & "' AND VOU_TYPE='CASH SALES' AND DOCTYPE='C'", Cnn, 2, 3
                    If Not MYRS_CASH.EOF Then
                        Cnn.Execute "UPDATE TEMPBOOKS SET AMT = AMT+" & Val(MYRS!AMT) & " WHERE COMPCODE = " & GCompCode & " AND DT = '" & Format(MYRS!DT, "yyyy/MM/dd") & "' AND ACC= 'CASH SALES' AND VOU_TYPE='C'"
                    Else
                        mysql = "INSERT INTO TEMPBOOKS(VOU_TYPE, DT, ACC, AMT, OP_BAL, CL_BAL, NARRATION, LOC_CODE,COMPCODE) VALUES('" & MDrCr & "','" & Format(MYRS!DT, "yyyy/MM/dd") & "','CASH SALES'," & IIf(MDrCr = "D", MYRS!AMT * -1, MYRS!AMT) & "," & NEXT_OP_BAL & "," & CL_BAL & ",'CASH SALES','" & MYRS!VOU_TYPE & "'," & GCompCode & ")"
                        Cnn.Execute mysql
                    End If
                Else
                    mysql = "INSERT INTO TEMPBOOKS(VOU_NO, VOU_TYPE, DT, ACC, AMT, OP_BAL, CL_BAL, NARRATION,COMPCODE) VALUES('" & MYRS!VOU_NO & "','" & MDrCr & "','" & Format(MYRS!DT, "yyyy/MM/dd") & "','" & MYRS!ACC & "'," & IIf(MDrCr = "D", MYRS!AMT * -1, MYRS!AMT) & "," & NEXT_OP_BAL & "," & CL_BAL & ",'" & NARR_CHNO & "'," & GCompCode & ")"
                    Cnn.Execute mysql
                End If
            Else
                Cnn.Execute "INSERT INTO TEMPBOOKS(VOU_NO, VOU_TYPE, DT, ACC, AMT, OP_BAL, CL_BAL, NARRATION,COMPCODE) VALUES('" & MYRS!VOU_NO & "','" & MDrCr & "','" & Format(MYRS!DT, "yyyy/MM/dd") & "','" & MYRS!ACC & "'," & IIf(MDrCr = "D", MYRS!AMT * -1, MYRS!AMT) & "," & NEXT_OP_BAL & "," & CL_BAL & ",'" & NARR_CHNO & "'," & GCompCode & ")"
            End If

        ElseIf MFormat = "Bank Book Regular Format" Or MFormat = "Bank Book Ledger Format" Then
             BOOKREC.AddNew
             BOOKREC!VOU_NO = MYRS!VOU_NO
             BOOKREC!DR_CR = MDrCr
             BOOKREC!VOU_TYPE = MYRS!VOU_TYPE
             BOOKREC!VOU_DT = Format(MYRS!DT, "yyyy/MM/dd")
             BOOKREC!AC_NAME = MYRS!ACC
             BOOKREC!AMOUNT = IIf(MDrCr = "D", MYRS!AMT * -1, MYRS!AMT)
             BOOKREC!OP_BAL = NEXT_OP_BAL
             BOOKREC!CL_BAL = CL_BAL
             BOOKREC!NARRATION = NARR_CHNO
             BOOKREC!CHQ_NO = CHEQUE_NO
             BOOKREC!CHQ_DT = CHEQUE_DT
             BOOKREC!OPENBAL = OP_BAL
             BOOKREC.Update
        End If
        MYRS.MoveNext
FLAG1:
        
    Next I

    Call PERCENTAGE(GETMAIN.ProgressBar1.Max, GETMAIN.ProgressBar1.Value)
    GETMAIN.ProgressBar1.Value = Val(GETMAIN.ProgressBar1.Value) + Val(1)

    On Error GoTo Error1
    

    COUNT_LOOP = -1: REPT_AMT = 0: PAMT_AMT = 0: NEXT_OP_BAL = CL_BAL: CL_BAL = 0

    DATE1 = MYRS!DT
    MYRS.MovePrevious
Error1: If err.Number = 3021 Then
        End If
End Sub
Sub RecSet()    ''Sub Routine to Open Recordset Without Table
    If MFormat = "Voucher List" Then
        Set Rec = Nothing
        Set Rec = New ADODB.Recordset
        Rec.Fields.Append "VOU_DT", adDate, , adFldIsNullable
        Rec.Fields.Append "VOU_TYPE", adVarChar, 2, adFldIsNullable
        Rec.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
        Rec.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
        Rec.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
        Rec.Fields.Append "NARRATION", adVarChar, 100, adFldIsNullable
        Rec.Fields.Append "CL_BAL", adDecimal, , adFldIsNullable
        Rec.Fields.Append "G_CODE", adInteger, , adFldIsNullable
        Rec.Fields.Append "AMOUNT", adNumeric, 4, adFldIsNullable
        Rec.Fields.Append "VOU_NO", adVarChar, 20, adFldIsNullable
        Rec.Fields.Append "VOU_ID", adNumeric, , adFldIsNullable
        Rec.Fields.Append "BILL_NO", adVarChar, 15, adFldIsNullable
        Rec.Fields.Append "BILL_DT", adDate, , adFldIsNullable
        
        Rec.Open , , adOpenKeyset, adLockOptimistic
    End If
End Sub


Sub BBOOKREC()    ''Sub Routine to Open Recordset Without Table
    Set BOOKREC = Nothing
    Set BOOKREC = New ADODB.Recordset
    BOOKREC.Fields.Append "VOU_NO", adVarChar, 20, adFldIsNullable
    BOOKREC.Fields.Append "VOU_DT", adDate, , adFldIsNullable
    BOOKREC.Fields.Append "VOU_TYPE", adVarChar, 2, adFldIsNullable
    BOOKREC.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
    BOOKREC.Fields.Append "AC_NAME", adVarChar, 100, adFldIsNullable
    BOOKREC.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
    BOOKREC.Fields.Append "NARRATION", adVarChar, 500, adFldIsNullable
    BOOKREC.Fields.Append "CHQ_NO", adVarChar, 7, adFldIsNullable
    BOOKREC.Fields.Append "CHQ_DT", adVarChar, 50, adFldIsNullable
    BOOKREC.Fields.Append "CL_BAL", adDouble, , adFldIsNullable
    BOOKREC.Fields.Append "OP_BAL", adDouble, , adFldIsNullable
    BOOKREC.Fields.Append "AMOUNT", adDouble, 4, adFldIsNullable
    BOOKREC.Open , , adOpenKeyset, adLockOptimistic
End Sub

