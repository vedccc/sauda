VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_VCH 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14445
   Icon            =   "FRM_VCH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   14445
   WindowState     =   2  'Maximized
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13500
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Voucher"
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
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   13725
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   13485
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRM_VCH.frx":000C
         Left            =   5160
         List            =   "FRM_VCH.frx":0031
         TabIndex        =   3
         Top             =   165
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DComboAcc 
         Height          =   360
         Left            =   7440
         TabIndex        =   4
         Top             =   165
         Width           =   3255
         _ExtentX        =   5741
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
      Begin vcDateTimePicker.vcDTP vcDTPToDate 
         Height          =   360
         Left            =   2640
         TabIndex        =   2
         Top             =   165
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   43156.5303240741
      End
      Begin VB.CommandButton CmdFrameCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11880
         TabIndex        =   6
         Top             =   165
         Width           =   1000
      End
      Begin VB.CommandButton CmdFrameOk 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10800
         TabIndex        =   5
         Top             =   165
         Width           =   1000
      End
      Begin vcDateTimePicker.vcDTP VcDtpVouDate 
         Height          =   360
         Left            =   720
         TabIndex        =   1
         Top             =   165
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37680.7163888889
      End
      Begin VB.Frame pr_frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "VouType"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "A/c"
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
         Left            =   6960
         TabIndex        =   18
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   218
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   218
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   13485
      Begin VB.CommandButton CmdFormOk 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10680
         TabIndex        =   17
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton CmdFormCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11760
         TabIndex        =   16
         Top             =   6120
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid_Detail 
         Height          =   3495
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   -2147483641
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Voucher Amount"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "AC_NAME"
            Caption         =   "      Account Description"
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
            DataField       =   "DR_CR"
            Caption         =   "Dr/Cr"
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
            DataField       =   "AMOUNT"
            Caption         =   "     Amount"
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
         BeginProperty Column03 
            DataField       =   "NARRATION"
            Caption         =   "                          Narration"
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
            MarqueeStyle    =   2
            ScrollBars      =   3
            BeginProperty Column00 
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   3960
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   1514.835
            EndProperty
            BeginProperty Column03 
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   10649.76
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid_Voucher 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   -2147483641
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Voucher"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "VOU_NO"
            Caption         =   "                      Voucher No."
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
            DataField       =   "VOU_DT"
            Caption         =   "         Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d.MMMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "VOU_TYPE"
            Caption         =   "Vou_Type"
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
            MarqueeStyle    =   2
            ScrollBars      =   2
            BeginProperty Column00 
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   4155.024
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   3
               Locked          =   -1  'True
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   7695
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   13500
   End
End
Attribute VB_Name = "FRM_VCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte:          Dim VCHREC As ADODB.Recordset
Dim VouRec As ADODB.Recordset:    Dim VouDRec As ADODB.Recordset
Dim LFAc_Code As String:          Dim AccRec As ADODB.Recordset
Private Sub CmdFrameCancel_Click()
    Call Get_Selection(5)
    LFAc_Code = vbNullString
    
    If GVoucherFormat = "Y" Then 'new
        Call VouFrmNew.CLEAR_SCREEN
        Call Get_Selection(5)
    Else
        Call VouFrm.CLEAR_SCREEN
        Call Get_Selection(5)
    End If
    Unload Me
End Sub
Private Sub CmdFormOk_Click()
    Dim MVou_No As String
    Dim MVou_DT As Date
    MVou_No = DataGrid_voucher.Columns(0).text
        
    Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
    mysql = "SELECT V.VOU_NO,V.VOU_DT FROM VOUCHER AS V WHERE V.COMPCODE=" & GCompCode & " AND V.VOU_NO='" & MVou_No & "'"
    VouRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not VouRec.EOF Then
        MVou_DT = VouRec!VOU_DT

        If (GSysLockDt >= MVou_DT) Then
            If MsgBox("System is locked till date " & GSysLockDt & vbNewLine & "Do you still want to view document?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
        If GVoucherFormat = "Y" Then 'new
            VouFrmNew.Fb_Press = FRM_VCH.Fb_Press
            VouFrmNew.Show
            If VouFrmNew.VOUCHER_ACCESS(VouRec!VOU_NO) Then
                If FRM_VCH.Fb_Press = 2 Then
                    VouFrmNew.Frame6.Enabled = True
                    If VouFrmNew.TXT_NARR.Visible = True Then VouFrmNew.TXT_NARR.SetFocus
                ElseIf FRM_VCH.Fb_Press = 3 Then
                    If (GSysLockDt < MVou_DT) Then
                        If MsgBox("  Confirm Deletion ?        ", vbYesNo + vbQuestion, "Confirmation") = 6 Then
                            Call VouFrmNew.Delete_Entry
                            Unload VouFrmNew
                            Frame1.Enabled = True: VcDtpVouDate.SetFocus: Frame2.Visible = False: CmdFormOk.Visible = False: CmdFormCancel.Visible = False
                            Exit Sub
                        Else
                            Unload VouFrmNew
                            Exit Sub
                        End If
                        'Call VouFrmNew.CLEAR_SCREEN
                        'Call Get_Selection(5)
                    End If
                End If
            End If
        Else
            VouFrm.Fb_Press = FRM_VCH.Fb_Press
            VouFrm.Show
            If VouFrm.VOUCHER_ACCESS(VouRec!VOU_NO) Then
                If FRM_VCH.Fb_Press = 2 Then
                    VouFrm.Frame6.Enabled = True
                    If VouFrm.TXT_NARR.Visible = True Then VouFrm.TXT_NARR.SetFocus
                ElseIf FRM_VCH.Fb_Press = 3 Then
                    If (GSysLockDt < MVou_DT) Then
                        If MsgBox("  Confirm Deletion ?        ", vbYesNo + vbQuestion, "Confirmation") = 6 Then
                            Call VouFrm.Delete_Entry
                            Unload VouFrm
                            Frame1.Enabled = True: VcDtpVouDate.SetFocus: Frame2.Visible = False: CmdFormOk.Visible = False: CmdFormCancel.Visible = False
                            Exit Sub
                        Else
                            Unload VouFrm
                            Exit Sub
                        End If
                        'Call VouFrm.CLEAR_SCREEN
                        'Call Get_Selection(5)
                    End If
                End If
            End If
        End If
        Unload FRM_VCH
    End If
End Sub
Private Sub CmdFormCancel_Click()
    Frame1.Enabled = True
    VcDtpVouDate.SetFocus
    Frame2.Visible = False
    CmdFormOk.Visible = False
    CmdFormCancel.Visible = False
End Sub
Private Sub Combo1_GotFocus()
Sendkeys "%{DOWN}"
End Sub
Private Sub DataGrid_voucher_Click()
    Update_DataGrid
End Sub
Private Sub DataGrid_voucher_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Update_DataGrid
    End If
End Sub
Private Sub DataGrid_voucher_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Update_DataGrid
End Sub
Private Sub DComboAcc_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub DComboAcc_Validate(Cancel As Boolean)
If LenB(DComboAcc.BoundText) > 1 Then
    LFAc_Code = Get_AccountMCode(DComboAcc.BoundText)
    If LenB(LFAc_Code) < 1 Then
        DComboAcc.BoundText = vbNullString
        MsgBox "Invalid Acccount Selctted "
        Cancel = True
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call Get_Selection(5)
        Unload Me
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Form_Load()
    Combo1.ListIndex = 8
    Frame2.ZOrder
    Frame1.Visible = True:              Frame2.Visible = False
    VcDtpVouDate.Value = Date:          VcDtpToDate.Value = Date
    CmdFormOk.Visible = False:          CmdFormCancel.Visible = False
    CmdFrameCancel.Visible = True
    Set AccRec = Nothing
    Set AccRec = New ADODB.Recordset
    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTM WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        Set DComboAcc.RowSource = AccRec
        DComboAcc.ListField = "NAME"
        DComboAcc.BoundColumn = "AC_CODE"
    End If
End Sub
Private Sub Form_Paint()
    GETMAIN.StatusBar1.Panels(1).text = "Voucher Details"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    GETMAIN.StatusBar1.Panels(1).text = vbNullString
End Sub
Private Sub CmdFrameOk_Click()
    LFAc_Code = vbNullString
    If LenB(DComboAcc.BoundText) > 0 Then LFAc_Code = DComboAcc.BoundText
    Set VouRec = Nothing: Set VouRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT V.VOU_NO,V.VOU_DT,V.VOU_TYPE  FROM VOUCHER AS V , VCHAMT AS VT  WHERE V.COMPCODE=" & GCompCode & "  "
    mysql = mysql & " AND V.COMPCODE =VT.COMPCODE AND V.VOU_NO=VT.VOU_NO"
    If Combo1.ListIndex >= 0 Then
        If Combo1.ListIndex = 0 Then mysql = mysql & " AND V.VOU_TYPE ='CV'"
        If Combo1.ListIndex = 1 Then mysql = mysql & " AND V.VOU_TYPE='BV'"
        If Combo1.ListIndex = 2 Then mysql = mysql & " AND V.VOU_TYPE='JV'"
        If Combo1.ListIndex = 3 Then mysql = mysql & " AND V.VOU_TYPE='S'"
        If Combo1.ListIndex = 4 Then mysql = mysql & " AND V.VOU_TYPE='H'"
        If Combo1.ListIndex = 5 Then mysql = mysql & " AND V.vou_type='K'"
        If Combo1.ListIndex = 6 Then mysql = mysql & " AND V.vou_type='M'"
        If Combo1.ListIndex = 7 Then mysql = mysql & " AND V.vou_type='F'"
    Else
        mysql = mysql & " AND V.VOU_TYPE NOT IN ('S','H','B','M')"
    End If
    mysql = mysql & " AND V.VOU_DT>='" & Format(VcDtpVouDate.Value, "yyyy/MM/dd") & "' AND V.VOU_DT<='" & Format(VcDtpToDate.Value, "yyyy/MM/dd") & "'"
    If LenB(LFAc_Code) > 1 Then mysql = mysql & " AND VT.AC_CODE ='" & LFAc_Code & "' ORDER BY V.VOU_DT,V.VOU_TYPE,V.VOU_NO"
    VouRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If VouRec.RecordCount <> 0 Then
        Set VouDRec = Nothing: Set VouDRec = New ADODB.Recordset
        mysql = "SELECT A.NAME AS AC_NAME, VT.DR_CR, VT.AMOUNT, VT.NARRATION FROM VCHAMT AS VT, ACCOUNTM AS A WHERE "
        mysql = mysql & " A.COMPCODE=" & GCompCode & " AND A.COMPCODE=VT.COMPCODE AND VT.VOU_NO='" & VouRec!VOU_NO & "'"
        mysql = mysql & " AND A.AC_CODE=VT.AC_CODE ORDER BY VT.VOUID"  'VT.VOUID
        VouDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        Set DataGrid_voucher.DataSource = VouRec
        DataGrid_voucher.ReBind
        DataGrid_voucher.Refresh
        Set DataGrid_detail.DataSource = VouDRec
        DataGrid_detail.ReBind
        DataGrid_detail.Refresh
        Frame2.Visible = True:             CmdFormOk.Visible = True
        CmdFormCancel.Visible = True:      DataGrid_voucher.Col = 0
        DataGrid_voucher.Row = 0:          DataGrid_voucher.SetFocus
    ElseIf VouRec.RecordCount = 0 Then
        MsgBox "Transaction not exist.", vbInformation, "Message"
        Frame1.Enabled = True
        Call Get_Selection(5)
    End If
End Sub
Private Sub Update_DataGrid()
     Set VouDRec = Nothing: Set VouDRec = New ADODB.Recordset
     mysql = "SELECT A.NAME AS AC_NAME, VT.DR_CR, VT.AMOUNT, VT.NARRATION FROM VCHAMT AS VT, ACCOUNTM AS A "
     mysql = mysql & " WHERE a.COMPCODE=" & GCompCode & " AND A.COMPCODE=VT.COMPCODE "
     mysql = mysql & " AND VT.VOU_NO='" & DataGrid_voucher.Columns(0).text & "' AND A.AC_CODE=VT.AC_CODE ORDER BY VT.VOUID"
     VouDRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
     Set DataGrid_detail.DataSource = VouDRec
     DataGrid_detail.ReBind
     DataGrid_detail.Refresh
End Sub
