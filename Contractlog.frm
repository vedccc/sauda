VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form Contractlog 
   Caption         =   "Contract Log"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   5295
      Begin VB.CommandButton OkCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Ok"
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CancelCmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin vcDateTimePicker.vcDTP datefrom 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   42079.7973611111
      End
      Begin vcDateTimePicker.vcDTP dateto 
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   42079.7973611111
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
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
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   1335
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
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
   Begin VB.Label Label27 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contract Log Report"
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
      TabIndex        =   3
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "Contractlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecRpt As ADODB.Recordset

Private Sub CancelCmd_Click()
   Unload Me
End Sub


Private Sub Form_Load()
    Label27.Caption = MFormat + " Report"
    datefrom.Value = Date
    dateto.Value = Date
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        CRViewer1.Visible = False
        Cancel = 1
    Else
        Unload Me
    End If
End Sub

Private Sub OkCmd_Click()
    On Error GoTo err1
     
    Dim MRec As ADODB.Recordset
                   
    If MFormat = "Voucher Log" Then
        mysql = " EXEC GET_VCHAMT_LOG '" & GCompCode & "','" & Format(datefrom.Value, "YYYY/MM/DD") & "','" & Format(dateto.Value, "YYYY/MM/DD") & "'"
    'ElseIf MFormat = "Contract Log" Then
    '    mysql = " EXEC GET_CTR_D_LOG '" & GCompCode & "','" & Format(datefrom.Value, "YYYY/MM/DD") & "','" & Format(dateto.Value, "YYYY/MM/DD") & "'"
    End If
    
    
    Set MRec = Nothing
    Set MRec = New ADODB.Recordset
    MRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not MRec.EOF Then
                        
            Me.MousePointer = Val(11): OkCmd.Enabled = False
                        
            Set RecRpt = Nothing
            Set RecRpt = New ADODB.Recordset
            If MFormat = "Contract Log" Then
                RecRpt.Fields.Append "CompCode", adVarChar, 15, adFldIsNullable
                RecRpt.Fields.Append "CONSNO", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "CONNO", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "PARTY", adVarChar, 150, adFldIsNullable
                RecRpt.Fields.Append "SAUDA", adVarChar, 150, adFldIsNullable
                RecRpt.Fields.Append "ITEMCODE", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "Condate", adDate, , adFldIsNullable
                RecRpt.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
                RecRpt.Fields.Append "QTY", adDouble, , adFldIsNullable
                RecRpt.Fields.Append "RATE", adDouble, , adFldIsNullable
                RecRpt.Fields.Append "BILLNO", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "ROWNO1", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "loginuser", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "datetm", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "tran", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "type", adVarChar, 50, adFldIsNullable
                RecRpt.Open , , adOpenKeyset, adLockBatchOptimistic
                                                                       
                Do While Not MRec.EOF
                
                    RecRpt.AddNew
                    RecRpt!CompCode = MRec!CompCode
                    RecRpt!CONSNO = MRec!CONSNO
                    RecRpt!CONNO = MRec!CONNO
                    RecRpt!PARTY = MRec!PARTY
                    RecRpt!Sauda = MRec!Sauda
                    RecRpt!ITEMCODE = MRec!ITEMCODE
                    RecRpt!Condate = MRec!Condate
                    
                    RecRpt!CONTYPE = MRec!CONTYPE
                    RecRpt!QTY = MRec!QTY
                    RecRpt!Rate = MRec!Rate
                    RecRpt!BILLNO = MRec!BILLNO
                    RecRpt!ROWNO1 = MRec!ROWNO1
                    RecRpt!loginuser = MRec!loginuser
                    RecRpt!datetm = MRec!datetm
                    RecRpt!tran = MRec!tran
                    RecRpt!Type = MRec!rowtype
                    
                    RecRpt.Update
                    MRec.MoveNext
                Loop
            
            ElseIf MFormat = "Voucher Log" Then
                RecRpt.Fields.Append "CompCode", adVarChar, 15, adFldIsNullable
                RecRpt.Fields.Append "VOU_NO", adVarChar, 25, adFldIsNullable
                RecRpt.Fields.Append "VOU_TYPE", adVarChar, 5, adFldIsNullable
                RecRpt.Fields.Append "VOU_DT", adDate, , adFldIsNullable
                RecRpt.Fields.Append "DR_CR", adVarChar, 1, adFldIsNullable
                RecRpt.Fields.Append "AC_CODE", adVarChar, 150, adFldIsNullable
                RecRpt.Fields.Append "AMOUNT", adDouble, , adFldIsNullable
                RecRpt.Fields.Append "CHEQUE_NO", adVarChar, 25, adFldIsNullable
                RecRpt.Fields.Append "CHEQUE_DT", adVarChar, 25, adFldIsNullable
                RecRpt.Fields.Append "NARRATION", adVarChar, 150, adFldIsNullable
                RecRpt.Fields.Append "BRANCH", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "BANK_NAME", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "vouid", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "invno", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "EXCODE", adVarChar, 25, adFldIsNullable
                RecRpt.Fields.Append "IAMOUNT", adDouble, , adFldIsNullable
                RecRpt.Fields.Append "CURRRATE", adDouble, , adFldIsNullable
                RecRpt.Fields.Append "EXID", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "VOU_ID", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "ACCID", adInteger, , adFldIsNullable
                RecRpt.Fields.Append "loginuser", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "datetm", adVarChar, 50, adFldIsNullable
                RecRpt.Fields.Append "tran", adVarChar, 15, adFldIsNullable
                RecRpt.Fields.Append "rowtype", adVarChar, 10, adFldIsNullable
                RecRpt.Open , , adOpenKeyset, adLockBatchOptimistic
                                                                       
                Do While Not MRec.EOF
                    RecRpt.AddNew
                    RecRpt!CompCode = MRec!CompCode
                    RecRpt!VOU_NO = MRec!VOU_NO
                    RecRpt!VOU_TYPE = MRec!VOU_TYPE
                    RecRpt!VOU_DT = MRec!VOU_DT
                    RecRpt!DR_CR = MRec!DR_CR
                    RecRpt!AC_CODE = MRec!AC_CODE
                    RecRpt!AMOUNT = MRec!AMOUNT
                    RecRpt!CHEQUE_NO = MRec!CHEQUE_NO
                    RecRpt!CHEQUE_DT = MRec!CHEQUE_DT
                    RecRpt!NARRATION = MRec!NARRATION
                    RecRpt!BRANCH = MRec!BRANCH
                    RecRpt!BANK_NAME = MRec!BANK_NAME
                    RecRpt!vouid = MRec!vouid
                    RecRpt!invno = MRec!invno
                    RecRpt!excode = MRec!excode
                    RecRpt!IAMOUNT = MRec!IAMOUNT
                    RecRpt!CURRRATE = MRec!CURRRATE
                    RecRpt!EXID = MRec!EXID
                    RecRpt!VOU_ID = MRec!VOU_ID
                    RecRpt!ACCID = MRec!ACCID
                    RecRpt!loginuser = MRec!loginuser
                    RecRpt!datetm = MRec!datetm
                    RecRpt!tran = MRec!tran
                    RecRpt!rowtype = MRec!rowtype
                    
                    RecRpt.Update
                    MRec.MoveNext
                Loop
            
            End If
            Set MRec = Nothing
        
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        Set MRec = RecRpt.Clone
        
        mysql = "'" & MFormat & " report " & datefrom.Value & " to " & dateto.Value & "'"
        If MFormat = "Contract Log" Then
            Set RDCREPO = RDCAPP.OpenReport(GReportPath & "ContractLog.rpt", 1)
        ElseIf MFormat = "Voucher Log" Then
            'mysql = "'Contract log report " & datefrom.Value & " to " & dateto.Value & "'"
            Set RDCREPO = RDCAPP.OpenReport(GReportPath & "VoucherLog.rpt", 1)
        End If
        
        RDCREPO.DiscardSavedData
        RDCREPO.Database.SetDataSource MRec
        RDCREPO.FormulaFields.GetItemByName("COMPANY").text = "'" & GCompanyName & "'"
        RDCREPO.FormulaFields.GetItemByName("TITLE").text = mysql
        CRViewer1.ZOrder
        CRViewer1.Move 0, 0, CInt(GETMAIN.Width - 100), CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1100)
        CRViewer1.Visible = True
        CRViewer1.ReportSource = RDCREPO
        CRViewer1.Zoom 1
        CRViewer1.ViewReport
        Me.MousePointer = 0
        OkCmd.Enabled = True
        Exit Sub
    Else
        MsgBox "No Records Found", vbInformation
    End If
    
err1:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    End If
    Me.MousePointer = 0
    OkCmd.Enabled = True
End Sub


