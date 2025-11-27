VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SELCOMP 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3585
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   11280
   ControlBox      =   0   'False
   ForeColor       =   &H80000016&
   Icon            =   "SELCOMP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16744576
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   21
      TabAction       =   1
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "NAME"
         Caption         =   "Firm Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d.MMMyyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FINBEGIN"
         Caption         =   "                From"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d.MMMyyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FINEND"
         Caption         =   "             To"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d.MMMyyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CITY"
         Caption         =   "City"
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
         MarqueeStyle    =   3
         RecordSelectors =   0   'False
         BeginProperty Column00 
            DividerStyle    =   3
            ColumnWidth     =   6449.953
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   3
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   3
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   3
            Object.Visible         =   0   'False
            ColumnWidth     =   1514.835
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3615
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
   End
End
Attribute VB_Name = "SELCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_DblClick()
    Call DBGRIDCLICK
    Text1.SetFocus
End Sub
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Call DataGrid1_DblClick
End Sub

Private Sub Form_Activate()
    If SelComp_Ado.EOF Or SelComp_Ado.BOF Then
        MsgBox "Create company first AND restart the program.", vbInformation, "Information"
        Call LogOff
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    DataGrid1.RowHeight = Val(550)
    If Not SelComp_Ado.EOF Then
        Set DataGrid1.DataSource = SelComp_Ado
        DataGrid1.ReBind
        DataGrid1.Refresh
        Call Get_Selection(12)
    End If
End Sub
Private Sub Form_Paint()
    GETMAIN.StatusBar1.Panels(1).text = vbNullString
    If SelComp_Ado.RecordCount = Val(1) Then Sendkeys "{ENTER}"
End Sub
Sub Form_Unload(Cancel As Integer)
On Error GoTo err1
    Unload SELCOMP
    Exit Sub
err1:
    If err.Number = Val(91) Then ''IF CNN ARE NOT INIT
        Resume Next
    Else
        MsgBox err.Description, vbCritical, "Error : " & err.Number
    End If
End Sub
Sub DBGRIDCLICK()
    On Error GoTo Error1
    Call CompanySelection(Val(SelComp_Ado!COMPCODE))
    GETMAIN.TRANS.Enabled = True
    GETMAIN.report.Enabled = True
    GETMAIN.master.Enabled = True
    GETMAIN.Database.Enabled = True
    GETMAIN.utilities.Enabled = True
    GETMAIN.StatusBar1.Panels(5).text = GCompanyName
    MenuOptfrm.Show
Error1:

'If err.Number <> 0 Then
'MsgBox err.Description
'End If
End Sub
Private Sub Text1_GotFocus()
    Unload SELCOMP
End Sub
