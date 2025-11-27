VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmQueryBill1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Query On Bill Summary"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   8115
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   2775
      Left            =   7680
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
      Begin MSComctlLib.ListView ExListView 
         Height          =   2460
         Left            =   90
         TabIndex        =   4
         Top             =   120
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   4339
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   27375
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
         Left            =   15480
         TabIndex        =   9
         Top             =   840
         Value           =   -1  'True
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
         Left            =   15480
         TabIndex        =   8
         Top             =   480
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
         Left            =   18480
         TabIndex        =   11
         Top             =   720
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
         Left            =   8280
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton CmdUpd 
         Caption         =   "Update"
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
         Left            =   18480
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
         Left            =   3720
         TabIndex        =   18
         Top             =   960
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
         Left            =   15480
         TabIndex        =   7
         Top             =   120
         Width           =   2415
      End
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   780
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
         Left            =   11760
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
         Left            =   11760
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
         Left            =   6000
         TabIndex        =   0
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
         Left            =   5400
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Summary"
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
         TabIndex        =   29
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Party"
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
         Left            =   10560
         TabIndex        =   28
         Top             =   240
         Width           =   1215
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
         Left            =   11040
         TabIndex        =   27
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
         Left            =   5640
         TabIndex        =   26
         Top             =   840
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   20400
      Top             =   600
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   12255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   27375
      Begin TabDlg.SSTab SSTab1 
         Height          =   12015
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   27135
         _ExtentX        =   47863
         _ExtentY        =   21193
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
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
         TabPicture(0)   =   "FRMQUERYBILL.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "MSFdetail"
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Bill Summary with Sharing "
         TabPicture(1)   =   "FRMQUERYBILL.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "MSFRPT"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid MSFdetail 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   3
            Top             =   480
            Visible         =   0   'False
            Width           =   19575
            _ExtentX        =   34528
            _ExtentY        =   12303
            _Version        =   393216
            Cols            =   18
            FixedCols       =   0
            RowHeightMin    =   30
            BackColor       =   0
            ForeColor       =   16777215
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   11535
            Left            =   -74760
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   26655
            _ExtentX        =   47016
            _ExtentY        =   20346
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   21
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
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
         Begin MSFlexGridLib.MSFlexGrid MSFRPT 
            Height          =   6975
            Left            =   0
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   19815
            _ExtentX        =   34951
            _ExtentY        =   12303
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   30
            BackColor       =   0
            ForeColor       =   16777215
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   12585
      Left            =   120
      Top             =   1200
      Width           =   27210
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   1065
      Left            =   120
      Top             =   120
      Width           =   27210
   End
End
Attribute VB_Name = "FrmQueryBill1"
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
    Call Get_BillSum
End Sub
Private Sub Get_BillSum()

    '>>> Do not import LIVE rate on saturday and sunday
    Dim Weekdy As Integer
    Weekdy = Weekday(vcDTP1.Value)
    If Weekdy = 1 Or Weekdy = 7 Then 'Sunday, Saturday
        Exit Sub
    End If
    
    Me.MousePointer = 11
    If FlagLiveRate = "Y" And Val(DataCombo1.BoundText) = 0 And Val(DataCombo2.BoundText) = 0 Then  '>>> call live rate only if no filter
        
        If RadioMCX.Value Then
            RadioOpt = "MCX"
        ElseIf RadioNSE.Value Then
            RadioOpt = "NSE"
        ElseIf RadioMN.Value Then
            RadioOpt = "BOTH"
        End If
        
        Call Get_LiveRate(vcDTP1.Value, RadioOpt, 0)
    
    End If
    
    Frame5.Visible = False
    Call Get_ExIDs
    DoEvents
    If GETMAIN.ProgressBar1.Value + 1 < GETMAIN.ProgressBar1.Max Then
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
    End If
    DoEvents
        
    Call Fill_BuySell
    DoEvents
    If GETMAIN.ProgressBar1.Value < GETMAIN.ProgressBar1.Max Then
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
    End If
    DoEvents
        
    Call Update_ClRate
    DoEvents
    If GETMAIN.ProgressBar1.Value < GETMAIN.ProgressBar1.Max Then
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
    End If
    DoEvents
                    
    If DataCombo1.BoundText = "" And DataCombo2.BoundText = "" Then
        Call Shree_Posting(DateValue(vcDTP1.Value))
        DoEvents
        If GETMAIN.ProgressBar1.Value < GETMAIN.ProgressBar1.Max Then
            GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        End If
        DoEvents
        
        Call Update_Charges("", "", "", "", vcDTP1.Value, vcDTP2.Value, True)
        DoEvents
        If GETMAIN.ProgressBar1.Value < GETMAIN.ProgressBar1.Max Then
            GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        End If
        DoEvents
    
        Cnn.BeginTrans
    
        If BILL_GENERATION(vcDTP1.Value, GFinEnd, "", "", "") Then
            Cnn.CommitTrans: CNNERR = False
        Else
           Cnn.RollbackTrans: CNNERR = False
        End If
        DoEvents
        GETMAIN.Label1.Caption = ""
        If GETMAIN.ProgressBar1.Value < GETMAIN.ProgressBar1.Max Then
            GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        End If
        DoEvents
            
    End If
    
    Set LAccRec = New ADODB.Recordset
    mysql = " SELECT DISTINCT A.ACCID, A.NAME FROM ACCOUNTD AS A, BILLSUMREC AS B WHERE convert(varchar,A.ACCID )=convert(varchar,B.ACCID ) ORDER BY NAME "
    LAccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LAccRec.EOF Then
        Set DataCombo1.RowSource = LAccRec
        DataCombo1.BoundColumn = "ACCID"
        DataCombo1.ListField = "NAME"
    End If
    
    Set LSaudaRec = Nothing
    Set LSaudaRec = New ADODB.Recordset
    mysql = " SELECT DISTINCT A.SAUDACODE ,A.SAUDAID FROM SAUDAMAST  AS A, BILLSUMREC AS B WHERE A.SAUDAID =B.SAUDAID ORDER BY SAUDACODE  "
    LSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LSaudaRec.EOF Then
        Set DataCombo2.RowSource = LSaudaRec
        DataCombo2.BoundColumn = "SAUDAID"
        DataCombo2.ListField = "SAUDACODE"
    End If
    
    Call FLEX_GRID_REFRESH
    If GETMAIN.ProgressBar1.Value < GETMAIN.ProgressBar1.Max Then
        DoEvents
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        DoEvents
    End If
    Call FLEX_GRID_REFRESH_REPORTFORMAT
    If GETMAIN.ProgressBar1.Value < GETMAIN.ProgressBar1.Max Then
        DoEvents
        GETMAIN.ProgressBar1.Value = GETMAIN.ProgressBar1.Value + 1
        DoEvents
    End If
    
    
    SSTab1.Tab = 1
    Me.MousePointer = 0
    
End Sub
Private Sub DataCombo1_Validate(Cancel As Boolean)
    LFAccID = 0
    If LenB(DataCombo1.BoundText) > 0 Then
        LFAccID = Val(DataCombo1.BoundText)
        If LFAccID = 0 Then
            MsgBox "Imvalid Party  Please select Again"
            Cancel = True
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
        If ExRec.RecordCount > 1 Then
            Me.MousePointer = 11:
            ExListView.Visible = False
            Do While Not ExRec.EOF
                ExListView.ListItems.Add , , ExRec!excode
                ExListView.ListItems(ExListView.ListItems.Count).ListSubItems.Add , , ExRec!EXID
                ExRec.MoveNext
            Loop
        Else
            ExListView.Enabled = False:
        End If
        Me.MousePointer = 0
        ExListView.Visible = True
    End If
    
     
End Sub
Private Sub CmdEx_Click()
    If Frame5.Visible = True Then
        Frame5.Visible = False
    Else
        Frame5.Visible = True
    End If
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
End Sub
Private Sub CmdLoad_Click()
mysql = "TRUNCATE TABLE BILLSUMREC"
Cnn.Execute mysql
Call Get_ExIDs
'Call Fill_Opening

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
Private Sub FLEX_GRID_REFRESH_REPORTFORMAT()
  
    Dim LSParties As String
    Dim LSSaudas As String
    Dim LSExCodes As String
    Dim LSInst As String
    Dim Check4 As Integer
    Dim LSFmlyIDs As String
    Dim AllParties As Boolean
    Dim AllFmly As Boolean
    Dim AllSaudas As Boolean
    Dim AllExcodes As Boolean
    Dim AllInst As Boolean
    Gridflag = True
    LSParties = ""
    LSSaudas = ""
    LSExCodes = ""
    LSInst = "FUT"
    Check4 = 0
    LSFmlyIDs = ""
    AllParties = True
    AllFmly = True
    AllSaudas = True
    AllExcodes = True
    AllInst = True
    
    AllParties = True
    If DataCombo1.BoundText <> "" Then
        Set LREC = New ADODB.Recordset
        mysql = " SELECT A.AC_CODE FROM ACCOUNTD AS A WHERE A.ACCID='" & LFAccID & "'"
        LREC.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not LREC.EOF Then
            LSParties = LREC!AC_CODE
        End If
        AllParties = False
    End If
    AllSaudas = True
    If DataCombo2.BoundText <> "" Then
        LSSaudas = LFSaudaID
        AllSaudas = False
    End If
    
                
    Call Bill_summary_share(LSParties, LSSaudas, LSExCodes, LSInst, vcDTP1.Value, vcDTP2.Value, Check4, LSFmlyIDs, AllParties, AllFmly, AllSaudas, AllExcodes, AllInst, 0, "", "1", 0)
    
    If Not GlobalRecRpt.EOF Then
        
        Dim Vgridrow As Integer
        Dim Vgridcol As Integer
        
        Dim LDIFFER As Double:        Dim LBrokAmt As Double:        Dim LTranAmt As Double:        Dim LStdAmt As Double:        Dim LBIllAmt As Double
        Dim RDIFFER As Double:        Dim RBrokAmt As Double:        Dim RTranAmt As Double:        Dim RStdAmt As Double:        Dim RBIllAmt As Double
        LDIFFER = 0:        LBrokAmt = 0:        LTranAmt = 0:        LStdAmt = 0:        LBIllAmt = 0
        RDIFFER = 0:        RBrokAmt = 0:        RTranAmt = 0:        RStdAmt = 0:        RBIllAmt = 0
        
        MSFRPT.Visible = True
        'DataGrid1.Visible = False
        '>>> Set grid columns
        DoEvents
        MSFRPT.Row = 0
        MSFRPT.Col = 0: MSFRPT.ColWidth(0) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 0) = "Party": MSFRPT.CellAlignment = 1
        MSFRPT.Col = 1: MSFRPT.ColWidth(1) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 1) = "Gross MTM"
        MSFRPT.Col = 2: MSFRPT.ColWidth(2) = TextWidth("XXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 2) = "Brok."
        MSFRPT.Col = 3: MSFRPT.ColWidth(3) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 3) = "MTM Share"
        MSFRPT.Col = 4: MSFRPT.ColWidth(4) = TextWidth("XXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 4) = "Brok Share"
        MSFRPT.Col = 5: MSFRPT.ColWidth(5) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 5) = "Credit Amount"
        MSFRPT.Col = 6: MSFRPT.ColWidth(6) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 6) = "Party"
        MSFRPT.Col = 7: MSFRPT.ColWidth(7) = TextWidth("XXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 7) = "Gross MTM"
        MSFRPT.Col = 8: MSFRPT.ColWidth(8) = TextWidth("XXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 8) = "Brok."
        MSFRPT.Col = 9: MSFRPT.ColWidth(9) = TextWidth("XXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 9) = "MTM Share"
        MSFRPT.Col = 10: MSFRPT.ColWidth(10) = TextWidth("XXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 10) = "Brok Share."
        MSFRPT.Col = 11: MSFRPT.ColWidth(11) = TextWidth("XXXXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 11) = "Debit Amount"

        MSFRPT.Rows = 1
        Vgridrow = 0
        DoEvents
        While Not GlobalRecRpt.EOF
            DoEvents
            Vgridrow = Vgridrow + 1
            MSFRPT.Rows = MSFRPT.Rows + 1
            
            MSFRPT.Row = Vgridrow
            MSFRPT.Col = 0
            MSFRPT.CellAlignment = 1
            
            MSFRPT.Col = 6
            MSFRPT.CellAlignment = 1
            
            MSFRPT.CellAlignment = 1
            If (InStr(1, GlobalRecRpt!LName, "ZZZZ") > 0) Then
                MSFRPT.TextMatrix(Vgridrow, 0) = ""
                MSFRPT.TextMatrix(Vgridrow, 1) = ""
                MSFRPT.TextMatrix(Vgridrow, 2) = ""
                MSFRPT.TextMatrix(Vgridrow, 3) = ""
                MSFRPT.TextMatrix(Vgridrow, 4) = ""
                MSFRPT.TextMatrix(Vgridrow, 5) = ""
            Else
                MSFRPT.TextMatrix(Vgridrow, 0) = GlobalRecRpt!LName
                MSFRPT.TextMatrix(Vgridrow, 1) = Format(GlobalRecRpt!LDIFFER, "0.00")
                If GlobalRecRpt!LDIFFER > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 1
                    MSFRPT.CellBackColor = vbBlue
                ElseIf GlobalRecRpt!LDIFFER < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 1
                    MSFRPT.CellBackColor = vbRed
                End If
                
                MSFRPT.TextMatrix(Vgridrow, 2) = Format(GlobalRecRpt!LBrokAmt, "0.00")
                If GlobalRecRpt!LBrokAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 2
                    MSFRPT.CellBackColor = vbBlue
                ElseIf GlobalRecRpt!LBrokAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 2
                    MSFRPT.CellBackColor = vbRed
                End If
                
                MSFRPT.TextMatrix(Vgridrow, 3) = Format(GlobalRecRpt!LTranAmt, "0.00")
                If GlobalRecRpt!LTranAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 3
                    MSFRPT.CellBackColor = vbBlue
                ElseIf GlobalRecRpt!LTranAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 3
                    MSFRPT.CellBackColor = vbRed
                End If
                
                MSFRPT.TextMatrix(Vgridrow, 4) = Format(GlobalRecRpt!LStdAmt, "0.00")
                If GlobalRecRpt!LStdAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 4
                    MSFRPT.CellBackColor = vbBlue
                ElseIf GlobalRecRpt!LStdAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 4
                    MSFRPT.CellBackColor = vbRed
                End If
                
                MSFRPT.TextMatrix(Vgridrow, 5) = Format(GlobalRecRpt!LBIllAmt, "0.00")
                If GlobalRecRpt!LBIllAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 5
                    MSFRPT.CellBackColor = vbBlue
                ElseIf GlobalRecRpt!LBIllAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 5
                    MSFRPT.CellBackColor = vbRed
                End If
                
                LDIFFER = LDIFFER + GlobalRecRpt!LDIFFER
                LBrokAmt = LBrokAmt + GlobalRecRpt!LBrokAmt
                LTranAmt = LTranAmt + GlobalRecRpt!LTranAmt
                LStdAmt = LStdAmt + GlobalRecRpt!LStdAmt
                LBIllAmt = LBIllAmt + GlobalRecRpt!LBIllAmt
            End If
            
            If (InStr(1, GlobalRecRpt!RNAME, "ZZZZ") > 0) Then
                MSFRPT.TextMatrix(Vgridrow, 6) = ""
                MSFRPT.TextMatrix(Vgridrow, 7) = ""
                MSFRPT.TextMatrix(Vgridrow, 8) = ""
                MSFRPT.TextMatrix(Vgridrow, 9) = ""
                MSFRPT.TextMatrix(Vgridrow, 10) = ""
                MSFRPT.TextMatrix(Vgridrow, 11) = ""
            Else
                MSFRPT.TextMatrix(Vgridrow, 6) = GlobalRecRpt!RNAME
                MSFRPT.TextMatrix(Vgridrow, 7) = Format(GlobalRecRpt!RDIFFER, "0.00")
                If GlobalRecRpt!RDIFFER < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 7
                    MSFRPT.CellBackColor = vbRed
                ElseIf GlobalRecRpt!RDIFFER > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 7
                    MSFRPT.CellBackColor = vbBlue
                End If
                MSFRPT.TextMatrix(Vgridrow, 8) = Format(GlobalRecRpt!RBrokAmt, "0.00")
                If GlobalRecRpt!RBrokAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 8
                    MSFRPT.CellBackColor = vbRed
                ElseIf GlobalRecRpt!RBrokAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 8
                    MSFRPT.CellBackColor = vbBlue
                End If
                MSFRPT.TextMatrix(Vgridrow, 9) = Format(GlobalRecRpt!RTranAmt, "0.00")
                If GlobalRecRpt!RTranAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 9
                    MSFRPT.CellBackColor = vbRed
                ElseIf GlobalRecRpt!RTranAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 9
                    MSFRPT.CellBackColor = vbBlue
                End If
                MSFRPT.TextMatrix(Vgridrow, 10) = Format(GlobalRecRpt!RStdAmt, "0.00")
                If GlobalRecRpt!RStdAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 10
                    MSFRPT.CellBackColor = vbRed
                ElseIf GlobalRecRpt!RStdAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 10
                    MSFRPT.CellBackColor = vbBlue
                End If
                MSFRPT.TextMatrix(Vgridrow, 11) = Format(GlobalRecRpt!RBIllAmt, "0.00")
                If GlobalRecRpt!RBIllAmt < 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 11
                    MSFRPT.CellBackColor = vbRed
                ElseIf GlobalRecRpt!RBIllAmt > 0 Then
                    MSFRPT.Row = Vgridrow
                    MSFRPT.Col = 11
                    MSFRPT.CellBackColor = vbBlue
                End If
                RDIFFER = RDIFFER + GlobalRecRpt!RDIFFER
                RBrokAmt = RBrokAmt + GlobalRecRpt!RBrokAmt
                RTranAmt = RTranAmt + GlobalRecRpt!RTranAmt
                RStdAmt = RStdAmt + GlobalRecRpt!RStdAmt
                RBIllAmt = RBIllAmt + GlobalRecRpt!RBIllAmt
            End If
        

'            TOTGrossMTM = TOTGrossMTM + Format(GlobalRecRpt!GrossMTM, "0.00")
'            TOTBROKAMT = TOTBROKAMT + Format(GlobalRecRpt!BROKAMT, "0.00")
'            TOTBILLAMT = TOTBILLAMT + Format(GlobalRecRpt!BILLAMT, "0.00")
'
'            TOTM2MSHARE = TOTM2MSHARE + Format(GlobalRecRpt!M2MSHARE, "0.00")
'            TOTBROKSHARE = TOTBROKSHARE + Format(GlobalRecRpt!BROKSHARE, "0.00")
'            TOTNETAMT = TOTNETAMT + Format(GlobalRecRpt!M2MSHARE, "0.00") + Format(GlobalRecRpt!BROKSHARE, "0.00") + Format(GlobalRecRpt!BILLAMT, "0.00")

            DoEvents
            GlobalRecRpt.MoveNext
        Wend
        DoEvents
        Vgridrow = Vgridrow + 1
        MSFRPT.Rows = MSFRPT.Rows + 1
        MSFRPT.Row = Vgridrow
                            
        MSFRPT.TextMatrix(Vgridrow, 0) = "Total": MSFRPT.Col = 0: MSFRPT.CellBackColor = vbBlue
        MSFRPT.TextMatrix(Vgridrow, 1) = Format(LDIFFER, "0.00"): MSFRPT.Col = 1:
        If LDIFFER < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        MSFRPT.TextMatrix(Vgridrow, 2) = Format(LBrokAmt, "0.00"): MSFRPT.Col = 2:
        If LBrokAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
                
        MSFRPT.TextMatrix(Vgridrow, 3) = Format(LTranAmt, "0.00"): MSFRPT.Col = 3:
        If LTranAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        MSFRPT.TextMatrix(Vgridrow, 4) = Format(LStdAmt, "0.00"): MSFRPT.Col = 4:
        If LStdAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        MSFRPT.TextMatrix(Vgridrow, 5) = Format(LBIllAmt, "0.00"): MSFRPT.Col = 5:
        If LBIllAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        MSFRPT.TextMatrix(Vgridrow, 6) = "Total": MSFRPT.Col = 6: MSFRPT.CellBackColor = vbBlue
        MSFRPT.TextMatrix(Vgridrow, 7) = Format(RDIFFER, "0.00"): MSFRPT.Col = 7:
        If RDIFFER < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        MSFRPT.TextMatrix(Vgridrow, 8) = Format(RBrokAmt, "0.00"): MSFRPT.Col = 8:
        If RBrokAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        MSFRPT.TextMatrix(Vgridrow, 9) = Format(RTranAmt, "0.00"): MSFRPT.Col = 9:
        If RTranAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        MSFRPT.TextMatrix(Vgridrow, 10) = Format(RStdAmt, "0.00"): MSFRPT.Col = 10:
        If RStdAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        MSFRPT.TextMatrix(Vgridrow, 11) = Format(RBIllAmt, "0.00"): MSFRPT.Col = 11:
        If RBIllAmt < 0 Then
            MSFRPT.CellBackColor = vbRed
        Else
            MSFRPT.CellBackColor = vbBlue
        End If
        
        
    End If
    Gridflag = False
End Sub
Private Sub FLEX_GRID_REFRESH()

'>>>select opqty, oprate, opamt,sellqty, sellrate, sellamt, buyqty,buyrate, buyamt, clqty,clrate,clamt, calval, (sellamt - opamt - buyamt + clamt) as 'GrossMTM'  from billsumrec where accid=34 and saudaid=461

    Dim BillRec  As ADODB.Recordset:      Dim BillSumm As ADODB.Recordset
    Dim LOpBal As Double:                       Dim LBal As Double
    Dim LAC_CODE As String:             Dim LSTR  As String
    Dim TOTGrossMTM As Double
    Dim TOTBROKAMT As Double
    Dim TOTBILLAMT As Double
    Dim TOTM2MSHARE As Double
    Dim TOTBROKSHARE As Double
    Dim TOTNETAMT As Double
        
    TOTGrossMTM = 0
    TOTBROKAMT = 0
    TOTBILLAMT = 0
    TOTM2MSHARE = 0
    TOTBROKSHARE = 0
    TOTNETAMT = 0
     
    DoEvents
    Label4.Caption = "Getting Final Data "
    Label10.Caption = LFAccID
    Label11.Caption = LFSaudaID
    
    DoEvents
    LSTR = "."
    Call SetRec
    mysql = "EXEC INSERT_BILLSUM2"
    Cnn.Execute mysql
    Set BillRec = Nothing
    Set BillRec = Nothing
    Set BillRec = New ADODB.Recordset
    mysql = "SELECT B.ACCID ,B.NAME,B.CONTRACT,"
    mysql = mysql & " B.OpenQty,B.OpenRate,B.OpenAmt,B.BuyQty,B.BuyRate,B.SellQty,B.SellRate,B.NetQty,"
    mysql = mysql & " B.NetRate,B.CloseQty,B.LTP,B.GrossMTM,B.BrokAmt,B.GrossMTM+B.BROKAMT AS BillAmt, "
    mysql = mysql & " isnull((SELECT  sum(Z.Samount) FROM INV_D1 Z WHERE Z.stdate >= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND Z.stdate <= '" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' AND Z.ACCID=B.ACCID AND Z.SAUDAID=B.SAUDAID) ,0) as 'M2MSHARE', "
    mysql = mysql & " isnull ((SELECT  sum(Z.amount) FROM INV_D1 Z WHERE Z.stdate >= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND Z.stdate <= '" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' AND Z.ACCID=B.ACCID AND Z.SAUDAID=B.SAUDAID),0)  as 'BROKSHARE', "
    mysql = mysql & " 0 AS NETAMT "
    mysql = mysql & " FROM BILLSUM B"
             
    If LFAccID <> 0 Then
        mysql = mysql & "   Where  ACCID  =" & LFAccID & ""
    End If
    
    If LFSaudaID <> 0 Then
        If InStr(1, mysql, "Where") > 0 Then
            mysql = mysql & "  AND SaudaID ='" & LFSaudaID & "'"
        Else
            mysql = mysql & "  where  SaudaID ='" & LFSaudaID & "'"
        End If
    End If
    
    'Else
      '  If LenB(LFSauda) > 1 Then
       ' MYSQL = MYSQL & "  AND CONTRACT ='" & LFSauda & "'"
    'End If
    mysql = mysql & " ORDER BY NAME,CONTRACT"
    
    'MYSQL = "SELECT B.Party,A.Name,S.SAUDACODE,(B.BUYQTY + (B.SELLQTY*-1)) AS NetQty,"
    'MYSQL = MYSQL & " CASE WHEN (B.BUYQTY + (B.SELLQTY*-1))= 0  THEN 0 WHEN (B.BUYQTY + (B.SELLQTY*-1))<> 0 THEN ABS(B.SELLAMT-B.BUYAMT)/(ABS((B.BUYQTY + (B.SELLQTY*-1))*CALVAL)) END  AS NetRate "
    'MYSQL = MYSQL & " ,B.CLRATE AS LTP,(B.BUYAMT*-1)+B.SELLAMT+(B.CLAMT*-1) as GrossMTM ,B.BROKAMT,"
    'MYSQL = MYSQL & " CASE WHEN MARGIN <1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))*B.CLRATE*B.CALVAL*B.MARGIN WHEN MARGIN >=1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))* MARGIN END  AS MarginAmt  "
    'MYSQL = MYSQL & " FROM ACCOUNTD AS A, BILLSUMREC AS B , SAUDAMAST AS S "
    'MYSQL = MYSQL & " WHERE A.AC_CODE=B.PARTY and S.SAUDAID =B.SAUDAID "
    
    'MYSQL = MYSQL & " ORDER BY A.Name,S.SAUDACODE "
    'BillRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        
    'Set BillRec = New ADODB.Recordset
    'MYSQL = "SELECT B.Party,A.Name,S.SAUDACODE,(B.BUYQTY + (B.SELLQTY*-1)) AS NetQty,"
    'MYSQL = MYSQL & " CASE WHEN (B.BUYQTY + (B.SELLQTY*-1))= 0  THEN 0 WHEN (B.BUYQTY + (B.SELLQTY*-1))<> 0 THEN ABS(B.SELLAMT-B.BUYAMT)/(ABS((B.BUYQTY + (B.SELLQTY*-1))*CALVAL)) END  AS NetRate "
    'MYSQL = MYSQL & " ,B.CLRATE AS LTP,(B.BUYAMT*-1)+B.SELLAMT+(B.CLAMT*-1) as GrossMTM ,B.BROKAMT,"
    'MYSQL = MYSQL & " CASE WHEN MARGIN <1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))*B.CLRATE*B.CALVAL*B.MARGIN WHEN MARGIN >=1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))* MARGIN END  AS MarginAmt  "
    'MYSQL = MYSQL & " FROM ACCOUNTD AS A, BILLSUMREC AS B , SAUDAMAST AS S "
    'MYSQL = MYSQL & " WHERE A.AC_CODE=B.PARTY and S.SAUDAID =B.SAUDAID "
    'If Len(DataCombo1.BoundText) > 1 Then MYSQL = MYSQL & " AND  A.AC_CODE='" & DataCombo1.BoundText & "'"
    'MYSQL = MYSQL & " ORDER BY A.Name,S.SAUDACODE "
    
    BillRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
           
    
    If Not BillRec.EOF Then
        Dim Vgridrow As Integer
        Dim Vgridcol As Integer
        
        MSFdetail.Visible = True
        DataGrid1.Visible = False
        '>>> Set grid columns
        DoEvents
        MSFdetail.Row = 0
        MSFdetail.Col = 0: MSFdetail.ColWidth(0) = TextWidth("XXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 0) = "AccID"
        MSFdetail.Col = 1: MSFdetail.ColWidth(1) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 1) = "Name"
        MSFdetail.Col = 2: MSFdetail.ColWidth(2) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 2) = "Contract"
        MSFdetail.Col = 3: MSFdetail.ColWidth(3) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 3) = "Op.Qty."
        MSFdetail.Col = 4: MSFdetail.ColWidth(4) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 4) = "Op.Rate"
        'MSFdetail.Col = 5: MSFdetail.ColWidth(5) = TextWidth("XXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 5) = "Op.Amt."
        MSFdetail.Col = 5: MSFdetail.ColWidth(5) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 5) = "BuyQty."
        MSFdetail.Col = 6: MSFdetail.ColWidth(6) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 6) = "BuyRate"
        MSFdetail.Col = 7: MSFdetail.ColWidth(7) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 7) = "SellQty."
        MSFdetail.Col = 8: MSFdetail.ColWidth(8) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 8) = "SellRate"
        MSFdetail.Col = 9: MSFdetail.ColWidth(9) = TextWidth("XXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 9) = "Tot.(Buy-Sell)"
        'MSFdetail.Col = 11: MSFdetail.ColWidth(11) = TextWidth("XXXXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 11) = "NetRate"
        MSFdetail.Col = 10: MSFdetail.ColWidth(10) = TextWidth("XXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 10) = "CloseQty."
        MSFdetail.Col = 11: MSFdetail.ColWidth(11) = TextWidth("XXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 11) = "LTP"
        MSFdetail.Col = 12: MSFdetail.ColWidth(12) = TextWidth("XXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 12) = "GrossMTM"
        MSFdetail.Col = 13: MSFdetail.ColWidth(13) = TextWidth("XXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 13) = "BrokAmt."
        MSFdetail.Col = 14: MSFdetail.ColWidth(14) = TextWidth("XXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 14) = "BillAmt."
        
        MSFdetail.Col = 15: MSFdetail.ColWidth(15) = TextWidth("XXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 15) = "M2M Share"
        MSFdetail.Col = 16: MSFdetail.ColWidth(16) = TextWidth("XXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 16) = "Brok. Share"
        MSFdetail.Col = 17: MSFdetail.ColWidth(17) = TextWidth("XXXXXXXXXXXXXXX"): MSFdetail.CellFontBold = True: MSFdetail.TextMatrix(0, 17) = "NetlAmt."

        MSFdetail.Rows = 1
        Vgridrow = 0
        DoEvents
        While Not BillRec.EOF
            DoEvents
            Vgridrow = Vgridrow + 1
            MSFdetail.Rows = MSFdetail.Rows + 1
            MSFdetail.TextMatrix(Vgridrow, 0) = BillRec!ACCID
            MSFdetail.TextMatrix(Vgridrow, 1) = BillRec!NAME
            MSFdetail.TextMatrix(Vgridrow, 2) = BillRec!CONTRACT
            MSFdetail.TextMatrix(Vgridrow, 3) = BillRec!OpenQty
            MSFdetail.TextMatrix(Vgridrow, 4) = BillRec!OpenRate
            'MSFdetail.TextMatrix(Vgridrow, 5) = Format(BillRec!OpenAmt, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 5) = BillRec!BUYQTY
            MSFdetail.TextMatrix(Vgridrow, 6) = Format(BillRec!BuyRATE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 7) = BillRec!SELLQTY
            MSFdetail.TextMatrix(Vgridrow, 8) = Format(BillRec!SellRATE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 9) = BillRec!NETQTY
            'MSFdetail.TextMatrix(Vgridrow, 11) = Format(BillRec!NETRATE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 10) = BillRec!CLOSEQTY
            MSFdetail.TextMatrix(Vgridrow, 11) = Format(BillRec!LTP, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 12) = Format(BillRec!GrossMTM, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 13) = Format(BillRec!BROKAMT, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 14) = Format(BillRec!Billamt, "0.00")
            
            MSFdetail.TextMatrix(Vgridrow, 15) = Format(BillRec!M2MSHARE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 16) = Format(BillRec!BROKSHARE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 17) = Format(BillRec!NETAMT, "0.00")
            
            TOTGrossMTM = TOTGrossMTM + Format(BillRec!GrossMTM, "0.00")
            TOTBROKAMT = TOTBROKAMT + Format(BillRec!BROKAMT, "0.00")
            TOTBILLAMT = TOTBILLAMT + Format(BillRec!Billamt, "0.00")
            
            TOTM2MSHARE = TOTM2MSHARE + Format(BillRec!M2MSHARE, "0.00")
            TOTBROKSHARE = TOTBROKSHARE + Format(BillRec!BROKSHARE, "0.00")
            TOTNETAMT = TOTNETAMT + Format(BillRec!M2MSHARE, "0.00") + Format(BillRec!BROKSHARE, "0.00") + Format(BillRec!Billamt, "0.00")
            
            MSFdetail.Row = Vgridrow
            MSFdetail.Col = 12
            If BillRec!GrossMTM < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf BillRec!GrossMTM > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            
            MSFdetail.Row = Vgridrow
            MSFdetail.Col = 13
            If BillRec!BROKAMT < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf BillRec!BROKAMT > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
                        
            MSFdetail.Row = Vgridrow
            MSFdetail.Col = 14
            If BillRec!Billamt < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf BillRec!Billamt > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            
            MSFdetail.Row = Vgridrow
            MSFdetail.Col = 15
            If BillRec!M2MSHARE < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf BillRec!M2MSHARE > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
                        
            MSFdetail.Row = Vgridrow
            MSFdetail.Col = 16
            If BillRec!BROKSHARE < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf BillRec!BROKSHARE > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            
            DoEvents
            BillRec.MoveNext
        Wend
         
        If LFAccID > 0 Then
            Vgridrow = Vgridrow + 1
            MSFdetail.Rows = MSFdetail.Rows + 1
            MSFdetail.TextMatrix(Vgridrow, 0) = ""
            MSFdetail.TextMatrix(Vgridrow, 1) = ""
            MSFdetail.TextMatrix(Vgridrow, 2) = ""
            MSFdetail.TextMatrix(Vgridrow, 3) = ""
            MSFdetail.TextMatrix(Vgridrow, 4) = ""
            'MSFdetail.TextMatrix(Vgridrow, 5) = ""
            MSFdetail.TextMatrix(Vgridrow, 5) = ""
            MSFdetail.TextMatrix(Vgridrow, 6) = ""
            MSFdetail.TextMatrix(Vgridrow, 7) = ""
            MSFdetail.TextMatrix(Vgridrow, 8) = ""
            MSFdetail.TextMatrix(Vgridrow, 9) = ""
            'MSFdetail.TextMatrix(Vgridrow, 11) = ""
            MSFdetail.TextMatrix(Vgridrow, 10) = ""
            MSFdetail.TextMatrix(Vgridrow, 11) = ""
            MSFdetail.TextMatrix(Vgridrow, 12) = Format(TOTGrossMTM, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 13) = Format(TOTBROKAMT, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 14) = Format(TOTBILLAMT, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 15) = Format(TOTM2MSHARE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 16) = Format(TOTBROKSHARE, "0.00")
            MSFdetail.TextMatrix(Vgridrow, 17) = Format(TOTNETAMT, "0.00")
            
            MSFdetail.Row = Vgridrow
            MSFdetail.Col = 12
            If TOTGrossMTM < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf TOTGrossMTM > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            MSFdetail.Col = 13
            If TOTBROKAMT < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf TOTBROKAMT > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            MSFdetail.Col = 14
            If TOTBILLAMT < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf TOTBILLAMT > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            MSFdetail.Col = 15
            If TOTM2MSHARE < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf TOTM2MSHARE > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            MSFdetail.Col = 16
            If TOTBROKSHARE < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf TOTBROKSHARE > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
            MSFdetail.Col = 17
            If TOTNETAMT < 0 Then
                MSFdetail.CellBackColor = vbRed
            ElseIf TOTNETAMT > 0 Then
                MSFdetail.CellBackColor = vbBlue
            End If
        End If
    
    End If '>>>     If Not BillRec.EOF Then
    
    
    'MSFsummary.Visible = False
'    Set BillSumm = Nothing
'    Set BillSumm = New ADODB.Recordset
'
'    MYSQL = "SELECT ACCID ,NAME,EXCODE,SUM(GROSSMTM) AS GROSSMTM, SUM(BROKAMT) AS BROKAMT, SUM(GROSSMTM+BROKAMT) AS NETAMT FROM BILLSUM"
'    MYSQL = MYSQL & " GROUP BY ACCID ,NAME,EXCODE"
'    MYSQL = MYSQL & " ORDER BY ACCID ,NAME,EXCODE"
'    BillSumm.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'
'    If Not BillSumm.EOF Then
'        DataGrid2.Visible = False
'        MSFsummary.Visible = True
'        '>>> Set grid columns
'        DoEvents
'        MSFsummary.Row = 0
'        MSFsummary.Col = 0: MSFsummary.ColWidth(0) = TextWidth("XXXXXXX"): MSFsummary.CellFontBold = True: MSFsummary.TextMatrix(0, 0) = "AccID"
'        MSFsummary.Col = 1: MSFsummary.ColWidth(1) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFsummary.CellFontBold = True: MSFsummary.TextMatrix(0, 1) = "Name"
'        MSFsummary.Col = 2: MSFsummary.ColWidth(2) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFsummary.CellFontBold = True: MSFsummary.TextMatrix(0, 2) = "ExCode"
'        MSFsummary.Col = 3: MSFsummary.ColWidth(3) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFsummary.CellFontBold = True: MSFsummary.TextMatrix(0, 3) = "GrossMTM"
'        MSFsummary.Col = 4: MSFsummary.ColWidth(4) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFsummary.CellFontBold = True: MSFsummary.TextMatrix(0, 4) = "BrokAmt."
'        MSFsummary.Col = 5: MSFsummary.ColWidth(5) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MSFsummary.CellFontBold = True: MSFsummary.TextMatrix(0, 5) = "NetAmt."
'        MSFsummary.Rows = 1
'        Vgridrow = 0
'        DoEvents
'        While Not BillSumm.EOF
'            DoEvents
'            Vgridrow = Vgridrow + 1
'            MSFsummary.Rows = MSFsummary.Rows + 1
'            MSFsummary.TextMatrix(Vgridrow, 0) = BillSumm!ACCID
'            MSFsummary.TextMatrix(Vgridrow, 1) = BillSumm!NAME
'            MSFsummary.TextMatrix(Vgridrow, 2) = BillSumm!EXCODE
'            MSFsummary.TextMatrix(Vgridrow, 3) = Format(BillSumm!GrossMTM, "0.00")
'            MSFsummary.TextMatrix(Vgridrow, 4) = Format(BillSumm!BROKAMT, "0.00")
'            MSFsummary.TextMatrix(Vgridrow, 5) = Format(BillSumm!NETAMT, "0.00")
'
''            MSFsummary.Row = Vgridrow
''            MSFsummary.Col = 16
''            If BillRec!BillAmt < 0 Then
''                MSFsummary.CellBackColor = vbRed
''            ElseIf BillRec!BillAmt > 0 Then
''                MSFsummary.CellBackColor = vbBlue
''            End If
'            DoEvents
'            BillSumm.MoveNext
'        Wend
'    End If
    
End Sub

'Private Sub DATA_GRID_REFRESH()
'
'    Dim BillRec  As ADODB.Recordset:    Dim BillSumm As ADODB.Recordset
'    Dim LOpBal As Double:               Dim LBal As Double
'    Dim LAC_CODE As String:             Dim LSTR  As String
'    DoEvents
'    Label4.Caption = "Getting Final Data "
'    Label10.Caption = LFAccID
'    Label11.Caption = LFSaudaID
'
'    DoEvents
'    LSTR = "."
'    Call SetRec
'    MYSQL = "EXEC INSERT_BILLSUM2"
'    Cnn.Execute MYSQL
'    Set BillRec = Nothing
'    Set BillRec = Nothing
'    Set BillRec = New ADODB.Recordset
'    MYSQL = "SELECT ACCID ,NAME,CONTRACT,"
'    MYSQL = MYSQL & " OpenQty,OpenRate,OpenAmt,BuyQty,BuyRate,SellQty,SellRate,NetQty,"
'    MYSQL = MYSQL & " NetRate,CloseQty,LTP,GrossMTM,BrokAmt,GrossMTM+BROKAMT AS BillAmt"
'    MYSQL = MYSQL & " FROM BILLSUM "
'
'    If LFAccID <> 0 Then
'        MYSQL = MYSQL & "   Where  ACCID  =" & LFAccID & ""
'    End If
'
'    If LFSaudaID <> 0 Then
'        If InStr(1, MYSQL, "Where") > 0 Then
'            MYSQL = MYSQL & "  AND SaudaID ='" & LFSaudaID & "'"
'        Else
'            MYSQL = MYSQL & "  where  SaudaID ='" & LFSaudaID & "'"
'        End If
'    End If
'
'    'Else
'      '  If LenB(LFSauda) > 1 Then
'       ' MYSQL = MYSQL & "  AND CONTRACT ='" & LFSauda & "'"
'    'End If
'    MYSQL = MYSQL & " ORDER BY NAME,CONTRACT"
'
'    'MYSQL = "SELECT B.Party,A.Name,S.SAUDACODE,(B.BUYQTY + (B.SELLQTY*-1)) AS NetQty,"
'    'MYSQL = MYSQL & " CASE WHEN (B.BUYQTY + (B.SELLQTY*-1))= 0  THEN 0 WHEN (B.BUYQTY + (B.SELLQTY*-1))<> 0 THEN ABS(B.SELLAMT-B.BUYAMT)/(ABS((B.BUYQTY + (B.SELLQTY*-1))*CALVAL)) END  AS NetRate "
'    'MYSQL = MYSQL & " ,B.CLRATE AS LTP,(B.BUYAMT*-1)+B.SELLAMT+(B.CLAMT*-1) as GrossMTM ,B.BROKAMT,"
'    'MYSQL = MYSQL & " CASE WHEN MARGIN <1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))*B.CLRATE*B.CALVAL*B.MARGIN WHEN MARGIN >=1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))* MARGIN END  AS MarginAmt  "
'    'MYSQL = MYSQL & " FROM ACCOUNTD AS A, BILLSUMREC AS B , SAUDAMAST AS S "
'    'MYSQL = MYSQL & " WHERE A.AC_CODE=B.PARTY and S.SAUDAID =B.SAUDAID "
'
'    'MYSQL = MYSQL & " ORDER BY A.Name,S.SAUDACODE "
'    'BillRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'
'
'    'Set BillRec = New ADODB.Recordset
'    'MYSQL = "SELECT B.Party,A.Name,S.SAUDACODE,(B.BUYQTY + (B.SELLQTY*-1)) AS NetQty,"
'    'MYSQL = MYSQL & " CASE WHEN (B.BUYQTY + (B.SELLQTY*-1))= 0  THEN 0 WHEN (B.BUYQTY + (B.SELLQTY*-1))<> 0 THEN ABS(B.SELLAMT-B.BUYAMT)/(ABS((B.BUYQTY + (B.SELLQTY*-1))*CALVAL)) END  AS NetRate "
'    'MYSQL = MYSQL & " ,B.CLRATE AS LTP,(B.BUYAMT*-1)+B.SELLAMT+(B.CLAMT*-1) as GrossMTM ,B.BROKAMT,"
'    'MYSQL = MYSQL & " CASE WHEN MARGIN <1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))*B.CLRATE*B.CALVAL*B.MARGIN WHEN MARGIN >=1 THEN abs(B.BUYQTY + (B.SELLQTY*-1))* MARGIN END  AS MarginAmt  "
'    'MYSQL = MYSQL & " FROM ACCOUNTD AS A, BILLSUMREC AS B , SAUDAMAST AS S "
'    'MYSQL = MYSQL & " WHERE A.AC_CODE=B.PARTY and S.SAUDAID =B.SAUDAID "
'    'If Len(DataCombo1.BoundText) > 1 Then MYSQL = MYSQL & " AND  A.AC_CODE='" & DataCombo1.BoundText & "'"
'    'MYSQL = MYSQL & " ORDER BY A.Name,S.SAUDACODE "
'
'    BillRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'    Set DataGrid1.DataSource = BillRec
'    DataGrid1.Visible = True
'    MSFdetail.Visible = False
'
'    DataGrid1.ReBind
'    DataGrid1.Refresh
'    DataGrid1.Columns(0).Width = 1100 ' party
'    DataGrid1.Columns(1).Width = 1600 ' name
'    DataGrid1.Columns(2).Width = 3600 ' contract
'
'        DataGrid1.Columns(3).Alignment = dbgRight ' openqty
'        'DataGrid1.Columns(3).NumberFormat = "0.00"
'        DataGrid1.Columns(3).Width = 1200
'        DataGrid1.Columns(4).Width = 1400 ' openrate
'        DataGrid1.Columns(5).Width = 1700 ' open amt
'        DataGrid1.Columns(6).Width = 1200 ' buyqty
'        DataGrid1.Columns(7).Width = 1400
'        DataGrid1.Columns(8).Width = 1200 ' sellqty
'        DataGrid1.Columns(9).Width = 1400
'        DataGrid1.Columns(10).Width = 1200 ' NetQty
'        DataGrid1.Columns(11).Width = 1400
'        DataGrid1.Columns(12).Width = 1200 ' CloseQty
'        DataGrid1.Columns(13).Width = 1400
'
'        DataGrid1.Columns(14).Width = 1700
'        DataGrid1.Columns(15).Width = 1700
'        DataGrid1.Columns(16).Width = 1700
'        'DataGrid1.Columns(17).Width = 1800
'
'        DataGrid1.Columns(3).NumberFormat = "0.00"
'        DataGrid1.Columns(6).NumberFormat = "0.00"
'        DataGrid1.Columns(8).NumberFormat = "0.00"
'        DataGrid1.Columns(8).NumberFormat = "0.00"
'        DataGrid1.Columns(12).NumberFormat = "0.00"
'        DataGrid1.Columns(10).NumberFormat = "0.00"
'        DataGrid1.Columns(4).NumberFormat = "0.0000"
'        DataGrid1.Columns(4).Alignment = dbgRight
'        DataGrid1.Columns(5).Alignment = dbgRight
'        DataGrid1.Columns(6).Alignment = dbgRight
'        DataGrid1.Columns(7).Alignment = dbgRight
'        DataGrid1.Columns(8).Alignment = dbgRight
'        DataGrid1.Columns(9).Alignment = dbgRight
'        DataGrid1.Columns(10).Alignment = dbgRight
'        DataGrid1.Columns(11).Alignment = dbgRight
'        DataGrid1.Columns(12).Alignment = dbgRight
'        DataGrid1.Columns(13).Alignment = dbgRight
'        DataGrid1.Columns(14).Alignment = dbgRight
'        DataGrid1.Columns(15).Alignment = dbgRight
'        DataGrid1.Columns(16).Alignment = dbgRight
'        'DataGrid1.Columns(17).Alignment = dbgRight
'
'        DataGrid1.Columns(4).NumberFormat = "0.0000"
'        DataGrid1.Columns(5).NumberFormat = "0.0000"
'        'DataGrid1.Columns(6).NumberFormat = "0.0000"
'        DataGrid1.Columns(7).NumberFormat = "0.0000"
'        'DataGrid1.Columns(8).NumberFormat = "0.0000"
'        DataGrid1.Columns(9).NumberFormat = "0.0000"
'        'DataGrid1.Columns(10).NumberFormat = "0.0000"
'        DataGrid1.Columns(11).NumberFormat = "0.0000"
'        'DataGrid1.Columns(12).NumberFormat = "0.0000"
'        DataGrid1.Columns(13).NumberFormat = "0.0000"
'
'        DataGrid1.Columns(14).NumberFormat = "##,###0.00"
'        DataGrid1.Columns(15).NumberFormat = "##,###0.00"
'        DataGrid1.Columns(16).NumberFormat = "##,###0.00"
'        'DataGrid1.Columns(17).NumberFormat = "##,###0.00"
'    Set BillSumm = Nothing
'    Set BillSumm = New ADODB.Recordset
'
'    MYSQL = "SELECT ACCID ,NAME,EXCODE,SUM(GROSSMTM) AS GROSSMTM, SUM(BROKAMT) AS BROKAMT, SUM(GROSSMTM+BROKAMT) AS NETAMT FROM BILLSUM"
'    MYSQL = MYSQL & " GROUP BY ACCID ,NAME,EXCODE"
'    MYSQL = MYSQL & " ORDER BY ACCID ,NAME,EXCODE"
'    BillSumm.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'
'    Set DataGrid2.DataSource = BillSumm
'    DataGrid2.ReBind
'    DataGrid2.Refresh
'    DataGrid2.Columns(0).Width = 1100
'    DataGrid2.Columns(1).Width = 4600
'    DataGrid2.Columns(2).Width = 2600
'    DataGrid2.Columns(3).Width = 2600
'    DataGrid2.Columns(4).Width = 2600
'    DataGrid2.Columns(5).Width = 2600
'    'DataGrid2.Columns(6).Width = 2500
'    DataGrid2.Columns(3).Alignment = dbgRight
'    DataGrid2.Columns(4).Alignment = dbgRight
'    DataGrid2.Columns(5).Alignment = dbgRight
'    'DataGrid2.Columns(6).Alignment = dbgRight
'    DataGrid2.Columns(3).NumberFormat = "##,###0.00"
'    DataGrid2.Columns(4).NumberFormat = "##,###0.00"
'    DataGrid2.Columns(5).NumberFormat = "##,###0.00"
'    'DataGrid2.Columns(6).NumberFormat = "##,###0.00"
'
'End Sub

Private Sub Fill_BuySell()
Dim TRec As ADODB.Recordset:    Dim LBuyQty As Double:      Dim LSellQty As Double:         Dim LBuyAmt As Double
Dim LSellAmt As Double:         Dim LBrokAmt  As Double:    Dim LParty As String:            Dim LSaudaID As Long
Dim LExID As Double:            Dim LCalval As Double:      Dim lOpQty As Double:            Dim LOpAmt As Double: Dim LClQty As Double
Dim LACCID As Long
Dim LOpenRate As Double
    Timer1.Enabled = False
    mysql = "TRUNCATE TABLE BILLSUMREC"
    Cnn.Execute mysql
    DoEvents
    Label4.Caption = "Creating New Bill Summ  "
    
    '>>> Opening + Closing
    DoEvents
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = " SELECT ACCID,EXID,SAUDAID,CALVAL,CONTYPE,SUM(QTY) AS LQTY ,SUM(QTY*RATE*CALVAL) AS AMT "
    mysql = mysql & " FROM CTR_D "
    mysql = mysql & " Where  COMPCODE = " & GCompCode & " "
    mysql = mysql & " AND SAUDAID IN (SELECT SAUDAID FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
    mysql = mysql & " AND CONDATE <'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
    If LenB(LExIDS) > 1 Then mysql = mysql & " AND EXID IN (" & LExIDS & ") "
    mysql = mysql & " GROUP BY ACCID , EXID,SAUDAID,CALVAL,CONTYPE"
    mysql = mysql & " ORDER BY ACCID , EXID,SAUDAID,CONTYPE"
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Do While Not TRec.EOF
        LACCID = TRec!ACCID:          LSaudaID = TRec!SAUDAID:          LExID = TRec!EXID:          LCalval = TRec!CALVAL
        
        LBuyQty = 0: LSellQty = 0: LBuyAmt = 0: LSellAmt = 0: LBrokAmt = 0
        
        Do While LACCID = TRec!ACCID And LSaudaID = TRec!SAUDAID
            If TRec!CONTYPE = "B" Then
                LBuyQty = TRec!LQTY:
                LBuyAmt = TRec!AMT
            Else
                LSellQty = TRec!LQTY
                LSellAmt = TRec!AMT
            End If
            TRec.MoveNext
            If TRec.EOF Then Exit Do
        Loop
        lOpQty = Round(LBuyQty - LSellQty, 2)
        LOpAmt = 0
        If lOpQty <> 0 Then
            LOpenRate = SDCLRATE(LSaudaID, vcDTP1.Value - 1, "O")
            LOpAmt = LOpenRate * lOpQty * LCalval
            mysql = "EXEC INSERT_BILLSUMOP " & LACCID & "," & LExID & "," & LSaudaID & "," & lOpQty & "," & LOpenRate & "," & LOpAmt & "," & LCalval & ""
            Cnn.Execute mysql
        End If
    Loop
        
    '>>> buy + sell + closing
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    mysql = " SELECT ACCID ,EXID,SAUDAID,CALVAL,CONTYPE,SUM(QTY) AS LQTY ,SUM(QTY*RATE*CALVAL) AS AMT ,"
    mysql = mysql & " ROUND(SUM( CASE BROKTYPE WHEN 'P' THEN (QTY*RATE*CALVAL)*(BROKRATE/100 ) WHEN 'O' THEN BROKQTY*BROKRATE  WHEN 'T' THEN QTY*BROKRATE END),2) AS BROKAMT FROM CTR_D "
    mysql = mysql & " Where  COMPCODE = " & GCompCode & "  "
    mysql = mysql & " AND SAUDAID IN (SELECT SAUDAID FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
    mysql = mysql & " AND CONDATE >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND CONDATE <='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
    If LenB(LExIDS) > 1 Then mysql = mysql & " AND EXID IN (" & LExIDS & ") "
    mysql = mysql & " GROUP BY ACCID , EXID,SAUDAID,CALVAL,CONTYPE"
    mysql = mysql & " ORDER BY ACCID , EXID,SAUDAID,CONTYPE"
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Do While Not TRec.EOF
        LACCID = TRec!ACCID:         LSaudaID = TRec!SAUDAID:          LExID = TRec!EXID:          LCalval = TRec!CALVAL
        LBuyQty = 0:        LSellQty = 0:         LBuyAmt = 0:        LSellAmt = 0:        LBrokAmt = 0
        
        Do While LACCID = TRec!ACCID And LSaudaID = TRec!SAUDAID
            If TRec!CONTYPE = "B" Then
                LBuyQty = TRec!LQTY
                LBuyAmt = TRec!AMT
                LBrokAmt = LBrokAmt + Val(TRec!BROKAMT & vbNullString)
            Else
                LSellQty = TRec!LQTY
                LSellAmt = TRec!AMT
                LBrokAmt = LBrokAmt + Val((TRec!BROKAMT & vbNullString))
            End If
            TRec.MoveNext
            If TRec.EOF Then Exit Do
        Loop
        LBrokAmt = Round(LBrokAmt * -1, 2)
        mysql = "EXEC INSERT_BILLSUMTRD " & LACCID & "," & LExID & "," & LSaudaID & "," & LBuyQty & "," & LBuyAmt & "," & LSellQty & "," & LSellAmt & "," & LBrokAmt & "," & LCalval & ""
        Cnn.Execute mysql
    Loop
    
'Closing
    'Set TRec = Nothing
   ' Set TRec = New ADODB.Recordset
  '  MYSQL = " SELECT ACCID,EXID,SAUDAID,CALVAL,CONTYPE,SUM(QTY) AS LQTY ,SUM(QTY*RATE*CALVAL) AS AMT "
 '   MYSQL = MYSQL & " FROM CTR_D "
'    MYSQL = MYSQL & " Where  COMPCODE = " & GCompCode & " "
    'MYSQL = MYSQL & " AND SAUDAID IN (SELECT SAUDAID FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "')"
   ' MYSQL = MYSQL & " AND CONDATE <='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
  '  If LenB(LExIDS) > 1 Then MYSQL = MYSQL & " AND EXID IN (" & LExIDS & ") "
 '   MYSQL = MYSQL & " GROUP BY ACCID , EXID,SAUDAID,CALVAL,CONTYPE"
'    MYSQL = MYSQL & " ORDER BY ACCID , EXID,SAUDAID,CONTYPE"'
    'TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
   ' Do While Not TRec.EOF
  '      LACCID = TRec!ACCID:          LSaudaID = TRec!SAUDAID
 '       LExID = TRec!EXID
 '       LCalval = TRec!CALVAL
 '       LBuyQty = 0: LSellQty = 0: LBuyAmt = 0: LSellAmt = 0: LBrokAmt = 0
 '
 '       Do While LACCID = TRec!ACCID And LSaudaID = TRec!SAUDAID
 '           If TRec!CONTYPE = "B" Then
 '               LBuyQty = TRec!LQTY:
 '               LBuyAmt = TRec!AMT
 '           Else
 '               LSellQty = TRec!LQTY
 '               LSellAmt = TRec!AMT
 '           End If
 '           TRec.MoveNext
 '           If TRec.EOF Then Exit Do
 '       Loop
 '       LClQty = Round(LBuyQty - LSellQty, 2)
 '       'LAMT = 0
 '
 '
 '       If LClQty <> 0 Then
 '           'LOpenRate = SDCLRATE(LSaudaID, vcDTP1.Value - 1, "O")
 '           'LOpAmt = LOpenRate * lOpQty * LCalval
 '           MYSQL = "EXEC INSERT_BILLSUMCL " & LACCID & "," & LExID & "," & LSaudaID & "," & LClQty & "," & LOpenRate & "," & LOpAmt & "," & LCalval & ""
 '           Cnn.Execute MYSQL
 '       End If
 '   Loop


'
    DoEvents
    Label4.Caption = "Creating New Bill Summ Complete "
    DoEvents
End Sub

Private Sub Update_ClRate()
    Dim TRec As ADODB.Recordset
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    DoEvents
    Label4.Caption = "Getting New LTP "
    DoEvents
    Timer1.Enabled = False
    mysql = "SELECT SAUDAID,CLOSERATE FROM CTR_R WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "'"
    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Do While Not TRec.EOF
        mysql = "UPDATE BILLSUMREC SET CLRATE =" & TRec!CLOSERATE & " WHERE SAUDAID =" & TRec!SAUDAID & ""
        Cnn.Execute mysql
        TRec.MoveNext
    Loop
    
    mysql = "UPDATE BILLSUMREC SET CLAMT = (CLQTY*CLRATE*CALVAL) WHERE CLQTY<>0"
    Cnn.Execute mysql
    
    
    
    DoEvents
    Label4.Caption = "Getting New LTP Complete "
    DoEvents
    Timer1.Enabled = True
End Sub

Public Sub GET_JCnn(LPDSource As String)
    Set Jcnn = Nothing
    Set Jcnn = New ADODB.Connection
    Jcnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & App.Path & LPDSource & _
    "Extended Properties=""TEXT;HDR=No;IMEX=1;FMT=Delimited"""
End Sub

Private Sub MSFRPT_EnterCell()
    Gridbackflag = False
    If MSFRPT.Col < 16 And MSFRPT.Row > 0 And MSFRPT.TextMatrix(MSFRPT.Row, 0) <> "Total" And Gridflag = False Then
        If MSFRPT.CellBackColor = vbBlack Then
            Gridbackflag = True
            MSFRPT.CellBackColor = vbHighlight
        End If
    End If
End Sub
Private Sub MSFRPT_LeaveCell()
    If MSFRPT.Col < 16 And MSFRPT.Row > 0 And MSFRPT.TextMatrix(MSFRPT.Row, 0) <> "Total" And Gridflag = False Then
        If Gridbackflag Then
            MSFRPT.CellBackColor = vbBlack
        End If
    End If
End Sub

Private Sub MSFdetail_EnterCell()
    If MSFdetail.Col < 12 Then MSFdetail.CellBackColor = vbHighlight
End Sub
Private Sub MSFdetail_LeaveCell()
    If MSFdetail.Col < 12 Then MSFdetail.CellBackColor = vbBlack
End Sub

Private Sub Timer1_Timer()
If Check1.Value = 0 Then
    I = I + 1
    Dim L As Integer
    L = Round((500 - I) / 10)
    DoEvents
    Label4.Caption = " Get data in next " & L & " sec."
    DoEvents
    If I = 500 Then
        Timer1.Enabled = False
        'Call CmdUpd_Click
        Get_BillSum
        I = 0
        'MsgBox "TIMER DONE"
        Timer1.Enabled = True
    End If
End If

End Sub
