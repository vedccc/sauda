VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCrDe 
   Caption         =   "Correction/Dividend Entry"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   24
      Top             =   840
      Width           =   15135
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   15
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   14895
         Begin VB.CommandButton CmdSave 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   13320
            TabIndex        =   13
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Txtnarr 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            MaxLength       =   150
            TabIndex        =   12
            Top             =   840
            Width           =   11700
         End
         Begin VB.TextBox Txttype 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   11160
            MaxLength       =   2
            TabIndex        =   11
            Text            =   "Dr"
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox TxtPtyCode 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   9
            Top             =   120
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   390
            Left            =   4230
            TabIndex        =   10
            Top             =   120
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox TXTROWNO 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   12000
            TabIndex        =   36
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Party"
            Height          =   375
            Left            =   3480
            TabIndex        =   35
            Top             =   128
            Width           =   615
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   15720
            TabIndex        =   31
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dr/Cr"
            Height          =   375
            Left            =   10560
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Party Code"
            Height          =   495
            Left            =   240
            TabIndex        =   29
            Top             =   195
            Width           =   1335
         End
         Begin VB.Label Label30 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Narration"
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   14895
         Begin VB.TextBox Txtamt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   11160
            TabIndex        =   4
            Top             =   120
            Width           =   1440
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "Cancel"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   13800
            TabIndex        =   6
            Top             =   120
            Width           =   1000
         End
         Begin VB.CommandButton CmdModify 
            Caption         =   "Modify"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   13800
            TabIndex        =   8
            Top             =   135
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12720
            TabIndex        =   5
            Top             =   120
            Width           =   1000
         End
         Begin VB.ComboBox InstCombo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "FrmCrDe.frx":0000
            Left            =   7800
            List            =   "FrmCrDe.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   120
            Width           =   2415
         End
         Begin vcDateTimePicker.vcDTP CDdate 
            Height          =   360
            Left            =   960
            TabIndex        =   1
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   3
            Value           =   41160.4222453704
         End
         Begin MSDataListLib.DataCombo DComboTSauda 
            Height          =   390
            Left            =   3720
            TabIndex        =   2
            Top             =   120
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   375
            Left            =   10320
            TabIndex        =   37
            Top             =   135
            Width           =   720
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   375
            Left            =   7200
            TabIndex        =   34
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
            Height          =   375
            Left            =   3000
            TabIndex        =   33
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   120
            Width           =   495
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Correction/Dividend Entries"
      TabPicture(0)   =   "FrmCrDe.frx":0024
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.CommandButton Command9 
         Caption         =   "Delete All Entries"
         Height          =   400
         Left            =   12840
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   -120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6135
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   14655
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7350
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   12965
            _Version        =   393216
            AllowUpdate     =   0   'False
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
         Begin VB.ComboBox cmbtype 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "FrmCrDe.frx":0040
            Left            =   8520
            List            =   "FrmCrDe.frx":004D
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   2775
         End
         Begin MSDataListLib.DataCombo DComboSauda 
            Height          =   390
            Left            =   4680
            TabIndex        =   17
            Top             =   120
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DComboParty 
            Height          =   390
            Left            =   1680
            TabIndex        =   18
            Top             =   120
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DComboCode 
            Height          =   390
            Left            =   720
            TabIndex        =   20
            Top             =   120
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   688
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label16 
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
            Left            =   120
            TabIndex        =   23
            Top             =   180
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
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
            Left            =   4080
            TabIndex        =   22
            Top             =   180
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label31 
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
            Height          =   375
            Left            =   7680
            TabIndex        =   21
            Top             =   180
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   6975
      Left            =   -120
      TabIndex        =   32
      Top             =   1920
      Width           =   15255
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Correction/Dividend Entry"
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
      TabIndex        =   0
      Top             =   120
      Width           =   15135
   End
End
Attribute VB_Name = "FrmCrDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LFExCode As String:                  Dim LFParty As String:              Dim LFBroker As String:            Dim LBillExCodes As String
Dim LBillParties As String:              Dim LBillSaudas As String:          Dim LSExCodes As String:           Dim LSPNames As String
Dim LSType As String:                    Dim LSUserIds As String:            Dim LItemCodeDBCombo As String:    Dim LFSauda As String
Dim LFBPress As Integer:                 Dim LPDataImport As Byte:           Dim SaveCalled As Boolean:         Dim LBillItems As String
Dim LOldParty As String:                 Dim LOldBroker As String:           Dim LOldContype As String:         Dim LOldSauda As String
Dim LOldEXCode As String:                Dim LOldRate2 As Double:            Dim LOldQty As Double:             Dim LOldRate As Double
Dim LOldConno As Long:                   Dim ExRec As ADODB.Recordset:       Dim PartyRec As ADODB.Recordset:   Dim ItemRec As ADODB.Recordset
Dim AllSaudaRec As ADODB.Recordset:      Dim SaudaRec As ADODB.Recordset:    Dim LFPartyRec As ADODB.Recordset: Dim ContRec As ADODB.Recordset
Dim LFSaudaRec As ADODB.Recordset:       Dim LFBrokerRec As ADODB.Recordset: Dim LListSaudas As String:         Dim LlistParties As String
Dim LFExID As Integer


'Public Sub ShowStanding()
'Dim NStandRec As ADODB.Recordset
'If LenB(LFParty) > 1 Then
'    mysql = "EXEC Get_PartyNetQtyPARTY " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "'," & LFExID & ",'" & LFParty & "' "
'Else
'    mysql = "EXEC Get_PartyNetQty " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "'," & LFExID & " "
'End If
'Set NStandRec = Nothing: Set NStandRec = New ADODB.Recordset
'NStandRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'If Not NStandRec.EOF Then
'    Set DataGrid2.DataSource = NStandRec
'    DataGrid2.ReBind
'    DataGrid2.Refresh
'    DataGrid2.Columns(0).Width = 3000:
'    DataGrid2.Columns(1).Width = 3000
'    DataGrid2.Columns(2).Width = 1200
'    DataGrid2.Columns(2).Alignment = dbgRight:
'    DataGrid2.Columns(2).NumberFormat = "0.00"
'End If
'End Sub
Private Sub cmbtype_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub cmbtype_Validate(Cancel As Boolean)
 Call DATA_GRID_REFRESH
End Sub

Private Sub CmdSave_Click()
    Dim LExCode As String:          Dim LDelFlag As Boolean:        Dim LOConNo As String:      Dim LContime As String:     Dim LCSauda As String
    Dim LCItemCode As String:       Dim LConType As String:         Dim LSInstType As String:   Dim LStatus As String:      Dim LST_Time As String
    Dim mparty As String:           Dim LCLot As Double:            Dim LCRefLot As Double:     Dim LCBrokLot As Double:    Dim LSCondate As Date
    Dim LConNo As Long:             Dim LClient As String:          Dim LExCont As String:      Dim MSaudaCode As String:   Dim LItemCode As String
    Dim MQty As Double:             Dim MRate As Double:            Dim LCalval As Double:      Dim MConRate As Double:     Dim LSConSno As Long:
    Dim LSOptType  As String:       Dim LSStrike As Double:         Dim LBSParty As String:     Dim LBrokFlag As String:    Dim LATime As String
    Dim TRec As ADODB.Recordset:    Dim NRec As ADODB.Recordset:    Dim LOrdNo As String:       Dim LSaudaID As Long
    Dim LExID As Integer:           Dim LItemID  As Integer
    
    On Error GoTo err1
    LDelFlag = False
    DoEvents
    LSCondate = CDdate.Value
    If GSysLockDt > LSCondate Then
        MsgBox "Can Not Add/Modfify/ Delete Trades. Settlement Locked Till " & GSysLockDt & ""
        Exit Sub
    End If
 
    Frame1.Enabled = False
    If LenB(TxtPtyCode.text) = 0 Then
        MsgBox "Party Code can not be Blank":        Frame1.Enabled = True
        TxtPtyCode.SetFocus
        Exit Sub
    Else
        mparty = Get_AccountDCode(TxtPtyCode.text)
        If LenB(mparty) < 1 Then
            MsgBox "Invalid Party Code":            Frame1.Enabled = True
            TxtPtyCode.SetFocus
            Exit Sub
        Else
            LClient = mparty
        End If
    End If
   
    
    LCSauda = vbNullString
    If LenB(DComboTSauda.BoundText) = 0 Then
        MsgBox "Sauda Code can not be Blank"
        Frame1.Enabled = True
        TxtPtyCode.SetFocus:
        Exit Sub
    End If
        
    If Val(Txtamt.text) = 0 Then
        MsgBox "Amount can not be Zero ":
        Frame1.Enabled = True
        Txtamt.SetFocus
        Exit Sub
    End If
            
    If Txttype.text = "Dr" Then
        LConType = "D"
    Else
        LConType = "C"
    End If
    
    Dim VCD As String
    VCD = "D"
    If InstCombo.text = "Correction" Then
        VCD = "C"
    End If
        
    DoEvents
    CNNERR = True
    Cnn.BeginTrans
    
        If LFBPress = 2 Then
            mysql = "DELETE FROM INV_AD WHERE COMPCODE =" & GCompCode & " AND ROWNO='" & TXTROWNO.text & "' " '>>>CONDATE ='" & Format(CDdate.Value, "YYYY/MM/DD") & "'"
            Cnn.Execute mysql
        End If
                    
        mysql = "INSERT INTO INV_AD (CompCode,CDDATE,PARTY,SAUDA,CD,CRDR,AMOUNT,NARRATION) VALUES "
        mysql = mysql & "('" & GCompCode & "','" & Format(CDdate.Value, "YYYY/MM/DD") & "','" & TxtPtyCode.text & "','" & DComboTSauda.BoundText & "','" & VCD & "','" & LConType & "','" & Txtamt.text & "','" & Txtnarr.text & "')"
        Cnn.Execute mysql
                
    Cnn.CommitTrans
    CNNERR = False
    
    Call DATA_GRID_REFRESH

    Txtnarr.text = vbNullString:
    Txtamt.text = vbNullString:
    Frame1.Enabled = True
    DataCombo2.BoundText = vbNullString:        DComboTSauda.BoundText = vbNullString
    
    CmdCancel.SetFocus
    SaveCalled = True
    Exit Sub
err1:
If err.Number <> 0 Then
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    Frame1.Enabled = True
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End If
End Sub
Private Sub CmdAdd_Click()
   
    If DComboTSauda.BoundText = "" Then
        MsgBox "Please select sauda", vbCritical
        DComboTSauda.SetFocus
        Exit Sub
    ElseIf Txtamt.text = "" Then
        MsgBox "Please enter amount", vbCritical
        Txtamt.SetFocus
        Exit Sub
    End If
    
'    TxtPtyCode.text = vbNullString: DataCombo2.BoundText = vbNullString: Txtnarr.text = vbNullString: Txtamt.text = vbNullString:
'    CmdModify.Enabled = False: Frame2.Enabled = True:
    CmdAdd.Enabled = False: CmdCancel.Enabled = True: CDdate.Enabled = False
    LFBPress = 1:
    LPDataImport = "0"
    'TxtPtyCode.SetFocus
    Call CALCULATE_STANDING
End Sub
Sub CALCULATE_STANDING()
    
    Dim NStandRec As ADODB.Recordset
    mysql = "EXEC Get_PartySaudaStd '" & GCompCode & "','" & Format(CDdate.Value, "YYYY/MM/DD") & "','" & DComboTSauda.BoundText & "','" & Txtamt.text & "' "

    Set NStandRec = Nothing: Set NStandRec = New ADODB.Recordset
    NStandRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not NStandRec.EOF Then
        Set DataGrid1.DataSource = NStandRec
        DataGrid1.ReBind
        DataGrid1.Refresh
        
        Call Resize_Grid
        
        Cnn.BeginTrans
        If BILL_GENERATION(CDdate.Value, GFinEnd, DComboTSauda.BoundText, "", "") Then
            Cnn.CommitTrans
            CNNERR = False
        
        End If
        GETMAIN.Label1.Caption = ""
        
        Command9.Visible = True
    Else
        Set DataGrid1.DataSource = Nothing
        DataGrid1.Refresh
        Command9.Visible = False
    End If
End Sub
Private Sub CmdModify_Click()
    '---------------
    'Call FillTradeSaudaCombo
    'Call FillLFPartyCombo
    'Call FillFSaudaCombo
    Call DATA_GRID_REFRESH
    'Call FillFBrokerCombo
    'If GShowStd = "Y" Then Call ShowStanding
    CDdate.Enabled = False
    'ChkShowContract.Enabled = False
    InstCombo.Enabled = False
    
    '---------------
    
    Call Mod_Rec
End Sub
Private Sub CmdCancel_Click()
    Call CANCEL_REC
End Sub
Private Sub Command1_Click()
'Dim I As Integer
'LListSaudas = vbNullString
'For I = 1 To ListView1.ListItems.Count
'    If ListView1.ListItems(I).Checked = True Then
'        If LenB(LListSaudas) > 1 Then
'            LListSaudas = LListSaudas & ",'" & ListView1.ListItems(I).text & "'"
'        Else
'            LListSaudas = "'" & ListView1.ListItems(I).text & "'"
'        End If
'    End If
'Next
Call DATA_GRID_REFRESH
End Sub

Private Sub Command2_Click()

'Dim I As Integer
'LlistParties = vbNullString
'For I = 1 To ListView2.ListItems.Count
'    If ListView2.ListItems(I).Checked = True Then
'        If LenB(LlistParties) > 1 Then
'            LlistParties = LlistParties & ",'" & ListView2.ListItems(I).text & "'"
'        Else
'            LlistParties = "'" & ListView2.ListItems(I).text & "'"
'        End If
'    End If
'Next
Call DATA_GRID_REFRESH

End Sub
Private Sub Command9_Click()

If MsgBox("Are You Sure You Want to Delete all entries", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then

    mysql = "DELETE FROM CTR_D  WHERE CONDATE = '" & Format(CDdate.Value, "YYYY/MM/DD") & "' and SAUDAID='" & DComboTSauda.BoundText & "' AND USERID='CorrDiv'  "
    Cnn.Execute mysql
    
    DATA_GRID_REFRESH
  
    Cnn.BeginTrans
    If BILL_GENERATION(CDdate.Value, GFinEnd, DComboTSauda.BoundText, "", "") Then
        Cnn.CommitTrans
        CNNERR = False
    End If
    GETMAIN.Label1.Caption = ""
End If

'Dim LDel As Boolean
'Dim LATime As String
'Dim LSCondate As Date
'Dim LSaudaID As Double
'If ContRec.RecordCount > 0 Then
'    LSCondate = CDdate.Value
'    If GSysLockDt > LSCondate Then
'        MsgBox "Can Not Add/Modfify/Delete Trades. Settlement Locked Till " & GSysLockDt & ""
'        Exit Sub
'    End If
'    'If OptTrade.Value Then
'        If MsgBox("Are You Sure You Want to Delete all Trades of " & CDdate.Value & " of " & "", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
'            If Not ContRec.EOF Then
'                ContRec.MoveFirst
'                Do While Not ContRec.EOF
'
'                    'sACHIN -- to move contrat entry in log table before edit
'                    'mysql = "INSERT INTO CTR_D_LOG (CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,loginuser,datetm,[tran])  "
'                    'mysql = mysql & "SELECT CompCode,CONSNO,CONNO,PARTY,SAUDA,ITEMCODE,CONDATE,CONTYPE,QTY,RATE,PATTAN,BROKRATE,TRANRATE,BILLNO,ROWNO1,PERCONT,INVNO,dataimport,ROWNO,CONTIME,BrokType,TranType,SrvTax,UserId,BrokQty,BrokAmt,CLCODE,STMRATE,BROKRATE2,ORDNO,STTRATE,NOTENO,TRANTAX,CONCODE,EXCODE,CALVAL,CONFIRM,ADJQTY,INSTTYPE,OPTTYPE,STRIKE,UPDBROK,BrokQty2,CGSTRATE,IGSTRATE,ORDTIME,SGSTRATE,SBC_TAX,SEBITAX,UTTRATE,UPDBQTY,SAUDAID,BROKFLAG,MULTI,FILETYPE,EXID,ITEMID,ACCID,EQ_STT,EQ_STAMP,ContractType,'" & GUserName & "',getdate(),'3' FROM CTR_D  "
'                    'mysql = mysql & "WHERE COMPCODE =" & GCompCode & " AND CONNO=" & ContRec!CONNO & "  AND CONDATE ='" & Format(CDdate.Value, "YYYY/MM/DD") & "'"
'                    'Cnn.Execute mysql
'
'                    mysql = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND EXCODE ='" & ContRec!excode & "' AND CONNO = " & ContRec!CONNO & " AND CONDATE ='" & Format(CDdate.Value, "YYYY/MM/DD") & "'"
'                    Cnn.Execute mysql
'                    LATime = CStr(Date) & " " & CStr(Time)
'
'                    LSaudaID = Get_SaudaID(ContRec!Sauda)
'
'                    If LenB(LBillParties) < 1 Then
'                        LBillParties = "'" & ContRec!PARTY & "','" & ContRec!Code & "'"
'                     '   If TRec!PARTY <> TRec!CONCODE Then
'                     '       LBillParties = "'" & TRec!PARTY & "','" & TRec!CONCODE & "'"
'                     '   Else
'                     '       LBillParties = "'" & TRec!PARTY & "'"
'                     '   End If
'                    Else
'                        If InStr(LBillParties, "'" & ContRec!PARTY & "'") < 1 Then LBillParties = LBillParties & ",'" & ContRec!PARTY & "'"
'                        If InStr(LBillParties, "'" & ContRec!Code & "") < 1 Then LBillParties = LBillParties & ",'" & ContRec!Code & "'"
'                    End If
'
'                   ' If LenB(LBillItems) < 1 Then
'                   '     LBillItems = "'" & TRec!ITEMCODE & "'"
'                   ' Else
'                   '     If InStr(LBillItems, TRec!ITEMCODE) < 1 Then LBillItems = LBillItems & "," & "'" & TRec!ITEMCODE & "'"
'                   ' End If
'
'                    If LenB(LBillSaudas) < 1 Then
'                        LBillSaudas = Trim(str(LSaudaID))
'                    Else
'                        If LStr_Exists(LBillSaudas, LSaudaID) = False Then LBillSaudas = LBillSaudas & "," & Trim(str(LSaudaID)) & ""
'                    End If
'
'                    'Call PInsert_Ctr_Log(ContRec!CONNO, ContRec!TRADENO, "Delete", LATime, CDdate.Value, ContRec!EXCODE, ContRec!Sauda, ContRec!PARTY, ContRec!BROKER, ContRec!BS, ContRec!QTY, ContRec!Rate, ContRec!CONRATE, GUserName)
'                    ContRec.MoveNext
'                Loop
'                SaveCalled = True
'                Call CANCEL_REC
'            End If
'            'Call Delete_Inv_D(vbNullString, vbNullString, vbNullString, CDdate.Value)
'        End If
'   ' End If
'    DATA_GRID_REFRESH
'End If
End Sub

Private Sub DataCombo2_Validate(Cancel As Boolean)
'Dim NRec As ADODB.Recordset
Dim LAcCode As String

If LenB(DataCombo2.text) = 0 Then
    MsgBox "Party can not be blank"
    Cancel = True
    Sendkeys "%{DOWN}"
Else
    LAcCode = Get_AccountDCode(DataCombo2.BoundText)
    If LenB(LAcCode) > 1 Then
        TxtPtyCode.text = LAcCode
        If Frame2.Enabled = False Then
            Frame10.Enabled = True
            CDdate.Enabled = True
            CDdate.SetFocus
        Else
           DComboTSauda.SetFocus
        End If
    Else
        DataCombo2.SetFocus
        Cancel = True
        Sendkeys "%{DOWN}"
    End If
End If
End Sub
Private Sub DComboTSauda_Validate(Cancel As Boolean)
'DoEvents
'
'Dim LSaudaID As Long
'If LenB(DComboTSauda.text) = 0 Then
'   ' MsgBox "Sauda can not be blank"
'    'Cancel = True
'  '  Sendkeys "%{DOWN}"
'Else
'    Call Get_Value
'End If
'    LSaudaID = Get_SaudaID(DComboTSauda.BoundText)
'
'    If GCINNo = "3000" Then Call Get_Value
'DoEvents

End Sub
Private Sub DComboParty_Validate(Cancel As Boolean)
If LenB(DComboParty.BoundText) <> 0 Then
    LFParty = DComboParty.BoundText
Else
    LFParty = vbNullString
End If
'If GShowStd = "Y" Then Call ShowStanding
Call FillFSaudaCombo
Call FillFBrokerCombo
Call DATA_GRID_REFRESH
End Sub

Private Sub DComboCode_Validate(Cancel As Boolean)
If LenB(DComboCode.BoundText) <> 0 Then
    LFParty = DComboCode.BoundText
    DComboParty.BoundText = LFParty
Else
    LFParty = vbNullString
End If
'If GShowStd = "Y" Then Call ShowStanding
Call FillFSaudaCombo
Call FillFBrokerCombo
Call DATA_GRID_REFRESH
End Sub

Private Sub DComboSauda_Validate(Cancel As Boolean)
FillFBrokerCombo
'If LenB(DComboSauda.BoundText) <> 0 Then
'    LFSauda = DComboSauda.BoundText
'Else
'    LFSauda = vbNullString
'End If
'Set LFBrokerRec = Nothing
'Set LFBrokerRec = New ADODB.Recordset
'MYSQL = "SELECT DISTINCT A.AC_CODE,A.NAME FROM ACCOUNTD  AS A, CTR_D AS B WHERE A.COMPCODE =" & GCompCode & " AND A.COMPCODE =B.COMPCODE"
'MYSQL = MYSQL & " AND A.AC_CODE  =B.CONCODE  AND B.CONDATE ='" & Format(CDdate.Value, "YYYY/MM/DD") & "'"
'If LenB(LFExCode) > 0 Then MYSQL = MYSQL & " AND B.EXCODE='" & LFExCode & "'"
'If LenB(LFParty) > 0 Then MYSQL = MYSQL & " AND B.PARTY ='" & LFParty & "'"
'If LenB(LFSauda) > 0 Then MYSQL = MYSQL & " AND B.SAUDA ='" & LFSauda & "'"
'MYSQL = MYSQL & " ORDER BY A.NAME"

'Call DATA_GRID_REFRESH
End Sub
Private Sub DataGrid1_DblClick()
    Dim LPConNo As Long:         Dim LPSauda As String:      Dim LPConType As String:        Dim TRec As ADODB.Recordset
'    If GCINNo = "2000" Then
'        DataGrid1.Col = 4:              LPSauda = DataGrid1.text
'        DataGrid1.Col = 10:             LPConNo = DataGrid1.text
'        DataGrid1.Col = 2:              LPConType = DataGrid1.text
'        'DataGrid1.Col = 9:              LPDataImport = Trim(DataGrid1.text)
'    Else
        DataGrid1.Col = 7:                 TXTROWNO.text = DataGrid1.text
'        DataGrid1.Col = 10:             LPConNo = DataGrid1.text
'        DataGrid1.Col = 3:              LPConType = DataGrid1.text
'    End If
'    TXTROWNO.text = LPConNo
    Call Mod_Rec
    Call Get_Trade_Details(LPConNo)
    
    CmdAdd.Enabled = True:                          CmdModify.Enabled = False
    CmdCancel.Enabled = True:                       CDdate.Enabled = False
    
    Frame2.Enabled = True:                          LFBPress = 2
    Label12.Caption = "Modifty Trade"
    
    If (GSysLockDt >= CDdate.Value) Then
        CmdAdd.Enabled = False
        CmdModify.Enabled = False
        CmdSave.Enabled = False
    End If
    
    'TxtConNo.SetFocus
'End If
Set TRec = Nothing
End Sub

'Private Sub DataGridOrder_DblClick()
'    Dim LPConNo As Long:         Dim LPSauda As String:      Dim LPConType As String:        Dim TRec As ADODB.Recordset
''    If GCINNo = "2000" Then
''        DataGrid1.Col = 4:              LPSauda = DataGrid1.text
''        DataGrid1.Col = 10:             LPConNo = DataGrid1.text
''        DataGrid1.Col = 2:              LPConType = DataGrid1.text
''        'DataGrid1.Col = 9:              LPDataImport = Trim(DataGrid1.text)
''    Else
'        DataGridOrder.Col = 2:              LPSauda = DataGridOrder.text
'        DataGridOrder.Col = 10:             LPConNo = DataGridOrder.text
'        DataGridOrder.Col = 3:              LPConType = DataGridOrder.text
''    End If
'
'    Call Mod_Rec
'    Call Get_Trade_Details(LPConNo)
'
'
'
'    CmdAdd.Enabled = True:                          CmdModify.Enabled = False
'    CmdCancel.Enabled = True:                       CDdate.Enabled = False
'
'    Frame2.Enabled = True:                          LFBPress = 2
'    Label12.Caption = "Modifty Order"
'
'
'    CmdSave.Enabled = True
'
'
'    If (GSysLockDt >= CDdate.Value) Then
'        CmdAdd.Enabled = False
'        CmdModify.Enabled = False
'        CmdSave.Enabled = False
'    End If
'
'
'    'TxtConNo.SetFocus
''End If
'Set TRec = Nothing
'End Sub


Private Sub DataGrid3_DblClick()
Dim LPConNo As Long:            Dim LPSauda As String:          Dim LPConType As String:        Dim TRec As ADODB.Recordset
'    'DataGrid3.Col = 2:          LPSauda = DataGrid3.text
'    'DataGrid3.Col = 7:          LPConNo = DataGrid3.text
'    'DataGrid3.Col = 3:          LPConType = DataGrid3.text
'    Call Mod_Rec
'    MYSQL = "SELECT CONSNO,CONNO, QTY,RATE,PARTY,CONTYPE,SAUDA,ITEMCODE,EXCODE,CONCODE,STATUS,ST_TIME FROM CTR_L "
'    MYSQL = MYSQL & " WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(CDdate.Value, "YYYY/MM/DD") & "'AND SAUDA='" & LPSauda & "'AND CONNO=" & LPConNo & ""
'    Set TRec = Nothing:    Set TRec = New ADODB.Recordset
'    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Not TRec.EOF Then
'        Do While Not TRec.EOF
'            TxtConNo.text = TRec!CONNO:                         TxtPtyCode.text = TRec!PARTY
'            DataCombo2.BoundText = TRec!PARTY:                  Txtamt.text = Format(TRec!Rate, "0.0000")
'            TxtBrokerCode.text = TRec!CONCODE:
'            DComboTSauda.BoundText = TRec!Sauda
'            If LPConType = "B" Then
'                Txttype.text = "Buy"
'            Else
'                Txttype.text = "Sel"
'            End If
'            Txtnarr.text = TRec!QTY
'            TRec.MoveNext
'        Loop
'        Get_Value
'        CmdAdd.Enabled = True:              CmdModify.Enabled = False
'        CmdCancel.Enabled = True:           CDdate.Enabled = False

'        Frame2.Enabled = True:              LFBPress = 2
'        Label12.Caption = "Modifying Existing Trades"
'        TxtConNo.SetFocus
'    End If
'    Set TRec = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "CDdate" Then
            Sendkeys "{tab}"
        End If
    End If
    'If CmdAdd.Enabled = False Then
    '   If KeyCode = 121 Then Frame4.Visible = True
    'End If
End Sub
Private Sub Form_Load()
Dim TRec As ADODB.Recordset
LPDataImport = 0:           LSExCodes = vbNullString:    LSPNames = vbNullString:    LSType = vbNullString
LSUserIds = vbNullString:   LFExCode = vbNullString:     LFParty = vbNullString:     LFSauda = vbNullString
LFExID = 0:                 LFBroker = vbNullString:
LListSaudas = vbNullString: LlistParties = vbNullString: SSTab1.Tab = 0
   
   
Call Connect_TSaudaCombo

    
CDdate.Value = Date
Frame2.Visible = True
'InstCombo.Clear
'InstCombo.AddItem "All"
'Set TRec = Nothing
'Set TRec = New ADODB.Recordset
'mysql = "SELECT DISTINCT INSTTYPE FROM SCRIPTMASTER  ORDER BY INSTTYPE"
'TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'If Not TRec.EOF Then
'    If TRec.RecordCount > 0 Then
'        TRec.MoveFirst
'        Do While Not TRec.EOF
'            InstCombo.AddItem (TRec!INSTTYPE)
'            TRec.MoveNext
'        Loop
'        InstCombo.Visible = True
'    End If
'Else
'    Set TRec = Nothing
'    Set TRec = New ADODB.Recordset
'    mysql = "SELECT DISTINCT INSTTYPE FROM SAUDAMAST ORDER BY INSTTYPE"
'    TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'    If Not TRec.EOF Then
'        If TRec.RecordCount > 0 Then
'            TRec.MoveFirst
'            Do While Not TRec.EOF
'                InstCombo.AddItem (TRec!INSTTYPE)
'                TRec.MoveNext
'            Loop
'            InstCombo.Visible = True
'        End If
'    End If
'End If
'InstCombo.ListIndex = 0:
'If TRec.RecordCount = 1 Then
'    InstCombo.Locked = True
'End If
'InstCombo.Visible = True

If GCINNo = "2000" Then
         DComboTSauda.Left = 7300
    Label7.Left = 5280:
              Txttype.Left = 5280
    Txtnarr.Left = 6080:             Txttype.TabIndex = 11
    Txtnarr.TabIndex = 12:           DComboTSauda.TabIndex = 13
ElseIf GUniqClientId = "2175AHM" Then
    Label7.Left = 1080:             Txttype.Left = 1080
    Txttype.TabIndex = 10:       Label27.Visible = False
          
    Txtnarr.Left = 1700:             Txtnarr.TabIndex = 12
           DComboTSauda.Left = 2800
    
    Label9.Left = 5800:             DComboTSauda.TabIndex = 13
    Txtamt.Left = 5800:            Txtamt.TabIndex = 14
    Label3.Left = 7200:             TxtPtyCode.Left = 7200
    TxtPtyCode.TabIndex = 15:
    DataCombo2.Left = 8150:         DataCombo2.TabIndex = 16
    
    
ElseIf GUniqClientId = "1207BIK" Then
    Label7.Left = 1200:             Txttype.Left = 1200
      Txtnarr.Left = 1845
           DComboTSauda.Left = 3000
    Label9.Left = 6600:             Txtamt.Left = 6600
    Label27.Visible = False:
    Label3.Left = 8100:             TxtPtyCode.Left = 8100
           DataCombo2.Left = 9100
          DataCombo2.Width = 3250
    
       
    
    Txttype.TabIndex = 8:        Txtnarr.TabIndex = 9
    Txtamt.TabIndex = 11:          DComboTSauda.TabIndex = 10
    TxtPtyCode.TabIndex = 12:       DataCombo2.TabIndex = 13
    'TxtBrokerCode.TabIndex = 14:
    
End If



'If cmbtype.Visible Then
    cmbtype.ListIndex = 0
'End If

End Sub
Private Sub DComboParty_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboCode_GotFocus()
    Sendkeys "%{DOWN}"
End Sub

Private Sub DComboSauda_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo2_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DComboTSauda_GotFocus()
    Sendkeys "%{DOWN}"
    If LenB(DComboTSauda.BoundText) > 0 Then
        DComboTSauda.BoundText = DComboTSauda.BoundText
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CANCEL_REC
End Sub


Private Sub InstCombo_Validate(Cancel As Boolean)
    Call DATA_GRID_REFRESH
End Sub

Private Sub TxtPtyCode_Validate(Cancel As Boolean)
'Dim NRec As ADODB.Recordset
Dim LAcCode As String
If Frame2.Enabled = True Then
    If LenB(TxtPtyCode.text) = 0 Then
        DataCombo2.SetFocus
    Else
        LAcCode = Get_AccountDCode(TxtPtyCode.text)
        If LenB(LAcCode) > 1 Then
            DataCombo2.BoundText = LAcCode
            DComboTSauda.SetFocus
        Else
            DataCombo2.SetFocus
        End If
    End If
Else
    Frame1.Enabled = True
    Frame10.Enabled = True
    CDdate.Enabled = True
    CDdate.SetFocus
End If
End Sub

Private Sub txttype_KeyPress(KeyAscii As Integer)
If Val(KeyAscii) >= 48 And KeyAscii <= 122 Then
    If Val(KeyAscii) = 66 Or Val(KeyAscii) = 98 Or Val(KeyAscii) = 83 Or Val(KeyAscii) = 115 Then
    Else
        If Txttype.text = "Dr" Then
            Txttype.text = "Cr"
        Else
            Txttype.text = "Dr"
        End If
    End If
End If
If KeyAscii = 32 Then
    If Txttype.text = "Dr" Then
        Txttype.text = "Cr"
    Else
        Txttype.text = "Dr"
    End If
End If
If KeyAscii = 43 Then Txttype.text = "Dr"
If KeyAscii = 45 Then Txttype.text = "Cr"
    
End Sub

Private Sub Txttype_Validate(Cancel As Boolean)
If Txttype.text <> "Dr" Then
    If Txttype.text <> "Cr" Then
        Txttype.text = "Dr"
        Cancel = True
        Txttype.SetFocus
    End If
End If
End Sub
Private Sub Txtamt_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Txtamt_Validate(Cancel As Boolean)
    Txtamt.text = Format(Txtamt.text, "0.0000")
    
 ' If GCINNo <> "3000" Then Call Get_Value
End Sub
Private Sub Txtamt_GotFocus()
    Txtamt.SelStart = 0
    Txtamt.SelLength = Len(Txtamt.text)
End Sub
Private Sub TxtPtyCode_GotFocus()
    TxtPtyCode.SelStart = 0
    TxtPtyCode.SelLength = Len(TxtPtyCode.text)
End Sub
Public Sub DATA_GRID_REFRESH()

    mysql = "SELECT D.PARTY,M.Name,D.CONTYPE, D.QTY,D.RATE,D.CONCODE , A.NAME "
    mysql = mysql & " FROM CTR_D D WITH(NOLOCK) , AccountM M, AccountM A  "
    mysql = mysql & " WHERE D.CONDATE = '" & Format(CDdate.Value, "YYYY/MM/DD") & "' and SAUDAID='" & DComboTSauda.BoundText & "' AND D.PARTY=D.CLCODE and D.USERID='CorrDiv' AND D.PARTY=M.AC_Code  AND D.CONCODE=A.AC_Code ORDER BY ROWNO "
    Set ContRec = Nothing
    Set ContRec = New ADODB.Recordset
    ContRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not ContRec.EOF Then
        Set DataGrid1.DataSource = ContRec
        DataGrid1.ReBind
        DataGrid1.Refresh
        Call Resize_Grid
        Command9.Visible = True
    Else
        Set DataGrid1.DataSource = Nothing
        DataGrid1.Refresh
        
        Command9.Visible = False
    End If
    

End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    'SORTING ***
    Dim LSortOrder As String
    On Error GoTo Error1
    DataGrid1.MarqueeStyle = dbgHighlightCell
    DoEvents
'    If Left$(Label13.Caption, 1) = "A" Then
'        LSortOrder = DataGrid1.Columns.Item(ColIndex).DataField & "  DESC"
'        ContRec.Sort = ("" & LSortOrder & "")
'        Label13.Caption = "Desc. ORDER ON  " & DataGrid1.Columns.Item(ColIndex).Caption
'    Else
        LSortOrder = DataGrid1.Columns.Item(ColIndex).DataField & "  ASC"
        ContRec.Sort = ("" & LSortOrder & "")
        
'    End If
    DoEvents
    If Not ContRec.EOF Then Set DataGrid1.DataSource = ContRec: DataGrid1.ReBind: DataGrid1.Refresh
    Call Resize_Grid
    
Error1:    Exit Sub
End Sub
'Private Sub DataGridOrder_HeadClick(ByVal ColIndex As Integer)
'    'SORTING ***
'    Dim LSortOrder As String
'    On Error GoTo Error1
'    '---------------
'    CDdate.Enabled = True
'    'ChkShowContract.Enabled = True
'    InstCombo.Enabled = True
'
'    '---------------
'
'    DataGridOrder.MarqueeStyle = dbgHighlightCell
'    DoEvents
''    If Left$(Label13.Caption, 1) = "A" Then
''        LSortOrder = DataGridOrder.Columns.Item(ColIndex).DataField & "  DESC"
''        ContRec.Sort = ("" & LSortOrder & "")
''        Label13.Caption = "Desc. ORDER ON  " & DataGridOrder.Columns.Item(ColIndex).Caption
''    Else
'        LSortOrder = DataGridOrder.Columns.Item(ColIndex).DataField & "  ASC"
'        ContRec.Sort = ("" & LSortOrder & "")
'
'    'End If
'    DoEvents
'    If Not ContRec.EOF Then Set DataGridOrder.DataSource = ContRec: DataGridOrder.ReBind: DataGrid1.Refresh
'    Call Resize_Grid
'
'Error1:    Exit Sub
'End Sub
Private Sub CANCEL_REC()

    Dim SREC As ADODB.Recordset:        Dim PREC As ADODB.Recordset
'Dim LBSaudas As String:                     Dim LBParties As String:            Dim LBItems As String

LFParty = vbNullString:                     LFExCode = vbNullString:            LFSauda = vbNullString
LFExID = 0
LFBroker = vbNullString:
LFExCode = vbNullString:
        Frame2.Enabled = False:             'TxtConfirm.text = "0"
                  SSTab1.Tab = 0:                     LListSaudas = vbNullString
LlistParties = vbNullString:                'ListView1.ListItems.Clear:          ListView2.ListItems.Clear
'TxtOptType.text = "CE"

Label12.Caption = "Updating Bills"
GETMAIN.Toolbar1_Buttons(6).Enabled = False
    On Error GoTo err1
    CmdAdd.Enabled = True:                      CmdModify.Enabled = True: CmdSave.Enabled = True
    CmdCancel.Enabled = False:                  CDdate.Enabled = True
      DataCombo2.BoundText = vbNullString
    DComboTSauda.BoundText = vbNullString:
    DComboParty.BoundText = vbNullString:       DComboSauda.BoundText = vbNullString
    DComboCode.BoundText = vbNullString
    'TxtConNo.text = vbNullString:               TxtPtyCode.text = vbNullString
    
    'ChkShowContract.Enabled = True
    'TxtConNo.Locked = False
    Frame2.Enabled = False
    LFBPress = 0:
    
'    If OptOrder.Value Then
'        SaveCalled = False
'    End If
    
'    If SaveCalled = True Then
'        Frame1.Enabled = False:
'        DataGrid1.Enabled = False
''        DataGrid2.Enabled = False
'        DataGridOrder.Enabled = False
'        Call RATE_TEST(CDdate.Value, , , frmcont4)
'
'        Call Shree_Posting(DateValue(CDdate.Value))
'
'        CNNERR = True:                 Cnn.BeginTrans
'        Call Update_Charges(LBillParties, LBillExCodes, LBillSaudas, vbNullString, CDdate.Value, CDdate.Value, True)
'        GETMAIN.Label1.Caption = "Updating Brokerage Rate Itemwise Complete"
'        Cnn.CommitTrans
'        Cnn.BeginTrans
'        If BILL_GENERATION(CDdate.Value, GFinEnd, LBillSaudas, LBillParties, LBillExCodes) Then
'            Cnn.CommitTrans
'            CNNERR = False
'        End If
'        'Call Chk_Billing
'        DataGrid1.Enabled = True
''       DataGrid2.Enabled = True
'        DataGridOrder.Enabled = True
'    End If
    LBillParties = vbNullString:    LBillExCodes = vbNullString
    LBillSaudas = vbNullString:     LBillItems = vbNullString
    SaveCalled = False:             Frame1.Enabled = True
    Frame3.Enabled = True:          Frame10.Enabled = True
    Frame2.Enabled = True:          GETMAIN.Toolbar1_Buttons(6).Enabled = True
    CDdate.Enabled = True:      CDdate.SetFocus
    
    'Label12.Caption = "Bills Updated"
    Call DATA_GRID_REFRESH
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    If CNNERR = True Then
        Cnn.RollbackTrans: CNNERR = False
        Frame1.Enabled = True
    End If
End Sub


Private Sub CDdate_Validate(Cancel As Boolean)
    Dim NRec As ADODB.Recordset
  '  Dim lstr  As String
    Dim MDt As String
    
    If (GSysLockDt >= CDdate.Value) Then
        CmdAdd.Enabled = False
        'CmdModify.Enabled = False
        CmdSave.Enabled = False
        If MsgBox("System is locked till date " & GSysLockDt & vbNewLine & "Do you still want to view document?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    Else
        CmdAdd.Enabled = True
        'CmdModify.Enabled = True
        CmdSave.Enabled = True
    End If
    
    Call Connect_TSaudaCombo
    
    MDt = Format(Now, "dd/mm/yyyy")
'
    If CDdate.Value > DateValue(MDt) Then
        MsgBox "Trade Date Is Greater than Current Date"
    End If
    'Call DATA_GRID_REFRESH
'
'    If InstCombo.ListIndex = -1 Then
'        InstCombo.ListIndex = 1
'    End If
    
    'Call FillLFPartyCombo
    'Call FillFSaudaCombo
    'Call FillFBrokerCombo
    'FillTradeSaudaCombo
End Sub
Private Sub Mod_Rec()
    'If ContRec.RecordCount > 0 Then
        Frame2.Enabled = True:              CmdAdd.Enabled = False
        CmdModify.Enabled = False:          CmdCancel.Enabled = True
        CDdate.Enabled = False:
        'ChkShowContract.Enabled = False:    'TxtConNo.text = vbNullString
        'TxtConNo.Locked = False:            Txtnarr.Locked = False
        Txtnarr.text = vbNullString:         Txtamt.text = vbNullString
        
        LFBPress = 2
        'Call Connect_TSaudaCombo
        Label12.Caption = "Modify Trades"
        'TxtConNo.SetFocus
'    Else
'        MsgBox "No Records to Modify "
'    End If
End Sub
Private Sub Get_Value()
     Call Connect_TSaudaCombo
    
    If 1 = 1 Then
        AllSaudaRec.Filter = adFilterNone
        If Not AllSaudaRec.EOF Then
            AllSaudaRec.MoveFirst
            AllSaudaRec.Find "SAUDACODE ='" & DComboTSauda.BoundText & "'"
            If AllSaudaRec.EOF Then
                MsgBox "Invalid Contract"
            
            End If
        Else
            MsgBox "No Contract"
        End If
    Else
        SaudaRec.Filter = adFilterNone
        If SaudaRec.RecordCount > 0 Then SaudaRec.MoveFirst
        SaudaRec.Find "SAUDACODE ='" & DComboTSauda.BoundText & "'", , adSearchForward
        If SaudaRec.EOF Then
            MsgBox "Invalid Contract"
            Sendkeys "%{DOWN}"
        Else
'            If SaudaRec!excode = "MCX" Or SaudaRec!excode = "NSE" Then
'                If SaudaRec!LOTWISE = "Y" Then
'                    TxtCalval.text = Format(SaudaRec!TRADEABLELOT, "0.00")
'                Else
'                    TxtCalval.text = Format(SaudaRec!lot, "0.00")
'                End If
'            Else
'
'                TxtCalval.text = Format(SaudaRec!lot, "0.00")
'            End If
            
            
        End If
    End If
    If DComboTSauda.Enabled Then
    DComboTSauda.SetFocus
    End If

End Sub
'Public Sub FillTradeSaudaCombo()
'    'MYSQL = " EXEC Get_ScriptContract " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "','" & LFExCode & "','" & InstCombo.text & "'"
'    'Set AllSaudaRec = Nothing
'    'Set AllSaudaRec = New ADODB.Recordset
'    'AllSaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'    '
'    'MYSQL = "EXEC Get_SaudaContract " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "'," & LFExID & ",'" & InstCombo.text & "'"
'    'Set SaudaRec = Nothing
'    'Set SaudaRec = New ADODB.Recordset
'    'SaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'
'    Call Connect_TSaudaCombo
'End Sub

Public Sub Connect_TSaudaCombo()

    mysql = " SELECT S.SAUDAID,S.SAUDANAME "
    mysql = mysql & "FROM SAUDAMAST AS S,EXMAST AS EX WHERE EX.EXCODE =S.EXCODE AND EX.COMPCODE =" & GCompCode & " AND EX.COMPCODE=S.COMPCODE and S.MATURITY>='" & Format(CDdate.Value, "YYYY/MM/DD") & "' "
    mysql = mysql & "ORDER BY S.SAUDANAME "
    Set AllSaudaRec = Nothing: Set AllSaudaRec = New ADODB.Recordset: AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AllSaudaRec.EOF Then
        Set DComboTSauda.RowSource = AllSaudaRec:
        DComboTSauda.BoundColumn = "SAUDAID"
        DComboTSauda.ListField = "SAUDANAME"
    End If
    
'    Set PartyRec = Nothing: Set PartyRec = New ADODB.Recordset
'    mysql = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " ORDER BY NAME"
'    PartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'    If Not PartyRec.EOF Then
'        Set DataCombo2.RowSource = PartyRec
'        DataCombo2.BoundColumn = "AC_CODE"
'        DataCombo2.ListField = "NAME"
'    End If


'If 1 = 1 Then

    'mysql = " EXEC Get_ScriptContract " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "','" & LFExCode & "','" & InstCombo.text & "','" & TxtOptType.text & "'," & Val(TxtStrike.text) & ""
    'mysql = " EXEC Get_ScriptContract " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "','" & LFExCode & "','" & InstCombo.text & "','" & TxtOptType.text & "'," & Val(TxtStrike.text) & ""
    
'    mysql = " SELECT S.SAUDACODE,S.SAUDANAME " ',S.ITEMCODE,C.ITEMNAME,S.MATURITY,S.EXCODE,S.INSTTYPE,S.OPTTYPE,S.STRIKEPRICE,C.LOT, S.LOT AS TRADEABLELOT,S.BROKLOT,S.REFLOT,EX.LOTWISE,S.EX_SYMBOL "
'    mysql = mysql & " FROM SAUDAMAST AS S,EXMAST AS EX WHERE EX.EXCODE =S.EXCODE AND EX.COMPCODE =" & GCompCode & ""
'    'mysql = mysql & " AND C.EX_SYMBOL=S.EX_SYMBOL AND C.EXCODE=S.EXCODE AND S.MATURITY>='" & Format(CDdate.Value, "YYYY/MM/DD") & "'"
'    'If LenB(LFExCode) > 0 Then mysql = mysql & " AND EX.EXCODE ='" & LFExCode & "'"
'    'If InstCombo.text <> "All" Then mysql = mysql & " AND  S.INSTTYPE ='" & InstCombo.text & "'"
'    'If InstCombo.text = "OPT" Then
'    '    If LenB(TxtOptType.text) > 0 Then mysql = mysql & " AND OPTTYPE ='" & TxtOptType.text & "'"
'    '    If Val(TxtStrike.text & vbNullString) <> 0 Then mysql = mysql & "  AND STRIKEPRICE =" & Val(TxtStrike.text & vbNullString) & ""
'    'End If
'    mysql = mysql & " ORDER BY S.SAUDANAME " ' ,S.INSTTYPE,S.OPTTYPE,S.STRIKEPRICE,S.MATURITY  "
'
'    Set AllSaudaRec = Nothing
'    Set AllSaudaRec = New ADODB.Recordset
'    AllSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'    If Not AllSaudaRec.EOF Then
'        Set DComboTSauda.RowSource = AllSaudaRec:
'        DComboTSauda.BoundColumn = "SAUDACODE"
'        DComboTSauda.ListField = "SAUDANAME"
'    Else
'        Set DComboTSauda.RowSource = Nothing
'        MsgBox "No Contract Exists "
'        If Frame2.Enabled Then
'            TxtPtyCode.SetFocus
'        Else
'            Frame10.Enabled = True
'            CDdate.SetFocus
'        End If
'    End If
    
'Else
'        mysql = "EXEC Get_SaudaContract " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "'," & LFExID & ""
'        Set SaudaRec = Nothing
'        Set SaudaRec = New ADODB.Recordset
'        SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
'        Set DComboTSauda.RowSource = SaudaRec
'        DComboTSauda.BoundColumn = "SAUDACODE"
'        DComboTSauda.ListField = "SAUDANAME"
'        If Not SaudaRec.EOF Then
'
'        Set DComboTSauda.RowSource = SaudaRec:
'            DComboTSauda.BoundColumn = "SAUDACODE"
'            DComboTSauda.ListField = "SAUDANAME"
'        Else
'            If LFExID <> -1 Then  ' If order then LFExID = -1
'                MsgBox "No Contract Exists "
'                Set DComboTSauda.RowSource = Nothing
'                If Frame2.Enabled Then
'                    TxtPtyCode.SetFocus
'                Else
'                    Frame10.Enabled = True
'                    CDdate.SetFocus
'                End If
'            End If
'        End If
'
'
'End If


End Sub

Private Sub Resize_Grid()

        DataGrid1.Columns(0).Width = 1500:
        DataGrid1.Columns(1).Width = 4000
        DataGrid1.Columns(2).Width = 1500
        DataGrid1.Columns(3).Width = 1500
        DataGrid1.Columns(4).Width = 1500
        DataGrid1.Columns(5).Width = 1500
        DataGrid1.Columns(6).Width = 4000

        DataGrid1.Columns(5).Alignment = dbgRight:
        DataGrid1.Columns(5).NumberFormat = "0.00"
    
End Sub

Private Sub Get_Trade_Details(LZConNo As Long)
    Dim TRec As ADODB.Recordset
'    Dim TRec1 As ADODB.Recordset
    Dim LAcCode As String

    mysql = "SELECT A.CDDATE,A.SAUDA,A.CD,A.PARTY,A.AMOUNT,A.CRDR,A.NARRATION,A.ROWNO "
    mysql = mysql & "FROM INV_AD A WHERE A.COMPCODE=" & GCompCode & " and ROWNO='" & TXTROWNO.text & "'"

    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not TRec.EOF Then
        CDdate.Value = TRec!CDdate
        TxtPtyCode.text = TRec!PARTY
        
        LAcCode = Get_AccountDCode(TxtPtyCode.text)
        If LenB(LAcCode) > 1 Then
            DataCombo2.BoundText = LAcCode
        End If
        
        DComboTSauda.BoundText = TRec!Sauda
        Txtamt.text = TRec!AMOUNT
        Txtnarr.text = TRec!NARRATION
        
        If TRec!CRDR = "D" Then
            Txttype.text = "Dr"
        Else
            Txttype.text = "Cr"
        End If
        
        If TRec!CD = "C" Then
            InstCombo.text = "Correction"
        Else
            InstCombo.text = "Dividend"
        End If
        TxtPtyCode.SetFocus
    End If
        
        
        
'        'TxtConNo.text = TRec!CONNO
'        'Text7.text = (TRec!ROWNO1 & vbNullString)
'
'        TxtPtyCode.text = TRec!PARTY
'        DataCombo2.BoundText = TRec!PARTY
'        Txtamt.text = Format(TRec!Rate, "0.0000")
'
'        'TxtBrokerCode.text = TRec!CONCODE
'
'        DComboTSauda.BoundText = TRec!Sauda
'
'        If TRec!CONTYPE = "B" Then
'            Txttype.text = "Dr"
'        Else
'            Txttype.text = "Cr"
'        End If
'        'TxtConfirm.text = TRec!CONFIRM
'        Txtnarr.text = TRec!QTY
'        LOldEXCode = TRec!excode
'        LOldParty = TRec!PARTY
'        LOldBroker = TRec!CONCODE
'        LOldQty = TRec!QTY
'        LOldContype = TRec!CONTYPE
'        LOldRate = TRec!Rate
'        LOldRate2 = TRec!BROKAMT
'        LOldSauda = TRec!Sauda
'        LOldConno = TRec!CONNO
'
'
'        'If OptTrade.Value Then
'            If LenB(LBillExCodes) < 1 Then
'                LBillExCodes = Str(TRec!EXID)
'            Else
'                If LStr_Exists(LBillExCodes, Str(TRec!EXID)) < 1 Then LBillExCodes = LBillExCodes & "," & Str(TRec!EXID) & ""
'            End If
''        ElseIf OptOrder.Value Then
''            LBillExCodes = "0"
''            CmbOrderStatus.ListIndex = TRec!EXID
'        'End If
'
'        If LenB(LBillParties) < 1 Then
'            If TRec!PARTY <> TRec!CONCODE Then
'                LBillParties = "'" & TRec!PARTY & "','" & TRec!CONCODE & "'"
'            Else
'                LBillParties = "'" & TRec!PARTY & "'"
'            End If
'        Else
'            If InStr(LBillParties, "'" & TRec!PARTY & "'") < 1 Then LBillParties = LBillParties & ",'" & TRec!PARTY & "'"
'            If InStr(LBillParties, "'" & TRec!CONCODE & "") < 1 Then LBillParties = LBillParties & ",'" & TRec!CONCODE & "'"
'        End If
'        If LenB(LBillItems) < 1 Then
'            LBillItems = "'" & TRec!ITEMCODE & "'"
'        Else
'            If InStr(LBillItems, TRec!ITEMCODE) < 1 Then LBillItems = LBillItems & "," & "'" & TRec!ITEMCODE & "'"
'        End If
'
'        If LenB(LBillSaudas) < 1 Then
'            LBillSaudas = Trim(Str(TRec!SAUDAID))
'        Else
'            If LStr_Exists(LBillSaudas, TRec!SAUDAID) = False Then LBillSaudas = LBillSaudas & "," & Trim(Str(TRec!SAUDAID)) & ""
'        End If
'        LSaudaID = TRec!SAUDAID
'        'LSaudaId = Get_SaudaID(TRec!Sauda)
'
'        Call Get_Value:
''        mysql = "SELECT CONFIRM FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(CDdate.Value, "YYYY/MM/DD") & "'"
''        mysql = mysql & " AND CONNO=" & TRec!CONNO & " AND PARTY='" & TRec!CONCODE & "'"
''        Set TRec1 = Nothing
''        Set TRec1 = New ADODB.Recordset
''        TRec1.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
''        If Not TRec1.EOF Then
''            TxtBrokerConfirm.text = TRec1!CONFIRM
''        Else
''            TxtBrokerConfirm.text = "0"
''        End If
''        Set TRec = Nothing
''        Set TRec1 = Nothing
'    End If
    
End Sub
Private Sub FillFSaudaCombo()
    Set LFSaudaRec = Nothing
    Set LFSaudaRec = New ADODB.Recordset
    mysql = "EXEC Get_SaudaCtr_d " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "'," & LFExID & ",'" & LFParty & "'"
    LFSaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LFSaudaRec.EOF Then
        DComboSauda.Enabled = True:                 Set DComboSauda.RowSource = LFSaudaRec
        DComboSauda.BoundColumn = "SAUDACODE":      DComboSauda.ListField = "SAUDANAME"
    Else
        DComboSauda.Enabled = False
    End If
'    ListView1.Visible = False
'    ListView1.ListItems.Clear
'    Do While Not LFSaudaRec.EOF
'        ListView1.ListItems.Add , , LFSaudaRec!saudacode
'        LFSaudaRec.MoveNext
'    Loop
'    ListView1.Visible = True
    
End Sub
Private Sub FillFBrokerCombo()
    
    If LenB(DComboSauda.BoundText) <> 0 Then
        LFSauda = DComboSauda.BoundText
    Else
        LFSauda = vbNullString
    End If
    Set LFBrokerRec = Nothing
    Set LFBrokerRec = New ADODB.Recordset
    
    Call DATA_GRID_REFRESH
End Sub

Private Sub FillLFPartyCombo()
    Set LFPartyRec = Nothing
    
    Set LFPartyRec = New ADODB.Recordset
    mysql = " EXEC Get_PartyCtr_d " & GCompCode & ",'" & Format(CDdate.Value, "YYYY/MM/DD") & "'," & LFExID & ""
    LFPartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not LFPartyRec.EOF Then
        DComboParty.Enabled = True
        Set DComboParty.RowSource = LFPartyRec
        DComboParty.BoundColumn = "AC_CODE"
        DComboParty.ListField = "NAME"
        DComboCode.Enabled = True
        Set DComboCode.RowSource = LFPartyRec
        DComboCode.BoundColumn = "AC_CODE"
        DComboCode.ListField = "AC_CODE"
    Else
        DComboParty.Enabled = False
        DComboCode.Enabled = False
    End If
'    ListView2.Visible = False
'    ListView2.ListItems.Clear
'    Do While Not LFPartyRec.EOF
'        ListView2.ListItems.Add , , LFPartyRec!AC_CODE
'        ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , LFPartyRec!NAME
'        LFPartyRec.MoveNext
'    Loop
'    ListView2.Visible = True
End Sub



