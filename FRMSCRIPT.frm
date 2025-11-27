VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmscript 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   12090
   Begin VB.Frame Frame5 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1695
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H80000013&
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
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1815
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4920
      TabIndex        =   10
      Top             =   240
      Width           =   5415
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Commodity && Script Master "
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
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00004080&
      Height          =   7695
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   11895
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   5880
         TabIndex        =   7
         Top             =   1200
         Width           =   5775
         Begin MSComctlLib.ListView ListView2 
            Height          =   5415
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select item specific."
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   9551
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EXCode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Commodity Name"
               Object.Width           =   6722
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "Commodities && Scripts Master"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   5415
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   5775
         Begin MSComctlLib.ListView ListView1 
            Height          =   5415
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select item specific."
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   9551
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EXCode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Commodity Name"
               Object.Width           =   6722
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "iTEMCODE"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "Commodities && Scripts Master"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   5415
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11535
         Begin VB.CommandButton Command1 
            Caption         =   "Ok"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8880
            TabIndex        =   1
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10200
            TabIndex        =   2
            Top             =   120
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   420
            Left            =   1320
            TabIndex        =   3
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   741
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404000&
            BackStyle       =   0  'Transparent
            Caption         =   "Exchange"
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
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmscript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DataCombo1.Text = "" Then
    MsgBox "PLease Select Exchange"
Else
    Set Rec = Nothing
    Set Rec = New ADODB.Recordset
    MYSQL = "SELECT COMPCODE FROM EXMAST WHERE COMPCODE =" & MC_CODE & " AND EXCODE ='" & DataCombo1.BoundText & "'"
    Rec.Open MYSQL, cnn, adOpenStatic, adLockReadOnly
    If Not Rec.EOF Then
        FILL_DATA
    End If
End If

End Sub

Private Sub DataCombo1_GotFocus()
SendKeys "%{DOWN}"
End Sub

Private Sub Form_Load()
Frame3.Enabled = False
Frame4.Enabled = False
Set Rec = Nothing
Set Rec = New ADODB.Recordset
MYSQL = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE =" & MC_CODE & " ORDER BY EXNAME "
Rec.Open MYSQL, cnn, adOpenStatic, adLockReadOnly
If Not Rec.EOF Then Set DataCombo1.RowSource = Rec: DataCombo1.ListField = "EXNAME": DataCombo1.BoundColumn = "EXCODE"
End Sub


Private Sub FILL_DATA()
Frame3.Enabled = True
Frame4.Enabled = True
Call UPDATE_ACTIVE
Dim RECSAUDA As ADODB.Recordset
    Set RECSAUDA = Nothing
    Set RECSAUDA = New ADODB.Recordset
    MYSQL = "SELECT EXCODE,ITEMCODE,ACTIVE FROM SCRIPTMASTER WHERE EXCODE ='" & DataCombo1.BoundText & "' ORDER BY ITEMNAME"
    RECSAUDA.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
       If Not RECSAUDA.EOF Then
        Do While Not RECSAUDA.EOF
            ListView1.ListItems.ADD , , RECSAUDA!ExCode
            ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.ADD , , RECSAUDA!itemcode
            If RECSAUDA!Active = "Y" Then
                ListView1.ListItems(ListView1.ListItems.Count).Checked = True
            End If
            RECSAUDA.MoveNext
        Loop
    End If
    Set RECSAUDA = Nothing
    Set RECSAUDA = New ADODB.Recordset
    MYSQL = "SELECT EXCHANGECODE,ITEMCODE FROM ITEMMAST WHERE COMPCODE =" & MC_CODE & " AND EXCHANGECODE ='" & DataCombo1.BoundText & "' ORDER BY ITEMNAME"
    RECSAUDA.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
       If Not RECSAUDA.EOF Then
        Do While Not RECSAUDA.EOF
            ListView2.ListItems.ADD , , RECSAUDA!EXCHANGECODE
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.ADD , , RECSAUDA!itemcode
            RECSAUDA.MoveNext
        Loop
    End If
    
End Sub

