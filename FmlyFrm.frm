VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FmlyFrm 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
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
   ScaleHeight     =   10935
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
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
      Height          =   7335
      Left            =   8880
      TabIndex        =   33
      Top             =   1560
      Width           =   7215
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7215
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   12726
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   27
      Top             =   840
      Width           =   7095
      Begin VB.CommandButton Command3 
         Caption         =   "Filter"
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox TxtFFmlyNAME 
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   31
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox TxtFFmlyCode 
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   29
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   195
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Branch List"
      Height          =   495
      Left            =   16080
      TabIndex        =   25
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Branch Wise Sub Brokerage and Sharing"
      Height          =   975
      Left            =   16080
      TabIndex        =   24
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -240
      TabIndex        =   21
      Top             =   0
      Width           =   16215
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1455
         Left            =   240
         TabIndex        =   22
         Top             =   0
         Width           =   17415
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Branch Master Setup"
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
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   15855
         End
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   1095
      Left            =   16320
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
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
      BackColor       =   &H00FFFFFF&
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
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   8895
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   285
            Left            =   2880
            TabIndex        =   43
            Top             =   0
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   5775
            Left            =   0
            TabIndex        =   6
            Top             =   480
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   10186
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   2716
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ACCID"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSFRPT 
            Height          =   3615
            Left            =   0
            TabIndex        =   40
            Top             =   2280
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   6376
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   30
            BackColor       =   16777215
            ForeColor       =   0
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
         Begin MSComctlLib.ListView ListView1 
            Height          =   5775
            Left            =   4440
            TabIndex        =   19
            Top             =   480
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   10186
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
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   2716
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ACCID"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "Selected Account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   26
            Top             =   0
            Width           =   4215
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select Account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   4335
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
         Height          =   1455
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   8655
         Begin VB.TextBox TELEID 
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   840
            MaxLength       =   50
            TabIndex        =   41
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Update Members"
            Height          =   285
            Left            =   6480
            TabIndex        =   39
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox UCCTEXT 
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   4560
            MaxLength       =   6
            TabIndex        =   37
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox TxtFmlyID 
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   840
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   120
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Post Settlement Entries "
            Height          =   285
            Left            =   3240
            TabIndex        =   5
            Top             =   1200
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox FmlyNameTxt 
            Height          =   405
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   2
            Top             =   120
            Width           =   4455
         End
         Begin VB.TextBox FmlyCodeTxt 
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   1
            Top             =   120
            Width           =   975
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   360
            Left            =   840
            TabIndex        =   3
            Top             =   600
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   16711680
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
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   360
            Left            =   5160
            TabIndex        =   4
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   16711680
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TeleId"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UCC"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3960
            TabIndex        =   38
            Top             =   1080
            Width           =   405
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1560
            TabIndex        =   35
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contra A/c"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3960
            TabIndex        =   17
            Top             =   675
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3240
            TabIndex        =   16
            Top             =   195
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   15
            Top             =   195
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Head"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   14
            Top             =   675
            Width           =   495
         End
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   8145
      Left            =   8520
      Top             =   720
      Width           =   7365
   End
   Begin VB.Label LblNew 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&New (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   12
      Top             =   -600
      Width           =   930
   End
   Begin VB.Label LblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit (F2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   11
      Top             =   -600
      Width           =   885
   End
   Begin VB.Label LblDelete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete (F3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7800
      TabIndex        =   10
      Top             =   -240
      Width           =   1170
   End
   Begin VB.Label LblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6720
      TabIndex        =   9
      Top             =   -480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label LblSave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save (F4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   8
      Top             =   -240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   7905
      Left            =   0
      Top             =   840
      Width           =   9045
   End
End
Attribute VB_Name = "FmlyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TRec As ADODB.Recordset:            Dim RecFmlyID1 As ADODB.Recordset
Dim RecFmlyID2 As ADODB.Recordset:      Dim AcFmlyRec As ADODB.Recordset
Dim AccRec As ADODB.Recordset:          Public Fb_Press As Byte
Dim LRpt As Integer:                    Dim old_Code As String
Dim LFmlyID As Integer: Dim Vgridrow As Integer:                Dim old_Family As String
Dim UCCTXT As String
Dim MTELEID As String
Dim AccRec2 As ADODB.Recordset

Sub Add_Rec()
    Fb_Press = 1: old_Family = vbNullString: old_Code = vbNullString
    Call Get_Selection(1)
    FmlyCodeTxt.text = vbNullString
    FmlyNameTxt.text = vbNullString
'    DataList1.Locked = True
    Frame1.Enabled = True: FmlyCodeTxt.SetFocus
End Sub
Sub Save_Rec()
    Dim TRec As ADODB.Recordset:        Dim LFmlyCode As String
    Dim LPostSettle As String:          Dim LHeadID As Long
    Dim LContraID As Long:              Dim LHeadCode As String
    Dim LContraCode As String:          Dim LFmlyID As Long
    Dim AccRec As ADODB.Recordset
    
    On Error GoTo err1
    If LenB(FmlyCodeTxt.text) < 1 Then MsgBox "Family Code required before saving record.", vbCritical, "Error": FmlyCodeTxt.SetFocus: Exit Sub
    If LenB(FmlyNameTxt.text) < 1 Then MsgBox "Family Name required before saving record.", vbCritical, "Error": FmlyNameTxt.SetFocus: Exit Sub
    If LenB(DataCombo1.BoundText) < 1 Then MsgBox "Family head name required before saving record.", vbCritical, "Error": DataCombo1.SetFocus: Exit Sub
    If LenB(DataCombo2.BoundText) < 1 Then MsgBox "Contra Account Name required before saving record.", vbCritical, "Error": DataCombo2.SetFocus: Exit Sub
    CNNERR = True
    Cnn.BeginTrans
    If Check2.Value = 1 Then
        LPostSettle = "Y"
    Else
        LPostSettle = "N"
    End If
    LHeadCode = vbNullString
    LHeadCode = Get_AccountMCode(DataCombo1.BoundText)
    If LenB(LHeadCode) < 1 Then
        MsgBox " PLEASE SELECT VALID HEAD ACCOUNT"
        DataCombo1.SetFocus
        Exit Sub
    End If
    LContraCode = vbNullString
    LContraCode = Get_AccountMCode(DataCombo2.BoundText)
    If LenB(LContraCode) < 1 Then
        MsgBox " PLEASE SELECT VALID CONTRA  ACCOUNT"
        DataCombo2.SetFocus
        Exit Sub
    End If
    LHeadID = Get_AccID(LHeadCode)
    LContraID = Get_AccID(LContraCode)
    LFmlyID = Val(TxtFmlyID.text)
    If Fb_Press = 1 Then
    
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT FMLYID,FMLYCODE,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " AND FMLYCODE ='" & FmlyCodeTxt.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then MsgBox "Duplicate family code. Already exists with  family name " & TRec!FmlyNAME, vbExclamation, "Warning": FmlyCodeTxt.SetFocus: Exit Sub
            
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT FmlyName FROM AccFmly WHERE COMPCODE =" & GCompCode & " AND FmlyName='" & FmlyNameTxt.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then MsgBox "Duplicate family name.", vbExclamation, "Warning": FmlyNameTxt.SetFocus: Exit Sub
        mysql = "INSERT INTO ACCFMLY( COMPCODE,FMLYCODE,FMLYNAME,FMLYHEAD,CONTRA_AC,POSTSETTLE,HEADID,CONTRAID )"
        mysql = mysql & " VALUES (" & GCompCode & ",'" & FmlyCodeTxt.text & "','" & FmlyNameTxt.text & "','" & LHeadCode & "','" & LContraCode & "','" & LPostSettle & "'," & LHeadID & "," & LContraID & " )"
        Cnn.Execute mysql
    Else
        If UCase(old_Code) = UCase(FmlyCodeTxt.text) Then
        Else
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            TRec.Open "SELECT FmlyCode FROM AccFmly WHERE COMPCODE=" & GCompCode & " AND FmlyCode='" & FmlyCodeTxt.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then MsgBox "Family code already exists.", vbExclamation, "Warning": FmlyCodeTxt.SetFocus: Exit Sub
        End If
        If UCase(old_Family) = UCase(FmlyNameTxt.text) Then
        Else
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            TRec.Open "SELECT FmlyName FROM AccFmly WHERE COMPCODE=" & GCompCode & " AND FmlyName='" & FmlyNameTxt.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then MsgBox "Family name already exists.", vbExclamation, "Warning": FmlyNameTxt.SetFocus: Exit Sub
        End If
    End If
    If Fb_Press <> 1 Then
        mysql = "UPDATE ACCFMLY SET "
        mysql = mysql & " FMLYNAME ='" & FmlyNameTxt.text & "'"
        mysql = mysql & " ,FMLYHEAD ='" & LHeadCode & "'"
        mysql = mysql & " ,CONTRA_AC ='" & LContraCode & "'"
        mysql = mysql & " ,HEADID =" & LHeadID & ""
        mysql = mysql & " ,CONTRAID =" & LContraID & ""
        mysql = mysql & " ,POSTSETTLE   ='" & LPostSettle & "'"
        mysql = mysql & " WHERE COMPCODE =" & GCompCode & "  AND FMLYCODE ='" & FmlyCodeTxt.text & "'"
        Cnn.Execute mysql
    End If
    LFmlyCode = FmlyCodeTxt.text
    mysql = "DELETE FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND FMLYCODE ='" & FmlyCodeTxt.text & "'"
    Cnn.Execute mysql
    
    
    Dim LREC As ADODB.Recordset
    Dim lrec2 As ADODB.Recordset
    Dim LAC_CODE As String
    Dim LACCID  As Long
    Dim I As Integer
    Dim LParty As String
    UCCTXT = UCCTEXT.text
    MTELEID = TELEID.text

    LFmlyID = Get_Fmlyid(LFmlyCode)
    For I = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(I).Checked = True Then
            LParty = ListView2.ListItems(I).ListSubItems(1)
            LACCID = ListView2.ListItems(I).ListSubItems(2)
            
            mysql = "INSERT INTO ACCFMLYD (COMPCODE,FMLYCODE,PARTY,FMLYID,ACCID) VALUES "
            mysql = mysql & "(" & GCompCode & ",'" & LFmlyCode & "','" & LParty & "'," & LFmlyID & "," & LACCID & ")"
            Cnn.Execute mysql
            mysql = "SELECT EXID,EXCODE,OPTIONS FROM EXMAST  WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE "
            Set lrec2 = Nothing
            Set lrec2 = New ADODB.Recordset
            lrec2.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            If Not lrec2.EOF Then
                Do While Not lrec2.EOF
                    LAC_CODE = Get_PEXSBROK_AC_CODE(LFmlyID, LACCID, lrec2!EXID, GFinEnd, "FUT")
                    If LenB(LAC_CODE) < 1 Then
                        If lrec2!excode = "BEQ" Or lrec2!excode = "EQ" Then
                            LAC_CODE = Get_PEXSBROK_AC_CODE(LFmlyID, LACCID, lrec2!EXID, GFinEnd, "CSH")
                            If LenB(LAC_CODE) < 1 Then Call PINSERT_PEXSBROK(LParty, LFmlyCode, lrec2!excode, "P", 0, "N", 0, GFinEnd, "CSH", 0, lrec2!EXID, LFmlyID, LACCID)
                        Else
                            Call PINSERT_PEXSBROK(LParty, LFmlyCode, lrec2!excode, "P", 0, "N", 0, GFinEnd, "FUT", 0, lrec2!EXID, LFmlyID, LACCID)
                            If lrec2!Options = "Y" Then
                                If lrec2!excode = "NSE" Or lrec2!excode = "MCX" Or lrec2!excode = "NCDX" Then
                                    LAC_CODE = Get_PEXSBROK_AC_CODE(LFmlyID, LACCID, lrec2!EXID, GFinEnd, "OPT")
                                    If LenB(LAC_CODE) < 1 Then Call PINSERT_PEXSBROK(LParty, LFmlyCode, lrec2!excode, "P", 0, "N", 0, GFinEnd, "OPT", 0, lrec2!EXID, LFmlyID, LACCID)
                                End If
                            End If
                        End If
                    End If
                    If Check1.Value = 1 Then
                        mysql = "UPDATE ACCOUNTD SET UCC =  '" & UCCTXT & "' ,DIRECTOR =  '" & MTELEID & "'  WHERE COMPCODE = " & GCompCode & " AND AC_CODE = '" & LAC_CODE & "'"
                        Cnn.Execute mysql
                    End If
                    lrec2.MoveNext
                Loop
            End If
        Else
            LACCID = ListView2.ListItems(I).ListSubItems(2)
            mysql = "DELETE FROM PEXSBROK WHERE FMLYID = '" & LFmlyID & "' AND ACCID = '" & LACCID & "' AND BROKRATE=0 AND SHRATE=0"
            Cnn.Execute mysql
        End If
    Next
    Call CANCEL_REC
    GETMAIN.bwtbal.Visible = True
    Cnn.CommitTrans
    CNNERR = False
    Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    Screen.MousePointer = 0:
    
    Command1.Enabled = True
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End Sub
Sub CANCEL_REC()
    Set AcFmlyRec = Nothing: Set AcFmlyRec = New ADODB.Recordset
    mysql = "SELECT * FROM AccFmly WHERE COMPCODE=" & GCompCode & " ORDER BY FmlyNAME"
    AcFmlyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    Fb_Press = 0
    Call Get_Selection(10)
    FmlyCodeTxt.text = vbNullString
    FmlyNameTxt.text = vbNullString
    DataCombo1.BoundText = vbNullString
    DataCombo2.BoundText = vbNullString
    TxtFFmlyCode.text = vbNullString
    TxtFFmlyNAME.text = vbNullString
    Check2.Value = 0
    'DataList1.Locked = False
    Call Fill_DataGrid
    Frame1.Enabled = False
    ListView1.ListItems.Clear
    Dim I As Integer
    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = False
    Next
End Sub
Sub MODIFY_REC()
    Dim I As Integer
    Dim TRec As ADODB.Recordset
    Dim ACREC As ADODB.Recordset
    
    If Trim(FmlyCodeTxt.text) <> "" Then
        'DataList1.Locked = True
        AcFmlyRec.MoveFirst
        AcFmlyRec.Find "FMLYCODE='" & FmlyCodeTxt.text & "'", , adSearchForward
        If Not AcFmlyRec.EOF Then
            LFmlyID = Val(AcFmlyRec!FMLYID & vbNullString)
            TxtFmlyID.text = AcFmlyRec!FMLYID & vbNullString
            FmlyCodeTxt.text = AcFmlyRec!FMLYCODE
            FmlyNameTxt.text = AcFmlyRec!FmlyNAME
            DataCombo1.BoundText = AcFmlyRec!FMLYHEAD & vbNullString
            AccRec2.MoveFirst
            AccRec2.Find "AC_CODE ='" & (AcFmlyRec!Contra_Ac & vbNullString) & "'"
            If Not AccRec2.EOF Then
                DataCombo2.BoundText = AccRec2!AC_CODE
            End If
            old_Code = FmlyCodeTxt.text: old_Family = FmlyNameTxt.text
            Frame1.Enabled = True: FmlyCodeTxt.SetFocus
            mysql = "SELECT B.ACCID,B.AC_CODE ,B.NAME FROM ACCFMLYD A, ACCOUNTM B WHERE A.COMPCODE =" & GCompCode & " "
            mysql = mysql & " AND A.ACCID =B.ACCID AND  A.FMLYCODE ='" & FmlyCodeTxt.text & "' ORDER BY B.NAME"
            Set RecFmlyID1 = Nothing
            Set RecFmlyID1 = New ADODB.Recordset
            RecFmlyID1.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            ListView1.Visible = False
            If Not RecFmlyID1.EOF Then
                ListView1.ListItems.Clear
                Do While Not RecFmlyID1.EOF
                    ListView1.ListItems.Add , , RecFmlyID1!NAME
                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , RecFmlyID1!AC_CODE
                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , RecFmlyID1!ACCID
                    ListView1.ListItems(ListView1.ListItems.Count).Checked = True
                    RecFmlyID1.MoveNext
                Loop
            End If
            ListView1.Visible = True
            If RecFmlyID1.RecordCount > 0 Then RecFmlyID1.MoveFirst
            ListView2.Visible = False
            For I = 1 To ListView2.ListItems.Count
                ListView2.ListItems(I).Checked = False
            Next
            If Not RecFmlyID1.EOF Then
                Do While Not RecFmlyID1.EOF
                    For I = 1 To ListView2.ListItems.Count
                        If ListView2.ListItems(I).ListSubItems(1).text = RecFmlyID1!AC_CODE Then
                            ListView2.ListItems(I).Checked = True
                        End If
                    Next
                    RecFmlyID1.MoveNext
                Loop
            End If
            ListView2.Visible = True
            
            'Getucc
            Set ACREC = Nothing: Set ACREC = New ADODB.Recordset
            mysql = "SELECT UCC,DIRECTOR FROM ACCOUNTD WHERE COMPCODE= " & GCompCode & " AND AC_CODE = '" & DataCombo1.BoundText & "' "
            ACREC.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            If Not ACREC.EOF Then
                UCCTEXT.text = IIf(IsNull(ACREC!UCC), "", ACREC!UCC)
                TELEID.text = ACREC!DIRECTOR
            End If
        End If
        If Fb_Press = 3 Then
            If MsgBox("You are about to Delete one record. Confirm Delete ?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                Set TRec = Nothing: Set TRec = New ADODB.Recordset
                mysql = "SELECT * FROM PITSBROK  WHERE COMPCODE=" & GCompCode & " AND FmlyCODE ='" & FmlyCodeTxt.text & "' AND ( BROKRATE <> 0 OR  SHRATE<>0 )"
                TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then
                    MsgBox "Transaction Exists  ", vbExclamation, "Error"
                    Call CANCEL_REC
                    Exit Sub
                Else
                    Set TRec = Nothing: Set TRec = New ADODB.Recordset
                    mysql = "SELECT * FROM PEXSBROK  WHERE COMPCODE=" & GCompCode & " AND FmlyCODE ='" & FmlyCodeTxt.text & "' AND ( BROKRATE <> 0 OR  SHRATE<>0 )"
                    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                    If Not TRec.EOF Then
                        MsgBox "Transaction Exists  ", vbExclamation, "Error"
                        Call CANCEL_REC
                        Exit Sub
                    Else
                        mysql = "DELETE FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " AND FMLYCODE ='" & FmlyCodeTxt.text & "'"
                        Cnn.Execute mysql
                    
                        mysql = "DELETE FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND FMLYCODE ='" & FmlyCodeTxt.text & "'"
                        Cnn.Execute mysql
                    
                        mysql = "DELETE FROM PITSBROK WHERE COMPCODE =" & GCompCode & " AND FMLYCODE ='" & FmlyCodeTxt.text & "'"
                        Cnn.Execute mysql
                        
                        mysql = "DELETE FROM PEXSBROK WHERE COMPCODE =" & GCompCode & " AND FMLYCODE ='" & FmlyCodeTxt.text & "'"
                        Cnn.Execute mysql
                        
                        mysql = "SELECT FmlyCode FROM AccFmly WHERE COMPCODE=" & GCompCode & ""
                        Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                        If Not TRec.EOF Then
                            GETMAIN.bwtbal.Visible = True
                        Else
                            GETMAIN.bwtbal.Visible = False
                        End If
                    End If
                End If
            End If
            Call CANCEL_REC
            'DataList1.Locked = False
            'DataList1.SetFocus
        End If
    Else
        MsgBox "Please SELECT family.", vbCritical
        Call CANCEL_REC
        'DataList1.Locked = False
        'DataList1.SetFocus
    End If
End Sub


Private Sub Check3_Click()
Dim I As Integer
    For I = 1 To ListView2.ListItems.Count
        If Check3.Value = 1 Then
            ListView2.ListItems.Item(I).Checked = True
        Else
            ListView2.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub

Private Sub Command1_Click()
    LRpt = 1
    Call FmlyList_Rec
End Sub

Private Sub Command2_Click()
LRpt = 2
 Call FmlyList_Rec
End Sub


Private Sub Command3_Click()
Call Fill_DataGrid
End Sub

Private Sub DataCombo1_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub DataCombo2_GotFocus()
    Sendkeys "%{down}"
End Sub

'Private Sub DataList1_Click()
'    FmlyCodeTxt.text = DataList1.BoundText
'    FmlyNameTxt.text = DataList1.text
'End Sub
'P 'rivate Sub DataList1_DblClick()
' '   If DataList1.Locked Then
'    Else
'        Call Get_Selection(2)
'        Fb_Press = 2
'        Call MODIFY_REC
 '   End If
'End Sub
'Private Sub DataList1_KeyPress(KeyAscii As Integer)
'    If DataList1.Locked Then
'    Else
'        If KeyAscii = 13 Then
'            Call Get_Selection(2)
'            Fb_Press = 2
 '           Call MODIFY_REC
 '       End If
''    End If
'End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Form_Load()
    Call CANCEL_REC
    Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
    mysql = "SELECT ACCID,AC_CODE, NAME FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " ORDER BY NAME"
    AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRec.EOF Then
        Set DataCombo1.RowSource = AccRec
        DataCombo1.ListField = "name"
        DataCombo1.BoundColumn = "ac_code"
    
    End If
    Set AccRec2 = Nothing: Set AccRec2 = New ADODB.Recordset
    mysql = "SELECT ACCID,AC_CODE, NAME FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & " ORDER BY NAME"
    AccRec2.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    
    If Not AccRec2.EOF Then
        Set DataCombo2.RowSource = AccRec2
        DataCombo2.ListField = "NAME"
        DataCombo2.BoundColumn = "AC_CODE"
    End If
    Call Fill_DataGrid
    ListView2.ListItems.Clear
    Do While Not AccRec.EOF
        ListView2.ListItems.Add , , AccRec!NAME
        ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , AccRec!AC_CODE
        ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , AccRec!ACCID
        AccRec.MoveNext
    Loop

End Sub
Private Sub Form_Paint()
'    Me.BackColor = GETMAIN.BackColor
End Sub
Private Sub Form_Unload(Cancel As Integer)
If CRViewer1.Visible = True Then
        Call Get_Selection(10)
        CRViewer1.Visible = False
        Frame4.Visible = True
        Frame6.Visible = True
        Cancel = 1
    Else
        Call CANCEL_REC
        GETMAIN.StatusBar1.Panels(1).text = vbNullString
        Unload Me
    End If
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim LREC As ADODB.Recordset
Set LREC = Nothing
Set LREC = New ADODB.Recordset
    If Item.Checked = False Then
        
        mysql = "SELECT COMPCODE FROM PITSBROK WHERE COMPCODE =" & GCompCode & "  AND FMLYCODE ='" & FmlyCodeTxt.text & "' AND PARTY='" & Item.ListSubItems(1) & "' AND (BROKRATE <>0 OR SHRATE<>0 )"
        LREC.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not LREC.EOF Then
            MsgBox "Can Not Remove This Party as Sub Brokerage or Sharing Already Set. Please Set Zero in Sub Brokerage and Sharing to Remove This Party From This Branch"
            Item.Checked = True
        Else
            mysql = "SELECT COMPCODE FROM PEXSBROK WHERE COMPCODE =" & GCompCode & "  AND FMLYCODE ='" & FmlyCodeTxt.text & "' AND PARTY='" & Item.ListSubItems(1) & "' AND (BROKRATE <>0 OR SHRATE<>0 )"
            Set LREC = Nothing:        Set LREC = New ADODB.Recordset
            LREC.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            If Not LREC.EOF Then
                MsgBox "Can Not Remove This Party as Sub Brokerage or Sharing Already Set. Please Set Zero in Sub Brokerage and Sharing to Remove This Party From This Branch"
                Item.Checked = True
            End If
        End If
    End If
End Sub
Sub FmlyList_Rec()
    Screen.MousePointer = 11
    Call Get_Selection(12)
    If LRpt = 1 Then
        mysql = " SELECT A.FMLYCODE,A.FMLYNAME,A.PARTY,B.NAME FROM ACCFMLYD AS A, ACCOUNTM AS B "
        mysql = mysql & " WHERE A.COMPCODE = B.COMPCODE  AND A.PARTY=B.AC_CODE AND  A.COMPCODE=" & GCompCode & " ORDER BY A.FMLYNAME, B.NAME "
    Else
        mysql = " SELECT A.FMLYCODE,A.FMLYNAME,A.PARTY,B.NAME,P.ITEMCODE,P.BROKTYPE,P.BROKRATE,PS.BROKTYPE AS BBROKTYPE,PS.BROKRATE  AS BBROKRATE,PS.SHTYPE,PS.SHRATE,P.UPTOSTDT FROM ACCFMLYD AS A, ACCOUNTM AS B,PITBROK AS P, PITSBROK AS PS "
        mysql = mysql & " WHERE A.COMPCODE = B.COMPCODE And A.PARTY = B.AC_CODE And A.COMPCODE = " & GCompCode & ""
        mysql = mysql & " AND P.COMPCODE=A.COMPCODE AND P.AC_CODE=B.AC_CODE AND P.AC_CODE = PS.PARTY AND P.UPTOSTDT = PS.UPTOSTDT AND PS.COMPCODE =A.COMPCODE AND PS.PARTY = A.PARTY AND PS.FMLYCODE = A.FMLYCODE AND P.ITEMCODE = PS.ITEMCODE "
        mysql = mysql & " ORDER BY A.FMLYNAME, B.NAME "
    End If
    Dim TRec As ADODB.Recordset
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If LRpt = 1 = True Then
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "Fmly_List.RPT", 1)
    Else
        Set RDCREPO = RDCAPP.OpenReport(GReportPath & "Fmly_BrokList.RPT", 1)
    End If
    Frame4.Visible = False: Frame6.Visible = False
    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource TRec
    CRViewer1.Width = CInt(GETMAIN.Width - 100)
    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Top = 0: CRViewer1.Left = 0
    CRViewer1.Visible = True
    CRViewer1.ReportSource = RDCREPO
    CRViewer1.ViewReport
    'Set RPT = Nothing
    Screen.MousePointer = 0
End Sub
Private Sub Text20_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub DataGrid1_Click()
If AcFmlyRec.RecordCount > 0 Then
    If AcFmlyRec.EOF Then AcFmlyRec.MoveFirst
    DataGrid1.Col = 1
    FmlyCodeTxt.text = DataGrid1.text
    DataGrid1.Col = 2
    FmlyNameTxt = DataGrid1.text
End If
End Sub
Private Sub DataGrid1_DblClick()
    Call Get_Selection(2)
    Fb_Press = 2
    Call MODIFY_REC
End Sub
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    Dim LChar As String
    LChar = UCase(Chr(KeyAscii))
    If KeyAscii = 13 Then
        Call Get_Selection(2)
        Fb_Press = 2
        Call MODIFY_REC
    Else
        AcFmlyRec.MoveFirst
        Do While Not AcFmlyRec.EOF
            If Left$(AcFmlyRec!FmlyNAME, 1) <> LChar Then
                AcFmlyRec.MoveNext
            Else
                Exit Do
            End If
        Loop
        If AcFmlyRec.EOF Then AcFmlyRec.MoveFirst
    End If
End Sub

Private Sub Fill_DataGrid()
    Set AcFmlyRec = Nothing: Set AcFmlyRec = New ADODB.Recordset
    mysql = "SELECT FmlyID,FmlyCode,FmlyName ,FmlyHead,Contra_Ac,Postsettle FROM AccFmly WHERE COMPCODE=" & GCompCode & " "
    If LenB(TxtFFmlyCode.text) > 0 Then mysql = mysql & "AND FMLYCODE LIKE '" & TxtFFmlyCode.text & "%' "
    If LenB(TxtFFmlyNAME.text) > 0 Then mysql = mysql & "AND FMLYNAME LIKE '" & TxtFFmlyNAME.text & "%' "
    mysql = mysql & " ORDER BY FmlyNAME"
    AcFmlyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AcFmlyRec.EOF Then
        Set DataGrid1.DataSource = AcFmlyRec
        DataGrid1.ReBind
        DataGrid1.Refresh
        DataGrid1.Columns(0).Width = 500
        DataGrid1.Columns(1).Width = 800
        DataGrid1.Columns(2).Width = 3500
        DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Width = 1000
    End If
End Sub

Private Sub MSF_GRID()
    MSFRPT.Row = 0
    MSFRPT.Col = 0: MSFRPT.ColWidth(0) = TextWidth("XXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 0) = "Select"
    MSFRPT.Col = 1: MSFRPT.ColWidth(1) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 1) = "Name"
    MSFRPT.Col = 2: MSFRPT.ColWidth(2) = TextWidth("XXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 2) = "Include MTM"
    MSFRPT.Col = 3: MSFRPT.ColWidth(3) = TextWidth("XXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 3) = "Code"
    MSFRPT.Col = 4: MSFRPT.ColWidth(4) = TextWidth("XXXXXXX"): MSFRPT.CellFontBold = True: MSFRPT.TextMatrix(0, 4) = "ACCID"
End Sub

