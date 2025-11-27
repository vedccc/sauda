VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Splash 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7770
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   2895
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Visible         =   0   'False
      Width           =   7215
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5280
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Select"
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
         Left            =   6000
         TabIndex        =   24
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Restore Database Backup"
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
         Left            =   1320
         TabIndex        =   23
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Backup file to be Restore"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Please wait ..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   2655
         Left            =   120
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   6735
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   120
         Width           =   975
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   7215
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "This Software is mainly for Traing Purpose and will stop working after 15 Days"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   6975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "This is a Demo Software.  All the Accounts, Trades and Ledger Positings are ficticious."
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   6975
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Caption         =   $"Splash.frx":000C
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   6735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   6720
         Begin VB.CommandButton Command2 
            Caption         =   "Log In"
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
            Left            =   5640
            TabIndex        =   4
            Top             =   120
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Change Password"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3720
            TabIndex        =   3
            Top             =   180
            Width           =   1815
         End
         Begin VB.TextBox txtUsername 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   960
            MaxLength       =   12
            TabIndex        =   1
            Top             =   120
            Width           =   1005
         End
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   2640
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pass"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   12
            Top             =   165
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "User Id"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   165
            Width           =   690
         End
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "A Complete Solution"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   6735
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         X1              =   6840
         X2              =   6840
         Y1              =   1560
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         X1              =   360
         X2              =   6840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         X1              =   360
         X2              =   360
         Y1              =   360
         Y2              =   1560
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         X1              =   360
         X2              =   6840
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   6720
      End
      Begin VB.Line Line6 
         BorderWidth     =   4
         X1              =   120
         X2              =   7080
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line7 
         BorderWidth     =   4
         X1              =   7080
         X2              =   7080
         Y1              =   120
         Y2              =   2880
      End
      Begin VB.Line Line8 
         BorderWidth     =   4
         X1              =   120
         X2              =   7080
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line9 
         BorderWidth     =   3
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   2880
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "14 Aug 2010"
      Height          =   255
      Left            =   6360
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EncryptionFlag As Boolean:  Dim HardDiskNo As String:           Dim MCnn As ADODB.Connection
Dim LDEMO As String:            Dim FlagVerAdj As Boolean:          Dim FlagIntigrity As Boolean
Dim Flag_Pitbrok As Boolean:    Dim LCompRec As ADODB.Recordset:    Dim GNew_Sub_Brok_Updt As Boolean
Dim NEWSOFT0 As Boolean:        Dim Rec As ADODB.Recordset
Dim DBRestoreflag  As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Frame4.Visible = True
    Text1.SetFocus
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{tab}"
End Sub

Private Sub Command1_Click()
If Text1.text <> Text2.text Then
    MsgBox "Confirm Password Does not match"
    Text2.text = vbNullString
    Text2.SetFocus
Else
    If EncryptionFlag Then
        GUserName = Get_UserName(txtUsername.text, txtPassword.text)
    Else
        GUserName = Get_UserName(txtUsername.text, EncryptNEW(txtPassword.text, 13))
    End If
    If LenB(GUserName) > 0 Then
        MYSQL = "UPDATE USERMASTER SET PASSWD='" & EncryptNEW(Text1.text, 13) & "' WHERE USER_NAME ='" & txtUsername.text & "'"
        Cnn.Execute MYSQL
        Call Check_Login
    Else
        MsgBox "Invalid Login"
        Frame4.Visible = False
        txtUsername.text = vbNullString
        txtPassword.text = vbNullString
        txtUsername.SetFocus
    End If
End If
End Sub

Private Sub Command2_Click()


                        
    If txtUsername.text = "" Then
        MsgBox "Please enter username!!!"
        txtUsername.SetFocus
    ElseIf txtPassword.text = "" Then
        MsgBox "Please enter password!!!"
        txtPassword.SetFocus
    Else
        Call Set_SystemTable
        Call ColChanges
        Call Check_Login
    End If
End Sub

Private Sub Command3_Click()

On Error GoTo err1

    If Text3.text = "" Then
        MsgBox "Please provide database name to be restore."
        Text3.SetFocus
    ElseIf Text4.text = "" Then
        MsgBox "Please select backup file to be restore."
        Text4.SetFocus
    Else
        Dim CnnString As String
        Cnn.close
        DoEvents
        Command3.Enabled = False
        Label11.Visible = True
        DoEvents
        CnnString = ServerString
        Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
        MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
        MCnn.Open
    
        MYSQL = "EXECUTE RESTOREDB '" & Text3.text & "','" & Text4.text & "'"
        MCnn.Execute MYSQL
        
        DoEvents
        Label11.Visible = False
        DBRestoreflag = False
        Me.Height = 3300
        Frame6.Visible = False
        DoEvents
                
    End If
    
err1:
If err.Number <> 0 Then
    MsgBox err.Description
    DoEvents
    Label11.Visible = False
End If
End Sub

Private Sub Command4_Click()
    Dim VBakFile As Variant
    CommonDialog1.ShowOpen
    VBakFile = CommonDialog1.FileName
    Text4.text = VBakFile
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim A, B, C As String
    If KeyCode = 27 Then Unload Me
    If KeyCode = 114 Then FlagVerAdj = True
    If KeyCode = 115 Then FlagBrok = True
    If KeyCode = 116 Then FlagIntigrity = True
    If KeyCode = 121 Then FlagStored = True
    If KeyCode = 117 Then Flag_Pitbrok = True
    If KeyCode = 113 Then EncryptionFlag = True
    If KeyCode = 120 Then Call Registration
    If KeyCode = 122 Then GFlag_Fin = True
    
    If KeyCode = 113 Then 'F2 key to restore backup
        DBRestoreflag = True
        Me.Height = 7500
        Frame6.Visible = True
        Text3.SetFocus
    End If
End Sub

Private Sub Form_Load()
  On Error GoTo err1
  Dim FIELD_FOUND As Boolean
  Dim CnnString As String
  Dim Rec As ADODB.Recordset
  FIELD_FOUND = True
  DBRestoreflag = False
  If App.PrevInstance = True Then
    MsgBox "Sauda Application is already Open"
    Unload Me
    Exit Sub
  End If
    ServerString = MServer
    ServerString = Left$(MServer, (InStr(1, MServer, "database") + 8)) & "MASTER"
    CnnString = ServerString
    Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
    MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
    MCnn.Open
    FlagVerAdj = False
    FlagIntigrity = False
    GNew_Sub_Brok_Updt = False
    Flag_Pitbrok = False
    Exit Sub
err1:
    If Val(err.Number) = Val(-2147217900) Then  '       Z x'FIELD IS NOT IN THE TABLE
        FIELD_FOUND = False
        Resume Next
    Else
        MsgBox err.Description, vbCritical, "Error Number : " & err.Number
       'Resume
    End If
End Sub
Private Sub Form_Paint()
    txtUsername.SetFocus
    If LDEMO = "Y" Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
        If Not DBRestoreflag Then
            Me.Height = 3300
        End If
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{tab}"
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If LenB(Text1.text) = 0 Then
    MsgBox "Please Enter new Password.", vbCritical
    Cancel = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{tab}"
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If LenB(Text2.text) = 0 Then
        MsgBox "Please Enter Password.", vbCritical
        Cancel = True
    End If
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2.SetFocus
    End If
        
End Sub

Private Sub txtPassword_Validate(Cancel As Boolean)
    
    If LenB(txtPassword.text) < 1 Then
'        MsgBox "Please Enter Password.", vbCritical
'        Cancel = True
    Else
        Command2.SetFocus
    End If
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then 'F2 key to restore backup
        DBRestoreflag = True
        Me.Height = 7500
        Frame6.Visible = True
        Text3.SetFocus
    End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{tab}"
End Sub
Private Sub ColChanges()
    Cnn.Execute " IF EXISTS(SELECT OBJECT_ID FROM SYS.OBJECTS WHERE TYPE='P' AND NAME='COL_CHNAGES') DROP PROCEDURE COL_CHNAGES"
    MYSQL = "CREATE PROCEDURE COL_CHNAGES AS SET NOCOUNT ON BEGIN "
    MYSQL = MYSQL & " DELETE FROM USER_RIGHTS WHERE MENUNAME LIKE 'mnufilefor%'"
    
    MYSQL = MYSQL & " IF  col_length('COMPANY','BillingCycle') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD BillingCycle BIT NOT NULL DEFAULT 1 WITH VALUES"
    
    MYSQL = MYSQL & " IF  col_length('COMPANY','QTY_DECIMAL') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD QTY_DECIMAL BIT NOT NULL DEFAULT 0 WITH VALUES"
    
    MYSQL = MYSQL & " IF  col_length('COMPANY','ConNoType') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD ConNoType INT NOT NULL DEFAULT 0 WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','ONLYBROK') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD ONLYBROK INT NOT NULL DEFAULT  0 WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','STMDT') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD STMDT SMALLDATETIME  NOT NULL DEFAULT  '2015/06/01' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','SHARE') IS NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD SHARE VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','MARGIN') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD MARGIN  VARCHAR(1) NOT NULL DEFAULT 'N' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','EQ') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD EQ VARCHAR(1) NOT NULL DEFAULT 'N' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','GENQUERY') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD GENQUERY VARCHAR(1) NOT NULL DEFAULT '1' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','SHOWLOT') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD SHOWLOT VARCHAR(1) NOT NULL DEFAULT  'N' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','SEBIREGNO') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD SEBIREGNO VARCHAR(20) NOT NULL DEFAULT  '' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','GSTIN') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD GSTIN VARCHAR(20) NOT NULL DEFAULT  '' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','UNIQCLIENTID') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD UNIQCLIENTID VARCHAR(50) NOT NULL DEFAULT 'A' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','TRANFEES') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD TRANFEES VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','STAMPDUTY') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD STAMPDUTY VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','VALUEWISE') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD VALUEWISE VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','STANDING') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD STANDING VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','MINBROKYN') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD MINBROKYN VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','STT') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD STT VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','SUBBROK') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD SUBBROK VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','SRVTAX') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD SRVTAX VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('SAUDAMAST','OPTTYPE') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE SAUDAMAST ADD OPTTYPE VARCHAR(1) NULL DEFAULT '' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('SAUDAMAST','INSTTYPE') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE SAUDAMAST ADD INSTTYPE VARCHAR(3) NULL DEFAULT 'FUT' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('SAUDAMAST','STRIKEPRICE ') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE SAUDAMAST ADD STRIKEPRICE  FLOAT  NULL DEFAULT 0 WITH VALUES END"
    MYSQL = MYSQL & " IF  col_length('COMPANY','SHOWSTD') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD SHOWSTD VARCHAR(1) NOT NULL DEFAULT 'N' WITH VALUES"
    MYSQL = MYSQL & " IF  col_length('COMPANY','AMTDIVIDE') IS  NULL"
    MYSQL = MYSQL & " ALTER TABLE COMPANY ADD AMTDIVIDE VARCHAR(1) NOT NULL DEFAULT '0' WITH VALUES"
    Cnn.Execute MYSQL
    
    MYSQL = "EXEC COL_CHNAGES"
    Cnn.Execute MYSQL
    
    MYSQL = "IF  col_length('SYSCOMP','SYSLOCKDT') IS  NULL "
    MYSQL = MYSQL & "ALTER TABLE SYSCOMP ADD SYSLOCKDT  SmallDateTime NOT NULL DEFAULT '2000/01/01' WITH VALUES "
    MCnn.Execute MYSQL
    
    MYSQL = "UPDATE SYSCOMP SET SYSLOCKDT ='2000/01/01' WHERE SYSLOCKDT IS NULL"
    MCnn.Execute MYSQL
    
    Cnn.Execute "  IF EXISTS(SELECT OBJECT_ID FROM SYS.OBJECTS WHERE TYPE='P' AND NAME='GET_USERNAME') DROP PROCEDURE GET_USERNAME "
    MYSQL = "CREATE PROCEDURE GET_USERNAME @USERNAME VARCHAR(12), @PASSWD VARCHAR(12),@PUSERNAME VARCHAR(12) OUTPUT"
    MYSQL = MYSQL & " AS SET NOCOUNT ON BEGIN DECLARE @LPUSERNAME AS VARCHAR(12) SET @PUSERNAME =''"
    MYSQL = MYSQL & " SELECT @LPUSERNAME =USER_NAME FROM USERMASTER WHERE USER_NAME =@USERNAME AND PASSWD = @PASSWD "
    MYSQL = MYSQL & " IF @LPUSERNAME IS NOT NULL SET @PUSERNAME   =@LPUSERNAME  End"
    Cnn.Execute MYSQL
    
End Sub

Private Sub Set_SystemTable()
On Error GoTo err1
    Dim TABLE_NAME As String:    Dim MRec As ADODB.Recordset:    Dim MCompCode As Integer:    Dim LDate As Date
    Dim MLFinBeg As Date:        Dim MLFinEnd As Date:           Dim MRpt_Path As String:     Dim MDPath As String
    Dim MSysLockDt As Date
                
    TABLE_NAME = "SYSTNO"
    MYSQL = "SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[" & TABLE_NAME & "]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1"
    Set MRec = Nothing
    Set MRec = New ADODB.Recordset
    MRec.Open MYSQL, MCnn, adOpenForwardOnly, adLockReadOnly
    If MRec.EOF Then
        MYSQL = "CREATE TABLE SYSTNO (SNO INT IDENTITY(1,1),SERIALNO VARCHAR(50) NOT NULL,CUSTID VARCHAR(50),REGNO2 VARCHAR(50),REGNO3 VARCHAR(50),"
        MYSQL = MYSQL & " REGNO4 VARCHAR(50),REGNO5 VARCHAR(50),REGNO6 VARCHAR(50),OPTIONS VARCHAR(1),TDATE SMALLDATETIME,LDATE SMALLDATETIME,CCODE INT )"
        MCnn.Execute MYSQL
    Else
        MYSQL = "IF  col_length('SYSTNO','CUSTID') IS  NULL "
        MYSQL = MYSQL & "ALTER TABLE systno ADD CUSTID   VARCHAR(50) NULL DEFAULT 'A' WITH VALUES "
        MCnn.Execute MYSQL
        MYSQL = "IF  col_length('SYSTNO','REGNO2') IS  NULL "
        MYSQL = MYSQL & "ALTER TABLE SYSTNO  ADD REGNO2 VARCHAR(50),REGNO3 VARCHAR(50),REGNO4 VARCHAR(50),REGNO5 VARCHAR(50),REGNO6 VARCHAR(50),OPTIONS VARCHAR(1),TDATE SMALLDATETIME"
        MCnn.Execute MYSQL
        MYSQL = "IF  col_length('SYSTNO','TDATE') is NULL "
        MYSQL = MYSQL & "ALTER TABLE SYSTNO ADD TDATE SMALLDATETIME NOT NULL DEFAULT '2014/03/31' WITH VALUES "
        MCnn.Execute MYSQL
        MYSQL = "IF  col_length('SYSTNO','LDATE') IS  NULL "
        MYSQL = MYSQL & "ALTER TABLE SYSTNO ADD LDATE SmallDateTime NOT NULL DEFAULT '2000/01/01' WITH VALUES "
        MCnn.Execute MYSQL
        MYSQL = "IF  col_length('SYSTNO','CCODE') IS  NULL "
        MYSQL = MYSQL & "ALTER TABLE SYSTNO ADD CCODE  INT NOT NULL DEFAULT 0 WITH VALUES "
        MCnn.Execute MYSQL
    End If
    MCnn.Execute " IF EXISTS (SELECT OBJECT_ID FROM SYS.OBJECTS WHERE TYPE = 'P' AND NAME='INSERT_SYSTNO') DROP PROCEDURE INSERT_SYSTNO"
    MYSQL = "CREATE PROCEDURE INSERT_SYSTNO @SNO VARCHAR(50),@CUSTID VARCHAR(50),@SNO2 VARCHAR(50),@SNO3 VARCHAR(50),@SNO4 VARCHAR(50),@SNO5 VARCHAR(50),@SNO6 VARCHAR(50),@OPTIONS VARCHAR(1),@TDATE SMALLDATETIME,@LDATE SMALLDATETIME,@CCODE INT  AS "
    MYSQL = MYSQL & " INSERT INTO SYSTNO (SERIALNO,CUSTID,REGNO2,REGNO3,REGNO4,REGNO5,REGNO6,OPTIONS,TDATE,LDATE,CCODE) "
    MYSQL = MYSQL & " VALUES ( @SNO, @CUSTID,@SNO2,@SNO3,@SNO4,@SNO5,@SNO6,@OPTIONS,@TDATE,@LDATE,@CCODE) "
    MCnn.Execute MYSQL
    TABLE_NAME = "SYSCOMP"
    LDate = DateValue("2000/01/01")
    MYSQL = "SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[" & TABLE_NAME & "]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1"
    Set MRec = Nothing:    Set MRec = New ADODB.Recordset
    MRec.Open MYSQL, MCnn, adOpenForwardOnly, adLockReadOnly
    If MRec.EOF Then
        MYSQL = "CREATE TABLE SYSCOMP (SNO INT  IDENTITY(1,1),DataBaseName VARCHAR(50) Not Null ,Compcode INT NOT NULL,FinBegin SmallDatetime  NOT NULL,"
        MYSQL = MYSQL & " FinEnd SmallDatetime NOT NULL,RPT_PATH VARCHAR(100), D_PATH VARCHAR(100),SYSLOCKDT SmallDatetime NOT NULL )"
        MCnn.Execute MYSQL
        MYSQL = "SELECT * FROM COMPANY ORDER BY COMPCODE"
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        MRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not MRec.EOF Then
            Do While Not MRec.EOF
                MCompCode = MRec!COMPCODE
                MLFinBeg = MRec!finbegin
                MLFinEnd = MRec!finend
                MRpt_Path = MRec!Rpt_Path
                MDPath = MRec!DPATH
                MSysLockDt = LDate
                MYSQL = "INSERT INTO SYSCOMP (DATABASENAME,COMPCODE,FINBEGIN,FINEND,RPT_PATH,D_PATH,SYSLOCKDT) VALUES"
                MYSQL = MYSQL & "('" & GDatabaseName & "'," & MCompCode & ",'" & Format(MLFinBeg, "YYYY/MM/DD") & "','" & Format(MLFinEnd, "YYYY/MM/DD") & "','" & MRpt_Path & "','" & MDPath & "','" & Format(MSysLockDt, "YYYY/MM/DD") & "')"
                MCnn.Execute MYSQL
                MRec.MoveNext
            Loop
        End If
    End If
    Exit Sub
err1:
    'Resume
    MsgBox err.Description & " :   ", vbCritical, err.HelpFile
End Sub
Private Sub Check_Login()
Dim LHDNo As String:

    On Error GoTo err1:
    Dim MYRS As ADODB.Recordset
    HardDiskNo = GetDriveSerialNumber
    LHDNo = HardDiskNo
    GRegNo = LHDNo
    GUserName = vbNullString
    If EncryptionFlag Then
        GUserName = Get_UserName(txtUsername.text, txtPassword.text)
    Else
        GUserName = Get_UserName(txtUsername.text, EncryptNEW(txtPassword.text, 13))
    End If
    If LenB(GUserName) > 0 Then
        GETMAIN.Show
        MYSQL = "DELETE FROM SELCOMP WHERE COMPCODE =0 "
        Cnn.Execute MYSQL
        Set SelComp_Ado = Nothing
        Set SelComp_Ado = New ADODB.Recordset
        MYSQL = "SELECT COMPCODE,NAME,ACORDER FROM COMPANY ORDER BY COMPCODE"
        SelComp_Ado.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        If CHK_REG(LHDNo) Then
            Registered = True
            If SelComp_Ado.RecordCount = 1 Then
                GAcOrder = SelComp_Ado!ACORDER
                GETMAIN.comsel.Visible = False: GETMAIN.sp2.Visible = False
                Call CompanySelection(Val(SelComp_Ado!COMPCODE))
                If FlagVerAdj Then Call VERSION_ADJUSTMENT
                If FlagIntigrity Then Call IntigrityCheck
                If FlagVerAdj Then Call BROK_CHECK
                
                Call LogIn
                Call Get_Selection(12)
                'If FlagVerAdj Then
                    
                'End If
            ElseIf SelComp_Ado.RecordCount > Val(1) Then
                If FlagVerAdj Then Call VERSION_ADJUSTMENT
                MYSQL = "SELECT COMPCODE FROM SelComp"
                Set MYRS = Nothing: Set MYRS = New ADODB.Recordset
                MYRS.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
                If MYRS.EOF Then
ERR2:
                    SELCOMP.Show
                Else
                    Call CompanySelection(MYRS!COMPCODE)
                    
                End If
                If FlagIntigrity Then Call IntigrityCheck
                If FlagBrok Then Call BROK_CHECK
                Call LogIn
                Call Get_Selection(12)
                
            End If
'
'            '>>>>>>>>> temp
'            Dim Vexcelcolval As String
'            Dim Vitemcode As String
'            Dim Vday As Integer
'            Dim Vmonth As String
'            Dim Vrate As Double
'            Dim Vlastchar As String
'            Dim Vdate As Date
'
'            'Vexcelcolval = "NIY28JN14200P"
'            Vexcelcolval = "BNTY28JN30500P"
'
'            If IsNumeric(Right(Mid(Vexcelcolval, 1, 4), 1)) Then
'                Vitemcode = Left(Vexcelcolval, 3)
'            Else
'                Vitemcode = Left(Vexcelcolval, 4)
'            End If
'            Vexcelcolval = Replace(Vexcelcolval, Vitemcode, "")
'            Vlastchar = Right(Vexcelcolval, 1)
'            Vexcelcolval = Mid(Vexcelcolval, 1, Len(Vexcelcolval) - 1)
'            If IsNumeric(Left(Vexcelcolval, 2)) Then
'                Vday = Left(Vexcelcolval, 2)
'                Vexcelcolval = Mid(Vexcelcolval, 3, Len(Vexcelcolval))
'            Else
'                Vday = Left(Vexcelcolval, 1)
'                Vexcelcolval = Mid(Vexcelcolval, 2, Len(Vexcelcolval))
'            End If
'            Vmonth = Left(Vexcelcolval, 2)
'            Vexcelcolval = Mid(Vexcelcolval, 3, Len(Vexcelcolval))
'            Vrate = Vexcelcolval
'            If Vmonth = "JN" Then
'                Vdate = CStr(Vday) + "/01/" + CStr(Year(Date))
'            ElseIf Vmonth = "FB" Then
'                Vdate = CStr(Vday) + "/02/" + CStr(Year(Date))
'            ElseIf Vmonth = "MR" Then
'                Vdate = CStr(Vday) + "/03/" + CStr(Year(Date))
'            ElseIf Vmonth = "AP" Then
'                Vdate = CStr(Vday) + "/04/" + CStr(Year(Date))
'            ElseIf Vmonth = "MA" Then
'                Vdate = CStr(Vday) + "/05/" + CStr(Year(Date))
'            ElseIf Vmonth = "JN" Then
'                Vdate = CStr(Vday) + "/06/" + CStr(Year(Date))
'            ElseIf Vmonth = "JU" Then
'                Vdate = CStr(Vday) + "/07/" + CStr(Year(Date))
'            ElseIf Vmonth = "AU" Then
'                Vdate = CStr(Vday) + "/08/" + CStr(Year(Date))
'            ElseIf Vmonth = "SE" Then
'                Vdate = CStr(Vday) + "/09/" + CStr(Year(Date))
'            ElseIf Vmonth = "OC" Then
'                Vdate = CStr(Vday) + "/10/" + CStr(Year(Date))
'            ElseIf Vmonth = "NO" Then
'                Vdate = CStr(Vday) + "/11/" + CStr(Year(Date))
'            ElseIf Vmonth = "DE" Then
'                Vdate = CStr(Vday) + "/12/" + CStr(Year(Date))
'            End If
'            '>>>>>>>>> temp end
                        
            
            
            If FlagVerAdj Then Call Check_Opening
                        
            '>>> BUL - 17 JAN 2021 -- Bhavcopy download if not exists -- STARTUK
                FlagLoggedIn = True
                GExCode = ""
                Dim LTExCode As String
                Dim LFileName As String
                Dim ERec As ADODB.Recordset
                Dim Weekdy As Integer
                Weekdy = Weekday(Date)
                FlagLoggedIn = False
                If Weekdy = 1 Or Weekdy = 7 Then 'Sunday, Saturday
                    FlagLoggedIn = False
                End If
                
                If FlagLoggedIn Then
                    MYSQL = "select excode from EXMAST where COMPCODE =" & GCompCode & " and excode in ('NSE', 'NCDX','EQ') order by excode"
                    Set ERec = Nothing
                    Set ERec = New ADODB.Recordset
                    ERec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
                    Do While Not ERec.EOF
                        GExCode = ERec("excode")
                        If ERec("excode") = "EQ" Then
                            LFileName = App.Path & "\" & GExCode & "\" & "cm" & CStr(Left$((Date - 1), 2)) & UCase(Left$(MonthName(Val(month((Date - 1)))), 3)) & CStr(Year((Date - 1))) & "bhav.csv.zip"
                        ElseIf ERec("excode") = "NCDX" Then
                            LFileName = App.Path & "\" & GExCode & "\" & "FO_" & CStr(Left((Date - 1), 2)) & CStr(month((Date - 1))) & Right(CStr(Year((Date - 1))), 2) & "_FINAL.CSV"
                        ElseIf ERec("excode") = "NSE" Then
                            LFileName = App.Path & "\" & GExCode & "\" & "fo" & CStr(Left$((Date - 1), 2)) & UCase(Left$(MonthName(Val(month(Date - 1))), 3)) & Year(Date - 1) & "bhav.csv"
                        End If
                        If Not FileExist(LFileName) Then
                            If Weekdy > 2 And Weekdy < 7 Then
                                frmdata.vcDTP1.Value = Date - 1
                                frmdata.vcDTP2.Value = Date - 1
                            ElseIf Weekdy = 2 Then 'if Monday -- Get date of last Friday
                                frmdata.vcDTP1.Value = Date - 3
                                frmdata.vcDTP2.Value = Date - 3
                            End If
                            If ERec("excode") = "NCDX" Then
                                frmdata.CHKNCDXExcelClosing.Value = 1
                            ElseIf ERec("excode") = "NSE" Then
                                frmdata.ChkNSEBhavCopy.Value = 1
                            ElseIf ERec("excode") = "EQ" Then
                                frmdata.ChkNSEEQClosing.Value = 1
                            End If
                            frmdata.Frame3.Visible = False
                            frmdata.Frame4.Visible = False
                            frmdata.Frame5.Visible = False
                            frmdata.Label5.Visible = False
                            frmdata.Label6.Visible = False
                            frmdata.Label8.Visible = False
                            frmdata.Label9(7).Visible = False
                            frmdata.Label9(5).Visible = False
                            frmdata.Label4.Visible = False
                            frmdata.Label15.Visible = False
                            frmdata.Label11.Visible = False
                            frmdata.Label10.Visible = True
                            frmdata.okcmdclick
                            
                            Exit Do
                        End If
                        ERec.MoveNext
                    Loop
                End If
            '>>>  - 17 JAN 2021 -- Bhavcopy download if not exists ENDUK
            Unload Me
            Exit Sub
        Else
            MsgBox "Please Call Sauda Support Staff to Get your Software Registered Reg ID: " & HardDiskNo & ""
            Unload Me
            Exit Sub
        End If
    Else
        MsgBox "Sorry Invalid User Try Again"
        txtUsername.text = vbNullString
        txtPassword.text = vbNullString
        txtUsername.SetFocus
        Exit Sub
    End If
    
    Set MCnn = Nothing
    Unload Me
    Exit Sub
err1:
    If err.Number <> 0 Then
        MsgBox err.Description
        'Resume
        MsgBox err.Description & " :   ", vbCritical, err.HelpFile
    End If
End Sub

Private Sub txtUsername_Validate(Cancel As Boolean)
    If LenB(txtUsername.text) = 0 Then
'        MsgBox "Please Enter User Name.", vbCritical
'        Cancel = True
    End If
End Sub

Private Sub Check_Opening()

Dim TRec As ADODB.Recordset
Dim TRec2 As ADODB.Recordset
Dim LMinDate  As Date
Dim LMinDate2  As Date
LMinDate2 = GFinBegin
MYSQL = "SELECT MIN(CONDATE) AS MDATE  FROM CTR_D WHERE COMPCODE =" & GCompCode & ""
Set TRec = Nothing
Set TRec = New ADODB.Recordset
TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
If Not TRec.EOF Then
    If Not IsNull(TRec!MDate) Then
        LMinDate = TRec!MDate
        MYSQL = "SELECT MIN(VOU_DT) AS MDATE FROM VCHAMT  WHERE COMPCODE =" & GCompCode & ""
        Set TRec2 = Nothing
        Set TRec2 = New ADODB.Recordset
        TRec2.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec2.EOF Then
            If Not IsNull(TRec2!MDate) Then
                If LMinDate > TRec!MDate Then
                    LMinDate2 = TRec2!MDate
                Else
                    LMinDate2 = LMinDate
                End If
            End If
        End If
        MYSQL = "UPDATE COMPANY SET FINBEGIN='" & Format(LMinDate2, "YYYY/MM/DD") & "' WHERE COMPCODE =" & GCompCode & ""
        Cnn.Execute MYSQL
        
        MYSQL = "UPDATE SYSCOMP SET FINBEGIN='" & Format(LMinDate2, "YYYY/MM/DD") & "' WHERE COMPCODE =" & GCompCode & " AND DATABASENAME ='" & GDatabaseName & "'"
        MCnn.Execute MYSQL
        
        MYSQL = "SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND "
        MYSQL = MYSQL & " EXISTS (SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PATTAN ='O' AND CONDATE >'" & Format(LMinDate, "YYYY/MM/DD") & "')"
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Please call Sauda Support Staff As there are opening Trade Beyond Open Date. "
        End If
    End If
End If
End Sub
