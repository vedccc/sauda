VERSION 5.00
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Begin VB.Form frmSysLDt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "System Lock Date"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtnew 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtold 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.TextBox txtmpwd 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1080
   End
   Begin VB.CommandButton OkCmd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1080
   End
   Begin vcDateTimePicker.vcDTP vcDTP3 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
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
      Value           =   39796.5153356481
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Password:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      Height          =   3300
      Left            =   30
      Top             =   480
      Width           =   5445
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Settlemt Lock Date:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1860
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Settlement Lock Date"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmSysLDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 And Not Label6.Visible Then
        txtold.PasswordChar = ""
        txtnew.PasswordChar = ""
        txtmpwd.PasswordChar = ""
        Label6.Visible = True
    ElseIf KeyCode = 114 And Label6.Visible Then
        txtold.PasswordChar = "*"
        txtnew.PasswordChar = "*"
        txtmpwd.PasswordChar = "*"
        Label6.Visible = False
    End If
End Sub
Private Sub Form_Load()
    OkCmd.Enabled = False
    vcDTP3.Value = GSysLockDt
End Sub
Private Sub Label5_Click()
    Frame1.Visible = True
    txtold.text = ""
    txtnew.text = ""
    txtold.SetFocus
    OkCmd.Enabled = True
End Sub



Private Sub OkCmd_Click()

    If Frame1.Visible Then
        If txtnew.text = "" Then
            MsgBox "Password can't be blank!!!", vbCritical
            txtnew.SetFocus
            Exit Sub
        ElseIf txtold.text = "" Then
            MsgBox "Password can't be blank!!!", vbCritical
            txtold.SetFocus
            Exit Sub
        ElseIf txtold.text = txtnew.text Then
            MsgBox "Password can't be same!!!", vbCritical
            txtold.SetFocus
            Exit Sub
        End If
        Cnn.BeginTrans
            mysql = "update USERMASTER set PASSWD='" & EncryptNEW(txtnew.text, 13) & "' where USER_NAME='Master'"
            Cnn.Execute mysql
        Cnn.CommitTrans
        Frame1.Visible = False
        OkCmd.Enabled = False
    Else
        mysql = "UPDATE COMPANY SET SYSLOCKDT = '" & Format(vcDTP3.Value, "YYYY/MM/DD") & "' WHERE COMPCODE =" & GCompCode & ""
        Cnn.Execute mysql
        
        Dim CnnString As String
        ServerString = MServer
        ServerString = Left$(MServer, (InStr(1, MServer, "database") + 8)) & "MASTER"
        CnnString = ServerString
        
        If MCnn.State = 0 Then ' 0=closed
            Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
            MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
            MCnn.Open
        End If
        mysql = "UPDATE SYSCOMP SET SYSLOCKDT = '" & Format(vcDTP3.Value, "YYYY/MM/DD") & "' WHERE COMPCODE =" & GCompCode & ""
        MCnn.Execute mysql
    
        GSysLockDt = vcDTP3.Value
            
        Unload Me
    End If
End Sub
Private Sub txtmpwd_Validate(Cancel As Boolean)
    Label2.Visible = False
    If txtmpwd.text <> "" Then
    
        
        
        Dim MRec As ADODB.Recordset
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        mysql = "SELECT USER_NAME FROM USERMASTER WHERE USER_NAME ='Master' and PASSWD='" & EncryptNEW(txtmpwd.text, 13) & "'"
        MRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not MRec.EOF Then
            OkCmd.Enabled = True
        Else
            Label2.Visible = True
            Label2.Caption = "Invalid master password!!!"
            txtmpwd.SetFocus
        End If
    End If
End Sub
