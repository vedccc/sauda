VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000000&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2220
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000000&
      Caption         =   "OK"
      Height          =   390
      Left            =   600
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC As ADODB.Recordset
Dim REC1 As ADODB.Recordset
Private Sub cmdCancel_Click()
    End
End Sub
Private Sub cmdOK_Click()
On Error GoTo ERR1
    MYSQL = "SELECT * FROM USERMASTER WHERE USER_NAME='" & txtUsername.Text & "' AND PASSWD = '" & txtPassword.Text & "'"
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    If Not REC.EOF Then
        USER_ID = REC!USER_NAME
        GETMAIN.SELCOMP_ADO.ConnectionString = CNN
        GETMAIN.SELCOMP_ADO.RecordSource = "SELECT * FROM COMPANY ORDER BY COMPANY.COMPCODE"
        GETMAIN.SELCOMP_ADO.Refresh
        GETMAIN.Show
        If GETMAIN.SELCOMP_ADO.Recordset.RecordCount = Val(1) Then
            Call companyselection(Val(GETMAIN.SELCOMP_ADO.Recordset!CompCode))
        Else
            SELCOMP.Show
        End If
        Call LogIn
        Call Get_Selection(12)
        Unload Splash
        Unload Me
        Exit Sub
    Else
        MsgBox "Invalid User. Try again....", vbInformation, "Error"
        frmLogin.txtUsername.Text = ""
        frmLogin.txtPassword.Text = ""
        frmLogin.txtUsername.SetFocus
    End If
    Exit Sub
ERR1:
    MsgBox Err.Number & " :   " & Err.Description, vbCritical, Err.HelpFile
End Sub
