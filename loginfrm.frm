VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form loginfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   5540
      TabIndex        =   28
      Top             =   6600
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "This Software is mainly for Training Purpose and will stop working after 15 Days"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   0
         TabIndex        =   30
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "This is a Demo Software.  All the Accounts, Trades and Ledger Positings are ficticious."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   0
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   9255
         Begin VB.CommandButton Command3 
            Caption         =   "Restore Database Backup"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   24
            Top             =   3720
            Width           =   4455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Select"
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
            Left            =   6000
            TabIndex        =   23
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
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
            IMEMode         =   3  'DISABLE
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   2040
            Width           =   5535
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
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
            TabIndex        =   21
            Top             =   1080
            Width           =   1935
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   5280
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "Please wait ..."
            BeginProperty Font 
               Name            =   "Verdana"
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
            Top             =   3120
            Visible         =   0   'False
            Width           =   7935
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Database Backup file to be Restore"
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
            Left            =   360
            TabIndex        =   26
            Top             =   1800
            Width           =   3975
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Database Name"
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
            Left            =   360
            TabIndex        =   25
            Top             =   840
            Width           =   2415
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   4455
            Left            =   120
            Top             =   120
            Width           =   9015
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   5520
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   6
            Top             =   2525
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   1325
            Width           =   2055
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFFFF&
            X1              =   840
            X2              =   2880
            Y1              =   2775
            Y2              =   2775
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            X1              =   840
            X2              =   2880
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   5
            Left            =   240
            TabIndex        =   19
            Top             =   2040
            Width           =   2040
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "New Password"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   2010
         End
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
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
         Left            =   6720
         TabIndex        =   1
         Top             =   2995
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   5520
         TabIndex        =   3
         Top             =   4920
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   6720
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   5520
         TabIndex        =   9
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "User Id"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   5520
         TabIndex        =   8
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Hello! Let's get started"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   5520
         TabIndex        =   11
         Top             =   1080
         Width           =   2280
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   2
         Left            =   5520
         TabIndex        =   10
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "88897-40123, 88898-40123, 88390-84261, 88390-85057"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   6960
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "INDORE- M.P. 452001"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   6480
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Back-Office Accounting Software For Commodity And Stock Exchanges"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Software"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   1440
         TabIndex        =   13
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   6720
         X2              =   8760
         Y1              =   4095
         Y2              =   4095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   6720
         X2              =   8760
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Sauda"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2415
         Left            =   0
         TabIndex        =   4
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   8775
         Left            =   5400
         Top             =   -120
         Width           =   5415
      End
   End
End
Attribute VB_Name = "loginfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EncryptionFlag As Boolean:  Dim HardDiskNo As String:           'Dim MCnn As ADODB.Connection
Dim LDEMO As String:            Dim FlagVerAdj As Boolean:          Dim FlagIntigrity As Boolean
Dim Flag_Pitbrok As Boolean:    Dim LCompRec As ADODB.Recordset:    Dim GNew_Sub_Brok_Updt As Boolean
Dim NEWSOFT0 As Boolean:        Dim Rec As ADODB.Recordset
Dim DBRestoreflag  As Boolean:
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

If LenB(Text1.text) = 0 Then
    MsgBox "Please Enter new Password.", vbCritical
    Text1.SetFocus
ElseIf LenB(Text2.text) = 0 Then
    MsgBox "Please Enter confirm Password.", vbCritical
    Text2.SetFocus
ElseIf Text1.text <> Text2.text Then
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
        mysql = "UPDATE USERMASTER SET PASSWD='" & EncryptNEW(Text1.text, 13) & "' WHERE USER_NAME ='" & txtUsername.text & "'"
        Cnn.Execute mysql
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
'
'Private Sub LUnZipFile(LPZipFile As String, LPTarget As String)
'Dim oUnZip As CGUnzipFiles
'    Set oUnZip = New CGUnzipFiles
'    With oUnZip
'        .ZipFileName = LPZipFile
'        .ExtractDir = LPTarget
'        .HonorDirectories = False
'        If .Unzip <> 0 Then MsgBox .GetLastMessage
'    End With
'    Set oUnZip = Nothing
'    Exit Sub
'End Sub

Private Sub Command2_Click()

On Error GoTo err1

'Dim LFileName As String
'Dim LRemoteFileName As String
'Dim LRenameFile As String
'
'    LFileName = App.Path & "\nse\" & "F_CN01_NSE_21052021.csv"
'    LRemoteFileName = "/faocommon/marketreports/" & "F_CN01_NSE_21052021.csv.gz" & ""
'    'If Not FileExist(LFileName) Then
'        Call Connect_Ftp("ftp.connect2nse.com", "FAOGUEST", "FAOGUEST", LRemoteFileName, LFileName)
'    'End If
'
'    LRenameFile = RenameFileOrDir(LFileName, App.Path & "\nse\" & "F_CN01_NSE_21052021.zip")
'    'LTargetFileName = App.Path & "\" & LTExCode & "\" & "F_CN01_NSE_" & CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & Year(LTradeDt) & ".csv"
'    'If FileExist(LTargetFileName) Then
'    Call LUnZipFile(LRenameFile, App.Path & "\nse\" & "F_CN01_NSE_21052021.csv")
'    'TxtPath = "F_CN01_NSE_" & CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & Year(LTradeDt) & ".csv"
                    

'    Dim NewRec As ADODB.Recordset
'    Dim TRec As ADODB.Recordset
'    Dim LCompCode As String
'    Dim LCondate As Date
'    Dim MCount As Integer
'
'    MYSQL = "SELECT COMPCODE FROM COMPANY"
'    Set NewRec = Nothing
'    Set NewRec = New ADODB.Recordset
'    NewRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'    Do While Not NewRec.EOF
'        LCompCode = NewRec!COMPCODE
'        Set TRec = Nothing
'        Set TRec = New ADODB.Recordset
'        MYSQL = "SELECT DISTINCT CONDATE,ROWNO FROM CTR_D WHERE COMPCODE =" & LCompCode & " ORDER BY CONDATE,ROWNO"
'        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'        Do While Not TRec.EOF
'            DoEvents
'            MCount = 0
'            'GETMAIN.Label1.Caption = "Generating ConNo " & TRec!Condate & ""
'            LCondate = TRec!Condate
'            Do While LCondate = TRec!Condate
'                'DoEvents
'            '    GETMAIN.Label1.Caption = "Generating ConNo " & TRec!Condate & " " & MCount & ""
'                MCount = MCount + 1
'
'                MYSQL = " UPDATE CTR_D SET CONNO = " & MCount & " WHERE COMPCODE=" & LCompCode & " "
'                MYSQL = MYSQL & " AND CONDATE ='" & Format(LCondate, "YYYY/MM/DD") & "'"
'                MYSQL = MYSQL & " AND ROWNO =" & TRec!ROWNO & " "
'                Cnn.Execute MYSQL
'
'                TRec.MoveNext
'                If TRec.EOF Then Exit Do
'
'                MYSQL = " UPDATE CTR_D SET CONNO = " & MCount & "  WHERE COMPCODE=" & LCompCode & " "
'                MYSQL = MYSQL & " AND CONDATE ='" & Format(LCondate, "YYYY/MM/DD") & "'"
'                MYSQL = MYSQL & " AND ROWNO =" & TRec!ROWNO & " "
'                Cnn.Execute MYSQL
'
'                TRec.MoveNext
'               If TRec.EOF Then Exit Do
'
'            Loop
'            If TRec.EOF Then Exit Do
'        Loop
'        NewRec.MoveNext
'    Loop
    
    
    If txtUsername.text = "" Then
        MsgBox "Please enter username!!!"
        txtUsername.SetFocus
    ElseIf txtPassword.text = "" Then
        MsgBox "Please enter password!!!"
        txtPassword.SetFocus
    Else
'        ServerString = MServer
'        ServerString = Left$(MServer, (InStr(1, MServer, "database") + 8)) & "MASTER"
'        CnnString = ServerString
'        Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
'        MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
        If FlagVerAdj Then
            If MCnn.State = 0 Then
                MCnn.Open
            End If
            
            Call Set_SystemTable
            Call ColChanges
        End If
        
        Call Check_Login
    End If

err1:

If err.Number <> 0 Then
    MsgBox "Err1:" & err.Description
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
        
        ServerString = MServer
        ServerString = Left$(MServer, (InStr(1, MServer, "database") + 8)) & "MASTER"
        CnnString = ServerString
        Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
        MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
        MCnn.Open
    
        mysql = "EXECUTE RESTOREDB '" & Text3.text & "','" & Text4.text & "'"
        MCnn.Execute mysql
        
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
    If KeyCode = 27 Then
        If Frame4.Visible Then
            Frame4.Visible = False
            Check1.Value = 0
        ElseIf Frame6.Visible Then
            Frame6.Visible = False
        Else
            Unload Me
        End If
    
    End If
    
    If KeyCode = 114 Then 'F3
        FlagVerAdj = True
    End If
    If KeyCode = 115 Then FlagBrok = True
    If KeyCode = 116 Then FlagIntigrity = True
    If KeyCode = 121 Then FlagStored = True
    If KeyCode = 117 Then Flag_Pitbrok = True
    If KeyCode = 113 Then EncryptionFlag = True
    If KeyCode = 120 Then
        Call Registration
    End If
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
  
  Dim Rec As ADODB.Recordset
  FIELD_FOUND = True
  DBRestoreflag = False
  If App.PrevInstance = True Then
    MsgBox "Sauda Application is already Open"
    Unload Me
    Exit Sub
  End If
  
    FlagVerAdj = False
    FlagIntigrity = False
    GNew_Sub_Brok_Updt = False
    Flag_Pitbrok = False
    
    ServerString = MServer
    ServerString = Left$(MServer, (InStr(1, MServer, "database") + 8)) & "MASTER"
    Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = ServerString
    MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
        
    Exit Sub
err1:
    If Val(err.Number) = Val(-2147217900) Then  '       Z x'FIELD IS NOT IN THE TABLE
        FIELD_FOUND = False
    Else
        MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    End If
End Sub
Private Sub Form_Paint()
    txtUsername.SetFocus
    If LDEMO = "Y" Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
        If Not DBRestoreflag Then
            Me.Height = 7935
        End If
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFFFFF
    Text1.ForeColor = &H0&
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &HFF8080
    Text1.ForeColor = &HFFFFFF
End Sub
'Private Sub Text1_Validate(Cancel As Boolean)
'If LenB(Text1.text) = 0 Then
'    MsgBox "Please Enter new Password.", vbCritical
'    Cancel = True
'End If
'End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFFFFF
    Text2.ForeColor = &H0&
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &HFF8080
    Text2.ForeColor = &HFFFFFF
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.BackColor = &HFFFFFF
    txtPassword.ForeColor = &H0&
End Sub
'Private Sub Text2_Validate(Cancel As Boolean)
'    If LenB(Text2.text) = 0 Then
'        MsgBox "Please Enter Password.", vbCritical
'        Cancel = True
'    End If
'End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2.SetFocus
    End If
        
End Sub
Private Sub txtPassword_LostFocus()
    txtPassword.BackColor = &HFF8080
    txtPassword.ForeColor = &HFFFFFF
End Sub
Private Sub txtPassword_Validate(Cancel As Boolean)
    
    If LenB(txtPassword.text) < 1 Then
'        MsgBox "Please Enter Password.", vbCritical
'        Cancel = True
    Else
        Command2.SetFocus
    End If
End Sub

Private Sub txtUsername_GotFocus()
    txtUsername.BackColor = &HFFFFFF
    txtUsername.ForeColor = &H0&
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
    mysql = "CREATE PROCEDURE COL_CHNAGES AS SET NOCOUNT ON BEGIN "
    mysql = mysql & " DELETE FROM USER_RIGHTS WHERE MENUNAME LIKE 'mnufilefor%'"
    
    mysql = mysql & " IF  col_length('COMPANY','BillingCycle') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD BillingCycle BIT NOT NULL DEFAULT 1 WITH VALUES"
    
    mysql = mysql & " IF  col_length('COMPANY','QTY_DECIMAL') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD QTY_DECIMAL BIT NOT NULL DEFAULT 0 WITH VALUES"
    
    mysql = mysql & " IF  col_length('COMPANY','ConNoType') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD ConNoType INT NOT NULL DEFAULT 0 WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','ONLYBROK') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD ONLYBROK INT NOT NULL DEFAULT  0 WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','STMDT') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD STMDT SMALLDATETIME  NOT NULL DEFAULT  '2015/06/01' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','SHARE') IS NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD SHARE VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','MARGIN') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD MARGIN  VARCHAR(1) NOT NULL DEFAULT 'N' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','EQ') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD EQ VARCHAR(1) NOT NULL DEFAULT 'N' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','GENQUERY') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD GENQUERY VARCHAR(1) NOT NULL DEFAULT '1' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','SHOWLOT') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD SHOWLOT VARCHAR(1) NOT NULL DEFAULT  'N' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','SEBIREGNO') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD SEBIREGNO VARCHAR(20) NOT NULL DEFAULT  '' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','GSTIN') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD GSTIN VARCHAR(20) NOT NULL DEFAULT  '' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','UNIQCLIENTID') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD UNIQCLIENTID VARCHAR(50) NOT NULL DEFAULT 'A' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','TRANFEES') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD TRANFEES VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','STAMPDUTY') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD STAMPDUTY VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','VALUEWISE') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD VALUEWISE VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','STANDING') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD STANDING VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','MINBROKYN') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD MINBROKYN VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','STT') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD STT VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','SUBBROK') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD SUBBROK VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','SRVTAX') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD SRVTAX VARCHAR(1) NOT NULL DEFAULT 'Y' WITH VALUES"
    mysql = mysql & " IF  col_length('SAUDAMAST','OPTTYPE') IS  NULL"
    mysql = mysql & " ALTER TABLE SAUDAMAST ADD OPTTYPE VARCHAR(1) NULL DEFAULT '' WITH VALUES"
    mysql = mysql & " IF  col_length('SAUDAMAST','INSTTYPE') IS  NULL"
    mysql = mysql & " ALTER TABLE SAUDAMAST ADD INSTTYPE VARCHAR(3) NULL DEFAULT 'FUT' WITH VALUES"
    mysql = mysql & " IF  col_length('SAUDAMAST','STRIKEPRICE ') IS  NULL"
    mysql = mysql & " ALTER TABLE SAUDAMAST ADD STRIKEPRICE  FLOAT  NULL DEFAULT 0 WITH VALUES END"
    mysql = mysql & " IF  col_length('COMPANY','SHOWSTD') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD SHOWSTD VARCHAR(1) NOT NULL DEFAULT 'N' WITH VALUES"
    mysql = mysql & " IF  col_length('COMPANY','AMTDIVIDE') IS  NULL"
    mysql = mysql & " ALTER TABLE COMPANY ADD AMTDIVIDE VARCHAR(1) NOT NULL DEFAULT '0' WITH VALUES"
    Cnn.Execute mysql
    
    'MYSQL = "EXEC COL_CHNAGES"
    'Cnn.Execute MYSQL
    
    mysql = "IF  col_length('SYSCOMP','SYSLOCKDT') IS  NULL "
    mysql = mysql & "ALTER TABLE SYSCOMP ADD SYSLOCKDT  SmallDateTime NOT NULL DEFAULT '2000/01/01' WITH VALUES "
    MCnn.Execute mysql

    mysql = "UPDATE SYSCOMP SET SYSLOCKDT ='2000/01/01' WHERE SYSLOCKDT IS NULL"
    MCnn.Execute mysql
    
    Cnn.Execute "  IF EXISTS(SELECT OBJECT_ID FROM SYS.OBJECTS WHERE TYPE='P' AND NAME='GET_USERNAME') DROP PROCEDURE GET_USERNAME "
    mysql = "CREATE PROCEDURE GET_USERNAME @USERNAME VARCHAR(12), @PASSWD VARCHAR(12),@PUSERNAME VARCHAR(12) OUTPUT"
    mysql = mysql & " AS SET NOCOUNT ON BEGIN DECLARE @LPUSERNAME AS VARCHAR(12) SET @PUSERNAME =''"
    mysql = mysql & " SELECT @LPUSERNAME =USER_NAME FROM USERMASTER WHERE USER_NAME =@USERNAME AND PASSWD = @PASSWD "
    mysql = mysql & " IF @LPUSERNAME IS NOT NULL SET @PUSERNAME   =@LPUSERNAME  End"
    Cnn.Execute mysql
    
End Sub

Private Sub Set_SystemTable()
On Error GoTo err1
    Dim TABLE_NAME As String:    Dim MRec As ADODB.Recordset:    Dim MCompCode As Integer:    Dim ldate As Date
    Dim MLFinBeg As Date:        Dim MLFinEnd As Date:           Dim MRpt_Path As String:     Dim MDPath As String
    Dim MSysLockDt As Date
                
    TABLE_NAME = "SYSTNO"
    mysql = "SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[" & TABLE_NAME & "]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1"
    Set MRec = Nothing
    Set MRec = New ADODB.Recordset
    MRec.Open mysql, MCnn, adOpenForwardOnly, adLockReadOnly
    If MRec.EOF Then
        mysql = "CREATE TABLE SYSTNO (SNO INT IDENTITY(1,1),SERIALNO VARCHAR(50) NOT NULL,CUSTID VARCHAR(50),REGNO2 VARCHAR(50),REGNO3 VARCHAR(50),"
        mysql = mysql & " REGNO4 VARCHAR(50),REGNO5 VARCHAR(50),REGNO6 VARCHAR(50),OPTIONS VARCHAR(1),TDATE SMALLDATETIME,LDATE SMALLDATETIME,CCODE INT )"
        MCnn.Execute mysql
    Else
        mysql = "IF  col_length('SYSTNO','CUSTID') IS  NULL "
        mysql = mysql & "ALTER TABLE systno ADD CUSTID   VARCHAR(50) NULL DEFAULT 'A' WITH VALUES "
        MCnn.Execute mysql
        mysql = "IF  col_length('SYSTNO','REGNO2') IS  NULL "
        mysql = mysql & "ALTER TABLE SYSTNO  ADD REGNO2 VARCHAR(50),REGNO3 VARCHAR(50),REGNO4 VARCHAR(50),REGNO5 VARCHAR(50),REGNO6 VARCHAR(50),OPTIONS VARCHAR(1),TDATE SMALLDATETIME"
        MCnn.Execute mysql
        mysql = "IF  col_length('SYSTNO','TDATE') is NULL "
        mysql = mysql & "ALTER TABLE SYSTNO ADD TDATE SMALLDATETIME NOT NULL DEFAULT '2014/03/31' WITH VALUES "
        MCnn.Execute mysql
        mysql = "IF  col_length('SYSTNO','LDATE') IS  NULL "
        mysql = mysql & "ALTER TABLE SYSTNO ADD LDATE SmallDateTime NOT NULL DEFAULT '2000/01/01' WITH VALUES "
        MCnn.Execute mysql
        mysql = "IF  col_length('SYSTNO','CCODE') IS  NULL "
        mysql = mysql & "ALTER TABLE SYSTNO ADD CCODE  INT NOT NULL DEFAULT 0 WITH VALUES "
        MCnn.Execute mysql
    End If
    
    MCnn.Execute " IF EXISTS (SELECT OBJECT_ID FROM SYS.OBJECTS WHERE TYPE = 'P' AND NAME='INSERT_SYSTNO') DROP PROCEDURE INSERT_SYSTNO"
    mysql = "CREATE PROCEDURE INSERT_SYSTNO @SNO VARCHAR(50),@CUSTID VARCHAR(50),@SNO2 VARCHAR(50),@SNO3 VARCHAR(50),@SNO4 VARCHAR(50),@SNO5 VARCHAR(50),@SNO6 VARCHAR(50),@OPTIONS VARCHAR(1),@TDATE SMALLDATETIME,@LDATE SMALLDATETIME,@CCODE INT  AS "
    mysql = mysql & " INSERT INTO SYSTNO (SERIALNO,CUSTID,REGNO2,REGNO3,REGNO4,REGNO5,REGNO6,OPTIONS,TDATE,LDATE,CCODE) "
    mysql = mysql & " VALUES ( @SNO, @CUSTID,@SNO2,@SNO3,@SNO4,@SNO5,@SNO6,@OPTIONS,@TDATE,@LDATE,@CCODE) "
    MCnn.Execute mysql
        
    TABLE_NAME = "SYSCOMP"
    ldate = DateValue("2000/01/01")
    mysql = "SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[" & TABLE_NAME & "]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1"
    Set MRec = Nothing:    Set MRec = New ADODB.Recordset
    MRec.Open mysql, MCnn, adOpenForwardOnly, adLockReadOnly
    If MRec.EOF Then
        mysql = "CREATE TABLE SYSCOMP (SNO INT  IDENTITY(1,1),DataBaseName VARCHAR(50) Not Null ,Compcode INT NOT NULL,FinBegin SmallDatetime  NOT NULL,"
        mysql = mysql & " FinEnd SmallDatetime NOT NULL,RPT_PATH VARCHAR(100), D_PATH VARCHAR(100),SYSLOCKDT SmallDatetime NOT NULL )"
        MCnn.Execute mysql
        mysql = "SELECT * FROM COMPANY ORDER BY COMPCODE"
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        MRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not MRec.EOF Then
            Do While Not MRec.EOF
                MCompCode = MRec!CompCode
                MLFinBeg = MRec!finbegin
                MLFinEnd = MRec!finend
                MRpt_Path = MRec!Rpt_Path
                MDPath = MRec!DPATH
                MSysLockDt = ldate
                mysql = "INSERT INTO SYSCOMP (DATABASENAME,COMPCODE,FINBEGIN,FINEND,RPT_PATH,D_PATH,SYSLOCKDT) VALUES"
                mysql = mysql & "('" & GDatabaseName & "'," & MCompCode & ",'" & Format(MLFinBeg, "YYYY/MM/DD") & "','" & Format(MLFinEnd, "YYYY/MM/DD") & "','" & MRpt_Path & "','" & MDPath & "','" & Format(MSysLockDt, "YYYY/MM/DD") & "')"
                MCnn.Execute mysql
                MRec.MoveNext
            Loop
        End If
    End If
    Exit Sub
err1:
    
    MsgBox err.Description & " :   ", vbCritical, err.HelpFile
End Sub
Private Sub Check_Login()

Dim LHDNo As String:
Dim errstep As String
    
    VER_ADJ "COMPANY", "ECODE", "INT", , "0"
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
    
    'GUserName = DecryptNEW("@?>", 13)
    
    'GUserName = "s"
    If LenB(GUserName) > 0 Then
    
        '>>> Master user and password
        Dim MRec As ADODB.Recordset
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        mysql = "SELECT USER_NAME FROM USERMASTER WHERE USER_NAME ='Master'"
        MRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If MRec.EOF Then
            Cnn.BeginTrans
                mysql = "INSERT INTO USERMASTER(USER_NAME, PASSWD, MAINUSER, SW) VALUES('Master','" & EncryptNEW("54321", 13) & "',0,'" & SW & "')"
                Cnn.Execute mysql
            Cnn.CommitTrans
        End If
           
       Set MYRS = Nothing
       Set MYRS = New ADODB.Recordset
       mysql = "SELECT ECODE FROM COMPANY"
       MYRS.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If MYRS!ECODE <> "191023" Then
            FlagVerAdj = True
        End If
        GETMAIN.Show
        mysql = "DELETE FROM SELCOMP WHERE COMPCODE =0 "
        Cnn.Execute mysql
        Set SelComp_Ado = Nothing
        Set SelComp_Ado = New ADODB.Recordset
        mysql = "SELECT COMPCODE,NAME,ACORDER FROM COMPANY ORDER BY COMPCODE"
        SelComp_Ado.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If CHK_REG(LHDNo) Then
            Registered = True
            If SelComp_Ado.RecordCount = 1 Then
                GAcOrder = SelComp_Ado!ACORDER
                GETMAIN.comsel.Visible = False: GETMAIN.sp2.Visible = False
                
                If FlagVerAdj Then Call VERSION_ADJUSTMENT
                
                Call CompanySelection(Val(SelComp_Ado!CompCode))
                If FlagIntigrity Then Call IntigrityCheck
                If FlagVerAdj Then Call BROK_CHECK

                Call LogIn
                Call Get_Selection(12)
                
                'If FlagVerAdj Then
                'End If
            ElseIf SelComp_Ado.RecordCount > Val(1) Then
                If FlagVerAdj Then Call VERSION_ADJUSTMENT
                mysql = "SELECT COMPCODE FROM SelComp"
                Set MYRS = Nothing: Set MYRS = New ADODB.Recordset
                MYRS.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                If MYRS.EOF Then
ERR2:
                    SELCOMP.Show
                Else
                    Call CompanySelection(MYRS!CompCode)
                    
                End If
                If FlagIntigrity Then Call IntigrityCheck
                If FlagBrok Then Call BROK_CHECK
                Call LogIn
                Call Get_Selection(12)
            End If
            
            If FlagVerAdj Then Call Check_Opening
                        
            '>>> BUL - 17 JAN 2021 -- Bhavcopy download if not exists -- STARTUK
            
                If FlagDataImport = "Y" Then  'And GHDNO <> "1713138570"
                    FlagLoggedIn = True
                Else
                    FlagLoggedIn = False
                End If
                
                GEXCODE = ""
                Dim LTExCode As String
                Dim LFileName As String
                Dim ERec As ADODB.Recordset
                Dim Weekdy As Integer
                Dim filedate As Date
                Dim LFileSystemObject As FileSystemObject
                Set LFileSystemObject = CreateObject("Scripting.FileSystemObject")
                                
                Weekdy = Weekday(Date)

'                If Weekdy = 1 Or Weekdy = 7 Then 'Sunday, Saturday
'                    FlagLoggedIn = False
'                End If
                
                If Weekdy = 1 Then 'Sunday
                    filedate = Date - 2 '>> file of Friday
                ElseIf Weekdy = 7 Then 'Saturday
                    filedate = Date - 1  '>> file of Friday
                ElseIf Weekdy = 2 Then 'Monday
                    filedate = Date - 3  '>> file of Friday
                Else
                    filedate = Date - 1
                End If
                
                If FlagLoggedIn Then
                
                    Dim VFiledownload As Boolean
                    Dim Vdate As String
                    Dim Vdate1 As String
'''''                    If Len(month(Date - 1)) = 1 Then
'''''                        Vdate = "0" & month(Date - 1)
'''''                    Else
'''''                        Vdate = month(Date - 1)
'''''                    End If
'''''                    If Len(Day(Date - 1)) = 1 Then
'''''                        Vdate = Vdate & "-" & "0" & Day(Date - 1)
'''''                    Else
'''''                        Vdate = Vdate & "-" & Day(Date - 1)
'''''                    End If
'''''                    Vdate = Vdate & "-" & Year(Date - 1)
'''''
'''''                    Vdate1 = Year(Date - 1)
'''''                    If Len(month(Date - 1)) = 1 Then
'''''                        Vdate1 = Vdate1 & "0" & month(Date - 1)
'''''                    Else
'''''                        Vdate1 = Vdate1 & month(Date - 1)
'''''                    End If
'''''                    If Len(Day(Date - 1)) = 1 Then
'''''                        Vdate1 = Vdate1 & "0" & Day(Date - 1)
'''''                    Else
'''''                        Vdate1 = Vdate1 & Day(Date - 1)
'''''                    End If
                    If Len(month(filedate)) = 1 Then
                        Vdate = "0" & month(filedate)
                    Else
                        Vdate = month(filedate)
                    End If
                    If Len(Day(filedate)) = 1 Then
                        Vdate = Vdate & "-" & "0" & Day(filedate)
                    Else
                        Vdate = Vdate & "-" & Day(filedate)
                    End If
                    Vdate = Vdate & "-" & Year(filedate)
                    
                    Vdate1 = Year(filedate)
                    If Len(month(filedate)) = 1 Then
                        Vdate1 = Vdate1 & "0" & month(filedate)
                    Else
                        Vdate1 = Vdate1 & month(filedate)
                    End If
                    If Len(Day(filedate)) = 1 Then
                        Vdate1 = Vdate1 & "0" & Day(filedate)
                    Else
                        Vdate1 = Vdate1 & Day(filedate)
                    End If
                                                                        
                    mysql = "select excode from EXMAST where COMPCODE =" & GCompCode & " and excode in ('NSE', 'NCDX','EQ', 'MCX') order by excode"
                    Set ERec = Nothing
                    Set ERec = New ADODB.Recordset
                    ERec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                    Do While Not ERec.EOF
                        If GEXCODE = "" Then
                           GEXCODE = ERec("excode")
                        Else
                           GEXCODE = GEXCODE & "," & ERec("excode")
                        End If
                        
''''                        If ERec("excode") = "EQ" Then
''''                            LFileName = App.Path & "\" & GExCode & "\" & "cm" & CStr(Left$((Date - 1), 2)) & UCase(Left$(MonthName(Val(month((Date - 1)))), 3)) & CStr(Year((Date - 1))) & "bhav.csv.zip"
''''                        ElseIf ERec("excode") = "NCDX" Then
''''                            LFileName = App.Path & "\" & GExCode & "\" & "FO_" & CStr(Left((Date - 1), 2)) & CStr(month((Date - 1))) & Right(CStr(Year((Date - 1))), 2) & "_FINAL.CSV"
''''                        ElseIf ERec("excode") = "NSE" Then
''''                            LFileName = App.Path & "\" & GExCode & "\" & "fo" & CStr(Left$((Date - 1), 2)) & UCase(Left$(MonthName(Val(month(Date - 1))), 3)) & Year(Date - 1) & "bhav.csv"
''''                        ElseIf ERec("excode") = "MCX" Then
''''                            LFileName = App.Path & "\" & GExCode & "\" & "MCX_MS" & Vdate1 & ".csv"
''''                        End If

                        '>>> create folder if not exists
                        If Not LFileSystemObject.FolderExists(App.Path & "\" & ERec("excode")) Then
                            LFileSystemObject.CreateFolder (App.Path & "\" & ERec("excode"))
                        End If

                        If ERec("excode") = "EQ" Then
                            LFileName = App.Path & "\" & GEXCODE & "\" & "cm" & CStr(Left$((filedate), 2)) & UCase(Left$(MonthName(Val(month((filedate)))), 3)) & CStr(Year((filedate))) & "bhav.csv.zip"
                        ElseIf ERec("excode") = "NCDX" Then
                            LFileName = App.Path & "\" & GEXCODE & "\" & "FO_" & CStr(Left((filedate), 2)) & CStr(month((filedate))) & Right(CStr(Year((filedate))), 2) & "_FINAL.CSV"
                        ElseIf ERec("excode") = "NSE" Then
                            LFileName = App.Path & "\" & GEXCODE & "\" & "fo" & CStr(Left$((filedate), 2)) & UCase(Left$(MonthName(Val(month(filedate))), 3)) & Year(filedate) & "bhav.csv"
                        ElseIf ERec("excode") = "MCX" Then
                            LFileName = App.Path & "\" & GEXCODE & "\" & "MCX_MS" & Vdate1 & ".csv"
                        End If
                        If Not FileExist(LFileName) Then
                            If ERec("excode") = "MCX" Then
                                VFiledownload = DownloadFile("http://178.33.73.246/mcxbhavcopy.aspx?date=" & Vdate, App.Path & "\" & "MCX\MCX_MS" & Vdate1 & ".csv")
                            End If
                            DoEvents
                            If Not VFiledownload Then
                                VFiledownload = DownloadFile("http://178.33.73.246/mcxbhavcopy.aspx?date=" & Vdate, App.Path & "\" & "MCX\MCX_MS" & Vdate1 & ".csv")
                                DoEvents
                            End If
                                                                                
''''                            If Weekdy > 2 And Weekdy < 7 Then
''''                                frmdata.vcDTP1.Value = Date - 1
''''                                frmdata.vcDTP2.Value = Date - 1
''''                            ElseIf Weekdy = 2 Then 'if Monday -- Get date of last Friday
''''                                frmdata.vcDTP1.Value = Date - 3
''''                                frmdata.vcDTP2.Value = Date - 3
''''                            End If
                            frmdata.vcDTP1.Value = filedate
                            frmdata.vcDTP2.Value = filedate

                            If ERec("excode") = "NCDX" Then
                                frmdata.CHKNCDXExcelClosing.Value = 1
                            ElseIf ERec("excode") = "NSE" Then
                                frmdata.ChkNSEBhavCopy.Value = 1
                            ElseIf ERec("excode") = "EQ" Then
                                frmdata.ChkNSEEQClosing.Value = 1
                            ElseIf ERec("excode") = "MCX" Then
                                frmdata.Check6.Value = 1
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
                            frmdata.Label10.Caption = frmdata.Label10.Caption + " -- " & ERec("excode")
                            frmdata.okcmdclick
                            Exit Do
                        End If
                        ERec.MoveNext
                    Loop
                End If
            GETMAIN.OrdRegRPT.Visible = False
            If GOrderEntryYN = "Y" Then
                GETMAIN.OrdRegRPT.Visible = True
            End If
            '>>>  - 17 JAN 2021 -- Bhavcopy download if not exists ENDUK
            GEXECODE = 191023
            Cnn.Execute "UPDATE COMPANY SET ECODE = " & GEXECODE & " "
            
            Unload Me
            Exit Sub
        Else
            MsgBox "Verification Status:" & GlobalVClientStatus & "," & vbNewLine & vbNewLine & "Contact Sauda Support Staff to Get your Software" & vbNewLine & "Registered Reg ID: " & HardDiskNo & vbNewLine & vbNewLine & "Call : " & Label5.Caption
            Unload Me
            End
        End If
    Else
        MsgBox "Sorry Invalid User Try Again"
        txtUsername.text = vbNullString
        txtPassword.text = vbNullString
        txtUsername.SetFocus
        Exit Sub
    End If
    
    'Set MCnn = Nothing
    Unload Me
    Exit Sub
err1:
    If err.Number <> 0 Then
        MsgBox "Err2:" & err.Description
        MsgBox err.Description & " :   ", vbCritical, err.HelpFile
   End If
End Sub
Private Sub txtUsername_LostFocus()
    txtUsername.BackColor = &HFF8080
    txtUsername.ForeColor = &HFFFFFF
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
mysql = "SELECT MIN(CONDATE) AS MDATE  FROM CTR_D WHERE COMPCODE =" & GCompCode & ""
Set TRec = Nothing
Set TRec = New ADODB.Recordset
TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
If Not TRec.EOF Then
    If Not IsNull(TRec!MDate) Then
        LMinDate = TRec!MDate
        mysql = "SELECT MIN(VOU_DT) AS MDATE FROM VCHAMT  WHERE COMPCODE =" & GCompCode & ""
        Set TRec2 = Nothing
        Set TRec2 = New ADODB.Recordset
        TRec2.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec2.EOF Then
            If Not IsNull(TRec2!MDate) Then
                If LMinDate > TRec!MDate Then
                    LMinDate2 = TRec2!MDate
                Else
                    LMinDate2 = LMinDate
                End If
            End If
        End If
        mysql = "UPDATE COMPANY SET FINBEGIN='" & Format(LMinDate2, "YYYY/MM/DD") & "' WHERE COMPCODE =" & GCompCode & ""
        Cnn.Execute mysql
        
        mysql = "UPDATE SYSCOMP SET FINBEGIN='" & Format(LMinDate2, "YYYY/MM/DD") & "' WHERE COMPCODE =" & GCompCode & " AND DATABASENAME ='" & GDatabaseName & "'"
        MCnn.Execute mysql
        
        mysql = "SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND "
        mysql = mysql & " EXISTS (SELECT TOP 1 COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND PATTAN ='O' AND CONDATE >'" & Format(LMinDate, "YYYY/MM/DD") & "')"
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Please call Sauda Support Staff As there are opening Trade Beyond Open Date. "
        End If
    End If
End If
End Sub
