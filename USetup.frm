VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form USetup 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   12
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -240
         TabIndex        =   14
         Top             =   120
         Width           =   12255
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "User Setup"
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
            Height          =   495
            Left            =   360
            TabIndex        =   15
            Top             =   120
            Width           =   12135
         End
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   120
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
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8190
      Left            =   2880
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   14446
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User Setup"
      TabPicture(0)   =   "USetup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Permissions"
      TabPicture(1)   =   "USetup.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MfGrid"
      Tab(1).Control(1)=   "Label5(0)"
      Tab(1).Control(2)=   "Label5(2)"
      Tab(1).Control(3)=   "Label5(3)"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   8775
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   8175
            Begin VB.Frame Frame14 
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               ForeColor       =   &H00C0C0FF&
               Height          =   1455
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Visible         =   0   'False
               Width           =   5535
               Begin VB.TextBox TxtAdminPass 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1560
                  MaxLength       =   20
                  PasswordChar    =   "*"
                  TabIndex        =   19
                  Top             =   720
                  Width           =   2295
               End
               Begin VB.Label Label27 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Enter Master Password"
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
                  Left            =   1560
                  TabIndex        =   20
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.Shape Shape2 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H00FFC0C0&
                  Height          =   1215
                  Left            =   120
                  Top             =   120
                  Width           =   5295
               End
            End
            Begin VB.TextBox PASSWORD 
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
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   12
               PasswordChar    =   "*"
               TabIndex        =   2
               ToolTipText     =   "********"
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox USERNAME 
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
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   1320
               MaxLength       =   12
               TabIndex        =   1
               Top             =   240
               Width           =   4215
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
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
               Height          =   240
               Left            =   120
               TabIndex        =   11
               Top             =   900
               Width           =   1035
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User Name"
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
               Height          =   240
               Left            =   120
               TabIndex        =   10
               Top             =   300
               Width           =   1215
               WordWrap        =   -1  'True
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MfGrid 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   12726
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   12632256
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "1. F2  Select/Unselect row,   2. F3  Select/Unselect all for Active Group  3. F4  Select/Unselect all options"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   5
         Top             =   7620
         Width           =   8175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. F6  Select/Unselect active col. for active group"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   4
         Top             =   7920
         Width           =   3720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. F7  Select/Unselect active col."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   -70800
         TabIndex        =   3
         Top             =   7920
         Width           =   2460
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   8220
      Left            =   120
      TabIndex        =   7
      Top             =   1035
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   14499
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   16777215
      ListField       =   ""
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1390
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Authors"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   7980
      Left            =   75
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   5940
      Left            =   2880
      Top             =   1080
      Width           =   8925
   End
End
Attribute VB_Name = "USetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte
Public MUsername As String
Dim OldUserName As String
Dim CompRec As ADODB.Recordset
Dim MenuRec As ADODB.Recordset
Dim UserRec As ADODB.Recordset

Sub ADD_NEW_RECORD()
    Fb_Press = 1: OldUserName = "": USERNAME.Locked = False: PASSWORD.Locked = False:  SSTab1.Enabled = True
    Call Get_Selection(Fb_Press)
    DataList1.Locked = True
    USERNAME.SetFocus
End Sub
Sub Save_Record()
    On Error GoTo err1
    Dim flag As Boolean
    Dim TRec As ADODB.Recordset
    Dim GR As Integer
    flag = False
    If UCase(OldUserName) <> UCase(USERNAME.text) Then
        mysql = "SELECT USER_NAME FROM USERMASTER WHERE USER_NAME='" & USERNAME.text & "'"
        Set TRec = Nothing: Set TRec = New ADODB.Recordset:
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox String(5, " ") & "Duplicate User Name." & String(10, " "), vbCritical, "Error"
            USERNAME.SetFocus
            Exit Sub
        End If
        Set TRec = Nothing
    End If
    Cnn.BeginTrans
    CNNERR = True
        If Fb_Press = 2 Then
            
            mysql = "DELETE FROM USERMASTER WHERE USER_NAME='" & MUsername & "'"
            Cnn.Execute mysql
    
            mysql = "DELETE FROM USER_RIGHTS WHERE USER_NAME='" & MUsername & "'"
            Cnn.Execute mysql
        End If
        mysql = "INSERT INTO USERMASTER(USER_NAME, PASSWD, MAINUSER, SW) VALUES('" & USERNAME.text & "','" & EncryptNEW(PASSWORD.text, 13) & "',0,'" & SW & "')"
        Cnn.Execute mysql
        For GR = 1 To (MfGrid.Rows - 1)
            If MfGrid.TextMatrix(GR, 1) = "" Then
            Else
                mysql = "INSERT INTO USER_RIGHTS(SW, USER_NAME, MENUNAME, M_VISIBLE, M_ADDITION, M_MODIFICATION, M_DELETION, M_PRINT) VALUES('" & SW & "','" & USERNAME.text & "','" & MfGrid.TextMatrix(GR, 1) & "'," & IIf(MfGrid.TextMatrix(GR, 3) = Chr(254), 1, 0) & "," & IIf(MfGrid.TextMatrix(GR, 4) = Chr(254), 1, 0) & "," & IIf(MfGrid.TextMatrix(GR, 5) = Chr(254), 1, 0) & "," & IIf(MfGrid.TextMatrix(GR, 6) = Chr(254), 1, 0) & "," & IIf(MfGrid.TextMatrix(GR, 7) = Chr(254), 1, 0) & ")"
                Cnn.Execute mysql
            End If
        Next GR
        Cnn.CommitTrans
        CNNERR = False
        GETMAIN.MousePointer = 0
        Call LogIn
        Call CANCEL_RECORD
        Exit Sub
err1:
    MsgBox err.Description, vbCritical, err.HelpFile
    If CNNERR = True Then Cnn.RollbackTrans: CNNERR = False
End Sub
Sub CANCEL_RECORD()
    Dim I As Integer
    PASSWORD.text = vbNullString
    
    USERNAME.text = vbNullString
    
    SSTab1.Tab = 0: SSTab1.Enabled = False: Fb_Press = 0
    DataList1.Locked = False
    
    TxtAdminPass.text = ""
    Frame14.Visible = False
            
            
    Call Get_Selection(10)
End Sub
Sub DATA_ACCESS()
    
    Dim TRec As ADODB.Recordset
    Dim GR As Integer
    SSTab1.Enabled = True
    MUsername = DataList1.text
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT USER_NAME,PASSWD FROM USERMASTER WHERE USER_NAME='" & MUsername & "' ", Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        USERNAME.text = TRec!user_name: OldUserName = TRec!user_name
        PASSWORD.text = DecryptNEW(TRec!PASSWD, 13)
    End If
    USERNAME.Locked = False: PASSWORD.Locked = False
    
    mysql = "SELECT MENUNAME,M_VISIBLE,M_ADDITION,M_MODIFICATION,M_DELETION,M_PRINT FROM USER_RIGHTS WHERE USER_NAME='" & MUsername & "'ORDER BY MENUNAME"
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then Exit Sub
    TRec.MoveFirst
    Do While Not TRec.EOF
        For GR = 1 To (MfGrid.Rows - 1)
            MfGrid.Row = GR
            If MfGrid.TextMatrix(MfGrid.Row, 1) = TRec!MENUNAME Then Exit For
        Next GR
        If LenB(TRec!MENUNAME) > 0 Then
            MfGrid.Col = 3: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12
            If TRec!M_VISIBLE Then
                MfGrid.TextMatrix(MfGrid.Row, 3) = Chr(254)
            Else
                MfGrid.TextMatrix(MfGrid.Row, 3) = Chr(111)
            End If
            MfGrid.Col = 4: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12
            If TRec!M_ADDITION Then
                MfGrid.TextMatrix(MfGrid.Row, 4) = Chr(254)
            Else
                MfGrid.TextMatrix(MfGrid.Row, 4) = Chr(111)
            End If
            MfGrid.Col = 5: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 5) = Chr(111)
            If TRec!M_MODIFICATION Then
                MfGrid.TextMatrix(MfGrid.Row, 5) = Chr(254)
            Else
                MfGrid.TextMatrix(MfGrid.Row, 5) = Chr(111)
            End If
            MfGrid.Col = 6: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 6) = Chr(111)
            If TRec!M_DELETION Then
                MfGrid.TextMatrix(MfGrid.Row, 6) = Chr(254)
            Else
                MfGrid.TextMatrix(MfGrid.Row, 6) = Chr(111)
            End If
            MfGrid.Col = 7: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 7) = Chr(111)
            If TRec!M_PRINT Then
                MfGrid.TextMatrix(MfGrid.Row, 7) = Chr(254)
            Else
                MfGrid.TextMatrix(MfGrid.Row, 7) = Chr(111)
            End If
        End If
        TRec.MoveNext
    Loop
    MfGrid.Row = 1: MfGrid.Col = 3: DataList1.Locked = True
    USERNAME.SetFocus
    If Fb_Press = 3 Then
        If USERNAME.text = "ankan" And PASSWORD.text = "ankan" Then
            MsgBox "You are not allowed to Delete this record.", vbInformation
            Call CANCEL_RECORD
            Exit Sub
        End If
        If MsgBox("Are you sure about to DELETE ?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            mysql = "DELETE FROM USERMASTER   WHERE USER_NAME='" & MUsername & "'"
            Cnn.Execute mysql
            mysql = "DELETE FROM USER_RIGHTS WHERE USER_NAME='" & MUsername & "'"
            Cnn.Execute mysql
        End If
        Call CANCEL_RECORD
    End If

End Sub

Private Sub DataList1_Click()
    USERNAME.text = DataList1.text
End Sub
Private Sub DataList1_DblClick()
    If DataList1.Locked Then
    Else
        Call Get_Selection(2)
        Fb_Press = 2
        MUsername = DataList1.text
        
        If UCase(MUsername) = "MASTER" Then
            SSTab1.Enabled = True
            Frame14.Visible = True
            TxtAdminPass.SetFocus
        Else
            Call DATA_ACCESS
        End If
    End If
End Sub
Private Sub DataList1_KeyPress(KeyAscii As Integer)
    If DataList1.Locked Then
    Else
        If KeyAscii = 13 Then
            Call Get_Selection(2)
            Fb_Press = 2
            MUsername = DataList1.text
            
            If UCase(MUsername) = "MASTER" Then
                SSTab1.Enabled = True
                Frame14.Visible = True
                TxtAdminPass.SetFocus
            Else
                Call DATA_ACCESS
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
Dim LArray() As String
Dim CountSplit As Integer
Dim MMenuCaption As String
Dim I As Integer
On Error GoTo err1
    SSTab1.Tab = 0
    Set UserRec = Nothing: Set UserRec = New ADODB.Recordset
    mysql = "SELECT * FROM USERMASTER ORDER BY USER_NAME "
    UserRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not UserRec.EOF Then
        Set DataList1.RowSource = UserRec
        DataList1.ListField = "USER_NAME"
        DataList1.BoundColumn = "USER_NAME"
    End If
    
    Call FILL_MSFLEXGRID
    Set MenuRec = Nothing: Set MenuRec = New ADODB.Recordset
    mysql = "SELECT * FROM USER_MENU ORDER BY MCAPTION"
    MenuRec.Open mysql, Cnn, adOpenStatic, adLockOptimistic
    Do While Not MenuRec.EOF
        ' REMOVES & CHARACTER
        MfGrid.Row = MfGrid.Rows - 1: MfGrid.Rows = MfGrid.Rows + 1
        MMenuCaption = ""
        LArray = Split(GETMAIN.Controls(MenuRec!MENUNAME).Caption, "&")
        CountSplit = UBound(LArray)
        For I = 0 To CountSplit
            MMenuCaption = MMenuCaption & LArray(I)
        Next
        If MenuRec!GROUP_MENU = "T" Then
            MfGrid.TextMatrix(MfGrid.Row, 0) = "Tools"
        ElseIf MenuRec!GROUP_MENU = "S" Then
            MfGrid.TextMatrix(MfGrid.Row, 0) = "Setup"
        ElseIf MenuRec!GROUP_MENU = "R" Then
            MfGrid.TextMatrix(MfGrid.Row, 0) = "Reports"
        ElseIf MenuRec!GROUP_MENU = "Q" Then
            MfGrid.TextMatrix(MfGrid.Row, 0) = "Query"
        ElseIf MenuRec!GROUP_MENU = "E" Then
            MfGrid.TextMatrix(MfGrid.Row, 0) = "Entries"
        End If
        MfGrid.TextMatrix(MfGrid.Row, 1) = MenuRec!MENUNAME
        MfGrid.TextMatrix(MfGrid.Row, 2) = MMenuCaption
        MfGrid.Col = 3: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 3) = Chr(111)
        MfGrid.Col = 4: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 4) = Chr(111)
        MfGrid.Col = 5: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 5) = Chr(111)
        MfGrid.Col = 6: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 6) = Chr(111)
        MfGrid.Col = 7: MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontName = "WingDings": MfGrid.CellFontSize = 12: MfGrid.TextMatrix(MfGrid.Row, 7) = Chr(111)
        MenuRec.MoveNext
    Loop
    Call Get_Selection(10)
    SSTab1.Enabled = False
    Exit Sub
err1:
    If err.Number = Val(730) Then
        MenuRec.Delete
        MenuRec.MoveNext
        If Not MenuRec.EOF Then MenuRec.MoveLast
    Else
        MsgBox err.Description, vbCritical, err.HelpFile
        
    End If
End Sub
Private Sub Form_Paint()
    Call Get_Selection(Fb_Press)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Fb_Press = 0:
    Call Get_Selection(12)
End Sub

Private Sub ListView1_Validate(Cancel As Boolean)
    SSTab1.Tab = 1
End Sub
Sub FILL_MSFLEXGRID()
    MfGrid.Row = 0
    MfGrid.Col = 0: MfGrid.ColWidth(0) = TextWidth("XXXXXXXXXXXXX"): MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 0) = "Group"
    MfGrid.Col = 1: MfGrid.ColWidth(1) = TextWidth(""): MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 1) = "MenuName"
    MfGrid.Col = 2: MfGrid.ColWidth(2) = TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"): MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 2) = "MenuCaption"
    MfGrid.Col = 3: MfGrid.ColWidth(3) = TextWidth("XXXXXX"): MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 3) = "Show"
    MfGrid.Col = 4: MfGrid.ColWidth(4) = TextWidth("XXXXX"): MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 4) = "New"
    MfGrid.Col = 5: MfGrid.ColWidth(5) = TextWidth("XXXXX"): MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 5) = "Edit"
    MfGrid.Col = 6: MfGrid.ColWidth(6) = TextWidth("XXXXXXXXX"): MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 6) = "Delete"
    MfGrid.Col = 7: MfGrid.ColWidth(7) = TextWidth(""): MfGrid.CellAlignment = flexAlignCenterCenter: MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 7) = "Print"
    MfGrid.Col = 8: MfGrid.ColWidth(8) = TextWidth(""): MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 8) = ""
    MfGrid.Col = 9: MfGrid.ColWidth(9) = TextWidth(""): MfGrid.CellFontBold = True: MfGrid.TextMatrix(0, 9) = ""
    MfGrid.FixedRows = 1
End Sub

Private Sub MfGrid_Click()
    If MfGrid.Col = 3 Or MfGrid.Col = 4 Or MfGrid.Col = 5 Or MfGrid.Col = 6 Or MfGrid.Col = 7 Then
        If MfGrid.TextMatrix(MfGrid.Row, MfGrid.Col) = Chr(111) Then
            MfGrid.TextMatrix(MfGrid.Row, MfGrid.Col) = Chr(254)
        Else
            MfGrid.TextMatrix(MfGrid.Row, MfGrid.Col) = Chr(111)
        End If
    End If
End Sub
Private Sub MfGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim GCol As Integer
    Dim GR As Integer
    Dim I As Integer
    If SSTab1.Tab = 1 Then
        If KeyCode = 113 And Shift = 0 Then 'f2 to SELECT/unSELECT row
            For GCol = 3 To 7
                If MfGrid.TextMatrix(MfGrid.Row, GCol) = Chr(254) Then
                    MfGrid.TextMatrix(MfGrid.Row, GCol) = Chr(111)
                Else
                    MfGrid.TextMatrix(MfGrid.Row, GCol) = Chr(254)
                End If
            Next GCol
        ElseIf KeyCode = 114 And Shift = 0 Then 'f3
            For GR = 1 To (MfGrid.Rows - 1)
                If MfGrid.TextMatrix(MfGrid.Row, 0) = MfGrid.TextMatrix(GR, 0) Then
                    For GCol = 3 To 7
                        If MfGrid.TextMatrix(GR, GCol) = Chr(254) Then
                            MfGrid.TextMatrix(GR, GCol) = Chr(111)
                        Else
                            MfGrid.TextMatrix(GR, GCol) = Chr(254)
                        End If
                    Next GCol
                End If
            Next GR
        ElseIf KeyCode = 115 Then  'f4
            For GR = 1 To (MfGrid.Rows - 1)
                If MfGrid.TextMatrix(GR, 0) = "" Then
                    
                Else
                    For GCol = 3 To 7
                        If MfGrid.TextMatrix(GR, GCol) = Chr(254) Then
                            MfGrid.TextMatrix(GR, GCol) = Chr(111)
                        Else
                            MfGrid.TextMatrix(GR, GCol) = Chr(254)
                        End If
                    Next GCol
                End If
            Next GR
        ElseIf KeyCode = 117 Then  'f6
            For GR = 1 To (MfGrid.Rows - 1)
                If MfGrid.TextMatrix(MfGrid.Row, 0) = MfGrid.TextMatrix(GR, 0) Then
                    If MfGrid.TextMatrix(GR, MfGrid.Col) = Chr(254) Then
                        MfGrid.TextMatrix(GR, MfGrid.Col) = Chr(111)
                    Else
                        MfGrid.TextMatrix(GR, MfGrid.Col) = Chr(254)
                    End If
                End If
            Next GR
        ElseIf KeyCode = 118 Then  'f7
            For GR = 1 To (MfGrid.Rows - 1)
                If MfGrid.TextMatrix(GR, 0) = "" Then
                Else
                    If MfGrid.TextMatrix(GR, MfGrid.Col) = Chr(254) Then
                        MfGrid.TextMatrix(GR, MfGrid.Col) = Chr(111)
                    Else
                        MfGrid.TextMatrix(GR, MfGrid.Col) = Chr(254)
                    End If
                End If
            Next GR
        End If
    End If
End Sub
Private Sub MfGrid_KeyPress(KeyAscii As Integer)
    If MfGrid.Col = 3 Or MfGrid.Col = 4 Or MfGrid.Col = 5 Or MfGrid.Col = 6 Or MfGrid.Col = 7 Then
        If MfGrid.TextMatrix(MfGrid.Row, MfGrid.Col) = Chr(111) Then
            MfGrid.TextMatrix(MfGrid.Row, MfGrid.Col) = Chr(254)
        Else
            MfGrid.TextMatrix(MfGrid.Row, MfGrid.Col) = Chr(111)
        End If
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.NAME = "vcDTP1" Then
            Sendkeys "{tab}"
        End If
    End If
End Sub

Private Sub TxtAdminPass_Validate(Cancel As Boolean)
    If TxtAdminPass.text <> "" Then
        Dim MRec As ADODB.Recordset
        Set MRec = Nothing
        Set MRec = New ADODB.Recordset
        Dim TEMP As String
        TEMP = EncryptNEW(TxtAdminPass.text, 13)
        mysql = "SELECT USER_NAME FROM USERMASTER WHERE USER_NAME ='Master' and PASSWD='" & TEMP & "'"
        MRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not MRec.EOF Then
            Call DATA_ACCESS
            Frame14.Visible = False
        Else
            MsgBox "Invalid master password!!!", vbInformation
            TxtAdminPass.text = ""
            SSTab1.Enabled = True
            Frame14.Visible = True
            TxtAdminPass.SetFocus
        End If
    End If
End Sub
