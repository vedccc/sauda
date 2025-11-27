VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form MenuOptfrm 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   15600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   15600
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
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
      Height          =   8250
      Left            =   12240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "MenuOptfrm.frx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download Latest Sauda Updates"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7920
      Width           =   8085
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contract Entry"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1815
      Width           =   2565
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14760
      Top             =   720
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4950
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8731
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Company (F10)"
      Height          =   720
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7080
      Width           =   2685
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Account Statement"
      Height          =   720
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3330
      Width           =   2565
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Account"
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2400
      Width           =   2445
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Backup"
      Height          =   720
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7080
      Width           =   2565
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balance Sheet"
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6120
      Width           =   2445
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bill Summary Share"
      Height          =   720
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4260
      Width           =   2685
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Voucher Entry"
      Height          =   720
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   2685
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Standing Report"
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3330
      Width           =   2445
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trial Balance"
      Height          =   720
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5190
      Width           =   2685
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sub Brokerage"
      Height          =   720
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4260
      Width           =   2565
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Correcting Books"
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   2445
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contract Entry"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1440
      Width           =   2565
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contract Register"
      Height          =   720
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Width           =   2685
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bill Summary"
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4260
      Width           =   2445
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Margin Report"
      Height          =   720
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5190
      Width           =   2565
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company"
      Height          =   720
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   2685
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Brokerage"
      Height          =   720
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   2565
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Query Trial Balance"
      Height          =   720
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   2565
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Acc. Stm Summary"
      Height          =   720
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3330
      Width           =   2685
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General Ledger"
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5190
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Import"
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   2445
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      Caption         =   "Terms and Conditions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   9120
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   6480
      TabIndex        =   30
      Top             =   9360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   $"MenuOptfrm.frx":0300
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   10215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4560
      Top             =   120
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Live M2M Available"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   36
      Top             =   8760
      Width           =   8055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   1275
      Width           =   5895
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   6000
      Y1              =   6660
      Y2              =   6660
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   6000
      Y1              =   1660
      Y2              =   1660
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   6360
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   7920
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sauda Software "
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   5895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back-Office Accounting Software For Commodity     And Stock Exchanges"
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
      Height          =   600
      Left            =   120
      TabIndex        =   7
      Top             =   7320
      Width           =   5895
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "  88897-40123, 88898-40123, 88390-84261, 88390-85057"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   8280
      Width           =   5895
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1000
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   14595
      Left            =   -120
      Picture         =   "MenuOptfrm.frx":03C0
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   14655
   End
End
Attribute VB_Name = "MenuOptfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flagload As Boolean

Private Sub Command1_Click()
    FlagLoggedIn = False
    Call GETMAIN.mnudata_Click
End Sub
Private Sub Command10_Click()
    Call GETMAIN.mnubillsmry_Click
End Sub
Private Sub Command11_Click()
    Call GETMAIN.mnuexbrok_Click
End Sub
Private Sub Command12_Click()
If Text1.Visible = True Then
    Text1.Visible = False
Else
    Text1.Visible = True
End If
End Sub

Private Sub Command13_Click()
    Call GETMAIN.DBkp_Click
End Sub

Private Sub Command14_Click()
    Call GETMAIN.mnuaccsmry_Click
End Sub

Private Sub Command15_Click()
    Call GETMAIN.udtb_Click
End Sub

Private Sub Command16_Click()
    Call GETMAIN.balancesht_Click
End Sub

Private Sub Command17_Click()
Call GETMAIN.Marsry_Click
End Sub

Private Sub Command18_Click()
Call GETMAIN.mnuexsbrok_Click
End Sub

Private Sub Command19_Click()
Call GETMAIN.mnudass_Click
End Sub

Private Sub Command2_Click()
    Call GETMAIN.CONTRACTENTRY_Click
End Sub

Private Sub Command20_Click()
Call GETMAIN.rwb_Click
End Sub

Private Sub Command21_Click()
Call GETMAIN.COMPSETUP_Click
End Sub

Private Sub Command22_Click()
Call GETMAIN.mnuQTB_Click
End Sub

Private Sub Command23_Click()
    Call GETMAIN.CONTRACTENTRY7_Click
End Sub

Private Sub Command24_Click()
    Dim vreturn As String
    vreturn = ShellExecute(Me.hwnd, "open", "http://103.186.184.40/websauda/download/sauda.zip", "", "", 4)
End Sub
Private Sub Command3_Click()
    Call GETMAIN.VCHENT_Click
End Sub
Private Sub Command4_Click()
    Call GETMAIN.MENUACCSTT_Click
End Sub
Private Sub Command5_Click()
    Call GETMAIN.genled_Click
End Sub
Private Sub Command6_Click()
    Call GETMAIN.SAUDAWSSTND_Click
End Sub
Private Sub Command7_Click()
    'Call GETMAIN.DBkp_Click
    GETMAIN.comsel_Click
    Unload Me
End Sub
Private Sub Command8_Click()
Call GETMAIN.AccountHead_Click
End Sub
Private Sub Command9_Click()
    Call GETMAIN.CONTRACTREG_Click
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then GETMAIN.comsel_Click
End Sub
Private Sub Form_Load()
Dim ListIt As ListItem
Dim TRec As ADODB.Recordset
Dim MMenuCaption As String
Dim LStrArray() As String
Dim I As Integer
Dim CountSplit As Integer
    FlagLoggedIn = False
'    Label42.Caption = vbNullString
    If Registered Then
        Label16.Caption = "Registration No:" & GRegNo & ", Client ID :" & GUniqClientId '>>>& "" & vbNewLine & " Till Date  " & GTillDate & " "
        Label4.Caption = "Till Date  " & GTillDate & " "
    Else
        Label16.Caption = "UnRegistered User Registration No:" & GRegNo '>>>& "" & " Till Date  " & GTillDate & " "
        Label4.Caption = "Till Date  " & GTillDate & " "
    End If
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    mysql = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnubrok'"
    Cnn.Execute mysql
    mysql = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnusubbrok'"
    Cnn.Execute mysql
    mysql = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnubrokslab2'"
    Cnn.Execute mysql
    mysql = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnukyc'"
    Cnn.Execute mysql
    mysql = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnuslab'"
    Cnn.Execute mysql
    mysql = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnusaudalot'"
    Cnn.Execute mysql
    If GSysLockDt > GFinBegin Then
        Label9.Visible = True
        Label9.Caption = " Settlement Locked  " & GSysLockDt & ""
    Else
        Label9.Caption = vbNullString
    End If
    
    mysql = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnuitemgroup'"
    Cnn.Execute mysql
    
    mysql = "SELECT B.MENUNAME,M_VISIBLE FROM USERMASTER A,USER_RIGHTS B WHERE A.USER_NAME=B.USER_NAME AND A.USER_NAME='" & GUserName & "'"
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockOptimistic
    Do While Not TRec.EOF
        If TRec!M_VISIBLE = 1 Then
            MMenuCaption = vbNullString
            LStrArray = Split(GETMAIN.Controls(TRec!MENUNAME).Caption, "&")
            CountSplit = UBound(LStrArray)
            For I = 0 To CountSplit
                MMenuCaption = MMenuCaption & LStrArray(I)
            Next
            Set ListIt = ListView1.ListItems.Add(, , MMenuCaption)
            ListIt.SubItems(1) = TRec!MENUNAME
        Else
            Select Case TRec!MENUNAME
            Case "mnudata" 'Data Import
                Command1.Enabled = False
                Command1.Visible = False
            Case "CONTRACTENTRY" ' Contract Entry
                Command2.Enabled = False
                Command2.Visible = False
            Case "VCHENT" ' Voucher Entry
                Command3.Enabled = False
                Command3.Visible = False
            Case "mnuexbrok" ' Brokerage            Setup
                Command11.Enabled = False
                Command11.Visible = False
            Case "ACCOUNTHEAD" ' Account Setup
                Command8.Enabled = False
                Command8.Visible = False
            Case "CONTRACTREG" 'contract Register
                Command9.Enabled = False
                Command9.Visible = False
            Case "SAUDAWSSTND" 'Standing
                Command6.Enabled = False
                Command6.Visible = False
            Case "MENUACCSTT" 'Account Statement
                Command4.Enabled = False
                Command4.Visible = False
            Case "mnuaccsmry" 'Account Statement Summary
                Command14.Enabled = False
                Command14.Visible = False
            Case "mnubillsmry" 'Bill Summary "
                Command10.Enabled = False
                Command10.Visible = False
            
            
            Case "mnuexsbrok" ' sub brokerage
                Command18.Enabled = False
                Command18.Visible = False
            Case "mnudass" ' billsummary share
                Command19.Enabled = False
                Command19.Visible = False
            Case "mnuQTB" ' query on trial balance
                Command22.Enabled = False
                Command22.Visible = False
            Case "rwb" ' correcting books
                Command20.Enabled = False
                Command20.Visible = False
            Case "COMPSETUP" ' company
                Command21.Enabled = False
                Command21.Visible = False
                
                
            
            Case "genled" ' General Ledger
                Command5.Enabled = False
                Command5.Visible = False
            Case "Marsry"
                Command17.Enabled = False
                Command17.Visible = False
            Case "udtb" ' trail Balance
                Command15.Enabled = False
                Command18.Visible = False
            Case "balancesht"
                Command16.Enabled = False
                Command16.Visible = False
            Case "DBkp" ' Data Backup
                Command13.Enabled = False
                Command13.Visible = False
            End Select
        End If
        TRec.MoveNext
    Loop
    If GFlag_Fin = True Then Call FIN_UPDATE
    
    

    
    flagload = True
End Sub
Private Sub Form_Paint()
'    If GETMAIN.ActiveForm.NAME = "MenuOptfrm" Then
'        If ListView1.ListItems.Count > 0 Then
'            ListView1.ListItems(1).Selected = True
'            'If ListView1.Visible Then
'                ListView1.SetFocus
'            'End If
'        End If
'    End If
    If SelComp_Ado.RecordCount = Val(1) Then Command7.Value = False  'Label5.Visible = False
    Label1.Caption = GCompanyName
    Label2.Caption = "Accounting Period " & GFinBegin & " to " & GFinEnd


    'DoEvents
End Sub
Private Sub Form_Resize()
    If flagload Then
        Image1.Height = Me.Height + 500
        Image1.Width = Me.Width
        flagload = False
            
        Dim tempGCtrType As String
        tempGCtrType = ";" + GCtrType + ";"
        
        Command2.Visible = False
        Command23.Visible = False
            
        If (InStr(1, tempGCtrType, ";1;") > 0) Or (InStr(1, tempGCtrType, ";2;") > 0) Or (InStr(1, tempGCtrType, ";3;") > 0) Or (InStr(1, tempGCtrType, ";4;") > 0) Or (InStr(1, tempGCtrType, ";5;") > 0) Or (InStr(1, tempGCtrType, ";6;") > 0) Then
            Command2.Visible = True
            Command2.Height = 720
        End If
        If InStr(1, tempGCtrType, ";7;") > 0 Then
            If Command2.Visible Then
                Command2.Height = 360
                Command23.Visible = True
                Command23.Height = 360
                If Command2.Visible Then
                    Command23.Top = 1815
                Else
                    Command23.Top = 1440
                End If
                Command23.Caption = "Contract Entry (7)"
            Else
                Command23.Visible = True
                Command23.Height = 720
                If Command2.Visible Then
                    Command23.Top = 1815
                Else
                    Command23.Top = 1440
                End If
                Command23.Caption = "Contract Entry (7)"
            End If
        End If
        If (Command2.Visible) And (Not Command23.Visible) Then
            Command2.FontSize = 12
        ElseIf (Not Command2.Visible) And (Command23.Visible) Then
            Command23.FontSize = 12
        End If
        

    End If
End Sub

'Private Sub Label5_Click()
'    GETMAIN.comsel_Click
'End Sub
Private Sub ListView1_Click()
    Call SelectMenuOption
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call SelectMenuOption
End Sub
Sub SelectMenuOption()
        Select Case ListView1.SelectedItem.SubItems(1)
        Case "mnudass"
            Call GETMAIN.mnudass_Click
        Case "ACCOUNTHEAD"
            Call GETMAIN.AccountHead_Click
        Case "ACTYPE"
            Call GETMAIN.ACTYPE_Click
        Case "mnutrdconfirm"
            Call GETMAIN.mnutrdconfirm_Click
        Case "mnusubshare"
            Call GETMAIN.mnusubshare_Click
        Case "mnudtst"
            Call GETMAIN.mnudtst_Click
        Case "balancesht"
            Call GETMAIN.balancesht_Click
        Case "mnutrdreg"
            Call GETMAIN.mnutrdreg_Click
        Case "mnutrurnover"
            Call GETMAIN.mnutrurnover_Click
        Case "bankbook_f1"
            Call GETMAIN.bankbook_f1_Click
        Case "bankbook_f2"
            Call GETMAIN.bankbook_f2_Click
        Case "mnuNewStm"
            Call GETMAIN.mnuNewStm_Click
        Case "PackUpData"
            Call GETMAIN.PackUpData_Click
        Case "mnubrshare"
            Call GETMAIN.mnubrshare_Click
        Case "QonBlst"
             Call GETMAIN.QonBlst_Click
        Case "bcrpt"
            Call GETMAIN.bcrpt_Click
        Case "BrokLst"
            Call GETMAIN.BrokLst_Click
        Case "bwtbal"
            Call GETMAIN.bwtbal_Click
        Case "cashbook_f1"
            Call GETMAIN.cashbook_f1_Click
        Case "cashbook_f2"
            Call GETMAIN.cashbook_f2_Click
        Case "mnubrbrok"
            Call GETMAIN.mnubrbrok_Click
        Case "mnucnote"
            Call GETMAIN.mnucnote_Click
        Case "mnudaily"
            Call GETMAIN.mnudaily_Click
        Case "DBkp"
            Call GETMAIN.DBkp_Click
        Case "mnutrialdt"
            Call GETMAIN.mnutrialdt_Click
        Case "cbcrpt"
            Call GETMAIN.cbcrpt_Click
        Case "ccrpt"
            Call GETMAIN.ccrpt_Click
        Case "chqreg"
            Call GETMAIN.chqreg_Click
        Case "CLOSERATE"
            Call GETMAIN.CLOSERATE_Click
        Case "COMPSETUP"
            Call GETMAIN.COMPSETUP_Click
        Case "CONTRACTENTRY"
            Call GETMAIN.CONTRACTENTRY_Click
        Case "CONTRACTREG"
            Call GETMAIN.CONTRACTREG_Click
        Case "DATEWSCONTLIST"
            Call GETMAIN.DATEWSCONTLIST_Click
        Case "ftptb"
            Call GETMAIN.ftptb_Click
        Case "Exstp"
            Call GETMAIN.Exstp_Click
        Case "FmlyStup"
            Call GETMAIN.FmlyStup_Click
        Case "genled"
            Call GETMAIN.genled_Click
        Case "GenQry"
            Call GETMAIN.GenQry_Click
        Case "mnuoutstanding"
            Call GETMAIN.mnuoutstanding_Click
        Case "ITEMSETUP"
            Call GETMAIN.ITEMSETUP_Click
        Case "loginoff"
            Call GETMAIN.loginoff_Click
        Case "Marsry"
            Call GETMAIN.Marsry_Click
        Case "MENUACCSTT"
            Call GETMAIN.MENUACCSTT_Click
        Case "MenuInvPrint"
            Call GETMAIN.MenuInvPrint_Click
        Case "mnuaccsmry"
            Call GETMAIN.mnuaccsmry_Click
        Case "mnubillsmry"
            Call GETMAIN.mnubillsmry_Click
        Case "mnublist"
            Call GETMAIN.mnublist_Click
        Case "mnudata"
            Call GETMAIN.mnudata_Click
        Case "mnuINVLIST"
            Call GETMAIN.mnuINVLIST_Click
        Case "MNUInvWsLedg"
            Call GETMAIN.MNUInvWsLedg_Click
        Case "MNUInvWsLedg"
            Call GETMAIN.MNUInvWsLedg_Click
        Case "mnuQTB"
            Call GETMAIN.mnuQTB_Click
        Case "opntb"
            Call GETMAIN.opntb_Click
        Case "PANDLMENU"
            Call GETMAIN.PANDLMENU_Click
        Case "RPTBROKSMRY"
            Call GETMAIN.RPTBROKSMRY_Click
        Case "RtLst"
            Call GETMAIN.RtLst_Click
        Case "rwb"
            Call GETMAIN.rwb_Click
        Case "SAUDAMAST"
            Call GETMAIN.SAUDAMAST_Click
        Case "SAUDAWSSTND"
            Call GETMAIN.SAUDAWSSTND_Click
        Case "SETMASTER"
            Call GETMAIN.SETMASTER_Click
        Case "swmrt"
            Call GETMAIN.swmrt_Click
        Case "udtb"
            Call GETMAIN.udtb_Click
        Case "UrSetup"
            Call GETMAIN.UrSetup_Click
        Case "VCHENT"
            Call GETMAIN.VCHENT_Click
        Case "voulist"
            Call GETMAIN.voulist_Click
        Case "YrUpdate"
            Call GETMAIN.YrUpdate_Click
        Case "Reindex"
            Call GETMAIN.Reindex_Click
        End Select
End Sub

Private Sub Timer1_Timer()
'>>> 24 FEB. 2021 11:42 PM
If DateDiff("D", Now, GTillDate) < 7 Then
    DoEvents
    Label4.BackColor = &HFF&
    Label4.ForeColor = &HFFFFFF
    Label4.FontSize = 16
    If Registered Then
        If InStr(1, Label4.Caption, "Till Date") > 0 Then
        'If InStr(1, Label16.Caption, "Till Date") > 0 Then
            'Label16.Caption = "Registration No:" & GRegNo & ", Client ID :" & GUniqClientId & ""
            Label4.Caption = ""
        Else
            'Label16.Caption = "Registration No:" & GRegNo & ", Client ID :" & GUniqClientId '>>>& "" & vbNewLine & " Till Date  " & GTillDate & " "
            Label4.Caption = " Till Date  " & GTillDate & " "
        End If
    Else
        'If InStr(1, Label16.Caption, "Till Date") > 0 Then
        If InStr(1, Label4.Caption, "Till Date") > 0 Then
            'Label16.Caption = "UnRegistered User Registration No:" & GRegNo & ""
            Label4.Caption = " "
        Else
            'Label16.Caption = "UnRegistered User Registration No:" & GRegNo '>>>& "" & vbNewLine & " Till Date  " & GTillDate & " "
            Label4.Caption = "Till Date  " & GTillDate & " "
        End If
    End If
End If
DoEvents
If Label5.Caption = "" Then
    Label5.Caption = "Live M2M Available"
Else
    Label5.Caption = ""
End If
DoEvents

End Sub
