VERSION 5.00
Begin VB.Form serverfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sauda Setup"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      Begin VB.TextBox datatxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox ServerNameTxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox LogInTxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox PasswordTxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "driver={SQL Server};server="
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox DefaultDBName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "SwColDdb"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   12
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   9
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "serverfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub
Private Sub Command2_Click()
    If ServerNameTxt.Text = "" Then
        MsgBox "Please enter server name.", vbInformation: ServerNameTxt.SetFocus
    Else
        Call CreateDb
    End If
End Sub
Sub CreateDb()
    On Error GoTo Error1
    Dim sql As String
    Dim oRs As ADODB.Recordset
    Screen.MousePointer = 11
    If serverfrm.datatxt.Text = "" Then MsgBox "Please enter database name.", vbCritical: serverfrm.datatxt.SetFocus: Exit Sub
    Set CNN = Nothing: Set CNN = New ADODB.Connection
    ServerString = "driver={SQL Server};server=" & serverfrm.ServerNameTxt.Text & ";"
    If serverfrm.LogInTxt.Text = "" Then
        ServerString = ServerString & "uid=;"
    Else
        ServerString = ServerString & "uid=" & serverfrm.LogInTxt.Text & ";"
    End If
    If serverfrm.PasswordTxt.Text = "" Then
        ServerString = ServerString & "pwd=;"
    Else
        ServerString = ServerString & "pwd=" & serverfrm.PasswordTxt.Text & ";"
    End If
    ServerString = ServerString & "database=master;"
    With CNN
        .Provider = "MSDASQL"
        .CursorLocation = adUseServer
        .ConnectionString = ServerString
        .IsolationLevel = adXactIsolated
        .Mode = adModeShareExclusive
'                    .ConnectionTimeout = 5
        .Open
        .Execute "SET LOCK_TIMEOUT 0"
    End With
    CompanyDatabaseName = serverfrm.datatxt.Text
    ServerString = "driver={SQL Server};server=" & serverfrm.ServerNameTxt.Text & ";"
    If serverfrm.LogInTxt.Text = "" Then
        ServerString = ServerString & "uid=;"
    Else
        ServerString = ServerString & "uid=" & serverfrm.LogInTxt.Text & ";"
    End If
    If serverfrm.PasswordTxt.Text = "" Then
        ServerString = ServerString & "pwd=;"
    Else
        ServerString = ServerString & "pwd=" & serverfrm.PasswordTxt.Text & ";"
    End If
    CNN.Execute "CREATE DATABASE [" & CompanyDatabaseName & "]"
Error1: 'Database exists or not exists
        Err.Number = 0
        
        CNN.close: Set CNN = Nothing
        ServerString = ServerString & "database=" & CompanyDatabaseName
        Set CNN = Nothing: Set CNN = New ADODB.Connection: CNN.ConnectionString = ServerString
        CNN.Mode = adModeShareExclusive: CNN.IsolationLevel = adXactIsolated: CNN.CursorLocation = adUseClient: CNN.Open
        
        Call CreateDatabase
        
        Open App.Path & "\WINSQL32.INI" For Append As #1
            Write #1, Mid(ServerString, 1, Len(ServerString))
        Close #1
        
        GETMAIN.comsel.Visible = False: GETMAIN.sp2.Visible = False
        
        GETCOMP.Show
        
        Screen.MousePointer = 0
        MsgBox "Database successfully created.", vbInformation
        Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Sub CreateDatabase()
    On Error GoTo Error1
Error1:
    ' 1
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "Company", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE Company (COMPCODE smallint NULL , NAME nvarchar (100)  NULL , ADD1 nvarchar (200)  NULL , ADD2 nvarchar (30)  NULL ,CITY nvarchar (15) NULL , PIN nvarchar (6) NULL , STATE nvarchar (20)  NULL ,PHONEO nvarchar (50) NULL , PHONER nvarchar (50) NULL ,MOBILE nvarchar (50)  NULL , FAX nvarchar (15) NULL ,EMAIL nvarchar (50)  NULL , FINBEGIN smalldatetime NULL ,FINEND smalldatetime NULL , CCONTRACT nvarchar (50)  NULL , BROKAC int NULL ,STANDAC int NULL ,TRANAC int NULL ,SHREEAC int NULL ,TRADEAC int NULL ,NBOTBANK int NULL ,AMTCALBY real NULL , AUTOBROK bit NULL ,RPT_PATH nvarchar (100) NULL , PARTYCODETYPE nvarchar (1)  NULL ,INVOICE bit NULL , voupr bit NULL ,broktype char (1)  NULL) "
    End If
    CreateTblRs.close: Set CreateTblRs = Nothing
    '2
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "ACCT_EX ", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE ACCT_EX  (CompCode int NULL ,AC_CODE nvarchar (6)   NULL ,EXCODE nvarchar (10)  NULL , ACEXCODE nvarchar (30)  NULL) "
    End If
    '3
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "AC_GROUP ", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE AC_GROUP  (CompCode int NULL ,CODE smallint NULL ,G_CODE smallint NULL ,G_NAME nvarchar (40)  NULL , PCODE smallint NULL ,G_CAT nvarchar (18)  NULL ,TYPE nvarchar (1)  NULL ,TEMP_CAT nvarchar (18) NULL ,GPEFFECT bit NOT NULL ,KEYGROUP nvarchar (1) NULL , MAINGRP nvarchar (1) NULL ,SEQ int NULL ,TAMT float NULL) "
        
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,1,1,'FIXED ASSETS','010100000000000000',36,'A','010100000000000000',0,'P','',2,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,2,2,'INVESTMENTS','010200000000000000',36,'A','010200000000000000',0,'P','',3,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,3,3,'CURRENT ASSETS LOANS & ADVANCES','010300000000000000',36,'A','010300000000000000',0,'P','',4,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,4,4,'Misc. Expenditure Capitalized','010500000000000000',36,'A','010500000000000000',0,'P','',13,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,5,5,'CAPITAL','020100000000000000','37','L','020100000000000000',0,'P','',21,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,6,6,'RESERVES & SURPLUS','020200000000000000',37,'L','020200000000000000',0,'P','',22,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,7,7,'SECURED LOANS','020300000000000000',37,'L','020300000000000000',0,'P','',23,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,8,8,'UNSECURED LOANS','020400000000000000',37,'L','020400000000000000',0,'P','',24,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,9,9,'CURRENT LIABILITIE PROVISION','020500000000000000',37,'L','020500000000000000',0,'P','',25,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,10,10,'Cash in Hand','010301000000000000',3,'A','010301000000000000',0,'P','',5,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,11,11,'CASH AT BANK','010302000000000000',3,'A','010302000000000000',0,'P','',6,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,12,12,'SUNDRY DEBTORS','010303000000000000',3,'A','010303000000000000',0,'P','',7,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,13,13,'SUNDRY CREDITORS','020501000000000000',9,'L','020501000000000000',0,'P','',26,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,14,14,'CUSTOMER /  SUPPLIERS','010304000000000000',3,'A','010304000000000000',0,'P','',8,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,15,15,'SALES','030100000000000000',38,'I','030100000000000000',0,'P','',44,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,16,16,'PURCHASE','040100000000000000',39,'E','040100000000000000',0,'P','',50,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,17,17,'Debtors Credit Balance','020502000000000000',9,'L','020502000000000000',0,'P','',29,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,18,18,'SELLING EXPENSES','040300000000000000',39,'E','040300000000000000',0,'P','',52,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,19,19,'ADVANCE TO SUNDARIES','010305000000000000',3,'A','010305000000000000',0,'P','',9,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,20,20,'FINANCIAL EXPENSES','040400000000000000',39,'E','040400000000000000',0,'P','',53,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,21,21,'Deposits Paid','010400000000000000',36,'A','010400000000000000',0,'P','',12,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,22,22,'DEPRECIATION','040500000000000000',36,'A','040500000000000000',0,'P','',54,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,23,23,'Sundry Creditor For Expenses','020503000000000000',9,'L','020503000000000000',0,'P','',30,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,24,24,'Miscellaneous Sales','030200000000000000',38,'I','030200000000000000',0,'P','',45,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,27,27,'MISC.EXPENSES NOT WRITTEN OFF','030300000000000000',38,'I','030300000000000000',0,'P','',46,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,28,28,'MISCELLANEOUS EXPENSES','040600000000000000',39,'E','040600000000000000',0,'P','',55,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,32,32,'OPENING STOCK','040200000000000000',39,'E','040200000000000000',0,'P','',51,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,35,35,'OTHER MANUFACTURING EXP.','040700000000000000',39,'E','040700000000000000',0,'P','',56,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,36,36,'ASSETS','010000000000000000',0,'A','010000000000000000',0,'P','Y',1,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,37,37,'LIABILITIES','020000000000000000',0,'L','020000000000000000',0,'P','Y',20,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,38,38,'INCOME','030000000000000000',0,'I','030000000000000000',0,'P','Y',43,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,39,39,'EXPENCES','040000000000000000',0,'E','040000000000000000',0,'P','Y',47,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,49,40,'Transporter A/c.','020501010000000000',13,'L','020501010000000000',0,'','',27,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,50,41,'Stores Suppliers','020501020000000000',13,'L','020501020000000000',0,'','',28,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,51,3,'ADMINISTRATIVE EXPENSES','010306000000000000',39,'E','010306000000000000',0,'','',48,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,52,37,'ADVANCE TO STAFF','020600000000000000',37,'L','020600000000000000',0,'','',32,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,53,37,'MARKETING STAFF','020601000000000000',52,'L','020601000000000000',0,'','',33,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,54,37,'MARKETING STAFF IMPEREST)','020602000000000000',52,'L','020602000000000000',0,'','',34,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,55,37,'STAFF  ADVANCE  (LEFT OUT)','020603000000000000',52,'L','020603000000000000',0,'','',35,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,56,36,'RAW MATERIAL','010600000000000000',36,'A','010600000000000000',0,'','',14,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,57,36,'CAPITAL WORK IN PROGRESS','010700000000000000',36,'A','010700000000000000',0,'','',15,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,58,9,'SALARY ALLOWANCE & BENEFITE','020504000000000000',39,'L','020504000000000000',0,'','',49,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,59,39,'REPAIR & MAINTNANCE EXP.','040800000000000000',39,'E','040800000000000000',0,'','',57,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,60,36,'RAW MATERIAL CONSUMED','010800000000000000',36,'A','010800000000000000',0,'','',16,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,61,39,'STORES & SPARES CONSUMED','040900000000000000',39,'E','040900000000000000',0,'','',58,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,62,37,'OTHER LIABILITIES','020700000000000000',37,'L','020700000000000000',0,'','',37,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,63,37,'STAFF LOAN','020604000000000000',52,'L','020604000000000000',0,'','',36,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,64,37,'DEFERRED TAX','020800000000000000',37,'L','020800000000000000',0,'','',38,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,65,37,'INCOME TAX','020900000000000000',37,'L','020900000000000000',0,'','',39,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,66,37,'SALES TAX','021000000000000000',37,'L','021000000000000000',0,'','',40,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,67,37,'INTER OFFICE A/C','021000000000000000',37,'L','021000000000000000',0,'','',41,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,68,36,'LOANS & ADVANCES','010900000000000000',36,'A','010900000000000000',0,'','',17,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,69,37,'SECURITY DEPOSIT','021000000000000000',37,'L','021000000000000000',0,'','',42,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,70,36,'STORES & SPARES PARTS','011000000000000000',36,'A','011000000000000000',0,'','',19,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,71,36,'DEPRECIATION RESERVE','011000000000000000',36,'A','011000000000000000',0,'','',18,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,72,3,'OTHER ASSETS, LOANS & ADVANCES','010307000000000000',3,'A','010307000000000000',0,'','',11,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,73,39,'REBATE AND DISCOUNT','041000000000000000',39,'E','041000000000000000',0,'','',59,0)"
        CNN.Execute "INSERT INTO Ac_Group (COMPCODE,CODE,G_CODE,G_NAME,G_CAT,PCODE,TYPE,TEMP_CAT,GPEFFECT,KEYGROUP,MAINGRP,SEQ,TAMT) VALUES (1001,74,38,'OTHER INCOME & RECEIPTS','030400000000000000',38,'I','030400000000000000',0,'','',76,0)"
    End If
    '4
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "AccFmly", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE AccFmly  (CompCode int NULL,FmlyId int IDENTITY (1, 1) NOT NULL , FmlyCode nvarchar (6) NULL ,FmlyName nvarchar (50) NULL ,FmlyHead nvarchar (50) NULL) "
    End If
    '5
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "AccountD", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE AccountD (CompCode int NULL,AC_CODE nvarchar (6) NULL ,GCODE int NULL,NAME nvarchar (100) NULL,Add1 nvarchar (100) NULL,Add2 nvarchar (30) NULL,City nvarchar (30) NULL,Pin nvarchar (30) NULL,PhoneO nvarchar (30) NULL,PhoneR nvarchar (30) NULL,Fax nvarchar (30) NULL, Mobile nvarchar (30) NULL ,Email nvarchar (30) NULL,Contact nvarchar (30) NULL,FmlyId int NULL) "
    End If
    '6
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "AccountM", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE AccountM (CompCode int NULL,AC_Code nvarchar (6) NULL,Name nvarchar (100) NULL,GCode int NULL,GRPCODE int NULL ,OP_BAL float NULL , ClBal float NULL , PostDays smallint NULL ,Add1 nvarchar (30) NULL ,Add2 nvarchar (30) NULL ,City nvarchar (15) NULL ,pin nvarchar (6) NULL ,phoneR nvarchar (15) NULL,phoneO nvarchar (15) NULL ,mobile nvarchar (15) NULL ,fax nvarchar (15) NULL ,email nvarchar (30) NULL ,contact nvarchar (30) NULL ,ACTIVE bit NULL , INTDATE nvarchar (10) NULL , CREDIT float NULL , DEBIT float NULL) "
    End If
    '7
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "CTR_D", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE CTR_D (CompCode int NULL ,CONSNO int NULL ,CONNO int NULL ,PARTY nvarchar (6) NULL ,SAUDA nvarchar (15) NULL , ITEMCODE nvarchar (15) NULL ,CONDATE smalldatetime NULL ,CONTYPE nvarchar (6) NULL ,QTY float NULL , RATE float NULL ,PATTAN nvarchar (6) NULL ,BROKRATE float NULL ,TRANRATE float NULL ,STDRATE float NULL ,BILLNO int NULL ,ROWNO1 int NULL ,PERCONT nvarchar (1) NULL , INVNO int NULL ,INVDATE smalldatetime NULL ,dataimport bit NULL , ROWNO int IDENTITY (1, 1) NOT NULL , CONTIME nvarchar (50) NULL ,BrokType char (1)  NULL) "
    End If
    '8
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "CTR_M", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE CTR_M (CompCode int NULL ,CONSNO int NULL ,CONDATE smalldatetime NULL ,SAUDA nvarchar (15) NULL , ITEMCODE nvarchar (15) NULL ,CLOSERATE float NULL ,PATTAN nvarchar (1) NULL ,VOU_NO nvarchar (15) NULL ,dataimport bit NULL) "
    End If
    '9
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "CTR_R", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE CTR_R (CompCode int NULL ,CONSNO int NULL ,SAUDA nvarchar (15) NULL ,CONDATE smalldatetime NULL ,PATTAN nvarchar (1) NULL ,OPRATE float NULL , HGRATE float NULL ,LOWRATE float NULL ,CLOSERATE float NULL ,dataimport bit NULL) "
    End If
    '10
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "CodeM", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE CodeM (CompCode int NULL ,Code int NULL ,Id nvarchar (3)  NULL , Name nvarchar (30)  NULL ,Actual bit NULL ,Cal smallint NULL) "
    End If
    '11
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "EXMAST", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE EXMAST (CompCode int NULL ,EXCODE nvarchar (10) NULL , EXNAME nvarchar (100) NULL ,TRADINGACC nvarchar (6) NULL ,BANKACC nvarchar (6) NULL , ContractACC nvarchar (6) NULL , brokac nvarchar (6) NULL ,standac nvarchar (6) NULL ,tranac nvarchar (6) NULL ,shreeac nvarchar (6) NULL , roundacc nvarchar (6) NULL ,TradeDI char (1) NULL , ClosingDI char (1) NULL) "
    End If
    '12
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "ExTax", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE ExTax (COMPCODE int NULL , EXCHANGECODE nvarchar (6) NULL , FROMDT smalldatetime NULL ,TODT smalldatetime NULL ,SERVICETAX float NULL ,TURNOVERTAX float NULL ,STAMPDTAX float NULL) "
    End If
    '13
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "INV_D", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE INV_D (CompCode int NULL , BILLNO int NULL ,STDATE smalldatetime NULL , PARTY nvarchar (6) NULL ,SAUDA nvarchar (15) NULL , ITEMCODE nvarchar (15) NULL ,OPQTY float NULL ,OPRATE float NULL , CLQTY float NULL , CLRATE float NULL , DIFFAMT float NULL ,BROKRATE float NULL , BROKAMT float NULL ,STDRATE float NULL ,STDAMT float NULL , TRANRATE float NULL ,TRANAMT float NULL ,VCHNO nvarchar (15) NULL , BILLAMT float NULL ,SETNO int NULL ,INVNO int NULL ,INVDATE smalldatetime NULL) "
    End If
    '14
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "ITEMMAST", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE ITEMMAST (CompCode int NULL , ITEMCODE nvarchar (15) NULL , ITEMNAME nvarchar (30) NULL ,  ITEMRATE real NULL , LOT float NULL , EXHCODE nvarchar (30) NULL , BROKRATE float NULL , BROKTYPE nvarchar (1) NULL , STDRATE float NULL ,TRANRATE float NULL ,MARTYPE nvarchar (1) NULL , MARRATE float NULL , EXCHANGECODE nvarchar (6) NULL) "
    End If
    '15
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "PITBROK", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE PITBROK ( CompCode int NULL , AC_CODE nvarchar (6) NULL , ITEMCODE nvarchar (15) NULL , BROKTYPE nvarchar (1)  NULL , BROKRATE float NULL ,  STDRATE real NULL ,TRANRATE float NULL , PARTYTYPE nvarchar (1) NULL , BCYCLE nvarchar (1)  NULL , PERSONNELAC nvarchar (1) NULL ,uptostdt smalldatetime NULL , MARTYPE nvarchar (1) NULL ,MARRATE float NULL) "
    End If
    '16
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "SAUDAMAST", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE SAUDAMAST ( CompCode int NULL ,SAUDACODE nvarchar (15) NULL ,SAUDANAME nvarchar (30) NULL , ITEMCODE nvarchar (15) NULL , MATURITY smalldatetime NULL) "
    End If
    '17
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "TEMPBOOKS", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE TEMPBOOKS ( OP_BAL float NULL , CL_BAL float NULL , DT smalldatetime NULL ,ACC nvarchar (65) NULL ,TOTAL_R float NULL , TOTAL_P float NULL , TOTAL float NULL , AMT float NULL , BALANCE float NULL , VOU_TYPE nvarchar (10) NULL , ITEM_CODE nvarchar (15) NULL , LOC_CODE nvarchar (4) NULL , NARRATION nvarchar (80) NULL,CHQ_NO nvarchar (10) NULL ,CHQ_DT nvarchar (10) NULL ,VOU_NO nvarchar (17) NULL , docno nvarchar (11)  NULL , docdt smalldatetime NULL , doctype nvarchar (2) NULL ,docqty smallint NULL , BILL_NO nvarchar (15) NULL , BILL_DT smalldatetime NULL ,CompCode int NULL) "
    End If
    '18
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "TEMP_ACCSTT", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE TEMP_ACCSTT (CONTDATE datetime NULL , SAUDACODE nvarchar (15) NULL , SAUDANAME nvarchar (50) NULL , ITEMCODE nvarchar (15) NULL , ITEMNAME nvarchar (50) NULL , PARTYCODE nvarchar (6) NULL , PARTYNAME nvarchar (100) NULL , ADDRESS nvarchar (100) NULL , CITY nvarchar (50) NULL , BQNTY float NULL , BRATE float NULL ,SQNTY float NULL ,SRATE float NULL , CTYPE nvarchar (1) NULL , PATTAN nvarchar (1) NULL , DRAMOUNT money NULL ,CRAMOUNT money NULL , BROKERAGE float NULL , STANDING float NULL ,TRFEES float NULL , CONTDATE1 datetime NULL , SRNO int IDENTITY (1, 1) NOT NULL , INVNO int NULL , INVDATE datetime NULL , WORDAMT nvarchar (250) NULL , CALVAL float NULL , COMPCODE int NULL , rpattan nvarchar (1) NULL , ServiceTax float NULL ,StampDty float NULL ,TurnOverTax float NULL ,BrokRate float NULL , BrokType char (1) NULL, CONNO int NULL , CONTIME nvarchar (50) NULL,MarRate float NULL ,MarAmt float NULL, MarType char (1) NULL,TranRate float NULL, TranType char (1) NULL ) "
    End If
    '19
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "TxDb", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        MYSQL = "CREATE TABLE TxDb (SerialNo int IDENTITY (1, 1) NOT NULL ,F1 nvarchar (50) NULL , f2 nvarchar (50) NULL , f3 nvarchar (50) NULL , f4 nvarchar (50) NULL , f5 nvarchar (50) NULL , f6 nvarchar (50) NULL ,f7 nvarchar (50) NULL , f8 nvarchar (50) NULL , f9 nvarchar (50) NULL , f10 nvarchar (50) NULL ,f11 nvarchar (50) NULL ,f12 nvarchar (50) NULL , f13 nvarchar (50) NULL , f14 nvarchar (50) NULL , f15 nvarchar (50) NULL ,f16 nvarchar (50) NULL ,f17 nvarchar (50) NULL , f18 nvarchar (50) NULL , f19 nvarchar (50) NULL , f20 nvarchar (50) NULL , f21 nvarchar (50)  NULL , f22 nvarchar (50) NULL , f23 nvarchar (50) NULL , f24 nvarchar (50) NULL , f25 nvarchar (50) NULL , f26 nvarchar (50)  NULL , f27 nvarchar (50) NULL , f28 nvarchar (50) NULL , f29 nvarchar (50) NULL , f30 nvarchar (50) NULL ,f31 nvarchar (50) NULL ,f32 nvarchar (50)  NULL ,f33 nvarchar (50) NULL , f34 nvarchar (50)  NULL ,f35 nvarchar (50)  NULL ,"
        MYSQL = MYSQL & "f36 nvarchar (50) NULL ,f37 nvarchar (50) NULL , f38 nvarchar (50) NULL , f39 nvarchar (50) NULL ,f40 nvarchar (50) NULL , f41 nvarchar (50) NULL , f42 nvarchar (50) NULL , f43 nvarchar (50) NULL , F44 nvarchar (50)NULL)"
        CNN.Execute MYSQL
    End If
    '20
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "USERMASTER", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE  USERMASTER ( USER_NAME nvarchar (12) NULL , PASSWD nvarchar (12) NULL , MAINUSER smallint NULL , SW nvarchar (5) NULL) "
    End If
    '21
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "USER_COMPANY", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE  USER_COMPANY ( USER_NAME nvarchar (12) NULL ,COMPCODE smallint NULL , SW nvarchar (5) NULL) "
    End If
    '22
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "USER_MENU", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE  USER_MENU ( GROUP_MENU nvarchar (100) NULL , MENUNAME nvarchar (100) NULL , SW nvarchar (5)  NULL) "
    End If
    '23
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "USER_RIGHTS", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE  USER_RIGHTS ( USER_NAME nvarchar (12) NULL , SW nvarchar (5) NULL , MENUNAME nvarchar (100) NULL , M_VISIBLE tinyint NULL , M_ADDITION tinyint NULL , M_MODIFICATION tinyint NULL , M_DELETION tinyint NULL ,M_PRINT tinyint NULL) "
    End If
    '24
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "VCHAMT", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE  VCHAMT ( CompCode int NULL , VOU_NO nvarchar (15) NULL , VOU_TYPE nvarchar (2) NULL , VOU_DT smalldatetime NULL , DR_CR nvarchar (1) NULL , AC_CODE nvarchar (6) NULL , AMOUNT float NULL ,CL_BAL float NULL , CHEQUE_NO nvarchar (7) NULL , CHEQUE_DT nvarchar (12) NULL ,NARR_CD float NULL , NARRATION nvarchar (100) NULL , BRANCH nvarchar (30) NULL , BANK_NAME nvarchar (30) NULL , ADJAMT float NULL , vouid int IDENTITY (1, 1) NOT NULL) "
    End If
    '25
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "VOUCHER", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE  VOUCHER ( CompCode int NULL , VOU_NO nvarchar (15) NULL , VOU_DT smalldatetime NULL , VOU_TYPE nvarchar (2) NULL , VOU_PR nvarchar (10) NULL , BILLNO nvarchar (15) NULL , BILLDT smalldatetime NULL , USER_NAME nvarchar (30) NULL , USER_DATE smalldatetime NULL , USER_TIME nvarchar (15) NULL , User_ACTION nvarchar (6) NULL ,VCHNO int NULL , sauda nvarchar (15) NULL , INVNO int NULL , INVDATE smalldatetime NULL ,    narration nvarchar (100) NULL) "
    End If
    '26
    
    Set CreateTblRs = CNN.OpenSchema(adSchemaTables, _
        Array(Empty, Empty, "settle", "TABLE"))
    If Not CreateTblRs.EOF Then
    Else
        CNN.Execute "CREATE TABLE  settle ( CompCode int NULL , setno int NULL , setdate smalldatetime NULL) "
    End If
    
    Call CreateSProc
    
End Sub
Sub CreateSProc()
'   Stored Procedure -- BkpDev
    MYSQL = "CREATE PROCEDURE BkpDev  @DevName  nvarchar (10),@DevPath nvarchar (500) AS "
    MYSQL = MYSQL & "EXEC sp_addumpdevice 'disk', @DevName,@DevPath "
    CNN.Execute MYSQL
'   Stored Procedure -- BkpDB
    MYSQL = "CREATE PROCEDURE BkpDB @DbName nvarchar (25),@DevName nvarchar (25) AS "
    MYSQL = MYSQL & "BACKUP DATABASE @DbName TO @DevName with INIT"
    CNN.Execute MYSQL
'   Stored Procedure -- ReIndexing
    MYSQL = "CREATE PROCEDURE ReIndexTable AS "
    MYSQL = MYSQL & "DBCC DBREINDEX(AC_GROUP) "
    MYSQL = MYSQL & "DBCC DBREINDEX(VOUCHER) "
    MYSQL = MYSQL & "DBCC DBREINDEX(VCHAMT) "
    MYSQL = MYSQL & "DBCC DBREINDEX(ACCOUNTM) "
    MYSQL = MYSQL & "DBCC DBREINDEX(ACCOUNTD) "
    MYSQL = MYSQL & "DBCC DBREINDEX(ACCT_EX) "
    MYSQL = MYSQL & "DBCC DBREINDEX(CTR_M) "
    MYSQL = MYSQL & "DBCC DBREINDEX(CTR_D) "
    MYSQL = MYSQL & "DBCC DBREINDEX(CTR_R) "
    MYSQL = MYSQL & "DBCC DBREINDEX(EXMAST) "
    MYSQL = MYSQL & "DBCC DBREINDEX(EXTAX) "
    MYSQL = MYSQL & "DBCC DBREINDEX(INV_D) "
    MYSQL = MYSQL & "DBCC DBREINDEX(ITEMMAST) "
    MYSQL = MYSQL & "DBCC DBREINDEX(PITBROK) "
    MYSQL = MYSQL & "DBCC DBREINDEX(SAUDAMAST) "
    MYSQL = MYSQL & "DBCC DBREINDEX(SETTLE) "
    CNN.Execute MYSQL
    Call CreateIndex
End Sub
Sub CreateIndex()
    MYSQL = MYSQL & "CREATE INDEX Index27 ON Ac_Group (COMPCODE) "
End Sub
