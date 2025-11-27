VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MenuOpt 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   11280
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11280
   ScaleWidth      =   15330
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   4440
      TabIndex        =   37
      Top             =   9600
      Width           =   10455
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   $"menuopt.frx":0000
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
         TabIndex        =   38
         Top             =   120
         Width           =   10215
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Terms and Conditions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8520
      Width           =   10215
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
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
      Left            =   -120
      TabIndex        =   15
      Top             =   0
      Width           =   14655
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   14325
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Happy Diwali"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   1575
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   15
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   4095
      Left            =   4485
      TabIndex        =   6
      Top             =   2160
      Width           =   10245
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Data Import"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   2325
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "General Ledger"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2550
         Width           =   2325
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Account Stm Summary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1740
         Width           =   2325
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Query Trial Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3360
         Width           =   2325
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2680
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3360
         Width           =   2325
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Correcting Books"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3360
         Width           =   2325
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bill Summary Share"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1740
         Width           =   2325
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sub Brokerage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1740
         Width           =   2325
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Margin Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2550
         Width           =   2325
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Balance Sheet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2550
         Width           =   2325
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Trial Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2550
         Width           =   2325
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Data Backup"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3360
         Width           =   2325
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   2325
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Brokerage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   930
         Width           =   2325
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bill Summary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1740
         Width           =   2325
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Contract Register"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1000
         Width           =   2325
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Data Base Backup"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7800
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Standing Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   930
         Width           =   2325
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Account Statement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   930
         Width           =   2325
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Voucher Entry"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   2325
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Contract Entry"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   2325
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8070
      Left            =   240
      TabIndex        =   0
      Top             =   1185
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   14235
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
      ForeColor       =   64
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   315
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
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "menuopt.frx":00C0
      Top             =   9600
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Left            =   4560
      TabIndex        =   39
      Top             =   6360
      Width           =   6855
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   15480
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Change Company  (F10)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   270
      Left            =   11595
      TabIndex        =   21
      Top             =   6360
      Width           =   2985
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "88897-40123, 88898-40123, 88390-84261, 88390-85057"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   8160
      Width           =   10335
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
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
      Left            =   4485
      TabIndex        =   16
      Top             =   1680
      Width           =   10245
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back-Office Accounting Software For Commodity And Stock Exchanges"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   7440
      Width           =   10335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      FillColor       =   &H00800000&
      Height          =   8340
      Left            =   120
      Top             =   1080
      Width           =   4245
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Sauda Software "
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   6840
      Width           =   10335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "INDORE- M.P. 452001"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   7800
      Width           =   10335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Menu Option"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   1935
      TabIndex        =   2
      Top             =   645
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4485
      TabIndex        =   1
      Top             =   1200
      Width           =   10245
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      FillColor       =   &H00400000&
      Height          =   8460
      Left            =   4320
      Top             =   1080
      Width           =   10605
   End
End
Attribute VB_Name = "MenuOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
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
Call GETMAIN.DBkp_Click
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
    Label42.Caption = vbNullString
    If Registered Then
        Label16.Caption = "Registration No:" & GRegNo & " Client ID :" & GUniqClientId & "" & " Till Date  " & GTillDate & " "
    Else
        Label16.Caption = "UnRegistered User Registration No:" & GRegNo & "" & " Till Date  " & GTillDate & " "
    End If
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    MYSQL = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnubrok'"
    Cnn.Execute MYSQL
    MYSQL = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnusubbrok'"
    Cnn.Execute MYSQL
    MYSQL = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnubrokslab2'"
    Cnn.Execute MYSQL
    MYSQL = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnukyc'"
    Cnn.Execute MYSQL
    MYSQL = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnuslab'"
    Cnn.Execute MYSQL
    MYSQL = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnusaudalot'"
    Cnn.Execute MYSQL
    If GSysLockDt > GFinBegin Then
        Label9.Caption = " Settlement Locked  " & GSysLockDt & ""
    Else
        Label9.Caption = vbNullString
    End If
    
    MYSQL = "DELETE FROM USER_RIGHTS WHERE MENUNAME ='mnuitemgroup'"
    Cnn.Execute MYSQL
    
    MYSQL = "SELECT B.MENUNAME,M_VISIBLE FROM USERMASTER A,USER_RIGHTS B WHERE A.USER_NAME=B.USER_NAME AND A.USER_NAME='" & GUserName & "'"
    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockOptimistic
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
End Sub
Private Sub Form_Paint()

    If GETMAIN.ActiveForm.NAME = "MenuOptfrm" Then
        If ListView1.ListItems.Count > 0 Then
            ListView1.ListItems(1).Selected = True
            ListView1.SetFocus
        End If
    End If
    If SelComp_Ado.RecordCount = Val(1) Then Label5.Visible = False
    Label1.Caption = GCompanyName
    Label2.Caption = "Accounting Period " & GFinBegin & " to " & GFinEnd
End Sub

Private Sub Label5_Click()
    GETMAIN.comsel_Click
End Sub

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
