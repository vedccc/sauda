VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form SINGLECONTRACT 
   BackColor       =   &H80000000&
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00404000&
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
      TabIndex        =   31
      Top             =   0
      Width           =   1815
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
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
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   615
      Left            =   240
      TabIndex        =   25
      Top             =   9120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SINGLECONTRACT.frx":0000
         Left            =   0
         List            =   "SINGLECONTRACT.frx":000A
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Text2"
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4313
      TabIndex        =   19
      Top             =   120
      Width           =   3255
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         Caption         =   "Contract Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   -720
      Width           =   9975
      Begin VB.Label LblExit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close (F6)"
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
         Left            =   8640
         TabIndex        =   14
         Top             =   240
         Width           =   1080
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
         Left            =   5094
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1020
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
         Left            =   6777
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
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
         Left            =   3261
         TabIndex        =   11
         Top             =   240
         Width           =   1170
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
         Left            =   1713
         TabIndex        =   10
         Top             =   240
         Width           =   885
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SINGLECONTRACT.frx":0019
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SINGLECONTRACT.frx":046B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5520
      Top             =   9240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7485
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   13980
      Begin VB.Frame Frame7 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   13575
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            ItemData        =   "SINGLECONTRACT.frx":08BD
            Left            =   10005
            List            =   "SINGLECONTRACT.frx":08C7
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   120
            Width           =   1335
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   360
            Left            =   720
            TabIndex        =   35
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   37860.8625462963
         End
         Begin MSDataListLib.DataCombo ITEMCMB 
            Bindings        =   "SINGLECONTRACT.frx":08DE
            Height          =   360
            Left            =   3645
            TabIndex        =   36
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ForeColor       =   64
            Text            =   "DataCombo2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo Saudacmb 
            Bindings        =   "SINGLECONTRACT.frx":08E9
            Height          =   360
            Left            =   6765
            TabIndex        =   37
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ForeColor       =   64
            Text            =   "DataCombo2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   180
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   2400
            TabIndex        =   40
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   18
            Left            =   9360
            TabIndex        =   39
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   5880
            TabIndex        =   38
            Top             =   180
            Width           =   690
         End
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Text10"
         Top             =   6802
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Text9"
         Top             =   6802
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Text8"
         Top             =   6802
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Text7"
         Top             =   6802
         Width           =   975
      End
      Begin VB.CommandButton cmdImportFromExcel 
         Caption         =   "..."
         Height          =   285
         Left            =   -360
         TabIndex        =   15
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   6802
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   6802
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   165
         TabIndex        =   4
         Top             =   960
         Width           =   11415
      End
      Begin MSDataListLib.DataCombo Saudacombo 
         Height          =   360
         Left            =   6480
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   64
         Text            =   "SAUDACOMBO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   360
         Left            =   480
         TabIndex        =   0
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   64
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5580
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   13770
         _ExtentX        =   24289
         _ExtentY        =   9843
         _Version        =   393216
         AllowArrows     =   -1  'True
         ForeColor       =   128
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "buysell"
            Caption         =   "B/S"
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
            DataField       =   "Code"
            Caption         =   "Code"
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
         BeginProperty Column02 
            DataField       =   "name"
            Caption         =   "Party Name"
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
         BeginProperty Column03 
            DataField       =   "QNTY"
            Caption         =   "Qnty."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "RATE"
            Caption         =   "Rate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "SAUDACODE"
            Caption         =   "Sauda Code"
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
         BeginProperty Column06 
            DataField       =   "SaudaName"
            Caption         =   "Sauda Name"
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
         BeginProperty Column07 
            DataField       =   "Concode"
            Caption         =   "Con. Code"
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
         BeginProperty Column08 
            DataField       =   "ConName"
            Caption         =   "Con. Name"
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
         BeginProperty Column09 
            DataField       =   "rate1"
            Caption         =   "Con. Rate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "SRNO"
            Caption         =   "Trade No."
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
         BeginProperty Column11 
            DataField       =   "trdtime"
            Caption         =   "Trade Time"
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
         BeginProperty Column12 
            DataField       =   "CONTIME"
            Caption         =   "CONTIME"
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
         BeginProperty Column13 
            DataField       =   "LCLCODE"
            Caption         =   "LCLCODE"
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
         BeginProperty Column14 
            DataField       =   "LInvNo"
            Caption         =   "LInvNo"
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
         BeginProperty Column15 
            DataField       =   "RInvNo"
            Caption         =   "RInvNo"
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
         BeginProperty Column16 
            DataField       =   "RCLCODE"
            Caption         =   "RCLCODE"
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
         BeginProperty Column17 
            DataField       =   "DImport"
            Caption         =   "DImport"
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
         BeginProperty Column18 
            DataField       =   "UserId"
            Caption         =   "UserId"
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   1874.835
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column17 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1019.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "Sell"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6480
         TabIndex        =   24
         Top             =   6840
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "Buy"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   6840
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "Totals"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   6840
         Width           =   630
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1605
         TabIndex        =   21
         Top             =   5325
         Width           =   45
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   0
         Left            =   720
         Picture         =   "SINGLECONTRACT.frx":08F4
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F10 to open new party"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   11520
         TabIndex        =   7
         Top             =   720
         Width           =   2115
      End
      Begin VB.Image Image1 
         Height          =   195
         Index           =   1
         Left            =   1080
         Picture         =   "SINGLECONTRACT.frx":0BFE
         Stretch         =   -1  'True
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   12
      Height          =   7740
      Left            =   75
      Top             =   1080
      Width           =   14325
   End
End
Attribute VB_Name = "SINGLECONTRACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim LEXCODE  As String
Dim Lparty As String
Dim LConNo As Long
Dim LUserId As String
Dim LCONTRACTACC As String
Dim LConSno As Long
Dim LDataImport As Byte
Dim OldDate As Date
Dim GCALVAL As Long
Dim FLOWDIR As Byte
Dim VCHNO As String
Dim LSaudaCode As String
Dim LITEMCODE As String
Dim GRIDPOS As Byte
Dim GRIDREC As ADODB.Recordset
Dim CONFLAGE As Boolean
Public fb_press As Byte
Dim REC As ADODB.Recordset
Dim ITEMREC As ADODB.Recordset
Dim SaudaRec As ADODB.Recordset
Dim TEMPORARY As ADODB.Recordset
Dim RECEX As ADODB.Recordset
Dim RECGRID As ADODB.Recordset
Dim TempParty As ADODB.Recordset
Dim TempSauda As ADODB.Recordset
Dim TempItem As ADODB.Recordset
Dim REC_SAUDA As ADODB.Recordset
Dim REC_ACCOUNT As ADODB.Recordset
Dim REC_CloRate As ADODB.Recordset
Dim REC_CTRM As ADODB.Recordset
Sub SaudaList()
If ITEMCMB.BoundText <> "" Then
    MYSQL = "SELECT SAUDACODE ,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & MC_CODE & " AND ITEMCODE='" & ITEMCMB.BoundText & "' AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY SAUDANAME"
Else
    MYSQL = "SELECT SAUDACODE ,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & MC_CODE & "  AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY SAUDANAME"
End If
    Set SaudaRec = Nothing: Set SaudaRec = New ADODB.Recordset
    SaudaRec.Open MYSQL, cnn
    If Not SaudaRec.EOF Then
        Set Saudacmb.RowSource = SaudaRec
        Saudacmb.ListField = "SAUDANAME"
        Saudacmb.BoundColumn = "SAUDACODE"
    End If

End Sub
Sub ITEMLIST()
MYSQL = "SELECT DISTINCT I.ITEMCODE,I.ITEMNAME FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & MC_CODE & " AND S.COMPCODE=I.COMPCODE AND I.ITEMCODE IN(SELECT ITEMCODE FROM SAUDAMAST WHERE COMPCODE='" & MC_CODE & "'AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "') ORDER BY I.ITEMNAME"
Set ITEMREC = Nothing: Set ITEMREC = New ADODB.Recordset: ITEMREC.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
    If Not ITEMREC.EOF Then
        Set ITEMCMB.RowSource = ITEMREC: ITEMCMB.BoundColumn = "ITEMCODE": ITEMCMB.ListField = "ITEMNAME"
    End If
End Sub
Sub ADD_REC()
    If REC_ACCOUNT.RecordCount > 0 Then
        LDataImport = 0
        Frame1.Enabled = True: Combo1.ListIndex = 0
        Call Get_Selection(1)
        If vcDTP1.Enabled Then vcDTP1.SetFocus
    Else
        Call CANCEL_REC
    End If
    If fb_press = 1 Then
        MYSQL = "SELECT MAX(CAST(CONNO AS INT)) AS CONNO FROM CTR_D WHERE COMPCODE =" & MC_CODE & ""
        Set CONREC = Nothing
        Set CONREC = New ADODB.Recordset
        CONREC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
        If Not CONREC.EOF Then
            LConNo = Val(CONREC!conno & "") + Val(1)
        Else
            LConNo = 1
        End If
    End If
    RECGRID.AddNew
    RECGRID!DIMPORT = 0
    RECGRID!CONTIME = Time
    RECGRID!userid = LUserId
    RECGRID.Update
    LConNo = LConNo
    RECGRID!SRNO = LConNo  'RECGRID.AbsolutePosition
    DataGrid1.Col = 0
    
End Sub
Sub SAVE_REC()
    Dim SAUDACODE As String
    On Error GoTo ERR1
    'validation
    If vcDTP1.Value < MFIN_BEG Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
    If vcDTP1.Value > MFIN_END Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.Enabled = True: vcDTP1.SetFocus: Exit Sub
    
    'If Val(Text1.Text) = 0 And Val(Text4.Text) = 0 Then MsgBox "Please Check Entries.", vbCritical: Exit Sub
    'If Val(Text4.Text) = 0 Then MsgBox "Please Check Entries.", vbCritical:  Exit Sub
    cnn.BeginTrans
    If fb_press = 2 Then
        Set REC_CTRM = Nothing
        Set REC_CTRM = New ADODB.Recordset
        MYSQL = "SELECT * FROM CTR_M where COMPCODE=" & MC_CODE & " AND "
        If ITEMCMB.BoundText <> "" Then
            MYSQL = MYSQL & " itemcode='" & ITEMCMB.BoundText & "'  AND  "
        End If
        If Saudacmb.BoundText <> "" Then
            MYSQL = MYSQL & " sauda ='" & Saudacmb.BoundText & "'  AND  "
        End If
        MYSQL = MYSQL & " CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' ORDER BY CONSNO"
        REC_CTRM.Open MYSQL, cnn, adOpenStatic, adLockReadOnly
        If Not REC_CTRM.EOF Then
            REC_CTRM.MoveFirst
            While Not REC_CTRM.EOF
                CONSNO = REC_CTRM!CONSNO
                VCHNO = REC_CTRM!VOU_NO & ""
                Call DELETE_VOUCHER(VCHNO)
                cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & MC_CODE & " AND CONSNO=" & Val(REC_CTRM!CONSNO) & ""
                cnn.Execute "DELETE FROM CTR_M WHERE COMPCODE =" & MC_CODE & " AND CONSNO=" & Val(REC_CTRM!CONSNO) & ""
                REC_CTRM.MoveNext
            Wend
        End If
    End If
    Set GRIDREC = Nothing: Set GRIDREC = New ADODB.Recordset
    Set GRIDREC = RECGRID.Clone
    GRIDREC.MoveFirst
    While Not GRIDREC.EOF
        If GRIDREC!Qnty <= 0 Then
            GRIDREC.Delete
        End If
        GRIDREC.MoveNext
    Wend
    GRIDREC.Sort = "SAUDACODE"
    GRIDREC.MoveFirst
    SAUDACODE = ""
    MSAMT = 0
    MBAMT = 0
    While Not GRIDREC.EOF
        If SAUDACODE = GRIDREC!SAUDACODE Then
        Else
            SAUDACODE = GRIDREC!SAUDACODE
            Set REC_SAUDA = Nothing: Set REC_SAUDA = New ADODB.Recordset
            REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE COMPCODE =" & MC_CODE & " AND SAUDACODE='" & SAUDACODE & "'", cnn, adOpenForwardOnly, adLockReadOnly
            If REC_SAUDA.EOF Then
                MsgBox "Invalid Sauda Code.", vbExclamation, "Error": Text2.SetFocus: Exit Sub
            Else
                Set GeneralRec1 = Nothing: Set GeneralRec1 = New ADODB.Recordset
                GeneralRec1.Open "SELECT EX.EXCODE,EX.SHREEAC,EX.TRADINGACC  FROM EXMAST AS EX , ITEMMAST AS IM WHERE EX.COMPCODE=" & MC_CODE & " AND EX.COMPCODE=IM.COMPCODE AND EX.EXCODE=IM.EXCHANGECODE  AND  IM.ITEMCODE = '" & REC_SAUDA!ItemCode & "'", cnn, adOpenForwardOnly, adLockReadOnly
                If Not GeneralRec1.EOF Then
                    GSHREE = GeneralRec1!shreeac
                    GTRADING = GeneralRec1!TRADINGACC
                    LEXCODE = GeneralRec1!EXCODE
                End If
            End If
            If GRIDREC.RecordCount > 0 Then
                CNNERR = True
                VCHNO = VOUCHER_NUMBER("CONT", FIN_YEAR(vcDTP1.Value))
                Set REC = Nothing: Set REC = New ADODB.Recordset
                REC.Open "SELECT CONSNO FROM CTR_M WHERE COMPCODE =" & MC_CODE & " AND SAUDA='" & SAUDACODE & "' AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND pattan = '" & Mid(Combo1.Text, 1, 1) & "' ", cnn, adOpenForwardOnly, adLockReadOnly
                If Not REC.EOF Then
                    CONSNO = REC!CONSNO
                Else
                    Set REC = Nothing: Set REC = New ADODB.Recordset
                    REC.Open "SELECT MAX(CONSNO) FROM CTR_M WHERE COMPCODE =" & MC_CODE & "", cnn, adOpenForwardOnly, adLockReadOnly
                    CONSNO = Val(REC.Fields(0) & "") + Val(1)
                End If
                Set REC = Nothing
                LDataImport = IIf(IsNull(LDataImport), 0, LDataImport)
                If GRIDREC!ItemCode = "" Then
                    MsgBox "PLEASE CHECK ITEM IN THIS CONTRACT. ENTRY NOT SAVED"
                    Exit Sub
                Else
                    MYSQL = "INSERT INTO CTR_M(COMPCODE,CONSNO, CONDATE, SAUDA, ITEMCODE, VOU_NO, PATTAN,DataImport) VALUES(" & MC_CODE & "," & CONSNO & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "', '" & SAUDACODE & "', '" & GRIDREC!ItemCode & "', '" & VCHNO & "', '" & Left(Combo1.Text, 1) & "'," & LDataImport & ")"
                    cnn.Execute MYSQL
                    Dim BOOLAC As String * 1
                    Dim RC As ADODB.Recordset
                    'do not initialized LPARTY here
                    MBAMT = 0
                    MSAMT = 0
                End If
            End If
        End If
        If TempParty.EOF Then
            TempParty.AddNew
            TempParty!Acode = GRIDREC!conCode
            TempParty.Update
        Else
            TempParty.MoveFirst
            TempParty.Find "ACODE='" & GRIDREC!conCode & "'", , adSearchForward
            If TempParty.EOF Then
                TempParty.AddNew
                TempParty!Acode = GRIDREC!conCode
                TempParty.Update
            End If
        End If
        If TempSauda.EOF Then
            TempSauda.AddNew
            TempSauda!SAUDACODE = GRIDREC!SAUDACODE
            TempSauda.Update
        Else
            TempSauda.MoveFirst
            TempSauda.Find "SAUDACODE='" & GRIDREC!SAUDACODE & "'", , adSearchForward
            If TempSauda.EOF Then
                TempSauda.AddNew
                TempSauda!SAUDACODE = GRIDREC!SAUDACODE
                TempSauda.Update
            End If
        End If
        If TempItem.EOF Then
            TempItem.AddNew
            TempItem!ItemCode = GRIDREC!ItemCode
            TempItem.Update
        Else
            TempItem.MoveFirst
            TempItem.Find "ITEMCODE='" & GRIDREC!ItemCode & "'", , adSearchForward
            If TempItem.EOF Then
                TempItem.AddNew
                TempItem!ItemCode = GRIDREC!ItemCode
                TempItem.Update
            End If
        End If
        
        If TempParty.EOF Then
            TempParty.AddNew
            TempParty!Acode = GRIDREC!Code
            TempParty.Update
        Else
            TempParty.MoveFirst
            TempParty.Find "ACODE='" & GRIDREC!Code & "'", , adSearchForward
            If TempParty.EOF Then
                TempParty.AddNew
                TempParty!Acode = GRIDREC!Code
                TempParty.Update
            End If
        End If
        ''RECORDSET RC IS CHECKING WHETHER THE PARTY IS PERSONNEL OR NOT
        BOOLAC = "N"
        MCL = ""
        If Len(GRIDREC!Name & "") > Val(0) And Len(GRIDREC!conName & "") > Val(0) Then   ''WHEN BUYER AND SELLER BOTH ARE THERE
        '        If RECGRID!QNTY > Val(0) And RECGRID!BRate > Val(0) Then                   ''QNTY AND RATE REQUIRED
        '        If RECGRID!DIMPORT = 0 Then
        '            MCL = RECGRID!BCODE
        '        Else
        '            MCL = RECGRID!LCLCODE
        '        End If
        '        MBAMT = MBAMT + (Val(RECGRID!QNTY & "") * (Round(Val(RECGRID!BRate & ""), 2)) * GCALVAL)
        '        LDataImport = Abs(RECGRID!DIMPORT)
        '        MYSQL = "INSERT INTO CTR_D (COMPCODE ,CLCODE,CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT,DATAIMPORT,CONTIME,USERID) VALUES(" & MC_CODE & ",'" & MCL & "'," & Val(CONSNO) & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'," & Val(RECGRID!SRNO) & ",'" & Text2.Text & "', '" & DataCombo2.BoundText & "', '" & RECGRID!BCODE & "', " & Val(RECGRID!QNTY) & "," & Val(RECGRID!BRate) & ",'B', '" & BOOLAC & "'," & LDataImport & ",'" & RECGRID!CONTIME & "','" & (RECGRID!UserId & "") & "')"
        '        cnn.Execute MYSQL
        '    End If
        '    If RECGRID!QNTY > Val(0) And RECGRID!SRate > Val(0) Then               ''QNTY AND RATE REQUIRED
        '        LDataImport = Abs(RECGRID!DIMPORT)
        '        If RECGRID!DIMPORT = 0 Then
        '            MCL = RECGRID!scode
        '            Else
        '            MCL = RECGRID!RCLCODE
        '        End If
        '        MSAMT = MSAMT + (Val(RECGRID!QNTY & "") * Round(Val(RECGRID!SRate) & "", 2) * GCALVAL)
        '        MYSQL = "INSERT INTO CTR_D(COMPCODE,CLCODE,CONSNO, CONDATE, CONNO, SAUDA, ITEMCODE, PARTY, QTY, RATE, CONTYPE, PERCONT,DataImport,CONTIME,UserId) VALUES(" & MC_CODE & ",'" & MCL & "'," & Val(CONSNO) & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'," & Val(RECGRID!SRNO) & ",'" & Text2.Text & "', '" & DataCombo2.BoundText & "', '" & RECGRID!scode & "', " & Val(RECGRID!QNTY) & "," & Round(Val(RECGRID!SRate), 2) & ",'S', '" & BOOLAC & "' ," & LDataImport & ",'" & RECGRID!CONTIME & "','" & (RECGRID!UserId & "") & "')"
        '        cnn.Execute MYSQL
        '        End If
            If GRIDREC!Qnty > Val(0) And GRIDREC!Rate > Val(0) And GRIDREC!Rate1 > Val(0) Then                   ''QNTY AND RATE REQUIRED
                'If GRIDREC!DIMPORT = 0 Then
                    If GRIDREC!BUYSELL = "B" Then
                        MBCL = GRIDREC!Code
                        MSCL = GRIDREC!conCode
                        
                    Else
                        MBCL = GRIDREC!conCode
                        MSCL = GRIDREC!Code
                    End If
                'Else
                 '   MBCL = RECGRID!LCLCODE
                  '  MSCL = RECGRID!RCLCODE
                'End If
                If GRIDREC!BUYSELL = "B" Then
                    MBAMT = MBAMT + (Val(GRIDREC!Qnty & "") * (Round(Val(GRIDREC!Rate & ""), 2)) * GRIDREC!LOT)
                    LDataImport = Abs(GRIDREC!DIMPORT)
                    MSAMT = MSAMT + (Val(GRIDREC!Qnty & "") * Round(Val(GRIDREC!Rate1) & "", 2) * GRIDREC!LOT)
                    BUYYER = GRIDREC!Code
                    SELLER = GRIDREC!conCode
                    BRate = GRIDREC!Rate
                    SRate = GRIDREC!Rate1
                Else
                    MSAMT = MSAMT + (Val(GRIDREC!Qnty & "") * (Round(Val(GRIDREC!Rate & ""), 2)) * GRIDREC!LOT)
                    LDataImport = Abs(GRIDREC!DIMPORT)
                    MBAMT = MBAMT + (Val(GRIDREC!Qnty & "") * Round(Val(GRIDREC!Rate1) & "", 2) * GRIDREC!LOT)
                    SELLER = GRIDREC!Code
                    BUYYER = GRIDREC!conCode
                    SRate = GRIDREC!Rate
                    BRate = GRIDREC!Rate1
                End If
                If GRIDREC!userid = "" Then
                    MYSQL = "SELECT FY.FMLYCODE,EX.CONTRACTACC FROM EXMAST AS EX , ITEMMAST AS IT, ACCFMLY AS FY WHERE EX.COMPCODE =" & MC_CODE & " AND  EX.COMPCODE = IT.COMPCODE  AND EX.EXCODE = IT.ExchangeCode AND IT.ITEMCODE = '" & REC_SAUDA!ItemCode & "' AND EX.COMPCODE  = FY.COMPCODE AND EX.CONTRACTACC = FY.FMLYHEAD "
                    Set RECEX = Nothing: Set RECEX = New ADODB.Recordset: RECEX.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
                    If Not RECEX.EOF Then
                        GRIDREC!userid = RECEX!FMLYCode
                    End If
                End If
                If Combo1.ListIndex = 0 Then
                    PATTAN = "C"
                Else
                     PATTAN = "O"
                End If
            
                MYSQL = "EXEC INSERT_CTR_D " & MC_CODE & ",'" & MBCL & "','" & MSCL & "'," & Val(CONSNO) & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'," & Val(GRIDREC!SRNO) & ",'" & GRIDREC!SAUDACODE & "','" & GRIDREC!ItemCode & "','" & BUYYER & "'," & Val(GRIDREC!Qnty) & "," & BRate & ",'" & BOOLAC & "','" & SELLER & "'," & Val(GRIDREC!Qnty) & "," & Round(Val(SRate), 2) & "," & LDataImport & ",'" & GRIDREC!CONTIME & "','" & IIf(IsNull(GRIDREC!userid), "", GRIDREC!userid) & " " & "','" & IIf(IsNull(GRIDREC!ORDER_NO), "", GRIDREC!ORDER_NO) & "'," & IIf(IsNull(GRIDREC!TRADE_NO), Val(GRIDREC!SRNO), GRIDREC!TRADE_NO) & ",'" & LEXCODE & "','" & PATTAN & "','" & GRIDREC!conCode & "'"
            '                            @COMPCODE    , @BCLCODE     ,    @SCLCODE  ,     @CONSNO        ,     @CONDATE                              ,     @CONNO FLOAT        ,    @SAUDA          ,    @ITEMCODE                 ,   @BPARTY            ,     @BQTY                ,     @BRATE               , @PERCONT       ,@SPARTY                ,    @SQTY                 , @SRATE                             ,  @DATAIMPORT              , @CONTIME , @USERID NVARCHAR(50),@ORDNO NVARCHAR(100),@ROWNO1
                cnn.Execute MYSQL
            End If
            
            GRIDREC.MoveNext
        End If
        If GRIDREC.EOF Then
            If (MBAMT - MSAMT) <> Val(0) Then
                MYSQL = "INSERT INTO VOUCHER(COMPCODE,VOU_NO, VOU_DT, VOU_TYPE, VOU_PR, BILLNO, BILLDT, USER_NAME, USER_DATE, USER_TIME, USER_ACTION) VALUES(" & MC_CODE & ",'" & VCHNO & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','O','','" & 0 & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & USER_ID & "','" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD')"
                cnn.Execute MYSQL
                MAMOUNT = Abs(Val((MBAMT - MSAMT)))
                If (MBAMT - MSAMT) < Val(0) Then
                    MCR = "C"
                    MDR = "D"
                    sql = "DEBIT=DEBIT+"
                    SQL1 = "CREDIT=CREDIT+"
                Else
                    MCR = "D"
                    MDR = "C"
                    sql = "CREDIT=CREDIT+"
                    SQL1 = "DEBIT=DEBIT+"
                End If
                MNARATION = "Shree for : " & SAUDACODE & ", " & Format(vcDTP1.Value, "DD/MM/YYYY")
                ''SHREE POSTING
                MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MDR & "','" & GSHREE & "'," & Val(MAMOUNT) & ",'" & MNARATION & "')"
                cnn.Execute MYSQL
                ''TRADING AC POSTING
                MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MCR & "','" & GTRADING & "'," & Val(MAMOUNT) & ",'" & MNARATION & "')"
                cnn.Execute MYSQL
                MBAMT = 0
                MSAMT = 0
            End If
        Else
            If SAUDACODE = GRIDREC!SAUDACODE Then
            Else
                If (MBAMT - MSAMT) <> Val(0) Then
                    MYSQL = "INSERT INTO VOUCHER(COMPCODE,VOU_NO, VOU_DT, VOU_TYPE, VOU_PR, BILLNO, BILLDT, USER_NAME, USER_DATE, USER_TIME, USER_ACTION) VALUES(" & MC_CODE & ",'" & VCHNO & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','O','','" & 0 & "','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & USER_ID & "','" & Format(Date, "yyyy/MM/dd") & "','" & Time & "','ADD')"
                    cnn.Execute MYSQL
                    MAMOUNT = Abs(Val((MBAMT - MSAMT)))
                    If (MBAMT - MSAMT) < Val(0) Then
                        MCR = "C"
                        MDR = "D"
                        sql = "DEBIT=DEBIT+"
                        SQL1 = "CREDIT=CREDIT+"
                    Else
                        MCR = "D"
                        MDR = "C"
                        sql = "CREDIT=CREDIT+"
                        SQL1 = "DEBIT=DEBIT+"
                    End If
                    MNARATION = "Shree for : " & SAUDACODE & ", " & Format(vcDTP1.Value, "DD/MM/YYYY")
                    ''SHREE POSTING
                    MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MDR & "','" & GSHREE & "'," & Val(MAMOUNT) & ",'" & MNARATION & "')"
                    cnn.Execute MYSQL
                    ''TRADING AC POSTING
                    MYSQL = "INSERT INTO VCHAMT(COMPCODE,VOU_NO, VOU_TYPE, VOU_DT, DR_CR, AC_CODE, AMOUNT, NARRATION) VALUES(" & MC_CODE & ",'" & VCHNO & "','O','" & Format(vcDTP1.Value, "yyyy/MM/dd") & "','" & MCR & "','" & GTRADING & "'," & Val(MAMOUNT) & ",'" & MNARATION & "')"
                    cnn.Execute MYSQL
                    MBAMT = 0
                    MSAMT = 0
                End If
            End If
        End If
    Wend

            Lparty = ""
            If Not TempParty.EOF Then
                TempParty.MoveFirst
                Do While Not TempParty.EOF
                    If Lparty = "" Then
                        Lparty = "'" & TempParty!Acode & "'"
                        
                    Else
                        Lparty = Lparty & ",'" & TempParty!Acode & "'"
                    End If
                TempParty.MoveNext
                Loop
            End If
            LSaudaCode = ""
            If Not TempSauda.EOF Then
                TempSauda.MoveFirst
                Do While Not TempSauda.EOF
                    If LSaudaCode = "" Then
                        LSaudaCode = "'" & TempSauda!SAUDACODE & "'"
                        
                    Else
                        LSaudaCode = LSaudaCode & ",'" & TempSauda!SAUDACODE & "'"
                    End If
                TempSauda.MoveNext
                Loop
            End If
            If Not TempItem.EOF Then
                TempItem.MoveFirst
                Do While Not TempItem.EOF
                    If LITEMCODE = "" Then
                        LITEMCODE = "'" & TempItem!ItemCode & "'"
                        
                    Else
                        LITEMCODE = LITEMCODE & ",'" & TempItem!ItemCode & "'"
                    End If
                TempItem.MoveNext
                Loop
            End If
            
            Call UpdateBrokRateType(True, True, Lparty, LITEMCODE, vcDTP1.Value, vcDTP1.Value, vcDTP1.Value, LSaudaCode)
            cnn.CommitTrans
            If GAppSpread = "Y" Then
                Call UpdateMargin(Lparty, "" & LSaudaCode & "", vcDTP1.Value, vcDTP1.Value)
            End If
            CNNERR = False
            MFROMDATE = Format(vcDTP1.Value, "yyyy/MM/dd")
            MYSQL = "SELECT MATURITY FROM SAUDAMAST WHERE COMPCODE =" & MC_CODE & " AND SAUDACODE = '" & SAUDACODE & "'"
            Set REC = Nothing: Set REC = New ADODB.Recordset: REC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
            If Not REC.EOF Then MTODATE = REC.Fields(0)
            cnn.BeginTrans
            CNNERR = False
            If BILL_GENERATION(CDate(MFROMDATE), CDate(MTODATE), LSaudaCode, Lparty) Then
                cnn.CommitTrans: CNNERR = False
            Else
                cnn.RollbackTrans: CNNERR = False
            End If
    Call CANCEL_REC
    Exit Sub
ERR1:
    MsgBox Err.Description, vbCritical, "Error Number : " & Err.Number
   ' Resume
    If CNNERR = True Then
        cnn.RollbackTrans: CNNERR = False
    End If
    
 End Sub
Sub CANCEL_REC()
    vcDTP1.Enabled = True:  Combo1.Enabled = True
    Call RECSET
    Set REC_CTRM = Nothing
    Set REC_CTRM = New ADODB.Recordset
    MYSQL = "SELECT * FROM CTR_M where COMPCODE=" & MC_CODE & " ORDER BY CONSNO"
    REC_CTRM.Open MYSQL, cnn, adOpenStatic, adLockReadOnly
    Call ITEMLIST
    Call SaudaList
    CONFLAGE = False
    fb_press = 0
    Set DataGrid1.DataSource = RECGRID
    ITEMCMB.Enabled = True
    Saudacmb.Enabled = True
    Combo1.Enabled = True
    DataGrid1.Refresh
    Label2.Visible = False
    DataCombo3.Visible = False
    Call ClearFormFn(SINGLECONTRACT)
    Call Get_Selection(10)
    Combo1.ListIndex = -1: Frame1.Enabled = False
End Sub
Function MODIFY_REC(LCONDATE As Date, LSaudaCode As String, LITEMCODE As String, LPATTAN As String) As Boolean



Set REC_CTRM = Nothing
Set REC_CTRM = New ADODB.Recordset
MYSQL = "SELECT * FROM CTR_M WHERE COMPCODE =" & MC_CODE & " AND PATTAN='" & Left(LPATTAN, 1) & "'AND CONDATE='" & Format(LCONDATE, "YYYY/MM/DD") & "'  "
If LITEMCODE <> "" Then
    MYSQL = MYSQL & "AND itemcode='" & LITEMCODE & "'"
End If
If LSaudaCode <> "" Then
    MYSQL = MYSQL & "AND sauda='" & LSaudaCode & "'"
End If
MYSQL = MYSQL & " ORDER BY SAUDA"
REC_CTRM.Open MYSQL, cnn, , adLockReadOnly


If Not REC_CTRM.EOF Then
    vcDTP1.Value = LCONDATE
    ITEMCMB.BoundText = LITEMCODE
    Saudacmb.BoundText = LSaudaCode
    If Left(LPATTAN, 1) = "C" Then
        Combo1.ListIndex = 0
    Else
        Combo1.ListIndex = 1
    End If
    If fb_press = 1 Then
        If MsgBox("Contract Already Exist For Selected Criteria.Press OK To Modify The Existing Contracts.", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            fb_press = 2
        Else
            Call CANCEL_REC
            Exit Function
        End If
    End If
    Call RECSET
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    MYSQL = "SELECT I.LOT,C.* FROM CTR_D AS C ,ITEMMAST AS I WHERE C.COMPCODE =" & MC_CODE & " AND C.CONDATE='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND I.COMPCODE=C.COMPCODE AND I.ITEMCODE=C.ITEMCODE AND C.PATTAN='" & Left(LPATTAN, 1) & "' "
    If ITEMCMB.BoundText <> "" Then
        MYSQL = MYSQL & "AND C.itemcode='" & ITEMCMB.BoundText & "'"
    End If
    If Saudacmb.BoundText <> "" Then
        MYSQL = MYSQL & "AND C.sauda='" & Saudacmb.BoundText & "'"
    End If
    MYSQL = MYSQL & " ORDER BY C.CONNO,C.ROWNO"
    REC.Open MYSQL, cnn, , adLockReadOnly
    
    If Not REC.EOF Then
        REC.MoveFirst
        Lparty = ""
        While Not REC.EOF
            RECGRID.AddNew
            RECGRID!SRNO = REC!conno 'RECGRID.AbsolutePositi
            If Len(REC!conCode & "") > Val(0) Then
                If REC!PARTY <> REC!conCode Then
                    RECGRID!BUYSELL = REC!CONTYPE
                    RECGRID!Code = REC!PARTY & ""
                    REC_ACCOUNT.MoveFirst
                    REC_ACCOUNT.Find "AC_CODE='" & REC!PARTY & "'"
                    RECGRID!Name = REC_ACCOUNT!Name
                    RECGRID!Rate = Round(REC!Rate, 2)
                Else
                    RECGRID!BUYSELL = "S"
                    RECGRID!conCode = REC!PARTY & ""
                    REC_ACCOUNT.MoveFirst
                    REC_ACCOUNT.Find "AC_CODE='" & REC!PARTY & "'"
                    RECGRID!conName = REC_ACCOUNT!Name
                    RECGRID!Rate1 = Round(REC!Rate, 2)
                End If
            Else
                RECGRID!BUYSELL = REC!CONTYPE
                RECGRID!Code = REC!PARTY & ""
                REC_ACCOUNT.MoveFirst
                REC_ACCOUNT.Find "AC_CODE='" & REC!PARTY & "'"
                RECGRID!Name = REC_ACCOUNT!Name
            End If
            RECGRID!ItemCode = REC!ItemCode & ""
            RECGRID!LOT = Val(REC!LOT & "")
            RECGRID!SAUDACODE = REC!Sauda & ""
            RECGRID!LCLCODE = REC!CLCODE & ""
            ITEMREC.MoveFirst
            ITEMREC.Find "ITEMCODE ='" & REC!ItemCode & "'"
            SaudaRec.MoveFirst
            SaudaRec.Find "SAUDACODE ='" & REC!Sauda & "'"
            RECGRID!Qnty = REC!QTY
            RECGRID!ITEMName = ITEMREC!ITEMName
            RECGRID!SAUDANAME = SaudaRec!SAUDANAME
            
            RECGRID!LInvNo = Val(REC!INVNO & "")
            If Not IsNull(REC!DATAIMPORT) Then
                If REC!DATAIMPORT = True Then
                    RECGRID!DIMPORT = 1
                Else
                    RECGRID!DIMPORT = 0
                End If
            Else
                RECGRID!DIMPORT = 0
            End If
            RECGRID!CONTIME = IIf(IsNull(REC!CONTIME), Time, REC!CONTIME)
            RECGRID!userid = REC!userid & ""
            REC.MoveNext
           ' REC.MovePrevious
            If Len(REC!conCode & "") > Val(0) Then
                If REC!PARTY <> REC!conCode Then
                    RECGRID!Code = REC!PARTY & ""
                    REC_ACCOUNT.MoveFirst
                    REC_ACCOUNT.Find "AC_CODE='" & REC!PARTY & "'"
                    RECGRID!Name = REC_ACCOUNT!Name
                    RECGRID!Rate = Round(REC!Rate, 2)
                Else
                    RECGRID!conCode = REC!PARTY & ""
                    REC_ACCOUNT.MoveFirst
                    REC_ACCOUNT.Find "AC_CODE='" & REC!PARTY & "'"
                    RECGRID!conName = REC_ACCOUNT!Name
                    RECGRID!Rate1 = Round(REC!Rate, 2)
                End If
            Else
                RECGRID!conCode = REC!PARTY & ""
                REC_ACCOUNT.MoveFirst
                REC_ACCOUNT.Find "AC_CODE='" & REC!PARTY & "'"
                RECGRID!conName = REC_ACCOUNT!Name
            End If
            RECGRID!RCLCODE = REC!CLCODE & ""
            REC_ACCOUNT.MoveFirst
            
            RECGRID!RInvNo = Val(REC!INVNO & "")
            RECGRID.Update
            REC.MoveNext
        Wend
        If Saudacmb.BoundText <> "" Then
            DataGrid1.Columns(5).Locked = True
            Saudacmb.Enabled = False
        Else
            DataGrid1.Columns(5).Locked = False
        End If
        
        Set DataGrid1.DataSource = RECGRID
        DataGrid1.ReBind
        DataGrid1.Col = 0
        Call DataGrid1_AfterColEdit(0)
        ITEMCMB.Enabled = False
        Saudacmb.Enabled = False
        vcDTP1.Enabled = False
        Combo1.Enabled = False
        MODIFY_REC = True
    End If
Else
    If fb_press = 1 Then
        DataGrid1.Col = 0
        DataGrid1.SetFocus
        If Saudacmb.BoundText <> "" Then
            DataGrid1.Columns(5).Locked = True
            
        Else
            DataGrid1.Columns(5).Locked = False
        End If
        Saudacmb.Enabled = False
        ITEMCMB.Enabled = False
        vcDTP1.Enabled = False
        Combo1.Enabled = False
        fb_press = 1
        MODIFY_REC = True
    Else
        MsgBox "Record Does Not Exist For Selected Criteria.", vbInformation
        If fb_press = 2 Then
            DataGrid1.Col = 0
            DataGrid1.SetFocus
            If Saudacmb.BoundText <> "" Then
                DataGrid1.Columns(5).Locked = True
            Else
                DataGrid1.Columns(5).Locked = False
            End If
            Saudacmb.Enabled = False
            ITEMCMB.Enabled = False
            vcDTP1.Enabled = False
            Combo1.Enabled = False
            fb_press = 1
            MODIFY_REC = True
        Else
            Call CANCEL_REC
            Exit Function
        End If
    End If
End If
If fb_press = 3 Then
    If MsgBox("You are about to Delete all Contracts. Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
        cnn.BeginTrans
        CNNERR = True
        REC_CTRM.MoveFirst
        While Not REC_CTRM.EOF
            MYSQL = "DELETE FROM CTR_D WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & REC_CTRM!CONSNO & ""
            cnn.Execute MYSQL
            MYSQL = "DELETE FROM CTR_R WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & REC_CTRM!CONSNO & ""
            cnn.Execute MYSQL
            Call DELETE_VOUCHER(REC_CTRM!VOU_NO & "")
            MYSQL = "DELETE FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND CONSNO=" & REC_CTRM!CONSNO & ""
            cnn.Execute MYSQL
            REC_CTRM.MoveNext
        Wend
        cnn.CommitTrans
        Call CANCEL_REC
    End If
End If

        
End Function

Private Sub cmdImportFromExcel_Click()
    Dim exlCnn As New ADODB.Connection
    Dim conStr As String
    
    
    conStr = _
    "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=C:\cust.xls;DefaultDir=c:\;"
    exlCnn.ConnectionString = conStr
    exlCnn.Open
        
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open "SELECT * FROM [customers$]", exlCnn, adOpenDynamic, adLockOptimistic
    
    If (REC.RecordCount > 0) Then
        Set DataGrid1.DataSource = REC
    End If
    Set exlCnn = Nothing
            
End Sub

Private Sub Combo1_GotFocus()
    If FLOWDIR = 1 Then
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT * FROM CTR_M WHERE COMPCODE=" & MC_CODE & " AND SAUDA='" & DataCombo1.BoundText & "'", cnn, adOpenForwardOnly, adLockReadOnly
        If REC.EOF Then SendKeys "%{DOWN}"
    End If
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 1 Then
    flag = True
    End If
End Sub

Private Sub Combo1_LostFocus()
If MODIFY_REC(vcDTP1.Value, Saudacmb.BoundText, ITEMCMB.BoundText, Combo1.Text) Then
Else
    Combo1.SetFocus
End If

End Sub

''''Private Sub Combo2_GotFocus()
''''SendKeys "%{DOWN}"
''''DataGrid1.Col = 0
'''''DataGrid1.Text = ""
''''Combo2.Left = Val(480)
''''Combo2.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
''''End Sub
''
''Private Sub Combo2_Validate(Cancel As Boolean)
''    If Combo2.ListIndex = 0 Then
''        RECGRID!BUYSELL = "B"
''    Else
''        RECGRID!BUYSELL = "S"
''    End If
''    DataGrid1.SetFocus
''    DataGrid1.Col = 1
''    Combo2.Visible = False
''    DataGrid1.SetFocus
''
''End Sub

Private Sub DataCombo1_GotFocus()

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not RECGRID.EOF Then
    
    Set TEMPORARY = Nothing: Set TEMPORARY = New ADODB.Recordset
    Set TEMPORARY = RECGRID.Clone
    TEMPORARY.ActiveConnection = Nothing
    SQTY = 0: BQTY = 0: SRate = 0: BRate = 0: DIFFAMT = 0: SAVG = 0: BAVG = 0: LOT = 0: BTOT = 0: STOT = 0: TOTBRATE = 0: TOTSRate = 0: BUYAMT = 0: SELAMT = 0
    TEMPORARY.Filter = "SAUDACODE='" & RECGRID!SAUDACODE & "'"
    If Not TEMPORARY.EOF Then
        TEMPORARY.MoveFirst
        Label3.Caption = (RECGRID!SAUDACODE & "")
        While Not TEMPORARY.EOF
            
            If TEMPORARY!BUYSELL = "B" Then
                BRate = BRate + IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
                BQTY = BQTY + IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty)
                TOTBRATE = TOTBRATE + IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
                BUYAMT = BUYAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
                SELLAMT = SELLAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
            Else
                SRate = SRate + IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
                SQTY = SQTY + IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty)
                TOTSRate = TOTSRate + IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
                BUYAMT = BUYAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
                SELLAMT = SELLAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
            End If
            LOT = IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT)
            
            
            TEMPORARY.MoveNext
        Wend
    End If
    'DIFFAMT = (SRate - BRate) * lot
    If Not SQTY = 0 Then
        SAVG = TOTSRate / SQTY
        STOT = SRate * SQTY * LOT
    End If
    If Not BQTY = 0 Then
        BAVG = TOTBRATE / BQTY
        BTOT = BRate * BQTY * LOT
    End If
    'Total Shree Caculation
    TEMPORARY.Filter = adFilterNone
    TEMPORARY.MoveFirst
    TOTBUYAMT = 0: TOTSELLAMT = 0
    While Not TEMPORARY.EOF
        If TEMPORARY!BUYSELL = "B" Then
            TOTBUYAMT = TOTBUYAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
            TOTSELLAMT = TOTSELLAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
        Else
            TOTBUYAMT = TOTBUYAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
            TOTSELLAMT = TOTSELLAMT + (IIf(IsNull(TEMPORARY!Qnty), 0, TEMPORARY!Qnty) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
        End If
        TEMPORARY.MoveNext
    Wend
    
    Text2.Text = TOTSELLAMT - TOTBUYAMT
    Text6.Text = SELLAMT - BUYAMT
    Text1.Text = BQTY: Text4.Text = SQTY: Text7.Text = BAVG: Text8.Text = SAVG: Text9.Text = BTOT: Text10.Text = STOT:
    End If
End Sub


Private Sub ITEMCMB_GotFocus()
    Call ITEMLIST
    SendKeys "%{DOWN}"
End Sub

Private Sub DataCombo3_GotFocus()
    Set REC_ACCOUNT = Nothing: Set REC_ACCOUNT = New ADODB.Recordset
    REC_ACCOUNT.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & MC_CODE & " AND gcode in (12,14) ORDER BY NAME ", cnn, adOpenKeyset, adLockReadOnly
    If Not REC_ACCOUNT.EOF Then Set DataCombo3.RowSource = REC_ACCOUNT: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
    SendKeys "%{DOWN}"
    If DataGrid1.Col = 2 Or DataGrid1.Col = 1 Then
        DataGrid1.Col = 1
        DataGrid1.Text = ""
        Label2.Visible = True: Label2.Left = 1080
        DataCombo3.Left = Val(1080)
        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    ElseIf DataGrid1.Col = 7 Or DataGrid1.Col = 8 Then
        DataGrid1.Col = 8: DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
        DataCombo3.Left = DataGrid1.Columns(7).Left
        Label2.Visible = True: Label2.Left = DataGrid1.Columns(9).Left
    End If
    SendKeys "%{DOWN}"
    SendKeys "%{DOWN}"
End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DataGrid1.Col = 1 Or DataGrid1.Col = 2 Then
        If DataCombo3.BoundText <> "" Then
            RECGRID!Code = DataCombo3.BoundText
            RECGRID!Name = DataCombo3.Text
            RECGRID!userid = LUserId
            DataGrid1.Col = 2
            DataGrid1.SetFocus
            DataCombo3.Visible = False: Label2.Visible = False
        End If
    ElseIf DataGrid1.Col = 8 Or DataGrid1.Col = 7 Then
        If DataCombo3.BoundText <> "" Then
            RECGRID!conCode = DataCombo3.BoundText
            RECGRID!conName = DataCombo3.Text
            RECGRID!userid = LUserId
            DataCombo3.Visible = False: Label2.Visible = False
            DataGrid1.Col = 8
            DataGrid1.SetFocus
            
        End If
    End If
        
ElseIf KeyAscii = 27 Then
    DataGrid1.SetFocus
    DataCombo3.Visible = False: Label2.Visible = False
ElseIf KeyAscii = 121 Then   'F3  NEW PARTY
    GETACNT.Show
    GETACNT.ZOrder
    GETACNT.add_record
ElseIf KeyAscii = 18 Then
    DataCombo3.Visible = True: DataCombo3.SetFocus
End If

End Sub

Private Sub DataCombo3_Validate(Cancel As Boolean)
    If DataCombo3.Visible = True Then
        Cancel = True
    Else
        Label2.Visible = False
    End If
End Sub
Public Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)
    If ColIndex = Val(0) Then
        If UCase(Left(Trim(DataGrid1.Text), 1)) = "S" Or UCase(Left(Trim(DataGrid1.Text), 1)) = "B" Then
            DataGrid1.Text = Left(UCase(DataGrid1.Text), 1)
        Else
           DataGrid1.Text = "B"
           DataGrid1.Col = 0
           ' Combo2.Visible = True
           ' Combo2.SetFocus
        End If
    ElseIf ColIndex = Val(1) Then
        MYSQL = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & MC_CODE & " AND A.AC_CODE= '" & DataGrid1.Text & "'"
        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, cnn
        If TempRec.RecordCount > 0 Then
            RECGRID!Code = TempRec!AC_CODE
            RECGRID!Name = TempRec!Name
        Else
            DataCombo3.Visible = True
            DataCombo3.SetFocus
            RECGRID!Code = ""
            DataGrid1.Col = 1
            DataCombo3.SetFocus
            'Exit Sub
        End If
    ElseIf ColIndex = 4 Then
        ''IF CONTRACT THEN ONLY CHANGE OCCURS
        If Val(RECGRID!Rate & "") > 0 Then
            If Val(Round(RECGRID!Rate1, 2) & "") = Val(0) Then RECGRID!Rate1 = Round(RECGRID!Rate, 2)
        Else
            If ColIndex = 5 Then
            Else
                MsgBox "Rate can not be zero.Please enter rate.", vbCritical
                DataGrid1.Col = 6: DataGrid1.SetFocus
            End If
        End If
    ElseIf ColIndex = Val(5) Then
        MYSQL = "SELECT I.ITEMCODE,I.ITEMNAME,I.LOT,S.SAUDACODE,S.SAUDANAME FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & MC_CODE & " AND I.COMPCODE=S.COMPCODE AND I.ITEMCODE=S.ITEMCODE AND S.SAUDACODE= '" & DataGrid1.Text & "'"
        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, cnn
        If TempRec.RecordCount > 0 Then
            TempRec.MoveFirst
            TempRec.Find "saudaCODE='" & DataGrid1.Text & "'", , adSearchForward
            If Not TempRec.EOF Then
                RECGRID!ItemCode = TempRec!ItemCode
                RECGRID!ITEMName = TempRec!ITEMName
                RECGRID!SAUDACODE = TempRec!SAUDACODE
                RECGRID!ITEMName = TempRec!SAUDANAME
                RECGRID!LOT = TempRec!LOT
            Else
                RECGRID!SAUDACODE = ""
                DataGrid1.Col = 5
                Saudacombo.Visible = True
                Saudacombo.SetFocus
            End If
        End If
    End If
End Sub

Private Sub DataGrid1_GotFocus()
'    vcDTP1.Enabled = False
'    Text2.Enabled = False
   ' ITEMCMB.Enabled = False
 '   Combo1.Enabled = False
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If DataGrid1.Enabled = True Then
    If KeyCode = 13 And ((DataGrid1.Col = 9) Or (DataGrid1.Col = 4 And Saudacmb.BoundText <> "" And RECGRID!conCode <> "") Or (DataGrid1.Col = 6 And RECGRID!conCode <> "")) Then
        BCODE = RECGRID!Code
        BNAME = RECGRID!Name
        ItemCode = RECGRID!ItemCode
        ITEMName = RECGRID!ITEMName
        LOT = RECGRID!LOT
        SAUDACODE = RECGRID!SAUDACODE
        SAUDANAME = RECGRID!SAUDANAME
        BUYSELL = RECGRID!BUYSELL
        scode = RECGRID!conCode
        SNAME = RECGRID!conName
        
        RECGRID.MoveNext
        If RECGRID.EOF Then
            RECGRID.AddNew
            RECGRID!Code = BCODE
            RECGRID!Name = BNAME
            RECGRID!conCode = scode
            RECGRID!conName = SNAME
            RECGRID!ItemCode = ItemCode
            RECGRID!ITEMName = ITEMName
            RECGRID!LOT = LOT
            RECGRID!BUYSELL = BUYSELL
            RECGRID!SAUDACODE = SAUDACODE
            RECGRID!SAUDANAME = SAUDANAME
            RECGRID!Qnty = 0
            RECGRID!Rate = Round(Val(0), 2)
            RECGRID!Rate1 = Round(Val(0), 2)
            
            RECGRID!DIMPORT = 0
            RECGRID!userid = LUserId & ""
            RECGRID!CONTIME = Time
            MYSQL = "SELECT MAX(CAST(CONNO AS INT)) AS CONNO FROM CTR_D WHERE COMPCODE =" & MC_CODE & ""
            Set CONREC = Nothing
            Set CONREC = New ADODB.Recordset
            CONREC.Open MYSQL, cnn, adOpenForwardOnly, adLockReadOnly
            If fb_press = 2 And CONFLAGE = False Then
                LConNo = Val(CONREC!conno & "") + Val(1)
            Else
                LConNo = LConNo + 1
            End If
                CONFLAGE = True
            
                RECGRID!SRNO = LConNo 'RECGRID.AbsolutePosition
                RECGRID.Update
            End If
        DataGrid1.LeftCol = 0
        DataGrid1.Col = 0
    ElseIf DataGrid1.Col = 4 Or DataGrid1.Col = 11 Then
        If KeyCode = 13 Or KeyCode = 9 Then
            If Val(DataGrid1.Text) = 0 Then
               MsgBox "Rate Cannot Be Zero", vbCritical
               DataGrid1.SetFocus
               Exit Sub
            End If
        End If
    ElseIf DataGrid1.Col = Val(7) And (KeyCode = 13 Or KeyCode = 9) Then
        MYSQL = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & MC_CODE & " AND A.AC_CODE= '" & DataGrid1.Text & "'"
        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, cnn
        If TempRec.RecordCount > 0 Then
            RECGRID!conCode = TempRec!AC_CODE
            RECGRID!conName = TempRec!Name
        Else
            RECGRID!conCode = ""
            DataGrid1.Col = 7
            DataCombo3.Visible = True
            DataCombo3.SetFocus
            DataCombo3.SetFocus
            Exit Sub
        End If
    
    ElseIf DataGrid1.Col = Val(1) And (KeyCode = 13 Or KeyCode = 9) Then
        MYSQL = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & MC_CODE & " AND A.AC_CODE= '" & DataGrid1.Text & "'"
        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, cnn
        If TempRec.RecordCount > 0 Then
            RECGRID!Code = TempRec!AC_CODE
            RECGRID!Name = TempRec!Name
        Else
            RECGRID!Code = ""
            DataGrid1.Col = 1
            DataCombo3.Visible = True
            DataCombo3.SetFocus
            DataCombo3.SetFocus
            Exit Sub
        End If
    ElseIf DataGrid1.Col = Val(5) And (KeyCode = 13 Or KeyCode = 9) Then
        MYSQL = "SELECT I.ITEMCODE,I.ITEMNAME,I.LOT,S.SAUDACODE,S.SAUDANAME FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & MC_CODE & " AND I.COMPCODE=S.COMPCODE AND I.ITEMCODE=S.ITEMCODE AND S.SAUDACODE= '" & DataGrid1.Text & "'"
        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, cnn
        If TempRec.RecordCount > 0 Then
            TempRec.MoveFirst
            TempRec.Find "saudaCODE='" & DataGrid1.Text & "'", , adSearchForward
            If Not TempRec.EOF Then
                RECGRID!ItemCode = TempRec!ItemCode
                RECGRID!ITEMName = TempRec!ITEMName
                RECGRID!SAUDACODE = TempRec!SAUDACODE
                RECGRID!ITEMName = TempRec!SAUDANAME
                RECGRID!LOT = TempRec!LOT
            Else
                RECGRID!SAUDACODE = ""
                DataGrid1.Col = 5
                Saudacombo.Visible = True
                Saudacombo.SetFocus
            End If
        Else
            RECGRID!SAUDACODE = ""
            DataGrid1.Col = 5
            Saudacombo.Visible = True
            Saudacombo.SetFocus
        End If
        
    ElseIf KeyCode = 114 Then   'F3  NEW PARTY
        GETACNT.Show
        GETACNT.ZOrder
        GETACNT.add_record
    ElseIf KeyCode = 118 Then   ''F7 KEY
        RNO = InputBox("Enter the row number.", "Sauda")
        If Val(RNO) > Val(0) Then
            RECGRID.MoveFirst
            RECGRID.Find "SRNO=" & Val(RNO) & "", , adSearchForward
            If RECGRID.EOF Then
                MsgBox "Record not found.", vbCritical, "Error"
                RECGRID.MoveFirst
            End If
            DataGrid1.Col = 1
            DataGrid1.SetFocus
        End If
    ElseIf KeyCode = 46 And Shift = 2 Then
        RECGRID.Delete
        If RECGRID.RecordCount = 0 Then
            RECGRID.AddNew
            LConNo = LConNo + 1
            RECGRID!SRNO = LConNo 'RECGRID.RecordCount
            If Combo1.ListIndex = Val(1) Then
                RECGRID!BRate = Round(Val(Text3.Text), 2)
                RECGRID!SRate = Round(Val(Text3.Text), 2)
                RECGRID!userid = LUserId
            End If
            RECGRID.Update
        End If
        Call DataGrid1_AfterColEdit(0)
    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 0) Then
        If Len(Trim(DataGrid1.Text)) < 1 Then
           ' Combo2.Visible = True
           ' Combo2.ListIndex = 0
           ' Combo2.SetFocus
           DataGrid1.Col = 0
        Else
        End If
    ElseIf (KeyCode = 13 Or KeyCode = 9) And ((DataGrid1.Col = 1) Or (DataGrid1.Col = 7)) Then
        If Len(Trim(DataGrid1.Text)) < 1 Then
            DataCombo3.Visible = True
            DataCombo3.SetFocus
        Else
        End If
    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 5) Then
        If Len(Trim(DataGrid1.Text)) < 1 Then
            Saudacombo.Visible = True
            Saudacombo.SetFocus
        End If
    ElseIf (KeyCode = 13 Or KeyCode = 9) And DataGrid1.Col = 9 Then
        If Val(DataGrid1.Text & "") <= 0 Then
            MsgBox "Rate can not be zero.Please enter Rate.", vbCritical
            DataGrid1.Col = 3: DataGrid1.SetFocus
        End If
    ElseIf KeyCode = 27 Then
        KeyCode = 0
    End If
    DataGrid1.CurrentCellVisible = True
End If
End Sub

Private Sub EXCMB_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub EXCMB_Validate(Cancel As Boolean)
If excmb.BoundText = "" Then
    Cancel = True
    Exit Sub
End If
    Call ITEMLIST
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 112 Then
        Call lblnew_Click
    ElseIf KeyCode = 113 Then
        Call lbledit_Click
    ElseIf KeyCode = 114 Then
        Call LblDelete_Click
    ElseIf KeyCode = 115 Then
        Call lblsave_Click
    ElseIf KeyCode = 116 Then
        Call lblcancel_Click
    ElseIf KeyCode = 117 Then
        Call lblexit_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        On Error Resume Next
        If Me.ActiveControl.Name = "vcDTP1" Then
            SendKeys "{tab}"
        End If
    End If
End Sub
Private Sub Form_Load()
    Call CANCEL_REC
'----------

    'LblNew.Visible = False: LblEdit.Visible = False: LblCancel.Visible = False: LblDelete.Visible = False: LblSave.Visible = False: LblExit.Visible = False
    Set DataGrid1.DataSource = RECGRID: DataGrid1.Refresh
    Call ClearFormFn(SINGLECONTRACT)
    Frame1.Enabled = False
'--------
    LDataImport = 0
    DataCombo3.Top = Val(2055): DataCombo3.Left = Val(1080)
    MYSQL = "SELECT ITEMCODE,ITEMNAME AS ITEMNAME,Lot FROM ITEMMAST WHERE COMPCODE=" & MC_CODE & " ORDER BY ITEMCODE"
    Set REC = Nothing: Set REC = New ADODB.Recordset: REC.Open MYSQL, cnn, adOpenKeyset, adLockReadOnly
    If Not REC.EOF Then
        QACC_CHANGE = False: Set REC_ACCOUNT = Nothing: Set REC_ACCOUNT = New ADODB.Recordset
        REC_ACCOUNT.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & MC_CODE & " AND gcode in (12,14) ORDER BY NAME ", cnn, adOpenKeyset, adLockReadOnly
        If Not REC_ACCOUNT.EOF Then Set DataCombo3.RowSource = REC_ACCOUNT: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    Else
        Call Get_Selection(12)
    End If
    
    
End Sub

Private Sub Form_Paint()
    Me.BackColor = GETMAIN.BackColor
    If QACC_CHANGE Then
        QACC_CHANGE = False: Set REC_ACCOUNT = Nothing
        Set REC_ACCOUNT = New ADODB.Recordset
        REC_ACCOUNT.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & MC_CODE & " AND gcode in (12,14) ORDER BY NAME ", cnn, adOpenKeyset, adLockReadOnly
        If Not REC_ACCOUNT.EOF Then
            Set DataCombo3.RowSource = REC_ACCOUNT
            DataCombo3.BoundColumn = "AC_CODE"
            DataCombo3.ListField = "NAME"
        Else
            MsgBox "Please create customer account", vbInformation
            Call Get_Selection(12)
        End If
    End If
    If fb_press > 0 Then Call Get_Selection(fb_press)
End Sub



Private Sub headcmb_GotFocus()
    SendKeys "%{DOWN}"
End Sub


Private Sub headcmb_Validate(Cancel As Boolean)
If headcmb.BoundText = "" Then Cancel = True

End Sub

Private Sub partycmb_Validate(Cancel As Boolean)
If partycmb.BoundText <> "" Then
    RECGRID!Code = partycmb.BoundText
    RECGRID!Name = partycmb.Text
    DataGrid1.Columns(1).Locked = True
Else
    DataGrid1.Columns(1).Locked = False
End If
If ITEMCMB.BoundText <> "" Then
    RECGRID!ItemCode = ITEMCMB.BoundText
    RECGRID!ITEMName = ITEMCMB.Text
    DataGrid1.Columns(12).Locked = True
Else
    DataGrid1.Columns(12).Locked = False
End If

End Sub

Private Sub Saudacmb_GotFocus()
Call SaudaList
End Sub

Private Sub Saudacmb_LostFocus()
If MODIFY_REC(vcDTP1.Value, Saudacmb.BoundText, ITEMCMB.BoundText, Combo1.Text) Then
Else
    Combo1.SetFocus
End If

End Sub

Private Sub Saudacmb_Validate(Cancel As Boolean)
If Saudacmb.BoundText <> "" Then
    RECGRID!SAUDACODE = Saudacmb.BoundText
    RECGRID!SAUDANAME = Saudacmb.Text
    DataGrid1.Col = 0
    DataGrid1.SetFocus
    Saudacombo.Visible = False
    DataGrid1.Columns(12).Locked = True
    DataGrid1.SetFocus
End If

End Sub

Private Sub SAUDACOMBO_GotFocus()
    If ITEMCMB.BoundText <> "" Then
        MYSQL = "SELECT S.SAUDANAME,S.SAUDACODE FROM SAUDAMAST AS S WHERE S.COMPCODE =" & MC_CODE & " AND S.ITEMCODE='" & ITEMCMB.BoundText & "' AND MATURITY>= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,MATURITY"
    Else
        MYSQL = "SELECT SAUDANAME,SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & MC_CODE & "  AND MATURITY>= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,MATURITY"
    End If
    Set SAUDACMBREC = Nothing
    Set SAUDACMBREC = New ADODB.Recordset
    SAUDACMBREC.Open MYSQL, cnn
    If Not SAUDACMBREC.EOF Then
        Set Saudacombo.RowSource = SAUDACMBREC
        Saudacombo.ListField = "SAUDANAME"
        Saudacombo.BoundColumn = "SAUDACODE"
    End If
    SendKeys "%{DOWN}"
    Saudacombo.Left = DataGrid1.Columns(6).Left
    Saudacombo.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
    SendKeys "%{DOWN}"
    SendKeys "%{DOWN}"
End Sub

Private Sub SAUDACOMBO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Saudacombo.BoundText <> "" Then
        RECGRID!SAUDACODE = Saudacombo.BoundText
        RECGRID!SAUDANAME = Saudacombo.Text
        MYSQL = "SELECT I.ITEMCODE,I.ITEMNAME,I.LOT FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE= '" & MC_CODE & "' AND S.SAUDACODE='" & Saudacombo.BoundText & "'AND S.ITEMCODE=I.ITEMCODE AND I.COMPCODE=S.COMPCODE "
        Set TempRec = Nothing
        Set TempRec = New ADODB.Recordset
        TempRec.Open MYSQL, cnn
        If Not TempRec.EOF Then
            RECGRID!ItemCode = TempRec!ItemCode
            RECGRID!ITEMName = TempRec!ITEMName
            RECGRID!LOT = TempRec!LOT
        End If
        DataGrid1.Col = 6
        DataGrid1.SetFocus
        Saudacombo.Visible = False
    End If
End If

End Sub

Private Sub SAUDACOMBO_Validate(Cancel As Boolean)
If Saudacombo.Visible = True Then
    Cancel = True
End If
End Sub

Private Sub Text3_GotFocus()
    FLOWDIR = 0: Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Sub RECSET()
    Set TempParty = Nothing
    Set TempParty = New ADODB.Recordset
    TempParty.Fields.Append "ACODE", adVarChar, 6, adFldIsNullable
    TempParty.Open , , adOpenKeyset, adLockBatchOptimistic
    Set TempSauda = Nothing
    Set TempSauda = New ADODB.Recordset
    TempSauda.Fields.Append "SAUDACODE", adVarChar, 50, adFldIsNullable
    TempSauda.Open , , adOpenKeyset, adLockBatchOptimistic
    Set TempItem = Nothing
    Set TempItem = New ADODB.Recordset
    TempItem.Fields.Append "ITEMCODE", adVarChar, 50, adFldIsNullable
    TempItem.Open , , adOpenKeyset, adLockBatchOptimistic
    
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SRNO", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BUYSELL", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "CODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "NAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "QNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "CONCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "CONNAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "RATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LOT", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "ITEMCODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "ITEMNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SAUDACODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SAUDANAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "RATE1", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "LInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "RInvNo", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "DImport", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "CONTIME", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "UserId", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "LCLCODE", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "RCLCODE", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "ORDER_NO", adVarChar, 30, adFldIsNullable
    RECGRID.Fields.Append "TRADE_NO", adVarChar, 30, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
    
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    Text3.Text = Format(Text3.Text, "0.00")
End Sub
Sub DELETE_VOUCHER(VOU_NO As String)
    cnn.Execute "DELETE FROM VCHAMT  WHERE COMPCODE=" & MC_CODE & " AND VOU_NO='" & VOU_NO & "'"
    cnn.Execute "DELETE FROM VOUCHER WHERE COMPCODE=" & MC_CODE & " AND VOU_NO='" & VOU_NO & "'"
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
'    Static OLDVAL As Integer
'
'    Select Case ColIndex
'    Case 0
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "SRNO DESC"
'        Else
'            RECGRID.Sort = "SRNO"
'        End If
'
'    Case 1
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "BCODE DESC"
'        Else
'            RECGRID.Sort = "BCODE"
'        End If
'
'    Case 2
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "BNAME DESC"
'        Else
'            RECGRID.Sort = "BNAME"
'        End If
'
'    Case 3
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "BQNTY DESC"
'        Else
'            RECGRID.Sort = "BQNTY"
'        End If
'
'    Case 4
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "BRATE DESC"
'        Else
'            RECGRID.Sort = "BRATE"
'        End If
'    Case 5
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "SCODE DESC"
'        Else
'            RECGRID.Sort = "SCODE"
'        End If
'
'    Case 6
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "SNAME DESC"
'        Else
'            RECGRID.Sort = "SNAME"
'        End If
'
'    Case 7
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "SQNTY DESC"
'        Else
'            RECGRID.Sort = "SQNTY"
'        End If
'
'    Case 8
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "SRATE DESC"
'        Else
'            RECGRID.Sort = "SRATE"
'        End If
'    Case 13
'        If OLDVAL = -1 Then
'            RECGRID.Sort = "UserId DESC"
'        Else
'            RECGRID.Sort = "UserId"
'        End If
'    End Select
'
'    If OLDVAL = -1 Then
'        Call VISIBLE_IMAGE(0)
'    Else
'        Call VISIBLE_IMAGE(1)
'    End If
'
'    If OLDVAL = ColIndex Then
'        OLDVAL = -1
'    Else
'        OLDVAL = ColIndex
'    End If
'    Image1(0).Left = DataGrid1.Left + DataGrid1.Columns(ColIndex).Left + (DataGrid1.Columns(ColIndex).Width) / 2
'    Image1(1).Left = DataGrid1.Left + DataGrid1.Columns(ColIndex).Left + (DataGrid1.Columns(ColIndex).Width) / 2
End Sub
Sub VISIBLE_IMAGE(SORT_ORDER As Byte)
    If SORT_ORDER = 1 Then
        Image1(0).Visible = False
        Image1(1).Visible = True
    Else
        Image1(0).Visible = True
        Image1(1).Visible = False
    End If
End Sub
Function GetCloseRate() As Boolean
     Set REC_SAUDA = Nothing: Set REC_SAUDA = New ADODB.Recordset
     REC_SAUDA.Open "SELECT * FROM SAUDAMAST WHERE COMPCODE=" & MC_CODE & " AND SAUDACODE='" & Saudacmb.BoundText & "' AND MATURITY>= '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", cnn, adOpenForwardOnly, adLockReadOnly
     If REC_SAUDA.EOF Then
         MsgBox "Invalid SAUDA code.", vbExclamation, "Error"
         GetCloseRate = False
     Else
         GetCloseRate = True
         Set REC_CloRate = Nothing: Set REC_CloRate = New ADODB.Recordset
         REC_CloRate.Open "SELECT CloseRate,DataImport FROM CTR_R WHERE COMPCODE=" & MC_CODE & " AND SAUDA='" & Text2.Text & "' AND CONDATE  =  '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", cnn, adOpenForwardOnly, adLockReadOnly
         If Not REC_CloRate.EOF Then
            Text3.Text = Format(REC_CloRate!CLOSERATE, "0.00")
        End If
         Text2.Text = REC_SAUDA!SAUDACODE
         DataCombo1.BoundText = CStr(Text2.Text)
         ITEMCMB.BoundText = REC_SAUDA!ItemCode
    End If
End Function
Public Sub lblcancel_Click()
    Call GETMAIN.ButtonClick(5)
    LblNew.Visible = True: LblEdit.Visible = True: LblDelete.Visible = True
    LblSave.Visible = False: LblCancel.Visible = False
End Sub
Public Sub lblsave_Click()
    Call GETMAIN.ButtonClick(4)
    LblNew.Visible = True: LblEdit.Visible = True: LblDelete.Visible = True
    LblSave.Visible = False: LblCancel.Visible = False
End Sub
Public Sub lblexit_Click()
    Call GETMAIN.ButtonClick(6)
End Sub
Public Sub lblnew_Click()
    Call GETMAIN.ButtonClick(1)
    LblNew.Visible = False: LblEdit.Visible = False: LblDelete.Visible = False
    LblSave.Visible = True: LblCancel.Visible = True
End Sub
Public Sub lbledit_Click()
    Call GETMAIN.ButtonClick(2)
    LblNew.Visible = False: LblEdit.Visible = False: LblDelete.Visible = False
    LblSave.Visible = True: LblCancel.Visible = True
End Sub
Public Sub LblDelete_Click()
    Call GETMAIN.ButtonClick(3)
    LblNew.Visible = False: LblEdit.Visible = False: LblDelete.Visible = False
    LblSave.Visible = True: LblCancel.Visible = True
End Sub

