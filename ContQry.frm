VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form ContQry 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11655
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   6240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   38234.8441203704
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   4560
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   38234.8440856481
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   375
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Date from"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "to"
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   300
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "Date"
         Caption         =   "Date"
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
         DataField       =   "Sauda"
         Caption         =   "Sauda"
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
         DataField       =   "BTradeNo"
         Caption         =   "BTradeNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Buyer"
         Caption         =   "Buyer"
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
      BeginProperty Column04 
         DataField       =   "BQty"
         Caption         =   "BQty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "BRate"
         Caption         =   "BRate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "STradeNo"
         Caption         =   "STradeNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Seller"
         Caption         =   "Seller"
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
         DataField       =   "Sqty"
         Caption         =   "SQty"
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
         DataField       =   "SRate"
         Caption         =   "SRate"
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
            Locked          =   -1  'True
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adoacc 
      Height          =   330
      Left            =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
   Begin VB.Frame Frame2 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   11655
   End
End
Attribute VB_Name = "ContQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RECGRID As ADODB.Recordset
Sub ContList()
    Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
    MYSQL = "SELECT CTR_D.*, A.NAME AS NAME FROM CTR_D, ACCOUNTD AS A WHERE CTR_D.compcode=" & MC_CODE & " AND CTR_D.compcode=A.compcode and CTR_D.PARTY=A.AC_CODE AND CTR_D.CONSNO=" & Val(Adodc1.Recordset!CONSNO) & " ORDER BY CONNO, CONTYPE"
    GeneralRec.Open MYSQL, CNN, adOpenForwardOnly, adLockReadOnly
    Call RECSET
    Do While Not GeneralRec.EOF
        RECGRID.AddNew
                RECGRID!SRNO = GeneralRec!CONNO  'RECGRID.AbsolutePosition
                RECGRID!BCODE = GeneralRec!PARTY & ""
                RECGRID!BNAME = GeneralRec!Name
                RECGRID!BQNTY = GeneralRec!QTY
                RECGRID!BRate = GeneralRec!Rate
                RECGRID!LInvNo = Val(GeneralRec!INVNO & "")
                RECGRID!DIMPORT = IIf(GeneralRec!DATAIMPORT & "", 1, 0)
                RECGRID!CONTIME = IIf(IsNull(GeneralRec!CONTIME), Time, GeneralRec!CONTIME)
                
                GeneralRec.MoveNext
                RECGRID!scode = GeneralRec!PARTY & ""
                RECGRID!SNAME = GeneralRec!Name
                RECGRID!SQNTY = GeneralRec!QTY
                RECGRID!SRate = GeneralRec!Rate
                RECGRID!RInvNo = Val(GeneralRec!INVNO & "")
        RECGRID.Update
        GeneralRec.MoveNext
    Loop
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
End Sub
Sub RECSET()
    Set RECGRID = Nothing: Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "Date", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "BNAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "BQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "SNAME", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "SQNTY", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "CONTIME", adVarChar, 15, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
End Sub
