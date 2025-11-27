VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form YrUpdt 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11895
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Year Updation"
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
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   11535
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5610
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9895
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "YrUpdt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Combo1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5655
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   8775
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
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
            Height          =   375
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Enter City"
            Top             =   5160
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   615
            Left            =   120
            TabIndex        =   27
            Top             =   3120
            Width           =   8415
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
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
               Left            =   6600
               TabIndex        =   30
               Top             =   120
               Value           =   1  'Checked
               Width           =   735
            End
            Begin vcDateTimePicker.vcDTP vcDTP1 
               Height          =   375
               Left            =   1800
               TabIndex        =   28
               Top             =   120
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
               Value           =   39872.8817013889
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transfer Contract Opening "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   480
               Left            =   3600
               TabIndex        =   31
               Top             =   120
               Width           =   2835
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Settlement Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   240
               Index           =   1
               Left            =   120
               TabIndex        =   29
               Top             =   195
               Width           =   1995
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   1695
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   8415
            Begin VB.TextBox Text5 
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
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   4440
               Locked          =   -1  'True
               TabIndex        =   35
               TabStop         =   0   'False
               ToolTipText     =   "Enter City"
               Top             =   1080
               Width           =   3855
            End
            Begin VB.CheckBox Check4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Create New DataBase"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   33
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox Text4 
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
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "Enter City"
               Top             =   600
               Width           =   6495
            End
            Begin vcDateTimePicker.vcDTP vcDTP2 
               Height          =   360
               Left            =   1800
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   635
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   37680
            End
            Begin vcDateTimePicker.vcDTP vcDTP4 
               Height          =   360
               Left            =   6600
               TabIndex        =   22
               Top             =   120
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   37680
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "To Database"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   240
               Left            =   3000
               TabIndex        =   34
               Top             =   1200
               Width           =   1695
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "From Database"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   240
               Left            =   120
               TabIndex        =   26
               Top             =   660
               Width           =   1695
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fin Tear Begin"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   240
               Left            =   120
               TabIndex        =   24
               Top             =   195
               Width           =   1875
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Fin Year End"
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
               Height          =   240
               Left            =   5160
               TabIndex        =   23
               Top             =   195
               Width           =   1335
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   8415
            Begin VB.TextBox Text3 
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
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox Text1 
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
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   17
               TabStop         =   0   'False
               ToolTipText     =   "Enter City"
               Top             =   120
               Width           =   5655
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Company"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   240
               Left            =   240
               TabIndex        =   18
               Top             =   195
               Width           =   900
               WordWrap        =   -1  'True
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
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
            Left            =   4320
            TabIndex        =   15
            Top             =   5040
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Ok"
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
            Left            =   3120
            TabIndex        =   14
            Top             =   5040
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
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
            Left            =   240
            TabIndex        =   36
            Top             =   4080
            Width           =   8535
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   360
         TabIndex        =   11
         Top             =   3840
         Visible         =   0   'False
         Width           =   3255
         Begin MSDataListLib.DataCombo SDtDb 
            Height          =   360
            Left            =   480
            TabIndex        =   12
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ForeColor       =   16711680
            Text            =   "DataCombo1"
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
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Select All"
         Height          =   195
         Left            =   3240
         TabIndex        =   10
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   4440
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   4095
         Begin MSComctlLib.ListView ListView2 
            Height          =   2415
            Left            =   -720
            TabIndex        =   9
            ToolTipText     =   "Press : F2 to select all, F3 to unselect, F4 to select item specific."
            Top             =   480
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4260
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   6722
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Lot"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "exchange"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "SDutyType"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "SDutyRate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "SDutyPer"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "MTYPE"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "YrUpdt.frx":001C
         Left            =   6720
         List            =   "YrUpdt.frx":0026
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   1
         Top             =   2040
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Packup Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Party Opening Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4800
         TabIndex        =   3
         Top             =   1920
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "YrUpdt.frx":0046
      Height          =   5580
      Left            =   120
      TabIndex        =   2
      Top             =   795
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9843
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   16761024
      ForeColor       =   0
      ListField       =   "Author"
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   5820
      Left            =   2715
      Top             =   600
      Width           =   8925
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   5820
      Left            =   75
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "YrUpdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldCnn As ADODB.Connection:     Dim NewCnn As ADODB.Connection
Dim RecRate As ADODB.Recordset:     Dim ServerString As String
Dim RecSauda As ADODB.Recordset:    Dim MRec As ADODB.Recordset
Dim CompRec As ADODB.Recordset:     Dim AccRec As ADODB.Recordset
Dim MCnn  As ADODB.Connection
Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        Check1.Caption = "Yes"
    Else
        Check1.Caption = "No"
    End If
End Sub
Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        Check2.Caption = "Yes"
    Else
        Check2.Caption = "No"
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Text5.text = ""
        Text5.Locked = False
    Else
        Text5.text = ""
        Text5.Locked = True
    End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex = 1 Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
End Sub
Private Sub Command1_Click()
    
    On Error GoTo err1
    
    'create database checkbox is true
    If Check4.Value = 1 Then
    
        Set OldCnn = Nothing
        Set NewCnn = Nothing
        Set Cnn = Nothing
                
        Dim CnnString As String:
        Dim DBname As String
        Dim DBPath As String
        Dim DBBkpfilepath As String
        Dim MasterConn As ADODB.Connection
        Dim DBRec As ADODB.Recordset
        
        DBname = Left$(MServer, (InStr(1, MServer, "database") + 8))
        DBname = Replace(MServer, DBname, "")
    
        If Text5.text = "" Then
            DBname = DBname + "_N"
        Else
            DBname = Text5.text
        End If
                    
        ServerString = MServer
        ServerString = Left$(MServer, (InStr(1, MServer, "database") + 8)) & "MASTER"
                       
        Set MasterConn = Nothing:    Set MasterConn = New ADODB.Connection
        MasterConn.Open ServerString
        DoEvents
                
        mysql = "SELECT NAME FROM sysdatabases WHERE NAME='" & DBname & "'"
        Set DBRec = Nothing: Set DBRec = New ADODB.Recordset: DBRec.Open mysql, MasterConn, adOpenKeyset, adLockReadOnly
        If DBRec.EOF Then
            
        '1) creae NEW DB ----------------
            CnnString = ServerString
            Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
            MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
            MCnn.Open
            MCnn.CommandTimeout = 6000
            mysql = "EXECUTE DbGenerate '" & DBname & "'"
            MCnn.Execute mysql
            Set MCnn = Nothing
        'creae NEW DB ----------------end
        
                
        '2) backup existing DB----------------
            Call DataBackUp_OnLogOff("N")
        'backup existing DB----------------end
        
        
        '3) restore existng DB on NEW DB------------------
            DoEvents
            CnnString = ServerString
            Set MCnn = Nothing: Set MCnn = New ADODB.Connection: MCnn.ConnectionString = CnnString
            MCnn.Mode = adModeShareExclusive: MCnn.IsolationLevel = adXactIsolated: MCnn.CursorLocation = adUseClient
            MCnn.Open
        
            mysql = "EXECUTE RESTOREDB '" & DBname & "','" & BkpDevPath & "'"
            MCnn.Execute mysql
            MCnn.close
            Set MCnn = Nothing
            DoEvents
        'restore existng DB on NEW DB------------------end
        
        '4) Connect to NEW DB-------------
            MServer = Left$(MServer, (InStr(1, MServer, "database") + 8)) & DBname
            Set Cnn = Nothing: Set Cnn = New ADODB.Connection: Cnn.ConnectionString = MServer
            Cnn.Mode = adModeShareExclusive: Cnn.IsolationLevel = adXactIsolated: Cnn.CursorLocation = adUseClient
            Cnn.Open
            Cnn.CommandTimeout = 300
            
            Set OldCnn = Cnn
            Set NewCnn = OldCnn
        'Connect to NEW DB-------------end
        
        '5) change WinSql32.ini file for DB name--------------
            Dim LFileSystemObject As FileSystemObject
            Dim EFileSystemObject As Scripting.FileSystemObject
            Set EFileSystemObject = CreateObject("Scripting.FileSystemObject")
            Dim EFile As Variant
            EFileSystemObject.CreateTextFile (App.Path & "\WinSql32-1.ini")
            Set LFileSystemObject = CreateObject("Scripting.FileSystemObject")
            Open App.Path & "\WINSQL32.INI" For Input As #1
                Line Input #1, MServer
                If InStr(1, MServer, "{SQL Server};") > 0 Then
                    MServer = Left$(MServer, (InStr(1, MServer, "database") + 8)) & DBname
                End If
                Set EFile = EFileSystemObject.OpenTextFile(App.Path & "\WinSql32-1.ini", ForWriting)
                EFile.WriteLine MServer
            Close #1
            Set EFile = Nothing
            LFileSystemObject.DeleteFile App.Path & "\WinSql32.ini"
            LFileSystemObject.MoveFile App.Path & "\WinSql32-1.ini", App.Path & "\WinSql32.ini"
        'change WinSql32.ini file for DB name--------------end
        Else
            MsgBox "Database " & DBname & " already exists!!!, please change database name and try again.", vbCritical
            Exit Sub
        End If
            
    End If
        
err1:
    If err.Number <> 0 Then
        MsgBox err.Description
    Else
        Call Yearupdation
    End If
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub DataList1_Click()
    Text1.text = DataList1.text
    If InStr(1, MServer, "database") > 0 Then Text4.text = Mid(MServer, (InStr(1, MServer, "database") + 9), Len(MServer))
    Call COMPANY_ACCESS
End Sub
Private Sub Form_Load()
'    Call ClearFormFn(YrUpdt)
    vcDTP1.Value = GFinBegin
    Label11.Caption = MFormat
    Call Get_Selection(12)
    
    Set OldCnn = Cnn
    Set NewCnn = OldCnn
    Set CompRec = Nothing:    Set CompRec = New ADODB.Recordset
    mysql = "SELECT COMPCODE,NAME FROM COMPANY ORDER BY COMPCODE"
    CompRec.Open mysql, OldCnn, adOpenStatic, adLockReadOnly
    If Not CompRec.EOF Then
        Set DataList1.RowSource = CompRec
        DataList1.ListField = "Name"
        DataList1.BoundColumn = "COMPCODE"
    End If
    Combo1.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Private Sub SDtDb_GotFocus()
    Sendkeys "%{down}"
'    Call LSendKeys_Down
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then Text2.Locked = Not Text2.Locked
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then Text4.Locked = Not Text4.Locked
End Sub
Sub COMPANY_ACCESS()
    On Error GoTo Error1
    CompRec.MoveFirst
    CompRec.Find "COMPCODE=" & DataList1.BoundText & "", , adSearchForward
    If Not CompRec.EOF Then
        Text1.text = CompRec!NAME & ""
        Text3.text = CompRec!CompCode & ""
        vcDTP1.SetFocus
    Else
        MsgBox "Please Select Company", vbInformation
        DataList1.SetFocus
    End If
    Exit Sub
Error1:
    
    MsgBox err.Description, vbCritical, err.HelpFile
End Sub
Sub Yearupdation()
Dim BalRec As ADODB.Recordset:  Dim AccRec As ADODB.Recordset:  Dim TRec As ADODB.Recordset
Dim ldate As Date:              Dim LNetBal As Double:          Dim LBal As Double
Dim LAC_CODE As String:         Dim MConSno As Long:            Dim MConNo As Long
Dim VchNo As String:            Dim MSCode As String:           Dim MPartyCode As String
Dim MICode As String:           Dim MContCode  As String:       Dim MCLoseRate  As Double
Dim MSauda As String:           Dim MCONTYPE As String * 1:     Dim MQnty As Double
Dim MASAmt As Double:           Dim MABAmt As Double:           Dim RecRate As ADODB.Recordset
Dim LCalval As Double:          Dim LTime As String:            Dim MItemCode As String
Dim MInstType  As String:       Dim MOptType  As String:        Dim MStrike  As Double
Dim LDrive As String:           Dim LExCode As String:          Dim CnnString As String
Dim LFolderName  As String:     Dim LSaudaID As Long:        Dim ServerString As String
Dim MCompCode As Integer:       Dim MRpt_Path As String
Dim MLFinBeg As Date:           Dim MLFinEnd  As Date:          Dim MDPath  As String
Dim MSysLockDt As Date:         Dim LExID As Integer:           Dim LItemID As Integer
Dim LACCID As Long
    
If Combo1.ListIndex = 0 Then ' datewise
    On Error GoTo Error1
        If LenB(Text1.text) = 0 Then MsgBox "Please Select Company. ", vbCritical: DataList1.SetFocus: Exit Sub
        GCompCode = DataList1.BoundText
        Cnn.BeginTrans
        Cnn.CommandTimeout = 6000
        '*********
        'Transfer P  party opening FROM   MFinBeg   to   SELECTed   Settlement Date
        If Check1.Value = vbChecked Then
            Set AccRec = Nothing: Set AccRec = New ADODB.Recordset
            mysql = "SELECT AC_CODE,NAME,OP_BAL FROM ACCOUNTM WHERE COMPCODE=" & GCompCode & "  ORDER BY AC_CODE"
            AccRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            If Not AccRec.EOF Then
                'Set AccRec.ActiveConnection = Nothing
                Do While Not AccRec.EOF
                    DoEvents
                    LNetBal = 0
                    Label10.Caption = "Updating Account Balances" & AccRec!AC_CODE
                    DoEvents
                    LBal = Val(AccRec!OP_BAL & "")
                    LAC_CODE = AccRec!AC_CODE
                    LNetBal = Net_DrCr(LAC_CODE, CStr(vcDTP1.Value + 1))
                    LBal = LBal + LNetBal
                    Cnn.Execute "UPDATE ACCOUNTM SET OP_BAL = " & LBal & " WHERE COMPCODE = " & GCompCode & " AND AC_CODE = '" & AccRec!AC_CODE & "' "
                    AccRec.MoveNext
                Loop
                If MFormat = "PackUp Data" Then
                    DoEvents
                    Label10.Caption = "Deleteing Vouchers"
                    DoEvents
                    ldate = DateValue(vcDTP1.Value)
                    mysql = "DELETE FROM VOUCHER WHERE COMPCODE =" & GCompCode & " AND VOU_DT<='" & Format(ldate, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    mysql = "DELETE FROM VCHAMT WHERE COMPCODE =" & GCompCode & " AND VOU_DT<='" & Format(ldate, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    mysql = "DELETE FROM VOUCHER WHERE COMPCODE =" & GCompCode & " AND (VOU_TYPE = 'S' OR VOU_TYPE = 'B' OR VOU_TYPE = 'H')"
                    Cnn.Execute mysql
                    mysql = "DELETE FROM VCHAMT WHERE COMPCODE =" & GCompCode & " AND (VOU_TYPE = 'S' OR VOU_TYPE = 'B' OR VOU_TYPE = 'H')"
                    Cnn.Execute mysql
                End If
            End If
        End If
        'Transfer contract opening
        If Check2.Value = vbChecked Then
            Set BalRec = Nothing: Set BalRec = New ADODB.Recordset
            BalRec.Fields.Append "AC_CODE", adVarChar, 15, adFldIsNullable
            BalRec.Fields.Append "CONCODE", adVarChar, 15, adFldIsNullable
            BalRec.Fields.Append "SAUDA", adVarChar, 50, adFldIsNullable
            BalRec.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
            BalRec.Fields.Append "QTY", adDouble, , adFldIsNullable
            BalRec.Fields.Append "RATE", adDouble, , adFldIsNullable
            BalRec.Fields.Append "ConNo", adDouble, , adFldIsNullable
            BalRec.Fields.Append "EXCODE", adVarChar, 10, adFldIsNullable
            BalRec.Fields.Append "CALVAL", adDouble, , adFldIsNullable
            BalRec.Fields.Append "INSTTYPE", adVarChar, 3, adFldIsNullable
            BalRec.Fields.Append "OPTTYPE", adVarChar, 2, adFldIsNullable
            BalRec.Fields.Append "STRIKE", adDouble, , adFldIsNullable
            BalRec.Fields.Append "SAUDAID", adInteger, , adFldIsNullable
            BalRec.Fields.Append "ITEMID", adInteger, , adFldIsNullable
            BalRec.Fields.Append "EXID", adInteger, , adFldIsNullable
            BalRec.Fields.Append "ACCID", adInteger, , adFldIsNullable
            
            BalRec.Open , , adOpenKeyset, adLockOptimistic
            mysql = "SELECT EM.EXID,IM.ITEMID,EM.EXCODE,EM.LOTWISE,SU.SAUDAID,SU.SAUDACODE,SU.ITEMCODE,SU.INSTTYPE,SU.OPTTYPE,SU.STRIKEPRICE,P.NAME,P.ACCID,P.AC_CODE,EM.CONTRACTACC,"
            mysql = mysql & " SU.TRADEABLELOT,IM.LOT,CD.CALVAL, SUM(CASE CONTYPE WHEN 'B' THEN CD.QTY WHEN 'S' THEN CD.QTY *-1 END ) AS SUMOFQTY "
            mysql = mysql & " FROM CTR_D AS CD,SAUDAMAST AS SU,ACCOUNTD AS P,ITEMMAST AS IM,EXMAST AS EM "
            mysql = mysql & " WHERE CD.COMPCODE=" & DataList1.BoundText & " AND  CD.SAUDAID=SU.SAUDAID  "
            mysql = mysql & " AND CD.ACCID =P.ACCID "
            mysql = mysql & " AND CD.CONDATE <='" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND SU.MATURITY >'" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' "
            mysql = mysql & " AND SU.ITEMID =IM.ITEMID AND EM.EXID=IM.EXID   "
            mysql = mysql & " GROUP BY EM.EXID,IM.ITEMID,EM.EXCODE,EM.LOTWISE,SU.SAUDAID,SU.SAUDACODE,SU.ITEMCODE,SU.INSTTYPE,SU.OPTTYPE,SU.STRIKEPRICE,P.NAME,P.ACCID,P.AC_CODE,EM.CONTRACTACC,SU.TRADEABLELOT,IM.LOT,CD.CALVAL "
            mysql = mysql & " ORDER BY EM.EXCODE,SU.SAUDACODE,P.NAME"
            Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            'closing rate
            mysql = "SELECT SAUDAID,SAUDA,CLOSERATE FROM CTR_R WHERE COMPCODE =" & DataList1.BoundText & " AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'"
            Set RecRate = Nothing
            Set RecRate = New ADODB.Recordset
            RecRate.Open mysql, Cnn, adOpenStatic, adLockReadOnly
            MConSno = 0
            VchNo = vbNullString
            MConNo = 0
            Do While Not TRec.EOF
                MSCode = TRec!saudacode:                MICode = TRec!ITEMCODE
                MContCode = TRec!CONTRACTACC:           MCLoseRate = 0
                MInstType = TRec!INSTTYPE:              MOptType = TRec!OPTTYPE
                MStrike = TRec!STRIKEPRICE:             LSaudaID = TRec!SAUDAID
                LExID = TRec!EXID
                LItemID = TRec!itemid
                
                LACCID = TRec!ACCID
                DoEvents
                Label10.Caption = "Updating Standing " & TRec!NAME & " " & MSCode & ""
                DoEvents
                LCalval = TRec!CALVAL
                Do While LSaudaID = TRec!SAUDAID
                    MPartyCode = TRec!AC_CODE
                    MQnty = 0: MASAmt = 0: MABAmt = 0
                    If Round(TRec!SUMOFQTY, 2) <> 0 Then
                        If RecRate.EOF Then
                            MsgBox "No Closing Rates Found Retry "
                            CNNERR = False
                            Cnn.RollbackTrans
                            Exit Sub
                        End If
                        RecRate.MoveFirst
                        RecRate.Find "SAUDAID=" & LSaudaID & "", , adSearchForward
                        If Not RecRate.EOF Then
                            MCLoseRate = Format(RecRate!CLOSERATE, "0.0000")
                        Else
                            MsgBox "Please Enter Closing Rate for " & vcDTP1.Value & "  for " & MSCode & " "
                            Cnn.RollbackTrans
                            Exit Sub
                        End If
                
                        MQnty = TRec!SUMOFQTY
                        MConNo = MConNo + 1
                        BalRec.AddNew
                        BalRec!AC_CODE = MPartyCode:    BalRec!CONCODE = MContCode
                        BalRec!Sauda = MSCode:          BalRec!ITEMCODE = MICode
                        BalRec!QTY = MQnty:             BalRec!Rate = Val(MCLoseRate)
                        BalRec!CONNO = MConNo:          BalRec!excode = TRec!excode
                        BalRec!CALVAL = LCalval:        BalRec!INSTTYPE = MInstType
                        BalRec!OPTTYPE = MOptType:      BalRec!strike = MStrike
                        BalRec!SAUDAID = LSaudaID:      BalRec!itemid = LItemID
                        BalRec!EXID = LExID
                        BalRec!ACCID = LExID
                        BalRec.Update
                    End If
                    TRec.MoveNext
                    If TRec.EOF Then Exit Do
                Loop
                If TRec.EOF Then Exit Do
            Loop
            DoEvents
            Label10.Caption = "Deleteing Old Data "
            DoEvents
            Cnn.CommandTimeout = 6000
            mysql = "EXEC Get_PACKDATA " & GCompCode & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
            Cnn.Execute mysql
            
            If BalRec.RecordCount > 0 Then
                BalRec.MoveFirst
                Do While Not BalRec.EOF
                    MSauda = BalRec!Sauda
                    LSaudaID = BalRec!SAUDAID
                    MItemCode = BalRec!ITEMCODE
                    LItemID = BalRec!itemid
                    LExID = BalRec!EXID
                    LExCode = BalRec!excode
                    LACCID = BalRec!ACCID
                    MConSno = Get_ConSNo(vcDTP1.Value, MSauda, MItemCode, LExCode, LSaudaID, LItemID, LExID)
                    If BalRec!QTY > 0 Then
                        MCONTYPE = "B"
                    Else
                        MCONTYPE = "S"
                    End If
                    LTime = CStr(Time)
                    Call Add_To_Ctr_D(MCONTYPE, BalRec!AC_CODE, MConSno, vcDTP1.Value, BalRec!CONNO, BalRec!Sauda, BalRec!ITEMCODE, BalRec!AC_CODE, Abs(BalRec!QTY), BalRec!Rate, BalRec!CONCODE, LTime, "", "", Trim$(Str(BalRec!CONNO)), BalRec!excode, BalRec!CALVAL, 1, vbNullString, BalRec!INSTTYPE, BalRec!OPTTYPE, BalRec!strike, "0", "N", LExID, LItemID, LSaudaID)
                    'Call Add_To_Ctr_D2(LConType, LClient, LSConSno, LSCondate, LConNo, MSaudaCode, LItemCode, MParty, MQty, MRate, MConRate, LExCont, LContime, LOrdNo, vbNullString, LOConNo, LExCode, LCalval, LPDataImport, vbNullString, LSInstType, LSOptType, LSStrike, Left$(TxtFileType.text, 2), LBrokFlag, LExID, LItemID, LSaudaID)
                    BalRec.MoveNext
                Loop
            End If
        End If
        mysql = "UPDATE CTR_D SET PATTAN ='O' WHERE COMPCODE =" & GCompCode & "  AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
        Cnn.Execute mysql
        mysql = "UPDATE CTR_M SET PATTAN ='O' WHERE COMPCODE =" & GCompCode & "  AND CONDATE ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'"
        Cnn.Execute mysql
        
        'Delete Data FROM current database
        Cnn.CommitTrans: CNNERR = False
        Call Update_Charges(vbNullString, vbNullString, vbNullString, vbNullString, GFinBegin, GFinEnd, True)
        Cnn.BeginTrans: CNNERR = True
        DoEvents
        GETMAIN.Label1.Caption = "Bill Generation"
            
        If BILL_GENERATION(GFinBegin, GFinEnd, vbNullString, vbNullString, vbNullString) Then
            Cnn.CommitTrans: CNNERR = False
        Else
            Cnn.RollbackTrans: CNNERR = False
        End If
        Cnn.CommandTimeout = 60
        'Call Chk_Billing
            '*****************
            On Error GoTo Error1
            Open App.Path & "\WINSQL32.INI" For Input As #1
                Line Input #1, MServer
            Close #1
        If Mid(MServer, 1, 1) = "d" Then
        Else
            MServer = Mid(MServer, 2, Len(MServer) - 2)
        End If
        Set Cnn = Nothing: Set Cnn = New ADODB.Connection: Cnn.ConnectionString = MServer
        Cnn.Mode = adModeShareExclusive: Cnn.IsolationLevel = adXactIsolated: Cnn.CursorLocation = adUseClient
        Cnn.Open
        If MFormat = "PackUp Data" Then
            MsgBox "Data Packed Successfully"
        Else
            MsgBox "Year Updated Successfully"
        End If
End If
Exit Sub
Error1:
    If err.Number <> 0 Then MsgBox err.Description
    If CNNERR = True Then
        Cnn.RollbackTrans
    End If

End Sub
Private Sub Check3_Click()
    Dim I As Integer
    For I = 1 To ListView2.ListItems.Count
        If Check3.Value = 1 Then
            ListView2.ListItems.Item(I).Checked = True
        Else
            ListView2.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
