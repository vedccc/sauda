VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbrok 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   14040
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   240
      TabIndex        =   41
      Top             =   0
      Width           =   1695
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000013&
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
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   15255
      Begin VB.Frame Frame7 
         BackColor       =   &H00400000&
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
         Height          =   615
         Left            =   5520
         TabIndex        =   39
         Top             =   120
         Width           =   3255
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Brokerage Setup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "frmbrok.frx":0000
      Left            =   8160
      List            =   "frmbrok.frx":000A
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
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
      Height          =   3495
      Left            =   360
      TabIndex        =   18
      Top             =   1320
      Width           =   14655
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Only Traded Commodities"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   38
         Top             =   600
         Width           =   4695
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   8640
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   225
         Left            =   5640
         TabIndex        =   35
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   225
         Left            =   2520
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Branch wise"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   6840
         TabIndex        =   25
         Top             =   120
         Width           =   2895
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000011&
            BackStyle       =   0  'Transparent
            Caption         =   "Item List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   915
            TabIndex        =   28
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   120
         Width           =   3015
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000011&
            BackStyle       =   0  'Transparent
            Caption         =   "Exhange List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   840
            TabIndex        =   27
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   3495
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000011&
            BackStyle       =   0  'Transparent
            Caption         =   "Party List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   1245
            TabIndex        =   26
            Top             =   0
            Width           =   1005
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Index           =   0
         Left            =   9840
         TabIndex        =   20
         Top             =   960
         Width           =   4695
         Begin VB.ComboBox Combo6 
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
            ItemData        =   "frmbrok.frx":0028
            Left            =   120
            List            =   "frmbrok.frx":0035
            TabIndex        =   37
            Top             =   840
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   4455
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   360
               Left            =   2280
               TabIndex        =   4
               Top             =   120
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   635
               _Version        =   393216
               ForeColor       =   16711680
               Text            =   ""
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
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000011&
               BackStyle       =   0  'Transparent
               Caption         =   "Upto Settlement"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   285
               Left            =   120
               TabIndex        =   30
               Top             =   120
               Width           =   1620
            End
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FF8080&
            Caption         =   "Update Last Settlement Brokerage"
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
            Height          =   615
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   2895
         End
         Begin VB.CommandButton Command2 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   3120
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Press F7 to set all Rows by Current Cell Value"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   210
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   3990
            WordWrap        =   -1  'True
         End
      End
      Begin MSComctlLib.ListView PartyLst 
         Height          =   2580
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Party Name"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ItemLst 
         Height          =   2580
         Left            =   6840
         TabIndex        =   3
         Top             =   840
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Name"
            Object.Width           =   5185
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2580
         Left            =   3720
         TabIndex        =   2
         Top             =   840
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   5185
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Exchange Name"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   11520
         TabIndex        =   19
         Top             =   360
         Width           =   2115
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "frmbrok.frx":0051
      Left            =   4800
      List            =   "frmbrok.frx":0061
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   360
      Left            =   4320
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ForeColor       =   16711680
      Text            =   "DataCombo3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   360
      Left            =   4560
      TabIndex        =   13
      Top             =   6120
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ForeColor       =   16711680
      Text            =   "DataCombo2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      ItemData        =   "frmbrok.frx":00AE
      Left            =   360
      List            =   "frmbrok.frx":00E5
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      ItemData        =   "frmbrok.frx":0230
      Left            =   360
      List            =   "frmbrok.frx":023D
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "frmbrok.frx":025C
      Left            =   5400
      List            =   "frmbrok.frx":0266
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483626
      ForeColor       =   4194368
      HeadLines       =   1
      RowHeight       =   18
      TabAction       =   1
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   30
      BeginProperty Column00 
         DataField       =   "ITEMCODE"
         Caption         =   "Item Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ITEMNAME"
         Caption         =   "Item Name"
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
         DataField       =   "BROKTYPE"
         Caption         =   "Brok. Type"
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
         DataField       =   "BROKRATE"
         Caption         =   "Brok Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "STDRATE"
         Caption         =   "Std Rate"
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
      BeginProperty Column05 
         DataField       =   "TRANRATE"
         Caption         =   "Tran Rate"
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
      BeginProperty Column06 
         DataField       =   "TranType"
         Caption         =   "TranType"
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
         DataField       =   "PARTYTYPE"
         Caption         =   "Party Type"
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
         DataField       =   "BCYCLE"
         Caption         =   "Bill Cycle"
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
         DataField       =   "STTRATE"
         Caption         =   "STT Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "STMRATE"
         Caption         =   "Stamp Duty Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "BBROKRATE"
         Caption         =   "BBrokRate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "UPTOSTDT"
         Caption         =   "Set. Date"
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
      BeginProperty Column13 
         DataField       =   "Broktype2"
         Caption         =   "BrokType2"
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
         DataField       =   "BrokRate2"
         Caption         =   "BrokRate2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "MINRATE"
         Caption         =   "Min Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "MBROKTYPE"
         Caption         =   "MBroktype"
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
         DataField       =   "MBROKRATE"
         Caption         =   "MBrokrate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "MBROKRATE2"
         Caption         =   "MBrokrate2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Share"
         Caption         =   "Share(%)"
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
      BeginProperty Column20 
         DataField       =   "Applyon"
         Caption         =   "Apply On"
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
      BeginProperty Column21 
         DataField       =   "SHARE2"
         Caption         =   "Share 2 (%)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Applyon2"
         Caption         =   "Apply On 2"
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
      BeginProperty Column23 
         DataField       =   "MarType"
         Caption         =   "MarType"
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
      BeginProperty Column24 
         DataField       =   "MarRate"
         Caption         =   "MarRate"
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
      BeginProperty Column25 
         DataField       =   "New"
         Caption         =   "New"
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
      BeginProperty Column26 
         DataField       =   "Party"
         Caption         =   "Party"
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
      BeginProperty Column27 
         DataField       =   "PartyCode"
         Caption         =   "PartyCode"
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
      BeginProperty Column28 
         DataField       =   "EXCODE"
         Caption         =   "ExCode"
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
      BeginProperty Column29 
         DataField       =   "deleterow"
         Caption         =   "DeleteRow"
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
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2594.835
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column19 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column21 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column24 
            Alignment       =   1
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column25 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column26 
            Locked          =   -1  'True
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column27 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column28 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   9030
      Left            =   240
      TabIndex        =   16
      Top             =   1200
      Width           =   14895
   End
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   15840
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   4695
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4140
         _ExtentX        =   7303
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
      Begin MSDataListLib.DataCombo ItemDbComb 
         Height          =   360
         Left            =   6120
         TabIndex        =   7
         Top             =   240
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   16711680
         Text            =   ""
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
      Begin VB.Label Label2 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   293
         Width           =   495
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   9300
      Left            =   120
      Top             =   1080
      Width           =   15165
   End
End
Attribute VB_Name = "frmbrok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FlagBrok As Boolean:          Dim AllExcodes As Boolean:          Dim AllItems As Boolean:            Dim AllParties As Boolean
Public Fb_Press As Byte:             Dim LInstType As String * 3
Dim LSParties As String:             Dim LExCodes As String:             Dim LSItems As String
Dim LGridRow As Long:                Dim LGridCol As Integer:            Dim LSettlementDt As String
Dim GridColVal As String::           Dim CountRow As Long:               Dim SearchRow As Double:
Dim RECGRID As ADODB.Recordset:      Dim ItemRec As ADODB.Recordset:     Dim ExRec As ADODB.Recordset
Dim UptoStDtRec As ADODB.Recordset: Public LDataCol As Integer:          Dim AccRecADO As ADODB.Recordset

Sub ADD_NEW()
    Frame1.Enabled = True:    Frame2.Enabled = True
    Frame3.Enabled = True:    Call Get_Selection(1):    PartyLst.SetFocus
End Sub
Sub CANCEL_REC()
    Dim I As Integer
    AllExcodes = False: AllItems = False: AllParties = False
    Fb_Press = 0
    Check3.Value = 1
    ItemLst.Enabled = True:     PartyLst.Enabled = True
    ListView1.Enabled = True:   Check5.Enabled = True
    Check1.Enabled = True:      Check2.Enabled = True
    Check3.Enabled = True:      Check4.Enabled = True
    Check6.Enabled = True
    Command2.Enabled = True:    DataCombo4.Enabled = True
    For I = 1 To PartyLst.ListItems.Count
        PartyLst.ListItems.Item(I).Checked = False
    Next I
    For I = 1 To ItemLst.ListItems.Count
        ItemLst.ListItems.Item(I).Checked = False
    Next I
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(I).Checked = False
    Next I
    Check1.Value = 0:    Check4.Value = 0:    Check5.Value = 0
    Call RecSet
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataGrid1.Enabled = False
    DataCombo1.Enabled = True: ItemDbComb.Enabled = True: Command2.Enabled = True: Frame1.Enabled = False: Combo1.Visible = False: Combo2.Visible = False: Combo3.Visible = False: Combo4.Visible = False: Combo5.Visible = False: DataCombo2.Visible = False: DataCombo3.Visible = False: ItemDbComb.text = "": DataCombo1.text = ""
    Frame2.Enabled = False: Frame3.Enabled = False
    Call Get_Selection(13)
End Sub
Sub Save_Rec()
    Dim LNTtems As String
    Dim LStdTDate As Date
    
    Dim LSSaudas As String
    Dim AccRec As ADODB.Recordset
    Dim TempRec As ADODB.Recordset
    Dim TRec As New ADODB.Recordset
    Dim TRec2 As New ADODB.Recordset
    
    On Error GoTo ERR1
    ItemLst.Enabled = False:    PartyLst.Enabled = False
    Check1.Enabled = False:    Check2.Enabled = False
    Check3.Enabled = False:    Check4.Enabled = False
    Command2.Enabled = False:    DataCombo4.Enabled = False
    MYSQL = "DELETE FROM PITBROK WHERE UptoStDt IS NULL"
    Cnn.Execute MYSQL
    If Combo6.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf Combo6.ListIndex = 1 Then
        LInstType = "OPT"
    Else
        LInstType = "CSH"
    End If
    If IsDate(DataCombo4.text) Then LStdTDate = DateValue(DataCombo4.text)
    If RECGRID.RecordCount > 0 Then
        Set TempRec = RECGRID.Clone
        TempRec.MoveFirst
        Cnn.BeginTrans: CNNERR = True
        LSItems = vbNullString
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If Not IsNull(TempRec!ITEMCODE) Then
                If Trim(TempRec!ITEMCODE) <> "" Then
                    If AllItems = False Then
                        If LenB(LSItems) > 0 Then LSItems = LSItems & ", "
                        LSItems = LSItems & "'" & TempRec!ITEMCODE & "'"
                    End If
                    MYSQL = "DELETE FROM PITBROK WHERE COMPCODE =" & MC_CODE & " AND INSTTYPE='" & LInstType & "' AND AC_CODE = '" & TempRec!PARTYCODE & "'"
                    MYSQL = MYSQL & " AND ITEMCODE ='" & TempRec!ITEMCODE & "'"
                    MYSQL = MYSQL & " AND UpToStDt = '" & Format(TempRec!UPTOSTDT, "yyyy/MM/dd") & "' "
                    MYSQL = MYSQL & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                    Cnn.Execute MYSQL
                    If TempRec!DELETEROW <> "Y" Then
                        If IsDate(DataCombo4.text) Then
                            If DateValue(DataCombo4.text) > DateValue(GSysLockDt) Then
                                Call PInsert_PitBrok(MC_CODE, TempRec!PARTYCODE, TempRec!ITEMCODE, Left$(TempRec!BrokType, 1), Val(TempRec!BROKRATE & ""), Val(TempRec!STDRATE & ""), Val(TempRec!TranRate & ""), Left$(TempRec!TranType, 1), Val(TempRec!BROKRATE2), Left$(TempRec!MARTYPE, 1), Val(TempRec!MarRate & ""), Format(LStdTDate, "yyyy/MM/dd"), Val(TempRec!STMRATE), Val(TempRec!STTRATE), Val(TempRec!MINRATE), Left$(TempRec!MBrokType, 1), Val(TempRec!MBrokRate), Val(TempRec!MBROKRATE2), LInstType, TempRec!EXCODE)
                            Else
                                MsgBox "Sorry System Locked.  No Modification Allowed"
                                Exit Do
                            End If
                        Else
                            If DateValue(TempRec!UPTOSTDT) > DateValue(GSysLockDt) Then
                                Call PInsert_PitBrok(MC_CODE, TempRec!PARTYCODE, TempRec!ITEMCODE, Left$(TempRec!BrokType, 1), Val(TempRec!BROKRATE & ""), Val(TempRec!STDRATE & ""), Val(TempRec!TranRate & ""), Left$(TempRec!TranType, 1), Val(TempRec!BROKRATE2), Left$(TempRec!MARTYPE, 1), Val(TempRec!MarRate & ""), Format(TempRec!UPTOSTDT, "yyyy/MM/dd"), Val(TempRec!STMRATE), Val(TempRec!STTRATE), Val(TempRec!MINRATE), Left$(TempRec!MBrokType, 1), Val(TempRec!MBrokRate), Val(TempRec!MBROKRATE2), LInstType, TempRec!EXCODE)
                            Else
                                MsgBox "Sorry System Locked.  No Modification Allowed"
                                Exit Do
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            TempRec.MoveNext
        Loop
        If Check3.Value = 1 Then
            Set AccRec = Nothing
            Set AccRec = New ADODB.Recordset
            MYSQL = "SELECT AC_CODE FROM ACCOUNTD WHERE COMPCODE=" & MC_CODE & " "
            If AllParties = False Then MYSQL = MYSQL & " AND AC_CODE IN (" & LSParties & ") ORDER BY AC_CODE"
            AccRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
            While Not AccRec.EOF
                MYSQL = "SELECT EXCHANGECODE,ITEMCODE,MARTYPE,MARRATE FROM ITEMMAST WHERE COMPCODE  = " & MC_CODE & " "
                If AllExcodes = False Then MYSQL = MYSQL & " AND EXCHANGECODE  IN (" & LExCodes & ")"
                If AllItems = False Then MYSQL = MYSQL & " AND ITEMCODE IN (" & LSItems & ")"
                MYSQL = MYSQL & "  ORDER BY ITEMCODE "
                Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
                Do While Not TRec.EOF
                    DoEvents
                    MYSQL = "SELECT AC_CODE FROM PITBROK WHERE COMPCODE=" & MC_CODE & " AND AC_CODE ='" & AccRec!AC_CODE & "' AND ITEMCODE='" & TRec!ITEMCODE & "' AND UptoStdt='" & Format(GFinEnd, "yyyy/MM/dd") & "' AND INSTTYPE ='" & LInstType & "'"
                    Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
                    If TRec2.EOF Then
                        If DateValue(GFinEnd) > DateValue(GSysLockDt) Then
                            Call PInsert_PitBrok(MC_CODE, AccRec!AC_CODE, TRec!ITEMCODE, "P", 0, 0, 0, "P", 0, "I", 0, Format(GFinEnd, "yyyy/MM/dd"), 0, 0, 0, "P", 0, 0, LInstType, TRec!EXCHANGECODE)
                        End If
                    End If
                    TRec.MoveNext
                Loop
                AccRec.MoveNext
            Wend
        End If
        LSSaudas = vbNullString
        If AllItems = True Then
            LSItems = vbNullString
            LSSaudas = vbNullString
        Else
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            MYSQL = "SELECT DISTINCT SAUDACODE FROM SAUDAMAST WHERE COMPCODE = " & MC_CODE & " "
            If Trim(LSItems) <> "" Then MYSQL = MYSQL & " AND ITEMCODE IN (" & LSItems & ")"
            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
            While Not TRec.EOF
                If LenB(LSSaudas) <> 0 Then LSSaudas = LSSaudas & ", "
                LSSaudas = LSSaudas & "'" & TRec!SAUDACODE & "'"
                TRec.MoveNext
            Wend
        End If
        If AllParties = True Then LSParties = vbNullString
        If AllExcodes = True Then LExCodes = vbNullString
        Call UpdateBrokRateType(LSParties, LSItems, vbNullString, vbNullString, vbNullString, LExCodes)
        If GMarginYN = "Y" And GAppSpread = "Y" Then Call UpdateMargin(LSItems, LSParties, CStr(GFinBegin), CStr(GFinEnd), LExCodes)
        Cnn.CommitTrans: CNNERR = False
        If Check2.Value <> 1 Then
            Cnn.BeginTrans: CNNERR = True
            If BILL_GENERATION(GFinBegin, GFinEnd, LSSaudas, LSParties, LExCodes) Then
                Cnn.CommitTrans: CNNERR = False
            Else
                Cnn.RollbackTrans: CNNERR = False
            End If
        End If
    End If
    Call CANCEL_REC
    Exit Sub
ERR1:
    If CNNERR = True Then
        'Resume
        MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    End If
End Sub
Private Sub Check1_Click()
    Dim I As Integer
    For I = 1 To PartyLst.ListItems.Count
        If Check1.Value = 1 Then
            PartyLst.ListItems.Item(I).Checked = True
        Else
            PartyLst.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub Check2_Click()
Dim ListIt As ListItem:
    Set AccRecADO = Nothing: Set AccRecADO = New ADODB.Recordset
    If Check2.Value = 1 Then
        MYSQL = "SELECT DISTINCT ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC, CTR_D AS CT ,ACCFMLY AS AF WHERE ACC.COMPCODE=" & MC_CODE & " AND ACC.COMPCODE = CT.COMPCODE AND ACC.COMPCODE = AF.COMPCODE AND CT.USERID = AF.FMLYCODE AND AF.FMLYHEAD = ACC.AC_CODE ORDER BY ACC.NAME"
        
    Else
        MYSQL = "SELECT DISTINCT PB.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC,PITBROK AS PB WHERE ACC.COMPCODE=" & MC_CODE & " AND ACC.COMPCODE = PB.COMPCODE AND ACC.AC_CODE = PB.AC_CODE ORDER BY ACC.NAME"
    End If
    AccRecADO.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRecADO.EOF Then
        PartyLst.ListItems.clear
        While Not AccRecADO.EOF
            Set ListIt = PartyLst.ListItems.Add(, , AccRecADO!NAME)
            ListIt.SubItems(1) = AccRecADO!AC_CODE
            AccRecADO.MoveNext
        Wend
    End If
End Sub
Private Sub Check4_Click()
Dim I As Integer
    For I = 1 To ItemLst.ListItems.Count
        If Check4.Value = 1 Then
            ItemLst.ListItems.Item(I).Checked = True
        Else
            ItemLst.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub Check5_Click()
Dim I As Integer
For I = 1 To ListView1.ListItems.Count
    If Check5.Value = 1 Then
        ListView1.ListItems.Item(I).Checked = True
    Else
        ListView1.ListItems.Item(I).Checked = False
    End If
Next I
Call ListView1_Click
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
    Check6.Caption = "Show Only Traded Commodities"
Else
    Check6.Caption = "Show All Commodities"
End If
End Sub

Private Sub Combo4_GotFocus()
    If Mid(RECGRID!MARTYPE, 1, 1) = "Q" Then
        Combo4.ListIndex = Val(0)
    ElseIf Mid(RECGRID!MARTYPE, 1, 1) = "V" Then
        Combo4.ListIndex = Val(1)
    End If
    Combo4.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    Combo4.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    Combo4.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If KeyCode = 13 Then RECGRID!MARTYPE = Combo4.text
        Combo4.Visible = False: DataGrid1.Col = 23: DataGrid1.SetFocus
    ElseIf KeyCode = 27 Then
        Combo4.Visible = False
    End If
End Sub
Private Sub Combo4_Validate(Cancel As Boolean)
    If Len(Trim(Combo4.text)) < 1 Then Cancel = True: Exit Sub
End Sub

Private Sub Combo5_GotFocus()
    If Mid(RECGRID!APPLYON, 1, 1) = "N" Then
        Combo5.ListIndex = Val(0)
    ElseIf Mid(RECGRID!APPLYON, 1, 1) = "G" Then
        Combo5.ListIndex = Val(1)
    End If
    Combo5.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    Combo5.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    Combo5.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If LDataCol = 20 Then
            RECGRID!APPLYON = Combo5.text
            DataGrid1.Col = 20
        Else
            RECGRID!APPLYON2 = Combo5.text
            DataGrid1.Col = 22
        End If
        Combo5.Visible = False: DataGrid1.SetFocus
    ElseIf KeyCode = 27 Then
        Combo5.Visible = False
    End If
End Sub
Private Sub Combo5_Validate(Cancel As Boolean)
    If Len(Trim(Combo5.text)) < 1 Then Cancel = True: Exit Sub
End Sub
Private Sub Combo6_Validate(Cancel As Boolean)
If Combo6.ListIndex < 0 Then
    MsgBox "Please Select Instrument Type "
    Cancel = True
End If
End Sub
Public Sub Command2_Click()
    Dim PartyRec As ADODB.Recordset
    Dim ItemRec As ADODB.Recordset
    Dim BrokRec As ADODB.Recordset:
    Dim ChkCount  As Integer
    Dim LBrokType As String * 1
    Dim J As Integer
    DataCombo4.Enabled = False
    ItemLst.Enabled = False:        PartyLst.Enabled = False
    ListView1.Enabled = False:      Check5.Enabled = False
    Check1.Enabled = False:         Check2.Enabled = False
    Check3.Enabled = False:         Check4.Enabled = False
    Check6.Enabled = False
    Command2.Enabled = False:       LSParties = vbNullString
    ChkCount = 0
    For J = 1 To PartyLst.ListItems.Count
        If PartyLst.ListItems(J).Checked = True Then
            ChkCount = ChkCount + 1
            If LenB(LSParties) > 0 Then LSParties = LSParties & ", "
            LSParties = LSParties & "'" & PartyLst.ListItems(J).SubItems(1) & "'"
        End If
    Next
    If PartyLst.ListItems.Count = ChkCount Then AllParties = True
    If LenB(LSParties) = 0 Then
        PartyLst.Enabled = True:        ListView1.Enabled = True:        ItemLst.Enabled = True
        Command2.Enabled = True:        Check1.Enabled = True:           Check2.Enabled = True
        Check3.Enabled = True:          Check4.Enabled = True:           Check5.Enabled = True
        Check6.Enabled = True
        MsgBox "Please Select Party.", vbCritical:
        PartyLst.SetFocus:
        Exit Sub
    End If
    ChkCount = 0
    For J = 1 To ItemLst.ListItems.Count
        If ItemLst.ListItems(J).Checked = True Then
            ChkCount = ChkCount + 1
            If Len(LSItems) > 0 Then LSItems = LSItems & ", "
            LSItems = LSItems & "'" & ItemLst.ListItems(J).SubItems(1) & "'"
        End If
    Next
    If ChkCount = ItemLst.ListItems.Count Then
        AllItems = True
    Else
        AllItems = False
    End If
    If LenB(LSItems) = 0 Then
        PartyLst.Enabled = True:        ListView1.Enabled = True
        ItemLst.Enabled = True:         Command2.Enabled = True
        Check1.Enabled = True:          Check2.Enabled = True
        Check3.Enabled = True:          Check4.Enabled = True
        Check5.Enabled = True:          Check6.Enabled = True
        MsgBox "Please Select Commodity/Script ", vbCritical:
        ItemLst.SetFocus:
        Exit Sub
    End If
    CountRow = -1
    If Combo6.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf Combo6.ListIndex = 1 Then
        LInstType = "OPT"
    ElseIf Combo6.ListIndex = 2 Then
        LInstType = "CSH"
    End If
    Call RecSet
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    Set BrokRec = Nothing: Set BrokRec = New ADODB.Recordset
    MYSQL = "SELECT AM.NAME,AM.AC_CODE,A.TRANTYPE,A.ITEMCODE,B.ITEMNAME,A.BROKTYPE,A.BROKRATE,A.BROKRATE2,A.STMRATE,A.STTRATE,A.BBROKRATE,A.STDRATE,A.TRANRATE,"
    MYSQL = MYSQL & " A.UPTOSTDT,A.MARTYPE,A.MARRATE,A.MINRATE,A.MBROKRATE,A.MBROKRATE2,A.MBROKTYPE,B.EXCHANGECODE "
    MYSQL = MYSQL & " FROM PITBROK AS A, ITEMMAST AS B ,ACCOUNTM AS AM  WHERE A.COMPCODE=" & MC_CODE & " AND A.COMPCODE=B.COMPCODE AND A.ITEMCODE=B.ITEMCODE AND A.COMPCODE = AM.COMPCODE AND A.AC_CODE=AM.AC_CODE "
    MYSQL = MYSQL & " AND A.INSTTYPE='" & LInstType & "'"
    If AllParties = False Then MYSQL = MYSQL & " AND A.AC_CODE IN (" & LSParties & ")"
    If AllExcodes = False Then MYSQL = MYSQL & " AND B.EXCHANGECODE IN (" & LExCodes & ") "
    If AllItems = False Then MYSQL = MYSQL & " AND B.ITEMCODE IN (" & LSItems & ")"
    If Check6.Value = 1 Then
        MYSQL = MYSQL & " AND B.ITEMCODE IN (SELECT DISTINCT ITEMCODE FROM CTR_D WHERE COMPCODE =" & MC_CODE & " "
        If AllParties = False Then
            MYSQL = MYSQL & "     AND PARTY IN (" & LSParties & " ))"
        Else
            MYSQL = MYSQL & " )"
        End If
    End If
    If IsDate(DataCombo4.text) Then MYSQL = MYSQL & " AND A.UPTOSTDT = '" & Format(DataCombo4.text, "yyyy/MM/dd") & "'"
    MYSQL = MYSQL & " ORDER BY AM.NAME,B.EXCHANGECODE,A.ITEMCODE,A.UPTOSTDT "
    BrokRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not BrokRec.EOF Then
        DataGrid1.Enabled = True
        Do While Not BrokRec.EOF
            DoEvents
            RECGRID.AddNew
            RECGRID.Fields("ITEMCODE") = BrokRec!ITEMCODE
            RECGRID.Fields("ITEMNAME") = BrokRec!ITEMName
            If IsNull(BrokRec!BrokType) Or BrokRec!BrokType = "" Then
                RECGRID.Fields("BROKTYPE") = "Transaction"
            Else
                LBrokType = BrokRec!BrokType
                Select Case LBrokType
                Case "T"
                    RECGRID.Fields("BROKTYPE") = "Transaction"
                Case "O"
                    RECGRID.Fields("BROKTYPE") = "Opening Sauda"
                Case "C"
                    RECGRID.Fields("BROKTYPE") = "Closing Sauda"
                Case "Q"
                    RECGRID.Fields("BROKTYPE") = "Qtywise IntraDay"
                Case "P"
                    RECGRID.Fields("BROKTYPE") = "Percentage wise"
                Case "B"
                    RECGRID.Fields("BROKTYPE") = "Bought"
                Case "S"
                    RECGRID.Fields("BROKTYPE") = "Sold"
                Case "I"
                    RECGRID.Fields("BROKTYPE") = "IntraDay Brokerage"
                Case "V"
                    RECGRID.Fields("BROKTYPE") = "Valuewise Intraday"
                Case "D"
                    RECGRID.Fields("BROKTYPE") = "Delivery Wise Brokerage"
                Case "H"
                    RECGRID.Fields("BROKTYPE") = "Higher Value Percentage Wise"
                Case "L"
                    RECGRID.Fields("BROKTYPE") = "LotWise Higher Value"
                Case "W"
                    RECGRID.Fields("BROKTYPE") = "WHigher Value Intraday Wise"
                Case "X"
                    RECGRID.Fields("BROKTYPE") = "XIntraday Higher Wise"
                Case "Z"
                    RECGRID.Fields("BROKTYPE") = "ZLotwise"
                Case "R"
                    RECGRID.Fields("BROKTYPE") = "RZLotwise IntraDay Wise"
                Case "F"
                    RECGRID.Fields("BROKTYPE") = "Fixed Brokerage"
                Case "N"
                    RECGRID.Fields("BROKTYPE") = "N Per Trade"
                Case "U"
                    RECGRID.Fields("BROKTYPE") = "U ShareQty Wise"
                End Select
            End If
            RECGRID.Fields("MBROKRATE") = IIf(IsNull(BrokRec!MBrokRate), 0, BrokRec!MBrokRate)
            RECGRID.Fields("MBROKRATE2") = IIf(IsNull(BrokRec!MBROKRATE2), 0, BrokRec!MBROKRATE2)
            RECGRID.Fields("MINRATE") = IIf(IsNull(BrokRec!MINRATE), 0, BrokRec!MINRATE)
            RECGRID.Fields("BROKRATE2") = IIf(IsNull(BrokRec!BROKRATE2), 0, BrokRec!BROKRATE2)
            RECGRID.Fields("BBROKRATE") = IIf(IsNull(BrokRec!BBROKRATE), 0, BrokRec!BBROKRATE)
            RECGRID.Fields("STMRATE") = IIf(IsNull(BrokRec!STMRATE), 0, BrokRec!STMRATE)
            RECGRID.Fields("STTRATE") = IIf(IsNull(BrokRec!STTRATE), 0, BrokRec!STTRATE)
            RECGRID.Fields("STDRATE") = BrokRec!STDRATE
            RECGRID.Fields("TRANRATE") = BrokRec!TranRate
            If IsNull(BrokRec!TranType) Or BrokRec!TranType = "" Then
                RECGRID.Fields("TranType") = "Transaction"
            Else
                If BrokRec!TranType = "T" Then
                    RECGRID.Fields("TranType") = "Transaction"
                ElseIf BrokRec!TranType = "P" Then
                    RECGRID.Fields("TranType") = "Percentage wise"
                End If
            End If
            If IsNull(BrokRec!UPTOSTDT) Then
                RECGRID.Fields("UPTOSTDT") = GFinEnd
            Else
                If BrokRec!UPTOSTDT = "" Then
                    RECGRID.Fields("UPTOSTDT") = GFinEnd
                ElseIf DateValue(BrokRec!UPTOSTDT) = DateValue("01/01/1900") Then
                    RECGRID.Fields("UPTOSTDT") = GFinEnd
                Else
                    RECGRID.Fields("UPTOSTDT") = Format(BrokRec!UPTOSTDT, "DD/MM/YYYY")
                End If
            End If
            If IsNull(BrokRec!MARTYPE) Or BrokRec!MARTYPE = "" Then
                RECGRID.Fields("MARTYPE") = "Value Wise (In %)"
            Else
                If BrokRec!MARTYPE = "Q" Then
                    RECGRID.Fields("MARTYPE") = "Qtywise (Per Unit)"
                ElseIf BrokRec!MARTYPE = "V" Then
                    RECGRID.Fields("MARTYPE") = "Value Wise (In %)"
                ElseIf BrokRec!MARTYPE = "I" Then
                    RECGRID.Fields("MARTYPE") = "Import Rates"
                ElseIf BrokRec!MARTYPE = "C" Then
                    RECGRID.Fields("MARTYPE") = "Client Wise Margin"
                Else
                    RECGRID.Fields("MARTYPE") = "Import Wise"
                End If
            End If
            If IsNull(BrokRec!MBrokType) Or BrokRec!MBrokType = "" Then
                RECGRID.Fields("MBROKTYPE") = "Transaction"
            Else
                If BrokRec!MBrokType = "T" Then
                    RECGRID.Fields("MBROKTYPE") = "Transaction"
                ElseIf BrokRec!MBrokType = "O" Then
                    RECGRID.Fields("MBROKTYPE") = "Opening Sauda"
                ElseIf BrokRec!MBrokType = "C" Then
                    RECGRID.Fields("MBROKTYPE") = "Closing Sauda"
                ElseIf BrokRec!MBrokType = "Q" Then
                    RECGRID.Fields("MBROKTYPE") = "Qtywise IntraDay"
                ElseIf BrokRec!MBrokType = "P" Then
                    RECGRID.Fields("MBROKTYPE") = "Percentage wise"
                ElseIf BrokRec!MBrokType = "B" Then
                    RECGRID.Fields("MBROKTYPE") = "Bought"
                ElseIf BrokRec!MBrokType = "S" Then
                    RECGRID.Fields("MBROKTYPE") = "Sold"
                ElseIf BrokRec!MBrokType = "I" Then
                    RECGRID.Fields("MBROKTYPE") = "IntraDay Brokerage"
                ElseIf BrokRec!MBrokType = "V" Then
                    RECGRID.Fields("MBROKTYPE") = "Valuewise Intraday"
                ElseIf BrokRec!MBrokType = "D" Then
                    RECGRID.Fields("MBROKTYPE") = "Delivery Wise Brokerage"
                ElseIf BrokRec!MBrokType = "H" Then
                    RECGRID.Fields("MBROKTYPE") = "Higher Value Percentage Wise"
                ElseIf BrokRec!MBrokType = "L" Then
                    RECGRID.Fields("MBROKTYPE") = "LotWise Higher Value"
                ElseIf BrokRec!MBrokType = "W" Then
                    RECGRID.Fields("MBROKTYPE") = "WHigher Value Intraday Wise"
                ElseIf BrokRec!MBrokType = "X" Then
                    RECGRID.Fields("MBROKTYPE") = "XIntraday Higher Wise"
                ElseIf BrokRec!MBrokType = "Z" Then
                    RECGRID.Fields("MBROKTYPE") = "ZLotwise"
                ElseIf BrokRec!MBrokType = "R" Then
                    RECGRID.Fields("MBROKTYPE") = "RZLotwise IntraDay Wise"
                ElseIf BrokRec!MBrokType = "F" Then
                    RECGRID.Fields("MBROKTYPE") = "Fixed Brokerage"
                End If
            End If
            RECGRID.Fields("MARRATE") = Val(BrokRec!MarRate & "")
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            RECGRID.Fields("PARTY") = BrokRec!NAME & ""
            RECGRID.Fields("PARTYCODE") = BrokRec!AC_CODE & ""
            RECGRID.Fields("EXCODE") = BrokRec!EXCHANGECODE
            RECGRID.Fields("DELETEROW") = "N"
            RECGRID.Update
            BrokRec.MoveNext
        Loop
        RECGRID.AddNew
        RECGRID.Fields("ITEMCODE") = vbNullString:         RECGRID.Fields("ITEMNAME") = vbNullString
        RECGRID.Fields("BROKTYPE") = vbNullString:         RECGRID.Fields("BROKRATE") = 0
        RECGRID.Fields("BBROKRATE") = 0:                   RECGRID.Fields("STDRATE") = 0
        RECGRID.Fields("TRANRATE") = 0:                    RECGRID.Fields("TranType") = vbNullString
        RECGRID.Fields("UPTOSTDT") = vbNullString:         RECGRID.Fields("BROKRATE2") = 0
        RECGRID.Fields("MARTYPE") = vbNullString:          RECGRID.Fields("MARRATE") = 0
        RECGRID.Fields("SHARE") = 0:                       RECGRID.Fields("APPLYON") = vbNullString
        RECGRID.Fields("MBROKTYPE") = vbNullString:        RECGRID.Fields("MBROKRATE") = 0
        RECGRID.Fields("MINRATE") = 0:                     RECGRID.Fields("PARTY") = vbNullString
        RECGRID.Fields("PARTYCODE") = vbNullString:        RECGRID.Fields("EXCODE") = vbNullString
        CountRow = CountRow + 1
        RECGRID.Fields("New") = CountRow
        RECGRID.Update
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh:
        DataCombo1.Enabled = False: ItemDbComb.Enabled = False: Command2.Enabled = False: RECGRID.MoveFirst: DataGrid1.SetFocus
        DataGrid1.LeftCol = 0
        Label3.Visible = True
    Else
        Label3.Visible = False
        MsgBox "No Record does not exists.for selected Commodity/Script", vbExclamation
        If MsgBox("Do you really want to apply Seperate Brokergae for selected Commodity", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
            MYSQL = " SELECT AC_CODE FROM ACCOUNTD WHERE COMPCODE =" & MC_CODE & " AND AC_CODE IN (" & LSParties & ") ORDER BY AC_CODE "
            Set PartyRec = Nothing
            Set PartyRec = New ADODB.Recordset
            PartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
            Do While Not PartyRec.EOF
                MYSQL = "SELECT ITEMCODE,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & MC_CODE & " AND ITEMCODE IN (" & LSItems & ")"
                Set ItemRec = Nothing
                Set ItemRec = New ADODB.Recordset
                ItemRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
                Do While Not ItemRec.EOF
                    DoEvents
                    'MYSQL = "EXEC INSERT_PITBROK " & MC_CODE & ",'" & PartyRec!AC_CODE & "','" & ItemRec!ITEMCODE & "','P',0,0,0,'P',0,'V',0,'" & Format(GFinEnd, "yyyy/MM/dd") & "',0,0,0,'P',0,0,'" & LInstType & "'"
                    Call PInsert_PitBrok(MC_CODE, PartyRec!AC_CODE, ItemRec!ITEMCODE, "P", 0, 0, 0, "P", 0, "I", 0, Format(GFinEnd, "yyyy/MM/dd"), 0, 0, 0, "P", 0, 0, LInstType, ItemRec!EXCHANGECODE)
                   ' Cnn.Execute MYSQL
                    ItemRec.MoveNext
                Loop
                PartyRec.MoveNext
            Loop
        End If
        Call CANCEL_REC
    End If
    DataCombo4.Enabled = True
End Sub
Private Sub DataCombo1_GotFocus()
    DataCombo1.text = ""
    If Frame1.Enabled Then DataCombo1.SetFocus: Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub DataCombo2_GotFocus()
    If Fb_Press = 0 Then
    Else
        DataCombo2.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
        DataCombo2.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
        DataCombo2.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
        DataCombo2.BoundText = RECGRID!ITEMCODE & ""
        Sendkeys "%{DOWN}"
    End If
End Sub
Private Sub DataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTemp As ADODB.Recordset
    Dim TempParties As String
    If KeyCode = 27 Then
        DataCombo2.Visible = False
    ElseIf KeyCode = 13 Then
        If Trim(DataCombo2.BoundText) = "" Then
           MsgBox "Please Select Commodity ", vbCritical: DataCombo2.SetFocus
           Exit Sub
        Else
            If IsNull(RECGRID!BROKRATE) Then MsgBox "Invalid row to add new record", vbCritical: DataCombo2.Visible = False: Exit Sub
            If Trim(ItemDbComb.BoundText) = "" Then
            Else
                If ItemDbComb.BoundText = DataCombo2.BoundText Then
                Else
                    MsgBox "Please Select Commodity " & ItemDbComb.BoundText, vbCritical: DataCombo2.SetFocus
                    Exit Sub
                End If
            End If
            RECGRID!ITEMCODE = DataCombo2.BoundText
            RECGRID!ITEMName = DataCombo2.text
            
            Set RecTemp = RECGRID.Clone
            GridColVal = vbNullString
            RecTemp.MoveFirst
            GridColVal = DataCombo2.BoundText
            Do While Not RECGRID.EOF
                If RecTemp!ITEMCODE = GridColVal Then
                    GridColVal = IIf(IsNull(RecTemp!BrokType), "Transaction", RecTemp!BrokType)
                    Exit Do
                End If
                RecTemp.MoveNext
            Loop
                If LenB(GridColVal) = 0 Then
                    RECGRID!BrokType = "Transaction"
                Else
                    RECGRID!BrokType = GridColVal
                End If
                RECGRID!BROKRATE = "0.000000"
                RECGRID!BBROKRATE = "0.00"
                RECGRID!STDRATE = "0.00"
                RECGRID!TranRate = "0.00"
                GridColVal = vbNullString
                RecTemp.MoveFirst
                GridColVal = DataCombo2.BoundText
                Do While Not RECGRID.EOF
                    If RecTemp!ITEMCODE = GridColVal Then
                        GridColVal = IIf(IsNull(RecTemp!TranType), "Transaction", RecTemp!TranType)
                        Exit Do
                    End If
                    RecTemp.MoveNext
                Loop
                If LenB(GridColVal) = 0 Then
                    RECGRID!TranType = "Transaction"
                Else
                    RECGRID!TranType = GridColVal
                End If
                RECGRID!UPTOSTDT = ""  'LSettlementDt
                RECGRID!MARTYPE = "Value Wise (In %)"
                RECGRID!BROKRATE2 = "0.00"
                RECGRID!MarRate = "0.00"
                RECGRID!Party = Label7.Caption
                TempParties = Mid(LSParties, Val(InStr(LSParties, "'")) + 1, Val(Len(LSParties)))
                TempParties = Left(TempParties, Len(TempParties) - 1)
                RECGRID!PARTYCODE = TempParties
                RECGRID.Update
                Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataCombo1.Enabled = False: ItemDbComb.Enabled = False: Command2.Enabled = False: RECGRID.MoveFirst: DataGrid1.SetFocus
                DataGrid1.Row = RECGRID.RecordCount - 1: DataGrid1.Col = 3
        End If
        DataCombo2.Visible = False
    End If
End Sub
Private Sub DataCombo3_GotFocus()
    DataGrid1.LeftCol = 6
    DataCombo3.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    DataCombo3.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    DataCombo3.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    DataCombo3.text = RECGRID!UPTOSTDT & ""
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        DataCombo3.Visible = False
    ElseIf KeyCode = 13 Then
            'check item wise duplicate settlement ****************
        If DataCombo3.text <> "" Then
            If SYSTEMLOCK(DateValue(DataCombo3.text)) Then
                MsgBox "Sorry System Locked.  No Modification Allowed"
                Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataCombo1.Enabled = False: ItemDbComb.Enabled = False: Command2.Enabled = False: RECGRID.MoveFirst: DataGrid1.SetFocus
                DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol + 1: DataCombo3.Visible = False: DataGrid1.SetFocus
            Else
                LGridRow = DataGrid1.Row: LGridCol = DataGrid1.Col: GridColVal = RECGRID!ITEMCODE: SearchRow = RECGRID!New
                RECGRID.MoveFirst
                Do While Not RECGRID.EOF
                    If RECGRID!ITEMCODE = GridColVal Then
                        If SearchRow = RECGRID!New Then
                        Else
                            If RECGRID!UPTOSTDT = DataCombo3.text Then
                                MsgBox "Duplicate Settlement Date Found.", vbCritical:
                                RECGRID.MoveFirst: RECGRID.Find "New =" & SearchRow & "", , adSearchForward
                                DataCombo3.SetFocus: Exit Sub
                            End If
                        End If
                    Else
                        Exit Do
                    End If
                    RECGRID.MoveNext
                Loop
                RECGRID.MoveFirst: RECGRID.Find "New =" & SearchRow & "", , adSearchForward
                If KeyCode = 13 Then
                    RECGRID!UPTOSTDT = DataCombo3.text
                    LGridRow = SearchRow
                End If
                Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataCombo1.Enabled = False: ItemDbComb.Enabled = False: Command2.Enabled = False: RECGRID.MoveFirst: DataGrid1.SetFocus
                DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol + 1: DataCombo3.Visible = False: DataGrid1.SetFocus
            End If
        End If
    End If
End Sub
Private Sub DataCombo4_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub DataCombo4_Validate(Cancel As Boolean)
    If IsDate(DataCombo4.text) Then
        If SYSTEMLOCK(DateValue(DataCombo4.text)) Then
            MsgBox "Sorry System Locked.  No Addition, Modification or Deletion Allowed"
            Cancel = True
        End If
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error Resume Next
  Label7.Caption = DataGrid1.Columns(22).text
  
End Sub
Private Sub Form_Load()
    Dim TRec As ADODB.Recordset
    Dim ListIt As ListItem
    FlagBrok = False
    'Last Settlement Date
    Combo6.ListIndex = 0
    If GOptions = "Y" Then Combo6.Visible = True
    MYSQL = "UPDATE PITBROK SET BROKTYPE='P' WHERE BROKRATE =0"
    Cnn.Execute MYSQL
    LSettlementDt = vbNullString:
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT MAX(SETDATE) AS MAXSETTLEDATE FROM SETTLE WHERE COMPCODE = " & MC_CODE & "", Cnn, adOpenKeyset, adLockReadOnly
    If Not TRec.EOF Then LSettlementDt = TRec!MaxSettleDate
    Set AccRecADO = Nothing: Set AccRecADO = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC WHERE ACC.COMPCODE =" & MC_CODE & " ORDER BY ACC.NAME"
    AccRecADO.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    DataGrid1.Enabled = False
    If Not AccRecADO.EOF Then
        While Not AccRecADO.EOF
            Set ListIt = PartyLst.ListItems.Add(, , AccRecADO!NAME)
            ListIt.SubItems(1) = AccRecADO!AC_CODE
            AccRecADO.MoveNext
        Wend
        Call Get_Selection(13)
        Set ExRec = Nothing: Set ExRec = New ADODB.Recordset
        MYSQL = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE =" & MC_CODE & " ORDER BY EXCODE "
        ExRec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
        If Not ExRec.EOF Then
            ExRec.MoveFirst
            ListView1.ListItems.clear
            ListView1.Enabled = True: Check4.Enabled = True
            Do While Not ExRec.EOF
                If (ExRec!EXCODE = "EQ" Or ExRec!EXCODE = "BEQ") Then
                    Combo6.Visible = True
                End If
                ListView1.ListItems.Add , , ExRec!EXCODE
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ExRec!EXNAME
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ExRec!EXCODE
                ExRec.MoveNext
            Loop
            
            If ExRec.RecordCount = 1 Then
                Check4.Value = 1: ListView1.TabStop = False: Check4.TabStop = False
                Call Check4_Click
            Else
                ListView1.TabStop = True: Check4.TabStop = True
            End If
        Else
            ListView1.Enabled = False: Check4.Enabled = False
        End If
        Set UptoStDtRec = Nothing: Set UptoStDtRec = New ADODB.Recordset
        UptoStDtRec.Open "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PITBROK WHERE COMPCODE =" & MC_CODE & " ORDER BY UPTOSTDT", Cnn, adOpenKeyset, adLockReadOnly
        If Not UptoStDtRec.EOF Then
            Set DataCombo3.RowSource = UptoStDtRec
            DataCombo3.ListField = "CONDATE"
            DataCombo3.BoundColumn = "CONDATE"
            Set DataCombo4.RowSource = UptoStDtRec
            DataCombo4.ListField = "CONDATE"
            DataCombo4.BoundColumn = "CONDATE"
        End If
    End If
    If GMinBrokYN = "N" Then
        DataGrid1.Columns(15).Visible = False
        DataGrid1.Columns(16).Visible = False
        DataGrid1.Columns(17).Visible = False
        DataGrid1.Columns(18).Visible = False
    End If
    If GStandingYN = "N" Then
        DataGrid1.Columns(4).Visible = False
    End If
    If GSTTYN = "N" Then
        DataGrid1.Columns(9).Visible = False
    End If
    If GValueWiseYN = "N" Then
        DataGrid1.Columns(13).Visible = False
        DataGrid1.Columns(14).Visible = False
    End If
    'If GSubBrokYN = "N" Then
        DataGrid1.Columns(11).Visible = False
    'End If
    
    If GTranFeesYN = "N" Then
        DataGrid1.Columns(5).Visible = False
        DataGrid1.Columns(6).Visible = False
    End If
    If GStampDutyYN = "N" Then
        DataGrid1.Columns(10).Visible = False
    End If
    'If GShare = "N" Then
        DataGrid1.Columns(19).Visible = False
        DataGrid1.Columns(20).Visible = False
        DataGrid1.Columns(21).Visible = False
        DataGrid1.Columns(22).Visible = False
    'End If
    If GMarginYN = "N" Then
        DataGrid1.Columns(23).Visible = False
        DataGrid1.Columns(24).Visible = False
    End If
    
    Frame1.Enabled = False
    Call CANCEL_REC
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "ITEMCODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "ITEMNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "STDRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "TRANRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "TRANTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BBROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "STMRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "UPTOSTDT", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MARTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "MARRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MBROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "MBROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MBROKRATE2", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MINRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "PARTY", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "PARTYCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "STTRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "EXCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "DeleteRow", adVarChar, 1, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRec As ADODB.Recordset
    Dim MUPTOSTDT As Date
    If KeyCode = 13 And DataGrid1.Col = 2 Then ' BROKTYPE
        Combo1.Visible = True: Combo1.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 6 Then 'TRANTYPE
        Combo1.Visible = True: Combo1.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 16 Then 'MBROKTYPE
        Combo1.Visible = True: Combo1.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 23 Then 'MARGIN TYPE
        Combo4.Visible = True: Combo4.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 29 Then 'delete row
        
        DataGrid1.text = UCase(DataGrid1.text)
        If DataGrid1.text = "Y" Then
        Else
            DataGrid1.text = "N"
        End If
        DataGrid1.Col = 29
        
        DataGrid1.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 16 Then
        RECGRID.MoveNext ''ADDING NEW ROW
        If RECGRID.EOF Then
            RECGRID.AddNew
            RECGRID.Fields("ITEMCODE") = vbNullString:   RECGRID.Fields("ITEMNAME") = vbNullString
            RECGRID.Fields("BROKTYPE") = vbNullString:   RECGRID.Fields("BROKRATE") = 0
            RECGRID.Fields("BBROKRATE") = 0:             RECGRID.Fields("STMRATE") = 0
            RECGRID.Fields("STTRATE") = 0:               RECGRID.Fields("STDRATE") = 0
            RECGRID.Fields("TRANRATE") = 0:              RECGRID.Fields("TranType") = ""
            RECGRID.Fields("UPTOSTDT") = vbNullString:   RECGRID.Fields("BROKRATE2") = 0
            RECGRID.Fields("MARTYPE") = vbNullString:    RECGRID.Fields("MARRATE") = 0
            RECGRID.Fields("EXCODE") = vbNullString
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            RECGRID.Update
         End If
        DataGrid1.LeftCol = 0: DataGrid1.Col = 0
    ElseIf KeyCode = vbKeyF2 And Shift = 2 Then 'ctrl + f
        
    ElseIf KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
    'Press F7 to replace all rows with current cell value
    If KeyCode = 118 Then   'F7
        LGridRow = DataGrid1.Row
        LGridCol = DataGrid1.Col
        If DataGrid1.Col = 2 Then 'BROKTYPE
            GridColVal = RECGRID!BrokType
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BrokType = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 3 Then 'BROKRATE
            GridColVal = RECGRID!BROKRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BROKRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 11 Then 'BBROKRATE
            GridColVal = RECGRID!BBROKRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BBROKRATE = GridColVal
                RECGRID.MoveNext
            Wend
        
        ElseIf DataGrid1.Col = 10 Then 'STMRATE
            GridColVal = RECGRID!STMRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!STMRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 9 Then 'STTRATE
            GridColVal = RECGRID!STTRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!STTRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 4 Then 'STDRATE
            GridColVal = RECGRID!STDRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!STDRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 5 Then 'TRANRATE
            GridColVal = RECGRID!TranRate
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!TranRate = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 6 Then 'TRANTYPE
            GridColVal = IIf(IsNull(RECGRID!TranType), "P", RECGRID!TranType)
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!TranType = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 7 Then 'PARTYTYPE
            GridColVal = RECGRID!PARTYTYPE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!PARTYTYPE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 12 Then 'UPTOSTDT
            GridColVal = RECGRID!UPTOSTDT
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!UPTOSTDT = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 14 Then 'brokrate2
            GridColVal = RECGRID!BROKRATE2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BROKRATE2 = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 15 Then 'MINRATE
            GridColVal = RECGRID!MINRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MINRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 16 Then 'MBROKTYPE
            GridColVal = RECGRID!MBrokType
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MBrokType = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 17 Then 'MBROKTRATE
            GridColVal = RECGRID!MBrokRate
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MBrokRate = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 18 Then 'MBROKTRATE
            GridColVal = RECGRID!MBROKRATE2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MBROKRATE2 = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 19 Then 'SHARE
            GridColVal = RECGRID!SHARE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!SHARE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 20 Then 'APPLYON
            GridColVal = RECGRID!APPLYON
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!APPLYON = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 21 Then 'SHARE 2
            GridColVal = RECGRID!SHARE2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!SHARE2 = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 22 Then 'APPLYON 2
            GridColVal = RECGRID!APPLYON2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!APPLYON2 = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 23 Then 'MARTYPE
            GridColVal = RECGRID!MARTYPE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MARTYPE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 24 Then 'MARRATE
            GridColVal = RECGRID!MarRate
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MarRate = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 29 Then 'DELETEROW
            GridColVal = RECGRID!DELETEROW
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!DELETEROW = GridColVal
                RECGRID.MoveNext
            Wend
        End If
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol: DataGrid1.SetFocus
    End If
End Sub
Private Sub Combo1_GotFocus()
    If DataGrid1.Col = 2 Then
            If Mid(RECGRID!BrokType, 1, 1) = "T" Then
                Combo1.ListIndex = Val(0)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "O" Then
                Combo1.ListIndex = Val(1)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "P" Then
                Combo1.ListIndex = Val(2)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "I" Then
                Combo1.ListIndex = Val(3)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "C" Then
                Combo1.ListIndex = Val(4)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "V" Then
                Combo1.ListIndex = Val(5)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "Q" Then
                Combo1.ListIndex = Val(6)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "D" Then
                Combo1.ListIndex = Val(7)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "H" Then
                Combo1.ListIndex = Val(8)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "L" Then
                Combo1.ListIndex = Val(9)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "W" Then
                Combo1.ListIndex = Val(10)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "X" Then
                Combo1.ListIndex = Val(11)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "Z" Then
                Combo1.ListIndex = Val(12)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "R" Then
                Combo1.ListIndex = Val(13)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "F" Then
                Combo1.ListIndex = Val(14)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "N" Then
                Combo1.ListIndex = Val(15)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "U" Then
                Combo1.ListIndex = Val(16)
            End If
    ElseIf DataGrid1.Col = 6 Then
            If Mid(RECGRID!TranType, 1, 1) = "T" Then
                Combo1.ListIndex = Val(0)
            ElseIf Mid(RECGRID!TranType, 1, 1) = "P" Then
                Combo1.ListIndex = Val(1)
            End If
    ElseIf DataGrid1.Col = 16 Then
            If Mid(RECGRID!BrokType, 1, 1) = "T" Then
                Combo1.ListIndex = Val(0)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "O" Then
                Combo1.ListIndex = Val(1)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "P" Then
                Combo1.ListIndex = Val(2)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "I" Then
                Combo1.ListIndex = Val(3)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "C" Then
                Combo1.ListIndex = Val(4)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "V" Then
                Combo1.ListIndex = Val(5)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "Q" Then
                Combo1.ListIndex = Val(6)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "D" Then
                Combo1.ListIndex = Val(7)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "H" Then
                Combo1.ListIndex = Val(8)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "L" Then
                Combo1.ListIndex = Val(9)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "W" Then
                Combo1.ListIndex = Val(10)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "X" Then
                Combo1.ListIndex = Val(11)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "Z" Then
                Combo1.ListIndex = Val(12)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "R" Then
                Combo1.ListIndex = Val(13)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "F" Then
                Combo1.ListIndex = Val(14)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "N" Then
                Combo1.ListIndex = Val(15)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "U" Then
                Combo1.ListIndex = Val(16)
            End If
            
    End If
    Combo1.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    Combo1.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    Combo1.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        LGridRow = DataGrid1.Row: LGridCol = DataGrid1.Col: GridColVal = RECGRID!ITEMCODE: SearchRow = RECGRID!New
        If DataGrid1.Col = 2 Then
            If KeyCode = 13 Then RECGRID!BrokType = Combo1.text
                RECGRID.MoveFirst
                While Not RECGRID.EOF
                    If RECGRID!ITEMCODE = GridColVal Then RECGRID!BrokType = Combo1.text
                    RECGRID.MoveNext
                Wend
        ElseIf DataGrid1.Col = 6 Then
            If KeyCode = 13 Then RECGRID!TranType = Combo1.text
                RECGRID.MoveFirst
                While Not RECGRID.EOF
                    If RECGRID!ITEMCODE = GridColVal Then RECGRID!TranType = Combo1.text
                        RECGRID.MoveNext
                    Wend
        ElseIf DataGrid1.Col = 16 Then ' MBROKTYPE2
                    If KeyCode = 13 Then RECGRID!MBrokType = Combo1.text
                    RECGRID.MoveFirst
                    While Not RECGRID.EOF
                        If RECGRID!ITEMCODE = GridColVal Then RECGRID!MBrokType = Combo1.text
                        RECGRID.MoveNext
                    Wend
        End If
            Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataCombo1.Enabled = False: ItemDbComb.Enabled = False: Command2.Enabled = False: RECGRID.MoveFirst: DataGrid1.SetFocus
            RECGRID.Find "new =" & SearchRow & "", , adSearchForward
            If LGridCol = 6 Then
                DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol + 3: Combo1.Visible = False: DataGrid1.SetFocus
            Else
                DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol + 1: Combo1.Visible = False: DataGrid1.SetFocus
            End If
    ElseIf KeyCode = 27 Then
        Combo1.Visible = False
    End If
    
End Sub
Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 27 Then
        If KeyCode = 13 Then RECGRID!PARTYTYPE = Combo2.text
        DataGrid1.Col = 8: DataGrid1.SetFocus: Combo2.Visible = False
    ElseIf KeyCode = 27 Then
        Combo1.Visible = False
    End If
End Sub
Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 27 Then
        If KeyCode = 13 Then RECGRID!BCYCLE = Combo3.text
        DataGrid1.Col = 11: DataGrid1.SetFocus: Combo3.Visible = False
    ElseIf KeyCode = 27 Then
        Combo1.Visible = False
    End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If Len(Trim(Combo1.text)) < 1 Then
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
    If Len(Trim(Combo2.text)) < 1 Then
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
    If Len(Trim(Combo3.text)) < 1 Then
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub Combo2_GotFocus()
    DataGrid1.LeftCol = 3
    If Mid(RECGRID!PARTYTYPE, 1, 1) = "M" Then
        Combo2.ListIndex = Val(0)
    ElseIf Mid(RECGRID!PARTYTYPE, 1, 1) = "N" Then
        Combo2.ListIndex = Val(1)
    ElseIf Mid(RECGRID!PARTYTYPE, 1, 1) = "B" Then
        Combo2.ListIndex = Val(2)
    End If
    Combo2.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    Combo2.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    Combo2.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo3_GotFocus()
    If Mid(RECGRID!BCYCLE, 1, 1) = "S" Then
        Combo3.ListIndex = Val(0)
    ElseIf Mid(RECGRID!BCYCLE, 1, 1) = "D" Then
        Combo3.ListIndex = Val(1)
    End If
    Combo3.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    Combo3.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Combo3.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    Sendkeys "%{DOWN}"
End Sub
Private Sub Form_Paint()
    'If FlagBrok And LenB(MAc_Code) <> 0 Then
    '    DataCombo1.BoundText = MAc_Code
    '    Call Get_Selection(1)
    '    Call Command2_Click
    'End If
End Sub
Private Sub ItemDbComb_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub ItemDbComb_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub ListView1_Click()
    Dim TRec As ADODB.Recordset
    Dim ChkCount As Integer
    Dim I As Integer
    Dim ListIt As ListItem
    LExCodes = vbNullString
    ChkCount = 0
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            ChkCount = ChkCount + 1
            If LenB(LExCodes) > 1 Then LExCodes = LExCodes & ", "
            LExCodes = LExCodes & "'" & ListView1.ListItems(I).ListSubItems(2) & "'"
        End If
    Next I
    If ChkCount = ListView1.ListItems.Count Then
        AllExcodes = True
    Else
        AllExcodes = False
    End If
    ItemLst.ListItems.clear
    If LenB(LExCodes) = 0 Then Me.MousePointer = 0: Exit Sub
    MYSQL = "SELECT ITEMCODE,ITEMNAME FROM ITEMMAST WHERE COMPCODE =" & MC_CODE & " "
    MYSQL = MYSQL & " AND EXCHANGECODE in (" & LExCodes & ")  ORDER BY EXCHANGECODE,ITEMNAME"
    Set TRec = Nothing: Set TRec = New ADODB.Recordset:
    TRec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
    While Not TRec.EOF
        Set ListIt = ItemLst.ListItems.Add(, , TRec!ITEMName)
        ListIt.SubItems(1) = TRec!ITEMCODE
        TRec.MoveNext
    Wend
    Set TRec = Nothing
End Sub
