VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_SBRok 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15825
   ScaleWidth      =   28710
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      BackColor       =   &H00000040&
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
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   18015
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sub Brokerage && Sharing Setup"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   17655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   9360
      TabIndex        =   28
      Top             =   10680
      Visible         =   0   'False
      Width           =   6975
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   1200
         TabIndex        =   29
         Top             =   240
         Width           =   4500
         _ExtentX        =   7938
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
         TabIndex        =   30
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
         TabIndex        =   32
         Top             =   293
         Width           =   495
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
         TabIndex        =   31
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00404000&
      Height          =   8550
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   17655
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "Frm_SBRok.frx":0000
         Left            =   360
         List            =   "Frm_SBRok.frx":0016
         TabIndex        =   33
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
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
         ItemData        =   "Frm_SBRok.frx":0073
         Left            =   840
         List            =   "Frm_SBRok.frx":007D
         TabIndex        =   27
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
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
         ItemData        =   "Frm_SBRok.frx":0094
         Left            =   1320
         List            =   "Frm_SBRok.frx":00A4
         TabIndex        =   24
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
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
         ItemData        =   "Frm_SBRok.frx":00F1
         Left            =   3720
         List            =   "Frm_SBRok.frx":00FB
         TabIndex        =   23
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   17415
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFC0&
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
            Height          =   350
            Left            =   9120
            TabIndex        =   37
            Top             =   120
            Width           =   2415
            Begin VB.Label Label10 
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
               TabIndex        =   38
               Top             =   0
               Width           =   945
            End
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H0080C0FF&
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
            Left            =   10440
            TabIndex        =   36
            Top             =   480
            Width           =   1095
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFC0&
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
            Left            =   11760
            TabIndex        =   17
            Top             =   960
            Width           =   4335
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Update Last Settlement"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   1320
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.ComboBox Combo6 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               ItemData        =   "Frm_SBRok.frx":0119
               Left            =   120
               List            =   "Frm_SBRok.frx":0126
               TabIndex        =   4
               Top             =   840
               Visible         =   0   'False
               Width           =   2175
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
               Height          =   420
               Left            =   120
               TabIndex        =   5
               Top             =   1680
               Width           =   975
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0E0FF&
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
               TabIndex        =   18
               Top             =   120
               Width           =   4095
               Begin MSDataListLib.DataCombo DataCombo4 
                  Height          =   420
                  Left            =   1800
                  TabIndex        =   3
                  Top             =   120
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   741
                  _Version        =   393216
                  ForeColor       =   16711680
                  Text            =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   12
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
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000040&
                  Height          =   285
                  Left            =   120
                  TabIndex        =   19
                  Top             =   120
                  Width           =   1455
               End
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Press F7 to set all rows by current cell value"
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
               Height          =   1050
               Left            =   2400
               TabIndex        =   20
               Top             =   840
               Width           =   1815
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0E0FF&
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
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   3255
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFF80&
               Caption         =   "Branch List"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   350
               Left            =   15
               TabIndex        =   16
               Top             =   0
               Width           =   3225
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFC0&
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
            Height          =   350
            Left            =   6840
            TabIndex        =   13
            Top             =   120
            Width           =   2175
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFC0&
               Caption         =   "Exhange List"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   285
               Left            =   540
               TabIndex        =   14
               Top             =   0
               Width           =   1185
            End
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
            Height          =   350
            Left            =   3480
            TabIndex        =   11
            Top             =   120
            Width           =   3255
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFC0&
               Caption         =   "Party List"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   345
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Width           =   3285
            End
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080C0FF&
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
            ForeColor       =   &H00000040&
            Height          =   225
            Left            =   5640
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H0080C0FF&
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
            ForeColor       =   &H00000040&
            Height          =   225
            Left            =   8040
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H0080C0FF&
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
            ForeColor       =   &H00000040&
            Height          =   225
            Left            =   2280
            TabIndex        =   8
            Top             =   480
            Width           =   1095
         End
         Begin MSComctlLib.ListView PartyLst 
            Height          =   2580
            Left            =   3480
            TabIndex        =   1
            Top             =   720
            Width           =   3285
            _ExtentX        =   5794
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Branch Name"
               Object.Width           =   6350
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView ItemLst 
            Height          =   2580
            Left            =   120
            TabIndex        =   0
            Top             =   720
            Width           =   3285
            _ExtentX        =   5794
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Party Name"
               Object.Width           =   5185
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2580
            Left            =   6840
            TabIndex        =   2
            Top             =   720
            Width           =   2325
            _ExtentX        =   4101
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
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Exchange Code"
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
         Begin MSComctlLib.ListView ListView2 
            Height          =   2580
            Left            =   9240
            TabIndex        =   35
            Top             =   720
            Width           =   2445
            _ExtentX        =   4313
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
               Size            =   12
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
            Left            =   11280
            TabIndex        =   21
            Top             =   360
            Width           =   2115
            WordWrap        =   -1  'True
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4695
         Left            =   120
         TabIndex        =   22
         Top             =   3720
         Width           =   17400
         _ExtentX        =   30692
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   -2147483628
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
         ColumnCount     =   17
         BeginProperty Column00 
            DataField       =   "Party"
            Caption         =   "Party Code"
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
            DataField       =   "PartyName"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "BBROKTYPE"
            Caption         =   "Sub BrokType"
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
            DataField       =   "BBROKRATE"
            Caption         =   "Sub BrokRate"
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
         BeginProperty Column08 
            DataField       =   "BROKRATE2"
            Caption         =   "BrokRate2"
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         BeginProperty Column13 
            DataField       =   "FMLYCODE"
            Caption         =   "Branch"
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
            DataField       =   "FMLYNAME"
            Caption         =   "BranchName"
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
            DataField       =   "EXCODE"
            Caption         =   "Exchange"
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
            DataField       =   "DeleteR"
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
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   2009.764
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1590.236
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1590.236
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column14 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column15 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   360
         Left            =   840
         TabIndex        =   25
         Top             =   4680
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
         Left            =   840
         TabIndex        =   26
         Top             =   5160
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
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   8820
      Left            =   120
      Top             =   720
      Width           =   17925
   End
End
Attribute VB_Name = "Frm_SBRok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LExCodes As String:             Dim LFmlyCodes As String:    Dim LParties As String:            Dim COPYFROM As String
Dim Items As String:                Dim ListIt As ListItem:     Dim GridColVal As String:           Dim CountRow As Double
Dim SearchRow As Double:            Public FlagBrok As Boolean: Dim LSettlementDt As String:        Dim RECGRID As ADODB.Recordset
Dim TempRec As ADODB.Recordset:     Dim flag As Boolean:        Public LDataCol As Integer:         Public Fb_Press As Byte
Dim AddMode As Boolean:             Dim LInstType As String:    Dim RecAcc As ADODB.Recordset:       Dim LItems As String
Sub ADD_NEW()
    Frame1.Enabled = True:    Frame2.Enabled = True:    Frame3.Enabled = True
    Call Get_Selection(1)
    ItemLst.SetFocus
End Sub
Sub CANCEL_REC()
    Fb_Press = 0
    ItemLst.Enabled = True
    PartyLst.Enabled = True
    ListView1.Enabled = True
    Check5.Enabled = True
    Check1.Enabled = True
    Check4.Enabled = True
    Command2.Enabled = True
    DataCombo4.Enabled = True
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
    DataCombo1.Enabled = True
    ItemDbComb.Enabled = True:
    Command2.Enabled = True: Frame1.Enabled = False: Combo1.Visible = False
    'Combo2.Visible = False
    Combo3.Visible = False: Combo4.Visible = False: Combo5.Visible = False: DataCombo2.Visible = False: DataCombo3.Visible = False: ItemDbComb.text = "": DataCombo1.text = ""
    Frame2.Enabled = False: Frame3.Enabled = False
    Call Get_Selection(13)
End Sub
Sub Save_Rec()
    On Error GoTo ERR1
    ItemLst.Enabled = False
    PartyLst.Enabled = False
    Check1.Enabled = False
    Check4.Enabled = False
    Command2.Enabled = False
    DataCombo4.Enabled = False
    Dim LastStDate, LStdTDate As Date
    MYSQL = "DELETE FROM PITSBROK WHERE UPTOSTDT IS NULL"
    Cnn.Execute MYSQL
    If Combo6.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf Combo6.ListIndex = 1 Then
        LInstType = "OPT"
    Else
        LInstType = "CSH"
    End If
    If IsDate(DataCombo4.text) Then
        LStdTDate = DateValue(DataCombo4.text)
    End If
        
    If RECGRID.RecordCount > 0 Then
        Set TempRec = RECGRID.Clone
        TempRec.MoveFirst
        LastStDate = TempRec!UPTOSTDT
    End If
    
    If RECGRID.RecordCount > 0 Then
        Cnn.BeginTrans: CNNERR = True
        MYSQL = "DELETE FROM PITSBROK WHERE COMPCODE =" & GCompCode & " AND PARTY IN  (" & LParties & ") "
        If LenB(LExCodes) <> 0 Then MYSQL = MYSQL & " AND EXCODE IN ( " & LExCodes & ") "
        MYSQL = MYSQL & "  AND INSTTYPE='" & LInstType & "' "
        If LenB(LFmlyCodes) <> 0 Then MYSQL = MYSQL & " AND FMLYCODE IN (" & LFmlyCodes & ") "
        MYSQL = MYSQL & " AND ITEMCODE IN  (" & LItems & ")"
        If IsDate(DataCombo4.text) Then
            MYSQL = MYSQL & " AND UpToStDt = '" & Format(LastStDate, "yyyy/MM/dd") & "' "
        Else
            MYSQL = MYSQL & " AND UpToStDt = '" & Format(GFinEnd, "yyyy/MM/dd") & "' "
        End If
        MYSQL = MYSQL & " AND UPTOSTDT >'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
        Cnn.Execute MYSQL
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If TempRec!DELETER = "N" Then
                If IsNull(TempRec!itemcode) Then
                Else
                    If Trim(TempRec!itemcode) = "" Then
                    Else
                        If IsDate(DataCombo4.text) Then
                            If DateValue(DataCombo4.text) > DateValue(GSysLockDt) Then
                                MYSQL = "INSERT INTO PITSBROK (COMPCODE   ,PARTY                  ,FMLYCODE                  ,EXCODE                   ,ITEMCODE                ,BROKTYPE,BROKRATE,SHTYPE,SHRATE                           ,UPTOSTDT, INSTTYPE,BROKRATE2)"
                                MYSQL = MYSQL & " VALUES (" & GCompCode & " ,'" & TempRec!PARTY & "','" & TempRec!FMLYCODE & "','" & TempRec!EXCODE & " ','" & TempRec!itemcode & "','" & Left(TempRec!BBROKTYPE, 1) & "'," & Val(TempRec!BBROKRATE) & ",'" & Left(TempRec!APPLYON, 1) & "'," & Val(TempRec!SHARE) & ",'" & Format(DataCombo4.text, "YYYY/MM/DD") & "','" & LInstType & "'," & TempRec!BROKRATE2 & ")"
                                Cnn.Execute MYSQL
                            Else
                                MsgBox "Sorry System Locked.  No Modification Allowed"
                                Exit Do
                            End If
                        Else
                            If DateValue(TempRec!UPTOSTDT) > DateValue(GSysLockDt) Then
                                MYSQL = "INSERT INTO PITSBROK (COMPCODE   ,PARTY                  ,FMLYCODE                  ,EXCODE                   ,ITEMCODE                ,BROKTYPE,BROKRATE,SHTYPE,SHRATE                           ,UPTOSTDT, INSTTYPE,BROKRATE2)"
                                MYSQL = MYSQL & " VALUES (" & GCompCode & " ,'" & TempRec!PARTY & "','" & TempRec!FMLYCODE & "','" & TempRec!EXCODE & " ','" & TempRec!itemcode & "','" & Left(TempRec!BBROKTYPE, 1) & "'," & Val(TempRec!BBROKRATE) & ",'" & Left(TempRec!APPLYON, 1) & "'," & Val(TempRec!SHARE) & ",'" & Format(TempRec!UPTOSTDT, "YYYY/MM/DD") & "','" & LInstType & "'," & TempRec!BROKRATE2 & ")"
                                Cnn.Execute MYSQL
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
        If Check2.Value = 1 Then
            LSettlementDt = GFinEnd
            Set GeneralRec = Nothing:        Set GeneralRec = New ADODB.Recordset
            MYSQL = "SELECT * FROM ACCFMLYD WHERE COMPCODE=" & GCompCode & " AND FMLYCODE IN (" & LFmlyCodes & ") AND PARTY IN (" & LParties & ") ORDER BY FMLYCODE,PARTY "
            GeneralRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
            While Not GeneralRec.EOF
                LParty = GeneralRec!PARTY
                LFmlyCode = GeneralRec!FMLYCODE
                MYSQL = "SELECT ITEMCODE,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE  = " & GCompCode & " AND EXCHANGECODE IN (" & LExCodes & ")"
                MYSQL = MYSQL & " AND ITEMCODE IN  (" & LItems & ")"
                MYSQL = MYSQL & "  ORDER BY ITEMCODE "
                Set GeneralRec1 = Nothing: Set GeneralRec1 = New ADODB.Recordset: GeneralRec1.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
                Do While Not GeneralRec1.EOF
                    DoEvents
                    MYSQL = "SELECT * FROM PITSBROK WHERE COMPCODE=" & GCompCode & " AND  FMLYCODE ='" & LFmlyCode & "' AND PARTY  ='" & LParty & "' AND ITEMCODE='" & GeneralRec1!itemcode & "' AND UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE='" & LInstType & "'"
                    Set REC1 = Nothing: Set REC1 = New ADODB.Recordset: REC1.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
                    If REC1.EOF Then
                        If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                            MYSQL = "INSERT INTO PITSBROK (COMPCODE,PARTY,FMLYCODE,EXCODE,ITEMCODE,BROKTYPE,BROKRATE,SHTYPE,SHRATE,UPTOSTDT,INSTTYPE)"
                            MYSQL = MYSQL & " VALUES (" & GCompCode & " ,'" & LParty & "','" & LFmlyCode & "','" & GeneralRec1!EXCHANGECODE & " ','" & GeneralRec1!itemcode & "','P',0,'N',0,'" & Format(LSettlementDt, "YYYY/MM/DD") & "','" & LInstType & "')"
                            Cnn.Execute MYSQL
                        End If
                    End If
                    GeneralRec1.MoveNext
                Loop
                GeneralRec.MoveNext
            Wend
        End If
        Cnn.CommitTrans: CNNERR = False
        Cnn.BeginTrans: CNNERR = True
        
        If BILL_GENERATION(GFinBegin, GFinEnd, "", LParties, LExCodes) Then
            Cnn.CommitTrans: CNNERR = False
        Else
            Cnn.RollbackTrans: CNNERR = False
        End If
        Chk_Billing
    End If
    Call CANCEL_REC
    Exit Sub
ERR1:
    If CNNERR = True Then
        MsgBox err.Description, vbCritical, "Error Number : " & err.Number
        'Resume
    End If
End Sub

Private Sub Check1_Click()
    For I = 1 To PartyLst.ListItems.Count
        If Check1.Value = 1 Then
            PartyLst.ListItems.Item(I).Checked = True
        Else
            PartyLst.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub Check2_Click()
'    If Check2.Value = 1 Then
'        MYSQL = "SELECT DISTINCT ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC, CTR_D AS CT ,ACCFMLY AS AF WHERE ACC.COMPCODE=" & GCompCode  & " AND ACC.COMPCODE = CT.COMPCODE AND ACC.COMPCODE = AF.COMPCODE AND CT.USERID = AF.FMLYCODE AND AF.FMLYHEAD = ACC.AC_CODE ORDER BY ACC.NAME"
'    Else
'        MYSQL = "SELECT DISTINCT PB.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC,PEXBROK AS PB WHERE ACC.COMPCODE=" & GCompCode  & " AND ACC.COMPCODE = PB.COMPCODE AND ACC.AC_CODE = PB.AC_CODE ORDER BY ACC.NAME"
'    End If
'    Set RecAcc = Nothing
'    Set RecAcc = New ADODB.Recordset
'    RecAcc.Open MYSQL, cnn, adOpenStatic, adLockReadOnly
'    If Not RecAcc.EOF Then
'        PartyLst.ListItems.clear
'        While Not RecAcc.EOF
'            Set ListIt = PartyLst.ListItems.ADD(, , RecAcc!Name)
'            ListIt.SubItems(1) = RecAcc!AC_CODE
'            RecAcc.MoveNext
'        Wend
'    End If
End Sub
Private Sub Check3_Click()
    For I = 1 To ListView2.ListItems.Count
        If Check3.Value = 1 Then
            ListView2.ListItems.Item(I).Checked = True
        Else
            ListView2.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub

Private Sub Check4_Click()
    For I = 1 To ItemLst.ListItems.Count
        If Check4.Value = 1 Then
            ItemLst.ListItems.Item(I).Checked = True
        Else
            ItemLst.ListItems.Item(I).Checked = False
        End If
    Next I
    Call ItemLst_Click
End Sub
Private Sub Check5_Click()
For I = 1 To ListView1.ListItems.Count
    If Check5.Value = 1 Then
        ListView1.ListItems.Item(I).Checked = True
    Else
        ListView1.ListItems.Item(I).Checked = False
    End If
Next I
Call ListView1_Click
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
        RECGRID!APPLYON = Combo5.text
        DataGrid1.Col = 10
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
    Dim TRec As ADODB.Recordset
    ItemLst.Enabled = False
    PartyLst.Enabled = False
    ListView1.Enabled = False
    Check5.Enabled = False
    Check1.Enabled = False
    Check4.Enabled = False
    LParties = vbNullString
    For J = 1 To PartyLst.ListItems.Count
        If PartyLst.ListItems(J).Checked = True Then
            If LenB(LParties) > 0 Then LParties = LParties & ", "
            LParties = LParties & "'" & PartyLst.ListItems(J).SubItems(1) & "'"
        End If
    Next
    If Combo6.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf Combo6.ListIndex = 1 Then
        LInstType = "OPT"
    ElseIf Combo6.ListIndex = 2 Then
        LInstType = "CSH"
    End If
    
    If LParties = vbNullString Then
        MsgBox "Please Select Party.", vbCritical:
        ItemLst.Enabled = True:        PartyLst.Enabled = True
        ListView1.Enabled = True:        Check5.Enabled = True
        Check1.Enabled = True:        Check4.Enabled = True
        PartyLst.Enabled = True:        PartyLst.SetFocus:
        Exit Sub
    End If
    LFmlyCodes = vbNullString
    For J = 1 To ItemLst.ListItems.Count
        If ItemLst.ListItems(J).Checked = True Then
            If LenB(LFmlyCodes) > 0 Then LFmlyCodes = LFmlyCodes & ", "
            LFmlyCodes = LFmlyCodes & "'" & ItemLst.ListItems(J).SubItems(1) & "'"
        End If
    Next
    
    LExCodes = vbNullString
    For J = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(J).Checked = True Then
            If LenB(LExCodes) > 0 Then LExCodes = LExCodes & ", "
            LExCodes = LExCodes & "'" & ListView1.ListItems(J).SubItems(2) & "'"
        End If
    Next
    If LenB(LExCodes) = 0 Then
        MsgBox "Please Select Exchange .", vbCritical:
        ListView1.Enabled = True
        ItemLst.Enabled = True
        PartyLst.Enabled = True
        ListView1.Enabled = True
        Check5.Enabled = True
        Check1.Enabled = True
        Check4.Enabled = True
        ListView1.SetFocus
        Exit Sub
    End If
    LItems = vbNullString
    For J = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(J).Checked = True Then
            If LenB(LItems) > 0 Then LItems = LItems & ", "
            LItems = LItems & "'" & ListView2.ListItems(J).SubItems(1) & "'"
        End If
    Next
    If LItems = "" Then
        MsgBox "Please Select Commodity/Script", vbCritical:
        ListView1.Enabled = True
        ItemLst.Enabled = True
        PartyLst.Enabled = True
        ListView1.Enabled = True
        Check5.Enabled = True
        Check1.Enabled = True
        Check4.Enabled = True
        ListView1.SetFocus
        Exit Sub
    End If
    
    CountRow = -1
    Call RecSet
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
    MYSQL = "SELECT AM.NAME,A.PARTY,A.ITEMCODE,B.ITEMNAME,B.EXCHANGECODE,A.BROKTYPE AS BBROKTYPE,A.BROKRATE AS BBROKRATE,A.BROKRATE2,A.SHTYPE,A.SHRATE,A.FMLYCODE,FM.FMLYNAME,A.UPTOSTDT,A.SHTYPE,A.SHRATE "
    MYSQL = MYSQL & " FROM PITSBROK AS A, ITEMMAST AS B ,ACCOUNTD AS AM,ACCFMLY FM "
    MYSQL = MYSQL & " WHERE A.COMPCODE=" & GCompCode & " AND A.COMPCODE=B.COMPCODE AND A.ITEMCODE=B.ITEMCODE "
    MYSQL = MYSQL & " AND A.PARTY IN  (" & LParties & ") "
    MYSQL = MYSQL & " AND A.INSTTYPE='" & LInstType & "' "
    MYSQL = MYSQL & " AND A.COMPCODE = AM.COMPCODE AND A.PARTY=AM.AC_CODE  AND A.COMPCODE =FM.COMPCODE "
    MYSQL = MYSQL & " AND FM.FMLYCODE = A.FMLYCODE  "
    If LFmlyCodes <> "" Then MYSQL = MYSQL & " AND A.FMLYCODE   IN (" & LFmlyCodes & ")  "
    MYSQL = MYSQL & "AND B.EXCHANGECODE  IN (" & LExCodes & ")  "
    MYSQL = MYSQL & "AND B.ITEMCODE   IN (" & LItems & ")  "
    If IsDate(DataCombo4.text) Then
        MYSQL = MYSQL & " AND A.UPTOSTDT = '" & Format(DataCombo4.text, "yyyy/MM/dd") & "'  "
    Else
        MYSQL = MYSQL & " AND A.UPTOSTDT = '" & Format(GFinEnd, "yyyy/MM/dd") & "'  "
    End If
    MYSQL = MYSQL & "ORDER BY AM.NAME,B.EXCHANGECODE,A.ITEMCODE,A.UPTOSTDT "
    GeneralRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not GeneralRec.EOF Then
        DataGrid1.Enabled = True
        Do While Not GeneralRec.EOF
            RECGRID.AddNew
            MICode = GeneralRec!itemcode
            RECGRID.Fields("PARTY") = GeneralRec!PARTY
            RECGRID.Fields("PARTYNAME") = GeneralRec!NAME
            RECGRID.Fields("ITEMCODE") = GeneralRec!itemcode
            RECGRID.Fields("ITEMNAME") = GeneralRec!ITEMName
            MYSQL = "SELECT BROKTYPE,BROKRATE FROM PITBROK WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & GeneralRec!PARTY & "' AND ITEMCODE ='" & GeneralRec!itemcode & "' AND INSTTYPE='" & LInstType & "' AND UPTOSTDT>='" & Format(GeneralRec!UPTOSTDT, "YYYY/MM/DD") & "'ORDER BY UPTOSTDT"
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
            If Not TRec.EOF Then
                MDefBrokType = TRec!BrokType
                LBrokRate = TRec!BROKRATE
            Else
                MYSQL = "SELECT BROKTYPE,BROKRATE FROM PEXBROK WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & GeneralRec!PARTY & "' AND  INSTTYPE='" & LInstType & "' AND UPTOSTDT>='" & Format(GeneralRec!UPTOSTDT, "YYYY/MM/DD") & "'ORDER BY UPTOSTDT"
                Set TRec = Nothing: Set TRec = New ADODB.Recordset
                TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
                If Not TRec.EOF Then
                    MDefBrokType = TRec!BrokType
                    LBrokRate = TRec!BROKRATE
                Else
                    MDefBrokType = "P"
                    LBrokRate = 0
                End If
            End If
            If MDefBrokType = "T" Then
                RECGRID.Fields("BROKTYPE") = "Transaction"
            ElseIf MDefBrokType = "O" Then
                RECGRID.Fields("BROKTYPE") = "Opening Sauda"
            ElseIf MDefBrokType = "C" Then
                RECGRID.Fields("BROKTYPE") = "Closing Sauda"
            ElseIf MDefBrokType = "Q" Then
                RECGRID.Fields("BROKTYPE") = "Qtywise IntraDay"
            ElseIf MDefBrokType = "P" Then
                RECGRID.Fields("BROKTYPE") = "Percentage wise"
            ElseIf MDefBrokType = "B" Then
                RECGRID.Fields("BROKTYPE") = "Bought"
            ElseIf MDefBrokType = "S" Then
                RECGRID.Fields("BROKTYPE") = "Sold"
            ElseIf MDefBrokType = "I" Then
                RECGRID.Fields("BROKTYPE") = "IntraDay Brokerage"
            ElseIf MDefBrokType = "V" Then
                RECGRID.Fields("BROKTYPE") = "Valuewise Intraday"
            ElseIf MDefBrokType = "D" Then
                RECGRID.Fields("BROKTYPE") = "Delivery Wise Brokerage"
            ElseIf MDefBrokType = "H" Then
                RECGRID.Fields("BROKTYPE") = "Higher Value Percentage Wise"
            ElseIf MDefBrokType = "L" Then
                RECGRID.Fields("BROKTYPE") = "LotWise Higher Value"
            ElseIf MDefBrokType = "Z" Then
                RECGRID.Fields("BROKTYPE") = "ZLotWise"
            ElseIf MDefBrokType = "R" Then
                RECGRID.Fields("BROKTYPE") = "RZLotWise Intraday"
            End If
            RECGRID.Fields("BROKRATE") = LBrokRate
            RECGRID.Fields("BBROKRATE") = IIf(IsNull(GeneralRec!BBROKRATE), 0, GeneralRec!BBROKRATE)
            
            If GeneralRec!BBROKTYPE = "P" Then
                RECGRID.Fields("BBROKTYPE") = "Percentage Wise"
            ElseIf GeneralRec!BBROKTYPE = "T" Then
                RECGRID.Fields("BBROKTYPE") = "Transaction Wise"
            ElseIf GeneralRec!BBROKTYPE = "O" Then
                RECGRID.Fields("BBROKTYPE") = "Opening Sauda"
            ElseIf GeneralRec!BBROKTYPE = "L" Then
                RECGRID.Fields("BBROKTYPE") = "LotWise Higher Value"
            ElseIf GeneralRec!BBROKTYPE = "D" Then
                RECGRID.Fields("BBROKTYPE") = "Delivery Wise"
            ElseIf GeneralRec!BBROKTYPE = "I" Then
                RECGRID.Fields("BBROKTYPE") = "Delivery Wise"
            ElseIf GeneralRec!BBROKTYPE = "Z" Then
                RECGRID.Fields("BBROKTYPE") = "Delivery Wise"
                
            End If
            RECGRID.Fields("SHARE") = IIf(IsNull(GeneralRec!SHRATE), 0, GeneralRec!SHRATE)
            If GeneralRec!SHTYPE = "N" Then
                RECGRID.Fields("ApplyOn") = "Net Amount"
            Else
                RECGRID.Fields("ApplyOn") = "Gross Amount"
            End If
            If IsNull(GeneralRec!UPTOSTDT) Then
            
                RECGRID.Fields("UPTOSTDT") = LSettlementDt
            Else
                If GeneralRec!UPTOSTDT = "" Then
                    RECGRID.Fields("UPTOSTDT") = LSettlementDt
                ElseIf DateValue(GeneralRec!UPTOSTDT) = DateValue("01/01/1900") Then
                    RECGRID.Fields("UPTOSTDT") = LSettlementDt
                Else
                    RECGRID.Fields("UPTOSTDT") = GeneralRec!UPTOSTDT
                End If
            End If
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            RECGRID.Fields("FMLYCODE") = GeneralRec!FMLYCODE & ""
            RECGRID.Fields("FMLYNAME") = GeneralRec!FMLYNAME & ""
            RECGRID.Fields("EXCODE") = GeneralRec!EXCHANGECODE & ""
            RECGRID.Fields("BROKRATE2") = GeneralRec!BROKRATE2 & ""
            RECGRID.Fields("DELETER") = "N"
            RECGRID.Update
            GeneralRec.MoveNext
        Loop
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh: DataCombo1.Enabled = False: ItemDbComb.Enabled = False: Command2.Enabled = False: RECGRID.MoveFirst: DataGrid1.SetFocus
        DataGrid1.LeftCol = 0
        Label3.Visible = True
        Command2.Enabled = False
    Else
        Label3.Visible = False
        MsgBox "No Record does not exists.for selected Commodity/Script", vbExclamation
        If MsgBox("Do you really want to apply Seperate Brokergae for selected Commodity", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
            MYSQL = " SELECT FMLYCODE,PARTY FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND FMLYCODE IN (" & LFmlyCodes & ") AND PARTY IN (" & LParties & ") ORDER BY FMLYCODE,PARTY  "
            Set PartyRec = Nothing
            Set PartyRec = New ADODB.Recordset
            PartyRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
            Do While Not PartyRec.EOF
                MYSQL = "SELECT ITEMCODE,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE IN (" & LItems & ")"
                Set ItemRec = Nothing
                Set ItemRec = New ADODB.Recordset
                ItemRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
                Do While Not ItemRec.EOF
                    DoEvents
                    MYSQL = "INSERT INTO PITSBROK (COMPCODE   ,PARTY                  ,FMLYCODE                  ,EXCODE                   ,ITEMCODE                ,BROKTYPE,BROKRATE,SHTYPE,SHRATE                           ,UPTOSTDT, INSTTYPE,BROKRATE2)"
                    MYSQL = MYSQL & " VALUES (" & GCompCode & " ,'" & PartyRec!PARTY & "','" & PartyRec!FMLYCODE & "','" & ItemRec!EXCHANGECODE & " ','" & ItemRec!itemcode & "','P'," & Val(0) & ",'N'," & Val(0) & ",'" & Format(GFinEnd, "YYYY/MM/DD") & "','" & LInstType & "',0)"
                    Cnn.Execute MYSQL
                    ItemRec.MoveNext
                Loop
                PartyRec.MoveNext
            Loop
        End If
        Call CANCEL_REC
        Command2.Enabled = True
    End If
    
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
        DataCombo2.BoundText = RECGRID!itemcode & ""
        Sendkeys "%{DOWN}"
    End If
End Sub
Private Sub DataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
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
            flag = True
            RECGRID!itemcode = DataCombo2.BoundText
            RECGRID!ITEMName = DataCombo2.text
            Dim RecTemp As ADODB.Recordset
            Set RecTemp = RECGRID.Clone
            GridColVal = ""
            RecTemp.MoveFirst
            GridColVal = DataCombo2.BoundText
            Do While Not RECGRID.EOF
                If RecTemp!itemcode = GridColVal Then
                    GridColVal = IIf(IsNull(RecTemp!BrokType), "Transaction", RecTemp!BrokType)
                    Exit Do
                End If
                RecTemp.MoveNext
            Loop
                If GridColVal = "" Then
                    RECGRID!BrokType = "Transaction"
                Else
                    RECGRID!BrokType = GridColVal
                End If
                RECGRID!BROKRATE = "0.000000"
                RECGRID!BBROKRATE = "0.000000"
                RECGRID!STDRATE = "0.00"
                RECGRID!TRANRATE = "0.00"
                
                GridColVal = ""
                RecTemp.MoveFirst
                GridColVal = DataCombo2.BoundText
                Do While Not RECGRID.EOF
                    If RecTemp!itemcode = GridColVal Then
                        GridColVal = IIf(IsNull(RecTemp!TRANTYPE), "Transaction", RecTemp!TRANTYPE)
                        Exit Do
                    End If
                    RecTemp.MoveNext
                Loop
                If GridColVal = "" Then
                    RECGRID!TRANTYPE = "Transaction"
                Else
                    RECGRID!TRANTYPE = GridColVal
                End If
                RECGRID!UPTOSTDT = ""  'LSettlementDt
                RECGRID!MARTYPE = "Value Wise (In %)"
                RECGRID!BROKRATE2 = 0
                RECGRID!MarRate = 0
                RECGRID!PARTY = Label7.Caption
                TempParties = Mid(LParties, Val(InStr(LParties, "'")) + 1, Val(Len(LParties)))
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
                LGridRow = DataGrid1.Row: LGridCol = DataGrid1.Col: GridColVal = RECGRID!itemcode: SearchRow = RECGRID!New
                RECGRID.MoveFirst
                Do While Not RECGRID.EOF
                    If RECGRID!itemcode = GridColVal Then
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
Private Sub Form_Load()
    FlagBrok = False
    'Last Settlement Date
    Combo6.ListIndex = 0
    If GOptions = "Y" Then
        Combo6.Visible = True
    End If
    LSettlementDt = "": Set Rec = Nothing: Set Rec = New ADODB.Recordset
    Rec.Open "SELECT MAX(SETDATE) AS MAXSETTLEDATE FROM SETTLE WHERE COMPCODE = " & GCompCode & "", Cnn, adOpenKeyset, adLockReadOnly
    If Not Rec.EOF Then LSettlementDt = Rec!MaxSettleDate
    Set RecAcc = Nothing
    Set RecAcc = New ADODB.Recordset
    MYSQL = "SELECT DISTINCT PB.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC,PEXBROK AS PB WHERE ACC.COMPCODE =" & GCompCode & " AND ACC.COMPCODE =PB.COMPCODE AND ACC.AC_CODE =PB.AC_CODE ORDER BY ACC.NAME"
    RecAcc.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    DataGrid1.Enabled = False
    If Not RecAcc.EOF Then
        While Not RecAcc.EOF
            Set ListIt = PartyLst.ListItems.Add(, , RecAcc!NAME)
            ListIt.SubItems(1) = RecAcc!AC_CODE
            RecAcc.MoveNext
        Wend
        Call Get_Selection(13)
        Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
        GeneralRec.Open "SELECT FMLYCODE,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " ORDER BY FMLYNAME ", Cnn, adOpenKeyset, adLockReadOnly
        If Not GeneralRec.EOF Then
            Set DataCombo2.RowSource = GeneralRec
            DataCombo2.ListField = "FMLYNAME"
            DataCombo2.BoundColumn = "FMLYCODE"
            While Not GeneralRec.EOF
                Set ListIt = ItemLst.ListItems.Add(, , GeneralRec!FMLYNAME)
                ListIt.SubItems(1) = GeneralRec!FMLYCODE
                GeneralRec.MoveNext
            Wend
        End If
        Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
        MYSQL = "SELECT EXCODE,EXNAME FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE "
        GeneralRec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
        If Not GeneralRec.EOF Then
            GeneralRec.MoveFirst
            ListView1.ListItems.clear
            ListView1.Enabled = True: Check4.Enabled = True
            Do While Not GeneralRec.EOF
                If (GeneralRec!EXCODE = "EQ" Or GeneralRec!EXCODE = "BEQ") Then
                    Combo6.Visible = True
                End If
                ListView1.ListItems.Add , , GeneralRec!EXCODE
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , GeneralRec!EXNAME
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , GeneralRec!EXCODE
                GeneralRec.MoveNext
            Loop
            
            If GeneralRec.RecordCount = 1 Then
                Check4.Value = 1: ListView1.TabStop = False: Check4.TabStop = False
                Call Check4_Click
            Else
                ListView1.TabStop = True: Check4.TabStop = True
            End If
        Else
            ListView1.Enabled = False: Check4.Enabled = False
        End If
        If Not GeneralRec.EOF Then
            Set DataCombo2.RowSource = GeneralRec
            DataCombo2.ListField = "ITEMNAME"
            DataCombo2.BoundColumn = "ITEMCODE"
            While Not GeneralRec.EOF
                Set ListIt = ItemLst.ListItems.Add(, , GeneralRec!ITEMName)
                ListIt.SubItems(1) = GeneralRec!itemcode
                GeneralRec.MoveNext
            Wend
        End If
        Set GeneralRec = Nothing: Set GeneralRec = New ADODB.Recordset
        GeneralRec.Open "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PITSBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT", Cnn, adOpenKeyset, adLockReadOnly
        If Not GeneralRec.EOF Then
            Set DataCombo3.RowSource = GeneralRec
            DataCombo3.ListField = "CONDATE"
            DataCombo3.BoundColumn = "CONDATE"
            Set DataCombo4.RowSource = GeneralRec
            DataCombo4.ListField = "CONDATE"
            DataCombo4.BoundColumn = "CONDATE"
        End If
    End If
    Call CANCEL_REC
End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "PARTY", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "PARTYNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
    RECGRID.Fields.Append "ITEMNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BBROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BBROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "UPTOSTDT", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "APPLYON", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "SHARE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "New", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "FMLYCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "FMLYNAME", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "EXCODE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "DELETER", adVarChar, 1, adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And DataGrid1.Col = 6 Then ' BROKTYPE
        Combo1.Visible = True: Combo1.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 10 Then 'APPLYON
        Combo5.Visible = True: Combo5.SetFocus
    ElseIf KeyCode = vbKeyF2 And Shift = 2 Then 'ctrl + f
        Set GeneralRec = Nothing
        Set GeneralRec = New ADODB.Recordset
        GeneralRec.Open "SELECT MAX(UPTOSTDT) FROM PITBROK WHERE COMPCODE =" & GCompCode & " AND ITEMCODE='" & RECGRID!itemcode & "' AND AC_CODE='" & DataCombo1.BoundText & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not GeneralRec.EOF Then
            If Len(Trim(GeneralRec.Fields(0) & "")) > 1 Then
                MUPTOSTDT = GeneralRec.Fields(0)
                Set GeneralRec = Nothing
                Set GeneralRec = New ADODB.Recordset
                GeneralRec.Open "SELECT MAX(SETDATE) FROM SETTLE WHERE COMPCODE =" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
                If GeneralRec.Fields(0) <= MUPTOSTDT Then
                    Exit Sub
                End If
            End If
        Else
            Set GeneralRec = Nothing:            Set GeneralRec = New ADODB.Recordset
            GeneralRec.Open "SELECT MAX(UPTOSTDT) FROM PEXBROK WHERE COMPCODE =" & GCompCode & " AND  AC_CODE='" & DataCombo1.BoundText & "'", Cnn, adOpenForwardOnly, adLockReadOnly
            If Not GeneralRec.EOF Then
                If Len(Trim(GeneralRec.Fields(0) & "")) > 1 Then
                    MUPTOSTDT = GeneralRec.Fields(0)
                    Set GeneralRec = Nothing
                    Set GeneralRec = New ADODB.Recordset
                    GeneralRec.Open "SELECT MAX(SETDATE) FROM SETTLE WHERE COMPCODE =" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
                    If GeneralRec.Fields(0) <= MUPTOSTDT Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        itemcode = RECGRID!itemcode
        ITEMName = RECGRID!ITEMName
        BrokType = RECGRID!BrokType
        BROKRATE = RECGRID!BROKRATE
        BBROKRATE = RECGRID!BBROKRATE
        BROKRATE2 = RECGRID!BROKRATE2
        UPTOSTDT = RECGRID!UPTOSTDT
        RECGRID.MoveLast
        RECGRID.AddNew
        RECGRID!itemcode = itemcode
        RECGRID!ITEMName = ITEMName
        RECGRID!BrokType = BrokType
        RECGRID!BROKRATE = BROKRATE
        
        RECGRID!BBROKRATE = BROKRATE
        RECGRID!BROKRATE2 = BROKRATE2
        RECGRID!UPTOSTDT = "" 'UPTOSTDT
        RECGRID.Update
    ElseIf KeyCode = 13 And DataGrid1.Col = 16 Then
        DataGrid1.text = UCase(DataGrid1.text)
        If DataGrid1.text <> "Y" Then
            DataGrid1.text = "N"
        End If
        DataGrid1.Col = 16
        DataGrid1.SetFocus
    ElseIf KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
    'Press F7 to replace all rows with current cell value
    If KeyCode = 118 Then   'F7
        LGridRow = DataGrid1.Row
        LGridCol = DataGrid1.Col
        If DataGrid1.Col = 6 Then 'BBROKTYPE
            GridColVal = RECGRID!BBROKTYPE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BBROKTYPE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 7 Then 'BBROKRATE
            GridColVal = RECGRID!BBROKRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BBROKRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 8 Then 'BBROKRATE
            GridColVal = RECGRID!BROKRATE2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BBROKRATE2 = GridColVal
                RECGRID.MoveNext
            Wend
            
        ElseIf DataGrid1.Col = 11 Then 'APPLYON
            GridColVal = RECGRID!APPLYON
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!APPLYON = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 10 Then 'SHARE
            GridColVal = RECGRID!SHARE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!SHARE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 13 Then 'UPTOSTDT
            GridColVal = RECGRID!UPTOSTDT
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!UPTOSTDT = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 16 Then 'UPTOSTDT
            GridColVal = RECGRID!DELETER
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!DELETER = GridColVal
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
                Combo1.ListIndex = Val(5)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "D" Then
                Combo1.ListIndex = Val(6)
                
            End If
    ElseIf DataGrid1.Col = 6 Then
            If Mid(RECGRID!BrokType, 1, 1) = "T" Then
                Combo1.ListIndex = Val(0)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "P" Then
                Combo1.ListIndex = Val(1)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "O" Then
                Combo1.ListIndex = Val(2)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "I" Then
                Combo1.ListIndex = Val(3)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "Z" Then
                Combo1.ListIndex = Val(4)
            ElseIf Mid(RECGRID!BrokType, 1, 1) = "R" Then
                Combo1.ListIndex = Val(5)
            End If
    End If
    Combo1.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    Combo1.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    Combo1.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        LGridRow = DataGrid1.Row: LGridCol = DataGrid1.Col: GridColVal = RECGRID!itemcode: SearchRow = RECGRID!New
        If DataGrid1.Col = 6 Then
            RECGRID!BBROKTYPE = Combo1.text
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
Private Sub Form_Paint()
    If FlagBrok And MAc_Code <> "" Then
        DataCombo1.BoundText = MAc_Code
        Call Get_Selection(1)
        Call Command2_Click
    End If
End Sub

Private Sub ItemDbComb_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub ItemDbComb_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub ItemLst_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        Call ItemLst_Click
    End If
End Sub
Private Sub ItemLst_Click()
    Dim RecSauda As ADODB.Recordset
    LFmlyCodes = ""
    For I = 1 To ItemLst.ListItems.Count
        If ItemLst.ListItems(I).Checked = True Then
            LFmlyCodes = LFmlyCodes & "'"
            LFmlyCodes = LFmlyCodes & ItemLst.ListItems(I).ListSubItems(1)
            LFmlyCodes = LFmlyCodes & "'"
        End If
        If I < ItemLst.ListItems.Count Then
            If ItemLst.ListItems(I + 1).Checked = True And Len(LFmlyCodes) > Val(0) Then
                LFmlyCodes = LFmlyCodes & ", "
            End If
        End If
  Next I
  PartyLst.ListItems.clear
  If LFmlyCodes = "" Then Me.MousePointer = 0: Exit Sub
    MYSQL = "SELECT AC_CODE,NAME FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE IN (SELECT DISTINCT PARTY FROM ACCFMLYD WHERE COMPCODE =" & GCompCode & " AND FMLYCODE IN  (" & LFmlyCodes & "))  ORDER BY NAME "
    Set RecSauda = Nothing: Set RecSauda = New ADODB.Recordset: RecSauda.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
    While Not RecSauda.EOF
        Set ListIt = PartyLst.ListItems.Add(, , RecSauda!NAME)
        ListIt.SubItems(1) = RecSauda!AC_CODE
        RecSauda.MoveNext
    Wend
End Sub

Private Sub ListView1_Click()
    Dim RecSauda As ADODB.Recordset
    Dim ChkCount As Integer
    
    LExCodes = vbNullString
    ChkCount = 0
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = True Then
            ChkCount = ChkCount + 1
            If LenB(LExCodes) > 0 Then LExCodes = LExCodes & ", "
            LExCodes = LExCodes & "'" & ListView1.ListItems(I).ListSubItems(2) & "'"
        End If
    Next I
    If ChkCount = ListView1.ListItems.Count Then
        AllExcodes = True
    Else
        AllExcodes = False
    End If
    ListView2.ListItems.clear
    If LenB(LExCodes) = 0 Then Me.MousePointer = 0: Exit Sub
    MYSQL = "SELECT ITEMCODE,ITEMNAME,LOT,EXCHANGECODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND EXCHANGECODE in (" & LExCodes & ")  ORDER BY EXCHANGECODE,ITEMNAME"
    Set RecSauda = Nothing: Set RecSauda = New ADODB.Recordset: RecSauda.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
    While Not RecSauda.EOF
        Set ListIt = ListView2.ListItems.Add(, , RecSauda!ITEMName)
        ListIt.SubItems(1) = RecSauda!itemcode
        RecSauda.MoveNext
    Wend

End Sub
