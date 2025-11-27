VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSauda 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   16020
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtItemID 
      Height          =   495
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   8880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TxtExID 
      Height          =   495
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   120
      TabIndex        =   34
      Top             =   3000
      Width           =   14105
      Begin MSDataListLib.DataCombo FItemCombo 
         Height          =   360
         Left            =   5880
         TabIndex        =   39
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   0
         Text            =   ""
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
      Begin vcDateTimePicker.vcDTP vcDTP2 
         Height          =   375
         Left            =   9240
         TabIndex        =   40
         Top             =   120
         Width           =   1515
         _ExtentX        =   2672
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
         Value           =   43235.7562615741
      End
      Begin VB.CommandButton CmdFilterOk 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         TabIndex        =   42
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox TxtFilterCode 
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
         TabIndex        =   37
         Top             =   120
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5175
         Left            =   0
         TabIndex        =   35
         Top             =   600
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP3 
         Height          =   375
         Left            =   11160
         TabIndex        =   41
         Top             =   120
         Width           =   1515
         _ExtentX        =   2672
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
         Value           =   43235.7562615741
      End
      Begin MSDataListLib.DataCombo FEXCombo 
         Height          =   360
         Left            =   4080
         TabIndex        =   38
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   0
         Text            =   ""
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ExCode"
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
         Left            =   3240
         TabIndex        =   49
         Top             =   203
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Left            =   5400
         TabIndex        =   46
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   10800
         TabIndex        =   45
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Maturity From"
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
         Left            =   8280
         TabIndex        =   44
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SaudaCode"
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
         Left            =   120
         TabIndex        =   36
         Top             =   173
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   615
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   14295
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   14655
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Sauda Master Setup"
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
            TabIndex        =   30
            Top             =   120
            Width           =   14100
         End
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2535
      Left            =   14280
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   14085
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   14655
         Begin VB.TextBox TxtSaudaId 
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
            Left            =   7560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   195
            Width           =   1550
         End
         Begin VB.TextBox TxtNoCode 
            Alignment       =   1  'Right Justify
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
            Left            =   5040
            MaxLength       =   15
            TabIndex        =   12
            Top             =   1560
            Width           =   1100
         End
         Begin VB.TextBox TxtRefLot 
            Alignment       =   1  'Right Justify
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
            Left            =   10560
            TabIndex        =   14
            Top             =   1560
            Width           =   1245
         End
         Begin VB.TextBox TxtBrokLot 
            Alignment       =   1  'Right Justify
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
            Left            =   12840
            TabIndex        =   15
            Top             =   1560
            Width           =   1125
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            Height          =   495
            Left            =   2160
            TabIndex        =   27
            Top             =   1440
            Visible         =   0   'False
            Width           =   1695
            Begin VB.OptionButton Option6 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Put"
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
               Left            =   840
               TabIndex        =   11
               Top             =   195
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Call"
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
               Left            =   120
               TabIndex        =   10
               Top             =   195
               Width           =   735
            End
         End
         Begin VB.TextBox TxtStrikePrice 
            Alignment       =   1  'Right Justify
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
            Left            =   930
            TabIndex        =   9
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFC0C0&
            Height          =   615
            Left            =   10560
            TabIndex        =   26
            Top             =   720
            Width           =   3375
            Begin VB.OptionButton Option7 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Cash"
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
               Left            =   2400
               TabIndex        =   8
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Option"
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
               Left            =   1200
               TabIndex        =   7
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Future"
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
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.TextBox TxtTradeableLot 
            Alignment       =   1  'Right Justify
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
            Left            =   7560
            TabIndex        =   13
            Top             =   1560
            Width           =   1550
         End
         Begin VB.TextBox TxtSaudaName 
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
            Height          =   405
            Left            =   10560
            MaxLength       =   50
            TabIndex        =   2
            Top             =   195
            Width           =   3375
         End
         Begin VB.TextBox TxtSaudaCode 
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
            Left            =   930
            MaxLength       =   50
            TabIndex        =   1
            Top             =   195
            Width           =   5175
         End
         Begin MSDataListLib.DataCombo DComboItem 
            Height          =   360
            Left            =   3480
            TabIndex        =   4
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   16711680
            Text            =   ""
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
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   375
            Left            =   7560
            TabIndex        =   5
            Top             =   840
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
            Value           =   37861.9121759259
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   360
            Left            =   930
            TabIndex        =   3
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   16711680
            Text            =   ""
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda Id"
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
            Index           =   3
            Left            =   6480
            TabIndex        =   47
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Closing ID"
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
            Left            =   3960
            TabIndex        =   43
            Top             =   1635
            Width           =   975
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref Lot"
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
            Left            =   9480
            TabIndex        =   33
            Top             =   1620
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ExCode"
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
            Index           =   5
            Left            =   120
            TabIndex        =   32
            Top             =   885
            Width           =   720
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "BrokLot"
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
            Left            =   12000
            TabIndex        =   31
            Top             =   1635
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Strike "
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
            TabIndex        =   28
            Top             =   1635
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inst.Type"
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
            Index           =   4
            Left            =   9480
            TabIndex        =   25
            Top             =   945
            Width           =   960
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Trade Lot"
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
            Left            =   6480
            TabIndex        =   23
            Top             =   1635
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maturity"
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
            Index           =   2
            Left            =   6480
            TabIndex        =   22
            Top             =   885
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item/Script"
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
            Index           =   1
            Left            =   2280
            TabIndex        =   21
            Top             =   885
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   9480
            TabIndex        =   20
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
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
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   255
            Width           =   495
         End
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   2460
      Left            =   120
      Top             =   720
      Width           =   14085
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S a u d a   L i s t"
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
      Height          =   345
      Left            =   14280
      TabIndex        =   17
      Top             =   3240
      Width           =   2445
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   5100
      Left            =   120
      Top             =   3600
      Width           =   13965
   End
End
Attribute VB_Name = "FrmSauda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte:        Dim SaudaRec As ADODB.Recordset:    Dim ItemRec As ADODB.Recordset:     Dim ExRec As ADODB.Recordset
Dim LScriptCode As String:      Dim LInstType As String:            Dim LOptType As String:             Dim LStrike As Double
Dim FItemCode As String:        Dim LFExCode As String:             Dim LNoCode As Long
Sub Add_Rec()
    Fb_Press = 1
    Call Get_Selection(1)
    Frame1.Enabled = True
    TxtSaudaCode.text = vbNullString
    TxtSaudaName.text = vbNullString
    DComboItem.BoundText = vbNullString
    DataCombo3.BoundText = vbNullString

    Frame12.Enabled = False

    TxtSaudaCode.SetFocus
End Sub
Sub Save_Rec()
    On Error GoTo err1
    Dim LSaudas As String:          Dim LLItemCode As String
    Dim TRec As ADODB.Recordset:    Dim LLOT As Double
    Dim LLotWise  As String:        Dim LSaudaID As Long
    Dim LExID As Integer:           Dim LItemID As Integer
    
    If Len(Trim(TxtSaudaCode.text)) < 1 Then
        MsgBox "Sauda Code required before saving record.", vbCritical, "Error"
        TxtSaudaCode.SetFocus
        Exit Sub
    End If
    If Len(Trim(TxtSaudaName.text)) < 1 Then
        MsgBox "Sauda Name required before saving record.", vbCritical, "Error"
        TxtSaudaName.SetFocus
        Exit Sub
    End If
    
    If LenB(DComboItem.BoundText) = 0 Then
        MsgBox "Item Name required before saving record.", vbCritical, "Error"
        DComboItem.SetFocus
        Exit Sub
    End If
    If LenB(DataCombo3.BoundText) = 0 Then
        MsgBox "Exchange Code required before saving record.", vbCritical, "Error"
        DataCombo3.SetFocus
        Exit Sub
    End If
    
    If Option3.Value = True Then
        LInstType = "FUT"
    ElseIf Option4.Value = True Then
        LInstType = "OPT"
    ElseIf Option7.Value = True Then
        LInstType = "CSH"
    End If
    If (LInstType = "FUT" Or LInstType = "CSH") Then
        LOptType = ""
        LStrike = 0
    Else
        If Option5.Value = True Then
            LOptType = "CE"
        Else
            LOptType = "PE"
        End If
        LStrike = Val(TxtStrikePrice.text)
    End If
    LNoCode = Val(TxtNoCode.text & vbNullString)
    
    CNNERR = True
    Cnn.BeginTrans
    
    mysql = "SELECT COMPCODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDACODE  ='" & TxtSaudaCode.text & "'"
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If TRec.EOF Then
        Call PInsert_Saudamast(TxtSaudaCode.text, TxtSaudaName.text, FItemCode, vcDTP1.Value, Val(TxtTradeableLot.text), Val(TxtBrokLot.text), _
        LNoCode, LInstType, LOptType, LStrike, LFExCode, Val(TxtRefLot.text), Val(TxtExID.text), Val(TxtItemID.text))
    Else
        mysql = "UPDATE SAUDAMAST SET "
        mysql = mysql & " SAUDANAME   ='" & TxtSaudaName.text & "',"
        mysql = mysql & " ITEMCODE    ='" & FItemCode & "',"
        mysql = mysql & " MATURITY    ='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "',"
        mysql = mysql & " TRADEABLELOT= " & Val(TxtTradeableLot.text) & ","
        mysql = mysql & " BROKLOT     = " & Val(TxtBrokLot.text) & ","
        mysql = mysql & " NOCODE     = " & LNoCode & ","
        mysql = mysql & " INSTTYPE    ='" & LInstType & "',"
        mysql = mysql & " OPTTYPE     ='" & LOptType & "',"
        mysql = mysql & " STRIKEPRICE = " & LStrike & ","
        mysql = mysql & " EXCODE      ='" & LFExCode & "',"
        mysql = mysql & " EXID        = " & Val(TxtExID.text) & ","
        mysql = mysql & " ITEMID      = " & Val(TxtItemID.text) & ","
        mysql = mysql & " REFLOT      = " & Val(TxtRefLot.text) & ""
        mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND SAUDACODE ='" & TxtSaudaCode.text & "'"
        Cnn.Execute mysql
    End If
    
    GETMAIN.Label1.Caption = "Updateing Contract Details"
    LSaudaID = Get_SaudaID(TxtSaudaCode.text)
    LLotWise = "N"
    ExRec.MoveFirst
    ExRec.Find "EXCODE='" & LFExCode & "'"
    If Not ExRec.EOF Then
        
        LLotWise = ExRec!LOTWISE
        TxtExID.text = ExRec!EXID
        
    End If
    If Fb_Press = 2 Then
        LExID = Get_ExID(LFExCode)
        LItemID = Get_ITEMID(FItemCode)
        LLOT = Get_LotSize(LItemID, LSaudaID, LExID)
        mysql = "UPDATE CTR_D SET ITEMCODE ='" & FItemCode & "', EXCODE ='" & LFExCode & "',EXID=" & Val(TxtExID.text) & ",ITEMID=" & Val(TxtItemID.text) & " , CALVAL=" & LLOT & " "
        mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND SAUDAID =" & Val(TxtSaudaId.text) & ""
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_M SET ITEMCODE ='" & FItemCode & "',EXCODE='" & LFExCode & "',EXID=" & Val(TxtExID.text) & ",ITEMID=" & Val(TxtItemID.text) & "  WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND SAUDAID =" & Val(TxtSaudaId.text) & ""
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_R SET ITEMCODE ='" & FItemCode & "',EXCODE='" & LFExCode & "',EXID=" & Val(TxtExID.text) & ",ITEMID=" & Val(TxtItemID.text) & "   WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND SAUDAID =" & Val(TxtSaudaId.text) & ""
        Cnn.Execute mysql
    
        mysql = "UPDATE INV_D SET ITEMCODE ='" & FItemCode & "',EXCODE='" & LFExCode & "',EXID=" & Val(TxtExID.text) & ",ITEMID=" & Val(TxtItemID.text) & "   WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND SAUDAID =" & Val(TxtSaudaCode.text) & ""
        Cnn.Execute mysql
    End If

    Cnn.CommitTrans
    GETMAIN.Label1.Caption = vbNullString
    LSaudas = "'" & TxtSaudaCode.text & "'"
    'CNNERR = True
    'Frame2.Enabled = False
    'Dim LSParties As String
    'LSParties = vbNullString
    'Set TRec = Nothing: Set TRec = New ADODB.Recordset
    'mysql = "SELECT DISTINCT PARTY FROM CTR_D  WHERE SAUDAID   = " & LSaudaID & " "
    'TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    'While Not TRec.EOF
    '    If LenB(LSParties) > 0 Then LSParties = LSParties & ", "
    '    LSParties = LSParties & "'" & TRec!PARTY & "'"
    '    TRec.MoveNext
    'Wend
    'Cnn.CommitTrans
    'CNNERR = False
    'If Len(LSParties) > 0 Then
     '   Cnn.BeginTrans
      '  If BILL_GENERATION(CDate(GFinBegin), CDate(GFinEnd), Str(LSaudaID), LSParties, Trim(TxtExID.text)) Then
       '     MsgBox " Bills Updated Succesfully"
        '    Cnn.CommitTrans
         '   CNNERR = False
        'Else
         '   Cnn.RollbackTrans: CNNERR = False
        'End If
     'End If
    'Call Chk_Billing
    Frame2.Enabled = True
    Call CANCEL_REC
    Exit Sub
err1:
    If CNNERR = True Then
        Cnn.RollbackTrans
        CNNERR = False
    End If
    
    MsgBox err.Description
End Sub
Sub CANCEL_REC()
    DComboItem.Locked = False:    TxtSaudaCode.Enabled = True
    Fb_Press = 0
    TxtSaudaCode.text = vbNullString:       TxtSaudaName.text = vbNullString
    TxtNoCode.text = "0":                   TxtExID.text = vbNullString
    TxtRefLot.text = "1.00":                TxtBrokLot.text = "1.00"
    Call Get_Selection(10):                 TxtTradeableLot.text = "1.00"
    DComboItem.BoundText = vbNullString:    DataCombo3.BoundText = vbNullString
    Frame12.Enabled = True:                 Frame1.Enabled = False
    TxtStrikePrice.text = "0.00"
    CmdFilterOk.Caption = "Go":             LNoCode = 0
    FillDataGrid
    
End Sub
Sub MODIFY_REC()
Dim Rec3 As ADODB.Recordset
    If LenB(TxtSaudaCode.text) > 0 Then
        'DataGrid1.Enabled = False
        Frame12.Enabled = False
        SaudaRec.MoveFirst
        SaudaRec.Find "SAUDACODE='" & TxtSaudaCode.text & "'", , adSearchForward
        If Not SaudaRec.EOF Then
            Frame1.Enabled = True:                      DComboItem.Locked = False
            TxtSaudaCode.Enabled = False:               TxtSaudaCode.text = SaudaRec!saudacode
            TxtSaudaId.text = SaudaRec!SAUDAID:         TxtSaudaName.text = SaudaRec!SAUDANAME
            FItemCode = SaudaRec!ITEMCODE:              LFExCode = IIf(IsNull(SaudaRec!excode), "MCX", SaudaRec!excode)
            DComboItem.BoundText = FItemCode:           DataCombo3.BoundText = LFExCode
            vcDTP1.Value = SaudaRec!MATURITY:           TxtTradeableLot.text = IIf(IsNull(SaudaRec!lot), 1, SaudaRec!lot)
            LNoCode = Val(SaudaRec!NOCODE & vbNullString)
            TxtNoCode.text = LNoCode:
            
            TxtExID.text = SaudaRec!EXID
            TxtItemID.text = SaudaRec!itemid
            TxtBrokLot.text = Format(IIf(IsNull(SaudaRec!BROKLOT), 1, SaudaRec!BROKLOT), "0.00")
            TxtRefLot.text = Format(IIf(IsNull(SaudaRec!REFLOT), 1, SaudaRec!REFLOT), "0.00")
            TxtTradeableLot.text = Format(TxtTradeableLot.text, "0.00")
            TxtStrikePrice.text = Format(IIf(IsNull(SaudaRec!STRIKEPRICE), 1, SaudaRec!STRIKEPRICE), "0.00")
            If IsNull(SaudaRec!INSTTYPE) Then
                LInstType = "FUT"
                Option3.Value = True
                Option4.Value = False
            Else
                If SaudaRec!INSTTYPE = "FUT" Then
                    LInstType = "FUT":                  Frame9.Visible = False:     Label13.Visible = False
                    TxtStrikePrice.Visible = False:     Option3.Value = True:       Option4.Value = False
                    Option7.Value = False
                ElseIf SaudaRec!INSTTYPE = "OPT" Then
                    LInstType = "OPT":                  Frame9.Visible = True:      Label13.Visible = True
                    TxtStrikePrice.Visible = True:      Option3.Value = False:      Option4.Value = True
                    Option7.Value = False
                    Label13.Caption = "Strike Price"
                ElseIf SaudaRec!INSTTYPE = "CSH" Then
                    LInstType = "CSH":                  Frame9.Visible = False:     Label13.Visible = False
                    TxtStrikePrice.Visible = False:     Option3.Value = False:      Option4.Value = False
                    Option7.Value = True
                End If
            End If
            If IsNull(SaudaRec!OPTTYPE) Then
                LOptType = vbNullString:                Option5.Value = True:                Option6.Value = False
            Else
                If SaudaRec!OPTTYPE = "CE" Then
                    LOptType = "CE":                   Option5.Value = True:                    Option6.Value = False
                Else
                    LOptType = "PE":                    Option5.Value = False:                    Option6.Value = True
                End If
            End If
            TxtSaudaName.SetFocus
            If Fb_Press = 3 Then    ''FOR DELETE
                If MsgBox("You are about to Delete this record. Confirm Delete?", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
                    Set Rec3 = Nothing
                    Set Rec3 = New ADODB.Recordset
                    mysql = "SELECT COMPCODE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND SAUDA='" & TxtSaudaCode.text & "'"
                    Rec3.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                    If Not Rec3.EOF Then
                        MsgBox "Transaction Exists Can't Delete Sauda.", vbExclamation, "Error"
                    Else
                        mysql = "DELETE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDACODE='" & TxtSaudaCode.text & "'"
                        Cnn.Execute mysql
                    End If
                End If
                Call CANCEL_REC
                'DataList1.Locked = False
                'DataGrid1.Enabled = True
                Frame12.Enabled = True
                DataGrid1.SetFocus
                'DataList1.SetFocus
            End If
        Else
            MsgBox "Please Select Sauda.", vbInformation
            Call CANCEL_REC
            'DataGrid1.Enabled = True
            Frame12.Enabled = True
            DataGrid1.SetFocus
        End If
    End If
End Sub

Private Sub CmdFilterOK_Click()
If CmdFilterOk.Caption = "Go" Then
    CmdFilterOk.Caption = "Clear"
    FillDataGrid
   
Else
    TxtFilterCode.text = vbNullString
    FillDataGrid
    CmdFilterOk.Caption = "Go"
End If
End Sub
Private Sub DComboItem_GotFocus()
    If Fb_Press = 1 Then Sendkeys "%{DOWN}"
End Sub

Private Sub DataCombo3_GotFocus()
    If Fb_Press = 1 Then Sendkeys "%{DOWN}"
End Sub

Private Sub DataCombo2_GotFocus()
     Sendkeys "%{DOWN}"
End Sub

Private Sub DComboItem_Validate(Cancel As Boolean)
Dim Rec3 As ADODB.Recordset
If LenB(DComboItem.BoundText) = 0 Then
    MsgBox "Please Select Item/script"
    Cancel = True
    Exit Sub
End If
FItemCode = DComboItem.BoundText
mysql = "SELECT EXCHANGECODE,EXID,ITEMID FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND  ITEMCODE ='" & FItemCode & "' AND EXCHANGECODE ='" & LFExCode & "'"
Set Rec3 = Nothing
Set Rec3 = New ADODB.Recordset
Rec3.Open mysql, Cnn, adOpenStatic, adLockReadOnly

If Rec3.EOF Then
    MsgBox "Invalid Item/Script"
    DComboItem.text = vbNullString
    Cancel = True
Else
    TxtExID.text = Trim(str(Rec3!EXID))
    TxtItemID.text = Trim(str(Rec3!itemid))
End If

End Sub

Private Sub DataCombo3_Validate(Cancel As Boolean)
Dim Rec3 As ADODB.Recordset
LFExCode = DataCombo3.BoundText
If LenB(LFExCode) = 0 Then
    MsgBox " Exchange can not be blank"
    Cancel = True
Else
    mysql = "SELECT EXID,EXCODE FROM EXMAST  WHERE COMPCODE =" & GCompCode & " AND  EXCODE ='" & LFExCode & "'"
    Set Rec3 = Nothing
    Set Rec3 = New ADODB.Recordset
    Rec3.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Rec3.EOF Then
        DataCombo3.BoundText = vbNullString
        MsgBox "Invalid Exchange"
        Cancel = True
    Else
        TxtExID.text = Rec3!EXID
        Set ItemRec = Nothing
        Set ItemRec = New ADODB.Recordset
        ItemRec.Open "SELECT ITEMID,EXID,ITEMCODE,ITEMNAME AS ITEMNAME FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND EXCHANGECODE ='" & LFExCode & "' ORDER BY ITEMCODE", Cnn, adOpenKeyset, adLockReadOnly
        Set DComboItem.RowSource = ItemRec
        DComboItem.BoundColumn = "ITEMCODE"
        DComboItem.ListField = "ITEMNAME"
    End If
End If
End Sub
Private Sub DataGrid1_Click()
If SaudaRec.RecordCount > 0 Then
    If SaudaRec.EOF Then SaudaRec.MoveFirst
    DataGrid1.Col = 2
    TxtSaudaCode.text = DataGrid1.text
    DataGrid1.Col = 11
    TxtSaudaName.text = DataGrid1.text
End If
End Sub
Private Sub DataGrid1_DblClick()
    Call Get_Selection(2)
    Fb_Press = 2
    Call MODIFY_REC
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    Dim LChar As String
    LChar = UCase(Chr(KeyAscii))
    If KeyAscii = 13 Then
        Call Get_Selection(2)
        Fb_Press = 2
        Call MODIFY_REC
    Else
        If Not SaudaRec.EOF Then
            If Left$(SaudaRec!saudacode, 1) = LChar Then
                SaudaRec.MoveNext
            ElseIf LChar > Left$(SaudaRec!saudacode, 1) Then
                Do While Not SaudaRec.EOF
                    If Left$(SaudaRec!saudacode, 1) <> LChar Then
                        SaudaRec.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
            Else
                SaudaRec.MoveFirst
                Do While Not SaudaRec.EOF
                    If Left$(SaudaRec!saudacode, 1) <> LChar Then
                        SaudaRec.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If SaudaRec.EOF Then SaudaRec.MoveFirst
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub Form_Load()
    vcDTP2.Value = GFinBegin
    vcDTP3.Value = GFinEnd + 365
    
    Set ItemRec = Nothing:    Set ItemRec = New ADODB.Recordset
    ItemRec.Open "SELECT ITEMCODE, ITEMNAME AS ITEMNAME FROM ITEMMAST WHERE COMPCODE =" & GCompCode & "  ORDER BY ITEMCODE", Cnn, adOpenKeyset, adLockReadOnly
    Set DComboItem.RowSource = ItemRec
    DComboItem.BoundColumn = "ITEMCODE"
    DComboItem.ListField = "ITEMNAME"
    Set FItemCombo.RowSource = ItemRec
    FItemCombo.BoundColumn = "ITEMCODE"
    FItemCombo.ListField = "ITEMNAME"
    
    
    Set ExRec = Nothing:    Set ExRec = New ADODB.Recordset
    ExRec.Open "SELECT EXID,EXCODE,EXNAME,LOTWISE FROM EXMAST WHERE COMPCODE =" & GCompCode & "  ORDER BY EXCODE", Cnn, adOpenKeyset, adLockReadOnly
    Set DataCombo3.RowSource = ExRec
    
    Set FEXCombo.RowSource = ExRec
    
    DataCombo3.BoundColumn = "EXCODE"
    DataCombo3.ListField = "EXCODE"
    
    FEXCombo.BoundColumn = "EXCODE"
    FEXCombo.ListField = "EXCODE"
    Label13.Visible = True
    TxtStrikePrice.Visible = True
    Frame9.Visible = True
    Call CANCEL_REC
End Sub
Private Sub Form_Paint()
'    Me.BackColor = GETMAIN.BackColor
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CRViewer1.Visible = True Then
        Call Get_Selection(10)
        CRViewer1.Visible = False
        Cancel = 1
    Else
        Call CANCEL_REC
        GETMAIN.StatusBar1.Panels(1).text = vbNullString
        Unload Me
    End If
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label17.ForeColor = &H0&
End Sub
Private Sub Label17_Click()
    Call List_Sauda
End Sub
Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label17.ForeColor = &HC00000
End Sub
Private Sub Option3_Validate(Cancel As Boolean)
If Option3.Value = True Then
    LInstType = "FUT"
    Frame9.Enabled = False
    Frame9.Visible = False
    Label13.Visible = False
    Frame9.Enabled = False
    Label13.Visible = False
    TxtStrikePrice.Visible = False
ElseIf Option4.Value = True Then
    LInstType = "OPT"
    Frame9.Enabled = True
    Frame9.Visible = True
ElseIf Option7.Value = True Then
    LInstType = "CSH"
    Frame9.Enabled = False
    Frame9.Visible = False
    Label13.Visible = False
End If
End Sub
Private Sub Option4_Validate(Cancel As Boolean)
If Option4.Value = True Then
    Frame9.Enabled = True
    Frame9.Visible = True
    Label13.Visible = True
    TxtStrikePrice.Visible = True
Else
    Frame9.Enabled = False
    Label13.Visible = False
    TxtStrikePrice.Visible = False
    Frame9.Visible = False
End If
End Sub
Private Sub Option7_Validate(Cancel As Boolean)
If Option7.Value = True Then
    Frame9.Enabled = False
    Frame9.Visible = False
    Label13.Visible = False
    TxtStrikePrice.Visible = False
End If
End Sub
Private Sub TxtNoCode_GotFocus()
    TxtNoCode.SelLength = Len(TxtNoCode.text)
End Sub
Private Sub TxtSaudaCode_Validate(Cancel As Boolean)
    Dim TRec As ADODB.Recordset
    If Fb_Press = 1 Then
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open "SELECT COMPCODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDACODE='" & TxtSaudaCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Sauda code already exists.", vbExclamation, "Warning"
            Cancel = True
        End If
        TxtSaudaName.text = TxtSaudaCode.text
        Set TRec = Nothing
    End If
End Sub
Sub List_Sauda()
    Dim TRec As ADODB.Recordset
    Screen.MousePointer = 11
    Call Get_Selection(12)
    mysql = "SELECT SAUDACODE,SAUDANAME,ITEMCODE,MATURITY,EXCODE,INSTTYPE,OPTTYPE,STRIKEPRICE,TRADEABLELOT,BROKLOT,REFLOT,NOCODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE,ITEMCODE,MATURITY,INSTTYPE,OPTTYPE,STRIKEPRICE"
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "SaudaList.RPT", 1)
    RDCREPO.DiscardSavedData
    RDCREPO.Database.SetDataSource TRec
    CRViewer1.Width = CInt(GETMAIN.Width - 100)
    CRViewer1.Height = CInt(GETMAIN.Height - GETMAIN.Toolbar1.Height - 1000)
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Visible = True
    CRViewer1.ZOrder
    CRViewer1.ReportSource = RDCREPO
    CRViewer1.ViewReport
    Screen.MousePointer = 0
End Sub

Private Sub TxtStrikePrice_Validate(Cancel As Boolean)

If LenB(TxtStrikePrice.text) = 0 Then
    MsgBox "Strike Price Can not be blank"
    Cancel = True
Else
    If Option4.Value = True Then
        If Val(TxtStrikePrice.text) <= 0 Then
            MsgBox "Strike Price Can not less than Zero"
            Cancel = True
        End If
    Else
        TxtStrikePrice.text = "0.00"
    End If
    TxtStrikePrice.text = Format(TxtStrikePrice.text, "0.00")
End If
If (Option3.Value = True Or Option7.Value = True) Then TxtStrikePrice.text = "0.00"
   
End Sub
Private Sub TxtRefLot_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtRefLot_Validate(Cancel As Boolean)
    TxtRefLot.text = Format(TxtRefLot.text, "0.00")
End Sub
Private Sub TxtTradeableLot_Validate(Cancel As Boolean)
    TxtTradeableLot.text = Format(TxtTradeableLot.text, "0.00")
End Sub
Private Sub TxtBrokLot_Validate(Cancel As Boolean)
    TxtBrokLot.text = Format(TxtBrokLot.text, "0.00")
End Sub
Private Sub TxtTradeableLot_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtStrikePrice_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtNoCode_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub

Private Sub TxtBrokLot_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtTradeableLot_GotFocus()
    TxtTradeableLot.SelLength = Len(TxtTradeableLot.text)
End Sub
Private Sub TxtBrokLot_GotFocus()
    TxtBrokLot.SelLength = Len(TxtBrokLot.text)
End Sub
Private Sub TxtStrikePrice_GotFocus()
    TxtStrikePrice.SelLength = Len(TxtStrikePrice.text)
End Sub
Private Sub TxtFilterCode_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub FillDataGrid()
    Set SaudaRec = Nothing
    Set SaudaRec = New ADODB.Recordset
    mysql = "SELECT SaudaId,ExCode,SaudaCode,ItemCode,Maturity,InstType,TradeableLot AS Lot,RefLot,BROKLOT as BrokLot,OptType,StrikePrice,SaudaName,NOCODE  ,EXID,ITEMID"
    mysql = mysql & " FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " "
    If LenB(TxtFilterCode.text) <> 0 Then mysql = mysql & " AND  UPPER(SAUDANAME) LIKE '" & Trim(UCase(TxtFilterCode.text)) & "%' "
    If LenB(FItemCombo.BoundText) > 0 Then mysql = mysql & " AND  ITEMCODE ='" & FItemCombo.BoundText & "' "
    If LenB(FEXCombo.BoundText) > 0 Then mysql = mysql & " AND EXCODE ='" & FEXCombo.BoundText & "' "
    If CmdFilterOk.Caption = "Clear" Then mysql = mysql & " AND MATURITY  >='" & Format(vcDTP2.Value, "YYYY/MM/DD") & "' AND MATURITY  <='" & Format(vcDTP3.Value, "YYYY/MM/DD") & "'"
    mysql = mysql & " ORDER BY ITEMCODE,MATURITY"
    SaudaRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not SaudaRec.EOF Then
        Set DataGrid1.DataSource = SaudaRec
        DataGrid1.ReBind
        DataGrid1.Refresh
        DataGrid1.Columns(0).Width = 700:
        DataGrid1.Columns(1).Width = 700:
        DataGrid1.Columns(2).Width = 3000
        DataGrid1.Columns(3).Width = 2000:
        DataGrid1.Columns(4).Width = 1100
        DataGrid1.Columns(5).Width = 1000:
        DataGrid1.Columns(6).Width = 1000
        DataGrid1.Columns(7).Width = 1000:
        DataGrid1.Columns(8).Width = 1000
        DataGrid1.Columns(9).Width = 1000:
        DataGrid1.Columns(10).Width = 1000
        DataGrid1.Columns(11).Width = 3000:
        
        DataGrid1.Columns(6).Alignment = dbgRight
        
        DataGrid1.Columns(7).Alignment = dbgRight:  DataGrid1.Columns(8).Alignment = dbgRight
        DataGrid1.Columns(10).Alignment = dbgRight:  DataGrid1.Columns(6).NumberFormat = "0.00"
        DataGrid1.Columns(7).NumberFormat = "0.00": DataGrid1.Columns(8).NumberFormat = "0.00"
        DataGrid1.Columns(10).NumberFormat = "0.00"
        DataGrid1.Refresh
    Else
        MsgBox "No Records Found"
       ' TxtFilterCode.text = vbNullString
        'TxtFilterCode.SetFocus
    End If
End Sub
