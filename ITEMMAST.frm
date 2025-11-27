VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ITEMMAST 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   11430
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtExID 
      Height          =   495
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404080&
      Height          =   855
      Left            =   150
      TabIndex        =   12
      Top             =   0
      Width           =   13415
      Begin VB.Label Label27 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Item Master"
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
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   13455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF8080&
      Height          =   4335
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   13335
      Begin MSDataListLib.DataCombo ExCombo 
         Height          =   360
         Left            =   6360
         TabIndex        =   2
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.CommandButton Command3 
         Caption         =   "Go"
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
         Left            =   11280
         TabIndex        =   3
         Top             =   120
         Width           =   1290
      End
      Begin VB.TextBox TxtFilterCode 
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
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   3855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3615
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange"
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
         Left            =   5280
         TabIndex        =   14
         Top             =   173
         Width           =   1095
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Item"
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
         TabIndex        =   10
         Top             =   173
         Width           =   1095
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2295
      Left            =   15600
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
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
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13455
      Begin TabDlg.SSTab SSTab1 
         Height          =   3495
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   6165
         _Version        =   393216
         Tab             =   2
         TabHeight       =   706
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "ITEMMAST.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame6"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Taxes"
         TabPicture(1)   =   "ITEMMAST.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vcDTP1"
         Tab(1).Control(1)=   "Frame11"
         Tab(1).Control(2)=   "Frame9"
         Tab(1).Control(3)=   "Frame12"
         Tab(1).Control(4)=   "ITTaxGrid"
         Tab(1).Control(5)=   "Label10"
         Tab(1).Control(6)=   "Label28"
         Tab(1).Control(7)=   "Label30"
         Tab(1).Control(8)=   "Label31"
         Tab(1).Control(9)=   "Shape1"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Alias"
         TabPicture(2)   =   "ITEMMAST.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame4"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   -75000
            TabIndex        =   43
            Top             =   480
            Width           =   13455
            Begin VB.TextBox TXtCloseID 
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
               Height          =   405
               Left            =   11040
               TabIndex        =   52
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox TxtItemID 
               Alignment       =   1  'Right Justify
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
               Height          =   405
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox TxtPriceUnit 
               Alignment       =   1  'Right Justify
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
               Height          =   405
               Left            =   1680
               TabIndex        =   50
               Top             =   1380
               Width           =   2535
            End
            Begin VB.TextBox TxtQtyUnit 
               Alignment       =   1  'Right Justify
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
               Height          =   405
               Left            =   1680
               TabIndex        =   49
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox TxtGroup 
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
               Height          =   405
               Left            =   5760
               TabIndex        =   48
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox TxtItemCode 
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
               Height          =   405
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   47
               Top             =   240
               Width           =   3615
            End
            Begin VB.TextBox TxtItemName 
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
               Height          =   405
               Left            =   8400
               MaxLength       =   20
               TabIndex        =   46
               Top             =   240
               Width           =   3975
            End
            Begin VB.TextBox TxtLot 
               Alignment       =   1  'Right Justify
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
               Height          =   405
               Left            =   8400
               TabIndex        =   45
               Top             =   810
               Width           =   1215
            End
            Begin VB.TextBox TxtExSymbol 
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
               Height          =   405
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   44
               Top             =   810
               Width           =   2535
            End
            Begin MSDataListLib.DataCombo ExNameDb 
               Height          =   360
               Left            =   5760
               TabIndex        =   53
               Top             =   810
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ForeColor       =   16711680
               Text            =   "ExNamedb"
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
            Begin VB.Label Label9 
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
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   9720
               TabIndex        =   63
               Top             =   885
               Width           =   975
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   5400
               TabIndex        =   62
               Top             =   315
               Width           =   210
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Qty Unit"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   1995
               Width           =   855
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Price Unit"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   1455
               Width           =   1095
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   4320
               TabIndex        =   59
               Top             =   1440
               Width           =   570
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Item Code"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   58
               Top             =   315
               Width           =   1020
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Item Name"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   7200
               TabIndex        =   57
               Top             =   315
               Width           =   1065
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lot Size"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   7200
               TabIndex        =   56
               Top             =   885
               Width           =   795
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ex-Symbol"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   55
               Top             =   870
               Width           =   1230
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label13 
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
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   4320
               TabIndex        =   54
               Top             =   885
               Width           =   720
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   3255
            Left            =   0
            TabIndex        =   32
            Top             =   435
            Width           =   13455
            Begin VB.TextBox txtitalias3 
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
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   35
               Top             =   1260
               Width           =   2535
            End
            Begin VB.TextBox txtitalias4 
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
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   37
               Top             =   1820
               Width           =   2535
            End
            Begin VB.TextBox txtitalias1 
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
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   33
               Top             =   120
               Width           =   2535
            End
            Begin VB.TextBox txtitalias5 
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
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   39
               Top             =   2400
               Width           =   2535
            End
            Begin VB.TextBox txtitalias2 
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
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   34
               Top             =   690
               Width           =   2535
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Item Alias 4"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   1890
               Width           =   1215
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Item Alias 3"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   1335
               Width           =   1215
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Item Alias 1"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   40
               Top             =   180
               Width           =   1170
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Item Alias 5"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   38
               Top             =   2460
               Width           =   1170
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Item Alias 2"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   36
               Top             =   750
               Width           =   1230
               WordWrap        =   -1  'True
            End
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   375
            Left            =   -63720
            TabIndex        =   31
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
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
            Value           =   44126.5582407407
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   -70200
            TabIndex        =   23
            Top             =   600
            Width           =   1575
            Begin VB.OptionButton OptRiskMYes 
               BackColor       =   &H00FFC0C0&
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
               Height          =   285
               Left            =   0
               TabIndex        =   25
               Top             =   50
               Width           =   855
            End
            Begin VB.OptionButton OptRiskMNo 
               BackColor       =   &H00FFC0C0&
               Caption         =   "No"
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
               Left            =   840
               TabIndex        =   24
               Top             =   50
               Width           =   735
            End
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   -73800
            TabIndex        =   20
            Top             =   600
            Width           =   1695
            Begin VB.OptionButton OptCttYes 
               BackColor       =   &H00FFC0C0&
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
               Height          =   285
               Left            =   0
               TabIndex        =   22
               Top             =   50
               Width           =   855
            End
            Begin VB.OptionButton OptCttNo 
               BackColor       =   &H00FFC0C0&
               Caption         =   "No"
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
               Left            =   840
               TabIndex        =   21
               Top             =   50
               Width           =   855
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   -67080
            TabIndex        =   17
            Top             =   600
            Width           =   1575
            Begin VB.OptionButton OptSDutyNo 
               BackColor       =   &H00FFC0C0&
               Caption         =   "No"
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
               Left            =   840
               TabIndex        =   19
               Top             =   50
               Width           =   735
            End
            Begin VB.OptionButton OptSDutyYes 
               BackColor       =   &H00FFC0C0&
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
               Height          =   285
               Left            =   0
               TabIndex        =   18
               Top             =   50
               Width           =   855
            End
         End
         Begin MSDataGridLib.DataGrid ITTaxGrid 
            Height          =   2175
            Left            =   -74880
            TabIndex        =   26
            Top             =   1200
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            AllowAddNew     =   -1  'True
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
            Caption         =   "Item Wise Tax"
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
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "CTT Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   -64800
            TabIndex        =   30
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "CTT Appl"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   -74880
            TabIndex        =   29
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "RiskMngt.Appl"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   -71640
            TabIndex        =   28
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "StmpDuty Appl."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   -68160
            TabIndex        =   27
            Top             =   660
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFC0C0&
            Height          =   3135
            Left            =   -75000
            Top             =   450
            Width           =   13455
         End
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000A&
         Caption         =   "Per"
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
         Left            =   7080
         TabIndex        =   7
         Top             =   -360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000A&
         Caption         =   "S Duty Rate"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   -360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000A&
         Caption         =   "Stamp Duty Type"
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
         Left            =   0
         TabIndex        =   5
         Top             =   -360
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "I t e m   L i s t"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   16200
      TabIndex        =   11
      Top             =   960
      Width           =   1395
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   8460
      Left            =   0
      Top             =   480
      Width           =   13725
   End
End
Attribute VB_Name = "ITEMMAST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fb_Press As Byte::       Dim ItemRec As ADODB.Recordset:     Dim ExMastRec As ADODB.Recordset
Dim LPriceUnit As String:       Dim LQtyUnit As String:             Dim LRiskMApp As String
Dim LRegularLot As Double:      Dim LStmApp As String
Dim LOldLot As Double:          Dim LCTTApp As String:              Dim RecITTax As ADODB.Recordset
Dim LCLoseID  As Integer
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" _
    (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Long
    Private Declare Function InternetSetOption Lib "wininet.dll" _
    Alias "InternetSetOptionA" (ByVal hInternet As Long, _
    ByVal lOption As Long, lpBuffer As Any, _
    ByVal lBufferLength As Long) As Long
Function GetLastWinInetError() As String
    Dim dwError As Long
    Dim dwBufferLen As Long
    Dim szBuffer As String
    
    dwBufferLen = 1024
    szBuffer = Space$(dwBufferLen)
    
    If InternetGetLastResponseInfo(dwError, szBuffer, dwBufferLen) Then
        GetLastWinInetError = "Error " & dwError & ": " & Left$(szBuffer, dwBufferLen)
        Debug.Print GetLastWinInetError
    Else
        GetLastWinInetError = "No extended information available."
    End If
End Function


Sub Add_Rec()
    OptCttNo.Value = True
    Fb_Press = 1:
    Call Get_Selection(1)
    TxtItemCode.text = vbNullString:        TxtItemName.text = vbNullString
    txtitalias1.text = vbNullString: txtitalias2.text = vbNullString: txtitalias3.text = vbNullString: txtitalias4.text = vbNullString: txtitalias5.text = vbNullString
    Frame1.Enabled = True:
    Frame2.Enabled = False
    'DataList1.Locked = True
    TxtItemCode.SetFocus
End Sub
Sub Save_Rec()
On Error GoTo err1
Dim TRec As ADODB.Recordset:    Dim LSaudas As String:  Dim LIExCode  As String
Dim LExID  As Integer
Dim LItemID As Integer
On Error GoTo err1
    If LenB(TxtItemCode.text) = 0 Then
        MsgBox "Item Code required before saving record.", vbCritical, "Error"
        If TxtItemCode.Enabled Then TxtItemCode.SetFocus
        Exit Sub
    End If
    If LenB(ExNameDb.text) < 1 Then
        MsgBox "Please Select Exchange"
        ExNameDb.SetFocus
        Exit Sub
    Else
        ExMastRec.MoveFirst
        ExMastRec.Find "EXCODE ='" & ExNameDb.BoundText & "'"
        If ExMastRec.EOF Then
            ExNameDb.BoundText = vbNullString
            MsgBox "Please Select Exchange"
            ExNameDb.SetFocus
            Exit Sub
        Else
            LIExCode = ExMastRec!excode
            TxtExID.text = ExMastRec!EXID
        End If
    End If
    CNNERR = True
    Cnn.BeginTrans
    LExID = Get_ExID(LIExCode)
    If Fb_Press = 1 Then
        mysql = "INSERT INTO ITEMMAST (COMPCODE,ITEMCODE,ITEMNAME,LOT,SCGROUP,EXHCODE,EXCHANGECODE,PRICEUNIT,QTYUNIT,REGULARLOT,CTTAPP,RISKMAPP,STMAPP,EXID,CTTDATE,ITALIAS_1,ITALIAS_2,ITALIAS_3,ITALIAS_4,ITALIAS_5) "
        mysql = mysql & " VALUES (" & GCompCode & ",'" & TxtItemCode.text & "','" & TxtItemName.text & "'," & Val(TxtLot.text) & ",'" & TxtGroup.text & "','" & TxtExSymbol.text & "'"
        mysql = mysql & ",'" & LIExCode & "','" & Trim(TxtPriceUnit.text) & "','" & TxtQtyUnit.text & "'," & LRegularLot & ""
        mysql = mysql & ",'" & IIf(OptCttYes.Value, "Y", "N") & "','" & IIf(OptRiskMYes.Value, "Y", "N") & "','" & IIf(OptSDutyYes.Value, "Y", "N") & "'," & LExID & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "','" & txtitalias1.text & "','" & txtitalias2.text & "','" & txtitalias3.text & "','" & txtitalias4.text & "','" & txtitalias5.text & "')"
        Cnn.Execute mysql
    Else
        mysql = "UPDATE ITEMMAST SET "
        mysql = mysql & "  ITEMName = '" & TxtItemName.text & "'"
        mysql = mysql & " ,LOT = " & Val(TxtLot.text) & ""
        mysql = mysql & " ,SCGROUP ='" & Trim(TxtGroup.text) & "'"
        mysql = mysql & " ,EXHCODE ='" & Trim(TxtExSymbol.text) & "'"
        mysql = mysql & " ,EXCHANGECODE = '" & LIExCode & "'"
        mysql = mysql & " ,EXID = " & LExID & ""
        mysql = mysql & " ,CITEMID  = " & Val(TXtCloseID.text) & ""
        mysql = mysql & " ,PRICEUNIT = '" & Trim(TxtPriceUnit.text) & "'"
        mysql = mysql & " ,QTYUNIT = '" & Trim(TxtQtyUnit.text) & "'"
        mysql = mysql & " ,REGULARLOT =" & LRegularLot & ""
        mysql = mysql & " ,CTTAPP = '" & IIf(OptCttYes.Value, "Y", "N") & "'"
        mysql = mysql & " ,RISKMAPP = '" & IIf(OptRiskMYes.Value, "Y", "N") & "'"
        mysql = mysql & " ,STMAPP = '" & IIf(OptSDutyYes.Value, "Y", "N") & "'"
        mysql = mysql & " ,CTTDATE  = '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' "
        mysql = mysql & " ,ITALIAS_1 = '" & txtitalias1.text & "' "
        mysql = mysql & " ,ITALIAS_2 = '" & txtitalias2.text & "' "
        mysql = mysql & " ,ITALIAS_3 = '" & txtitalias3.text & "' "
        mysql = mysql & " ,ITALIAS_4 = '" & txtitalias4.text & "' "
        mysql = mysql & " ,ITALIAS_5 = '" & txtitalias5.text & "' "
        
        mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND ITEMCODE='" & TxtItemCode.text & "'"
        Cnn.Execute mysql
    End If
    LItemID = Get_ITEMID(TxtItemCode.text)
    mysql = "DELETE FROM ITEM_TAX WHERE COMPCODE =" & GCompCode & " AND ITEMID =" & LItemID & ""
    Cnn.Execute mysql
    
    If RecITTax.RecordCount > 0 Then
        RecITTax.MoveFirst
        Do While Not RecITTax.EOF
            If Not IsNull(RecITTax!StartDate) And Not IsNull(RecITTax!ENDDATE) Then
                If IsDate(RecITTax!StartDate) And IsDate(RecITTax!ENDDATE) Then
                    mysql = "INSERT INTO ITEM_TAX (COMPCODE,EXCODE,ITEMCODE,FROMDT,TODT,RISKMRATE,SEBITAX,EXID,ITEMID)"
                    mysql = mysql & " VALUES (" & GCompCode & " ,'" & LIExCode & "','" & TxtItemCode.text & "','" & Format(RecITTax!StartDate, "YYYY/MM/DD") & "'"
                    mysql = mysql & " ,'" & Format(RecITTax!ENDDATE, "YYYY/MM/DD") & "'," & Val(RecITTax!RISKMFEES & vbNullString) & "," & Val(RecITTax!SEBITAX & vbNullString) & "," & LExID & "," & LItemID & ")"
                    Cnn.Execute mysql
                End If
            End If
            RecITTax.MoveNext
        Loop
    End If
    Cnn.CommitTrans
    CNNERR = False
    If Fb_Press = 2 Then
        Cnn.BeginTrans
        CNNERR = True
        mysql = "UPDATE SAUDAMAST SET EXCODE='" & LIExCode & "',EXID=" & LExID & " ,ITEMID =" & LItemID & " WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND ITEMCODE='" & Trim(TxtItemCode.text) & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_D SET EXID = " & LExID & " ,  EXCODE='" & LIExCode & "', ITEMID =" & LItemID & " "
        mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND ITEMCODE='" & Trim(TxtItemCode.text) & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_M SET EXID = " & LExID & " ,  EXCODE='" & LIExCode & "', ITEMID =" & LItemID & " "
        mysql = mysql & " WHERE COMPCODE =" & GCompCode & " AND ITEMCODE='" & Trim(TxtItemCode.text) & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE CTR_D SET CALVAL=" & Val(TxtLot.text) & " WHERE COMPCODE =" & GCompCode & " AND EXCODE <>'NSE' AND ITEMCODE='" & Trim(TxtItemCode.text) & "'"
        Cnn.Execute mysql
        If LIExCode = "NSE" Then
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            mysql = "SELECT LOTWISE FROM EXMAST WHERE EXID =" & LExID & ""
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Not TRec.EOF Then
                If TRec!LOTWISE = "N" Then
                    mysql = "UPDATE CTR_D SET CALVAL=" & Val(TxtLot.text) & " WHERE EXID =" & LExID & " AND ITEMID =" & LItemID & ""
                    Cnn.Execute mysql
                End If
            End If
        End If
        mysql = "UPDATE CTR_R SET EXCODE='" & LIExCode & "' ,EXID=" & LExID & "  WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND ITEMCODE='" & Trim(TxtItemCode.text) & "'"
        Cnn.Execute mysql
        
        mysql = "UPDATE INV_D SET EXCODE='" & LIExCode & "'  ,EXID=" & LExID & " WHERE COMPCODE =" & GCompCode & " "
        mysql = mysql & " AND ITEMCODE='" & Trim(TxtItemCode.text) & "'"
        Cnn.Execute mysql
                
        Cnn.CommitTrans
        
        
        CNNERR = False
        LSaudas = vbNullString
        
        
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        mysql = "SELECT DISTINCT SAUDAID FROM CTR_D WHERE COMPCODE = " & GCompCode & " "
        mysql = mysql & " AND ITEMID = " & LItemID & " "
        TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        While Not TRec.EOF
            If LenB(LSaudas) > 0 Then LSaudas = LSaudas & ", "
            LSaudas = LSaudas & "" & Trim(str(TRec!SAUDAID)) & ""
            TRec.MoveNext
        Wend
        Dim LSParties As String
        If TxtLot.text <> LOldLot Then
        'If LenB(LSaudas) > 0 Then
            LSParties = vbNullString
            Set TRec = Nothing: Set TRec = New ADODB.Recordset
            mysql = "SELECT DISTINCT PARTY FROM CTR_D  WHERE exid IN  ( " & LExID & " )"
            TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            While Not TRec.EOF
                If LenB(LSParties) > 0 Then LSParties = LSParties & ", "
                LSParties = LSParties & "'" & TRec!PARTY & "'"
                TRec.MoveNext
            Wend
            If LenB(LSParties) > 0 Then
                Cnn.BeginTrans
                CNNERR = True
                Call Update_Charges(LSParties, Trim(str(LExID)), LSaudas, Trim(str(LItemID)), GFinBegin, GFinEnd, False)
                'Call Delete_Inv_D(vbNullString, Trim(Str(LExID)), LSaudas, GFinBegin)
                
                'Cnn.CommitTrans
                '''
                If BILL_GENERATION(CDate(GFinBegin), CDate(GFinEnd), LSaudas, LSParties, Trim(str(LExID))) Then
                    Cnn.CommitTrans
                    CNNERR = False
                    MsgBox "Bills updated Succesfully"
                Else
                    Cnn.RollbackTrans: CNNERR = False
                End If
                'Call Chk_Billing
            End If
        End If
    End If
    Call CANCEL_REC
    Exit Sub
err1:
    If err.Number <> 0 Then
        If CNNERR = True Then
            Cnn.RollbackTrans
            CNNERR = False
        End If
        MsgBox err.Description
        
    End If
    
End Sub
Sub CANCEL_REC()
    SSTab1.Tab = 0
    LPriceUnit = vbNullString:          LQtyUnit = vbNullString:
    LRegularLot = 0:                    TxtItemCode.Enabled = True
    LCTTApp = "N":                      LRiskMApp = "N":            LStmApp = "N"
    Fb_Press = 0:                       TxtItemCode.text = vbNullString
    TxtItemName.text = vbNullString:    TxtExSymbol.text = vbNullString
    TxtGroup.text = vbNullString:       TxtPriceUnit.text = vbNullString
    TxtQtyUnit.text = vbNullString:     TxtLot.text = "1.00"
    TxtItemID.text = vbNullString
    txtitalias1.text = vbNullString: txtitalias2.text = vbNullString: txtitalias3.text = vbNullString: txtitalias4.text = vbNullString: txtitalias5.text = vbNullString
    TXtCloseID.text = 0
    TxtExID = vbNullString
    FillDataGrid
    Call Get_Selection(10)
    Frame1.Enabled = False
    Call Set_TaxRec
    Set ITTaxGrid.DataSource = RecITTax
    ITTaxGrid.ReBind: ITTaxGrid.Refresh
    ITTaxGrid.Columns(0).Width = 1200
    ITTaxGrid.Columns(1).Width = 1200
    ITTaxGrid.Columns(2).Width = 1200
    ITTaxGrid.Columns(3).Width = 1200
    ITTaxGrid.Columns(2).Alignment = dbgRight
    ITTaxGrid.Columns(2).NumberFormat = "0.000000"
    ITTaxGrid.Columns(3).Alignment = dbgRight
    ITTaxGrid.Columns(3).NumberFormat = "0.000000"
    vcDTP1.Value = GFinBegin
    'DataList1.Locked = False
    Frame2.Enabled = True
End Sub
Sub MODIFY_REC()
    Dim TRec As ADODB.Recordset
    Frame2.Enabled = False
    If LenB(TxtItemCode.text) <> 0 Then
        'DataList1.Locked = True
        ItemRec.MoveFirst
        ItemRec.Find "ITEMCODE='" & TxtItemCode.text & "'", , adSearchForward
        If Not ItemRec.EOF Then
            TxtItemCode.Enabled = False
            TxtItemID.text = ItemRec!itemid
            TxtExID.text = ItemRec!EXID
            TxtItemName.text = ItemRec!ITEMName:
            LPriceUnit = ItemRec!PRICEUNIT & vbNullString:            LQtyUnit = (ItemRec!QTYUNIT & vbNullString)
            LRegularLot = Val(ItemRec!regularlot & vbNullString)
            TxtLot.text = IIf(IsNull(ItemRec!lot), "0.0000", Format(ItemRec!lot, "0.0000"))
            LOldLot = IIf(IsNull(ItemRec!lot), "0.0000", Format(ItemRec!lot, "0.0000"))
            TxtGroup.text = IIf(IsNull(ItemRec!SCGROUP), vbNullString, ItemRec!SCGROUP)
            TxtExSymbol.text = IIf(IsNull(ItemRec!EXHCODE), vbNullString, ItemRec!EXHCODE)
            LCTTApp = IIf(IsNull(ItemRec!CTTAPP), "N", ItemRec!CTTAPP)
            LRiskMApp = IIf(IsNull(ItemRec!RISKMAPP), "N", ItemRec!RISKMAPP)
            LStmApp = IIf(IsNull(ItemRec!STMAPP), "N", ItemRec!STMAPP)
            TXtCloseID.text = ItemRec!CITEMID
            vcDTP1.Value = ItemRec!CTTDATE
            If LCTTApp = "Y" Then
                OptCttYes.Value = True
            Else
                OptCttNo.Value = True
            End If
            If LRiskMApp = "Y" Then
                OptRiskMYes.Value = True
            Else
                OptRiskMNo.Value = True
            End If
            If LStmApp = "Y" Then
                OptSDutyYes.Value = True
            Else
                OptSDutyNo.Value = True
            End If
            ExNameDb.BoundText = ItemRec!excode
            TxtPriceUnit.text = ItemRec!PRICEUNIT & vbNullString
            TxtQtyUnit.text = ItemRec!QTYUNIT & vbNullString
            Call Set_TaxRec
            mysql = "SELECT * FROM ITEM_TAX WHERE COMPCODE =" & GCompCode & " AND ITEMCODE  ='" & TxtItemCode.text & "'"
            Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
            If Not TRec.EOF Then
                While Not TRec.EOF
                    RecITTax.AddNew:                    RecITTax!StartDate = TRec!FROMDT
                    RecITTax!ENDDATE = TRec!ToDt:       RecITTax!RISKMFEES = Val(TRec!RISKMRATE & vbNullString)
                    RecITTax!SEBITAX = Val(TRec!SEBITAX & vbNullString)

                    RecITTax.Update:                    TRec.MoveNext
                Wend
            End If
            Set ITTaxGrid.DataSource = RecITTax:            ITTaxGrid.ReBind: ITTaxGrid.Refresh
            ITTaxGrid.Columns(0).Width = 1200:              ITTaxGrid.Columns(1).Width = 1200
            ITTaxGrid.Columns(2).Width = 1200:              ITTaxGrid.Columns(2).Alignment = dbgRight
            ITTaxGrid.Columns(2).NumberFormat = "0.000000":   ITTaxGrid.Columns(3).Width = 1200
            ITTaxGrid.Columns(3).Alignment = dbgRight:      ITTaxGrid.Columns(3).NumberFormat = "0.000000"
                        
            txtitalias1.text = ItemRec!ITALIAS_1 & ""
            txtitalias2.text = ItemRec!ITALIAS_2 & ""
            txtitalias3.text = ItemRec!ITALIAS_3 & ""
            txtitalias4.text = ItemRec!ITALIAS_4 & ""
            txtitalias5.text = ItemRec!ITALIAS_5 & ""
            
            Frame1.Enabled = True:                          TxtItemName.SetFocus
        End If
        If Fb_Press = 3 Then
            If MsgBox("You are about to Delete one Record. Confirm Delete ?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                Set TRec = Nothing:                Set TRec = New ADODB.Recordset
                mysql = "SELECT SD.ITEMCODE FROM CTR_D AS CD,SAUDAMAST AS SD WHERE CD.COMPCODE=" & GCompCode & " AND CD.COMPCODE=SD.COMPCODE AND SD.SAUDACODE=CD.SAUDA AND SD.ITEMCODE='" & TxtItemCode.text & "'"
                TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                If Not TRec.EOF Then
                    MsgBox "Transaction Exists can not Delete Item.", vbExclamation, "Error"
                    Call CANCEL_REC
                    Exit Sub
                Else
                    Cnn.Execute "DELETE FROM PITBROK WHERE COMPCODE=" & GCompCode & " AND ITEMCODE='" & TxtItemCode.text & "'"
                    Cnn.Execute "DELETE FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND ITEMCODE='" & TxtItemCode.text & "'"
                End If
                mysql = "DELETE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & "  AND ITEMCODE ='" & TxtItemCode.text & "'"
                Cnn.Execute mysql
                Call CANCEL_REC
            End If
            
            'DataList1.Locked = False
            Frame2.Enabled = True
            DataGrid1.SetFocus
            'DataList1.SetFocus
        End If
    Else
        MsgBox "Please Select Item.", vbInformation
        Call CANCEL_REC
        Frame2.Enabled = True
        DataGrid1.SetFocus
        'DataList1.Locked = False
        'DataList1.SetFocus
    End If
End Sub
Private Sub DataCombo1_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub



Public Sub Connect_Ftp(LHostName As String, LUserName As String, LPass As String, LRemoteFileName As String, LLocalFileName As String)
    Dim hopen As Long
    Dim dwType As Long
    Dim hConnection As Long
    Dim extendedInfo As String
    Dim bPassiveMode As Boolean
    
    FlagDownloadedFromOurFTP = False
    
    ' Constants
    Const INTERNET_FLAG_RELOAD = &H80000000
    Const INTERNET_OPEN_TYPE_PRECONFIG = 0
    Const INTERNET_INVALID_PORT_NUMBER = 0
    Const INTERNET_SERVICE_FTP = 1
    Const FTP_TRANSFER_TYPE_BINARY = 2  ' Changed from ASCII for ZIP files
    Const FILE_ATTRIBUTE_NORMAL = &H80
    Const INTERNET_OPTION_PASSIVE = 12  ' For enabling passive mode
    
    ' Open internet connection
    hopen = InternetOpen("My VB FTP Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hopen = 0 Then
        ErrorOut err.LastDllError, "InternetOpen"
        Exit Sub
    End If
    
    ' Set transfer type to BINARY (critical for ZIP files)
    dwType = FTP_TRANSFER_TYPE_BINARY
    
    ' Enable passive mode (helps with firewall/NAT issues)
    bPassiveMode = True
    InternetSetOption hopen, INTERNET_OPTION_PASSIVE, bPassiveMode, Len(bPassiveMode)
    
    ' Connect to FTP server
    hConnection = InternetConnect(hopen, LHostName, INTERNET_INVALID_PORT_NUMBER, _
                                LUserName, LPass, INTERNET_SERVICE_FTP, 0, 0)
    If hConnection = 0 Then
        ErrorOut err.LastDllError, "InternetConnect"
        InternetCloseHandle hopen
        Exit Sub
    End If
    
    ' Create directory if it doesn't exist
    CreateDirectoryIfNotExists LLocalFileName
    
    ' Download the file with retry logic
    Dim iRetry As Integer
    For iRetry = 1 To 3  ' Try up to 3 times
        If FtpGetFile(hConnection, LRemoteFileName, LLocalFileName, False, _
                     FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) Then
            FlagDownloadedFromOurFTP = True
            Exit For
        Else
            Dim ErrorCode As Long
            ErrorCode = err.LastDllError
            
            ' Log error information
            Debug.Print "Attempt " & iRetry & " failed: " & ErrorCode
            Debug.Print "Error: " & GetLastWinInetError()
            
            If iRetry < 3 Then
                Sleep 3000  ' Wait 3 seconds before retry
            Else
                ' Final failure
                extendedInfo = GetLastWinInetError()
                MsgBox "FTP Transfer Failed:" & vbCrLf & _
                       "Error: " & ErrorCode & vbCrLf & _
                       "Details: " & extendedInfo & vbCrLf & _
                       "Remote: " & LRemoteFileName & vbCrLf & _
                       "Local: " & LLocalFileName, vbExclamation, "FTP Error"
            End If
        End If
    Next iRetry
    
    ' Clean up
    InternetCloseHandle hConnection
    InternetCloseHandle hopen
End Sub

Private Sub CreateDirectoryIfNotExists(filePath As String)
    Dim directory As String
    directory = Left(filePath, InStrRev(filePath, "\"))
    
    If Dir(directory, vbDirectory) = "" Then
        MkDir directory
    End If
End Sub
Public Sub Connect_Ftpp(LHostName As String, LUserName As String, LPass As String, LRemoteFileName As String, LLocalFileName As String)
    Dim ftp As New ChilkatFtp2
    Dim iRetry As Integer
    Dim success As Long
    Dim extendedInfo As String
    
    FlagDownloadedFromOurFTP = False

    ' Unlock the Chilkat component
    success = ftp.UnlockComponent("YOUR_UNLOCK_CODE")
    If success <> 1 Then
        MsgBox "Chilkat unlock failed: " & ftp.LastErrorText, vbCritical
        Exit Sub
    End If

    ' FTP connection settings
    ftp.HostName = LHostName
    ftp.USERNAME = LUserName
    ftp.PASSWORD = LPass

    ftp.Passive = 1         ' Enable passive mode
    ftp.UseEpsv = 1         ' Use EPSV instead of PASV
    ftp.AuthTls = 0         ' Set to 1 if server requires TLS

    ' Ensure local folder exists
    CreateDirectoryIfNotExists LLocalFileName

    ' Retry download logic
    For iRetry = 1 To 3
        success = ftp.Connect()
        If success = 1 Then
            success = ftp.GetFile(LRemoteFileName, LLocalFileName)
            If success = 1 Then
                FlagDownloadedFromOurFTP = True
                ftp.Disconnect
                Exit For
            Else
                Debug.Print "Attempt " & iRetry & " failed: " & ftp.LastErrorText
            End If
            ftp.Disconnect
        Else
            Debug.Print "Connection failed on attempt " & iRetry & ": " & ftp.LastErrorText
        End If
        
        If iRetry < 3 Then
            Sleep 3000  ' Wait before retrying
        Else
            MsgBox "FTP Transfer Failed after 3 attempts." & vbCrLf & _
                   "Details: " & ftp.LastErrorText & vbCrLf & _
                   "Remote: " & LRemoteFileName & vbCrLf & _
                   "Local: " & LLocalFileName, vbExclamation, "FTP Error"
        End If
    Next iRetry
End Sub

Public Sub Connect_Ftp_WinSCP( _
    ByVal LHostName As String, _
    ByVal LUserName As String, _
    ByVal LPass As String, _
    ByVal LRemoteFileName As String, _
    ByVal LLocalFileName As String)

    Dim sessionOptions As New WinSCPnet.sessionOptions
    Dim session As New WinSCPnet.session
    Dim transferResult As WinSCPnet.TransferOperationResult

    On Error GoTo ErrorHandler

    ' Setup session options
    With sessionOptions
        .Protocol = WinSCPnet.Protocol_Ftp
        .HostName = LHostName
        .USERNAME = LUserName
        .PASSWORD = LPass
        .FtpMode = WinSCPnet.FtpMode_Passive
        .FtpSecure = WinSCPnet.FtpSecure_Explicit
        .GiveUpSecurityAndAcceptAnyTlsHostCertificate = True
    End With

    ' Make sure the directory exists
    CreateDirectoryIfNotExists LLocalFileName

    ' Open the session
    session.Open sessionOptions

    ' Perform the file download
    Set transferResult = session.GetFiles("/Microsoft.NET" & LRemoteFileName, LLocalFileName, False, Nothing)

    If transferResult.IsSuccess Then
        MsgBox "Download succeeded.", vbInformation
        FlagDownloadedFromOurFTP = True
    Else
        MsgBox "Download failed: " & transferResult.Failures.Item(1).message, vbCritical
    End If

    ' Clean up
    session.Dispose
    Exit Sub

ErrorHandler:
    MsgBox "WinSCP Error: " & err.Description, vbCritical
    If Not session Is Nothing Then session.Dispose
End Sub

Public Sub SendWhatsAppPDF( _
    ByVal phoneNumberId As String, _
    ByVal recipientNumber As String, _
    ByVal PDFUrl As String, _
    ByVal FileName As String)

    Dim http As Object
    Dim url As String
    Dim jsonBody As String
    Dim accessToken As String

    ' Set your access token here (keep it private)
    accessToken = "EAAKvB3ZCC67oBO4ZAj1ZCLPyhXslXsjj9rZCYZCfkKzheuaZA5GZBDdqiaSI5dSBQvOOdNZBUcZBSbdFkSrECN2vyjktZClwOyIdczzD2YYRqZBC0rAbspJOQmkGvXzZABKqXTveIAaw4RHyng5syt9iZABw5aSVpjkiL4HAlrOP6Ekdr6SZCcQg0kfdSmZA1pZCW8bHjC2xrmsFFZAiCmiOmjI64OUFObyeQU39hzkZAMZBf1PnBDocAZDZD"

    ' Endpoint URL
    url = "https://graph.facebook.com/v19.0/" & 720804797779512# & "/messages"

    ' JSON payload
    jsonBody = "{""messaging_product"":""whatsapp""," & _
               """to"":""" & 7724958567# & """," & _
               """type"":""document""," & _
               """document"":{""link"":""" & PDFUrl & """,""filename"":""" & FileName & """}}"

    ' HTTP Request
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & accessToken
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    ' Show result
    MsgBox "Response: " & http.responseText

    Set http = Nothing
End Sub


Public Sub Upload_Ftp_WinSCP( _
    ByVal LHostName As String, _
    ByVal LUserName As String, _
    ByVal LPass As String, _
    ByVal LRemotePath As String, _
    ByVal LLocalFilePath As String)
    
    Dim sessionOptions As New WinSCPnet.sessionOptions
    Dim session As New WinSCPnet.session
    Dim transferResult As WinSCPnet.TransferOperationResult
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    ' Initialize FileSystemObject for file checks
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Validate local file exists before attempting upload
    If Not fso.FileExists(LLocalFilePath) Then
        MsgBox "Local file not found: " & LLocalFilePath, vbCritical, "File Error"
        Exit Sub
    End If
    
    ' Validate input parameters
    If Len(Trim(LHostName)) = 0 Then
        MsgBox "Host name cannot be empty.", vbCritical, "Parameter Error"
        Exit Sub
    End If
    
    If Len(Trim(LUserName)) = 0 Then
        MsgBox "Username cannot be empty.", vbCritical, "Parameter Error"
        Exit Sub
    End If
    
    If Len(Trim(LRemotePath)) = 0 Then
        MsgBox "Remote path cannot be empty.", vbCritical, "Parameter Error"
        Exit Sub
    End If
    
    ' Setup session options
    With sessionOptions
        .Protocol = WinSCPnet.Protocol_Ftp
        .HostName = LHostName
        .USERNAME = LUserName
        .PASSWORD = LPass
        .FtpMode = WinSCPnet.FtpMode_Passive
        .FtpSecure = WinSCPnet.FtpSecure_Explicit
        .GiveUpSecurityAndAcceptAnyTlsHostCertificate = True
    End With
    
    ' Open the session
    session.Open sessionOptions
    
    ' Perform the file upload
    Set transferResult = session.PutFiles(LLocalFilePath, LRemotePath, False, Nothing)
    
    ' Check transfer result
    If transferResult.IsSuccess Then
        MsgBox "Upload succeeded." & vbCrLf & "File: " & fso.GetFileName(LLocalFilePath), vbInformation, "Upload Success"
        ' Set your flag if needed (similar to download function)
        ' FlagUploadedToOurFTP = True
    Else
        ' Handle transfer failures
        Dim errorMsg As String
        errorMsg = "Upload failed:" & vbCrLf
        
        If transferResult.Failures.Count > 0 Then
            errorMsg = errorMsg & transferResult.Failures.Item(1).message
        Else
            errorMsg = errorMsg & "Unknown transfer error occurred."
        End If
        
        MsgBox errorMsg, vbCritical, "Upload Failed"
    End If
    
    ' Clean up
    session.Dispose
    Set fso = Nothing
    Exit Sub
    
ErrorHandler:
    Dim errorMessage As String
    
    ' Handle specific error types
    Select Case err.Number
        Case -2147467259 ' Common COM error for connection issues
            errorMessage = "Connection Error: Unable to connect to FTP server." & vbCrLf & _
                          "Please check hostname, username, and password." & vbCrLf & _
                          "Error: " & err.Description
        
        Case 53 ' File not found
            errorMessage = "Local File Path Error: " & vbCrLf & _
                          "Local file path is invalid or file doesn't exist." & vbCrLf & _
                          "Path: " & LLocalFilePath & vbCrLf & _
                          "Error: " & err.Description
        
        Case 76 ' Path not found
            errorMessage = "Remote Path Error: " & vbCrLf & _
                          "The remote path could not be found or is invalid." & vbCrLf & _
                          "Remote Path: " & LRemotePath & vbCrLf & _
                          "Error: " & err.Description
        
        Case 70 ' Permission denied
            errorMessage = "Permission Error: " & vbCrLf & _
                          "Access denied. Check file permissions or remote directory permissions." & vbCrLf & _
                          "Error: " & err.Description
        
        Case Else
            errorMessage = "WinSCP Upload Error: " & vbCrLf & _
                          "Error Number: " & err.Number & vbCrLf & _
                          "Description: " & err.Description
    End Select
    
    MsgBox errorMessage, vbCritical, "Upload Error"
    
    ' Clean up on error
    If Not session Is Nothing Then session.Dispose
    If Not fso Is Nothing Then Set fso = Nothing
End Sub

Option Explicit

' All-in-one function to send PDF via WhatsApp Business API
' Usage: result = SendPDFWhatsApp("1234567890", "C:\document.pdf", "Your_Access_Token", "Your_Phone_Number_ID", "Optional caption")

Public Function SendPDFWhatsApp(recipientNumber As String, pdfFilePath As String, accessToken As String, phoneNumberId As String, Optional caption As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim xhr As Object, fso As Object
    Dim boundary As String, requestBody As String, response As String
    Dim fileData As String, mediaId As String, jsonPayload As String
    Dim startPos As Long, endPos As Long
    
    ' Initialize result
    SendPDFWhatsApp = False
    
    ' Validate inputs
    If recipientNumber = "" Or pdfFilePath = "" Or accessToken = "" Or phoneNumberId = "" Then
        MsgBox "Missing required parameters", vbCritical
        Exit Function
    End If
    
    ' Check if file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(pdfFilePath) Then
        MsgBox "PDF file not found: " & pdfFilePath, vbCritical
        Exit Function
    End If
    
    ' STEP 1: Upload PDF file to WhatsApp
    ' Generate boundary for multipart form data
    boundary = "----FormBoundary" & Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")
    
    ' Read PDF file as binary
    fileData = ReadPDFBinary(pdfFilePath)
    If fileData = "" Then
        MsgBox "Failed to read PDF file", vbCritical
        Exit Function
    End If
    
    ' Build multipart form data
    requestBody = "--" & boundary & vbCrLf
    requestBody = requestBody & "Content-Disposition: form-data; name=""messaging_product""" & vbCrLf & vbCrLf
    requestBody = requestBody & "whatsapp" & vbCrLf
    requestBody = requestBody & "--" & boundary & vbCrLf
    requestBody = requestBody & "Content-Disposition: form-data; name=""type""" & vbCrLf & vbCrLf
    requestBody = requestBody & "application/pdf" & vbCrLf
    requestBody = requestBody & "--" & boundary & vbCrLf
    requestBody = requestBody & "Content-Disposition: form-data; name=""file""; filename=""" & fso.GetFileName(pdfFilePath) & """" & vbCrLf
    requestBody = requestBody & "Content-Type: application/pdf" & vbCrLf & vbCrLf
    requestBody = requestBody & fileData & vbCrLf
    requestBody = requestBody & "--" & boundary & "--" & vbCrLf
    
    ' Upload file
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "POST", "https://graph.facebook.com/v18.0/" & phoneNumberId & "/media", False
    xhr.setRequestHeader "Authorization", "Bearer " & accessToken
    xhr.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    xhr.send requestBody
    
    ' Check upload response
    Debug.Print "Upload Status: " & xhr.Status
    If xhr.Status <> 200 Then
        MsgBox "Upload failed. Status: " & xhr.Status & vbCrLf & "Response: " & xhr.responseText, vbCritical
        Exit Function
    End If
    
    ' Extract media ID from response
    response = xhr.responseText
    
    ' Debug: Show the actual response
    Debug.Print "Upload Response: " & response
    
    ' Try multiple JSON patterns for media ID extraction
    mediaId = ""
    
    ' Pattern 1: "id":"value"
    startPos = InStr(response, """id"":""")
    If startPos > 0 Then
        startPos = startPos + 6
        endPos = InStr(startPos, response, """")
        If endPos > startPos Then
            mediaId = Mid(response, startPos, endPos - startPos)
        End If
    End If
    
    ' Pattern 2: "id": "value" (with space)
    If mediaId = "" Then
        startPos = InStr(response, """id"": """)
        If startPos > 0 Then
            startPos = startPos + 7
            endPos = InStr(startPos, response, """")
            If endPos > startPos Then
                mediaId = Mid(response, startPos, endPos - startPos)
            End If
        End If
    End If
    
    ' Pattern 3: 'id':'value' (single quotes)
    If mediaId = "" Then
        startPos = InStr(response, "'id':'")
        If startPos > 0 Then
            startPos = startPos + 5
            endPos = InStr(startPos, response, "'")
            If endPos > startPos Then
                mediaId = Mid(response, startPos, endPos - startPos)
            End If
        End If
    End If
    
    If mediaId = "" Then
        MsgBox "Could not extract media ID from response:" & vbCrLf & response, vbCritical
        Exit Function
    End If
    
    Debug.Print "Extracted Media ID: " & mediaId
    
    ' STEP 2: Send WhatsApp message with PDF
    ' Build JSON payload
    jsonPayload = "{"
    jsonPayload = jsonPayload & """messaging_product"":""whatsapp"","
    jsonPayload = jsonPayload & """recipient_type"":""individual"","
    jsonPayload = jsonPayload & """to"":""" & recipientNumber & ""","
    jsonPayload = jsonPayload & """type"":""document"","
    jsonPayload = jsonPayload & """document"":{""id"":""" & mediaId & """"
    
    If caption <> "" Then
        ' Escape special characters in caption
        caption = Replace(caption, "\", "\\")
        caption = Replace(caption, """", "\""")
        caption = Replace(caption, vbCrLf, "\n")
        jsonPayload = jsonPayload & ",""caption"":""" & caption & """"
    End If
    
    jsonPayload = jsonPayload & "}}"
    
    ' Send message
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "POST", "https://graph.facebook.com/v18.0/" & phoneNumberId & "/messages", False
    xhr.setRequestHeader "Authorization", "Bearer " & accessToken
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.send jsonPayload
    
    ' Check send response
    If xhr.Status = 200 Then
        MsgBox "PDF sent successfully to " & recipientNumber, vbInformation
        SendPDFWhatsApp = True
    Else
        ' Parse error response for better error messages
        response = xhr.responseText
        If InStr(response, "131030") > 0 Then
            MsgBox "Error: Recipient phone number not in allowed list." & vbCrLf & _
                   "Add the number to your WhatsApp Business API allowed list in Meta Developers Console.", vbCritical
        ElseIf InStr(response, "OAuthException") > 0 Then
            MsgBox "Error: Invalid access token or permissions.", vbCritical
        Else
            MsgBox "Send failed. Status: " & xhr.Status & vbCrLf & response, vbCritical
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & err.Description, vbCritical
    SendPDFWhatsApp = False
End Function

Private Function ReadPDFBinary(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim stream As Object
    Dim fileData As Variant
    Dim i As Long
    Dim result As String
    
    ' Read file using ADODB.Stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.LoadFromFile filePath
    fileData = stream.Read
    stream.close
    
    ' Convert to string
    result = ""
    For i = 0 To UBound(fileData)
        result = result & Chr(fileData(i))
    Next i
    
    ReadPDFBinary = result
    Exit Function
    
ErrorHandler:
    ReadPDFBinary = ""
End Function

' Example usage in a button click or form
Private Sub Command1_Click()
    Dim result As Boolean
    Dim recipient As String
    Dim pdfPath As String
    Dim token As String
    Dim phoneId As String
    Dim message As String
    
    ' Set your values here
    recipient = "9630506575"  ' Phone number with country code, no +
    pdfPath = "C:\path\to\your\document.pdf"
    token = "EAAKvB3ZCC67oBOwiOkHzVsEBsYYkBx3u4T26bVqZCNqrTl8zqqZAG4j4BJLvFQpXmh0oZAu3mroBHCacMXNZC3qowRoAOR8ZAj0ZBehWyZABLti54u7TUHIwTNVPRNweNGowZCPRoZCX1gZC850C4uZBnXEQa41dd1hGqdua1Yjn8Op6jyzU3zZBtZBdol3PVNLZC84qgZDZD"
    phoneId = "720804797779512"
    message = "Here is your PDF document"
    
    ' Send PDF
    result = SendPDFWhatsApp(recipient, pdfPath, token, phoneId, message)
    
    If result Then
        Debug.Print "PDF sent successfully!"
    Else
        Debug.Print "Failed to send PDF"
    End If
End Sub




Private Sub Command3_Click()
Dim result As Boolean
    Dim recipient As String
    Dim pdfPath As String
    Dim token As String
    Dim phoneId As String
    Dim message As String
    
    ' Set your values here
    recipient = "7724958567"  ' Phone number with country code, no +
    pdfPath = "D:\Backup\BhavCopy_NSE_FO_0_0_0_20250506.csv\Docker_ From Basics to Advanced.pdf"
    token = "EAAKvB3ZCC67oBO5XgH5ONtpC1bH5XE22fVNECBE3Eu3ZCpzZBZCd924NotjGZAeZC9WD15et4nmPHfuEzOAjjqMAl6LtjAlamybxr1ayqkuMU71RvRxPnskXN0xoZCQ7bVc6GOz2N0NJcvRZBtvQt8uuh2poK8Xd9wbJiY2eyzeoJGJnrAWSbV6CMJZCG7RV9TtOsHu1mHoglLUdFqlpoSVsAGVSk2EdC3ua6Wu4Iryp6MQZDZD"
    phoneId = "720804797779512"
    message = "Here is your PDF document"
    
    ' Send PDF
    result = SendPDFWhatsApp(recipient, pdfPath, token, phoneId, message)
    
    If result Then
        Debug.Print "PDF sent successfully!"
    Else
        Debug.Print "Failed to send PDF"
    End If
End Sub

Private Sub ExNameDb_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub ExCombo_GotFocus()
    Sendkeys "%{down}"
End Sub
Private Sub ExNameDb_Validate(Cancel As Boolean)
If LenB(ExNameDb.text) < 1 Then
    MsgBox " Please Select Exchange"
    Cancel = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Form_Load()
    Call CANCEL_REC
    OptCttNo.Value = True
    Set ExMastRec = Nothing: Set ExMastRec = New ADODB.Recordset
    ExMastRec.Open "SELECT EXID,EXCODE,EXNAME FROM EXMAST  WHERE COMPCODE= " & GCompCode & " ORDER BY EXNAME ", Cnn, adOpenKeyset, adLockReadOnly
    If Not ExMastRec.EOF Then
        Set ExNameDb.RowSource = ExMastRec
        ExNameDb.ListField = "EXCODE"
        ExNameDb.BoundColumn = "EXCODE"
        Set ExCombo.RowSource = ExMastRec
        ExCombo.ListField = "EXCODE"
        ExCombo.BoundColumn = "EXCODE"
    End If
End Sub
Private Sub Form_Paint()
    'Me.BackColor = GETMAIN.BackColor
End Sub
Private Sub Form_Resize()
'    Frame1.Left = ((Me.ScaleWidth - 4024) * 15) / 100
'    Frame2.Left = ((Me.ScaleWidth - 4024) * 15) / 100
'    Frame3.Left = ((Me.ScaleWidth - 4024) * 15) / 100
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
    Call LIST_ITEM
End Sub
Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label17.ForeColor = &HC00000
End Sub

Private Sub TxtItemCode_Validate(Cancel As Boolean)
    Dim TRec As ADODB.Recordset
    TxtItemCode.text = UCase(TxtItemCode.text)
    If Fb_Press = 1 Then
        Set TRec = Nothing
        Set TRec = New ADODB.Recordset
        TRec.Open "SELECT COMPCODE FROM ITEMMAST WHERE COMPCODE=" & GCompCode & " AND ITEMCODE='" & TxtItemCode.text & "'", Cnn, adOpenForwardOnly, adLockReadOnly
        If Not TRec.EOF Then
            MsgBox "Item code already exists.", vbExclamation, "Warning"
            Cancel = True
        End If
    End If
    TxtItemName.text = UCase(TxtItemCode.text)
    TxtExSymbol.text = UCase(TxtItemCode.text)
End Sub

Private Sub TxtLot_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub TxtLot_GotFocus()
    TxtLot.SelLength = Len(TxtLot.text)
End Sub
Private Sub TxtLot_Validate(Cancel As Boolean)
    TxtLot.text = Format(TxtLot.text, "0.0000")
End Sub

Private Sub TXtCloseID_GotFocus()
    TXtCloseID.SelLength = Len(TXtCloseID.text)
End Sub
Private Sub TXtCloseID_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub

Sub LIST_ITEM()
On Error GoTo err1

    Dim TRec As ADODB.Recordset
    Screen.MousePointer = 11
    Call Get_Selection(12)
    mysql = "SELECT ITEMCODE,ITEMNAME,EXHCODE,LOT,EXCHANGECODE,PRICEUNIT,QTYUNIT,SCGROUP,'' as 'RISKMAPP',STMAPP,CTTDATE "  'CTTAPP.RISKMAPP
    mysql = mysql & " FROM ITEMMAST WHERE COMPCODE=" & GCompCode & " ORDER BY EXCHANGECODE,ITEMNAME"
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly

    Set RDCREPO = RDCAPP.OpenReport(GReportPath & "RptIList.RPT", 1)
    
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
    
err1:
If err.Number <> 0 Then
    MsgBox err.Description
End If
Screen.MousePointer = 0
End Sub
Private Sub TxtFilterCode_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
Private Sub Set_TaxRec()
    Set RecITTax = Nothing: Set RecITTax = New ADODB.Recordset
    RecITTax.Fields.Append "StartDate", adDate, , adFldIsNullable
    RecITTax.Fields.Append "EndDate", adDate, , adFldIsNullable
    RecITTax.Fields.Append "RiskMFees", adDouble, , adFldIsNullable
    RecITTax.Fields.Append "SEBITAX", adDouble, , adFldIsNullable
    RecITTax.Open , , adOpenKeyset, adLockBatchOptimistic
End Sub
Private Sub ITTaxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And ITTaxGrid.Col = 2 Then
        RecITTax.MoveNext
        If RecITTax.EOF Then
            RecITTax.AddNew
            RecITTax!RISKMFEES = 0
            RecITTax!SEBITAX = 0
            RecITTax.Update
        End If
        ITTaxGrid.Col = 0
    ElseIf KeyCode = 13 Then
          Sendkeys "{tab}"
    End If
End Sub

Private Sub FillDataGrid()
    Set ItemRec = Nothing
    Set ItemRec = New ADODB.Recordset
    mysql = "SELECT ExchangeCode as ExCode,ItemId,ItemCode,ItemName,ExhCode,Lot,CTTApp,RISKMApp,STMApp,SCGroup,PriceUnit,RegularLot,QtyUnit,EXID,CITEMID,CTTDATE   "
    mysql = mysql & ",ITALIAS_1,ITALIAS_2,ITALIAS_3,ITALIAS_4,ITALIAS_5  "
    mysql = mysql & " FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " "
    If LenB(TxtFilterCode.text) <> 0 Then mysql = mysql & " AND  UPPER(ITEMNAME) LIKE '" & Trim(UCase(TxtFilterCode.text)) & "%' "
    If LenB(ExCombo.BoundText) > 0 Then mysql = mysql & " AND EXCHANGECODE = '" & ExCombo.BoundText & "'"
    mysql = mysql & " ORDER BY ITEMNAME "
    ItemRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not ItemRec.EOF Then
        Set DataGrid1.DataSource = ItemRec
        DataGrid1.ReBind
        DataGrid1.Refresh
        DataGrid1.Columns(0).Width = 1500:           DataGrid1.Columns(1).Width = 800: DataGrid1.Columns(2).Width = 2500
        DataGrid1.Columns(3).Width = 2500:           DataGrid1.Columns(4).Width = 2000
        DataGrid1.Columns(5).Width = 1000:
        DataGrid1.Columns(1).Alignment = dbgRight
        DataGrid1.Columns(5).Alignment = dbgRight
        DataGrid1.Columns(5).NumberFormat = "0.00"
        DataGrid1.Refresh
    Else
        MsgBox "No Records Found"
    End If
End Sub
Private Sub DataGrid1_Click()
If ItemRec.RecordCount > 0 Then
    If ItemRec.EOF Then ItemRec.MoveFirst
    DataGrid1.Col = 2
    TxtItemCode.text = DataGrid1.text
    DataGrid1.Col = 3
    TxtItemName.text = DataGrid1.text
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
        ItemRec.MoveFirst
        Do While Not ItemRec.EOF
            If Left$(ItemRec!ITEMName, 1) <> LChar Then
                ItemRec.MoveNext
            Else
                Exit Do
            End If
        Loop
        If ItemRec.EOF Then ItemRec.MoveFirst
    End If
End Sub
