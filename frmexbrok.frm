VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmExBrok2 
   Caption         =   "Exchange Wise Brokerage Setup"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "Times New Roman"
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
   ScaleHeight     =   10935
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11760
      TabIndex        =   23
      Top             =   2040
      Width           =   6255
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3360
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox ChkUpdLastSettle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update Last Settlemet "
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
         TabIndex        =   26
         Top             =   1125
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox ChkBrokLock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lock Brokerage"
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
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   4560
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
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
         Value           =   43531.4583217593
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Brokerage Lock Setup"
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
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "Upto Date "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   3360
         TabIndex        =   27
         Top             =   660
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   18135
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Brokerage Setup"
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
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   18015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   18015
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   495
         Left            =   11640
         TabIndex        =   16
         Top             =   2760
         Width           =   6255
         Begin VB.OptionButton OptItem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item Wise"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   8
            Top             =   80
            Width           =   1455
         End
         Begin VB.OptionButton OptExchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Exchange Wise"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   80
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
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
         Height          =   5055
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   16695
         Begin VB.ComboBox BrokTypeCombo2 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            ItemData        =   "frmexbrok.frx":0000
            Left            =   14160
            List            =   "frmexbrok.frx":0055
            TabIndex        =   13
            Top             =   4080
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox BrokTypeCombo 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            ItemData        =   "frmexbrok.frx":0284
            Left            =   9360
            List            =   "frmexbrok.frx":02D9
            TabIndex        =   12
            Top             =   4080
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox MarginCombo 
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
            Height          =   360
            ItemData        =   "frmexbrok.frx":0507
            Left            =   11760
            List            =   "frmexbrok.frx":051D
            TabIndex        =   11
            Top             =   4080
            Visible         =   0   'False
            Width           =   1935
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   5055
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   16695
            _ExtentX        =   29448
            _ExtentY        =   8916
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Exchange Wise Brokerage"
            TabPicture(0)   =   "frmexbrok.frx":058A
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "DataGrid1"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "ItemWise Brokerage"
            TabPicture(1)   =   "frmexbrok.frx":05A6
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "DataGrid2"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "ChkDelItemBrok"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).ControlCount=   2
            Begin VB.CheckBox ChkDelItemBrok 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Delete Item Wise Brokerage of Selected Clients and Items"
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
               Height          =   375
               Left            =   5880
               TabIndex        =   31
               Top             =   360
               Width           =   6255
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   4545
               Left            =   -75000
               TabIndex        =   32
               Top             =   480
               Width           =   16680
               _ExtentX        =   29422
               _ExtentY        =   8017
               _Version        =   393216
               AllowUpdate     =   -1  'True
               AllowArrows     =   -1  'True
               BackColor       =   -2147483628
               ForeColor       =   4194368
               HeadLines       =   1
               RowHeight       =   19
               TabAction       =   1
               FormatLocked    =   -1  'True
               AllowAddNew     =   -1  'True
               AllowDelete     =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
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
               ColumnCount     =   20
               BeginProperty Column00 
                  DataField       =   "Party"
                  Caption         =   "Party"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "Excode"
                  Caption         =   "Ex Code"
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
               BeginProperty Column02 
                  DataField       =   "INSTTYPE"
                  Caption         =   "InstType"
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
               BeginProperty Column04 
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
               BeginProperty Column05 
                  DataField       =   "stamprate"
                  Caption         =   "Sell Brok"
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
                  DataField       =   "BrokRate2"
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
               BeginProperty Column07 
                  DataField       =   "SEBITAX"
                  Caption         =   "Sell Brok 2"
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
               BeginProperty Column09 
                  DataField       =   "TRANRATE"
                  Caption         =   "Tran Rate"
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
               BeginProperty Column10 
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
               BeginProperty Column11 
                  DataField       =   "UPTOSTDT"
                  Caption         =   "Set. Date"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "dd/MM/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   3
                  EndProperty
               EndProperty
               BeginProperty Column12 
                  DataField       =   "MINRATE"
                  Caption         =   "Min Rate"
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
               BeginProperty Column13 
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
               BeginProperty Column14 
                  DataField       =   "MBROKRATE"
                  Caption         =   "MBrokRate"
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
               BeginProperty Column15 
                  DataField       =   "MBROKRATE2"
                  Caption         =   "MBrokRate2"
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
               BeginProperty Column16 
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
               BeginProperty Column17 
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
               BeginProperty Column18 
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
               BeginProperty Column19 
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
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   2
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   2520
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   1005.165
                  EndProperty
                  BeginProperty Column02 
                     Locked          =   -1  'True
                     Object.Visible         =   -1  'True
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   2399.811
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     ColumnWidth     =   1395.213
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   1
                     Object.Visible         =   0   'False
                     ColumnWidth     =   1305.071
                  EndProperty
                  BeginProperty Column06 
                     Alignment       =   1
                     ColumnWidth     =   1395.213
                  EndProperty
                  BeginProperty Column07 
                     Alignment       =   1
                     Object.Visible         =   0   'False
                     ColumnWidth     =   1305.071
                  EndProperty
                  BeginProperty Column08 
                     Alignment       =   1
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column09 
                     Alignment       =   1
                     ColumnWidth     =   1094.74
                  EndProperty
                  BeginProperty Column10 
                     ColumnWidth     =   1814.74
                  EndProperty
                  BeginProperty Column11 
                     Locked          =   -1  'True
                     ColumnWidth     =   1305.071
                  EndProperty
                  BeginProperty Column12 
                     Alignment       =   1
                     ColumnWidth     =   900.284
                  EndProperty
                  BeginProperty Column13 
                     ColumnWidth     =   1814.74
                  EndProperty
                  BeginProperty Column14 
                     Alignment       =   1
                     ColumnWidth     =   1094.74
                  EndProperty
                  BeginProperty Column15 
                     Alignment       =   1
                     ColumnWidth     =   1094.74
                  EndProperty
                  BeginProperty Column16 
                     ColumnWidth     =   1395.213
                  EndProperty
                  BeginProperty Column17 
                     Alignment       =   1
                     ColumnWidth     =   900.284
                  EndProperty
                  BeginProperty Column18 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty Column19 
                     Locked          =   -1  'True
                     Object.Visible         =   0   'False
                     ColumnWidth     =   1200.189
                  EndProperty
               EndProperty
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   4185
               Left            =   0
               TabIndex        =   33
               Top             =   840
               Width           =   16680
               _ExtentX        =   29422
               _ExtentY        =   7382
               _Version        =   393216
               AllowUpdate     =   -1  'True
               AllowArrows     =   -1  'True
               BackColor       =   16777215
               ForeColor       =   4194368
               HeadLines       =   1
               RowHeight       =   19
               TabAction       =   1
               FormatLocked    =   -1  'True
               AllowAddNew     =   -1  'True
               AllowDelete     =   -1  'True
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
               ColumnCount     =   20
               BeginProperty Column00 
                  DataField       =   "Party"
                  Caption         =   "Party"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "ItemCode"
                  Caption         =   "ItemCode"
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
               BeginProperty Column02 
                  DataField       =   "INSTTYPE"
                  Caption         =   "InstType"
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
               BeginProperty Column04 
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
               BeginProperty Column05 
                  DataField       =   "SellBrok"
                  Caption         =   "Sell Brok"
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
                  DataField       =   "BrokRate2"
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
               BeginProperty Column07 
                  DataField       =   "SellBrok2"
                  Caption         =   "Sell Brok 2"
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
               BeginProperty Column09 
                  DataField       =   "TRANRATE"
                  Caption         =   "Tran Rate"
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
               BeginProperty Column10 
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
               BeginProperty Column11 
                  DataField       =   "UPTOSTDT"
                  Caption         =   "Set. Date"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "dd/MM/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   3
                  EndProperty
               EndProperty
               BeginProperty Column12 
                  DataField       =   "MINRATE"
                  Caption         =   "Min Rate"
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
               BeginProperty Column13 
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
               BeginProperty Column14 
                  DataField       =   "MBROKRATE"
                  Caption         =   "MBrokRate"
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
               BeginProperty Column15 
                  DataField       =   "MBROKRATE2"
                  Caption         =   "MBrokRate2"
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
               BeginProperty Column16 
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
               BeginProperty Column17 
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
               BeginProperty Column18 
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
               BeginProperty Column19 
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
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   2
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   2520
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   1995.024
                  EndProperty
                  BeginProperty Column02 
                     Locked          =   -1  'True
                     Object.Visible         =   -1  'True
                     ColumnWidth     =   1005.165
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   2399.811
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     ColumnWidth     =   1395.213
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   1
                     Object.Visible         =   0   'False
                     ColumnWidth     =   1005.165
                  EndProperty
                  BeginProperty Column06 
                     Alignment       =   1
                     ColumnWidth     =   1395.213
                  EndProperty
                  BeginProperty Column07 
                     Alignment       =   1
                     Object.Visible         =   0   'False
                     ColumnWidth     =   1005.165
                  EndProperty
                  BeginProperty Column08 
                     Alignment       =   1
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column09 
                     Alignment       =   1
                     ColumnWidth     =   1094.74
                  EndProperty
                  BeginProperty Column10 
                     ColumnWidth     =   1814.74
                  EndProperty
                  BeginProperty Column11 
                     Locked          =   -1  'True
                     ColumnWidth     =   1305.071
                  EndProperty
                  BeginProperty Column12 
                     Alignment       =   1
                     ColumnWidth     =   900.284
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
                     ColumnWidth     =   1200.189
                  EndProperty
                  BeginProperty Column16 
                     ColumnWidth     =   1395.213
                  EndProperty
                  BeginProperty Column17 
                     Alignment       =   1
                     ColumnWidth     =   900.284
                  EndProperty
                  BeginProperty Column18 
                     Object.Visible         =   -1  'True
                  EndProperty
                  BeginProperty Column19 
                     Locked          =   -1  'True
                     Object.Visible         =   0   'False
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
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
         Height          =   3375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   18015
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Branch wise Parties"
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
            TabIndex        =   37
            Top             =   50
            Width           =   2535
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
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
            TabIndex        =   35
            Top             =   2955
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1215
            Left            =   11640
            TabIndex        =   17
            Top             =   0
            Width           =   6255
            Begin VB.CommandButton OkCmd 
               Caption         =   "OK"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   120
               TabIndex        =   19
               Top             =   600
               Width           =   1215
            End
            Begin VB.ComboBox InstTypeCombo 
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
               ItemData        =   "frmexbrok.frx":05C2
               Left            =   1320
               List            =   "frmexbrok.frx":05CF
               TabIndex        =   18
               Top             =   0
               Width           =   1455
            End
            Begin MSDataListLib.DataCombo SettleDCombo 
               Height          =   360
               Left            =   4560
               TabIndex        =   20
               Top             =   120
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   635
               _Version        =   393216
               Style           =   2
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
            Begin VB.Label Label8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "To Activate Brokerae Lock SetUp First Select Brokerage Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1440
               TabIndex        =   34
               Top             =   600
               Width           =   3615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000011&
               BackStyle       =   0  'Transparent
               Caption         =   "Inst Type"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   240
               Left            =   120
               TabIndex        =   22
               Top             =   60
               Width           =   960
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000011&
               BackStyle       =   0  'Transparent
               Caption         =   "Show Brok Date "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   255
               Left            =   2880
               TabIndex        =   21
               Top             =   90
               Width           =   1680
            End
         End
         Begin VB.CheckBox PartyChk 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
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
            Height          =   285
            Left            =   2760
            TabIndex        =   3
            Top             =   2955
            Width           =   1335
         End
         Begin VB.CheckBox ExchangeChk 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
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
            Height          =   285
            Left            =   6480
            TabIndex        =   2
            Top             =   2955
            Width           =   1335
         End
         Begin MSComctlLib.ListView PartyLst 
            Height          =   2820
            Left            =   2760
            TabIndex        =   4
            Top             =   0
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   4974
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
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Party Name"
               Object.Width           =   6350
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ACCID"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView ExchnageList 
            Height          =   2820
            Left            =   6480
            TabIndex        =   5
            Top             =   0
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   4974
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
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Exchange"
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
         Begin MSComctlLib.ListView ItemList 
            Height          =   2820
            Left            =   8880
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   4974
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
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ITEMCODE"
               Object.Width           =   5185
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "EXCODE"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "EXID"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "ITEMID"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView BranchList 
            Height          =   2460
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   4339
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Branch Name"
               Object.Width           =   5185
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "FMLYCODE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "FMLYID"
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
            TabIndex        =   6
            Top             =   360
            Width           =   2115
            WordWrap        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "FrmExBrok2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LSParties As String::           Dim LInstType  As String:           Dim LSItems As String:              Dim LExCodes As String
Dim GridColVal As String:           Dim CountRow As Double:             Dim LSettlementDt As String:        Public Fb_Press As Byte
Dim RECGRID As ADODB.Recordset:     Dim RecGrid2 As ADODB.Recordset:    Dim UptoStDtRec As ADODB.Recordset
Dim LFmlyIDs  As String: Dim ListIt As ListItem:

Sub ADD_NEW()
    Frame3.Enabled = True:
'    Frame3.BackColor = &HFFC0C0
    Call Get_Selection(1):
    PartyLst.Enabled = True
    PartyLst.SetFocus
    SSTab1.Tab = 0
End Sub
Sub CANCEL_REC()
    Dim I As Integer
    Fb_Press = 0:
    vcDTP1.Value = Date:            vcDTP1.Visible = False
    Label6.Visible = False:         ChkUpdLastSettle.Visible = False:       CmdApply.Visible = False
    PartyChk.Value = 0:             ExchangeChk.Value = 0:
    ChkBrokLock.Value = 0:          ChkBrokLock.Visible = False:
    SettleDCombo.BoundText = vbNullString
    SettleDCombo.text = vbNullString:
    
    For I = 1 To PartyLst.ListItems.Count
        PartyLst.ListItems.Item(I).Checked = False
    Next I
    
    For I = 1 To ExchnageList.ListItems.Count
        ExchnageList.ListItems.Item(I).Checked = False
    Next I
    
    For I = 1 To ItemList.ListItems.Count
        ItemList.ListItems.Item(I).Checked = False
    Next I
    Call RecSet
    Call RecSet2
    UptoStDtRec.Requery
    
    SettleDCombo.Refresh
    Set DataGrid1.DataSource = RECGRID:    DataGrid1.ReBind:    DataGrid1.Refresh:
    
    Set DataGrid2.DataSource = RecGrid2:   DataGrid2.ReBind:    DataGrid2.Refresh:
    Frame4.BackColor = &H8080FF
    Frame4.Enabled = False
    'Frame3.BackColor = &H8080FF
    Frame3.Enabled = False
    Frame8.Enabled = True
    
    BrokTypeCombo.Visible = False:
    BrokTypeCombo2.Visible = False
    
    
    Call Get_Selection(13)
End Sub
Sub Save_Rec()
    Dim LInstType As String * 3
    On Error GoTo err1
    'Frame3.BackColor = &H8080FF
   ' Frame4.BackColor = &H8080FF
    Frame3.Enabled = False:
    Frame4.Enabled = False
    
    mysql = "DELETE FROM PEXBROK WHERE UptoStDt IS NULL"
    Cnn.Execute mysql
    mysql = "DELETE FROM PITBROK WHERE UptoStDt IS NULL"
    Cnn.Execute mysql
    
    If InstTypeCombo.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf InstTypeCombo.ListIndex = 1 Then
        LInstType = "OPT"
    Else
        LInstType = "CSH"
    End If
    If OptExchange.Value = True Then
        Call Save_ExBrok
        If ChkDelItemBrok.Value = True Then
            Save_ItemBrok
        End If
    Else
        Call Save_ItemBrok
    End If
        'If AllParties = True Then LSParties = vbNullString
        'If AllExcodes = True Then LExCodes = vbNullString
        'Call Delete_Inv_D(LSParties, LExCodes, vbNullString, GFinBegin)
        'Call Update_BrokTran(LSParties, LExCodes, vbNullString, vbNullString, GFinBegin, GFinEnd)
        Call Update_Charges(LSParties, LExCodes, vbNullString, vbNullString, GFinBegin, GFinEnd, False)
        Cnn.CommitTrans: CNNERR = False
        Cnn.BeginTrans: CNNERR = True
        If BILL_GENERATION(GFinBegin, GFinEnd, vbNullString, LSParties, LExCodes) Then
            Cnn.CommitTrans: CNNERR = False
        Else
            Cnn.RollbackTrans: CNNERR = False
        End If
        'Call Chk_Billing
        MsgBox "Brokerage Succesfully Updated "
        GETMAIN.ProgressBar1.Visible = False
        Call CANCEL_REC
        Exit Sub
err1:
    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
    
    If CNNERR = True Then
       Cnn.RollbackTrans
    End If
End Sub

Private Sub BranchList_Click()
    Dim I As Integer
    Dim RecSauda As ADODB.Recordset
    LFmlyIDs = vbNullString
    For I = 1 To BranchList.ListItems.Count
        If BranchList.ListItems(I).Checked = True Then
            If LenB(LFmlyIDs) > 0 Then LFmlyIDs = LFmlyIDs & ", "
            LFmlyIDs = LFmlyIDs & BranchList.ListItems(I).ListSubItems(2) & ""
        End If
  Next I
  PartyLst.ListItems.Clear
  If LFmlyIDs = "" Then
    Call get_partlist
    Me.MousePointer = 0
    Exit Sub
  End If
    mysql = "SELECT DISTINCT ACC.ACCID,ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC,PEXBROK AS PB WHERE ACC.COMPCODE =" & GCompCode & " "
    mysql = mysql & " AND ACC.ACCID =PB.ACCID AND ACC.ACCID IN (SELECT DISTINCT ACCID FROM ACCFMLYD WHERE  FMLYID IN  (" & LFmlyIDs & ")) ORDER BY ACC.NAME"
    Set RecSauda = Nothing: Set RecSauda = New ADODB.Recordset: RecSauda.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    While Not RecSauda.EOF
        Set ListIt = PartyLst.ListItems.Add(, , RecSauda!NAME)
        ListIt.SubItems(1) = RecSauda!AC_CODE
        ListIt.SubItems(2) = RecSauda!ACCID
        RecSauda.MoveNext
    Wend
End Sub

Private Sub Check1_Click()

    Dim I As Integer
    For I = 1 To BranchList.ListItems.Count
        BranchList.ListItems.Item(I).Checked = False
    Next I
    
    If Check1.Value = 1 Then
        BranchList.Visible = True: Check4.Visible = True
        BranchList.Enabled = True: Check4.Enabled = True
    Else
        BranchList.Visible = False: Check4.Visible = False
        BranchList.Enabled = False: Check4.Enabled = False
    End If
    Call get_partlist
End Sub
Private Sub get_partlist()
    PartyLst.ListItems.Clear
    Dim AccRecADO As ADODB.Recordset
    Set AccRecADO = Nothing: Set AccRecADO = New ADODB.Recordset
    mysql = "SELECT DISTINCT ACC.ACCID,ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC,PEXBROK AS PB WHERE ACC.COMPCODE =" & GCompCode & " "
    mysql = mysql & " AND ACC.ACCID =PB.ACCID ORDER BY ACC.NAME"
    AccRecADO.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRecADO.EOF Then
        While Not AccRecADO.EOF
            Set ListIt = PartyLst.ListItems.Add(, , AccRecADO!NAME)
            ListIt.SubItems(1) = AccRecADO!AC_CODE
            ListIt.SubItems(2) = AccRecADO!ACCID
            AccRecADO.MoveNext
        Wend
    End If
End Sub

Private Sub Check4_Click()
Dim I As Integer
    For I = 1 To BranchList.ListItems.Count
        If Check4.Value = 1 Then
            BranchList.ListItems.Item(I).Checked = True
        Else
            BranchList.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub

Private Sub ChkBrokLock_Click()
Dim TRec As ADODB.Recordset
vcDTP1.MinDate = GFinBegin
If ChkBrokLock.Value = 1 Then
    If OptExchange.Value = True Then
        mysql = "SELECT MAX(UPTOSTDT) AS MDT FROM PEXBROK WHERE COMPCODE =" & GCompCode & ""
        mysql = mysql & " AND INSTTYPE='" & LInstType & "'"
        mysql = mysql & " AND UPTOSTDT <'" & Format(DateValue(SettleDCombo.BoundText), "YYYY/MM/DD") & "'"
        mysql = mysql & " AND AC_CODE IN (" & LSParties & ")"
        mysql = mysql & " AND EXID IN (" & LExCodes & ") "
    Else
        mysql = "SELECT MAX(UPTOSTDT) AS MDT FROM PITBROK WHERE COMPCODE =" & GCompCode & ""
        mysql = mysql & " AND INSTTYPE='" & LInstType & "'"
        mysql = mysql & " AND UPTOSTDT <'" & Format(DateValue(SettleDCombo.BoundText), "YYYY/MM/DD") & "'"
        mysql = mysql & " AND AC_CODE IN (" & LSParties & ")"
        mysql = mysql & " AND EXID IN (" & LExCodes & ") "
        If LenB(LSItems) > 0 Then mysql = mysql & " AND ITEMID IN (" & LSItems & ") "
    End If
    
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then
        If Not IsNull(TRec!MDt) Then
            vcDTP1.MinDate = TRec!MDt + 1
        End If
    End If
    Set TRec = Nothing
    vcDTP1.Visible = True
    Label6.Visible = True
    CmdApply.Visible = True
    If OptExchange.Value = False Then
        ChkUpdLastSettle.Visible = True
    Else
        ChkUpdLastSettle.Visible = False
    End If
Else
    vcDTP1.Visible = False
    Label6.Visible = False
    ChkUpdLastSettle.Visible = False
    CmdApply.Visible = False
End If
End Sub

Private Sub CmdApply_Click()
Dim TempRec As ADODB.Recordset:         Dim TempRec2 As ADODB.Recordset
Dim LStdTDate As Date:                  Dim AccRec As ADODB.Recordset
Dim TRec As ADODB.Recordset:            Dim TRec2 As ADODB.Recordset
Dim LXAc_Code As String

SettleDCombo.BoundText = vbNullString
SettleDCombo.text = vbNullString
Frame8.Enabled = False
If OptExchange.Value = True Then
    If RECGRID.RecordCount > 0 Then
        Set TempRec = Nothing
        Set TempRec = RECGRID.Clone
        TempRec.MoveFirst
        Cnn.BeginTrans: CNNERR = True
        LSItems = vbNullString
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If Not IsNull(TempRec!excode) Then
                If LenB(TempRec!excode) > 0 Then
                    mysql = "DELETE FROM PEXBROK WHERE COMPCODE =" & GCompCode & " AND INSTTYPE='" & LInstType & "' AND AC_CODE = '" & TempRec!PARTYCODE & "'"
                    mysql = mysql & " AND EXCODE ='" & TempRec!excode & "' AND UpToStDt = '" & Format(TempRec!UPTOSTDT, "yyyy/MM/dd") & "' "
                    mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    If ChkBrokLock.Value = 1 Then
                        LStdTDate = DateValue(vcDTP1.Value)
                        If LStdTDate > DateValue(GSysLockDt) Then
                            'MYSQL = "EXEC INSERT_PEXBROK " & GCompCode & ",'" & TempRec!PARTYCODE & "','" & TempRec!EXCODE & "','" & Left$(TempRec!BrokType, 1) & "'," & Val(TempRec!BROKRATE) & "," & Val(TempRec!BROKRATE2) & "," & Val(TempRec!STDRATE) & "," & Val(TempRec!TRANRATE) & ",'" & Left$(TempRec!TRANTYPE, 1) & "',"
                            'MYSQL = MYSQL & "0," & Val(TempRec!MinRate) & ",'" & Left$(TempRec!MBROKTYPE, 1) & "'," & Val(TempRec!MBrokRate) & "," & Val(TempRec!MBrokRate2) & ",'" & Left$(TempRec!MARTYPE, 1) & "'," & Val(TempRec!MARRATE) & ",'" & Format(LStdTDate, "YYYY/MM/DD") & "','" & LInstType & "'," & TempRec!STAMPRATE & "," & TempRec!SEBITAX & "," & TempRec!EXID & "," & TempRec!ACCID & ""
                            'Cnn.Execute MYSQL
                            Call PInsert_PExBrok(TempRec!PARTYCODE, TempRec!excode, Left$(TempRec!broktype, 1), TempRec!brokrate, TempRec!BROKRATE2, TempRec!STDRATE, TempRec!TRANRATE, Left$(TempRec!TRANTYPE, 1), 0, Val(TempRec!MinRate), Left$(TempRec!MBROKTYPE, 1), Val(TempRec!MBROKRATE), Val(TempRec!MBROKRATE2), Left(TempRec!MARTYPE, 1), Val(TempRec!MARRATE), LStdTDate, LInstType, TempRec!STAMPRATE, TempRec!SEBITAX, TempRec!EXID, TempRec!ACCID)
                            
                        Else
                            MsgBox "Sorry System Locked.  No Modification Allowed"
                            Exit Do
                        End If
                    End If
                End If
            End If
            TempRec.MoveNext
        Loop
        LSettlementDt = GFinEnd
        Set TempRec = Nothing:        Set AccRec = Nothing:        Set AccRec = New ADODB.Recordset
        mysql = "SELECT ACCID ,AC_CODE FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " "
        mysql = mysql & " AND AC_CODE IN (" & LSParties & ") ORDER BY AC_CODE"
        AccRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        LXAc_Code = vbNullString
        While Not AccRec.EOF
            mysql = "SELECT EXID,EXCODE FROM EXMAST WHERE COMPCODE  = " & GCompCode & " "
            mysql = mysql & " AND EXID  IN (" & LExCodes & ")"
            mysql = mysql & "  ORDER BY EXCODE "
            Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not TRec.EOF
                DoEvents
                LXAc_Code = Get_PEXBROK_AC_CODE(AccRec!ACCID, TRec!EXID, DateValue(GFinEnd), LInstType)
                'MYSQL = "SELECT AC_CODE FROM PEXBROK WHERE COMPCODE=" & GCompCode & " AND AC_CODE ='" & AccRec!AC_CODE & "' "
                'MYSQL = MYSQL & " AND EXCODE ='" & TRec!EXCODE & "' AND UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE ='" & LInstType & "'"
                'Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
                If LenB(LXAc_Code) < 1 Then
                'If TRec2.EOF Then
                    If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                        mysql = "EXEC INSERT_PEXBROK " & GCompCode & ",'" & AccRec!AC_CODE & "','" & TRec!excode & "','P',0,0,0,0,'P',"
                        mysql = mysql & "0,0,'P',0,0,'I',0,'" & Format(LSettlementDt, "YYYY/MM/DD") & "','" & LInstType & "',0,0," & TRec!EXID & "," & AccRec!ACCID & ""
                        Cnn.Execute mysql
                    End If
                End If
                TRec.MoveNext
            Loop
            Set TRec = Nothing
            AccRec.MoveNext
        Wend
        Cnn.CommitTrans
        CNNERR = False
        Set AccRec = Nothing
    End If
    Fill_ExchangeGrid
Else
    If RecGrid2.RecordCount > 0 Then
        Set TempRec2 = Nothing:        Set TempRec2 = RecGrid2.Clone
        TempRec2.MoveFirst:        Cnn.BeginTrans: CNNERR = True
        TempRec2.MoveFirst
        Do While Not TempRec2.EOF
            If Not IsNull(TempRec2!ITEMCODE) Then
                If LenB(TempRec2!ITEMCODE) > 0 Then
                    mysql = "DELETE FROM PITBROK WHERE COMPCODE =" & GCompCode & " AND INSTTYPE='" & LInstType & "' AND AC_CODE = '" & TempRec2!PARTYCODE & "' "
                    mysql = mysql & " AND ITEMID =" & TempRec2!itemid & " AND UpToStDt = '" & Format(TempRec2!UPTOSTDT, "yyyy/MM/dd") & "' "
                    mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                    Cnn.Execute mysql
                    If ChkDelItemBrok.Value = False Then
                        If TempRec2!DELETEROW <> "Y" Then
                            If ChkBrokLock.Value = 1 Then
                                LStdTDate = DateValue(vcDTP1.Value)
                                If DateValue(LStdTDate) > DateValue(GSysLockDt) Then
                                    Call PInsert_PitBrok(GCompCode, TempRec2!PARTYCODE, TempRec2!ITEMCODE, Left$(TempRec2!broktype, 1), Val(TempRec2!brokrate & vbNullString), _
                                    Val(TempRec2!STDRATE & vbNullString), Val(TempRec2!TRANRATE & vbNullString), Left$(TempRec2!TRANTYPE, 1), TempRec2!BROKRATE2, Left$(TempRec2!MARTYPE, 1), _
                                    Val(TempRec2!MARRATE & vbNullString), Format(LStdTDate, "yyyy/MM/dd"), 0, 0, Val(TempRec2!MinRate & vbNullString), Left$(TempRec2!MBROKTYPE, 1), _
                                    Val(TempRec2!MBROKRATE & vbNullString), Val(TempRec2!MBROKRATE2 & vbNullString), LInstType, TempRec2!excode, TempRec2!EXID, TempRec2!itemid, TempRec2!ACCID)
                                Else
                                    MsgBox "Sorry System Locked.  No Modification Allowed"
                                    Exit Do
                               End If
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            TempRec2.MoveNext
        Loop
        Set TempRec2 = Nothing
        LSettlementDt = GFinEnd
        If ChkUpdLastSettle.Value = 1 Then
            Call Get_Items
            Set AccRec = Nothing
            Set AccRec = New ADODB.Recordset
            mysql = "SELECT ACCID,AC_CODE FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " "
            mysql = mysql & " AND AC_CODE IN (" & LSParties & ") ORDER BY AC_CODE"
            AccRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            While Not AccRec.EOF
                mysql = "SELECT EXID,ITEMID,EXCHANGECODE,ITEMCODE FROM ITEMMAST  WHERE COMPCODE  = " & GCompCode & " "
                If LenB(LSItems) > 0 Then mysql = mysql & " AND ITEMID   IN (" & LSItems & ")"
                mysql = mysql & "  ORDER BY EXCHANGECODE,ITEMCODE  "
                Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
                Do While Not TRec.EOF
                    DoEvents
                    mysql = "SELECT AC_CODE FROM PITBROK WHERE COMPCODE=" & GCompCode & " AND AC_CODE ='" & AccRec!AC_CODE & "' AND ITEMID =" & TRec!itemid & " AND UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE ='" & LInstType & "'"
                    Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                    If TRec2.EOF Then
                        If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                            Call PInsert_PitBrok(GCompCode, AccRec!AC_CODE, TRec!ITEMCODE, "P", 0, 0, 0, "P", 0, "I", 0, Format(LSettlementDt, "yyyy/MM/dd"), 0, 0, 0, "P", 0, 0, LInstType, TRec!EXCHANGECODE, TRec!EXID, TRec!itemid, AccRec!ACCID)
                        End If
                    End If
                    TRec.MoveNext
                Loop
                AccRec.MoveNext
            Wend
            Set TRec = Nothing
            Set AccRec = Nothing
            Cnn.CommitTrans
            CNNERR = False
        End If
        Fill_ItemGrid
    End If
End If
End Sub

Private Sub OptExchange_Click()
If OptItem.Value = True Then
    ItemList.Visible = True
    ChkUpdLastSettle.Visible = True
Else
    ItemList.Visible = False
    ChkUpdLastSettle.Visible = False
End If
    Set UptoStDtRec = Nothing: Set UptoStDtRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PEXBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT"
    UptoStDtRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not UptoStDtRec.EOF Then
        Set SettleDCombo.RowSource = UptoStDtRec
        SettleDCombo.ListField = "CONDATE"
        SettleDCombo.BoundColumn = "CONDATE"
    End If
    SSTab1.Tab = 0
End Sub
Private Sub OptItem_Click()
If OptItem.Value = True Then
    ItemList.Visible = True
    ChkUpdLastSettle.Visible = True
    Set UptoStDtRec = Nothing: Set UptoStDtRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PEXBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT"
    UptoStDtRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not UptoStDtRec.EOF Then
        Set SettleDCombo.RowSource = UptoStDtRec
        SettleDCombo.ListField = "CONDATE"
        SettleDCombo.BoundColumn = "CONDATE"
    End If
Else
    ItemList.Visible = False
    ChkUpdLastSettle.Visible = False
    Set UptoStDtRec = Nothing: Set UptoStDtRec = New ADODB.Recordset
    mysql = "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PITBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT"
    UptoStDtRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not UptoStDtRec.EOF Then
        Set SettleDCombo.RowSource = UptoStDtRec
        SettleDCombo.ListField = "CONDATE"
        SettleDCombo.BoundColumn = "CONDATE"
    End If
End If
SSTab1.Tab = 1
End Sub
Private Sub Partychk_Click()
Dim I As Integer
    For I = 1 To PartyLst.ListItems.Count
        If PartyChk.Value = 1 Then
            PartyLst.ListItems.Item(I).Checked = True
        Else
            PartyLst.ListItems.Item(I).Checked = False
        End If
    Next I
End Sub
Private Sub ExchangeChk_Click()
Dim I As Integer
For I = 1 To ExchnageList.ListItems.Count
    If ExchangeChk.Value = 1 Then
        ExchnageList.ListItems.Item(I).Checked = True
    Else
        ExchnageList.ListItems.Item(I).Checked = False
    End If
Next I
Call ExchnageList_Click
End Sub
Private Sub MarginCombo_GotFocus()
    Select Case Left$(RECGRID!MARTYPE, 1)
    Case "Q"
        MarginCombo.ListIndex = Val(0)
    Case "V"
        MarginCombo.ListIndex = Val(1)
    Case "I"
        MarginCombo.ListIndex = Val(2)
    Case "C"
        MarginCombo.ListIndex = Val(3)
    Case "L"
        MarginCombo.ListIndex = Val(4)
    Case "N"
        MarginCombo.ListIndex = Val(5)
    End Select
    MarginCombo.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    MarginCombo.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width) + 400
    MarginCombo.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub MarginCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If KeyCode = 13 Then RECGRID!MARTYPE = MarginCombo.text
        MarginCombo.Visible = False: DataGrid1.Col = 16: DataGrid1.SetFocus
    ElseIf KeyCode = 27 Then
        MarginCombo.Visible = False
    End If
End Sub
Private Sub MarginCombo_Validate(Cancel As Boolean)
    If Len(Trim(MarginCombo.text)) < 1 Then
        Cancel = True: Exit Sub
    Else
        RECGRID!MARTYPE = MarginCombo.text
    End If
End Sub
Private Sub InstTypeCombo_Validate(Cancel As Boolean)
If InstTypeCombo.ListIndex < 0 Then
    MsgBox "Please Select Instrument Type "
    Cancel = True
End If
End Sub
Public Sub OkCmd_Click()
    Dim ChkCount As Integer:    Dim J As Integer
    LSParties = vbNullString: LSItems = vbNullString
    Call Check_PExBrok
    ChkCount = 0
    Frame3.Enabled = False
    'Frame3.BackColor = &H8080FF
    For J = 1 To PartyLst.ListItems.Count
        If PartyLst.ListItems(J).Checked = True Then
            ChkCount = ChkCount + 1
            If LenB(LSParties) > 0 Then LSParties = LSParties & ", "
            LSParties = LSParties & "'" & PartyLst.ListItems(J).SubItems(1) & "'"
        End If
    Next
    Call Get_ExCodes
    If LenB(LSParties) = 0 Then
        Frame3.Enabled = True
        'Frame3.BackColor = &HFFC0C0
        MsgBox "Please Select Party.", vbCritical:
        PartyLst.SetFocus:
        Exit Sub
    End If
    ChkCount = 0
    If LenB(LExCodes) = 0 Then
        MsgBox "Please Select Commodity/Script ", vbCritical:
        Frame3.Enabled = True
        'Frame3.BackColor = &HFFC0C0
        ExchangeChk.Enabled = True:     ExchnageList.SetFocus
        Exit Sub
    End If
    
    CountRow = -1
    If InstTypeCombo.ListIndex = 0 Then
        LInstType = "FUT"
    ElseIf InstTypeCombo.ListIndex = 1 Then
        LInstType = "OPT"
    ElseIf InstTypeCombo.ListIndex = 2 Then
        LInstType = "CSH"
    End If
    If LenB(SettleDCombo.BoundText) > 1 Then ChkBrokLock.Visible = True
    Frame4.Enabled = True
    Frame4.BackColor = &HFFC0C0
    Call Fill_ExchangeGrid
    Call Fill_ItemGrid
    
End Sub

Private Sub SettleDCombo_GotFocus()
    Sendkeys "%{DOWN}"
End Sub
Private Sub SettleDCombo_Validate(Cancel As Boolean)
    If IsDate(SettleDCombo.text) Then
        If SYSTEMLOCK(DateValue(SettleDCombo.text)) Then
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
    Dim TRec As ADODB.Recordset:      Dim AccRecADO As ADODB.Recordset
    Dim ExRec As ADODB.Recordset:    Dim ListIt As ListItem
    FlagBrok = False
    SSTab1.Tab = 0
    InstTypeCombo.ListIndex = 0
    LSettlementDt = vbNullString: Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open "SELECT MAX(SETDATE) AS MAXSETTLEDATE FROM SETTLE WHERE COMPCODE = " & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
    If Not TRec.EOF Then LSettlementDt = TRec!MaxSettleDate
    Set TRec = Nothing
    PartyLst.Enabled = True
    Set AccRecADO = Nothing: Set AccRecADO = New ADODB.Recordset
    mysql = "SELECT DISTINCT ACC.ACCID,ACC.AC_CODE,ACC.NAME FROM ACCOUNTD AS ACC,PEXBROK AS PB WHERE ACC.COMPCODE =" & GCompCode & " "
    mysql = mysql & " AND ACC.ACCID =PB.ACCID ORDER BY ACC.NAME"
    AccRecADO.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not AccRecADO.EOF Then
        While Not AccRecADO.EOF
            Set ListIt = PartyLst.ListItems.Add(, , AccRecADO!NAME)
            ListIt.SubItems(1) = AccRecADO!AC_CODE
            ListIt.SubItems(2) = AccRecADO!ACCID
            AccRecADO.MoveNext
        Wend
        Call Get_Selection(13)
        
        'branch list
        Set TRec = Nothing: Set TRec = New ADODB.Recordset
        TRec.Open "SELECT FMLYID,FMLYCODE,FMLYNAME FROM ACCFMLY WHERE COMPCODE =" & GCompCode & " ORDER BY FMLYNAME ", Cnn, adOpenKeyset, adLockReadOnly
        If Not TRec.EOF Then
            While Not TRec.EOF
                Set ListIt = BranchList.ListItems.Add(, , TRec!FmlyNAME)
                ListIt.SubItems(1) = TRec!FMLYCODE
                ListIt.SubItems(2) = TRec!FMLYID
                TRec.MoveNext
            Wend
        End If
        BranchList.Enabled = False: Check4.Enabled = False
        
        Set ExRec = Nothing: Set ExRec = New ADODB.Recordset
        mysql = "SELECT EXID,EXCODE,EXNAME FROM EXMAST WHERE COMPCODE =" & GCompCode & " ORDER BY EXCODE "
        ExRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
        If Not ExRec.EOF Then
            ExRec.MoveFirst
            ExchnageList.ListItems.Clear
            Do While Not ExRec.EOF
                If (ExRec!excode = "EQ" Or ExRec!excode = "BEQ") Then InstTypeCombo.Visible = True
                ExchnageList.ListItems.Add , , ExRec!excode
                ExchnageList.ListItems(ExchnageList.ListItems.Count).ListSubItems.Add , , ExRec!EXNAME
                ExchnageList.ListItems(ExchnageList.ListItems.Count).ListSubItems.Add , , ExRec!EXID
                ExRec.MoveNext
            Loop
            
            If ExRec.RecordCount = 1 Then
                 ExchnageList.TabStop = False:
                 ExchangeChk.Value = 1
                 Call ExchangeChk_Click
                
            Else
                ExchnageList.TabStop = True:
            End If
        Else
            ExchnageList.Enabled = False:
        End If
        
        Set ExRec = Nothing
        Set UptoStDtRec = Nothing: Set UptoStDtRec = New ADODB.Recordset
        UptoStDtRec.Open "SELECT DISTINCT UPTOSTDT AS CONDATE FROM PEXBROK WHERE COMPCODE =" & GCompCode & " ORDER BY UPTOSTDT", Cnn, adOpenKeyset, adLockReadOnly
        If Not UptoStDtRec.EOF Then
            Set SettleDCombo.RowSource = UptoStDtRec
            SettleDCombo.ListField = "CONDATE"
            SettleDCombo.BoundColumn = "CONDATE"
        End If
    End If
    Set AccRecADO = Nothing
    If GValueWiseYN = "N" Then
        DataGrid1.Columns(6).Visible = False
        DataGrid2.Columns(6).Visible = False
    End If
    
    If GMinBrokYN = "N" Then
        DataGrid1.Columns(12).Visible = False
        DataGrid1.Columns(13).Visible = False
        DataGrid1.Columns(14).Visible = False
        DataGrid1.Columns(15).Visible = False
        
        DataGrid2.Columns(12).Visible = False
        DataGrid2.Columns(13).Visible = False
        DataGrid2.Columns(14).Visible = False
        DataGrid2.Columns(15).Visible = False
        
    End If
    If GOnlyBrok = "0" Then
        DataGrid1.Columns(5).Visible = True
        'DataGrid1.Columns(6).Visible = True
        DataGrid1.Columns(7).Visible = True
        DataGrid1.Columns(5).Caption = "StampRate"
        DataGrid1.Columns(7).Caption = "SEBITax"
        
    End If
    If GStandingYN = "N" Then
        DataGrid1.Columns(8).Visible = False
        DataGrid2.Columns(8).Visible = False
    End If
    If GTranFeesYN = "N" Then
        DataGrid1.Columns(9).Visible = False
        DataGrid1.Columns(10).Visible = False
        
        DataGrid2.Columns(9).Visible = False
        DataGrid2.Columns(10).Visible = False
        
    End If
    If GMarginYN = "N" Then
        DataGrid1.Columns(16).Visible = False
        DataGrid1.Columns(17).Visible = False
        
        DataGrid2.Columns(16).Visible = False
        DataGrid2.Columns(17).Visible = False
    End If
    Call CANCEL_REC
End Sub
Sub RecSet2()
    Set RecGrid2 = Nothing
    Set RecGrid2 = New ADODB.Recordset
    RecGrid2.Fields.Append "EXCODE", adVarChar, 6, adFldIsNullable
    RecGrid2.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
    RecGrid2.Fields.Append "INSTTYPE", adVarChar, 3, adFldIsNullable
    RecGrid2.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    RecGrid2.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "STDRATE", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "TRANRATE", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "TRANTYPE", adVarChar, 50, adFldIsNullable
    RecGrid2.Fields.Append "UPTOSTDT", adDate, , adFldIsNullable
    RecGrid2.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "MARTYPE", adVarChar, 50, adFldIsNullable
    RecGrid2.Fields.Append "MARRATE", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "MBROKTYPE", adVarChar, 50, adFldIsNullable
    RecGrid2.Fields.Append "MBROKRATE", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "MBROKRATE2", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "MINRATE", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "New", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "PARTY", adVarChar, 150, adFldIsNullable
    RecGrid2.Fields.Append "PARTYCODE", adVarChar, 15, adFldIsNullable
    RecGrid2.Fields.Append "DeleteRow", adVarChar, 1, adFldIsNullable
    RecGrid2.Fields.Append "SELLBROK", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "SELLBROK2", adDouble, , adFldIsNullable
    RecGrid2.Fields.Append "ITEMID", adInteger, , adFldIsNullable
    RecGrid2.Fields.Append "EXID", adInteger, , adFldIsNullable
    RecGrid2.Fields.Append "ACCID", adInteger, , adFldIsNullable
    
    RecGrid2.Open , , adOpenKeyset, adLockOptimistic
    
    

End Sub
Sub RecSet()
    Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "EXCODE", adVarChar, 6, adFldIsNullable
    RECGRID.Fields.Append "INSTTYPE", adVarChar, 3, adFldIsNullable
    RECGRID.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "STDRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "TRANRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "TRANTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "UPTOSTDT", adDate, , adFldIsNullable
    RECGRID.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MARTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "MARRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MBROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "MBROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MBROKRATE2", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MINRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "New", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "PARTY", adVarChar, 150, adFldIsNullable
    RECGRID.Fields.Append "PARTYCODE", adVarChar, 15, adFldIsNullable
    RECGRID.Fields.Append "DeleteRow", adVarChar, 1, adFldIsNullable
    RECGRID.Fields.Append "StampRate", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "SEBITAX", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "EXID", adInteger, , adFldIsNullable
    RECGRID.Fields.Append "ACCID", adInteger, , adFldIsNullable
    
    RECGRID.Open , , adOpenKeyset, adLockOptimistic
    
    
End Sub
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LGridRow As Integer
    Dim LGridCol  As Integer
    If KeyCode = 13 And DataGrid1.Col = 3 Then ' BROKTYPE
        BrokTypeCombo.Visible = True: BrokTypeCombo.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 10 Then 'TRANTYPE
        BrokTypeCombo.Visible = True: BrokTypeCombo.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 13 Then 'MBROKTYPE
        BrokTypeCombo.Visible = True: BrokTypeCombo.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 16 Then 'MARGIN TYPE
        MarginCombo.Visible = True: MarginCombo.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 18 Then
        RECGRID.MoveNext ''ADDING NEW ROW
        If RECGRID.EOF Then
            RECGRID.AddNew
            RECGRID.Fields("ITEMCODE") = vbNullString:      RECGRID.Fields("ITEMNAME") = vbNullString
            RECGRID.Fields("BROKTYPE") = vbNullString:      RECGRID.Fields("BROKRATE") = 0
            RECGRID.Fields("BBROKRATE") = 0:                RECGRID.Fields("STMRATE") = 0
            RECGRID.Fields("STTRATE") = 0:                  RECGRID.Fields("STDRATE") = 0
            RECGRID.Fields("TRANRATE") = 0:                 RECGRID.Fields("TranType") = vbNullString
            RECGRID.Fields("BROKRATE2") = 0:                RECGRID.Fields("EXCODE") = vbNullString
            RECGRID.Fields("MARTYPE") = vbNullString:       RECGRID.Fields("MARRATE") = 0
            RECGRID!EXID = 0
            RECGRID.Fields("UPTOSTDT") = Format(GFinEnd, "YYYY/MM/DD")
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            RECGRID.Update
         End If
        DataGrid1.LeftCol = 0: DataGrid1.Col = 0
    ElseIf KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
    'Press F7 to replace all rows with current cell value
    If KeyCode = 118 Then   'F7
        LGridRow = DataGrid1.Row
        LGridCol = DataGrid1.Col
        If DataGrid1.Col = 3 Then 'BROKTYPE
            GridColVal = RECGRID!broktype
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!broktype = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 4 Then 'BROKRATE
            GridColVal = RECGRID!brokrate
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!brokrate = Val(GridColVal)
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 6 Then 'brokrate2
            GridColVal = RECGRID!BROKRATE2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!BROKRATE2 = Val(GridColVal)
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 5 Then 'stamprate
            GridColVal = RECGRID!STAMPRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!STAMPRATE = Val(GridColVal)
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 7 Then 'SEBITAX
            GridColVal = RECGRID!SEBITAX
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!SEBITAX = Val(GridColVal)
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 8 Then 'STDRATE
            GridColVal = RECGRID!STDRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!STDRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 9 Then 'TRANRATE
            GridColVal = RECGRID!TRANRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!TRANRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 10 Then 'TRANTYPE
            GridColVal = IIf(IsNull(RECGRID!TRANTYPE), "P", RECGRID!TRANTYPE)
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!TRANTYPE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 12 Then 'MINRATE
            GridColVal = RECGRID!MinRate
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MinRate = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 13 Then 'MBROKTYPE
            GridColVal = RECGRID!MBROKTYPE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MBROKTYPE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 14 Then 'MBROKTRATE
            GridColVal = RECGRID!MBROKRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MBROKRATE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 15 Then 'MBROKTRATE
            GridColVal = RECGRID!MBROKRATE2
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MBROKRATE2 = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 16 Then 'MARTYPE
            GridColVal = RECGRID!MARTYPE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MARTYPE = GridColVal
                RECGRID.MoveNext
            Wend
        ElseIf DataGrid1.Col = 17 Then 'MARRATE
            GridColVal = RECGRID!MARRATE
            RECGRID.MoveFirst
            While Not RECGRID.EOF
                RECGRID!MARRATE = GridColVal
                RECGRID.MoveNext
            Wend
        End If
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol: DataGrid1.SetFocus
    End If
End Sub
Private Sub BrokTypeCombo_GotFocus()
    If DataGrid1.Col = 2 Then
        Select Case Left$(RECGRID!broktype, 1)
        Case "T"
            BrokTypeCombo.ListIndex = 0
        Case "O"
            BrokTypeCombo.ListIndex = 1
        Case "P"
            BrokTypeCombo.ListIndex = 2
        Case "I"
            BrokTypeCombo.ListIndex = 3
        Case "C"
            BrokTypeCombo.ListIndex = 4
        Case "V"
            BrokTypeCombo.ListIndex = 5
        Case "Q"
            BrokTypeCombo.ListIndex = 6
        Case "D"
            BrokTypeCombo.ListIndex = 7
        Case "H"
            BrokTypeCombo.ListIndex = 8
        Case "L"
            BrokTypeCombo.ListIndex = 9
        Case "W"
            BrokTypeCombo.ListIndex = 10
        Case "X"
            BrokTypeCombo.ListIndex = 11
        Case "Z"
            BrokTypeCombo.ListIndex = 12
        Case "R"
            BrokTypeCombo.ListIndex = 13
        Case "F"
            BrokTypeCombo.ListIndex = 14
        Case "M"
            BrokTypeCombo.ListIndex = 15
        Case "B"
            BrokTypeCombo.ListIndex = 16
        Case "N"
            BrokTypeCombo.ListIndex = 17
        Case "U"
            BrokTypeCombo.ListIndex = 18
        Case "Y"
            BrokTypeCombo.ListIndex = 19
        Case "S"
            BrokTypeCombo.ListIndex = 20
        Case "A"
            BrokTypeCombo.ListIndex = 21
        Case "1"
            BrokTypeCombo.ListIndex = 22
        Case "2"
            BrokTypeCombo.ListIndex = 23
        Case "3"
            BrokTypeCombo.ListIndex = 24
        Case "4"
            BrokTypeCombo.ListIndex = 25
        Case "5"
            BrokTypeCombo.ListIndex = 26
        End Select
    ElseIf DataGrid1.Col = 6 Then
            If Mid(RECGRID!TRANTYPE, 1, 1) = "T" Then
                BrokTypeCombo.ListIndex = Val(0)
            ElseIf Mid(RECGRID!TRANTYPE, 1, 1) = "P" Then
                BrokTypeCombo.ListIndex = Val(1)
            End If
    ElseIf DataGrid1.Col = 13 Then
        Select Case Left$(RECGRID!MBROKTYPE, 1)
        Case "T"
            BrokTypeCombo.ListIndex = 0
        Case "O"
            BrokTypeCombo.ListIndex = 1
        Case "P"
            BrokTypeCombo.ListIndex = 2
        Case "I"
            BrokTypeCombo.ListIndex = 3
        Case "C"
            BrokTypeCombo.ListIndex = 4
        Case "V"
            BrokTypeCombo.ListIndex = 5
        Case "Q"
            BrokTypeCombo.ListIndex = 6
        Case "D"
            BrokTypeCombo.ListIndex = 7
        Case "H"
            BrokTypeCombo.ListIndex = 8
        Case "L"
            BrokTypeCombo.ListIndex = 9
        Case "W"
            BrokTypeCombo.ListIndex = 10
        Case "X"
            BrokTypeCombo.ListIndex = 11
        Case "Z"
            BrokTypeCombo.ListIndex = 12
        Case "R"
            BrokTypeCombo.ListIndex = 13
        Case "F"
            BrokTypeCombo.ListIndex = 14
        Case "M"
            BrokTypeCombo.ListIndex = 15
        Case "B"
            BrokTypeCombo.ListIndex = 16
        Case "N"
            BrokTypeCombo.ListIndex = 17
        Case "U"
            BrokTypeCombo.ListIndex = 18
        Case "Y"
            BrokTypeCombo.ListIndex = 19
        Case "S"
            BrokTypeCombo.ListIndex = 20
        Case "A"
            BrokTypeCombo.ListIndex = 21
        Case "1"
            BrokTypeCombo.ListIndex = 22
        Case "2"
            BrokTypeCombo.ListIndex = 23
        Case "3"
            BrokTypeCombo.ListIndex = 24
        Case "4"
            BrokTypeCombo.ListIndex = 25
        Case "5"
            BrokTypeCombo.ListIndex = 26
        End Select
    End If
    BrokTypeCombo.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    BrokTypeCombo.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    BrokTypeCombo.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub BrokTypeCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LGridRow As Integer
    Dim LGridCol  As Integer
    Dim SearchRow As Integer
    If KeyCode = 13 Then
        LGridRow = DataGrid1.Row: LGridCol = DataGrid1.Col:  SearchRow = RECGRID!New
        If DataGrid1.Col = 3 Then
            If KeyCode = 13 Then RECGRID!broktype = BrokTypeCombo.text
        ElseIf DataGrid1.Col = 8 Then
            If KeyCode = 13 Then RECGRID!TRANTYPE = BrokTypeCombo.text
        ElseIf DataGrid1.Col = 11 Then
            If KeyCode = 13 Then RECGRID!MBROKTYPE = BrokTypeCombo.text
        ElseIf DataGrid1.Col = 13 Then
            If KeyCode = 13 Then RECGRID!MBROKTYPE = BrokTypeCombo.text
        End If
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        RECGRID.MoveFirst: DataGrid1.SetFocus
        RECGRID.Find "new =" & SearchRow & "", , adSearchForward
        DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol + 1: BrokTypeCombo.Visible = False: DataGrid1.SetFocus
    ElseIf KeyCode = 27 Then
        BrokTypeCombo.Visible = False
    End If
    
End Sub
Private Sub BrokTypeCombo_Validate(Cancel As Boolean)
    If Len(Trim(BrokTypeCombo.text)) < 1 Then
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub ExchnageList_Click()
    Call Get_ExCodes
    Call Fill_ItemList
End Sub
Private Sub Fill_ExchangeGrid()
    Dim BrokRec As ADODB.Recordset:     Dim LBrokType As String * 1:        Dim LMBrokType As String * 1
    Call RecSet
    Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    Set BrokRec = Nothing: Set BrokRec = New ADODB.Recordset
    mysql = "SELECT AM.ACCID,AM.NAME,AM.AC_CODE,A.TRANTYPE,A.BROKTYPE,A.BROKRATE,A.BROKRATE2,A.STDRATE,A.TRANRATE,"
    mysql = mysql & " A.UPTOSTDT,A.MARTYPE,A.MARRATE,A.MINRATE,A.MBROKRATE,A.MBROKRATE2,A.MBROKTYPE,B.EXCODE,A.INSTTYPE,A.SELLBROK,A.SELLBROK2,B.EXID "
    mysql = mysql & " FROM PEXBROK AS A, ACCOUNTD AS AM ,EXMAST AS B WHERE A.COMPCODE=" & GCompCode & " AND A.COMPCODE=B.COMPCODE "
    mysql = mysql & " AND A.EXID =B.EXID AND A.COMPCODE = AM.COMPCODE AND A.AC_CODE=AM.AC_CODE "
    mysql = mysql & " AND A.INSTTYPE='" & LInstType & "'"
    mysql = mysql & " AND A.AC_CODE IN (" & LSParties & ")"
    mysql = mysql & " AND B.EXID IN (" & LExCodes & ") "
    If ChkBrokLock.Value = 0 Then
        If IsDate(SettleDCombo.text) Then mysql = mysql & " AND A.UPTOSTDT = '" & Format(SettleDCombo.text, "yyyy/MM/dd") & "'"
    End If
    mysql = mysql & " ORDER BY AM.NAME,B.EXCODE,A.UPTOSTDT "
    BrokRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
    If Not BrokRec.EOF Then
        Do While Not BrokRec.EOF
            DoEvents
            RECGRID.AddNew
            RECGRID.Fields("EXCODE") = BrokRec!excode
            RECGRID.Fields("EXID") = BrokRec!EXID
            RECGRID.Fields("INSTTYPE") = BrokRec!INSTTYPE
            If IsNull(BrokRec!broktype) Or BrokRec!broktype = "" Then
                RECGRID.Fields("BROKTYPE") = "Transaction"
            Else
                LBrokType = BrokRec!broktype
                Select Case LBrokType
                    Case "A"
                        RECGRID.Fields("BROKTYPE") = "A Opening ZLotwise"
                    Case "B"
                        RECGRID.Fields("BROKTYPE") = "BuySell Intraday"
                    Case "C"
                        RECGRID.Fields("BROKTYPE") = "Closing Sauda"
                    Case "D"
                        RECGRID.Fields("BROKTYPE") = "Delivery Wise Brokerage"
                    Case "F"
                        RECGRID.Fields("BROKTYPE") = "Fixed Brokerage"
                    Case "H"
                        RECGRID.Fields("BROKTYPE") = "Higher Value Percentage Wise"
                    Case "I"
                        RECGRID.Fields("BROKTYPE") = "IntraDay Brokerage"
                    Case "L"
                        RECGRID.Fields("BROKTYPE") = "LotWise Higher Value "
                    Case "M"
                        RECGRID.Fields("BROKTYPE") = "MRate Wise IntraDay"
                    Case "N"
                        RECGRID.Fields("BROKTYPE") = "N Per Trade Wise"
                    Case "O"
                        RECGRID.Fields("BROKTYPE") = "Opening Sauda"
                    Case "P"
                        RECGRID.Fields("BROKTYPE") = "Percentage wise"
                    Case "Q"
                        RECGRID.Fields("BROKTYPE") = "Qtywise IntraDay"
                    Case "S"
                        RECGRID.Fields("BROKTYPE") = "S Qtywise IntraDay"
                    Case "R"
                        RECGRID.Fields("BROKTYPE") = "RZLotwise Intraday"
                    'Case "S"
                    '    RECGRID.Fields("BROKTYPE") = "Slab Wise Brokerage"
                    Case "T"
                        RECGRID.Fields("BROKTYPE") = "Transaction"
                    Case "U"
                        RECGRID.Fields("BROKTYPE") = "U ShareQty Wise"
                    Case "V"
                        RECGRID.Fields("BROKTYPE") = "Valuewise Intraday"
                    Case "X"
                        RECGRID.Fields("BROKTYPE") = "XIntraday Higher Wise"
                    Case "Y"
                        RECGRID.Fields("BROKTYPE") = "Y Qtywise Intraday"
                    Case "Z"
                        RECGRID.Fields("BROKTYPE") = "ZLotwise"
                    Case "1"
                        RECGRID.Fields("BROKTYPE") = "1 RateWise Percentage Wise"
                    Case "2"
                        RECGRID.Fields("BROKTYPE") = "2 MinRate Percentage Wise"
                    Case "3"
                        RECGRID.Fields("BROKTYPE") = "3 Closing Sauda %"
                    Case "4"
                        RECGRID.Fields("BROKTYPE") = "4 Valuewise Intraday 2"
                    Case "5"
                        RECGRID.Fields("BROKTYPE") = "5 Closing Sauda Zlotwise"
                    Case "W"
                        RECGRID.Fields("BROKTYPE") = "WHigher Value Intraday Wise"
                End Select
            End If
            RECGRID!brokrate = Val(BrokRec!brokrate & vbNullString)
            
            RECGRID.Fields("MBROKRATE") = Val(BrokRec!MBROKRATE & vbNullString)
            RECGRID.Fields("MBROKRATE2") = Val(BrokRec!MBROKRATE2 & vbNullString)
            RECGRID.Fields("MINRATE") = Val(BrokRec!MinRate & vbNullString)
            RECGRID.Fields("BROKRATE2") = Val(BrokRec!BROKRATE2) & vbNullString
            RECGRID.Fields("STDRATE") = BrokRec!STDRATE
            RECGRID.Fields("TRANRATE") = BrokRec!TRANRATE
            RECGRID.Fields("STAMPRATE") = Val(BrokRec!SELLBROK & vbNullString)
            RECGRID.Fields("SEBITAX") = Val(BrokRec!SELLBROK2 & vbNullString)
            If IsNull(BrokRec!TRANTYPE) Or BrokRec!TRANTYPE = "" Then
                RECGRID.Fields("TranType") = "Transaction"
            Else
                If BrokRec!TRANTYPE = "T" Then
                    RECGRID.Fields("TranType") = "Transaction"
                ElseIf BrokRec!TRANTYPE = "P" Then
                    RECGRID.Fields("TranType") = "Percentage wise"
                End If
            End If
            If IsNull(BrokRec!UPTOSTDT) Then
                RECGRID.Fields("UPTOSTDT") = Format(LSettlementDt, "YYYY/MM/DD")
            Else
                RECGRID.Fields("UPTOSTDT") = Format(BrokRec!UPTOSTDT, "YYYY/MM/DD")
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
                ElseIf BrokRec!MARTYPE = "L" Then
                    RECGRID.Fields("MARTYPE") = "LotWise Margin"
                Else
                    RECGRID.Fields("MARTYPE") = "Import Rates"
                End If
            End If
            If IsNull(BrokRec!MBROKTYPE) Or BrokRec!MBROKTYPE = "" Then
                RECGRID.Fields("MBROKTYPE") = "Transaction"
            Else
                LMBrokType = BrokRec!MBROKTYPE
                Select Case LMBrokType
                Case "A"
                    RECGRID.Fields("MBROKTYPE") = "A Opening ZLotwise"
                Case "B"
                    RECGRID.Fields("MBROKTYPE") = "BuySell Intraday"
                Case "C"
                    RECGRID.Fields("MBROKTYPE") = "Closing Sauda"
                Case "D"
                    RECGRID.Fields("MBROKTYPE") = "Delivery Wise Brokerage"
                Case "F"
                    RECGRID.Fields("MBROKTYPE") = "Fixed Brokerage"
                Case "H"
                    RECGRID.Fields("MBROKTYPE") = "Higher Value Percentage Wise"
                Case "I"
                    RECGRID.Fields("MBROKTYPE") = "IntraDay Brokerage"
                Case "L"
                    RECGRID.Fields("MBROKTYPE") = "LotWise Higher Value "
                Case "M"
                    RECGRID.Fields("MBROKTYPE") = "MRate Wise IntraDay"
                Case "N"
                    RECGRID.Fields("MBROKTYPE") = "N Per Trade Wise"
                Case "O"
                    RECGRID.Fields("MBROKTYPE") = "Opening Sauda"
                Case "P"
                    RECGRID.Fields("MBROKTYPE") = "Percentage wise"
                Case "Q"
                    RECGRID.Fields("MBROKTYPE") = "Qtywise IntraDay"
                Case "R"
                    RECGRID.Fields("MBROKTYPE") = "RZLotwise Intraday"
                Case "S"
                    RECGRID.Fields("MBROKTYPE") = "S Variable Qtywise IntraDay"
                Case "T"
                    RECGRID.Fields("MBROKTYPE") = "Transaction"
                Case "U"
                    RECGRID.Fields("MBROKTYPE") = "U ShareQty Wise"
                Case "V"
                    RECGRID.Fields("MBROKTYPE") = "Valuewise Intraday"
                Case "X"
                    RECGRID.Fields("MBROKTYPE") = "XIntraday Higher Wise"
                Case "Y"
                    RECGRID.Fields("MBROKTYPE") = "Y Qtywise Intraday"
                Case "Z"
                    RECGRID.Fields("MBROKTYPE") = "ZLotwise"
                Case "1"
                    RECGRID.Fields("MBROKTYPE") = "1 RateWise Percentage Wise"
                Case "2"
                    RECGRID.Fields("MBROKTYPE") = "2 MinRate Percentage Wise"
                Case "3"
                    RECGRID.Fields("MBROKTYPE") = "3 Closing Sauda %"
                Case "4"
                    RECGRID.Fields("MBROKTYPE") = "4 Valuewise Intraday 2"
                Case "5"
                    RECGRID.Fields("MBROKTYPE") = "5 Closing Sauda Zlotwise"
                Case "W"
                    RECGRID.Fields("BROKTYPE") = "WHigher Value Intraday Wise"
                End Select
            End If
            RECGRID.Fields("MARRATE") = Val(BrokRec!MARRATE & vbNullString)
            CountRow = CountRow + 1
            RECGRID.Fields("New") = CountRow
            RECGRID.Fields("PARTY") = BrokRec!NAME
            RECGRID.Fields("PARTYCODE") = BrokRec!AC_CODE
            RECGRID.Fields("ACCID") = BrokRec!ACCID
            
            RECGRID.Fields("DELETEROW") = "N"
            RECGRID.Update
            BrokRec.MoveNext
        Loop
        RECGRID.AddNew
        RECGRID.Fields("EXCODE") = vbNullString:        RECGRID.Fields("INSTTYPE") = vbNullString
        RECGRID.Fields("BROKTYPE") = vbNullString:      RECGRID.Fields("BROKRATE") = 0
        RECGRID.Fields("STAMPRATE") = 0:                RECGRID.Fields("SEBITAX") = 0
        RECGRID.Fields("STDRATE") = 0:                  RECGRID.Fields("TRANRATE") = 0
        RECGRID.Fields("TranType") = vbNullString:      RECGRID.Fields("UPTOSTDT") = Format(GFinEnd, "YYYY/MM/DD")
        RECGRID.Fields("BROKRATE2") = 0:                RECGRID.Fields("MARTYPE") = vbNullString
        RECGRID.Fields("MARRATE") = 0:                  RECGRID.Fields("MBROKTYPE") = vbNullString
        RECGRID.Fields("MBROKRATE") = 0:                RECGRID.Fields("MINRATE") = 0:
        RECGRID.Fields("PARTY") = vbNullString:         RECGRID.Fields("PARTYCODE") = vbNullString:
        RECGRID.Fields("ACCID") = 0
        RECGRID.Fields("EXID") = 0
        CountRow = CountRow + 1
        RECGRID.Fields("New") = CountRow
        RECGRID.Update
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh:
        If RECGRID.RecordCount > 0 Then
            RECGRID.MoveFirst:
            DataGrid1.SetFocus
            DataGrid1.LeftCol = 0
        End If
        
    Else
        MsgBox "Record does not exists.", vbExclamation
        
        Call CANCEL_REC
    End If
    Set BrokRec = Nothing
    
End Sub

Private Sub Fill_ItemGrid()
    Dim BrokRec As ADODB.Recordset:     Dim PartyRec As ADODB.Recordset
    Dim TRec As ADODB.Recordset:        Dim LBrokType As String * 1:    Dim LMBrokType As String * 1
    Call Get_Items
    Call RecSet2
    'If LenB(LlistParties) > 0 Or LenB(LListSaudas) > 0 Then
    If LenB(LSItems) > 1 Or OptItem.Value = False Then
        Set DataGrid2.DataSource = RecGrid2: DataGrid2.ReBind: DataGrid2.Refresh
        Set BrokRec = Nothing: Set BrokRec = New ADODB.Recordset
        If Len(LSItems) < 1 Then
            mysql = "SELECT AM.ACCID,AM.NAME,AM.AC_CODE,A.TRANTYPE,A.BROKTYPE,A.BROKRATE,A.BROKRATE2,A.STDRATE,A.TRANRATE,"
            mysql = mysql & " A.UPTOSTDT , A.MARTYPE,A.MARRATE,A.MINRATE,A.MBROKRATE,A.MBROKRATE2,A.MBROKTYPE,B.EXCHANGECODE,A.ITEMCODE,A.INSTTYPE,B.ITEMID,B.EXID"
            mysql = mysql & " FROM PITBROK AS A, ACCOUNTD AS AM ,ITEMMAST AS B WHERE A.COMPCODE=" & GCompCode & " AND A.COMPCODE=B.COMPCODE "
            mysql = mysql & "  AND A.ITEMID =B.ITEMID AND A.COMPCODE = AM.COMPCODE AND A.AC_CODE=AM.AC_CODE "
            mysql = mysql & " AND A.INSTTYPE='" & LInstType & "' "
            If LenB(LSItems) > 0 Then mysql = mysql & " AND B.ITEMID IN (" & LSItems & ")"
            mysql = mysql & " AND A.AC_CODE IN (" & LSParties & ")"
            mysql = mysql & " AND B.EXID IN (" & LExCodes & ") "
            If ChkBrokLock.Value = 0 Then
                If IsDate(SettleDCombo.text) Then mysql = mysql & " AND A.UPTOSTDT = '" & Format(SettleDCombo.text, "yyyy/MM/dd") & "'"
            End If
            mysql = mysql & " ORDER BY AM.NAME,B.EXCHANGECODE, A.ITEMCODE,A.UPTOSTDT "
        Else
            mysql = "SELECT AM.ACCID,AM.NAME,AM.AC_CODE,A.tRANTYPE,A.BROKTYPE,A.BROKRATE,A.BROKRATE2,A.STDRATE,A.TRANRATE, A.UPTOSTDT , A.MARTYPE,A.MARRATE,A.MINRATE,"
            mysql = mysql & " A.MBROKRATE , A.MBROKRATE2, A.MBROKTYPE, B.EXCHANGECODE, B.ITEMCODE, A.INSTTYPE, B.ITEMID, B.EXID "
            mysql = mysql & " FROM ACCOUNTD AM INNER JOIN ITEMMAST B ON AM.COMPCODE=B.COMPCODE "
            mysql = mysql & " LEFT OUTER JOIN PITBROK AS A ON B.ITEMID=A.ITEMID  AND AM.ACCID = A.ACCID "
            mysql = mysql & " WHERE AM.COMPCODE= " & GCompCode & " and B.EXID IN (" & LExCodes & ") and AM.AC_CODE IN (" & LSParties & ")"
            mysql = mysql & " AND A.INSTTYPE='" & LInstType & "' "
            If LenB(LSItems) > 0 Then mysql = mysql & " AND B.ITEMID IN (" & LSItems & ")"
            If ChkBrokLock.Value = 0 Then
                If IsDate(SettleDCombo.text) Then mysql = mysql & " AND A.UPTOSTDT = '" & Format(SettleDCombo.text, "yyyy/MM/dd") & "'"
            End If
            mysql = mysql & " ORDER BY AM.NAME,B.EXCHANGECODE, A.ITEMCODE,A.UPTOSTDT "
        End If
        
        BrokRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
        If Not BrokRec.EOF Then
            DataGrid2.Enabled = True
            Do While Not BrokRec.EOF
                DoEvents
                RecGrid2.AddNew
                RecGrid2.Fields("EXCODE") = BrokRec!EXCHANGECODE
                RecGrid2.Fields("ITEMCODE") = BrokRec!ITEMCODE
                RecGrid2.Fields("EXID") = BrokRec!EXID
                RecGrid2.Fields("ITEMID") = BrokRec!itemid
                RecGrid2.Fields("INSTTYPE") = IIf(IsNull(BrokRec!INSTTYPE) = True, LInstType, BrokRec!INSTTYPE)
                If IsNull(BrokRec!broktype) Or BrokRec!broktype = "" Then
                    RecGrid2.Fields("BROKTYPE") = "Transaction"
                Else
                    LBrokType = BrokRec!broktype
                    Select Case LBrokType
                        Case "A"
                            RecGrid2.Fields("BROKTYPE") = "A Opening ZLotwise"
                        Case "T"
                            RecGrid2.Fields("BROKTYPE") = "Transaction"
                        Case "O"
                            RecGrid2.Fields("BROKTYPE") = "Opening Sauda"
                        Case "C"
                            RecGrid2.Fields("BROKTYPE") = "Closing Sauda"
                        Case "Q"
                            RecGrid2.Fields("BROKTYPE") = "Qtywise IntraDay"
                        Case "S"
                            RecGrid2.Fields("BROKTYPE") = "S Variable Qtywise IntraDay"
                        Case "P"
                            RecGrid2.Fields("BROKTYPE") = "Percentage wise"
                        Case "B"
                            RecGrid2.Fields("BROKTYPE") = "BuySell Intraday"
                        'Case "S"
                        '    RecGrid2.Fields("BROKTYPE") = "Sold"
                        Case "I"
                            RecGrid2.Fields("BROKTYPE") = "IntraDay Brokerage"
                        Case "V"
                            RecGrid2.Fields("BROKTYPE") = "Valuewise Intraday"
                        Case "D"
                            RecGrid2.Fields("BROKTYPE") = "Delivery Wise Brokerage"
                        Case "H"
                            RecGrid2.Fields("BROKTYPE") = "Higher Value Percentage Wise"
                        Case "L"
                            RecGrid2.Fields("BROKTYPE") = "LotWise"
                        Case "X"
                            RecGrid2.Fields("BROKTYPE") = "XIntraday Higher Wise"
                        Case "Z"
                            RecGrid2.Fields("BROKTYPE") = "ZLotwise"
                        Case "R"
                            RecGrid2.Fields("BROKTYPE") = "RZLotwise Intraday"
                        Case "F"
                            RecGrid2.Fields("BROKTYPE") = "Fixed Brokerage"
                        Case "M"
                            RecGrid2.Fields("BROKTYPE") = "MRate Wise IntraDay"
                        Case "N"
                            RecGrid2.Fields("BROKTYPE") = "N Per Trade Wise"
                        Case "U"
                            RecGrid2.Fields("BROKTYPE") = "U ShareQty Wise"
                        Case "Y"
                            RecGrid2.Fields("BROKTYPE") = "Y Qtywise IntraDay"
                        Case "1"
                            RecGrid2.Fields("BROKTYPE") = "1 RateWise Percentage Wise"
                        Case "2"
                            RecGrid2.Fields("BROKTYPE") = "2 MinRate Percentage Wise"
                        Case "3"
                            RecGrid2.Fields("BROKTYPE") = "3 Closing Sauda %"
                        Case "4"
                            RecGrid2.Fields("BROKTYPE") = "4 Valuewise Intraday 2"
                        Case "5"
                            RecGrid2.Fields("BROKTYPE") = "5 Closing Sauda Zlotwise"
                    End Select
                End If
                    RecGrid2!brokrate = IIf(IsNull(BrokRec!brokrate), 0, Val(BrokRec!brokrate & ""))
                    RecGrid2.Fields("MBROKRATE") = IIf(IsNull(BrokRec!MBROKRATE), 0, BrokRec!MBROKRATE)
                    RecGrid2.Fields("MBROKRATE2") = IIf(IsNull(BrokRec!MBROKRATE2), 0, BrokRec!MBROKRATE2)
                    RecGrid2.Fields("MINRATE") = IIf(IsNull(BrokRec!MinRate), 0, BrokRec!MinRate)
                    RecGrid2.Fields("BROKRATE2") = IIf(IsNull(BrokRec!BROKRATE2), 0, BrokRec!BROKRATE2)
                    RecGrid2.Fields("STDRATE") = IIf(IsNull(BrokRec!STDRATE), 0, BrokRec!STDRATE)
                    RecGrid2.Fields("TRANRATE") = IIf(IsNull(BrokRec!TRANRATE), 0, BrokRec!TRANRATE)
                    If IsNull(BrokRec!TRANTYPE) Or BrokRec!TRANTYPE = "" Then
                        RecGrid2.Fields("TranType") = "Transaction"
                    Else
                        If BrokRec!TRANTYPE = "T" Then
                            RecGrid2.Fields("TranType") = "Transaction"
                        ElseIf BrokRec!TRANTYPE = "P" Then
                            RecGrid2.Fields("TranType") = "Percentage wise"
                        End If
                    End If
                    If IsNull(BrokRec!UPTOSTDT) Then
                        RecGrid2.Fields("UPTOSTDT") = Format(LSettlementDt, "YYYY/MM/DD")
                    Else
                        RecGrid2.Fields("UPTOSTDT") = Format(BrokRec!UPTOSTDT, "YYYY/MM/DD")
                    End If
                    If IsNull(BrokRec!MARTYPE) Or BrokRec!MARTYPE = "" Then
                    RecGrid2.Fields("MARTYPE") = "Value Wise (In %)"
                Else
                    If BrokRec!MARTYPE = "Q" Then
                        RecGrid2.Fields("MARTYPE") = "Qtywise (Per Unit)"
                    ElseIf BrokRec!MARTYPE = "V" Then
                        RecGrid2.Fields("MARTYPE") = "Value Wise (In %)"
                    ElseIf BrokRec!MARTYPE = "I" Then
                        RecGrid2.Fields("MARTYPE") = "Import Rates"
                    ElseIf BrokRec!MARTYPE = "C" Then
                        RecGrid2.Fields("MARTYPE") = "Client Wise Margin"
                    ElseIf BrokRec!MARTYPE = "L" Then
                        RecGrid2.Fields("MARTYPE") = "LotWise Margin"
                    Else
                        RecGrid2.Fields("MARTYPE") = "Import Rates"
                    End If
                End If
                If IsNull(BrokRec!MBROKTYPE) Or BrokRec!MBROKTYPE = "" Then
                    RecGrid2.Fields("MBROKTYPE") = "Transaction"
                Else
                    LMBrokType = BrokRec!MBROKTYPE
                    Select Case LMBrokType
                        Case "A"
                            RECGRID.Fields("MBROKTYPE") = "A Opening ZLotwise"
                    
                        Case "T"
                            RecGrid2.Fields("MBROKTYPE") = "Transaction"
                        Case "O"
                            RecGrid2.Fields("MBROKTYPE") = "Opening Sauda"
                        Case "C"
                            RecGrid2.Fields("MBROKTYPE") = "Closing Sauda"
                        Case "Q"
                            RecGrid2.Fields("MBROKTYPE") = "Qtywise IntraDay"
                        Case "P"
                            RecGrid2.Fields("MBROKTYPE") = "Percentage wise"
                        Case "B"
                            RecGrid2.Fields("MBROKTYPE") = "Buysell Intraday"
                        Case "S"
                            RecGrid2.Fields("MBROKTYPE") = "S Variable Qtywise IntraDay"
                        Case "I"
                            RecGrid2.Fields("MBROKTYPE") = "IntraDay Brokerage"
                        Case "V"
                            RecGrid2.Fields("MBROKTYPE") = "Valuewise Intraday"
                        Case "D"
                            RecGrid2.Fields("MBROKTYPE") = "Delivery Wise Brokerage"
                        Case "H"
                            RecGrid2.Fields("MBROKTYPE") = "Higher Value Percentage Wise"
                        Case "L"
                            RecGrid2.Fields("MBROKTYPE") = "LotWise"
                        Case "X"
                            RecGrid2.Fields("MBROKTYPE") = "XIntraday Higher Wise"
                        Case "Z"
                            RecGrid2.Fields("MBROKTYPE") = "ZLotwise"
                        Case "R"
                            RecGrid2.Fields("MBROKTYPE") = "RZLotwise Intraday"
                        Case "F"
                            RecGrid2.Fields("MBROKTYPE") = "Fixed Brokerage"
                        Case "M"
                            RecGrid2.Fields("MBROKTYPE") = "MRate Wise IntraDay"
                        Case "N"
                            RecGrid2.Fields("MBROKTYPE") = "N Per Trade Wise"
                        Case "U"
                            RecGrid2.Fields("MBROKTYPE") = "U ShareQty Wise"
                        Case "Y"
                            RecGrid2.Fields("MBROKTYPE") = "Y Qtywise IntraDay"
                        Case "1"
                            RecGrid2.Fields("MBROKTYPE") = "1 RateWise Percentage Wise"
                        Case "2"
                            RecGrid2.Fields("MBROKTYPE") = "2 MinRate Percentage Wise"
                        Case "3"
                            RecGrid2.Fields("MBROKTYPE") = "3 BuySell Percentagewise"
                        Case "4"
                            RECGRID.Fields("BROKTYPE") = "4 Valuewise Intraday 2"
                        Case "5"
                            RecGrid2.Fields("MBROKTYPE") = "5 Closing Sauda Zlotwise"
                    End Select
                
                End If
                RecGrid2.Fields("MARRATE") = IIf(IsNull(BrokRec!MARRATE), 0, Val(BrokRec!MARRATE & vbNullString))
                CountRow = CountRow + 1
                RecGrid2.Fields("New") = CountRow
                RecGrid2.Fields("PARTY") = BrokRec!NAME
                RecGrid2.Fields("PARTYCODE") = BrokRec!AC_CODE
                RecGrid2.Fields("ACCID") = BrokRec!ACCID
                RecGrid2.Fields("DELETEROW") = "N"
                RecGrid2.Update
                BrokRec.MoveNext
            Loop
            RecGrid2.AddNew
            RecGrid2.Fields("EXCODE") = vbNullString:       RecGrid2.Fields("INSTTYPE") = vbNullString
            RecGrid2.Fields("TranType") = vbNullString:     RecGrid2.Fields("UPTOSTDT") = Format(GFinEnd, "YYYY/MM/DD")
            RecGrid2.Fields("BROKTYPE") = vbNullString:     RecGrid2.Fields("MARTYPE") = vbNullString:
            RecGrid2.Fields("MBROKTYPE") = vbNullString:    RecGrid2.Fields("PARTY") = vbNullString:
            RecGrid2.Fields("PARTYCODE") = vbNullString:    RecGrid2.Fields("BROKRATE") = 0
            RecGrid2.Fields("SELLBROK") = 0:                RecGrid2.Fields("SELLBROK2") = 0
            RecGrid2.Fields("STDRATE") = 0:                 RecGrid2.Fields("TRANRATE") = 0:
            RecGrid2.Fields("BROKRATE2") = 0:               RecGrid2.Fields("MARRATE") = 0:
            RecGrid2.Fields("MBROKRATE") = 0:               RecGrid2.Fields("MINRATE") = 0:
            RecGrid2!EXID = 0
            RecGrid2!itemid = 0
            CountRow = CountRow + 1
            RecGrid2.Fields("New") = CountRow
            RecGrid2.Update
            Set DataGrid2.DataSource = RecGrid2: DataGrid2.ReBind: DataGrid1.Refresh: RecGrid2.MoveFirst: DataGrid1.SetFocus
            DataGrid1.LeftCol = 0
        Else
            If OptItem.Value = True Then
                If MsgBox("No Records. Do you really want to apply Separate Brokergae for selected Item/Script", vbYesNo + vbQuestion, "Confirm New Records") = vbYes Then
                    mysql = " SELECT ACCID,AC_CODE FROM ACCOUNTD WHERE COMPCODE =" & GCompCode & " AND AC_CODE IN (" & LSParties & ") ORDER BY AC_CODE "
                    Set PartyRec = Nothing
                    Set PartyRec = New ADODB.Recordset
                    PartyRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                    Do While Not PartyRec.EOF
                        mysql = "SELECT EXID,ITEMID,EXCHANGECODE,ITEMCODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMID IN (" & LSItems & ") ORDER BY EXCHANGECODE,ITEMCODE"
                        Set TRec = Nothing
                        Set TRec = New ADODB.Recordset
                        TRec.Open mysql, Cnn, adOpenStatic, adLockReadOnly
                        Do While Not TRec.EOF
                            DoEvents
                            Call PInsert_PitBrok(GCompCode, PartyRec!AC_CODE, TRec!ITEMCODE, "P", 0, 0, 0, "P", 0, "I", 0, Format(GFinEnd, "yyyy/MM/dd"), 0, 0, 0, "P", 0, 0, LInstType, TRec!EXCHANGECODE, TRec!EXID, TRec!itemid, PartyRec!ACCID)
                            TRec.MoveNext
                        Loop
                        PartyRec.MoveNext
                    Loop
                    Set PartyRec = Nothing
                    Set TRec = Nothing
                End If
                'CANCEL_REC
                Call OkCmd_Click
                Exit Sub
            End If
            
        End If
    Else
        MsgBox "Please Select Item / Script for which you want to set Seperate Brokerage "
        CANCEL_REC
    End If
    Set BrokRec = Nothing
End Sub
Private Sub Save_ExBrok()
On Error GoTo err1
Dim LAC_CODE As String
Dim LStdTDate As Date:          Dim TempRec As ADODB.Recordset:
Dim TRec As ADODB.Recordset:    Dim TRec2 As ADODB.Recordset:       Dim AccRec As ADODB.Recordset

    If RECGRID.RecordCount > 0 Then
        Set TempRec = Nothing
        Set TempRec = RECGRID.Clone
        TempRec.MoveFirst
        Cnn.BeginTrans: CNNERR = True
        LSItems = vbNullString
        TempRec.MoveFirst
        Do While Not TempRec.EOF
            If Not IsNull(TempRec!excode) Then
                If LenB(TempRec!excode) > 0 Then
                    If DateValue(TempRec!UPTOSTDT) > DateValue(GSysLockDt) Then
                        mysql = "DELETE FROM PEXBROK WHERE COMPCODE =" & GCompCode & " AND INSTTYPE='" & LInstType & "' AND AC_CODE = '" & TempRec!PARTYCODE & "'"
                        mysql = mysql & " AND EXID =" & TempRec!EXID & ""
                        mysql = mysql & " AND UPTOSTDT = '" & Format(TempRec!UPTOSTDT, "yyyy/MM/dd") & "' "
                        mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                        Cnn.Execute mysql
                        If TempRec!DELETEROW <> "Y" Then
                            'MYSQL = "EXEC INSERT_PEXBROK " & GCompCode & ",'" & TempRec!PARTYCODE & "','" & TempRec!EXCODE & "','" & Left$(TempRec!BROKTYPE, 1) & "'," & Val(TempRec!BROKRATE) & "," & Val(TempRec!BROKRATE2) & "," & Val(TempRec!STDRATE) & "," & Val(TempRec!TRANRATE) & ",'" & Left$(TempRec!TRANTYPE, 1) & "',"
                            'MYSQL = MYSQL & Val(0) & "," & Val(TempRec!MINRATE) & ",'" & Left$(TempRec!MBROKTYPE, 1) & "'," & Val(TempRec!MBrokRate) & "," & Val(TempRec!MBROKRATE2) & ",'" & Left$(TempRec!MARTYPE, 1) & "'," & Val(TempRec!MARRATE) & ","
                            ''" & Format(TempRec!uptostdt, "YYYY/MM/DD") & "','" & LInstType & "'," & TempRec!SELLBROK & "," & TempRec!SELLBROK2 & ""
                            'Cnn.Execute MYSQL
                            Call PInsert_PExBrok(TempRec!PARTYCODE, TempRec!excode, Left$(TempRec!broktype, 1), Val(TempRec!brokrate & vbNullString), Val(TempRec!BROKRATE2 & vbNullString), Val(TempRec!STDRATE & vbNullString), _
                            Val(TempRec!TRANRATE), Left$(TempRec!TRANTYPE, 1), 0, Val(TempRec!MinRate & vbNullString), Left$(TempRec!MBROKTYPE, 1), Val(TempRec!MBROKRATE & vbNullString), Val(TempRec!MBROKRATE2 & vbNullString) _
                            , Left$(TempRec!MARTYPE, 1), Val(TempRec!MARRATE), TempRec!UPTOSTDT, LInstType, Val(TempRec!STAMPRATE & vbNullString), TempRec!SEBITAX, TempRec!EXID, TempRec!ACCID)
                        End If
                    Else
                        MsgBox "Sorry System Locked.  No Modification Allowed for " & TempRec!UPTOSTDT & ""
                    End If
                End If
            End If
            DoEvents
            TempRec.MoveNext
        Loop
        LSettlementDt = GFinEnd
        Set TempRec = Nothing
        Set AccRec = Nothing
        Set AccRec = New ADODB.Recordset
        mysql = "SELECT ACCID,AC_CODE FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " "
        mysql = mysql & " AND AC_CODE IN (" & LSParties & ") ORDER BY AC_CODE"
        AccRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        While Not AccRec.EOF
            mysql = "SELECT EXCODE,EXID FROM EXMAST WHERE COMPCODE  = " & GCompCode & " "
            If LenB(LExCodes) > 0 Then mysql = mysql & " AND EXID  IN (" & LExCodes & ")"
            mysql = mysql & "  ORDER BY EXCODE "
            Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not TRec.EOF
                DoEvents
                LAC_CODE = Get_PEXBROK_AC_CODE(AccRec!ACCID, TRec!EXID, DateValue(LSettlementDt), LInstType)
                If Len(LAC_CODE) < 1 Then
                    If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                        Call PInsert_PExBrok(AccRec!AC_CODE, TRec!excode, "P", 0, 0, 0, 0, "P", 0, 0, "P", 0, 0, "V", 0, DateValue(LSettlementDt), LInstType, 0, 0, TRec!EXID, AccRec!ACCID)
                        'MYSQL = "EXEC INSERT_PEXBROK " & GCompCode & ",'" & AccRec!ac_code & "','" & TRec!EXCODE & "','P',0,0,0,0,'P',"
                        'MYSQL = MYSQL & "0,0,'P',0,0,'I',0,'" & Format(LSettlementDt, "YYYY/MM/DD") & "','" & LInstType & "',0,0"
                        'Cnn.Execute MYSQL
                    End If
                End If
                TRec.MoveNext
            Loop
            Set TRec = Nothing
            AccRec.MoveNext
        Wend
        Set AccRec = Nothing
    End If
    Exit Sub
err1:
    MsgBox err.Description
    If CNNERR = True Then Cnn.RollbackTrans
End Sub

Private Sub Save_ItemBrok()
On Error GoTo err1
Dim LStdTDate As Date:          Dim TempRec2 As ADODB.Recordset
Dim TRec As ADODB.Recordset:    Dim TRec2 As ADODB.Recordset
Dim AccRec As ADODB.Recordset
If RecGrid2.RecordCount > 0 Then
    Set TempRec2 = Nothing:        Set TempRec2 = RecGrid2.Clone
    TempRec2.MoveFirst:        Cnn.BeginTrans: CNNERR = True
    TempRec2.MoveFirst
    Do While Not TempRec2.EOF
        If Not IsNull(TempRec2!ITEMCODE) Then
            If Trim(TempRec2!ITEMCODE) <> "" Then
                mysql = "DELETE FROM PITBROK WHERE COMPCODE =" & GCompCode & " AND INSTTYPE='" & LInstType & "' AND AC_CODE = '" & TempRec2!PARTYCODE & "'"
                mysql = mysql & " AND ITEMID =" & TempRec2!itemid & ""
                mysql = mysql & " AND UPTOSTDT = '" & Format(TempRec2!UPTOSTDT, "yyyy/MM/dd") & "' "
                mysql = mysql & " AND UPTOSTDT>'" & Format(GSysLockDt, "YYYY/MM/DD") & "'"
                Cnn.Execute mysql
                If ChkDelItemBrok.Value = False Then
                    If TempRec2!DELETEROW <> "Y" Then
                        If DateValue(TempRec2!UPTOSTDT) > DateValue(GSysLockDt) Then
                            Call PInsert_PitBrok(GCompCode, TempRec2!PARTYCODE, TempRec2!ITEMCODE, Left$(TempRec2!broktype, 1), Val(TempRec2!brokrate & ""), Val(TempRec2!STDRATE & ""), Val(TempRec2!TRANRATE & ""), Left$(TempRec2!TRANTYPE, 1), TempRec2!BROKRATE2, Left$(TempRec2!MARTYPE, 1), Val(TempRec2!MARRATE & ""), Format(TempRec2!UPTOSTDT, "yyyy/MM/dd"), 0, 0, Val(TempRec2!MinRate), Left$(TempRec2!MBROKTYPE, 1), Val(TempRec2!MBROKRATE), Val(TempRec2!MBROKRATE2), LInstType, TempRec2!excode, TempRec2!EXID, TempRec2!itemid, TempRec2!ACCID)
                        Else
                            MsgBox "Sorry System Locked.  No Modification Allowed"
                            Exit Do
                        End If
                    End If
                End If
            End If
        End If
        DoEvents
        TempRec2.MoveNext
    Loop
    Set TempRec2 = Nothing
    LSettlementDt = GFinEnd
    If ChkUpdLastSettle.Value = 1 Then
        Call Get_Items
        Set AccRec = Nothing
        Set AccRec = New ADODB.Recordset
        mysql = "SELECT ACCID,AC_CODE FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " "
        If LenB(LSParties) > 0 Then mysql = mysql & " AND AC_CODE IN (" & LSParties & ") ORDER BY AC_CODE"
        AccRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
        While Not AccRec.EOF
            mysql = "SELECT EXID,ITEMID,EXCHANGECODE,ITEMCODE FROM ITEMMAST  WHERE COMPCODE  = " & GCompCode & " "
            If LenB(LSItems) > 1 Then mysql = mysql & " AND ITEMID  IN (" & LSItems & ")"
            mysql = mysql & "  ORDER BY EXCHANGECODE,ITEMCODE  "
            Set TRec = Nothing: Set TRec = New ADODB.Recordset: TRec.Open mysql, Cnn, adOpenForwardOnly, adLockReadOnly
            Do While Not TRec.EOF
                DoEvents
                mysql = "SELECT AC_CODE FROM PITBROK WHERE COMPCODE=" & GCompCode & " AND AC_CODE ='" & AccRec!AC_CODE & "' AND ITEMID  =" & TRec!itemid & " AND UptoStdt='" & Format(LSettlementDt, "yyyy/MM/dd") & "' AND INSTTYPE ='" & LInstType & "'"
                Set TRec2 = Nothing: Set TRec2 = New ADODB.Recordset: TRec2.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
                If TRec2.EOF Then
                    If DateValue(LSettlementDt) > DateValue(GSysLockDt) Then
                        Call PInsert_PitBrok(GCompCode, AccRec!AC_CODE, TRec!ITEMCODE, "P", 0, 0, 0, "P", 0, "I", 0, Format(LSettlementDt, "yyyy/MM/dd"), 0, 0, 0, "P", 0, 0, LInstType, TRec!EXCHANGECODE, TRec!EXID, TRec!itemid, AccRec!ACCID)
                    End If
                End If
                TRec.MoveNext
            Loop
            AccRec.MoveNext
        Wend
        Set TRec = Nothing
        Set AccRec = Nothing
    End If
End If
Exit Sub
err1:
    If CNNERR = True Then Cnn.RollbackTrans
    MsgBox err.Description
End Sub

Private Sub Fill_ItemList()
    Dim ItemRec As ADODB.Recordset:
    Call Get_ExCodes
    Set ItemRec = Nothing: Set ItemRec = New ADODB.Recordset
    mysql = "SELECT ITEMID,ITEMCODE,ITEMNAME,EXCHANGECODE ,EXID FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " "
    If LenB(LExCodes) > 0 Then mysql = mysql & " AND EXID IN (" & LExCodes & ")"
    If LenB(LExCodes) > 0 Then mysql = mysql & " AND EXCHANGECODE <>'BEQ' AND EXCHANGECODE <> 'EQ'"
    mysql = mysql & " ORDER BY ITEMCODE "
    ItemRec.Open mysql, Cnn, adOpenKeyset, adLockReadOnly
    If Not ItemRec.EOF Then
        ItemRec.MoveFirst
        ItemList.ListItems.Clear
        ItemList.Enabled = True:
        Do While Not ItemRec.EOF
            If (ItemRec!EXCHANGECODE = "EQ" Or ItemRec!EXCHANGECODE = "BEQ") Then InstTypeCombo.Visible = True
            ItemList.ListItems.Add , , ItemRec!ITEMCODE
            'ItemList.ListItems(ItemList.ListItems.Count).ListSubItems.Add , , (ItemRec!ITEMName & vbNullString)
            ItemList.ListItems(ItemList.ListItems.Count).ListSubItems.Add , , ItemRec!EXCHANGECODE
            ItemList.ListItems(ItemList.ListItems.Count).ListSubItems.Add , , ItemRec!EXID
            ItemList.ListItems(ItemList.ListItems.Count).ListSubItems.Add , , ItemRec!itemid
            ItemRec.MoveNext
        Loop
    End If
    Set ItemRec = Nothing
        
End Sub
Private Sub Get_ExCodes()
    Dim ChkCount As Integer
    Dim I As Integer
    LExCodes = vbNullString
    ChkCount = 0
    For I = 1 To ExchnageList.ListItems.Count
        If ExchnageList.ListItems(I).Checked = True Then
            ChkCount = ChkCount + 1
            If LenB(LExCodes) > 0 Then LExCodes = LExCodes & ","
            LExCodes = LExCodes & ExchnageList.ListItems(I).ListSubItems(2) & ""
        End If
    Next I

End Sub

Private Sub Get_Items()
    Dim ChkCount As Integer
    Dim I As Integer
    LSItems = vbNullString
    For I = 1 To ItemList.ListItems.Count
        If ItemList.ListItems(I).Checked = True Then
            If LenB(LSItems) <> 0 Then LSItems = LSItems & ","
            LSItems = LSItems & "" & ItemList.ListItems(I).SubItems(3) & ""
        End If
    Next I
End Sub

Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LGridRow As Integer:    Dim LGridCol As Integer
    If KeyCode = 13 And DataGrid2.Col = 3 Then ' BROKTYPE
        BrokTypeCombo2.Visible = True: BrokTypeCombo2.SetFocus
    ElseIf KeyCode = 13 And DataGrid2.Col = 10 Then 'TRANTYPE
        BrokTypeCombo2.Visible = True: BrokTypeCombo2.SetFocus
    ElseIf KeyCode = 13 And DataGrid2.Col = 13 Then 'MBROKTYPE
        BrokTypeCombo2.Visible = True: BrokTypeCombo2.SetFocus
    ElseIf KeyCode = 13 And DataGrid2.Col = 16 Then 'MARGIN TYPE
        MarginCombo.Visible = True: MarginCombo.SetFocus
    ElseIf KeyCode = 13 And DataGrid2.Col = 18 Then
        DataGrid2.text = UCase(DataGrid2.text)
        If DataGrid2.text = "Y" Then
        Else
            DataGrid2.text = "N"
        End If
        DataGrid2.SetFocus
        DataGrid2.LeftCol = 0: DataGrid2.Col = 0
    ElseIf KeyCode = 13 Then
        Sendkeys "{TAB}"
    End If
    'Press F7 to replace all rows with current cell value
    If KeyCode = 118 Then   'F7
        LGridRow = DataGrid2.Row
        LGridCol = DataGrid2.Col
        If DataGrid2.Col = 3 Then 'BROKTYPE
            GridColVal = RecGrid2!broktype
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!broktype = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 4 Then 'BROKRATE
            GridColVal = RecGrid2!brokrate
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!brokrate = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 6 Then 'brokrate2
            GridColVal = RecGrid2!BROKRATE2
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!BROKRATE2 = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 8 Then 'STDRATE
            GridColVal = RecGrid2!STDRATE
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!STDRATE = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 9 Then 'TRANRATE
            GridColVal = RecGrid2!TRANRATE
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!TRANRATE = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 10 Then 'TRANTYPE
            GridColVal = IIf(IsNull(RecGrid2!TRANTYPE), "P", RecGrid2!TRANTYPE)
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!TRANTYPE = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 12 Then 'MINRATE
            GridColVal = RecGrid2!MinRate
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!MinRate = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 13 Then 'MBROKTYPE
            GridColVal = RecGrid2!MBROKTYPE
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!MBROKTYPE = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 14 Then 'MBROKTRATE
            GridColVal = RecGrid2!MBROKRATE
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!MBROKRATE = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 15 Then 'MBROKTRATE
            GridColVal = RecGrid2!MBROKRATE2
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!MBROKRATE2 = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 16 Then 'MARTYPE
            GridColVal = RecGrid2!MARTYPE
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!MARTYPE = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 17 Then 'MARRATE
            GridColVal = RecGrid2!MARRATE
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!MARRATE = GridColVal
                RecGrid2.MoveNext
            Wend
        ElseIf DataGrid2.Col = 18 Then 'DELETEROW
            GridColVal = RecGrid2!DELETEROW
            RecGrid2.MoveFirst
            While Not RecGrid2.EOF
                RecGrid2!DELETEROW = GridColVal
                RecGrid2.MoveNext
            Wend
        End If
        Set DataGrid2.DataSource = RecGrid2: DataGrid2.ReBind: DataGrid2.Refresh
        DataGrid2.Row = LGridRow: DataGrid2.Col = LGridCol: DataGrid2.SetFocus
    End If
End Sub


Private Sub BrokTypeCombo2_GotFocus()
    If DataGrid2.Col = 2 Then
            Select Case Left$(RECGRID!broktype, 1)
            Case "T"
                BrokTypeCombo2.ListIndex = 0
            Case "O"
                BrokTypeCombo2.ListIndex = 1
            Case "P"
                BrokTypeCombo2.ListIndex = 2
            Case "I"
                BrokTypeCombo2.ListIndex = 3
            Case "C"
                BrokTypeCombo2.ListIndex = 4
            Case "V"
                BrokTypeCombo2.ListIndex = 5
            Case "Q"
                BrokTypeCombo2.ListIndex = 6
            Case "D"
                BrokTypeCombo2.ListIndex = 7
            Case "H"
                BrokTypeCombo2.ListIndex = 8
            Case "L"
                BrokTypeCombo2.ListIndex = 9
            Case "W"
                BrokTypeCombo2.ListIndex = 10
            Case "X"
                BrokTypeCombo2.ListIndex = 11
            Case "Z"
                BrokTypeCombo2.ListIndex = 12
            Case "R"
                BrokTypeCombo2.ListIndex = 13
            Case "F"
                BrokTypeCombo2.ListIndex = 14
            Case "M"
                BrokTypeCombo2.ListIndex = 15
            Case "B"
                BrokTypeCombo2.ListIndex = 16
            Case "N"
                BrokTypeCombo2.ListIndex = 17
            Case "U"
                BrokTypeCombo2.ListIndex = 18
            Case "Y"
                BrokTypeCombo2.ListIndex = 19
            Case "S"
                BrokTypeCombo2.ListIndex = 20
            Case "A"
                BrokTypeCombo2.ListIndex = 21
            Case "1"
                BrokTypeCombo2.ListIndex = 22
            Case "2"
                BrokTypeCombo2.ListIndex = 23
            Case "3"
                BrokTypeCombo2.ListIndex = 24
            Case "4"
                BrokTypeCombo2.ListIndex = 25
                
            End Select
    ElseIf DataGrid2.Col = 6 Then
            If Mid(RecGrid2!TRANTYPE, 1, 1) = "T" Then
                BrokTypeCombo2.ListIndex = Val(0)
            ElseIf Mid(RecGrid2!TRANTYPE, 1, 1) = "P" Then
                BrokTypeCombo2.ListIndex = Val(1)
            End If
    ElseIf DataGrid2.Col = 13 Then
            Select Case Left$(RecGrid2!MBROKTYPE, 1)
            Case "T"
                BrokTypeCombo2.ListIndex = 0
            Case "O"
                BrokTypeCombo2.ListIndex = 1
            Case "P"
                BrokTypeCombo2.ListIndex = 2
            Case "I"
                BrokTypeCombo2.ListIndex = 3
            Case "C"
                BrokTypeCombo2.ListIndex = 4
            Case "V"
                BrokTypeCombo2.ListIndex = 5
            Case "Q"
                BrokTypeCombo2.ListIndex = 6
            Case "D"
                BrokTypeCombo2.ListIndex = 7
            Case "H"
                BrokTypeCombo2.ListIndex = 8
            Case "L"
                BrokTypeCombo2.ListIndex = 9
            Case "W"
                BrokTypeCombo2.ListIndex = 10
            Case "X"
                BrokTypeCombo2.ListIndex = 11
            Case "Z"
                BrokTypeCombo2.ListIndex = 12
            Case "R"
                BrokTypeCombo2.ListIndex = 13
            Case "F"
                BrokTypeCombo2.ListIndex = 14
            Case "M"
                BrokTypeCombo2.ListIndex = 15
            Case "B"
                BrokTypeCombo2.ListIndex = 16
            Case "N"
                BrokTypeCombo2.ListIndex = 17
            Case "U"
                BrokTypeCombo2.ListIndex = 18
            Case "Y"
                BrokTypeCombo2.ListIndex = 19
            Case "S"
                BrokTypeCombo2.ListIndex = 20
            Case "1"
                BrokTypeCombo2.ListIndex = 22
            Case "2"
                BrokTypeCombo2.ListIndex = 23
            Case "3"
                BrokTypeCombo2.ListIndex = 24
            Case "4"
                BrokTypeCombo2.ListIndex = 25
            Case "5"
                BrokTypeCombo2.ListIndex = 26
            End Select
            
    End If
    BrokTypeCombo2.Top = Val(DataGrid2.Top) + Val(DataGrid2.RowTop(DataGrid2.Row))
    BrokTypeCombo2.Width = Val(DataGrid2.Columns(DataGrid2.Col).Width)
    BrokTypeCombo2.Left = Val(DataGrid2.Left) + Val(DataGrid2.Columns(DataGrid2.Col).Left)
    Sendkeys "%{DOWN}"
End Sub
Private Sub BrokTypeCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim LGridRow As Integer:        Dim LGridCol As Integer:        Dim SearchRow As Integer
    If KeyCode = 13 Then
        LGridRow = DataGrid2.Row: LGridCol = DataGrid2.Col:  SearchRow = RecGrid2!New
        If DataGrid2.Col = 3 Then
            If KeyCode = 13 Then RecGrid2!broktype = BrokTypeCombo2.text
        ElseIf DataGrid2.Col = 8 Then
            If KeyCode = 13 Then RecGrid2!TRANTYPE = BrokTypeCombo2.text
        ElseIf DataGrid2.Col = 11 Then
            If KeyCode = 13 Then RecGrid2!MBROKTYPE = BrokTypeCombo2.text
        ElseIf DataGrid2.Col = 13 Then
            If KeyCode = 13 Then RecGrid2!MBROKTYPE = BrokTypeCombo2.text
        End If
        Set DataGrid2.DataSource = RecGrid2: DataGrid2.ReBind: DataGrid2.Refresh
        RecGrid2.MoveFirst: DataGrid2.SetFocus
        RECGRID.Find "new =" & SearchRow & "", , adSearchForward
        DataGrid2.Row = LGridRow: DataGrid2.Col = LGridCol + 1: BrokTypeCombo2.Visible = False: DataGrid2.SetFocus
    ElseIf KeyCode = 27 Then
        BrokTypeCombo2.Visible = False
    End If
End Sub
Private Sub BrokTypeCombo2_Validate(Cancel As Boolean)
    If Len(Trim(BrokTypeCombo2.text)) < 1 Then
        Cancel = True
        Exit Sub
    End If
End Sub
