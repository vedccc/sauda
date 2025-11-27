VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "datectl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SINGLECONTRACT 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   12735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18765
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
   ScaleHeight     =   12735
   ScaleWidth      =   18765
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   1695
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00808000&
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
      TabIndex        =   29
      Top             =   0
      Width           =   15975
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5880
         TabIndex        =   47
         Top             =   120
         Width           =   3255
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Contract Entry"
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
            TabIndex        =   48
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   9960
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Text2"
         Top             =   120
         Width           =   975
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
      TabIndex        =   6
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   240
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   10080
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
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
      Height          =   8205
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   15420
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   7320
         Width           =   15135
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "Text4"
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "Text7"
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   10920
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "Text8"
            Top             =   120
            Width           =   1215
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
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "Text9"
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "Text10"
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "S Amount"
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
            Left            =   12240
            TabIndex        =   46
            Top             =   165
            Width           =   915
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Avg Rate"
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
            Left            =   9960
            TabIndex        =   45
            Top             =   165
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "B Amount"
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
            Left            =   4800
            TabIndex        =   44
            Top             =   165
            Width           =   930
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Avg Rate"
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
            Left            =   2520
            TabIndex        =   43
            Top             =   165
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Totals"
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
            TabIndex        =   42
            Top             =   165
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Buy"
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
            Left            =   960
            TabIndex        =   41
            Top             =   165
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sell"
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
            Left            =   8280
            TabIndex        =   40
            Top             =   165
            Width           =   330
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   15135
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   13440
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
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
            Left            =   10080
            List            =   "SINGLECONTRACT.frx":08C7
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   120
            Width           =   1335
         End
         Begin vcDateTimePicker.vcDTP vcDTP1 
            Height          =   360
            Left            =   720
            TabIndex        =   22
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
            Left            =   3405
            TabIndex        =   23
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
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
            Left            =   6885
            TabIndex        =   24
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
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
         Begin VB.Label Label8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Admin Pass"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12120
            TabIndex        =   32
            Top             =   180
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   27
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   18
            Left            =   9480
            TabIndex        =   26
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sauda"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   6240
            TabIndex        =   25
            Top             =   180
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdImportFromExcel 
         Caption         =   "..."
         Height          =   285
         Left            =   -360
         TabIndex        =   13
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
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
            Name            =   "Times New Roman"
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
         Height          =   6180
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   10901
         _Version        =   393216
         AllowArrows     =   -1  'True
         ForeColor       =   128
         HeadLines       =   1
         RowHeight       =   21
         TabAction       =   1
         FormatLocked    =   -1  'True
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
            Size            =   11.25
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
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
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
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
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
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
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
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
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
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
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
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
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
               Object.Visible         =   0   'False
               ColumnWidth     =   1019.906
            EndProperty
         EndProperty
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
         TabIndex        =   14
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
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000040&
      BorderWidth     =   12
      Height          =   8460
      Left            =   0
      Top             =   1080
      Width           =   15765
   End
End
Attribute VB_Name = "SINGLECONTRACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Dim flag As Boolean::               Dim CONFLAGE As Boolean:
'''Dim LBConfirm As Integer:           Dim LSConfirm As Integer:
'''Dim OldDate As Date:                Dim LExCode  As String:
'''Dim LParty As String:               Dim LUserId As String:
'''Dim LContractAcc As String:         Dim VchNo As String:
'''Dim LSaudaCode As String:           Dim LItemCode As String:
'''Dim SAUDACODE As String:            Public Fb_Press As Byte:
'''Dim LDataImport As Byte:            Dim FLOWDIR As Byte::
'''Dim GRIDPOS As Byte:                Dim LConNo As Long::
'''Dim LConSno As Long:
'''Dim GRIDREC As ADODB.Recordset:     Dim ItemRec As ADODB.Recordset:
'''Dim SaudaRec As ADODB.Recordset:    Dim TEMPORARY As ADODB.Recordset:
'''Dim RecEx As ADODB.Recordset:       Dim RECGRID As ADODB.Recordset:
'''Dim Rec_Sauda As ADODB.Recordset:   Dim Rec_Account As ADODB.Recordset:
'''Dim REC_CloRate As ADODB.Recordset: Dim REC_CTRM As ADODB.Recordset:
'''Dim SaudaCmbRec As ADODB.Recordset:
'''Sub SaudaList()
'''If ITEMCMB.BoundText <> "" Then
'''    MYSQL = "SELECT SAUDACODE ,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE='" & ITEMCMB.BoundText & "' AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY SAUDANAME"
'''Else
'''    MYSQL = "SELECT SAUDACODE ,SAUDANAME FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & "  AND MATURITY >='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY SAUDANAME"
'''End If
'''    Set SaudaCmbRec = Nothing: Set SaudaCmbRec = New ADODB.Recordset
'''    SaudaCmbRec.Open MYSQL, Cnn
'''    If Not SaudaCmbRec.EOF Then
'''        Set Saudacmb.RowSource = SaudaCmbRec
'''        Saudacmb.ListField = "SAUDANAME"
'''        Saudacmb.BoundColumn = "SAUDACODE"
'''    End If
'''
'''End Sub
'''Sub ItemList()
'''MYSQL = "SELECT DISTINCT I.ITEMCODE,I.ITEMNAME FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & GCompCode & " AND S.COMPCODE=I.COMPCODE AND I.ITEMCODE IN(SELECT ITEMCODE FROM SAUDAMAST WHERE COMPCODE='" & GCompCode & "'AND MATURITY>='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "') ORDER BY I.ITEMNAME"
'''Set ItemRec = Nothing: Set ItemRec = New ADODB.Recordset: ItemRec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
'''If Not ItemRec.EOF Then
'''    Set ITEMCMB.RowSource = ItemRec: ITEMCMB.BoundColumn = "ITEMCODE": ITEMCMB.ListField = "ITEMNAME"
'''End If
'''End Sub
'''Sub ADD_REC()
'''    If Rec_Account.RecordCount > 0 Then
'''        LDataImport = 0
'''        Frame1.Enabled = True: Combo1.ListIndex = 0
'''        Call Get_Selection(1)
'''        If vcDTP1.Enabled Then vcDTP1.SetFocus
'''    Else
'''        Call CANCEL_REC
'''    End If
'''    If Fb_Press = 1 Then
'''        MYSQL = "SELECT MAX(CAST(CONNO AS INT)) AS CONNO FROM CTR_D WHERE COMPCODE =" & GCompCode & ""
'''        Set CONREC = Nothing
'''        Set CONREC = New ADODB.Recordset
'''        CONREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'''        If Not CONREC.EOF Then
'''            LConNo = Val(CONREC!CONNO & "") + Val(1)
'''        Else
'''            LConNo = 1
'''        End If
'''    End If
'''    RECGRID.AddNew
'''    RECGRID!DIMPORT = 0
'''    RECGRID!CONTIME = Time
'''    RECGRID!USERID = LUserId
'''    RECGRID!bconfirm = 0
'''    RECGRID!SCONFIRM = 0
'''    RECGRID.Update
'''    LConNo = LConNo
'''    RECGRID!SrNo = LConNo  'RECGRID.AbsolutePosition
'''    DataGrid1.Col = 0
'''
'''End Sub
'''Sub Save_Rec()
'''
'''    Dim LBuyer  As String
'''    Dim LSeller As String
'''    Dim BRATE As Double
'''    Dim SRATE As Double
'''
'''
'''    Dim LPtyContype As String
'''    On Error GoTo ERR1
'''    'validation
'''    If vcDTP1.Value < GFinBegin Then MsgBox "Date can not be before financial year begin date.", vbCritical: vcDTP1.SetFocus: Exit Sub
'''    If vcDTP1.Value > GFinEnd Then MsgBox "Date can not be beyond financial year end date.", vbCritical: vcDTP1.Enabled = True: vcDTP1.SetFocus: Exit Sub
'''    Cnn.BeginTrans
'''    If Fb_Press = 2 Then
'''        Set REC_CTRM = Nothing
'''        Set REC_CTRM = New ADODB.Recordset
'''        MYSQL = "SELECT CONSNO,VOU_NO FROM CTR_M WHERE  COMPCODE=" & GCompCode & " AND "
'''        If ITEMCMB.BoundText <> "" Then
'''            MYSQL = MYSQL & " ITEMCODE='" & ITEMCMB.BoundText & "'  AND  "
'''        End If
'''        If Saudacmb.BoundText <> "" Then
'''            MYSQL = MYSQL & " SAUDA ='" & Saudacmb.BoundText & "'  AND  "
'''        End If
'''        MYSQL = MYSQL & " CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' ORDER BY CONSNO"
'''        REC_CTRM.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'''        If Not REC_CTRM.EOF Then
'''            REC_CTRM.MoveFirst
'''            While Not REC_CTRM.EOF
'''                LConSno = REC_CTRM!consno
'''                VchNo = REC_CTRM!VOU_NO & ""
'''                Call Delete_Voucher(VchNo)
'''                Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & Val(LConSno) & ""
'''                Cnn.Execute "DELETE FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND CONSNO=" & Val(LConSno) & ""
'''                REC_CTRM.MoveNext
'''            Wend
'''        End If
'''    End If
'''    Set GRIDREC = Nothing: Set GRIDREC = New ADODB.Recordset
'''    Set GRIDREC = RECGRID.Clone
'''    GRIDREC.MoveFirst
'''    While Not GRIDREC.EOF
'''        If GRIDREC!QNTY <= 0 Then
'''            GRIDREC.Delete
'''        End If
'''        GRIDREC.MoveNext
'''    Wend
'''    GRIDREC.Sort = "SAUDACODE"
'''    GRIDREC.MoveFirst
'''    SAUDACODE = ""
'''    MSAmt = 0
'''    MBAmt = 0
'''    While Not GRIDREC.EOF
'''        If SAUDACODE = GRIDREC!SAUDACODE Then
'''        Else
'''            SAUDACODE = GRIDREC!SAUDACODE
'''            Set Rec_Sauda = Nothing: Set Rec_Sauda = New ADODB.Recordset
'''            Rec_Sauda.Open "SELECT ITEMCODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDACODE='" & SAUDACODE & "'", Cnn, adOpenForwardOnly, adLockReadOnly
'''            If Rec_Sauda.EOF Then
'''                MsgBox "Invalid Sauda Code.", vbExclamation, "Error": Text2.SetFocus: Exit Sub
'''            Else
'''                Set GeneralRec1 = Nothing: Set GeneralRec1 = New ADODB.Recordset
'''                GeneralRec1.Open "SELECT EX.EXCODE,EX.SHREEAC,EX.TRADINGACC  FROM EXMAST AS EX , ITEMMAST AS IM WHERE EX.COMPCODE=" & GCompCode & " AND EX.COMPCODE=IM.COMPCODE AND EX.EXCODE=IM.EXCHANGECODE  AND  IM.ITEMCODE = '" & Rec_Sauda!ITEMCODE & "'", Cnn, adOpenForwardOnly, adLockReadOnly
'''                If Not GeneralRec1.EOF Then
'''                    GShree = GeneralRec1!shreeac
'''                    GTrading = GeneralRec1!TRADINGACC
'''                    LExCode = GeneralRec1!EXCODE
'''                End If
'''            End If
'''            If GRIDREC.RecordCount > 0 Then
'''                CNNERR = True
'''                VchNo = Get_VouNo("CONT", GFinYear)
'''                Set Rec = Nothing: Set Rec = New ADODB.Recordset
'''                Rec.Open "SELECT CONSNO FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND SAUDA='" & SAUDACODE & "' AND CONDATE = '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "' AND PATTAN = '" & Mid(Combo1.text, 1, 1) & "'", Cnn, adOpenForwardOnly, adLockReadOnly
'''                If Not Rec.EOF Then
'''                    consno = Rec!consno
'''                Else
'''                    Set Rec = Nothing: Set Rec = New ADODB.Recordset
'''                    Rec.Open "SELECT MAX(CONSNO) FROM CTR_M WHERE COMPCODE =" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
'''                    consno = Val(Rec.Fields(0) & "") + Val(1)
'''                End If
'''                Set Rec = Nothing
'''                LDataImport = 0
'''                If GRIDREC!ITEMCODE = "" Then
'''                    MsgBox "PLEASE CHECK ITEM IN THIS CONTRACT. ENTRY NOT SAVED"
'''                    Exit Sub
'''                Else
'''                    MYSQL = "INSERT INTO CTR_M(COMPCODE,CONSNO, CONDATE, SAUDA, ITEMCODE, VOU_NO, PATTAN,DataImport) VALUES(" & GCompCode & "," & consno & ", '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "', '" & SAUDACODE & "', '" & GRIDREC!ITEMCODE & "', '" & VchNo & "', '" & LEFT$(Combo1.text, 1) & "'," & LDataImport & ")"
'''                    Cnn.Execute MYSQL
'''                    MBAmt = 0
'''                    MSAmt = 0
'''                End If
'''            End If
'''        End If
'''        ''RECORDSET RC IS CHECKING WHETHER THE PARTY IS PERSONNEL OR NOT
'''        MCL = ""
'''        If Len(GRIDREC!NAME & "") > Val(0) And Len(GRIDREC!conName & "") > Val(0) Then   ''WHEN BUYER AND SELLER BOTH ARE THERE
'''            If GRIDREC!QNTY > Val(0) And GRIDREC!Rate > Val(0) And GRIDREC!Rate1 > Val(0) Then                   ''QNTY AND RATE REQUIRED
'''                If GRIDREC!BUYSELL = "B" Then
'''                    MBAmt = MBAmt + (Val(GRIDREC!QNTY & "") * (Round(Val(GRIDREC!Rate & ""), 4)) * GRIDREC!LOT)
'''                    LDataImport = Abs(GRIDREC!DIMPORT)
'''                    MSAmt = MSAmt + (Val(GRIDREC!QNTY & "") * Round(Val(GRIDREC!Rate1) & "", 4) * GRIDREC!LOT)
'''                    LBuyer = GRIDREC!Code
'''                    LSeller = GRIDREC!concode
'''                    BRATE = GRIDREC!Rate
'''                    SRATE = GRIDREC!Rate1
'''
'''                Else
'''                    MSAmt = MSAmt + (Val(GRIDREC!QNTY & "") * (Round(Val(GRIDREC!Rate & ""), 4)) * GRIDREC!LOT)
'''                    LDataImport = 0
'''                    MBAmt = MBAmt + (Val(GRIDREC!QNTY & "") * Round(Val(GRIDREC!Rate1) & "", 4) * GRIDREC!LOT)
'''                    LSeller = GRIDREC!Code
'''                    LBuyer = GRIDREC!concode
'''                    SRATE = GRIDREC!Rate
'''                    BRATE = GRIDREC!Rate1
'''                End If
'''                If Combo1.ListIndex = 0 Then
'''                    PATTAN = "C"
'''                Else
'''                     PATTAN = "O"
'''                End If
'''
'''                LPtyContype = "B"
'''
'''                'MYSQL = "EXEC INSERT_CTR_D " & GCompCode & ",'" & MBCL & "','" & MSCL & "'," & Val(consno) & ",'" & Format(vcDTP1.Value, "YYYY/MM/DD") & "'," & Val(GRIDREC!SrNo) & ",'" & GRIDREC!SAUDACODE & "','" &
'''                'GRIDREC!ITEMCODE & "','" & BUYER & "'," & Val(GRIDREC!Qnty) & "," & BRate & ",'N','" & seller & "'," & Val(GRIDREC!Qnty) & "," & Round(Val(SRate), 4) & "," & LDataImport & ",'" & GRIDREC!CONTIME
'''                '& "','" & IIf(IsNull(GRIDREC!USERID), "", GRIDREC!USERID) & " " & "','" & IIf(IsNull(GRIDREC!ORDER_NO), "", GRIDREC!ORDER_NO) & "','" & IIf(IsNull(GRIDREC!TRADE_NO), Val(GRIDREC!SrNo), GRIDREC!TRADE_NO) & "'
'''                ','" & LExCode & "','" & PATTAN & "','" & GRIDREC!concode & "',1," & GRIDREC!bconfirm & "," & GRIDREC!SCONFIRM & ",'',"
'''                'Cnn.Execute MYSQL
'''                Call Add_To_Ctr_D(LPtyContype, LBuyer, LConSno, vcDTP1.Value, Val(GRIDREC!SrNo), GRIDREC!SAUDACODE, GRIDREC!ITEMCODE, LBuyer, Val(GRIDFREC!QNTY), BRATE, LSeller, GRIDREC!CONTIME, (GRIDREC!ORDER_NO & vbNullString), (GRIDREC!USERID & ""), (GRIDREC!TRADENO & ""), LExCode, 1, 0, vbNullString, "FUT", vbNullString, 0, "0")
'''            End If
'''            GRIDREC.MoveNext
'''        End If
'''    Wend
'''    Call Update_Charges(vbNullString, vbNullString, vbNullString, vbNullString, vcDTP1.Value, vcDTP1.Value, True)
'''    'Call Delete_Inv_D(vbNullString, vbNullString, vbNullString, vcDTP1.Value)
'''
'''    'Call Updt_BrokEXQty(vbNullString, vbNullString, vbNullString, vcDTP1.Value, vcDTP1.Value, vbNullString)
'''    'Call UpdateBrokRateType(vbNullString, vbNullString, vcDTP1.Value, vcDTP1.Value, vbNullString, vbNullString)
'''    Call Shree_Posting(DateValue(vcDTP1.Value))
'''    Cnn.CommitTrans
'''    'If GAppSpread = "Y" And GMarginYN = "Y" Then
'''     '   Call UpdateMargin(vbNullString, vbNullString, vcDTP1.Value, CStr(GFinEnd), "")
'''    'End If
'''    CNNERR = False
'''    Cnn.BeginTrans
'''    CNNERR = True
'''    If BILL_GENERATION(vcDTP1.Value, CDate(GFinEnd), vbNullString, vbNullString, vbNullString) Then
'''        Cnn.CommitTrans: CNNERR = False
'''    Else
'''        Cnn.RollbackTrans: CNNERR = False
'''    End If
'''    Call Chk_Billing
'''    Call CANCEL_REC
'''    Exit Sub
'''ERR1:
'''    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
'''
'''    If CNNERR = True Then
'''        Cnn.RollbackTrans: CNNERR = False
'''    End If
''' End Sub
'''Sub CANCEL_REC()
'''    vcDTP1.Enabled = True:  Combo1.Enabled = True
'''    Call RecSet
'''    Set REC_CTRM = Nothing
'''    Set REC_CTRM = New ADODB.Recordset
'''    MYSQL = "SELECT CONSNO,VOU_NO FROM CTR_M where COMPCODE=" & GCompCode & " ORDER BY CONSNO"
'''    REC_CTRM.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'''    Call ItemList
'''    Call SaudaList
'''    CONFLAGE = False
'''    Fb_Press = 0
'''    Set DataGrid1.DataSource = RECGRID
'''    ITEMCMB.Enabled = True
'''    Saudacmb.Enabled = True
'''    Combo1.Enabled = True
'''    DataGrid1.Refresh
'''    Label2.Visible = False
'''    DataCombo3.Visible = False
'''    Call ClearFormFn(SINGLECONTRACT)
'''    Call Get_Selection(10)
'''    Combo1.ListIndex = -1: Frame1.Enabled = False
'''End Sub
'''Function MODIFY_REC(LCondate As Date, LSaudaCode As String, LItemCode As String, LPattan As String) As Boolean
'''Set REC_CTRM = Nothing
'''Set REC_CTRM = New ADODB.Recordset
'''
'''MYSQL = "SELECT CONSNO,VOU_NO FROM CTR_M WHERE COMPCODE =" & GCompCode & " AND PATTAN='" & LEFT$(LPattan, 1) & "'AND CONDATE='" & Format(LCondate, "YYYY/MM/DD") & "'  "
'''If LItemCode <> "" Then
'''    MYSQL = MYSQL & "AND itemcode='" & LItemCode & "'"
'''End If
'''If LSaudaCode <> "" Then
'''    MYSQL = MYSQL & "AND sauda='" & LSaudaCode & "'"
'''End If
'''MYSQL = MYSQL & " ORDER BY SAUDA"
'''REC_CTRM.Open MYSQL, Cnn, , adLockReadOnly
'''If Not REC_CTRM.EOF Then
'''    vcDTP1.Value = LCondate
'''    ITEMCMB.BoundText = LItemCode
'''    Saudacmb.BoundText = LSaudaCode
'''    If LEFT$(LPattan, 1) = "C" Then
'''        Combo1.ListIndex = 0
'''    Else
'''        Combo1.ListIndex = 1
'''    End If
'''    If Fb_Press = 1 Then
'''        If MsgBox("Contract Already Exist For Selected Criteria.Press OK To Modify The Existing Contracts.", vbQuestion + vbYesNo, "Confirm") = vbYes Then
'''            Fb_Press = 2
'''        Else
'''            Call CANCEL_REC
'''            Exit Function
'''        End If
'''    End If
'''    Call RecSet
'''    Set Rec = Nothing
'''    Set Rec = New ADODB.Recordset
'''    MYSQL = "SELECT I.LOT,C.* FROM CTR_D AS C ,ITEMMAST AS I WHERE C.COMPCODE =" & GCompCode & " AND C.CONDATE='" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' AND I.COMPCODE=C.COMPCODE AND I.ITEMCODE=C.ITEMCODE AND C.PATTAN='" & LEFT$(LPattan, 1) & "' "
'''    If ITEMCMB.BoundText <> "" Then
'''        MYSQL = MYSQL & "AND C.itemcode='" & ITEMCMB.BoundText & "'"
'''    End If
'''    If Saudacmb.BoundText <> "" Then
'''        MYSQL = MYSQL & "AND C.sauda='" & Saudacmb.BoundText & "'"
'''    End If
'''    MYSQL = MYSQL & " ORDER BY C.CONNO,C.ROWNO"
'''    Rec.Open MYSQL, Cnn, , adLockReadOnly
'''
'''    If Not Rec.EOF Then
'''        Rec.MoveFirst
'''        LParty = vbNullString
'''        While Not Rec.EOF
'''            RECGRID.AddNew
'''            RECGRID!SrNo = Rec!CONNO 'RECGRID.AbsolutePositi
'''            If Len(Rec!concode & "") > Val(0) Then
'''                If Rec!PARTY <> Rec!concode Then
'''                    RECGRID!BUYSELL = Rec!CONTYPE
'''                    RECGRID!Code = Rec!PARTY & ""
'''                    Rec_Account.MoveFirst
'''                    Rec_Account.Find "AC_CODE='" & Rec!PARTY & "'"
'''                    RECGRID!NAME = Rec_Account!NAME
'''                    RECGRID!Rate = Round(Rec!Rate, 4)
'''                Else
'''                    RECGRID!BUYSELL = "S"
'''                    RECGRID!concode = Rec!PARTY & ""
'''                    Rec_Account.MoveFirst
'''                    Rec_Account.Find "AC_CODE='" & Rec!PARTY & "'"
'''                    RECGRID!conName = Rec_Account!NAME
'''                    RECGRID!Rate1 = Round(Rec!Rate, 4)
'''                End If
'''            Else
'''                RECGRID!BUYSELL = Rec!CONTYPE
'''                RECGRID!Code = Rec!PARTY & ""
'''                Rec_Account.MoveFirst
'''                Rec_Account.Find "AC_CODE='" & Rec!PARTY & "'"
'''                RECGRID!NAME = Rec_Account!NAME
'''                RECGRID!Rate = Round(Rec!Rate, 4)
'''            End If
'''            RECGRID!ITEMCODE = Rec!ITEMCODE & ""
'''            RECGRID!LOT = Val(Rec!LOT & "")
'''            RECGRID!SAUDACODE = Rec!Sauda & ""
'''            RECGRID!bconfirm = Rec!CONFIRM & ""
'''            ItemRec.MoveFirst
'''            ItemRec.Find "ITEMCODE ='" & Rec!ITEMCODE & "'"
'''            SaudaCmbRec.MoveFirst
'''            SaudaCmbRec.Find "SAUDACODE ='" & Rec!Sauda & "'"
'''            RECGRID!QNTY = Rec!QTY
'''            RECGRID!ITEMName = ItemRec!ITEMName
'''            RECGRID!SAUDANAME = SaudaCmbRec!SAUDANAME
'''
'''            RECGRID!LInvNo = Val(Rec!INVNO & "")
'''            If Not IsNull(Rec!DATAIMPORT) Then
'''                If Rec!DATAIMPORT = True Then
'''                    RECGRID!DIMPORT = 1
'''                Else
'''                    RECGRID!DIMPORT = 0
'''                End If
'''            Else
'''                RECGRID!DIMPORT = 0
'''            End If
'''            RECGRID!CONTIME = IIf(IsNull(Rec!CONTIME), Time, Rec!CONTIME)
'''            RECGRID!USERID = Rec!USERID & ""
'''            Rec.MoveNext
'''           ' REC.MovePrevious
'''            If Len(Rec!concode & "") > Val(0) Then
'''                If Rec!PARTY <> Rec!concode Then
'''                    RECGRID!Code = Rec!PARTY & ""
'''                    Rec_Account.MoveFirst
'''                    Rec_Account.Find "AC_CODE='" & Rec!PARTY & "'"
'''                    RECGRID!NAME = Rec_Account!NAME
'''                    RECGRID!Rate = Round(Rec!Rate, 4)
'''                Else
'''                    RECGRID!concode = Rec!PARTY & ""
'''                    Rec_Account.MoveFirst
'''                    Rec_Account.Find "AC_CODE='" & Rec!PARTY & "'"
'''                    If Not Rec_Account.EOF Then
'''                        RECGRID!conName = IIf(IsNull(Rec_Account!NAME), "", Rec_Account!NAME)
'''                    Else
'''                        MsgBox "Seller Party Missing"
'''                    End If
'''                    RECGRID!Rate1 = Round(Rec!Rate, 4)
'''                End If
'''            Else
'''                RECGRID!concode = Rec!PARTY & ""
'''                Rec_Account.MoveFirst
'''                Rec_Account.Find "AC_CODE='" & Rec!PARTY & "'"
'''                RECGRID!conName = Rec_Account!NAME
'''                RECGRID!Rate1 = Round(Rec!Rate, 4)
'''
'''            End If
'''            RECGRID!SCONFIRM = Rec!CONFIRM
'''            Rec_Account.MoveFirst
'''
'''            RECGRID!RInvNo = Val(Rec!INVNO & "")
'''            RECGRID.Update
'''            Rec.MoveNext
'''        Wend
'''        If Saudacmb.BoundText <> "" Then
'''            DataGrid1.Columns(5).Locked = True
'''            Saudacmb.Enabled = False
'''        Else
'''            DataGrid1.Columns(5).Locked = False
'''        End If
'''
'''        Set DataGrid1.DataSource = RECGRID
'''        DataGrid1.ReBind
'''        DataGrid1.Col = 0
'''        Call DataGrid1_AfterColEdit(0)
'''        ITEMCMB.Enabled = False
'''        Saudacmb.Enabled = False
'''        vcDTP1.Enabled = False
'''        Combo1.Enabled = False
'''        MODIFY_REC = True
'''    End If
'''Else
'''    If Fb_Press = 1 Then
'''        DataGrid1.Col = 0
'''        DataGrid1.SetFocus
'''        If Saudacmb.BoundText <> "" Then
'''            DataGrid1.Columns(5).Locked = True
'''
'''        Else
'''            DataGrid1.Columns(5).Locked = False
'''        End If
'''        Saudacmb.Enabled = False
'''        ITEMCMB.Enabled = False
'''        vcDTP1.Enabled = False
'''        Combo1.Enabled = False
'''        Fb_Press = 1
'''        MODIFY_REC = True
'''    Else
'''        MsgBox "Record Does Not Exist For Selected Criteria.", vbInformation
'''        If Fb_Press = 2 Then
'''            DataGrid1.Col = 0
'''            DataGrid1.SetFocus
'''            If Saudacmb.BoundText <> "" Then
'''                DataGrid1.Columns(5).Locked = True
'''            Else
'''                DataGrid1.Columns(5).Locked = False
'''            End If
'''            Saudacmb.Enabled = False
'''            ITEMCMB.Enabled = False
'''            vcDTP1.Enabled = False
'''            Combo1.Enabled = False
'''            Fb_Press = 1
'''            MODIFY_REC = True
'''        Else
'''            Call CANCEL_REC
'''            Exit Function
'''        End If
'''    End If
'''End If
'''If Fb_Press = 3 Then
'''    If MsgBox("You are about to Delete all Contracts. Confirm Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
'''        Cnn.BeginTrans
'''        CNNERR = True
'''        REC_CTRM.MoveFirst
'''        While Not REC_CTRM.EOF
'''            MYSQL = "DELETE FROM CTR_D WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & REC_CTRM!consno & ""
'''            Cnn.Execute MYSQL
'''            MYSQL = "DELETE FROM CTR_R WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & REC_CTRM!consno & ""
'''            Cnn.Execute MYSQL
'''            Call Delete_Voucher(REC_CTRM!VOU_NO & "")
'''            MYSQL = "DELETE FROM CTR_M WHERE COMPCODE=" & GCompCode & " AND CONSNO=" & REC_CTRM!consno & ""
'''            Cnn.Execute MYSQL
'''            REC_CTRM.MoveNext
'''        Wend
'''        Cnn.CommitTrans
'''        Call CANCEL_REC
'''    End If
'''End If
'''End Function
'''
'''Private Sub Combo1_GotFocus()
'''    If FLOWDIR = 1 Then
'''        Set Rec = Nothing
'''        Set Rec = New ADODB.Recordset
'''        Rec.Open "SELECT COMPCODE FROM CTR_M WHERE COMPCODE=" & GCompCode & " AND SAUDA='" & DataCombo1.BoundText & "'", Cnn, adOpenForwardOnly, adLockReadOnly
'''        If Rec.EOF Then Sendkeys "%{DOWN}"
'''    End If
'''End Sub
'''Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If Shift = 1 Then
'''    flag = True
'''    End If
'''End Sub
'''
'''Private Sub Combo1_LostFocus()
'''If MODIFY_REC(vcDTP1.Value, Saudacmb.BoundText, ITEMCMB.BoundText, Combo1.text) Then
'''Else
'''    Combo1.SetFocus
'''End If
'''End Sub
'''Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'''
'''If Not RECGRID.EOF Then
'''    Set TEMPORARY = Nothing: Set TEMPORARY = New ADODB.Recordset
'''    Set TEMPORARY = RECGRID.Clone
'''    TEMPORARY.ActiveConnection = Nothing
'''    SQty = 0: BQty = 0: SRATE = 0: BRATE = 0: diffaMt = 0: SAVG = 0: BAVG = 0: LOT = 0: BTOT = 0: STOT = 0: TOTBRATE = 0: TOTSRate = 0: BUYAMT = 0: SELAMT = 0
'''    TEMPORARY.Filter = "SAUDACODE='" & RECGRID!SAUDACODE & "'"
'''    If Not TEMPORARY.EOF Then
'''        TEMPORARY.MoveFirst
'''        Label3.Caption = (RECGRID!SAUDACODE & "")
'''        While Not TEMPORARY.EOF
'''            If TEMPORARY!BUYSELL = "B" Then
'''                BRATE = BRATE + IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
'''                BQty = BQty + IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY)
'''                TOTBRATE = TOTBRATE + IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
'''                BUYAMT = BUYAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''                SELLAMT = SELLAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''            Else
'''                SRATE = SRATE + IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
'''                SQty = SQty + IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY)
'''                TOTSRate = TOTSRate + IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate)
'''                BUYAMT = BUYAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''                SELLAMT = SELLAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''            End If
'''            LOT = IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT)
'''            TEMPORARY.MoveNext
'''        Wend
'''    End If
'''    If Not SQty = 0 Then
'''        SAVG = TOTSRate / SQty
'''        STOT = SRATE * SQty * LOT
'''    End If
'''    If Not BQty = 0 Then
'''        BAVG = TOTBRATE / BQty
'''        BTOT = BRATE * BQty * LOT
'''    End If
'''    'Total Shree Caculation
'''    TEMPORARY.Filter = adFilterNone
'''    TEMPORARY.MoveFirst
'''    TOTBUYAMT = 0: TOTSELLAMT = 0
'''    While Not TEMPORARY.EOF
'''        If TEMPORARY!BUYSELL = "B" Then
'''            TOTBUYAMT = TOTBUYAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''            TOTSELLAMT = TOTSELLAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''        Else
'''            TOTBUYAMT = TOTBUYAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate1), 0, TEMPORARY!Rate1) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''            TOTSELLAMT = TOTSELLAMT + (IIf(IsNull(TEMPORARY!QNTY), 0, TEMPORARY!QNTY) * IIf(IsNull(TEMPORARY!Rate), 0, TEMPORARY!Rate) * IIf(IsNull(TEMPORARY!LOT), 0, TEMPORARY!LOT))
'''        End If
'''        TEMPORARY.MoveNext
'''    Wend
'''    Text2.text = TOTSELLAMT - TOTBUYAMT
'''    Text6.text = SELLAMT - BUYAMT
'''    Text1.text = BQty: Text4.text = SQty: Text7.text = BAVG: Text8.text = SAVG: Text9.text = BTOT: Text10.text = STOT:
'''End If
'''    Text7.text = Format(Text7.text, "0.00")
'''    Text8.text = Format(Text8.text, "0.00")
'''    Text9.text = Format(Text9.text, "0.00")
'''    Text10.text = Format(Text10.text, "0.00")
'''End Sub
'''
'''Private Sub ITEMCMB_GotFocus()
'''    Call ItemList
'''    Sendkeys "%{DOWN}"
'''End Sub
'''Private Sub DataCombo3_GotFocus()
'''    Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
'''    Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
'''    If Not Rec_Account.EOF Then Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
'''    Sendkeys "%{DOWN}"
'''    If DataGrid1.Col = 2 Or DataGrid1.Col = 1 Then
'''        DataGrid1.Col = 1
'''        DataGrid1.text = vbNullString
'''        Label2.Visible = True: Label2.Left = 1080
'''        DataCombo3.Left = Val(1080)
'''        DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
'''    ElseIf DataGrid1.Col = 7 Or DataGrid1.Col = 8 Then
'''        DataGrid1.Col = 8: DataCombo3.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
'''        DataCombo3.Left = DataGrid1.Columns(7).Left
'''        Label2.Visible = True: Label2.Left = DataGrid1.Columns(9).Left
'''    End If
'''    Sendkeys "%{DOWN}"
'''    Sendkeys "%{DOWN}"
'''End Sub
'''Private Sub DataCombo3_KeyPress(KeyAscii As Integer)
'''If KeyAscii = 13 Then
'''    If InStr(DataCombo3.BoundText, "'") Then
'''        DataCombo3.BoundText = Replace(DataCombo3.BoundText, "'", "", 1, Len(DataCombo3.BoundText))
'''    End If
'''    If DataGrid1.Col = 1 Or DataGrid1.Col = 2 Then
'''        If DataCombo3.BoundText <> "" Then
'''
'''            Rec_Account.Filter = adFilterNone
'''            Rec_Account.Filter = "AC_CODE='" & DataCombo3.BoundText & "'"
'''            If Rec_Account.EOF Then
'''                DataCombo3.BoundText = vbNullString
'''                Rec_Account.Filter = adFilterNone
'''            Else
'''                Rec_Account.Filter = adFilterNone
'''                RECGRID!Code = DataCombo3.BoundText
'''                RECGRID!NAME = DataCombo3.text
'''                RECGRID!USERID = LUserId
'''                DataGrid1.Col = 2
'''                DataGrid1.SetFocus
'''                DataCombo3.Visible = False: Label2.Visible = False
'''            End If
'''        End If
'''    ElseIf DataGrid1.Col = 8 Or DataGrid1.Col = 7 Then
'''        If DataCombo3.BoundText <> "" Then
'''            Rec_Account.Filter = adFilterNone
'''            Rec_Account.Filter = "AC_CODE='" & DataCombo3.BoundText & "'"
'''            If Rec_Account.EOF Then
'''                DataCombo3.BoundText = vbNullString
'''                Rec_Account.Filter = adFilterNone
'''            Else
'''                Rec_Account.Filter = adFilterNone
'''                RECGRID!concode = DataCombo3.BoundText
'''                RECGRID!conName = DataCombo3.text
'''                RECGRID!USERID = LUserId
'''                DataCombo3.Visible = False: Label2.Visible = False
'''                DataGrid1.Col = 8
'''                DataGrid1.SetFocus
'''            End If
'''        End If
'''    End If
'''ElseIf KeyAscii = 27 Then
'''    DataGrid1.SetFocus
'''    DataCombo3.Visible = False: Label2.Visible = False
'''ElseIf KeyAscii = 121 Then   'F3  NEW PARTY
'''    GETACNT.Show
'''    GETACNT.ZOrder
'''    GETACNT.Add_Record
'''ElseIf KeyAscii = 18 Then
'''    DataCombo3.Visible = True: DataCombo3.SetFocus
'''End If
'''
'''End Sub
'''
'''Private Sub DataCombo3_Validate(Cancel As Boolean)
'''    If DataCombo3.Visible = True Then
'''        Cancel = True
'''    Else
'''        Label2.Visible = False
'''    End If
'''End Sub
'''Public Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
'''    If Combo1.ListIndex < Val(0) Then Combo1.ListIndex = Val(0)
'''    If ColIndex = Val(0) Then
'''        If UCase(LEFT$(Trim(DataGrid1.text), 1)) = "S" Or UCase(LEFT$(Trim(DataGrid1.text), 1)) = "B" Then
'''            DataGrid1.text = LEFT$(UCase(DataGrid1.text), 1)
'''        Else
'''           DataGrid1.text = "B"
'''           DataGrid1.Col = 0
'''        End If
'''        DataGrid1.SetFocus
'''
'''    ElseIf ColIndex = Val(1) Then
'''''        MYSQL = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode  & " AND A.AC_CODE= '" & DataGrid1.Text & "'"
'''''        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, cnn
'''''        If TempRec.RecordCount > 0 Then
'''''            RECGRID!Code = TempRec!AC_CODE
'''''            RECGRID!Name = TempRec!Name
'''''        Else
'''''            DataCombo3.Visible = True
'''''            DataCombo3.SetFocus
'''''            RECGRID!Code = ""
'''''            DataGrid1.Col = 1
'''''            DataCombo3.SetFocus
'''''            'Exit Sub
'''''        End If
'''    ElseIf ColIndex = 4 Then
'''        If Val(RECGRID!Rate & "") > 0 Then
'''            If Val(Round(RECGRID!Rate1, 4) & "") = Val(0) Then RECGRID!Rate1 = Round(RECGRID!Rate, 4)
'''        Else
'''            If ColIndex = 5 Then
'''            Else
'''                MsgBox "Rate can not be zero.Please enter rate.", vbCritical
'''                DataGrid1.Col = 6: DataGrid1.SetFocus
'''            End If
'''        End If
'''    ElseIf ColIndex = Val(5) Then
'''        MYSQL = "SELECT I.ITEMCODE,I.ITEMNAME,I.LOT,S.SAUDACODE,S.SAUDANAME FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & GCompCode & " AND I.COMPCODE=S.COMPCODE AND I.ITEMCODE=S.ITEMCODE AND S.SAUDACODE= '" & DataGrid1.text & "'"
'''        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, Cnn
'''        If TempRec.RecordCount > 0 Then
'''            TempRec.MoveFirst
'''            TempRec.Find "SAUDACODE='" & DataGrid1.text & "'", , adSearchForward
'''            If Not TempRec.EOF Then
'''                RECGRID!ITEMCODE = TempRec!ITEMCODE
'''                RECGRID!ITEMName = TempRec!ITEMName
'''                RECGRID!SAUDACODE = TempRec!SAUDACODE
'''                RECGRID!ITEMName = TempRec!SAUDANAME
'''                RECGRID!LOT = TempRec!LOT
'''            Else
'''                RECGRID!SAUDACODE = ""
'''                DataGrid1.Col = 5
'''                Saudacombo.Visible = True
'''                Saudacombo.SetFocus
'''            End If
'''        End If
'''    End If
'''End Sub
'''Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If DataGrid1.Enabled = True Then
'''    If KeyCode = 13 And ((DataGrid1.Col = 9) Or (DataGrid1.Col = 4 And Saudacmb.BoundText <> "" And RECGRID!concode <> "") Or (DataGrid1.Col = 6 And RECGRID!concode <> "")) Then
'''        BCODE = RECGRID!Code
'''        BNAME = RECGRID!NAME
'''        ITEMCODE = RECGRID!ITEMCODE
'''        ITEMName = RECGRID!ITEMName
'''        LOT = RECGRID!LOT
'''        SAUDACODE = RECGRID!SAUDACODE
'''        SAUDANAME = RECGRID!SAUDANAME
'''        BUYSELL = RECGRID!BUYSELL
'''        scode = RECGRID!concode
'''        SNAME = RECGRID!conName
'''
'''        RECGRID.MoveNext
'''        If RECGRID.EOF Then
'''            RECGRID.AddNew
'''            RECGRID!Code = BCODE
'''            RECGRID!NAME = BNAME
'''            RECGRID!concode = scode
'''            RECGRID!conName = SNAME
'''            RECGRID!ITEMCODE = ITEMCODE
'''            RECGRID!ITEMName = ITEMName
'''            RECGRID!LOT = LOT
'''            RECGRID!BUYSELL = BUYSELL
'''            RECGRID!SAUDACODE = SAUDACODE
'''            RECGRID!SAUDANAME = SAUDANAME
'''            RECGRID!QNTY = 0
'''            RECGRID!Rate = Round(Val(0), 4)
'''            RECGRID!Rate1 = Round(Val(0), 4)
'''            RECGRID!bconfirm = Val(0)
'''            RECGRID!SCONFIRM = Val(0)
'''
'''            RECGRID!DIMPORT = 0
'''            RECGRID!USERID = LUserId & ""
'''            RECGRID!CONTIME = Time
'''            MYSQL = "SELECT MAX(CAST(CONNO AS INT)) AS CONNO FROM CTR_D WHERE COMPCODE =" & GCompCode & ""
'''            Set CONREC = Nothing
'''            Set CONREC = New ADODB.Recordset
'''            CONREC.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'''            If Fb_Press = 2 And CONFLAGE = False Then
'''                LConNo = Val(CONREC!CONNO & "") + Val(1)
'''            Else
'''                LConNo = LConNo + 1
'''            End If
'''                CONFLAGE = True
'''
'''                RECGRID!SrNo = LConNo 'RECGRID.AbsolutePosition
'''                RECGRID.Update
'''            End If
'''        DataGrid1.LeftCol = 0
'''        DataGrid1.Col = 0
'''    ElseIf DataGrid1.Col = 4 Or DataGrid1.Col = 11 Then
'''        If KeyCode = 13 Or KeyCode = 9 Then
'''            If Val(DataGrid1.text) = 0 Then
'''               MsgBox "Rate Cannot Be Zero", vbCritical
'''               DataGrid1.SetFocus
'''               Exit Sub
'''            End If
'''        End If
'''    ElseIf DataGrid1.Col = Val(7) And (KeyCode = 13 Or KeyCode = 9) Then
'''        MYSQL = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode & " AND A.AC_CODE= '" & DataGrid1.text & "'"
'''        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, Cnn
'''        If TempRec.RecordCount > 0 Then
'''            RECGRID!concode = TempRec!AC_CODE
'''            RECGRID!conName = TempRec!NAME
'''        Else
'''            RECGRID!concode = ""
'''            DataGrid1.Col = 7
'''            DataCombo3.Visible = True
'''            DataCombo3.SetFocus
'''            DataCombo3.SetFocus
'''            Exit Sub
'''        End If
'''
'''    ElseIf DataGrid1.Col = Val(1) And (KeyCode = 13 Or KeyCode = 9) Then
'''
'''        If DataGrid1.text <> "" Then
'''            DataGrid1.text = Replace(DataGrid1.text, "'", "")
'''        End If
'''        MYSQL = "SELECT A.AC_CODE,A.NAME FROM ACCOUNTM AS A WHERE A.COMPCODE=" & GCompCode & " AND A.AC_CODE= '" & DataGrid1.text & "'"
'''        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, Cnn
'''        If TempRec.RecordCount > 0 Then
'''            RECGRID!Code = TempRec!AC_CODE
'''            RECGRID!NAME = TempRec!NAME
'''            DataGrid1.Col = 2
'''            DataGrid1.SetFocus
'''        Else
'''            RECGRID!Code = ""
'''            DataGrid1.Col = 1
'''            DataCombo3.Visible = True
'''            DataCombo3.SetFocus
'''            DataCombo3.SetFocus
'''            Exit Sub
'''        End If
'''    ElseIf DataGrid1.Col = Val(5) And (KeyCode = 13 Or KeyCode = 9) Then
'''        MYSQL = "SELECT I.ITEMCODE,I.ITEMNAME,I.LOT,S.SAUDACODE,S.SAUDANAME FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE=" & GCompCode & " AND I.COMPCODE=S.COMPCODE AND I.ITEMCODE=S.ITEMCODE AND S.SAUDACODE= '" & DataGrid1.text & "'"
'''        Set TempRec = Nothing: Set TempRec = New ADODB.Recordset: TempRec.Open MYSQL, Cnn
'''        If TempRec.RecordCount > 0 Then
'''            TempRec.MoveFirst
'''            TempRec.Find "saudaCODE='" & DataGrid1.text & "'", , adSearchForward
'''            If Not TempRec.EOF Then
'''                RECGRID!ITEMCODE = TempRec!ITEMCODE
'''                RECGRID!ITEMName = TempRec!ITEMName
'''                RECGRID!SAUDACODE = TempRec!SAUDACODE
'''                RECGRID!ITEMName = TempRec!SAUDANAME
'''                RECGRID!LOT = TempRec!LOT
'''            Else
'''                RECGRID!SAUDACODE = ""
'''                DataGrid1.Col = 5
'''                Saudacombo.Visible = True
'''                Saudacombo.SetFocus
'''            End If
'''        Else
'''            RECGRID!SAUDACODE = ""
'''            DataGrid1.Col = 5
'''            Saudacombo.Visible = True
'''            Saudacombo.SetFocus
'''        End If
'''
'''    ElseIf KeyCode = 114 Then   'F3  NEW PARTY
'''        GETACNT.Show
'''        GETACNT.ZOrder
'''        GETACNT.Add_Record
'''    ElseIf KeyCode = 118 Then   ''F7 KEY
'''        RNO = InputBox("Enter the row number.", "Sauda")
'''        If Val(RNO) > Val(0) Then
'''            RECGRID.MoveFirst
'''            RECGRID.Find "SRNO=" & Val(RNO) & "", , adSearchForward
'''            If RECGRID.EOF Then
'''                MsgBox "Record not found.", vbCritical, "Error"
'''                RECGRID.MoveFirst
'''            End If
'''            DataGrid1.Col = 1
'''            DataGrid1.SetFocus
'''        End If
'''    ElseIf KeyCode = 46 And Shift = 2 Then
'''        RECGRID.Delete
'''        If RECGRID.RecordCount = 0 Then
'''            RECGRID.AddNew
'''            LConNo = LConNo + 1
'''            RECGRID!SrNo = LConNo 'RECGRID.RecordCount
'''            If Combo1.ListIndex = Val(1) Then
'''                RECGRID!BRATE = Round(Val(Text3.text), 4)
'''                RECGRID!SRATE = Round(Val(Text3.text), 4)
'''                RECGRID!USERID = LUserId
'''            End If
'''            RECGRID.Update
'''        End If
'''        Call DataGrid1_AfterColEdit(0)
'''    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 0) Then
'''        If UCase(LEFT$(Trim(DataGrid1.text), 1)) = "S" Or UCase(LEFT$(Trim(DataGrid1.text), 1)) = "B" Then
'''            DataGrid1.text = LEFT$(UCase(DataGrid1.text), 1)
'''        Else
'''           DataGrid1.text = "B"
'''           DataGrid1.Col = 0
'''           DataGrid1.SetFocus
'''        End If
'''    ElseIf (KeyCode = 13 Or KeyCode = 9) And ((DataGrid1.Col = 1) Or (DataGrid1.Col = 7)) Then
'''        If Len(Trim(DataGrid1.text)) < 1 Then
'''            DataCombo3.Visible = True
'''            DataCombo3.SetFocus
'''        Else
'''        End If
'''    ElseIf (KeyCode = 13 Or KeyCode = 9) And (DataGrid1.Col = 5) Then
'''        If Len(Trim(DataGrid1.text)) < 1 Then
'''            Saudacombo.Visible = True
'''            Saudacombo.SetFocus
'''        End If
'''    ElseIf (KeyCode = 13 Or KeyCode = 9) And DataGrid1.Col = 9 Then
'''        If Val(DataGrid1.text & "") <= 0 Then
'''            MsgBox "Rate can not be zero.Please enter Rate.", vbCritical
'''            DataGrid1.Col = 3: DataGrid1.SetFocus
'''        End If
'''    ElseIf KeyCode = 27 Then
'''        KeyCode = 0
'''    End If
'''    DataGrid1.CurrentCellVisible = True
'''End If
'''End Sub
'''
'''Private Sub EXCMB_GotFocus()
'''    Sendkeys "%{DOWN}"
'''End Sub
'''
'''Private Sub EXCMB_Validate(Cancel As Boolean)
'''If excmb.BoundText = "" Then
'''    Cancel = True
'''    Exit Sub
'''End If
'''    Call ItemList
'''End Sub
'''
'''
'''Private Sub Form_KeyPress(KeyAscii As Integer)
'''If KeyAscii = 13 Then Sendkeys "{tab}"
'''End Sub
'''
'''Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'''    If KeyCode = 13 Then
'''        On Error Resume Next
'''        If Me.ActiveControl.NAME = "vcDTP1" Then
'''            Sendkeys "{tab}"
'''        End If
'''    End If
'''End Sub
'''Private Sub Form_Load()
'''    Set DataGrid1.DataSource = RECGRID: DataGrid1.Refresh
'''    Call ClearFormFn(SINGLECONTRACT)
'''    Frame1.Enabled = False
'''    Call CANCEL_REC
'''    LDataImport = 0
'''    DataCombo3.Top = Val(2055): DataCombo3.Left = Val(1080)
'''    MYSQL = "SELECT ITEMCODE,ITEMNAME AS ITEMNAME,Lot FROM ITEMMAST WHERE COMPCODE=" & GCompCode & " ORDER BY ITEMCODE"
'''    Set Rec = Nothing: Set Rec = New ADODB.Recordset: Rec.Open MYSQL, Cnn, adOpenKeyset, adLockReadOnly
'''    If Not Rec.EOF Then
'''        QACC_CHANGE = False: Set Rec_Account = Nothing: Set Rec_Account = New ADODB.Recordset
'''        Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
'''        If Not Rec_Account.EOF Then Set DataCombo3.RowSource = Rec_Account: DataCombo3.BoundColumn = "AC_CODE": DataCombo3.ListField = "NAME"
'''        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
'''    Else
'''        Call Get_Selection(12)
'''    End If
'''End Sub
'''Private Sub Form_Paint()
'''    Me.BackColor = GETMAIN.BackColor
'''    If QACC_CHANGE Then
'''        QACC_CHANGE = False: Set Rec_Account = Nothing
'''        Set Rec_Account = New ADODB.Recordset
'''        Rec_Account.Open "SELECT AC_CODE, NAME FROM ACCOUNTD WHERE COMPCODE=" & GCompCode & " AND gcode in (12,14) ORDER BY NAME ", Cnn, adOpenKeyset, adLockReadOnly
'''        If Not Rec_Account.EOF Then
'''            Set DataCombo3.RowSource = Rec_Account
'''            DataCombo3.BoundColumn = "AC_CODE"
'''            DataCombo3.ListField = "NAME"
'''        Else
'''            MsgBox "Please create customer account", vbInformation
'''            Call Get_Selection(12)
'''        End If
'''    End If
'''    If Fb_Press > 0 Then Call Get_Selection(Fb_Press)
'''End Sub
'''Private Sub headcmb_GotFocus()
'''    Sendkeys "%{DOWN}"
'''End Sub
'''Private Sub headcmb_Validate(Cancel As Boolean)
'''If headcmb.BoundText = "" Then Cancel = True
'''End Sub
'''Private Sub partycmb_Validate(Cancel As Boolean)
'''If PartyCmb.BoundText <> "" Then
'''    RECGRID!Code = PartyCmb.BoundText
'''    RECGRID!NAME = PartyCmb.text
'''    DataGrid1.Columns(1).Locked = True
'''Else
'''    DataGrid1.Columns(1).Locked = False
'''End If
'''If ITEMCMB.BoundText <> "" Then
'''    RECGRID!ITEMCODE = ITEMCMB.BoundText
'''    RECGRID!ITEMName = ITEMCMB.text
'''    DataGrid1.Columns(12).Locked = True
'''Else
'''    DataGrid1.Columns(12).Locked = False
'''End If
'''
'''End Sub
'''Private Sub Saudacmb_GotFocus()
'''Call SaudaList
'''End Sub
'''
'''Private Sub Saudacmb_LostFocus()
'''   ' resu
'''    Combo1.Enabled = True
'''    Combo1.SetFocus
'''End Sub
'''
'''Private Sub Saudacmb_Validate(Cancel As Boolean)
'''If Saudacmb.BoundText <> "" Then
'''    SaudaCmbRec.Filter = adFilterNone
'''    SaudaCmbRec.Filter = "SaudaCode='" & Saudacmb.BoundText & "'"
'''    If SaudaCmbRec.EOF Then
'''        Saudacmb.BoundText = ""
'''        Cancel = True
'''    Else
'''        SaudaCmbRec.Filter = adFilterNone
'''        RECGRID!SAUDACODE = Saudacmb.BoundText
'''        RECGRID!SAUDANAME = Saudacmb.text
'''        DataGrid1.Col = 0
'''        DataGrid1.SetFocus
'''        Saudacombo.Visible = False
'''        DataGrid1.Columns(12).Locked = True
'''        DataGrid1.SetFocus
'''    End If
'''End If
'''If GRateSlab = 1 Then
'''    Combo1.Enabled = False
'''    MYSQL = "SELECT  * FROM  CTR_R WHERE COMPCODE  =" & GCompCode & "  AND CONDATE ='" & Format(vcDTP1.Value, "yyyy/mm/dd") & "'"
'''    Set NRec1 = Nothing
'''    Set NRec1 = New ADODB.Recordset
'''    NRec1.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'''    If NRec1.EOF Then
'''        Label8.Visible = False
'''        Text13.Visible = False
'''        Text13.TabStop = False
'''    Else
'''        Label8.Visible = True
'''        Text13.Visible = True
'''        Text13.SetFocus
'''        Text13.SetFocus
'''    End If
'''End If
'''End Sub
'''
'''Private Sub SAUDACOMBO_GotFocus()
'''    If ITEMCMB.BoundText <> "" Then
'''        MYSQL = "SELECT S.SAUDANAME,S.SAUDACODE FROM SAUDAMAST AS S WHERE S.COMPCODE =" & GCompCode & " AND S.ITEMCODE='" & ITEMCMB.BoundText & "' AND MATURITY>= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,MATURITY"
'''    Else
'''        MYSQL = "SELECT SAUDANAME,SAUDACODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & "  AND MATURITY>= '" & Format(vcDTP1.Value, "YYYY/MM/DD") & "' ORDER BY ITEMCODE,MATURITY"
'''    End If
'''    Set SaudaCmbRec = Nothing
'''    Set SaudaCmbRec = New ADODB.Recordset
'''    SaudaCmbRec.Open MYSQL, Cnn
'''    If Not SaudaCmbRec.EOF Then
'''        Set Saudacombo.RowSource = SaudaCmbRec
'''        Saudacombo.ListField = "SAUDANAME"
'''        Saudacombo.BoundColumn = "SAUDACODE"
'''    End If
'''    Sendkeys "%{DOWN}"
'''    Saudacombo.Left = DataGrid1.Columns(5).Left
'''    Saudacombo.Top = DataGrid1.Top + Val(DataGrid1.RowTop(DataGrid1.Row))
'''    Sendkeys "%{DOWN}"
'''    Sendkeys "%{DOWN}"
'''End Sub
'''
'''Private Sub Saudacombo_KeyPress(KeyAscii As Integer)
'''If KeyAscii = 13 Then
'''    If Saudacombo.BoundText <> "" Then
'''        RECGRID!SAUDACODE = Saudacombo.BoundText
'''        RECGRID!SAUDANAME = Saudacombo.text
'''        MYSQL = "SELECT I.ITEMCODE,I.ITEMNAME,I.LOT FROM ITEMMAST AS I,SAUDAMAST AS S WHERE I.COMPCODE= '" & GCompCode & "' AND S.SAUDACODE='" & Saudacombo.BoundText & "'AND S.ITEMCODE=I.ITEMCODE AND I.COMPCODE=S.COMPCODE "
'''        Set TempRec = Nothing
'''        Set TempRec = New ADODB.Recordset
'''        TempRec.Open MYSQL, Cnn
'''        If Not TempRec.EOF Then
'''            RECGRID!ITEMCODE = TempRec!ITEMCODE
'''            RECGRID!ITEMName = TempRec!ITEMName
'''            RECGRID!LOT = TempRec!LOT
'''        End If
'''        DataGrid1.Col = 6
'''        DataGrid1.SetFocus
'''        Saudacombo.Visible = False
'''    End If
'''End If
'''
'''End Sub
'''Private Sub Saudacombo_Validate(Cancel As Boolean)
'''If Saudacombo.Visible = True Then
'''    Cancel = True
'''End If
'''End Sub
'''
'''Private Sub Text3_GotFocus()
'''    FLOWDIR = 0: Text3.SelLength = Len(Text3.text)
'''End Sub
'''Sub RecSet()
'''    Set RECGRID = Nothing
'''    Set RECGRID = New ADODB.Recordset
'''    RECGRID.Fields.Append "SRNO", adDouble, , adFldIsNullable
'''    RECGRID.Fields.Append "BUYSELL", adVarChar, 1, adFldIsNullable
'''    RECGRID.Fields.Append "CODE", adVarChar, 6, adFldIsNullable
'''    RECGRID.Fields.Append "NAME", adVarChar, 150, adFldIsNullable
'''    RECGRID.Fields.Append "QNTY", adDouble, , adFldIsNullable
'''    RECGRID.Fields.Append "CONCODE", adVarChar, 6, adFldIsNullable
'''    RECGRID.Fields.Append "CONNAME", adVarChar, 150, adFldIsNullable
'''    RECGRID.Fields.Append "RATE", adDouble, , adFldIsNullable
'''    RECGRID.Fields.Append "LOT", adDouble, , adFldIsNullable
'''    RECGRID.Fields.Append "ITEMCODE", adVarChar, 20, adFldIsNullable
'''    RECGRID.Fields.Append "ITEMNAME", adVarChar, 50, adFldIsNullable
'''    RECGRID.Fields.Append "SAUDACODE", adVarChar, 50, adFldIsNullable
'''    RECGRID.Fields.Append "SAUDANAME", adVarChar, 50, adFldIsNullable
'''    RECGRID.Fields.Append "RATE1", adDouble, , adFldIsNullable
'''    RECGRID.Fields.Append "LInvNo", adDouble, , adFldIsNullable
'''    RECGRID.Fields.Append "RInvNo", adDouble, , adFldIsNullable
'''    RECGRID.Fields.Append "DImport", adInteger, , adFldIsNullable
'''    RECGRID.Fields.Append "CONTIME", adVarChar, 15, adFldIsNullable
'''    RECGRID.Fields.Append "UserId", adVarChar, 30, adFldIsNullable
'''    RECGRID.Fields.Append "BCONFIRM", adInteger, 30, adFldIsNullable
'''    RECGRID.Fields.Append "SCONFIRM", adInteger, 30, adFldIsNullable
'''    RECGRID.Fields.Append "CONTYPE", adVarChar, 1, adFldIsNullable
'''    RECGRID.Fields.Append "ORDER_NO", adVarChar, 30, adFldIsNullable
'''    RECGRID.Fields.Append "TRADE_NO", adVarChar, 30, adFldIsNullable
'''    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic
'''End Sub
'''
'''Private Sub Text13_Validate(Cancel As Boolean)
'''    If GRegNo2 = EncryptNEW(Text13.text, 13) Then
'''        Combo1.Enabled = True
'''        Combo1.SetFocus
'''    Else
'''        MsgBox "Invalid Password No Modificatiobn Allowed"
'''        Cancel = True
'''    End If
'''End Sub
'''Sub Delete_Voucher(VOU_NO As String)
'''    Cnn.Execute "DELETE FROM VCHAMT  WHERE COMPCODE=" & GCompCode & " AND VOU_NO='" & VOU_NO & "'"
'''    Cnn.Execute "DELETE FROM VOUCHER WHERE COMPCODE=" & GCompCode & " AND VOU_NO='" & VOU_NO & "'"
'''End Sub
'''Sub VISIBLE_IMAGE(SORT_ORDER As Byte)
'''    If SORT_ORDER = 1 Then
'''        Image1(0).Visible = False
'''        Image1(1).Visible = True
'''    Else
'''        Image1(0).Visible = True
'''        Image1(1).Visible = False
'''    End If
'''End Sub
'''Function GetCloseRate() As Boolean
'''     Set Rec_Sauda = Nothing: Set Rec_Sauda = New ADODB.Recordset
'''     Rec_Sauda.Open "SELECT ITEMCODE,SAUDACODE FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND SAUDACODE='" & Saudacmb.BoundText & "' AND MATURITY>= '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", Cnn, adOpenForwardOnly, adLockReadOnly
'''     If Rec_Sauda.EOF Then
'''         MsgBox "Invalid SAUDA code.", vbExclamation, "Error"
'''         GetCloseRate = False
'''     Else
'''         GetCloseRate = True
'''         Set REC_CloRate = Nothing: Set REC_CloRate = New ADODB.Recordset
'''         REC_CloRate.Open "SELECT CloseRate,DataImport FROM CTR_R WHERE COMPCODE=" & GCompCode & " AND SAUDA='" & Text2.text & "' AND CONDATE  =  '" & Format(vcDTP1.Value, "yyyy/MM/dd") & "'", Cnn, adOpenForwardOnly, adLockReadOnly
'''         If Not REC_CloRate.EOF Then
'''            Text3.text = Format(REC_CloRate!CLOSERATE, "0.0000")
'''        End If
'''         Text2.text = Rec_Sauda!SAUDACODE
'''         DataCombo1.BoundText = CStr(Text2.text)
'''         ITEMCMB.BoundText = Rec_Sauda!ITEMCODE
'''    End If
'''End Function
