VERSION 5.00
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form GENSETL 
   Caption         =   "Generate Settlement"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   7815
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   2280
         Width           =   1335
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   375
         Left            =   3240
         TabIndex        =   0
         Top             =   1230
         Width           =   1695
         _ExtentX        =   2990
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
         Value           =   37960.3556828704
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upto Date"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   4
         Top             =   1320
         Width           =   735
      End
   End
End
Attribute VB_Name = "GENSETL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC As ADODB.Recordset

Private Sub Command1_Click()
    Call BILL_GENERATION
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
