VERSION 5.00
Begin VB.Form frmstartprocess 
   Caption         =   "Start Process"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   5160
         Top             =   2160
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton CmdUpd 
         Caption         =   "Start Process"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Interval in mins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Import"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmstartprocess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim I As Integer
Private Sub Label13_Click()
End Sub

Private Sub CmdUpd_Click()
    If Text1.text = "" Then
        flag = False
    ElseIf IsNumeric(Text1.text) Then
        flag = True
        DoEvents
        Label4.Caption = " Schedule executes after " & Text1.text & " mins."
        DoEvents
        
        ''call frmdata
        frmdata.Check39.Value = 1
        frmdata.OkCmd_Click

    End If
    
End Sub

Private Sub Form_Load()
    flag = False
End Sub

Private Sub Timer1_Timer()

'1 min = 60000 ms

If flag Then
    I = I + 1
    Dim L As Integer
    'L = Round((500 - I) / 10)
    'DoEvents
    'Label4.Caption = " Get data in next " & L & " sec."
    DoEvents
    If I = Val(Text1.text) Then
        'Timer1.Enabled = False
       'Call CmdUpd_Click
        'Call MCXEXFILE
        'Call Get_LiveRate("01/01/1900", "", 0)
        'Call okcmdclick
        
        frmdata.Check39.Value = 1
        frmdata.OkCmd_Click
        
        I = 0
        DoEvents
        Label4.Caption = " Schedule executes after " & Text1.text & " mins."
        DoEvents
        'MsgBox "TIMER DONE"
        'Timer1.Enabled = True
    Else
        DoEvents
        Label4.Caption = " Schedule executes after " & CStr(Val(Text1.text) - I) & " mins."
        DoEvents
    End If
End If

End Sub
