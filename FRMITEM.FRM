VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmitem 
   Caption         =   "Itemwise Suada"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3660
   ScaleWidth      =   6885
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2160
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   576
      ImageHeight     =   316
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmitem.frx":0000
            Key             =   "Sauda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4260
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmitem.frx":85554
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmitem.frx":10AAA8
            Key             =   "Sauda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REC As ADODB.Recordset
Dim rec2 As ADODB.Recordset
Dim SQL As String
Dim sql2 As String
Private Sub Form_Load()
Set REC = Nothing
Set rec2 = Nothing
Set REC = New ADODB.Recordset
Set rec2 = New ADODB.Recordset
SQL = "Select * from ItemMast Order by ItemName"

REC.Open SQL, CNN, adOpenStatic, adLockOptimistic
REC.MoveFirst
Do While Not REC.EOF
    TreeView1.Nodes.Add , , "Item" & REC("Itemcode"), REC("ItemName"), "Item", "Item"
    REC.MoveNext
Loop
sql2 = "select * from saudamast order by saudaname"
rec2.Open sql2, CNN, adOpenStatic, adLockOptimistic
rec2.MoveFirst
Do While Not rec2.EOF
    TreeView1.Nodes.Add "Item" & rec2("itemcode"), tvwChild, "Saud" & rec2("Saudacode"), rec2("SaudaName"), 2, 2
    rec2.MoveNext
Loop

End Sub

Private Sub Form_Resize()
On Error Resume Next
TreeView1.Height = Me.ScaleHeight - (TreeView1.Top * 2)
'TreeView1.Width = Me.ScaleWidth - (TreeView1.Left * 2)
ListView1.Height = Me.ScaleHeight - (TreeView1.Top * 2)
ListView1.Width = Me.ScaleWidth - (ListView1.Left - TreeView1.Left)
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim recfilter As ADODB.Recordset
Dim msql As String
Dim mitemcode As String
Set recfilter = Nothing
Set recfilter = New ADODB.Recordset
ListView1.ListItems.clear
'rec2.Filter = "Saudacode = " & Mid$(Node.Key, 5)
'Set recfilter = rec2.Clone
mitemcode = LTrim(Mid$(Node.Key, 5))
msql = "select * from saudamast where ITEMCODE = '" & mitemcode & "'"

recfilter.Open msql, CNN, adOpenStatic, adLockReadOnly
MsgBox msql & recfilter.RecordCount
Do While Not recfilter.EOF
    ListView1.ListItems.Add , "sauda" & recfilter("saudacode"), recfilter("saudaname"), "Sauda", "Sauda"
    recfilter.MoveNext
Loop
recfilter.close
Set recfilter = Nothing
End Sub
