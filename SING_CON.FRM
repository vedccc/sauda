VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0102BD99-4A7D-11D3-AC0E-00C026A22F30}#5.1#0"; "DATECTL.OCX"
Begin VB.Form SING_CON 
   Caption         =   "Contract Entry"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   7680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   11655
      Begin VB.TextBox Text18 
         Height          =   360
         Left            =   1440
         TabIndex        =   22
         Text            =   "Text18"
         Top             =   6360
         Width           =   10095
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "SING_CON.frx":0000
         Left            =   7920
         List            =   "SING_CON.frx":000D
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   2693
         Width           =   1575
      End
      Begin VB.TextBox Text17 
         Height          =   360
         Left            =   7920
         TabIndex        =   21
         Text            =   "Text17"
         Top             =   5625
         Width           =   3615
      End
      Begin VB.TextBox Text16 
         Height          =   360
         Left            =   1440
         TabIndex        =   20
         Text            =   "Text16"
         Top             =   5625
         Width           =   5055
      End
      Begin VB.TextBox Text15 
         Height          =   360
         Left            =   7920
         TabIndex        =   19
         Text            =   "Text15"
         Top             =   4890
         Width           =   3615
      End
      Begin VB.TextBox Text14 
         Height          =   360
         Left            =   1440
         TabIndex        =   18
         Text            =   "Text14"
         Top             =   4890
         Width           =   5055
      End
      Begin VB.TextBox Text13 
         Height          =   360
         Left            =   10440
         TabIndex        =   17
         Text            =   "Text13"
         Top             =   4155
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Rs."
         Top             =   4155
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   360
         Left            =   7920
         TabIndex        =   15
         Text            =   "Text11"
         Top             =   4155
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Height          =   360
         Left            =   5400
         TabIndex        =   14
         Text            =   "Text10"
         Top             =   4155
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   360
         Left            =   3480
         TabIndex        =   13
         Text            =   "Text9"
         Top             =   4155
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   360
         Left            =   1440
         TabIndex        =   12
         Text            =   "Text8"
         Top             =   4155
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   360
         Left            =   8400
         TabIndex        =   11
         Text            =   "Text7"
         Top             =   3420
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   7920
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "T6"
         Top             =   3420
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Left            =   1440
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   3420
         Width           =   5055
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Left            =   8400
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   7920
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "T3"
         Top             =   1950
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1950
         Width           =   5055
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   7920
         TabIndex        =   2
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vcDateTimePicker.vcDTP vcDTP1 
         Height          =   360
         Left            =   5040
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   37860.8625462963
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1440
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         Top             =   1215
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   2685
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
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
         Caption         =   "Remark"
         Height          =   225
         Index           =   19
         Left            =   240
         TabIndex        =   43
         Top             =   6428
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   225
         Index           =   18
         Left            =   7080
         TabIndex        =   42
         Top             =   2753
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         Height          =   225
         Index           =   17
         Left            =   7080
         TabIndex        =   41
         Top             =   5693
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quality"
         Height          =   225
         Index           =   16
         Left            =   240
         TabIndex        =   40
         Top             =   5693
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
         Height          =   225
         Index           =   15
         Left            =   7080
         TabIndex        =   39
         Top             =   4958
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del. Period"
         Height          =   225
         Index           =   14
         Left            =   240
         TabIndex        =   38
         Top             =   4958
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax"
         Height          =   225
         Index           =   13
         Left            =   9840
         TabIndex        =   37
         Top             =   4223
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scale"
         Height          =   225
         Index           =   12
         Left            =   7080
         TabIndex        =   36
         Top             =   4223
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   225
         Index           =   11
         Left            =   4800
         TabIndex        =   35
         Top             =   4223
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         Height          =   225
         Index           =   10
         Left            =   2880
         TabIndex        =   34
         Top             =   4223
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         Height          =   225
         Index           =   9
         Left            =   240
         TabIndex        =   33
         Top             =   4223
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brokrage"
         Height          =   225
         Index           =   8
         Left            =   7080
         TabIndex        =   32
         Top             =   3488
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buy. Through"
         Height          =   225
         Index           =   7
         Left            =   240
         TabIndex        =   31
         Top             =   3488
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer Name"
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Top             =   2753
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brokrage"
         Height          =   225
         Index           =   5
         Left            =   7080
         TabIndex        =   29
         Top             =   2018
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sel. Through"
         Height          =   225
         Index           =   4
         Left            =   240
         TabIndex        =   28
         Top             =   2018
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seller Name"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   1283
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         Height          =   225
         Index           =   2
         Left            =   7080
         TabIndex        =   26
         Top             =   548
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract Date"
         Height          =   225
         Index           =   1
         Left            =   3840
         TabIndex        =   25
         Top             =   548
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract No."
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   548
         Width           =   1020
      End
   End
End
Attribute VB_Name = "SING_CON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FB_PRESS As Byte
Dim REC As ADODB.Recordset
Sub ADD_REC()
    Frame1.Enabled = True
    Call Get_Selection(1)
    If FB_PRESS = 1 Then
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT MAX(CON_NO) FROM CONTRACT", CNN, adOpenForwardOnly, adLockReadOnly
        Text1.Text = REC.Fields(0) + Val(1)
        Set REC = Nothing
    Else
        Text1.SetFocus
    End If
End Sub
Sub SAVE_REC()
    If FB_PRESS = 1 Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!CON_NO = Val(Text1.Text)
    End If
    With Adodc1.Recordset
        !CON_DT = Format(vcDTP1.Value, "dd/MM/yyyy")
        !BUYER = DataCombo3.BoundText
        !SELLER = DataCombo2.BoundText
        !THRBUY = Text5.Text
        !THRSEL = Text2.Text
        !ITCODE = DataCombo1.BoundText
        !LOW = Val(Text8.Text)
        !HIGH = Val(Text9.Text)
        !Scale = Val(Text11.Text)
        !Rate = Val(Text10.Text)
        !DEL_PER = Text14.Text
        !DEL_PLC = Text15.Text
        !QUALITY = Text16.Text
        !PAYMENT = Text17.Text
        !REM1 = Text18.Text
        !TAX = Val(Text13.Text)
        !BBROK = Val(Text7.Text)
        !SBROK = Val(Text4.Text)
        !BBROKTYPE = Text6.Text
        !SBROKTYPE = Text3.Text
        !Type = Combo1.Text
        .Update
    End With
    Call CANCEL_REC
End Sub
Sub CANCEL_REC()
    Call ClearFormFn(GETCont)
    Call Get_Selection(10)
    Text12.Text = "Rs."
    Text3.Text = "Q"
    Text6.Text = "Q"
    Frame1.Enabled = False
End Sub
Sub MODIFY_REC()
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find "CON_NO=" & Val(Text1.Text) & "", , adSearchForward
    If Adodc1.Recordset.EOF Then
        MsgBox "Contract Number not found.", vbExclamation, "Error"
        Exit Sub
    End If

    With Adodc1.Recordset
        vcDTP1.Value = !CON_DT
        DataCombo3.BoundText = !BUYER
        DataCombo2.BoundText = !SELLER
        Text5.Text = !THRBUY
        Text2.Text = !THRSEL
        DataCombo1.BoundText = !ITCODE
        Text8.Text = Format(Val(!LOW), "0.00")
        Text9.Text = Format(Val(!HIGH), "0.00")
        Text11.Text = Format(Val(!Scale), "0.00")
        Text10.Text = Format(Val(!Rate), "0.00")
        Text14.Text = !DEL_PER
        Text15.Text = !DEL_PLC
        Text16.Text = !QUALITY
        Text17.Text = !PAYMENT
        Text18.Text = !REM1
        Text13.Text = Format(Val(!TAX), "0.00")
        Text7.Text = Format(Val(!BBROK), "0.00")
        Text4.Text = Format(Val(!SBROK), "0.00")
        Text6.Text = !BBROKTYPE
        Text3.Text = !SBROKTYPE
        Combo1.Text = !Type & ""
    End With
    
    If FB_PRESS = 3 Then
        If MsgBox("You are about to delete this record. Confirm Delete?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            Adodc1.Recordset.Delete
            Call CANCEL_REC
        End If
    End If
End Sub
Private Sub Combo1_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo1_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo2_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub DataCombo3_GotFocus()
    SendKeys "%{DOWN}"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
    Call CANCEL_REC
    
    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open "SELECT ITEMCODE, ITEMCODE+', '+STRING(6-LEN(ITEMCODE),' ')+ITEMNAME AS ITEMNAME FROM ITEMMAST ORDER BY ITEMCODE", CNN, adOpenKeyset, adLockReadOnly

    Set DataCombo1.RowSource = REC
    DataCombo1.BoundColumn = "ITEMCODE"
    DataCombo1.ListField = "ITEMNAME"

    Set REC = Nothing
    Set REC = New ADODB.Recordset
    REC.Open "SELECT AC_CODE, AC_CODE+', '+STRING(6-LEN(AC_CODE),' ')+NAME AS NAME FROM ACCOUNTD ORDER BY AC_CODE", CNN, adOpenKeyset, adLockReadOnly

    Set DataCombo2.RowSource = REC
    DataCombo2.BoundColumn = "AC_CODE"
    DataCombo2.ListField = "NAME"

    Set DataCombo3.RowSource = REC
    DataCombo3.BoundColumn = "AC_CODE"
    DataCombo3.ListField = "NAME"

    Adodc1.ConnectionString = CNN
    Adodc1.RecordSource = "SELECT * FROM CONTRACT ORDER BY CON_NO"
    Adodc1.Refresh
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    If FB_PRESS = 1 Then
        Set REC = Nothing
        Set REC = New ADODB.Recordset
        REC.Open "SELECT * FROM CONTRACT WHERE CON_NO=" & Text1.Text & "", CNN, adOpenForwardOnly, adLockReadOnly
        If REC.EOF Then
            Set REC = Nothing
            Exit Sub
        Else
            MsgBox "Number already exists!", vbExclamation, "Error"
            Cancel = True
        End If
    Else
        Call MODIFY_REC
    End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text10_Validate(Cancel As Boolean)
    Text10.Text = Format(Text10.Text, "0.00")
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text11_Validate(Cancel As Boolean)
    Text11.Text = Format(Text11.Text, "0.00")
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text13_Validate(Cancel As Boolean)
    Text13.Text = Format(Text13.Text, "0.00")
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Not KeyAscii >= 80 Or Not KeyAscii <= 81 Then KeyAscii = 0
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    Text4.Text = Format(Text4.Text, "0.00")
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If Not KeyAscii >= 80 Or Not KeyAscii <= 81 Then KeyAscii = 0
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    Text7.Text = Format(Text7.Text, "0.00")
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    Text8.Text = Format(Text8.Text, "0.00")
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    KeyAscii = NUMBERChk(KeyAscii)
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    Text9.Text = Format(Text9.Text, "0.00")
End Sub
