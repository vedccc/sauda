VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_Brok_Slab 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Brokerage Slab"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   15030
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   4200
      TabIndex        =   8
      Top             =   2160
      Width           =   12615
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   480
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   11033
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "BROKTYPE"
            Caption         =   "BrokType"
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
            Caption         =   "BrokRate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "BROKRATE2"
            Caption         =   "BrokRate2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "MINRATE"
            Caption         =   "MinRate"
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
         BeginProperty Column06 
            DataField       =   "MBROKTYPE"
            Caption         =   "MBrokType"
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
            DataField       =   "MBROKRATE"
            Caption         =   "MBrokRate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "MBROKRATE2"
            Caption         =   "MBrokRate2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   12615
      Begin VB.TextBox Text2 
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
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text1 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Slab Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Slab Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   7560
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   13335
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   4194368
      ListField       =   "Author"
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7860
      Left            =   4080
      Top             =   1080
      Width           =   12915
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      BorderWidth     =   12
      Height          =   7860
      Index           =   0
      Left            =   120
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "Frm_Brok_Slab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fb_Press As Byte
Dim LSLABCODE  As Integer
Dim TRec As ADODB.Recordset
Dim RECGRID As ADODB.Recordset

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And DataGrid1.Col = 2 Then ' BROKTYPE
        Combo1.Visible = True: Combo1.SetFocus
    ElseIf KeyCode = 13 And DataGrid1.Col = 6 Then 'TRANTYPE
        Combo1.Visible = True: Combo1.SetFocus
        
    ElseIf KeyCode = 13 Then
     '   SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Load()
'
'DataList1.SetFocus
Text1.text = vbNullString
Text2.text = vbNullString
Frame1.Enabled = False
Frame2.Enabled = False
Call Get_Selection(10)
MYSQL = "SELECT DISTINCT SLABCODE,SLABNAME FROM BROKSLAB  WHERE COMPCODE =" & GCompCode & " ORDER BY SLABCODE"
Set SaudaRec = Nothing
Set SaudaRec = New ADODB.Recordset
SaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
Set DataList1.RowSource = SaudaRec
    DataList1.BoundColumn = "SLABCODE"
    DataList1.ListField = "SLABNAME"
End Sub

Sub RecSet()
Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "EXCODE", adVarChar, 10, adFldIsNullable
    RECGRID.Fields.Append "INSTTYPE", adVarChar, 3, adFldIsNullable
    RECGRID.Fields.Append "BROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "BROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BROKRATE2", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MINRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MBROKTYPE", adVarChar, 50, adFldIsNullable
    RECGRID.Fields.Append "MBROKRATE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "MBROKRATE2", adDouble, , adFldIsNullable
    
    
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
Public Sub Add_Rec()
    Fb_Press = 1
    Frame1.Enabled = True
    Frame1.Enabled = True
    MYSQL = "SELECT MAX(SLABCODE) AS MNO FROM BROKSLAB WHERE COMPCODE =" & GCompCode & ""
    Set TRec = Nothing
    Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If TRec.EOF Then
        LSLABCODE = 1
    Else
        If IsNull(TRec!MNo) Then
            LSLABCODE = 1
        Else
            LSLABCODE = TRec!MNo + 1
        End If
    End If
    Text1.text = LSLABCODE
    Set Rec = Nothing: Set Rec = New ADODB.Recordset
    Rec.Open "SELECT EXCODE,OPTIONS FROM EXMAST WHERE COMPCODE=" & GCompCode & "", Cnn, adOpenForwardOnly, adLockReadOnly
    If Not Rec.EOF Then
        Call RecSet
        Do While Not Rec.EOF
            If Rec!EXCODE = "EQ" Or Rec!EXCODE = "BEQ" Then
                RECGRID.AddNew
                RECGRID!EXCODE = Rec!EXCODE
                RECGRID!INSTTYPE = "CSH"
                RECGRID!BrokType = "Percentage Wise"
                RECGRID!BROKRATE = 0
                RECGRID!BROKRATE2 = 0
                RECGRID!MINRATE = 0
                RECGRID!MBrokType = "Transaction Wise"
                RECGRID!MBrokRate = 0
                RECGRID!MBROKRATE2 = 0
                RECGRID.Update
            Else
                RECGRID.AddNew
                RECGRID!EXCODE = Rec!EXCODE
                RECGRID!INSTTYPE = "FUT"
                RECGRID!BrokType = "Percentage Wise"
                RECGRID!BROKRATE = 0
                RECGRID!BROKRATE2 = 0
                RECGRID!MINRATE = 0
                RECGRID!MBrokType = "Transaction Wise"
                RECGRID!MBrokRate = 0
                RECGRID!MBROKRATE2 = 0
                RECGRID.Update
                If Rec!Options = "Y" Then
                    RECGRID.AddNew
                    RECGRID!EXCODE = Rec!EXCODE
                    RECGRID!INSTTYPE = "OPT"
                    RECGRID!BrokType = "Percentage Wise"
                    RECGRID!BROKRATE = 0
                    RECGRID!BROKRATE2 = 0
                    RECGRID!MINRATE = 0
                    RECGRID!MBrokType = "Transaction Wise"
                    RECGRID!MBrokRate = 0
                    RECGRID!MBROKRATE2 = 0
                    RECGRID.Update
                End If
            End If
            Rec.MoveNext
        Loop
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
    End If
    Call Get_Selection(Fb_Press)
    Text2.SetFocus
End Sub

Private Sub Combo1_GotFocus()
    If DataGrid1.Col = 2 Or DataGrid1.Col = 6 Then
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
            End If
    End If
    Combo1.Top = Val(DataGrid1.Top) + Val(DataGrid1.RowTop(DataGrid1.Row))
    Combo1.Width = Val(DataGrid1.Columns(DataGrid1.Col).Width)
    Combo1.Left = Val(DataGrid1.Left) + Val(DataGrid1.Columns(DataGrid1.Col).Left)
    Sendkeys "%{DOWN}"
    Sendkeys "%{DOWN}"
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        LGridRow = DataGrid1.Row: LGridCol = DataGrid1.Col:
        If DataGrid1.Col = 2 Then
            If KeyCode = 13 Then RECGRID!BrokType = Combo1.text
        ElseIf DataGrid1.Col = 6 Then
            If KeyCode = 13 Then RECGRID!MBrokType = Combo1.text
        End If
        Set DataGrid1.DataSource = RECGRID: DataGrid1.ReBind: DataGrid1.Refresh
        DataGrid1.Row = LGridRow: DataGrid1.Col = LGridCol - 1: Combo1.Visible = False: DataGrid1.SetFocus
    ElseIf KeyCode = 27 Then
        Combo1.Visible = False
    End If
End Sub
Public Sub Save_Rec()
On Error GoTo ERR1
Frame1.Enabled = False
Frame2.Enabled = False
CNNERR = True
Cnn.BeginTrans
If Fb_Press = 2 Then
    MYSQL = "DELETE FROM BROKSLAB WHERE COMPCODE =" & GCompCode & " AND SLABCODE =" & Val(Text1.text) & ""
    Cnn.Execute MYSQL
End If

    If Not RECGRID.EOF Then
        RECGRID.MoveFirst
        Do While Not RECGRID.EOF
            MYSQL = "INSERT INTO BROKSLAB (COMPCODE,SLABCODE,SLABNAME,EXCODE,INSTTYPE,BROKTYPE,BROKRATE,BROKRATE2,MINRATE,MBROKTYPE,MBROKRATE,MBROKRATE2)"
            MYSQL = MYSQL & " VALUES  (" & GCompCode & " ," & Val(Text1.text) & ",'" & Trim(Text2.text) & "','" & RECGRID!EXCODE & "','" & RECGRID!INSTTYPE & "','" & Left$(RECGRID!BrokType, 1) & "'," & Val(RECGRID!BROKRATE) & "," & Val(RECGRID!BROKRATE2) & "," & Val(RECGRID!MINRATE) & ",'" & Left$(RECGRID!MBrokType, 1) & "'," & Val(RECGRID!MBrokRate) & "," & Val(RECGRID!MBROKRATE2) & ")"
            Cnn.Execute MYSQL
            RECGRID.MoveNext
        Loop
    End If
    Cnn.CommitTrans
    CNNERR = False
    Text1.text = vbNullString
    Text2.text = vbNullString
    Call RecSet
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.ReBind
    DataGrid1.Refresh
Exit Sub
ERR1:
    MsgBox err.Description
    Cnn.RollbackTrans
    CNNERR = False
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If LenB(Text2.text) = 0 Then
    MsgBox "Slab Name Can No Be Blank"
    Cancel = True
    Exit Sub
Else
    MYSQL = "SELECT SLABNAME FROM BROKSLAB WHERE COMPCODE =" & GCompCode & "  AND SLABCODE =" & Val(Text1.text) & " AND SLABNAME  ='" & UCase(Text2.text) & "'"
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If Not TRec.EOF Then
        MsgBox "Duplicate Slab Name"
        Cancel = True
        Exit Sub
    End If
    
End If
DataGrid1.Col = 2
DataGrid1.SetFocus

End Sub

Public Sub DataList1_Click()
    Text2.text = DataList1.text
    Text1.text = DataList1.BoundText
End Sub
Private Sub DataList1_DblClick()
    If DataList1.Locked Then
    Else
        Call Get_Selection(2)
        Fb_Press = 2
        Call MODIFY_REC
    End If
End Sub
Private Sub DataList1_KeyPress(KeyAscii As Integer)
    If DataList1.Locked Then
    Else
        If KeyAscii = 13 Then
            Call Get_Selection(2)
            Fb_Press = 2
            Call MODIFY_REC
        End If
    End If
End Sub


Public Sub MODIFY_REC()
Frame1.Enabled = True
Frame2.Enabled = True
Set TRec = Nothing: Set TRec = New ADODB.Recordset
MYSQL = "SELECT * FROM BROKSLAB WHERE COMPCODE =" & GCompCode & "  AND SLABCODE =" & Val(Text1.text) & ""
TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
If Not TRec.EOF Then
    Call RecSet
    Do While Not TRec.EOF
        RECGRID.AddNew
        RECGRID!EXCODE = TRec!EXCODE
        RECGRID!INSTTYPE = TRec!INSTTYPE
        If TRec!BrokType = "P" Then
            RECGRID!BrokType = "Percentage Wise"
        ElseIf TRec!BrokType = "O" Then
            RECGRID!BrokType = "Opening Sauda"
        ElseIf TRec!BrokType = "C" Then
            RECGRID!BrokType = "Closing Sauda"
        ElseIf TRec!BrokType = "T" Then
            RECGRID!BrokType = "Transaction Wise"
        ElseIf TRec!BrokType = "I" Then
            RECGRID!BrokType = "IntraDay Wise"
        ElseIf TRec!BrokType = "V" Then
            RECGRID!BrokType = "Value Wise IntraDay Wise"
        ElseIf TRec!BrokType = "Z" Then
            RECGRID!BrokType = "ZLotWise"
        ElseIf TRec!BrokType = "D" Then
            RECGRID!BrokType = "Delivery Wise"
        ElseIf TRec!BrokType = "H" Then
            RECGRID!BrokType = "Higher Value Wise"
        ElseIf TRec!BrokType = "Q" Then
            RECGRID!BrokType = "Qtywise InttraDay Wise"
        ElseIf TRec!BrokType = "W" Then
            RECGRID!BrokType = "WHigher Valuse Wise Intraday Wise"
        ElseIf TRec!BrokType = "X" Then
            RECGRID!BrokType = "XIntraday Higher Wise"
        ElseIf TRec!BrokType = "R" Then
            RECGRID!BrokType = "RZLotWise IntraDay Wise"
        Else
            RECGRID!BrokType = "Percentage Wise"
        End If
        
        If TRec!MBrokType = "P" Then
            RECGRID!MBrokType = "Percentage Wise"
        ElseIf TRec!MBrokType = "O" Then
            RECGRID!MBrokType = "Opening Sauda"
        ElseIf TRec!MBrokType = "C" Then
            RECGRID!MBrokType = "Closing Sauda"
        ElseIf TRec!MBrokType = "T" Then
            RECGRID!MBrokType = "Transaction Wise"
        ElseIf TRec!MBrokType = "I" Then
            RECGRID!MBrokType = "IntraDay Wise"
        ElseIf TRec!MBrokType = "V" Then
            RECGRID!MBrokType = "Value Wise IntraDay Wise"
        ElseIf TRec!MBrokType = "Z" Then
            RECGRID!MBrokType = "ZLotWise"
        ElseIf TRec!MBrokType = "D" Then
            RECGRID!MBrokType = "Delivery Wise"
        ElseIf TRec!MBrokType = "H" Then
            RECGRID!MBrokType = "Higher Value Wise"
        ElseIf TRec!MBrokType = "Q" Then
            RECGRID!MBrokType = "Qtywise InttraDay Wise"
        ElseIf TRec!MBrokType = "W" Then
            RECGRID!MBrokType = "WHigher Valuse Wise Intraday Wise"
        ElseIf TRec!MBrokType = "X" Then
            RECGRID!MBrokType = "XIntraday Higher Wise"
        ElseIf TRec!MBrokType = "R" Then
            RECGRID!MBrokType = "RZLotWise IntraDay Wise"
        Else
            RECGRID!MBrokType = "Percentage Wise"
        End If
        RECGRID!BROKRATE = TRec!BROKRATE
        RECGRID!BROKRATE2 = TRec!BROKRATE2
        RECGRID!MINRATE = TRec!MINRATE
        RECGRID!MBrokRate = TRec!MBrokRate
        RECGRID!MBROKRATE2 = TRec!MBROKRATE2
        RECGRID.Update
        TRec.MoveNext
    Loop
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.ReBind
    DataGrid1.Refresh
End If
End Sub
