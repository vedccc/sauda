VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmsaudalot 
   Caption         =   "Contract Lot Size"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   14025
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   13695
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
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
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8520
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6015
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   10610
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Saudacode"
            Caption         =   "SaudaCode"
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
            DataField       =   "Maturity"
            Caption         =   "Maturity"
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
            DataField       =   "LOT"
            Caption         =   "Lot"
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
         BeginProperty Column04 
            DataField       =   "BROKLOT"
            Caption         =   "BrokLot"
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
         BeginProperty Column05 
            DataField       =   "TILLDATE"
            Caption         =   "TillDate"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin VB.Frame Frame4 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   4
         Top             =   120
         Width           =   3255
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Contract Lot Size"
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
            TabIndex        =   5
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   840
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
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmsaudalot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaudaRec As ADODB.Recordset
Dim GRIDREC As ADODB.Recordset
Dim TRec As ADODB.Recordset
Dim RECGRID As ADODB.Recordset

Private Sub Command1_Click()
Dim MFLAG As Boolean
Dim LENTRY As Boolean
Dim LSAUDA As String
Dim LMAT As Date
Dim LINST As String
Dim LLot As Double
Dim LBrokLot As Double
LENTRY = False
MFLAG = False
If RECGRID.RecordCount > 0 Then
    
    RECGRID.MoveFirst
    LSAUDA = RECGRID!SAUDACODE
    MYSQL = "DELETE FROM SAUDALOT WHERE COMPCODE =" & GCompCode & " AND SAUDACODE='" & LSAUDA & "'"
    Cnn.Execute MYSQL
    Do While Not RECGRID.EOF
        If RECGRID!MATURITY = RECGRID!TILLDATE Then
            MFLAG = True
        End If
        LSAUDA = RECGRID!SAUDACODE
        LINST = RECGRID!INSTTYPE
        LLot = RECGRID!LOT
        LBrokLot = RECGRID!BROKLOT
        LMAT = RECGRID!MATURITY
        If LLot <> 0 Then
            LENTRY = True
            MYSQL = "INSERT INTO SAUDALOT(COMPCODE,SAUDACODE,MATURITY,INSTTYPE,LOT,BROKLOT ,TILLDATE)"
            MYSQL = MYSQL & " VALUES(" & GCompCode & ",'" & RECGRID!SAUDACODE & "','" & Format(RECGRID!MATURITY, "YYYY/MM/DD") & "','" & RECGRID!INSTTYPE & "'," & RECGRID!LOT & "," & RECGRID!BROKLOT & ",'" & Format(RECGRID!TILLDATE, "YYYY/MM/DD") & "')"
            Cnn.Execute MYSQL
        End If
        
        RECGRID.MoveNext
    Loop
    If MFLAG = False And LENTRY = True Then
        MYSQL = "INSERT INTO SAUDALOT(COMPCODE,SAUDACODE,MATURITY,INSTTYPE,LOT,BROKLOT ,TILLDATE)"
        MYSQL = MYSQL & " VALUES(" & GCompCode & ",'" & LSAUDA & "','" & Format(LMAT, "YYYY/MM/DD") & "','" & LINST & "',1,1,'" & Format(LMAT, "YYYY/MM/DD") & "')"
        Cnn.Execute MYSQL
    End If
End If
MsgBox "Updated Successfully"
Call RecSet
Set DataGrid1.DataSource = RECGRID
DataGrid1.ReBind
DataGrid1.Refresh

Command1.Enabled = False
End Sub

Private Sub DataCombo1_GotFocus()
Sendkeys "%{DOWN}"
End Sub

Private Sub DataCombo1_Validate(Cancel As Boolean)
If DataCombo1.BoundText <> "" Then
    Call RecSet
    MYSQL = "SELECT * FROM SAUDALOT WHERE COMPCODE =" & GCompCode & "AND SAUDACODE ='" & DataCombo1.BoundText & "'"
    Set TRec = Nothing: Set TRec = New ADODB.Recordset
    TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
    If TRec.EOF Then
        MYSQL = "SELECT S.SAUDACODE,S.MATURITY,S.INSTTYPE,S.OPTTYPE,S.STRIKEPRICE,I.EXCHANGECODE,I.EXHCODE,S.TRADEABLELOT,S.BROKLOT AS BROKLOT FROM SAUDAMAST AS S, ITEMMAST AS I WHERE S.COMPCODE =" & GCompCode & " AND I.COMPCODE =S.COMPCODE"
        MYSQL = MYSQL & " AND S.ITEMCODE=I.ITEMCODE AND S.SAUDACODE ='" & DataCombo1.BoundText & "'"
        Set GRIDREC = Nothing
        Set GRIDREC = New ADODB.Recordset
        GRIDREC.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
        If Not GRIDREC.EOF Then
            RECGRID.AddNew
            RECGRID!SAUDACODE = GRIDREC!SAUDACODE
            RECGRID!MATURITY = GRIDREC!MATURITY
            RECGRID!INSTTYPE = GRIDREC!INSTTYPE
            RECGRID!OPTTYPE = GRIDREC!OPTTYPE
            RECGRID!STRIKEPRICE = GRIDREC!STRIKEPRICE
            RECGRID!ExCode = GRIDREC!EXCHANGECODE
            RECGRID!EXHCODE = GRIDREC!EXHCODE
            RECGRID!LOT = GRIDREC!TRADEABLELOT
            RECGRID!BROKLOT = GRIDREC!BROKLOT
            RECGRID!TILLDATE = GRIDREC!MATURITY
            RECGRID.Update
        End If
    Else
        TRec.MoveFirst
        Do While Not TRec.EOF
            RECGRID.AddNew
            RECGRID!SAUDACODE = TRec!SAUDACODE
            RECGRID!MATURITY = TRec!MATURITY
            RECGRID!INSTTYPE = TRec!INSTTYPE
            RECGRID!LOT = TRec!LOT
            RECGRID!BROKLOT = TRec!BROKLOT
            RECGRID!TILLDATE = TRec!TILLDATE
            RECGRID.Update
            TRec.MoveNext
        Loop
    End If
    
    
    Set DataGrid1.DataSource = RECGRID
    DataGrid1.ReBind
    DataGrid1.Refresh
    Command1.Enabled = True
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Load()
MYSQL = "SELECT SAUDACODE,SAUDANAME,ITEMCODE FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND EXCODE ='NSE' ORDER BY ITEMCODE, INSTTYPE"
Set SaudaRec = Nothing
Set SaudaRec = New ADODB.Recordset
SaudaRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
Set DataCombo1.RowSource = SaudaRec
    DataCombo1.BoundColumn = "SAUDACODE"
    DataCombo1.ListField = "SAUDACODE"
    
    
    
    
    
End Sub

Sub RecSet()
Set RECGRID = Nothing
    Set RECGRID = New ADODB.Recordset
    RECGRID.Fields.Append "SAUDACODE", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "MATURITY", adDate, , adFldIsNullable
    RECGRID.Fields.Append "INSTTYPE", adVarChar, 3, adFldIsNullable
    RECGRID.Fields.Append "OPTTYPE", adVarChar, 2, adFldIsNullable
    RECGRID.Fields.Append "STRIKEPRICE", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "EXCODE", adVarChar, 10, adFldIsNullable
    RECGRID.Fields.Append "EXHCODE", adVarChar, 100, adFldIsNullable
    RECGRID.Fields.Append "LOT", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "BROKLOT", adDouble, , adFldIsNullable
    RECGRID.Fields.Append "TILLDATE", adDate, , adFldIsNullable
    RECGRID.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
