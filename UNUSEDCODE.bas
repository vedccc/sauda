Attribute VB_Name = "ModUNUSED"
'Sub Save_AllTrade_Excel(LExcelType As Integer)
'
'Dim LIFLAG As Boolean:          Dim LBParty As String:          Dim LSParty As String:          Dim LBQty As Double
'Dim LSQty As Double:            Dim LBRate As Double:           Dim LSRate As Double:           Dim LDate As Date
'Dim LTradeDt As Date:           Dim LOverwrite As Boolean:      Dim TxtPath As String:          Dim LFileName As String
'Dim MCount As Long:             Dim LConType As String:         Dim LSConSno As Long:           Dim LQty As Double:
'Dim LContime As String:         Dim LOrdNo As String:           Dim LRate As Double:            Dim LOrdTime As String:
'Dim StrDateLen As Integer:      Dim LStrDate As String:         Dim LMonth As String:           Dim LYEAR As String
'Dim LClient As String:          Dim MParty As String:           Dim MBParty As String:          Dim MSParty As String
'Dim TRec As ADODB.Recordset:    Dim LLot As Double:             Dim LUserId As String:          Dim LPartyName As String
'Dim LExhCode As String:         Dim LSItemName As String:       Dim LStrConDate  As String:     Dim A As Integer
'Dim LFieldName As String:       Dim fld As ADODB.Field:         Dim LEx_Symbol  As String:      Dim LCMonth As String
'Dim LMFLPath As String:         Dim B As Integer:               Dim LSTR As String:             Dim LOConNo As String
'Dim LStrMonth As String:        Dim LDay As Integer:            Dim LSCondate As Date:
'
'On Error GoTo ERR1
'
'    Set TxtRec = Nothing: Set TxtRec = New ADODB.Recordset
'    If ChkNseExcel.Value = 1 Then
'        LSExCode = "NSE"
'    ElseIf Check19.Value = 1 Then
'        LSExCode = "CME"
'    ElseIf Check18.Value = 1 Then
'        LSExCode = "CME"
'    Else
'        LSExCode = "MCX"
'    End If
'    If Check15.Value = 1 Then LSExCode = "RACE"
'    If Check37.Value = 1 Then LSExCode = "REAL"
'    If Check39.Value = 1 And LExcelType = 25 Then LSExCode = "MCX"
'    Call GET_JCnn("\TRADE;")
'    If Check15.Value = 1 Then Call GET_JCnn("\RACE;")
'    If Check37.Value = 1 Then Call GET_JCnn("\REAL;")
'    Call Get_ExDetail(LSExCode)
'    MSParty = LExCont
'    Cnn.BeginTrans
'    CNNERR = True:
'    For LTradeDt = vcDTP1.Value To vcDTP2.Value
'        Select Case LExcelType
'        Case 1, 2, 6, 7, 10, 11, 12, 15, 16, 20, 21, 22, 23
'            TxtPath = CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & CStr(Year(LTradeDt)) & ".CSV"
'        Case 25
'            If Check15.Value = 1 Then
'                TxtPath = "TRD" & CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & Right$(CStr(Year(LTradeDt)), 2) & "_RACE.CSV"
'            ElseIf Check39.Value = 1 Then
'                TxtPath = CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & Right$(CStr(Year(LTradeDt)), 4) & ".CSV"
'            Else
'                TxtPath = "TRD" & CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & Right$(CStr(Year(LTradeDt)), 2) & "_REAL.CSV"
'            End If
'        Case 29, 30, 31, 32, 37, 38, 39, 40, 41, 42, 98, 99
'            TxtPath = CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & CStr(Year(LTradeDt)) & ".CSV"
'        End Select
'        If Check15.Value = 1 Then
'            LFileName = App.Path & "\RACE\" & TxtPath
'        ElseIf Check37.Value = 1 Then
'            LFileName = App.Path & "\REAL\" & TxtPath
'        Else
'            LFileName = App.Path & "\TRADE\" & TxtPath
'        End If
'        If Not FileExist(LFileName) Then
'            MsgBox LFileName & "  file not found", vbCritical
'        Else
'            If Check39.Value = 1 Then
'                LMFLPath = App.Path & "\TRADE\"
'                Call Create_Schema_Excel25(LMFLPath, TxtPath)
'            End If
'            If LExcelType = 39 Then
'                LMFLPath = App.Path & "\TRADE\"
'                Call Create_Schema_Excel25(LMFLPath, TxtPath)
'            End If
'
'            Set TxtRec = Nothing: Set TxtRec = New ADODB.Recordset
'            TxtRec.Open "Select * from " & TxtPath & "", Jcnn, adOpenStatic, adLockReadOnly, adCmdText
'            If Not TxtRec.EOF Then
'                FlagDataTrf = True
'                CNNERR = True
'                TxtRec.MoveFirst
'                LSCondate = DateValue(LTradeDt)
'                LDate = LSCondate
'                If MsgBox("Are you Sure to Delete Trades of  " & LSCondate & "", vbQuestion + vbYesNo, "Confirm") = vbYes Then
'                    If LExcelType = 25 Then
'                        If Check39.Value = 1 Then
'                            Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE = '" & Format(LSCondate, "yyyy/MM/dd") & "' AND DATAIMPORT=1 AND FILETYPE='" & LExcelType & "'"
'                        Else
'                            Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE = '" & Format(LSCondate, "yyyy/MM/dd") & "' AND DATAIMPORT=1 AND EXCODE ='" & LSExCode & "'AND FILETYPE='" & LExcelType & "'"
'                        End If
'                    ElseIf LExcelType = 99 Then
'
'                    Else
'                        Cnn.Execute "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE = '" & Format(LSCondate, "yyyy/MM/dd") & "' AND DATAIMPORT = 1 AND FILETYPE='" & Trim(Str(LExcelType)) & "'"
'                    End If
'                End If
'                Dim LIVouNo As String
'                Dim LVNo As Long
'                Dim LDR_CR As String
'                MCount = Get_Max_ConNo(LSCondate, 0)
'                GMaxVouNo = Val(Right$(Get_VouNo("JRNL", GFinYear), 7))
'                GMaxVouNo = GMaxVouNo + 1
'                LIVouNo = Get_Next_Vou_No(GMaxVouNo, "JRNL", Right$(GFinBegin, 2) & Right$(GFinEnd, 2))
'                Dim LNarration As String
'
'                LVNo = PInsert_Voucher(LIVouNo, LSCondate, "JV", "P", 0, "ADD", vbNullString, 0, vbNullString, "1", "0", 0, 0)
'
'                While Not TxtRec.EOF
'                    LBParty = vbNullString:       LSParty = vbNullString:       LSItemName = vbNullString:    LExhCode = vbNullString
'                    LOrdNo = vbNullString:        LUserId = vbNullString:       LContime = Time:              LSInstType = "FUT"
'                    LOrdTime = vbNullString:      LSOptType = vbNullString:     LSStrike = 0:                 LLot = 0:
'                    LBQty = 0:                    LSQty = 0:                    LBRate = 0:                   LSRate = 0
'                    LSSaudaID = 0:                LSItemID = 0
'
'                    If IsNull(TxtRec!f1) Then GoTo EFlag_Next
'                    Select Case LExcelType
'                    Case 98
'
'                        LSCondate = LTradeDt
'
'                        A = InStr(TxtRec!f25, ":")
'                        LBParty = Left(TxtRec!f25, A - 1)
'                        LPartyName = TxtRec!f25
'                        MBParty = Get_AccountMCode(LBParty)
'                        LSExCode = "MCX"
'                        If LenB(MBParty) < 1 Then
'                            MBParty = NewParty(LSExCode, LBParty, LPartyName, LBParty, True)
'                        End If
'                        If LenB(MBParty) < 1 Then
'                            MsgBox ""
'                        End If
'
'                        If Not IsNull(TxtRec!F22) Then
'                            LRate = Val(TxtRec!F22 & vbNullString)
'                            LDR_CR = "C"
'                        ElseIf TxtRec!f23 Then
'                            LRate = Val(TxtRec!f23 & vbNullString)
'                            LDR_CR = "D"
'                        End If
'                        LNarration = "Trf from " & Trim(TxtRec!F26)
'                        Call PInsert_Vchamt(LIVouNo, "JV", LSCondate, LDR_CR, MBParty, Abs(LRate), vbNullString, LSCondate, LNarration, vbNullString, vbNullString, 0, vbNullString, LVNo, 0, 0)
'                        GoTo EFlag_Next
'
'
'                    Case 99
'                        Dim LREC2 As ADODB.Recordset
'                        Dim LAC_CODE As String
'                            If IsNull(TxtRec!f1) Then GoTo EFlag_Next
'                            If Not IsNumeric(TxtRec!F8) Then GoTo EFlag_Next
'                            LSCondate = "07/09/2020"
'                            LSTR = TxtRec!F4 & "/" & TxtRec!f5 & "/" & TxtRec!f6 & ""
'                            LExhCode = TxtRec!f7
'                            LSMaturity = DateValue(LSTR)
'                            LBParty = Right(TxtRec!F10, 6)
'                            LPartyName = LBParty
'                            LBQty = Val(TxtRec!F8)
'                            If LBQty > 0 Then
'                                LConType = "B"
'                            Else
'                                LConType = "S"
'                            End If
'                            LBQty = Abs(LBQty)
'                            LBRate = Val(TxtRec!F9)
'                            LSQty = LBQty
'                            LSRate = LSRate
'                    Case 10
'                        If IsNull(TxtRec!f7) Then GoTo EFlag_Next
'                        If Not (Left(TxtRec!f7, 1) = "b" Or Left(TxtRec!f7, 1) = "s") Then GoTo EFlag_Next
'                        LSExCode = vbNullString
'                        LOConNo = TxtRec!f1
'                        LOrdNo = (TxtRec!F3 & vbNullString)
'                        LBParty = Trim(IIf(IsNull(TxtRec!F4), vbNullString, (TxtRec!F4))):
'                        LPartyName = TxtRec!f5
'                        LStrConDate = Left$(TxtRec!f6, 10)
'                        LSCondate = DateValue(Right(LStrConDate, 2) & "/" & Mid(LStrConDate, 6, 2) & "/" & Left(LStrConDate, 4))
'                        LPtyContype = UCase(Left(TxtRec!f7, 1))
'                        A = InStr(TxtRec!F9, "-")
'                        If A = 0 Then
'                            A = InStr(TxtRec!F9, "/")
'                            If A = 0 Then
'                                LExhCode = Left(TxtRec!F9, 2)
'                                LStrConDate = Right(TxtRec!F9, 3)
'                                LYEAR = "20" & Right(LStrConDate, 2)
'                                LMonth = Get_CMX_Month(Left(LStrConDate, 1))
'                                LStrConDate = "28" & "/" & LMonth & "/" & LYEAR
'                                LSMaturity = DateValue(LStrConDate)
'                            Else
'                                LExhCode = Trim(Left(TxtRec!F9, A - 1))
'                                LYEAR = Year(LSCondate)
'                                LMonth = Right(TxtRec!F9, 2)
'                                LSMaturity = DateValue("28/" & LMonth & "/" & LYEAR & "")
'                            End If
'                        Else
'                            B = InStr(TxtRec!F9, "/")
'                            If B = 0 Then
'                                LExhCode = Trim(Left(TxtRec!F9, A - 1))
'                                LStrConDate = Right(TxtRec!F9, 3)
'                                LYEAR = "20" & Right(LStrConDate, 2)
'                                LMonth = Get_CMX_Month(Left(LStrConDate, 1))
'                                LStrConDate = "28" & "/" & LMonth & "/" & LYEAR
'                                LSMaturity = DateValue(LStrConDate)
'                            Else
'                                LExhCode = Trim(Left(TxtRec!F9, B - 1))
'                                LYEAR = Year(LSCondate)
'                                LMonth = Right(TxtRec!F9, 2)
'                                LSMaturity = DateValue("28/" & LMonth & "/" & LYEAR & "")
'                            End If
'                        End If
'                        If Right(TxtRec!F10, 1) = "K" Then
'                            LBQty = Val(Left(TxtRec!F10, Len(TxtRec!F10) - 1)) * 1000
'                        Else
'                            LBQty = Val(TxtRec!F10)
'                        End If
'                        If month(LSMaturity) < month(LSCondate) Then
'                            LYEAR = Year(LSCondate) + 1
'                            LSMaturity = DateValue("28/" & LMonth & "/" & LYEAR & "")
'                        End If
'                        LBRate = Val(IIf(IsNull(TxtRec!F11), 0, TxtRec!F11)):
'                        LSRate = LBRate
'                        LSParty = LExCont
'                    Case 11
'                        If IsNull(TxtRec!f1) Then GoTo EFlag_Next
'                        If Not (UCase(TxtRec!f7) = "BUY" Or UCase(TxtRec!f7) = "SELL") Then GoTo EFlag_Next
'                        LBParty = Trim(IIf(IsNull(TxtRec!f1), vbNullString, (TxtRec!f1))):
'                        LPartyName = LBParty
'                        LSExCode = Left(TxtRec!F2, 3)
'                        LExhCode = Trim(TxtRec!F3)
'                        LSMaturity = DateValue(Left(TxtRec!F4, 2) & "/" & Mid(TxtRec!F4, 4, 3) & "/20" & Right(TxtRec!F4, 2))
'                        LPtyContype = UCase(Left(TxtRec!f7, 1))
'                        LBQty = Val(TxtRec!F8)
'                        LBRate = Val(TxtRec!F9)
'                        LSRate = LBRate
'                        LOConNo = TxtRec!F14
'                        LOrdNo = (TxtRec!F14 & vbNullString)
'                        LStrConDate = Left$(TxtRec!F11, 10)
'                        LSCondate = DateValue(LStrConDate)
'                        LContime = Right(TxtRec!F11, 5)
'                    Case 13
'                        If IsNull(TxtRec!f1) Then GoTo EFlag_Next
'                        If Not (UCase(TxtRec!f7) = "BUY" Or UCase(TxtRec!f7) = "SELL") Then GoTo EFlag_Next
'                        LOConNo = TxtRec!f1
'                        LBParty = Trim(IIf(IsNull(TxtRec!F2), vbNullString, (TxtRec!F2))):
'                        LPartyName = LBParty
'                        LSExCode = "NSE"
'                        LExhCode = Left(TxtRec!F4, (Len(TxtRec!F4) - 8))
'                        LSTR = Right(TxtRec!F4, 8)
'                        LSMaturity = DateValue("28/" & Mid(LSTR, 3, 3) & "/20" & Left(LSTR, 2))
'                        LPtyContype = UCase(Left(TxtRec!f5, 1))
'                        LBQty = Val(TxtRec!f7)
'                        LBRate = Val(TxtRec!f6)
'                        LSRate = LBRate
'                        LOConNo = TxtRec!F14
'                        LOrdNo = LOConNo
'                        LStrConDate = Left$(TxtRec!F3, 10)
'                        LSCondate = DateValue(LStrConDate)
'                        LContime = Right(TxtRec!F3, 5)
'
'                    Case 15
'                        If IsNull(TxtRec!f1) Then GoTo EFlag_Next
'                        If Not IsNumeric(TxtRec!f6) Then GoTo EFlag_Next
'                        LSCondate = DateValue(TxtRec!f1)
'                        LBParty = Trim(IIf(IsNull(TxtRec!F2), vbNullString, (TxtRec!F2))):
'
'                        LExhCode = Trim(IIf(IsNull(TxtRec!F3), vbNullString, TxtRec!F3))
'                        LSMaturity = DateValue(TxtRec!F4)
'                        LBQty = Val(IIf(IsNull(TxtRec!f5), 0, TxtRec!f5))
'                        If LBQty > 0 Then
'                            LConType = "B"
'                        Else
'                            LConType = "S"
'                            LBQty = Abs(LBQty)
'                        End If
'                        LBRate = Val(IIf(IsNull(TxtRec!f6), 0, TxtRec!f6)):
'                        LSRate = Val(IIf(IsNull(TxtRec!f7), 0, TxtRec!f7)):
'                        LSParty = Trim$(IIf(IsNull(TxtRec!F8), vbNullString, (TxtRec!F8))):
'                        If Not IsNull(TxtRec!F9) Then
'                            If LenB(TxtRec!F9) > 0 Then
'                                If TxtRec!F9 = "CSH" Then
'                                    LSInstType = "CSH"
'                                    LOptType = vbNullString
'                                    LStrike = 0
'                                End If
'                                If Not IsNull(TxtRec!F10) Then
'                                    If IsNumeric(TxtRec!F10) Then
'                                        If Val(TxtRec!F10) <> 0 Then
'                                        LSInstType = "OPT": LOptType = Left$(TxtRec!F9, 2): LStrike = Val(TxtRec!F10 & vbNullString)
'                                    End If
'                                End If
'                                LFieldName = "F11"
'                                For Each fld In TxtRec.Fields
'                                    If UCase(fld.NAME) = LFieldName Then
'                                        If Not IsNull(TxtRec!F11) Then LContime = TxtRec!F11
'                                        Exit For
'                                    End If
'                                Next
'                                End If
'                            End If
'                        End If
'                    Case 16
'                        If IsNull(TxtRec!f1) Then GoTo EFlag_Next
'                        If Not IsNumeric(TxtRec!f6) Then GoTo EFlag_Next
'                        LOConNo = (TxtRec!f1)
'                        LBParty = Trim(IIf(IsNull(TxtRec!F2), vbNullString, (TxtRec!F2))):
'                        LPartyName = LBParty
'                       ' A = InStr(TxtRec!F4, "|")
'                        LExhCode = Left$(TxtRec!F4, (Len(TxtRec!F4) - 8))
'                        LSTR = Right(TxtRec!F4, 8)
'                        LMonth = Mid(LSTR, 3, 3)
'                        LYEAR = "20" & Left(LSTR, 2)
'                        LDate = DateValue("28/" & LMonth & "/" & LYEAR)
'                        'If month(LDate) < month(LSCondate) Then
'                        '    LYEAR = Str(Val(LYEAR) + 1)
'                        '    LDate = DateValue("28/" & LMonth & "/" & LYEAR)
'                        'End If
'                        LSMaturity = DateValue(LDate)
'                        LConType = Left(TxtRec!f5, 1)
'                        LPtyContype = LConType
'                        LBQty = Val(IIf(IsNull(TxtRec!f7), 0, TxtRec!f7))
'                        LBRate = Val(IIf(IsNull(TxtRec!f6), 0, TxtRec!f6)):
'                        LSRate = LBRate
'                        LSParty = LExCont
'                        'LOConNo = TxtRec!F8
'                    Case 20
'                        If Not IsNumeric(TxtRec!F4) Then GoTo EFlag_Next
'                        LSExCode = "MCX"
'                        A = InStr(TxtRec!f1, "-")
'                        LStrConDate = Left(TxtRec!f1, A - 1) & "/" & Mid(TxtRec!f1, A + 1, 3) & "/20" & Right(TxtRec!f1, 2) & ""
'                        LSCondate = DateValue(LStrConDate)
'                        A = InStr(TxtRec!F2, " ")
'                        LExhCode = Left(TxtRec!F2, A - 1)
'                        LMonth = Right(RTrim(TxtRec!F2), 3)
'                        LSMaturity = DateValue("28/" & LMonth & "/" & Year(LTradeDt))
'                        If month(LSMaturity) < month(LTradeDt) Then
'                            LSMaturity = DateValue("28/" & LMonth & "/" & Year(LTradeDt) + 1)
'                        End If
'                        LPtyContype = Left(TxtRec!F3, 1)
'                        LBQty = Val(TxtRec!F4)
'                        LBRate = Val(TxtRec!f5)
'                        LSRate = LBRate
'                        LBParty = Trim(IIf(IsNull(TxtRec!f6), vbNullString, (TxtRec!f6))):
'                        LSParty = Trim$(IIf(IsNull(TxtRec!f7), vbNullString, (TxtRec!f7))):
'                        LOConNo = TxtRec!F8
'                    Case 21
'                        LOptType = vbNullString
'                        LStrike = 0
'                        LSInstType = "FUT"
'                        If Not IsNumeric(TxtRec!f7) Then GoTo EFlag_Next
'                        LSExCode = "NSE"
'                        LStrConDate = Left(TxtRec!f1, 2) & "/" & Mid(TxtRec!f1, 4, 3) & "/20" & Right(TxtRec!f1, 2) & ""
'                        LSCondate = DateValue(LStrConDate)
'                        LExhCode = Trim(TxtRec!F2)
'                        LStrConDate = Left(TxtRec!F3, 2) & "/" & Mid(TxtRec!F3, 4, 3) & "/20" & Right(TxtRec!F3, 2) & ""
'                        LSMaturity = DateValue(LStrConDate)
'                        If Not IsNull(TxtRec!f5) Then LOptType = Trim(TxtRec!f5)
'
'                        If LOptType = "CE" Or LOptType = "PE" Then
'                            LSInstType = "OPT"
'                            LStrike = Val(TxtRec!F4)
'                        End If
'                        LPtyContype = Left(TxtRec!f6, 1)
'                        LBQty = Val(TxtRec!f7)
'                        LBRate = Val(TxtRec!F8)
'                        LSRate = LBRate
'                        LBParty = Trim(IIf(IsNull(TxtRec!F9), vbNullString, (TxtRec!F9))):
'                        LSParty = Trim$(IIf(IsNull(TxtRec!F10), vbNullString, (TxtRec!F10))):
'                        LOConNo = TxtRec!F11
'                    Case 22
'                        LOptType = vbNullString
'                        LStrike = 0
'                        LSInstType = "FUT"
'                        If IsNull(TxtRec!F11) Then GoTo EFlag_Next
'                        If Not IsNumeric(TxtRec!F10) Then GoTo EFlag_Next
'                        LSExCode = "MCX"
'                        LStrConDate = Left(TxtRec!f6, 10)
'                        LStrConDate = Right(LStrConDate, 2) & "/" & Mid(LStrConDate, 6, 2) & "/" & Left(LStrConDate, 4) & ""
'                        LSCondate = DateValue(LStrConDate)
'                        A = InStr(TxtRec!F9, "-")
'
'                        LExhCode = Left(TxtRec!F9, A - 1)
'                        LMonth = Right(LExhCode, 3)
'                        LExhCode = Left(TxtRec!F9, A - 4)
'                        LYEAR = Year(LSCondate)
'                        LStrConDate = "28/" & LMonth & "/" & LYEAR
'                        LSMaturity = DateValue(LStrConDate)
'                        LPtyContype = UCase(Left(TxtRec!f7, 1))
'                        LBQty = Val(TxtRec!F10)
'                        LBRate = Val(TxtRec!F11)
'                        LSRate = LBRate
'                        LBParty = Trim(TxtRec!F4):
'                        LSParty = LExCont
'                        LOConNo = TxtRec!f1
'                    Case 25
'                        If IsNull(TxtRec!f7) Then GoTo EFlag_Next
'                        If Not (Left(TxtRec!f7, 1) = "b" Or Left(TxtRec!f7, 1) = "s") Then GoTo EFlag_Next
'                        LOConNo = TxtRec!f1:                        LOrdNo = (TxtRec!F3 & vbNullString)
'                        LPartyName = TxtRec!f5:                     LStrConDate = Left$(TxtRec!f6, 10)
'                        LContime = Right(TxtRec!f6, Len(TxtRec!f6) - 11)
'                        LBParty = Trim(IIf(IsNull(TxtRec!F4), vbNullString, (TxtRec!F4))):
'                        LSCondate = DateValue(Right(LStrConDate, 2) & "/" & Mid(LStrConDate, 6, 2) & "/" & Left(LStrConDate, 4))
'                        LPtyContype = UCase(Left(TxtRec!f7, 1))
'                        A = InStr(TxtRec!F9, "-")
'                        If LSExCode = "REAL" Then A = 0
'                        If A = 0 Then
'                            If Left(TxtRec!F9, 4) = "GOLD" Then
'                                LExhCode = "GOLD-" & Right(TxtRec!F9, 5)
'                                LYEAR = Year(LSCondate)
'                                LCMonth = Mid(TxtRec!F9, 5, 3)
'                                LStrConDate = "28" & "/" & Mid(TxtRec!F9, 5, 3) & "/" & LYEAR
'                                LSMaturity = DateValue(LStrConDate)
'                                If month(LSMaturity) < month(LSCondate) Then
'                                    LYEAR = (Year(LSCondate) + 1)
'                                    LStrConDate = "01" & "/" & month(LSMaturity) & "/" & LYEAR
'                                    LSMaturity = DateValue(LStrConDate)
'                                End If
'                            Else
'                                LExhCode = Left$(TxtRec!F9, Len(TxtRec!F9) - 3)
'                                LStrConDate = Right$(TxtRec!F9, 3)
'                                LCMonth = Left(LStrConDate, 1)
'                                LYEAR = "20" & Right(LStrConDate, 2)
'                                LMonth = Get_CMX_Month(Left(LStrConDate, 1))
'                                LStrConDate = "28" & "/" & LMonth & "/" & LYEAR
'                                LSMaturity = DateValue(LStrConDate)
'                            End If
'                        Else
'                            LExhCode = Trim(Left(TxtRec!F9, A - 1))
'                            LStrConDate = Right(TxtRec!F9, 3)
'                            LCMonth = Left(LStrConDate, 1)
'                            LYEAR = "20" & Right(LStrConDate, 2)
'                            LMonth = Get_CMX_Month(Left(LStrConDate, 1))
'                            LStrConDate = "28" & "/" & LMonth & "/" & LYEAR
'                            LSMaturity = DateValue(LStrConDate)
'                        End If
'                        If Right(TxtRec!F10, 1) = "K" Then
'                            LBQty = Val(Left(TxtRec!F10, Len(TxtRec!F10) - 1)) * 1000
'                        Else
'                            LBQty = Val(TxtRec!F10)
'                        End If
'                        LBRate = Val(IIf(IsNull(TxtRec!F11), 0, TxtRec!F11)):
'                        LSRate = LBRate
'                        LSParty = LExCont
'                    Case 38
'                        If IsNull(TxtRec!f7) Then GoTo EFlag_Next
'                        If Not (Left(TxtRec!f7, 1) = "b" Or Left(TxtRec!f7, 1) = "s") Then GoTo EFlag_Next
'                        LOConNo = TxtRec!f1:                        LOrdNo = (TxtRec!F3 & vbNullString)
'                        LBParty = Trim(IIf(IsNull(TxtRec!F4), vbNullString, (TxtRec!F4))):
'                        LPartyName = LBParty & TxtRec!f5:                     LStrConDate = Left$(TxtRec!f6, 10)
'
'                        LSCondate = DateValue(Right(LStrConDate, 2) & "/" & Mid(LStrConDate, 6, 2) & "/" & Left(LStrConDate, 4))
'                        LPtyContype = UCase(Left(TxtRec!f7, 1))
'                        DoEvents
'                        A = Len(TxtRec!F9)
'                        LExhCode = Left$(TxtRec!F9, A - 3)
'                        If TxtRec!F9 = "GCM20" Then
'                            LExhCode = "GC"
'                            LMonth = "06"
'                            LYEAR = Year(LSCondate)
'                        ElseIf TxtRec!F9 = "NASDAQ JUN20" Then
'                            LExhCode = "NASDAQ"
'                            LMonth = "06"
'                            LYEAR = Year(LSCondate)
'                        Else
'                            If Right(TxtRec!F9, 2) = "20" Or Right(TxtRec!F9, 2) = "19" Then
'                                LSTR = Right(TxtRec!F9, 5)
'                                LMonth = Left(LSTR, 3)
'                                LYEAR = "20" & Right(TxtRec!F9, 2)
'                                LExhCode = Left$(TxtRec!F9, A - 5)
'                            Else
'                                LSTR = Right(TxtRec!F9, 4)
'                                If LSTR = "SEPP" Or LSTR = "JANN" Or LSTR = "MARR" Or LSTR = "NOVV" Or LSTR = "MAYY" Then
'                                    LMonth = Left(LSTR, 3)
'                                    LExhCode = Left$(TxtRec!F9, A - 4)
'                                ElseIf LSTR = "SEP2" Then
'                                    LMonth = Left(LSTR, 3)
'                                    LExhCode = Left$(TxtRec!F9, A - 4)
'                                Else
'                                    LMonth = Right(TxtRec!F9, 3)
'                                    LExhCode = Left$(TxtRec!F9, A - 3)
'                                End If
'                                LYEAR = Year(LSCondate)
'                            End If
'
'                        End If
'                        LStrConDate = "28" & "/" & LMonth & "/" & LYEAR
'                        LSMaturity = DateValue(LStrConDate)
'                        If month(LSMaturity) < month(LSCondate) Then
'                            LYEAR = (Year(LSCondate) + 1)
'                            LStrConDate = "28" & "/" & LMonth & "/" & LYEAR
'                            LSMaturity = DateValue(LStrConDate)
'                            LMonth = month(LSMaturity)
'                        End If
'                        If Right(TxtRec!F10, 1) = "K" Then
'                            LBQty = Val(Left(TxtRec!F10, Len(TxtRec!F10) - 1)) * 1000
'                        Else
'                            LBQty = Val(TxtRec!F10)
'                        End If
'                        LBRate = Val(IIf(IsNull(TxtRec!F11), 0, TxtRec!F11)):
'                        LSRate = LBRate
'                        LSParty = LExCont
'                    Case 39, 42
'                        If IsNull(TxtRec!f7) Then GoTo EFlag_Next
'                        If Not (Left(TxtRec!f7, 1) = "b" Or Left(TxtRec!f7, 1) = "s") Then GoTo EFlag_Next
'                        LOConNo = TxtRec!f1:
'                        LOrdNo = (TxtRec!F3 & vbNullString)
'                        LPartyName = TxtRec!f5:                     LStrConDate = Left$(TxtRec!f6, 10)
'                        LBParty = Right(Trim(IIf(IsNull(TxtRec!F4), vbNullString, (TxtRec!F4))), 6):
'                        LSCondate = DateValue(Right(LStrConDate, 2) & "/" & Mid(LStrConDate, 6, 2) & "/" & Left(LStrConDate, 4))
'                        LPtyContype = UCase(Left(TxtRec!f7, 1))
'                        LYEAR = Val("20" & Right(TxtRec!F9, 2))
'                        LSTR = Right(TxtRec!F9, 3)
'
'                        If LExcelType = 39 Then
'                            LMonth = Get_CMX_Month(Left(LSTR, 1))
'                            LSMaturity = DateValue(TxtRec!F20)
'                            A = InStr(TxtRec!F9, "-")
'                            If A = 0 Then
'                                A = InStr(TxtRec!F9, "-R")
'                                If Left(TxtRec!F9, 4) = "GOLD" Then
'                                    LExhCode = "GOLD"
'                                ElseIf Left(TxtRec!F9, 5) = "CRUDE" Then
'                                    LExhCode = "CRUDEOIL"
'                                ElseIf Left(TxtRec!F9, 6) = "SILVER" Then
'                                    LExhCode = "SILVER"
'                                ElseIf Left(TxtRec!F9, 6) = "NICKEL" Then
'                                    LExhCode = "NICKEL"
'                                ElseIf Left(TxtRec!F9, 2) = "NG" Then
'                                    LExhCode = "NATURALGAS"
'                                Else
'                                    LExhCode = Left(TxtRec!F9, A - 1)
'                                End If
'                            Else
'                                If Left(TxtRec!F9, 4) = "GOLD" Then
'                                    LExhCode = "GOLD"
'                                ElseIf Left(TxtRec!F9, 5) = "CRUDE" Then
'                                    LExhCode = "CRUDEOIL"
'                                ElseIf Left(TxtRec!F9, 6) = "SILVER" Then
'                                    LExhCode = "SILVER"
'                                ElseIf Left(TxtRec!F9, 6) = "NICKEL" Then
'                                    LExhCode = "NICKEL"
'                                ElseIf Left(TxtRec!F9, 2) = "NG" Then
'                                    LExhCode = "NATURALGAS"
'                                Else
'                                    LExhCode = Left(TxtRec!F9, (A - 1))
'                                End If
'                            End If
'                        ElseIf LExcelType = 42 Then
'                            A = InStr(TxtRec!F9, "-")
'                            If A = 0 Then
'                              '  MsgBox ""
'                                A = InStr(TxtRec!F9, "/")
'                                If Left(TxtRec!F9, 6) = "SILVER" Then
'                                    LExhCode = "SILVERM"
'                                ElseIf Left(TxtRec!F9, 4) = "GOLD" Then
'                                    LExhCode = "GOLDM"
'                                End If
'                            Else
'                                B = InStr(TxtRec!F9, "/")
'                                If B <> 0 Then
'                                    If Left(TxtRec!F9, 6) = "SILVER" Then
'                                        LExhCode = "SILVERM"
'                                    ElseIf Left(TxtRec!F9, 4) = "GOLD" Then
'                                        LExhCode = "GOLDM"
'                                    End If
'                                Else
'
'                                    LExhCode = Left(TxtRec!F9, A - 1)
'                                End If
'                            End If
'                        End If
'                            If Right(TxtRec!F10, 1) = "K" Then
'                            LBQty = Val(Left(TxtRec!F10, Len(TxtRec!F10) - 1)) * 1000
'                        Else
'                            LBQty = Val(TxtRec!F10)
'                        End If
'                        LBRate = Val(IIf(IsNull(TxtRec!F11), 0, TxtRec!F11)):
'                        LSRate = LBRate
'                        LSParty = LExCont
'                    Case 40
'                        If IsNull(TxtRec!f7) Then GoTo EFlag_Next
'                        If Not (Left(TxtRec!f7, 1) = "b" Or Left(TxtRec!f7, 1) = "s") Then GoTo EFlag_Next
'                        LOConNo = TxtRec!f1:                        LOrdNo = (TxtRec!F3 & vbNullString)
'                        LBParty = Trim(IIf(IsNull(TxtRec!F4), vbNullString, (TxtRec!F4))):
'                        LPartyName = LBParty & TxtRec!f5:                     LStrConDate = Left$(TxtRec!f6, 10)
'
'                        LSCondate = DateValue(Right(LStrConDate, 2) & "/" & Mid(LStrConDate, 6, 2) & "/" & Left(LStrConDate, 4))
'                        LPtyContype = UCase(Left(TxtRec!f7, 1))
'                        DoEvents
'                        If TxtRec!f19 = "USD" Then
'                            GoTo EFlag_Next
'                            'A = Len(TxtRec!f9)
'                            'LExhCode = Left$(TxtRec!f9, A - 3)
'                            'LSTR = Right(TxtRec!f9, 3)
'                            'LStrMonth = Left(LSTR, 1)
'                            'LMonth = Get_CMX_Month(LStrMonth)
'
'                        ElseIf Right(TxtRec!F9, 4) = "JUNE" Or Right(TxtRec!F9, 4) = "JULY" Then
'                            A = Len(TxtRec!F9)
'                            LExhCode = Left$(TxtRec!F9, A - 4)
'                            LSTR = Right(TxtRec!F9, 4)
'                            LStrMonth = Left(LSTR, 3)
'                        ElseIf Right(TxtRec!F9, 3) = "JUN" Or Right(TxtRec!F9, 3) = "AUG" Or Right(TxtRec!F9, 3) = "JUL" Or Right(TxtRec!F9, 3) = "SEP" Or Right(TxtRec!F9, 3) = "OCT" Then
'                            A = Len(TxtRec!F9)
'                            LExhCode = Left$(TxtRec!F9, A - 3)
'                            LStrMonth = Right(TxtRec!F9, 3)
'                        ElseIf Right(TxtRec!F9, 2) = "20" Then
'                            A = Len(TxtRec!F9)
'                            LExhCode = Left$(TxtRec!F9, A - 5)
'                            LSTR = Right(TxtRec!F9, 5)
'                            LStrMonth = Left(LSTR, 3)
'                        ElseIf Right(TxtRec!F9, 1) = "2" Then
'                            A = Len(TxtRec!F9)
'                            LExhCode = Left$(TxtRec!F9, A - 4)
'                            LSTR = Right(TxtRec!F9, 4)
'                            LStrMonth = Left(LSTR, 3)
'                        End If
'                        LYEAR = Year(LSCondate)
'                        LSMaturity = DateValue("28/" & LStrMonth & "/" & LYEAR)
'                        If month(LSCondate) > month(LSMaturity) Then
'                            LYEAR = (Year(LSCondate) + 1)
'                            LSMaturity = DateValue("28/" & LStrMonth & "/" & LYEAR)
'                            LMonth = month(LSMaturity)
'                        End If
'                        If Right(TxtRec!F10, 1) = "K" Then
'                            LBQty = Val(Left(TxtRec!F10, Len(TxtRec!F10) - 1)) * 1000
'                        Else
'                            LBQty = Val(TxtRec!F10)
'                        End If
'                        LBRate = Val(IIf(IsNull(TxtRec!F11), 0, TxtRec!F11)):
'                        LSRate = LBRate
'                        LSParty = LExCont
'                    Case 41
'                        If IsNull(TxtRec!f5) Then GoTo EFlag_Next
'                        If Not IsNumeric(TxtRec!f5) Then GoTo EFlag_Next
'                        LSTR = Right(TxtRec!F3, 3):
'                        If LSTR = "FUT" Then
'                            LSInstType = "FUT"
'                            LSTR = Right(TxtRec!F3, 8):
'                            LStrMonth = Mid(LSTR, 3, 3):
'                            LYEAR = Val("20" & Left(LSTR, 2))
'                            A = Len(TxtRec!F3):
'                            LExhCode = Mid(TxtRec!F3, 1, A - 8)
'                            LSMaturity = DateValue("20/" & LStrMonth & "/" & LYEAR)
'                        Else
'                            LSInstType = "OPT"
'                            A = InStr(TxtRec!F3, "2")
'                            LSTR = Mid(TxtRec!F3, A, Len(TxtRec!F3))
'                            LExhCode = Left(TxtRec!F3, A - 1)
'                            LOptType = Right(LSTR, 2)
'                            LYEAR = Val("20" & Left(LSTR, 2))
'                            LMonth = Val(Mid(LSTR, 3, 1))
'                            LDay = Val(Mid(LSTR, 4, 2))
'                            LSMaturity = DateValue(LDay & "/" & LMonth & "/" & LYEAR)
'                            LStrike = Val(Mid(LSTR, 6, Len(LSTR) - 2))
'                        End If
'                        LOConNo = Trim(TxtRec!F8):                      LOrdNo = (TxtRec!F10 & vbNullString)
'                        LSCondate = DateValue(Left(TxtRec!f7, 10))
'                        LContime = Mid(TxtRec!f7, 11, Len(TxtRec!f7)):  LPtyContype = UCase(Left(TxtRec!F4, 1))
'                        LBQty = Val(TxtRec!f5):                         LBRate = Val(TxtRec!f6)
'                        LBParty = Trim(IIf(IsNull(TxtRec!F11), vbNullString, (TxtRec!F11))):
'                        LPartyName = LBParty
'                        LSRate = LBRate
'                        LOrdNo = Trim(TxtRec!F10)
'                        DoEvents
'                    End Select
'                    If LenB(LExhCode) < 1 Then GoTo EFlag_Next
'                    If LExcelType = 1 Then
'                        LExID = Get_ExID("MCX")
'                        LSItemCode = Get_ItemMaster(LExID, LExhCode)
'                        If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                        LItemID = Get_ITEMID(LSItemCode)
'                        LSSaudaCode = Get_SaudaMaster(LExID, LItemID, LSMaturity, LSInstType, LOptType, LStrike)
'
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        LSaudaID = Get_SaudaID(LSSaudaCode)
'
'                        LLot = Get_LotSize(LItemID, LSaudaID, LExID, LLotWise)
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'                        If LBQty <> 0 Then
'                            LBQty = LBQty / LLot
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            If LenB(MBParty) < 1 Then GoTo EFlag_Next
'                            LPtyContype = "B":
'                            MCount = MCount + 1
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            LOConNo = Trim(Str(MCount))
'                            If InStr(LBillExId, Str(LExID)) = 0 Then
'                                If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                LBillExId = LBillExId & Str(LExID)
'                            End If
'                            If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'
'                            Call Add_To_Ctr_D(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, "0", "Y", LExID, LItemID, LSaudaID)
'                        End If
'                        If LSQty <> 0 Then
'                            LSQty = LSQty / LLot
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            If LenB(MBParty) < 1 Then GoTo EFlag_Next
'                            LPtyContype = "S"
'                            MCount = MCount + 1
'                            LOConNo = Trim(Str(MCount))
'                            If InStr(LBillExId, Str(LExID)) = 0 Then
'                                If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                LBillExId = LBillExId & Str(LExID)
'                            End If
'
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                                        If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'                            Call Add_To_Ctr_D(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LSQty, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, "0", "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    ElseIf LExcelType = 99 Then
'                        LLot = 1
'                        'MYSQL = "SELECT EXID,ITEMCODE,EXCHANGECODE,LOT,ITEMID FROM ITEMMAST WHERE ITEMCODE   ='" & LExhCode & "'"
'                        MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "','" & vbNullString & "'"
'                        Set TRec = Nothing:   Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                        If Not TRec.EOF Then
'                            LSExCode = TRec!EXCHANGECODE: LSItemCode = TRec!ITEMCODE: LLot = TRec!LOT: LExID = TRec!EXID
'                            LItemID = TRec!ITEMID
'                        Else
'                            MYSQL = "SELECT ITEMCODE,EXCODE ,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL  ='" & LExhCode & "'"
'                            Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSExCode = TRec!EXCODE
'                                LSItemCode = TRec!ITEMCODE
'                                LLot = TRec!LOT
'                                LExID = Get_ExID(LSExCode)
'                                LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'                                'LSSauda = Create_SSaudaMast(LExhCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LSItemCode)
'                                LItemID = Get_ITEMID(LSItemCode)
'                            Else
'                                MsgBox LExhCode & " Not Found"
'                                GoTo EFlag_Next
'                            End If
'                        End If
'                        If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                        Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                        MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",''," & month(LSMaturity) & "," & Year(LSMaturity) & ",'FUT'" & ",'',0"
'                        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                        If Not TRec.EOF Then
'                            LSSaudaCode = TRec!SAUDACODE
'                            LSMaturity = TRec!MATURITY
'                            LSaudaID = TRec!SAUDAID
'                        Else
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT SAUDACODE,MATURITY FROM SCRIPTMASTER WHERE ITEMCODE='" & LSItemCode & "' AND INSTTYPE='" & LSInstType & "'"
'                            MYSQL = MYSQL & " AND MONTH(MATURITY)=" & month(LSMaturity) & " AND YEAR(MATURITY)=" & Year(LSMaturity) & ""
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSMaturity = TRec!MATURITY
'                                LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                                LSaudaID = Get_SaudaID(LSSaudaCode)
'                            Else
'                                LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                                LSaudaID = Get_SaudaID(LSSaudaCode)
'                            End If
'                        End If
'                        If InStr(LExchange, LSExCode) < 1 Then
'                            If LenB(LExchange) > 0 Then LExchange = LExchange & ","
'                            LExchange = LExchange & "'" & LSExCode & "'"
'                        End If
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        Call Get_ExDetail(LSExCode)
'                        MSParty = LExCont
'                        If LConType = "B" Then
'                            LPtyContype = "B"
'                        Else
'                            LPtyContype = "S"
'                        End If
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'                        If LBQty <> 0 Then
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            If LenB(MBParty) < 1 Then MBParty = NewParty(LSExCode, LBParty, LPartyName, LBParty, True)
'                            If LenB(MBParty) < 1 Then GoTo EFlag_Next
'                            LContime = Time
'                            MCount = MCount + 1:
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            LOConNo = Trim(Str(MCount)):
'                            If InStr(LBillExId, Str(LExID)) = 0 Then
'                                If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                LBillExId = LBillExId & Str(LExID)
'                            End If
'                                        If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'
'                            Call Add_To_Ctr_D(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, LExcelType, "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    ElseIf LExcelType = 2 Then
'                        LLot = 1
'                        MYSQL = "SELECT ITEMCODE,EXCODE,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL ='" & LExhCode & "'"
'                        Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                        If Not TRec.EOF Then
'                            LSExCode = TRec!EXCODE
'                            LSItemCode = TRec!ITEMCODE
'                            LLot = TRec!LOT
'                            LExID = Get_ExID(LSExCode)
'                            LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'
'                            LSSauda = Create_SSaudaMast(LExhCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LSItemCode)
'                            LItemID = Get_ITEMID(LSItemCode)
'                            If LenB(LSItemCode) > 0 Then LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                            If InStr(LExchange, LSExCode) < 1 Then
'                                If LenB(LExchange) > 0 Then LExchange = LExchange & ","
'                                LExchange = LExchange & "'" & LSExCode & "'"
'                            End If
'                            If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                            LSaudaID = Get_SaudaID(LSSaudaCode)
'                            Call Get_ExDetail(LSExCode)
'                            MSParty = LExCont
'                            If LConType = "B" Then
'                                LPtyContype = "B"
'                            Else
'                                LPtyContype = "S"
'                            End If
'                            If LSCondate > LSMaturity Then
'                                MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                                'GoTo EFlag_Next
'                            End If
'                            If LBQty <> 0 Then
'                                MBParty = Get_ACCT_EX(LExID, LBParty)
'                                If LenB(MBParty) < 1 Then GoTo EFlag_Next
'                                LContime = Time
'                                MCount = MCount + 1:
'                                LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                                LOConNo = Trim(Str(MCount)):
'                                If InStr(LBillExId, Str(LExID)) = 0 Then
'                                    If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                    LBillExId = LBillExId & Str(LExID)
'                                End If
'                                            If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'
'                                Call Add_To_Ctr_D(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, "0", "Y", LExID, LItemID, LSaudaID)
'                            End If
'                        End If
'                    ElseIf LExcelType = 10 Then
'                        LLot = 1
'                        'MYSQL = "SELECT EXID,ITEMID,ITEMCODE,EXCHANGECODE,LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE ='" & LExhCode & "'"
'                        MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "','" & vbNullString & "'"
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If TRec.EOF Then
'                            MYSQL = "SELECT ITEMCODE,EXCODE,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL ='" & LExhCode & "'"
'                            Set TRec = Nothing:
'                            Set TRec = New ADODB.Recordset
'                            TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                            If TRec.EOF Then
'                                If Right(TxtRec!F10, 1) = "K" Then LSExCode = "NSE"
'                                LSItemCode = Create_ItemMast(LExhCode, LExhCode, LExhCode, LLot, LSExCode, LExID)
'                                LItemID = Get_ITEMID(LSItemCode)
'                                LExID = Get_ExID(LSExCode)
'                            Else
'                                LSExCode = TRec!EXCODE
'                                LExID = Get_ExID(LSExCode)
'                                LSItemCode = TRec!ITEMCODE
'                                LExID = Get_ExID(LSExCode)
'                                LLot = TRec!LOT
'                                LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'                                LItemID = Get_ITEMID(LSItemCode)
'                            End If
'                        Else
'                            LSExCode = TRec!EXCHANGECODE:   LSItemCode = TRec!ITEMCODE
'                            LLot = TRec!LOT:                LExID = TRec!EXID
'                            LItemID = TRec!ITEMID
'                        End If
'                        If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                        If A = 0 Then
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT TOP 1 SAUDAID,SAUDACODE,MATURITY FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND ITEMID=" & LItemID & " AND INSTTYPE='FUT'"
'                            MYSQL = MYSQL & " AND MATURITY >='" & Format(LSMaturity, "YYYY/MM/DD") & " 'ORDER BY MATURITY"
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSSaudaCode = TRec!SAUDACODE
'                                LSMaturity = TRec!MATURITY
'                                LSaudaID = TRec!SAUDAID
'                            Else
'                                Set TRec = Nothing
'                                Set TRec = New ADODB.Recordset
'                                MYSQL = "SELECT TOP 1 SAUDACODE,MATURITY FROM SCRIPTMASTER WHERE ITEMCODE='" & LSItemCode & "' AND EXCODE='" & LSExCode & "' AND INSTTYPE='FUT'"
'                                MYSQL = MYSQL & " AND MATURITY>='" & Format(LSMaturity, "YYYY/MM/DD") & "' ORDER BY MATURITY "
'                                TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                                If Not TRec.EOF Then
'                                    LSMaturity = TRec!MATURITY
'                                    LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                                Else
'                                    If LSExCode = "CMX" Then
'                                        LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                                    Else
'                                        MsgBox "Please Import New Contracts " & LSItemCode & ""
'                                        Cnn.RollbackTrans
'                                        CNNERR = False
'                                        Exit Sub
'                                    End If
'                                End If
'                                LSaudaID = Get_SaudaID(LSSaudaCode)
'                            End If
'                        Else
'                            Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                            MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",''," & month(LSMaturity) & "," & Year(LSMaturity) & ",'FUT'" & ",'',0"
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSSaudaCode = TRec!SAUDACODE
'                                LSMaturity = TRec!MATURITY
'                                LSaudaID = TRec!SAUDAID
'                            Else
'                                Set TRec = Nothing
'                                Set TRec = New ADODB.Recordset
'                                MYSQL = "SELECT SAUDACODE,MATURITY FROM SCRIPTMASTER WHERE ITEMCODE='" & LSItemCode & "' AND INSTTYPE='" & LSInstType & "'"
'                                MYSQL = MYSQL & " AND MONTH(MATURITY)=" & month(LSMaturity) & " AND YEAR(MATURITY)=" & Year(LSMaturity) & ""
'                                TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                                If Not TRec.EOF Then
'                                    LSMaturity = TRec!MATURITY
'                                    LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                                Else
'                                    LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                                End If
'                                LSaudaID = Get_SaudaID(LSSaudaCode)
'                            End If
'                        End If
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'                        If LBQty <> 0 Then
'                            MBParty = NewParty(LSExCode, LBParty, LPartyName, LBParty, True)
'                            MSParty = Get_ACCT_EX(LExID, LSParty)
'                            If LenB(MBParty) < 1 Then
'                                MsgBox " Client does Not Exist " & LBParty & ""
'                                GoTo EFlag_Next
'                            End If
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT BROKER FROM EXBROKCLIENT WHERE COMPCODE =" & GCompCode & " AND CLIENT ='" & MBParty & "'"
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then MSParty = TRec!BROKER
'                            If LenB(MSParty) < 1 Then
'                                Call Get_ExDetail(LSExCode)
'                                MSParty = LExCont
'                            End If
'                            If Chk_TradeNo(LSSaudaCode, LOConNo, LSCondate) = False Then
'                                MCount = MCount + 1:
'                                If InStr(LBillExId, Str(LExID)) = 0 Then
'                                    If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                    LBillExId = LBillExId & Str(LExID)
'                                End If
'                                If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                                LNoTrd = LNoTrd + 1
'                                GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                                DoEvents
'
'
'                                LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                                Call Add_To_Ctr_D2(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, Trim(Str(LExcelType)), "Y", LExID, LItemID, LSaudaID)
'                            End If
'                        End If
'                    ElseIf (LExcelType = 11 Or LExcelType = 12) Then
'                        LLot = 1
'                        'MYSQL = "SELECT EXID,ITEMID,ITEMCODE,EXCHANGECODE,LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND EXHCODE  ='" & LExhCode & "'"
'                        MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "','" & vbNullString & "'"
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If TRec.EOF Then
'                            MYSQL = "SELECT ITEMCODE,EXCODE,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL ='" & LExhCode & "'"
'                            Set TRec = Nothing:
'                            Set TRec = New ADODB.Recordset
'                            TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                            If TRec.EOF Then
'                                LExID = Get_ExID(LSExCode)
'                                LSItemCode = Create_ItemMast(LExhCode, LExhCode, LExhCode, LLot, LSExCode, LExID)
'                            Else
'                                LSExCode = TRec!EXCODE
'                                LSItemCode = TRec!ITEMCODE
'                                LLot = TRec!LOT
'                                LExID = Get_ExID(LSExCode)
'                                LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'                                LItemID = Get_ITEMID(LSItemCode)
'                            End If
'                        Else
'                            LSExCode = TRec!EXCHANGECODE:   LSItemCode = TRec!ITEMCODE
'                            LLot = TRec!LOT:                LExID = TRec!EXID
'                            LItemID = TRec!ITEMID
'                        End If
'                        If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                        Set TRec = Nothing
'                        Set TRec = New ADODB.Recordset
'                        MYSQL = "SELECT TOP 1 SAUDAID,SAUDACODE,MATURITY FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND ITEMID=" & LItemID & ""
'                        MYSQL = MYSQL & " AND MATURITY >='" & Format(LSMaturity, "YYYY/MM/DD") & "' ORDER BY MATURITY"
'                        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                        If Not TRec.EOF Then
'                            LSSaudaCode = TRec!SAUDACODE
'                            LSMaturity = TRec!MATURITY
'                            LSaudaID = TRec!SAUDAID
'                        Else
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT SAUDACODE,MATURITY FROM SCRIPTMASTER WHERE ITEMCODE='" & LSItemCode & "' AND EXCODE='" & LSExCode & "'"
'                            MYSQL = MYSQL & " AND MATURITY>='" & Format(LSMaturity, "YYYY/MM/DD") & "' ORDER BY MATURITY "
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSMaturity = TRec!MATURITY
'                                LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                                LSaudaID = Get_SaudaID(LSSaudaCode)
'                            Else
'                                MsgBox "Please Import New Contracts " & LSItemCode & ""
'                                Cnn.RollbackTrans
'                                CNNERR = False
'                                Exit Sub
'                            End If
'                        End If
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'                        If LBQty <> 0 Then
'                            MBParty = NewParty(LSExCode, LBParty, LPartyName, LBParty, True)
'                            MSParty = LExCont
'                            If LenB(MBParty) < 1 Then
'                                MsgBox " Client does Not Exist " & LBParty & ""
'                                GoTo EFlag_Next
'                            End If
'                            MYSQL = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND  ROWNO1='" & LOConNo & "' AND CONDATE ='" & Format(LSCondate, "YYYY/MM/DD") & "'"
'                            Cnn.Execute MYSQL
'                            MCount = MCount + 1:
'                            If InStr(LBillExId, Str(LExID)) = 0 Then
'                                If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                LBillExId = LBillExId & Str(LExID)
'                            End If
'                            If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            Call Add_To_Ctr_D2(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, Trim(Str(LExcelType)), "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    ElseIf LExcelType = 20 Or LExcelType = 16 Or LExcelType = 21 Then
'                        LLot = 1
'                        'MYSQL = "SELECT ITEMID,EXID,ITEMCODE,EXCHANGECODE,LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND EXHCODE='" & LExhCode & "' AND EXCHANGECODE='" & LSExCode & "' "
'                        MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "','" & vbNullString & "'"
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If TRec.EOF Then
'                            MYSQL = "SELECT ITEMCODE,EXCODE,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL ='" & LExhCode & "'AND EXCODE='" & LSExCode & "'"
'                            Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                            TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                            If TRec.EOF Then
'                                MsgBox "Please Import/Create New Item/Script " & LExhCode & ""
'                                GoTo EFlag_Next
'                            Else
'                                LSExCode = TRec!EXCODE
'                                LSItemCode = TRec!ITEMCODE
'                                LLot = TRec!LOT
'                                LExID = Get_ExID(LSExCode)
'                                LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'                                LItemID = Get_ITEMID(LSItemCode)
'                            End If
'                        Else
'                            LSExCode = TRec!EXCHANGECODE:   LSItemCode = TRec!ITEMCODE
'                            LLot = TRec!LOT:                LExID = TRec!EXID:                            LItemID = TRec!ITEMID
'                        End If
'                        Set TRec = Nothing
'                        Set TRec = New ADODB.Recordset
'                        If LExcelType = 21 Then
'                            MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",'" & Format(LSMaturity, "YYYY/MM/DD") & "',0,0,'& lsinsttype &','" & LOptType & "'," & LStrike & ""
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                        Else
'                            MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",''," & month(LSMaturity) & "," & Year(LSMaturity) & ",'" & LSInstType & "','" & LOptType & "'," & LStrike & ""
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                        End If
'                        If Not TRec.EOF Then
'                            LSSaudaCode = TRec!SAUDACODE
'                            LSMaturity = TRec!MATURITY
'                            LSaudaID = TRec!SAUDAID
'                        Else
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT SAUDACODE,MATURITY FROM SCRIPTMASTER WHERE ITEMCODE='" & LSItemCode & "' AND EXCODE='" & LSExCode & "'"
'                            If LSInstType <> "FUT" Then
'                                MYSQL = MYSQL & " AND MATURITY='" & Format(LSMaturity, "YYYY/MM/DD") & "' AND INSTTYPE ='OPT' "
'                                MYSQL = MYSQL & " AND OPTTYPE='" & LOptType & "' AND STRIKEPRICE =" & LStrike & ""
'                            Else
'                                MYSQL = MYSQL & " AND INSTTYPE='FUT' AND MONTH(MATURITY)=" & month(LSMaturity) & "AND YEAR(MATURITY)=" & Year(LSMaturity) & ""
'                            End If
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSSaudaCode = TRec!SAUDACODE
'                                LSMaturity = TRec!MATURITY
'                                If LenB(LSItemCode) <> 0 Then LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                            Else
'                                MsgBox " Please Import New Contract" & LSItemCode & " " & LSMaturity & ""
'                                Cnn.RollbackTrans
'                                Exit Sub
'                            End If
'                            LSaudaID = Get_SaudaID(LSSaudaCode)
'                        End If
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        LSaudaID = Get_SaudaID(LSSaudaCode)
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'
'                        If LBQty <> 0 Then
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            If LExcelType = 21 Then
'                                LSParty = Get_ACCT_EX(LExID, LSParty)
'                            End If
'                            If LenB(MBParty) < 1 Then MBParty = NewParty(LSExCode, LBParty, LPartyName, LBParty, True)
'                            If LenB(MBParty) < 1 Then
'                                MsgBox " Client does Not Exist " & LBParty & ""
'                                GoTo EFlag_Next
'                            End If
'                            MSParty = Get_AccountMCode(LSParty)
'                            If LenB(MSParty) < 1 Then
'                                MsgBox " Broker does Not Exist " & LSParty & ""
'                                GoTo EFlag_Next
'                            End If
'                                'LContime = Time
'                            MCount = MCount + 1:
'                           If InStr(LBillExId, Str(LExID)) = 0 Then
'                                If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                LBillExId = LBillExId & Str(LExID)
'                            End If
'                            If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            'LOConNo = Trim(Str(MCount)):
'                            Call Add_To_Ctr_D2(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, Trim(Str(LExcelType)), "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    ElseIf LExcelType = 22 Then
'                        LLot = 1
'                        'MYSQL = "SELECT ITEMID,EXID,ITEMCODE,EXHCODE,EXCHANGECODE,LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE ='" & LExhCode & "' AND EXCHANGECODE='" & LSExCode & "' "
'                        MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "','" & vbNullString & "'"
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If TRec.EOF Then
'                            MsgBox "Please Create New Commodity  " & LExhCode & ""
'                            Cnn.RollbackTrans
'                            CNNERR = False
'                            Exit Sub
'                        Else
'                            LSExCode = TRec!EXCHANGECODE:   LSItemCode = TRec!ITEMCODE
'                            LEx_Symbol = TRec!EXHCODE:      LExID = TRec!EXID
'                            LItemID = TRec!ITEMID
'                            LLot = TRec!LOT
'                        End If
'                        Set TRec = Nothing
'                        Set TRec = New ADODB.Recordset
'                        MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",''," & month(LSMaturity) & "," & Year(LSMaturity) & ",'" & LSInstType & "','" & LOptType & "'," & LStrike & ""
'                        TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                        If Not TRec.EOF Then
'                            LSSaudaCode = TRec!SAUDACODE
'                            LSMaturity = TRec!MATURITY
'                            LSaudaID = TRec!SAUDAID
'                        Else
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT SAUDACODE,MATURITY FROM SCRIPTMASTER WHERE EX_SYMBOL ='" & LEx_Symbol & "' AND EXCODE='" & LSExCode & "'"
'                            If LSInstType <> "FUT" Then
'                                MYSQL = MYSQL & " AND MATURITY='" & Format(LSMaturity, "YYYY/MM/DD") & "' AND INSTTYPE ='OPT' "
'                                MYSQL = MYSQL & " AND OPTTYPE='" & LOptType & "' AND STRIKEPRICE =" & LStrike & ""
'                            Else
'                                MYSQL = MYSQL & " AND INSTTYPE='FUT' AND MONTH(MATURITY)=" & month(LSMaturity) & "AND YEAR(MATURITY)=" & Year(LSMaturity) & ""
'                            End If
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSSaudaCode = TRec!SAUDACODE
'                                LSMaturity = TRec!MATURITY
'                                If LenB(LSItemCode) <> 0 Then LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                            Else
'                                MsgBox " Please Import New Contract" & LSItemCode & " " & LSMaturity & ""
'                                Cnn.RollbackTrans
'                                Exit Sub
'                            End If
'                            LSaudaID = Get_SaudaID(LSSaudaCode)
'                        End If
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'
'                        If LBQty <> 0 Then
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            If LenB(MBParty) < 1 Then
'                                MsgBox " Client does Not Exist " & LBParty & ""
'                                GoTo EFlag_Next
'                            End If
'                           If InStr(LBillExId, Str(LExID)) = 0 Then
'                                If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                LBillExId = LBillExId & Str(LExID)
'                            End If
'                            If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'                            MCount = MCount + 1:
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            Call Add_To_Ctr_D2(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, Trim(Str(LExcelType)), "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    ElseIf LExcelType = 42 Then
'                        LLot = 1
'                        LEx_Symbol = LExhCode
'                        MYSQL = "SELECT EXID,ITEMID ,ITEMCODE,EXCHANGECODE,LOT,EXHCODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMNAME ='" & LExhCode & "'"
'                        'MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "'," & vbNullString & "'"
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If TRec.EOF Then
'                            MYSQL = "SELECT ITEMCODE,EXCODE,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL ='" & LExhCode & "'"
'                            Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                            TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                            If TRec.EOF Then
'                                MsgBox "Please Import/Create New Item/Script " & LExhCode & ""
'                                GoTo EFlag_Next
'                            Else
'                                LSExCode = TRec!EXCODE:     LSItemCode = TRec!ITEMCODE
'                                LEx_Symbol = LExhCode:      LExID = Get_ExID(LSExCode)
'                                LLot = TRec!LOT
'                                LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'                                LItemID = Get_ITEMID(LSItemCode)
'                            End If
'                        Else
'                            LSExCode = TRec!EXCHANGECODE:                            LSItemCode = TRec!ITEMCODE
'                            LEx_Symbol = TRec!EXHCODE:                               LExID = TRec!EXID
'                            LItemID = TRec!ITEMID:                                   LLot = TRec!LOT
'                        End If
'                        If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                        Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                        LSSaudaCode = vbNullString
'                        MYSQL = "SELECT TOP 1 SAUDACODE,SAUDAID ,MATURITY FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND ITEMID =" & LItemID & " "
'                        MYSQL = MYSQL & " AND EXID =" & LExID & " AND MATURITY>='" & Format(LSCondate, "YYYY/MM/DD") & "' ORDER BY MATURITY"
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If Not TRec.EOF Then
'                            LSSaudaCode = TRec!SAUDACODE
'                            LSMaturity = TRec!MATURITY
'                            LSaudaID = TRec!SAUDAID
'                        End If
'                        If LenB(LSSaudaCode) < 1 Then
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT TOP 1 MATURITY FROM SCRIPTMASTER WHERE EX_SYMBOL ='" & LEx_Symbol & "' AND EXCODE='" & LSExCode & "'"
'                            MYSQL = MYSQL & " AND INSTTYPE='FUT' AND MATURITY>='" & Format(LSCondate, "yyyy/mm/dd") & "' ORDER BY MATURITY "
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSMaturity = TRec!MATURITY
'                                LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                            Else
'                                MsgBox "Please Import new contracts  for " & LSItemCode & ""
'                                GoTo EFlag_Next
'                            End If
'                        End If
'                        LSaudaID = Get_SaudaID(LSSaudaCode)
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'
'                        If LBQty <> 0 Then
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            If LenB(MBParty) < 1 Then
'                                MBParty = NewParty(LSExCode, LBParty, LPartyName, LBParty, True)
'                            End If
'                            MSParty = LSParty
'                            If LenB(MBParty) < 1 Then
'                                MsgBox " Client does Not Exist " & LBParty & ""
'                                GoTo EFlag_Next
'                            End If
'                            MYSQL = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE= '" & Format(LSCondate, "YYYY/MM/DD") & "' AND ROWNO1 ='" & LOConNo & "'"
'                            Cnn.Execute MYSQL
'                           If InStr(LBillExId, Str(LExID)) = 0 Then
'                                    If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                    LBillExId = LBillExId & Str(LExID)
'                                End If
'
'                            MCount = MCount + 1:
'                            If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            Call Add_To_Ctr_D2(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, Trim(Str(LExcelType)), "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    ElseIf LExcelType = 25 Or LExcelType = 29 Or LExcelType = 37 Or LExcelType = 38 Or LExcelType = 39 Or LExcelType = 40 Or LExcelType = 41 Then
'                        LLot = 1
'                        LEx_Symbol = LExhCode
'                        If LExcelType = 41 Then
'                            MYSQL = "SELECT EXID,ITEMID ,ITEMCODE,EXCHANGECODE,LOT,EXHCODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND EXHCODE  ='" & LExhCode & "'"
'                            'MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "'," & vbNullString & "'"
'                        Else
'                            'MYSQL = "SELECT EXID,ITEMID ,ITEMCODE,EXCHANGECODE,LOT,EXHCODE FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE ='" & LExhCode & "'"
'                            MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "','" & vbNullString & "'"
'                        End If
'
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If TRec.EOF Then
'                            MYSQL = "SELECT ITEMCODE,EXCODE,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL ='" & LExhCode & "'"
'                            Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                            TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                            If TRec.EOF Then
'                                MsgBox "Please Import/Create New Item/Script " & LExhCode & ""
'                                GoTo EFlag_Next
'                            Else
'                                LSExCode = TRec!EXCODE
'                                LSItemCode = TRec!ITEMCODE
'                                LEx_Symbol = LExhCode
'                                LExID = Get_ExID(LSExCode)
'                                LLot = TRec!LOT
'                                LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'                                LItemID = Get_ITEMID(LSItemCode)
'                            End If
'                        Else
'                            LSExCode = TRec!EXCHANGECODE:   LSItemCode = TRec!ITEMCODE
'                            LEx_Symbol = TRec!EXHCODE:      LExID = TRec!EXID
'                            LItemID = TRec!ITEMID:          LLot = TRec!LOT
'                        End If
'                        If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                        Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                        MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",''," & month(LSMaturity) & "," & Year(LSMaturity) & ",'" & LSInstType & "','" & LOptType & "'," & LStrike & ""
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If Not TRec.EOF Then
'                            LSSaudaCode = TRec!SAUDACODE
'                            LSMaturity = TRec!MATURITY
'                            LSaudaID = TRec!SAUDAID
'                        Else
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT MATURITY FROM SCRIPTMASTER WHERE EX_SYMBOL ='" & LEx_Symbol & "' AND EXCODE='" & LSExCode & "'"
'                            If LSInstType <> "FUT" Then
'                                MYSQL = MYSQL & " AND MATURITY='" & Format(LSMaturity, "YYYY/MM/DD") & "' AND INSTTYPE ='OPT' "
'                                MYSQL = MYSQL & " AND OPTTYPE='" & LOptType & "' AND STRIKEPRICE =" & LStrike & ""
'                            Else
'                                MYSQL = MYSQL & " AND INSTTYPE='FUT' AND MONTH(MATURITY)=" & month(LSMaturity) & "AND YEAR(MATURITY)=" & Year(LSMaturity) & ""
'                            End If
'                            TRec.Open MYSQL, Cnn, adOpenStatic, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LSMaturity = TRec!MATURITY
'                                LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                            Else
'                                MsgBox "Please Import/Create New Contracts for Item/Script " & LExhCode & " " & LSMaturity & ""
'                                                         GoTo EFlag_Next
'                                'LSSaudaCode = Trim(LSItemCode) & " " & month(LSMaturity) & " " & LYEAR
'                                'If Duplicate_Sauda_Chk(LSSaudaCode) = True Then
'                                '    LSSaudaCode = LSExCode & " " & LSSaudaCode
'                                '    If Duplicate_Sauda_Chk(LSSaudaCode) = True Then
'                                '        MsgBox "Please Check Entry for " & LSSaudaCode & ""
'                                '        GoTo EFlag_Next
'                                '     End If
'                                'End If
'                                'LEXID = Get_ExID(LSExCode)
'                                'Call PInsert_Saudamast(LSSaudaCode, LSSaudaCode, LSItemCode, LSMaturity, 1, 1, 0, LSInstType, LOptType, LStrike, LSExCode, 1, LEXID, LItemID)
'                            End If
'                            LSaudaID = Get_SaudaID(LSSaudaCode)
'                        End If
'                        If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                        If LExcelType = 41 Then
'                            Call Get_ExDetail(LSExCode)
'                            LSParty = LExCont
'                        End If
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'
'                        If LBQty <> 0 Then
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            If LenB(MBParty) < 1 Then
'                                MBParty = NewParty(LSExCode, LBParty, LPartyName, LBParty, True)
'                                If LenB(MBParty) > 1 Then
'                                    If LSExCode = "RACE" Then
'                                        MYSQL = "UPDATE ACCOUNTM SET PTYHEAD =1 WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & MBParty & "'"
'                                        Cnn.Execute MYSQL
'                                    Else
'                                        MYSQL = "UPDATE ACCOUNTM SET PTYHEAD =2 WHERE COMPCODE =" & GCompCode & " AND AC_CODE ='" & MBParty & "'"
'                                        Cnn.Execute MYSQL
'                                    End If
'                                End If
'                            End If
'                            MSParty = LSParty
'                            If LenB(MBParty) < 1 Then
'                                MsgBox " Client does Not Exist " & LBParty & ""
'                                GoTo EFlag_Next
'                            End If
'                            If LenB(MSParty) < 1 Then
'                                MsgBox " Broker does Not Exist " & LSParty & ""
'                                GoTo EFlag_Next
'                            End If
'                            MYSQL = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE= '" & Format(LSCondate, "YYYY/MM/DD") & "' AND ROWNO1 ='" & LOConNo & "'"
'                            Cnn.Execute MYSQL
'                            MCount = MCount + 1:
'                            If InStr(LBillExId, Str(LExID)) = 0 Then
'                                If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                    LBillExId = LBillExId & Str(LExID)
'                                End If
'                            If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            Call Add_To_Ctr_D2(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, Trim(Str(LExcelType)), "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    ElseIf LExcelType = 6 Or LExcelType = 7 Or LExcelType = 15 Or LExcelType = 30 Or LExcelType = 31 Or LExcelType = 32 Then
'                        LLot = 1
'                        'MYSQL = "SELECT EXID,ITEMID,ITEMCODE,EXCHANGECODE,LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND EXHCODE='" & LExhCode & "'"
'                        MYSQL = "EXEC GET_ITEMREC " & GCompCode & ",'" & LExhCode & "',''"
'
'                        Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                        TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                        If TRec.EOF Then
'                            MYSQL = "SELECT ITEMCODE,EXCODE,LOT FROM CONTRACTMASTER WHERE EX_SYMBOL ='" & LExhCode & "'"
'                            Set TRec = Nothing:                            Set TRec = New ADODB.Recordset
'                            TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                            If TRec.EOF Then
'                                MsgBox "Please Import/Create New Item/Script " & LExhCode & ""
'                                GoTo EFlag_Next
'                            Else
'                                LSExCode = TRec!EXCODE
'                                LSItemCode = TRec!ITEMCODE
'                                LLot = TRec!LOT
'                                LExID = Get_ExID(LSExCode)
'                                LSItemCode = Create_ItemMast(LSItemCode, LSItemCode, LExhCode, LLot, LSExCode, LExID)
'                                LItemID = Get_ITEMID(LSItemCode)
'                            End If
'                        Else
'                            LSExCode = TRec!EXCHANGECODE:                            LSItemCode = TRec!ITEMCODE
'                            LExID = TRec!EXID:                            LLot = TRec!LOT
'                            LItemID = TRec!ITEMID
'                        End If
'                        Set TRec = Nothing
'                        LSSauda = Create_SSaudaMast(LExhCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LSItemCode)
'                        If LenB(LSItemCode) <> 0 Then LSSaudaCode = Create_SaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, LOptType, LStrike, LExID, LItemID)
'                        If InStr(LExchange, LSExCode) < 1 Then
'                            If LenB(LExchange) > 0 Then LExchange = LExchange & ","
'                            LExchange = LExchange & "'" & LSExCode & "'"
'                        End If
'                        If LSExCode = "CME" And LExcelType = 30 Then
'                            If LExhCode = "SI" Then LBRate = LBRate / 1000
'                            If LExhCode = "GC" Then LBRate = LBRate / 10
'                            If LExhCode = "CL" Then LBRate = LBRate / 100
'                            If LExhCode = "NG" Then LBRate = LBRate / 1000
'                            If LExhCode = "HG" Then LBRate = LBRate / 10000
'                            LSRate = LBRate
'                        End If
'
'                        If LenB(LSSaudaCode) = 0 Then GoTo EFlag_Next
'                        LSaudaID = Get_SaudaID(LSSaudaCode)
'                        If LConType = "B" Then
'                            LPtyContype = "B"
'                        Else
'                            LPtyContype = "S"
'                        End If
'                        If LExcelType = 32 And LSExCode = "NSE" Then
'                            Set TRec = Nothing
'                            Set TRec = New ADODB.Recordset
'                            MYSQL = "SELECT REFLOT FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND SAUDACODE ='" & LSSaudaCode & "'"
'                            TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                            If Not TRec.EOF Then
'                                LBQty = LBQty * TRec!REFLOT
'                            End If
'                        End If
'                        LBQty = Abs(LBQty)
'                        If LSCondate > LSMaturity Then
'                            MsgBox "Pls Check Trade " & LSSaudaCode & " is already expired "
'                            'GoTo EFlag_Next
'                        End If
'
'                        If LBQty <> 0 Then
'                            MBParty = Get_ACCT_EX(LExID, LBParty)
'                            MSParty = Get_ACCT_EX(LExID, LSParty)
'                            If LenB(MBParty) < 1 Then
'                                MsgBox " Client does Not Exist " & LBParty & ""
'                                GoTo EFlag_Next
'                            End If
'                            If LenB(MSParty) < 1 Then
'                                MsgBox " Broker does Not Exist " & LSParty & ""
'                                GoTo EFlag_Next
'                            End If
'                           If InStr(LBillExId, Str(LExID)) = 0 Then
'                                    If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                    LBillExId = LBillExId & Str(LExID)
'                                End If
'                                If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                            LNoTrd = LNoTrd + 1
'                            GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                            DoEvents
'
'                            MCount = MCount + 1:
'                            LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LExID)
'                            LOConNo = Trim(Str(MCount)):
'                            Call Add_To_Ctr_D2(LPtyContype, LBParty, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LBQty, LBRate, LSRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, Trim(Str(LExcelType)), "Y", LExID, LItemID, LSaudaID)
'                        End If
'                    End If
'EFlag_Next:
'                    TxtRec.MoveNext
'                Wend
'                If LExcelType = "41" Then
'                    MYSQL = "SELECT CONNO,ORDNO FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(LTradeDt, "YYYY/MM/DD") & "' AND EXCODE='NSE' And DATAIMPORT = 1 AND CONCODE <>PARTY "
'                    MYSQL = MYSQL & " ORDER BY ORDNO"
'                    Set TRec = Nothing
'                    Set TRec = New ADODB.Recordset
'                    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                    Do While Not TRec.EOF
'                        LOrdNo = TRec!ORDNO
'                        MCount = TRec!CONNO
'                        MYSQL = "UPDATE CTR_D SET BROKFLAG='N' ,BROKRATE =0 WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(LTradeDt, "YYYY/MM/DD") & "'"
'                        MYSQL = MYSQL & " AND EXCODE='NSE' And DATAIMPORT = 1 AND ORDNO ='" & LOrdNo & "' AND CONNO <>" & MCount & ""
'                        Cnn.Execute MYSQL
'                        Do While LOrdNo = TRec!ORDNO
'                            TRec.MoveNext
'                            If TRec.EOF Then Exit Do
'                        Loop
'                        If TRec.EOF Then Exit Do
'                    Loop
'                End If
'            'End If
'        End If
'    End If
'Next
'    Cnn.CommitTrans
'    CNNERR = False
'    Exit Sub
'ERR1:
'    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
'    If CNNERR = True Then
'
'        Cnn.RollbackTrans: CNNERR = False
'    End If
'End Sub

'Sub Save_AarnaTrade()
'    'date 20/07/2004
'    Dim LTradeDt As Date:           Dim LOverwrite As Boolean:      Dim TxtPath As String:          Dim LFileName As String
'    Dim MCount As Long:             Dim LSItemCode As String:       Dim LConType As String:         Dim LPtyContype As String
'    Dim LSConSno As Long:           Dim LQty As Double:             Dim LRate As Double:            Dim LSInstType As String
'    Dim LOptType As String:         Dim LStrike As Double:          Dim LContime As String:         Dim LOrdNo As String
'    Dim LOrdTime As String:         Dim LSSaudaCode As String:      Dim LSCondate As Date:          Dim LSMaturity As Date
'    Dim ITExchangeCode As String:   Dim LOConNo As String:          Dim LStrDate As String:         Dim LClient As String
'    Dim MParty As String:           Dim MBParty As String:          Dim MSParty As String:          Dim TRec As ADODB.Recordset
'    Dim LLot As Double:             Dim LUserId As String:          Dim LPartyName As String:       Dim LSExCode As String
'    Dim LFileSource As String:      Dim LSaudaID As Long:        Dim LEXID As Integer:           Dim LItemID As Integer
'
'    LOverwrite = True
'    On Error GoTo ERR1
'    'Call GET_JCnn("\SAP;")
'    If LenB(LBrokerCode) < 1 Then
'        MsgBox " Please Select Broker  "
'        BrokerCombo.SetFocus
'        Exit Sub
'    End If
'    'If LenB(LClientCode) < 1 Then
'    '    MsgBox " Please Select Client   "
'    '    ClientCombo.SetFocus
'    '    Exit Sub
'    'End If
'
'    Cnn.BeginTrans
'    CNNERR = True:
'    For LTradeDt = vcDTP1.Value To vcDTP2.Value
'        'TxtPath = "P013_" & LEFT$(CStr(LTradedt), 2) & Mid$(CStr(LTradedt), 4, 2) & Right$(CStr(Year(LTradedt)), 4) & ".CSV"
'        TxtPath = CommonDialog1.FileTitle
'        LFileName = CommonDialog1.FileName
'        LFileSource = Left$(LFileName, (Len(LFileName) - Len(TxtPath)) - 1) & ";"
'        Set Jcnn = Nothing
'        Set Jcnn = New ADODB.Connection
'        Jcnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'        "Data Source=" & LFileSource & _
'        "Extended Properties=""TEXT;HDR=No;IMEX=1;FMT=Delimited"""
'        If Not FileExist(LFileName) Then
'            MsgBox LFileName & "  file not found", vbCritical
'            GoTo flag210
'        Else
'            Set TxtRec = Nothing: Set TxtRec = New ADODB.Recordset
'            MYSQL = "Select * from " & TxtPath & " "
'            TxtRec.Open MYSQL, Jcnn, adOpenStatic, adLockReadOnly, adCmdText
'        End If
'        If Not TxtRec.EOF Then
'            FlagDataTrf = True
'            TxtRec.MoveFirst
'            CNNERR = True
'            MCount = Get_Max_ConNo(LTradeDt, 0)
'            TxtRec.MoveFirst
'            While Not TxtRec.EOF
'                LPtyContype = "B": LOrdTime = vbNullString: LOptType = vbNullString: LStrike = 0: LSInstType = "FUT"
'                LSaudaID = 0
'                MCount = MCount + 1
'                If Trim$(TxtRec!F4 & vbNullString) = "USD" Then
'                    LStrDate = Trim$(TxtRec!f1)
'                    If InStr(LStrDate, "/") = 2 Then
'                        If InStr(3, LStrDate, "/") = 4 Then
'                            LSCondate = DateValue(Mid$(LStrDate, 3, 1) & "/" & Left$(LStrDate, 1) & "/" & Mid$(LStrDate, 5, 4))
'                        Else
'                            LSCondate = DateValue(Mid$(LStrDate, 3, 2) & "/" & Left$(LStrDate, 1) & "/" & Mid$(LStrDate, 6, 4))
'                        End If
'                    Else
'                        LSCondate = DateValue(Left$(LStrDate, 2) & "/" & Mid$(LStrDate, 4, 2) & "/" & Mid$(LStrDate, 7, 4))
'                    End If
'                    LContime = Right$(TxtRec!f1, 10)
'
'                    LSExCode = TxtRec!f5
'                    ITExchangeCode = Trim$(TxtRec!f6)
'                    LOConNo = Trim$(TxtRec!F17):                        LOrdNo = Trim$(TxtRec!F17)
'                    LStrDate = Trim$(TxtRec!f7):
'                    If InStr(LStrDate, "/") = 2 Then
'                        If InStr(3, LStrDate, "/") = 4 Then
'                            LSMaturity = DateValue(Mid$(LStrDate, 3, 1) & "/" & Left$(LStrDate, 1) & "/" & Mid$(LStrDate, 5, 4))
'                        Else
'                            LSMaturity = DateValue(Mid$(LStrDate, 3, 2) & "/" & Left$(LStrDate, 1) & "/" & Mid$(LStrDate, 6, 4))
'                        End If
'                    Else
'                       LSMaturity = DateValue(Left$(LStrDate, 2) & "/" & Mid$(LStrDate, 4, 2) & "/" & Mid$(LStrDate, 7, 4))
'                    End If
'                    LSItemCode = Create_TItemMast(ITExchangeCode, ITExchangeCode, ITExchangeCode, 1, LSExCode)
'                    If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                    Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                    MYSQL = "SELECT ITEMID,EXID,ITEMCODE,LOT FROM ITEMMAST WHERE COMPCODE =" & GCompCode & " AND ITEMCODE ='" & LSItemCode & "'"
'                    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                    If Not TRec.EOF Then
'                        LLot = TRec!LOT
'                        LItemID = TRec!ITEMID
'                        LEXID = TRec!EXID
'                    End If
'
'                    Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                    'MYSQL = "SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND ITEMCODE ='" & LSItemCode & "' "
'                    'MYSQL = MYSQL & " AND MATURITY ='" & Format(LSMaturity, "YYYY/MM/DD") & "' AND INSTTYPE ='" & LSInstType & "'"
'                    MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",'" & Format(LSMaturity, "YYYY/MM/DD") & "',0,0,'FUT'" & ",'',0"
'                    'MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",''," & month(LSMaturity) & "," & Year(LSMaturity) & ",'FUT'" & ",'',0"
'                    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                    If Not TRec.EOF Then
'                        LSSaudaCode = TRec!SAUDACODE
'                    Else
'                        LSSaudaCode = Create_TSaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, vbNullString, 0)
'                    End If
'                    If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                    LSaudaID = Get_SaudaID(LSSaudaCode)
'                    LExCont = LBrokerCode
'                    MSParty = LExCont
'                    LConType = Trim$(TxtRec!F10)
'                    If LConType = "S" Then LPtyContype = "S"
'                    MParty = TxtRec!F2
'                    LClient = MParty:                    LUserId = MParty
'                    LPartyName = MParty
'                    LEXID = Get_ExID(LSExCode)
'                    MBParty = Get_ACCT_EX(LEXID, MParty)
'                    If LenB(MBParty) < 1 Then
'                        MBParty = NewParty(LSExCode, MParty, LPartyName & " " & LSExCode, LUserId, True)
'                        If LenB(MBParty) < 1 Then GoTo EFlag_Next
'                    End If
'
'                    LRate = Val(TxtRec!F12)
'                    LQty = Abs(Val(TxtRec!F11))
'                    MCount = MCount + 1
'                    If InStr(LBillExId, Str(LEXID)) = 0 Then
'                                    If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                    LBillExId = LBillExId & Str(LEXID)
'                                End If
'                                If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                        LNoTrd = LNoTrd + 1
'                        GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                        DoEvents
'
'
'                    LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LEXID)
'                    MYSQL = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND ROWNO1 ='" & LOConNo & "' COLLATE SQL_Latin1_General_CP1_CS_AS AND CONDATE ='" & Format(LSCondate, "YYYY/MM/DD") & "' AND EXCODE ='" & LSExCode & "'"
'                    Cnn.Execute MYSQL
'                    Call Add_To_Ctr_D(LPtyContype, LClient, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LQty, LRate, MSParty, LContime, "", LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, "0", "Y", LEXID, LItemID, LSaudaID)
'                End If
'EFlag_Next:
'                TxtRec.MoveNext
'            Wend
'        End If
'flag210:
'    Next LTradeDt
'    Cnn.CommitTrans
'    CNNERR = False
'    Exit Sub
'ERR1:
'    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
'    If CNNERR = True Then
'
'       Cnn.RollbackTrans: CNNERR = False
'    End If
'End Sub
'
'Sub Save_AarnaClosing()
'    'date 20/07/2004
'    Dim LTradeDt As Date:           'Dim LOverwrite As Boolean
':          Dim TxtPath As String:    Dim LFileName As String
'    Dim MCount As Long:             Dim LSItemCode As String:       Dim LConType As String * 1:     Dim LPtyContype As String * 1
'    Dim LSConSno As Long:           Dim LQty As Double:             Dim LRate As Double:            Dim LSInstType As String * 3
'    Dim LOptType As String:         Dim LStrike As Double:          Dim LContime As String:         Dim LOrdNo As String
'    Dim LOrdTime As String:         Dim LSSaudaCode As String:      Dim LSCondate As Date:          Dim LSMaturity As Date
'    Dim ITExchangeCode As String:   Dim LOConNo As String:          Dim LStrDate As String:         Dim LClient As String
'    Dim MParty As String:           Dim MBParty As String:          Dim MSParty As String:          Dim TRec As ADODB.Recordset
'    Dim LLot As Double:             Dim LUserId As String:          Dim LPartyName As String:       Dim LSExCode As String
'    Dim LFileSource As String:      Dim lOpRate  As Double:         Dim LHighRate  As Double:       Dim LLowRate  As Double
'    Dim LSettleRate  As Double:     Dim LCloseRate As Double:       Dim LSaudaID As Long:: Dim LItemID As Integer
'    Dim LEXID As Integer
'
'    On Error GoTo ERR1
'
'    Cnn.BeginTrans
'    CNNERR = True:
'        LSCondate = vcDTP1.Value
'        TxtPath = CommonDialog1.FileTitle
'        LFileName = CommonDialog1.FileName
'        LFileSource = Left$(LFileName, (Len(LFileName) - Len(TxtPath)) - 1) & ";"
'        Set Jcnn = Nothing
'        Set Jcnn = New ADODB.Connection
'        Jcnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'        "Data Source=" & LFileSource & _
'        "Extended Properties=""TEXT;HDR=No;IMEX=1;FMT=Delimited"""
'        If Not FileExist(LFileName) Then
'            MsgBox LFileName & "  file not found", vbCritical
'            GoTo flag210
'        Else
'            Set TxtRec = Nothing: Set TxtRec = New ADODB.Recordset
'            MYSQL = "Select * from " & TxtPath & " "
'            TxtRec.Open MYSQL, Jcnn, adOpenStatic, adLockReadOnly, adCmdText
'        End If
'        If Not TxtRec.EOF Then
'            FlagDataTrf = True
'            TxtRec.MoveFirst
'            CNNERR = True
'            TxtRec.MoveFirst
'            While Not TxtRec.EOF
'                LSaudaID = 0
'                If TxtRec!f7 = "FUTURES" Then
'                    LSExCode = TxtRec!F3
'                    ITExchangeCode = Trim$(TxtRec!F4)
'                    LSInstType = "FUT"
'                    LOptType = vbNullString: LStrike = 0: LSInstType = "FUT"
'                    LEXID = Get_ExID(LSExCode)
'                    LStrDate = Trim$(TxtRec!f5)
'                    LSMaturity = DateValue(Left$(LStrDate, 2) & "/" & Mid$(LStrDate, 4, 3) & "/" & Right$(LStrDate, 4))
'
'                    LSItemCode = Create_TItemMast(ITExchangeCode, ITExchangeCode, ITExchangeCode, 1, LSExCode)
'                    If LenB(LSItemCode) < 1 Then GoTo EFlag_Next
'                    LItemID = Get_ITEMID(LSItemCode)
'                    Set TRec = Nothing: Set TRec = New ADODB.Recordset
'                    MYSQL = "SELECT SAUDACODE FROM SAUDAMAST WHERE COMPCODE=" & GCompCode & " AND ITEMCODE ='" & LSItemCode & "' "
'                    MYSQL = MYSQL & " AND MATURITY ='" & Format(LSMaturity, "YYYY/MM/DD") & "' AND INSTTYPE ='" & LSInstType & "'"
'                    TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                    MYSQL = "EXEC GET_SAUDAREC " & GCompCode & "," & LItemID & ",'" & Format(LSMaturity, "YYYY/MM/DD") & "',0,0,'FUT'" & ",'',0"
'                    If Not TRec.EOF Then
'                        LSSaudaCode = TRec!SAUDACODE
'                    Else
'                        LSSaudaCode = Create_TSaudaMast(LSItemCode, LSMaturity, LSExCode, LSInstType, vbNullString, 0)
'                    End If
'                    If LenB(LSSaudaCode) < 1 Then GoTo EFlag_Next
'                    LSaudaID = Get_SaudaID(LSSaudaCode)
'                    LRate = Val(TxtRec!F15)
'                    lOpRate = 0: LHighRate = 0: LLowRate = 0:
'                    LSettleRate = LRate:        LCloseRate = LRate
'                    If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                        'LNoTrd = LNoTrd + 1
'                        'GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                        'DoEvents
'
'                    If LCloseRate <> 0 Then
'                        LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LEXID)
'                        MYSQL = "DELETE FROM CTR_R WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(LSCondate, "YYYY/MM/DD") & "' AND EXCODE ='" & LSExCode & "' AND SAUDA ='" & LSSaudaCode & "'"
'                        Cnn.Execute MYSQL
'                        LSaudaID = Get_SaudaID(LSSaudaCode)
'                        Call PInsert_Ctr_R(LSConSno, LSSaudaCode, LSCondate, lOpRate, LHighRate, LLowRate, LSettleRate, LCloseRate, LSExCode, LSItemCode, LSaudaID, LItemID, LEXID)
'                    End If
'                End If
'EFlag_Next:
'                TxtRec.MoveNext
'            Wend
'        End If
'flag210:
'    Cnn.CommitTrans
'    CNNERR = False
'    Exit Sub
'ERR1:
'    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
'    If CNNERR = True Then
'       Cnn.RollbackTrans: CNNERR = False
'    End If
'End Sub
'
'Sub SAVE_NCDEX_Assign(MOVERWRITE As Boolean)
'    Dim TRec As ADODB.Recordset
'    Dim LSCondate As Date:          Dim LSMaturity As Date:             Dim MCount As Long:                 Dim LLot As Double
'    Dim LQty As Double:             Dim LRate As Double:                Dim LStrike As Double:              Dim LUserId As String
'    Dim LSInstType As String:       Dim LOConNo As String:              Dim LSItemCode As String:           Dim LMFLPath  As String:
'    Dim LSItemName As String:       Dim LSSaudaCode As String:          Dim LOrdNo As String:               Dim LOrdTime As String:
'    Dim LContime As String:         Dim ITExchangeCode As String:       Dim LStrMaturity As String:         Dim LPartyName As String
'    Dim LPtyContype As String:      Dim LOptType As String:             Dim MParty As String:               Dim LClient As String
'    Dim LExhCode As String:         Dim LCITEM As String:               Dim LSSauda As String:              Dim MSParty As String
'    Dim LConType As String:         Dim MBParty As String:              Dim LSConSno As Long:               Dim LFileName As String
'    Dim LSExCode  As String:        Dim LSaudaID As Long:            Dim TxtPath As String:
'    Dim LEXID As Integer:           Dim LItemID  As Integer
'
'    On Error GoTo ERR1
'    LSExCode = "NCDX"
'    Dim LTradeDt As Date
'    For LTradeDt = vcDTP1.Value To vcDTP2.Value
'        TxtPath = "NCDEX_AS01_" & LMemberId & CStr(Left$(LTradeDt, 2)) & CStr(Mid(LTradeDt, 4, 2)) & CStr(Year(LTradeDt)) & ".CSV"
'        LFileName = App.Path & "\NCDX\" & TxtPath
'        LMFLPath = App.Path & "\NCDX\"
'        If Not FileExist(LFileName) Then
'            MsgBox LFileName & "  file not found", vbCritical
'        Else
'            Set TxtRec = Nothing: Set TxtRec = New ADODB.Recordset
'            TxtRec.Open "Select * from " & TxtPath & "", Jcnn, adOpenStatic, adLockReadOnly, adCmdText
'        End If
'        MCount = Get_Max_ConNo(LSCondate, 0)
'        TxtRec.MoveFirst
'        While Not TxtRec.EOF
'            DoEvents
'            LQty = 0:                            LRate = 0:                         LSInstType = "FUT":                 LStrike = 0
'            LUserId = vbNullString:              LOConNo = vbNullString:            LSItemCode = vbNullString:          LSItemName = vbNullString
'            LSSaudaCode = vbNullString:          LOrdNo = vbNullString:             LOrdTime = vbNullString:            LContime = vbNullString:
'            ITExchangeCode = vbNullString::      LStrMaturity = vbNullString:       LLot = 1
'            LPtyContype = "B":                   LOptType = vbNullString
'            If LenB(TxtRec!f1) <> 0 Then
'                LSCondate = DateValue(Left$(TxtRec!f1, 2) & "/" & Mid(TxtRec!f1, 3, 2) & "/" & Right$(TxtRec!f1, 4))
'                ITExchangeCode = TxtRec!F11
'                LExhCode = ITExchangeCode
'                LOConNo = Trim(TxtRec!F2):
'                LContime = vbNullString
'                LOrdNo = LOConNo:
'                LStrMaturity = DateValue(Left$(TxtRec!F12, 2) & "/" & Mid(TxtRec!F12, 3, 2) & "/" & Right$(TxtRec!F12, 4))
'                LUserId = vbNullString:
'                LConType = Left$(TxtRec!F17, 1)
'                LRate = Val(TxtRec!F13):
'                LQty = Val(TxtRec!F16)
'                LEXID = Get_ExID(LSExCode)
'                LSItemCode = Get_ItemMaster(LEXID, LExhCode)
'                If LenB(LSItemCode) < 1 Then
'                    LCITEM = Create_CItemMast(LExhCode, LExhCode, LExhCode, LLot, LSExCode)
'                    LSItemCode = Create_ItemMast(LCITEM, LCITEM, LExhCode, LLot, LSExCode, LEXID)
'                End If
'                If LenB(LSItemCode) < 1 Then GoTo FFlag_Next
'                LItemID = Get_ITEMID(LSItemCode)
'                MYSQL = "SELECT SAUDACODE,MATURITY,SAUDAID FROM SAUDAMAST WHERE COMPCODE =" & GCompCode & " AND  INSTTYPE ='FUT' AND MATURITY>'" & Format(LSMaturity, "YYYY/MM/DD") & "' ORDER BY MATURITY "
'                Set TRec = Nothing
'                Set TRec = New ADODB.Recordset
'                TRec.Open MYSQL, Cnn, adOpenForwardOnly, adLockReadOnly
'                If Not TRec.EOF Then
'                    LSSaudaCode = TRec!SAUDACODE
'                    LSaudaID = TRec!SAUDAID
'                End If
'                If LenB(LSSaudaCode) = 0 Then GoTo FFlag_Next
'                LLot = Get_LotSize(LItemID, LSaudaID, LEXID, LLotWise)
'                MSParty = LExCont
'                MParty = TxtRec!F8
'                MCount = MCount + 1
'                If InStr(LBillExId, Str(LEXID)) = 0 Then
'                                    If LenB(LBillExId) > 0 Then LBillExId = LBillExId & ","
'                                    LBillExId = LBillExId & Str(LEXID)
'                                End If
'                        If LMinBillDate > LSCondate Then LMinBillDate = LSCondate
'                        LNoTrd = LNoTrd + 1
'                        GETMAIN.Label1.Caption = "Trade No " & LNoTrd & " " & TxtRec!F9
'                        DoEvents
'
'
'                LSConSno = Get_ConSNo(LSCondate, LSSaudaCode, LSItemCode, LSExCode, LSaudaID, LItemID, LEXID)
'                MYSQL = "DELETE FROM CTR_D WHERE COMPCODE =" & GCompCode & " AND CONDATE ='" & Format(LSCondate, "YYYY/MM/DD") & "' AND EXCODE='NCDX' AND ROWNO1='" & LOConNo & "' "
'                Cnn.Execute MYSQL
'                Call Add_To_Ctr_D(LPtyContype, LClient, LSConSno, LSCondate, MCount, LSSaudaCode, LSItemCode, MBParty, LQty, LRate, MSParty, LContime, LOrdNo, LUserId, LOConNo, LSExCode, LLot, 1, LOrdTime, LSInstType, LOptType, LStrike, "0", "Y", LEXID, LItemID, LSaudaID)
'            End If
'FFlag_Next:
'            TxtRec.MoveNext
'        Wend
'    Next
'Exit Sub
'ERR1:
'    MsgBox err.Description, vbCritical, "Error Number : " & err.Number
'
'End Sub



