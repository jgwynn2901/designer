'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – TPAL-0139
'FNS DESIGNER
'Client			:	ALL
'Date: 05/16/2011	By: Syed Waqas Ahmed Shah
'Requirement	: 	The hardcoded DB name should be changed to get DB name from session
'*/
'---------------------------------------------------------------------------------------------------------------------->

Option Explicit On 
Option Strict On

Imports System.Data.OracleClient

Module summaryTieredAFF

    Sub getClaimsTieredPricingAFF(ByRef oRpt As AFFBillingSummaryTiered, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String)
        Dim cSQL As String
        Dim oConn As New OracleConnection
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oValueFld As CrystalDecisions.CrystalReports.Engine.FieldObject
        Dim oDiskOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
        Dim cFileName As String
        Dim dRepStart As Date
        Dim dRepEnd As Date
        Dim cStart, cEnd, cEndLabel As String
        Dim nFreeCount As Integer
        Dim nTotalClaimsReceived As Integer
        Dim nTotalTransmissions As Integer
        Dim nTotalFaxedPages As Integer
        Dim lIsINFO As Boolean
        Dim nTotalINFClaims As Integer
        Dim nINFFreePercentage As Integer
        Dim nINFFreeCalls As Integer
        Dim nCountDifference As Double
        Dim nTotalClaimsCalls As Integer
        Dim nTotalClaimsFaxes As Integer
        Dim nTotalClaimsInt As Integer

        Const nSERVICE_FEE As Integer = 1
        Const nFAX_FEE As Integer = 2
        Const nTEMP_FEE As Integer = 3
        Const nESCALATE_FEE As Integer = 4
        Const nVENDOR_REFERRAL_FEE As Integer = 5
        Const nPRINT_FEE As Integer = 6

        dRepStart = CDate(cReportStartDate)
        dRepEnd = CDate(cReportEndDate)
        cEndLabel = UCase$(Format(dRepEnd, "dd-MMM-yyyy"))
        dRepEnd = DateAdd(DateInterval.Day, 1, dRepEnd)
        cStart = UCase$(Format(dRepStart, "dd-MMM-yyyy"))
        cEnd = UCase$(Format(dRepEnd, "dd-MMM-yyyy"))
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("cPeriodFrom")
        oParamFld.Text = "'" & cStart & "'"
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("cPeriodTo")
        oParamFld.Text = "'" & cEndLabel & "'"
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("cClientName")
        oParamFld.Text = "'" & cClient & "'"

        '   get total number of claims (excluding INF calls)
        cSQL = "Select COUNT(call.call_id) AS totalCalls " & _
                "From CALL Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND LOB_CD <> 'INF' " & _
                "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID

        oConn.ConnectionString = CStr(HttpContext.Current.Session("ConnectionString")).Replace("DSN=", "Data Source=")
        oConn.Open()
        oCmd.CommandText = "ALTER SESSION SET NLS_DATE_FORMAT = 'DD-MON-YYYY HH:MI:SS'"
        oCmd.Connection = oConn
        oCmd.ExecuteNonQuery()
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            nTotalClaimsReceived = CInt(oReader.GetValue(oReader.GetOrdinal("totalCalls")))
        End If
        oReader.Close()
        '   get INF claims
        cSQL = "Select count(*) as nTotal From CALL " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
             " AND LOB_CD = 'INF' " & _
             "Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        nTotalINFClaims = 0
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFTotal")
            getFees(oConn, oRpt, cAHS_ID, "INF", "I", nSERVICE_FEE, "nINFPrice")
            nTotalINFClaims = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            oParamFld.Text = CStr(nTotalINFClaims)
        End If
        oReader.Close()
        '
        '   get call claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
             " AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND (CALL_CALLER.CALLER_TYPE IS NULL OR (SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) <> 'F' " & _
             "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) <> 'N')) " & _
             "Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalClaimsCalls = 0
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUCalls")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLICalls")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRCalls")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUCalls")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLICalls")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPRCalls")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORCalls")
                    Case "CRI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRICalls")
                    Case "INF"
                        lIsINFO = True
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                If Not lIsINFO Then
                    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                    nTotalClaimsCalls = nTotalClaimsCalls + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        '   gert faxed claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
             " AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'F' " & _
                "Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalClaimsFaxes = 0
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUFaxes")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLIFaxes")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRFaxes")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUFaxes")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLIFaxes")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPRFaxes")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORFaxes")
                    Case "INF"
                        lIsINFO = True
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                If Not lIsINFO Then
                    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                    nTotalClaimsFaxes = nTotalClaimsFaxes + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        '   get internet claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
             " AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'N' " & _
                "Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalClaimsInt = 0
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUInternet")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLIInternet")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRInternet")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUInternet")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLIInternet")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPRInternet")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORInternet")
                    Case "INF"
                        lIsINFO = True
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                If Not lIsINFO Then
                    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                    nTotalClaimsInt = nTotalClaimsInt + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        '   get temped claims
        cSQL = "Select count(*) as nTotal From CALL, CALL_CLAIM " & _
                "Where CALL.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND CALL.STATUS = 'COMPLETED' " & _
            "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
            " AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
            "AND CALL.LOB_CD ='PAU'" & _
            "AND CALL_CLAIM.TEMPEDPOLICY_FLG = 'Y'"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nTEMP_FEE, "nTempPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nTempTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()
        '   get escalations
        cSQL = "Select count(*) as nTotal From CALL, ESCALATION_OUTCOME " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND CALL.STATUS = 'COMPLETED' " & _
            "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
            " AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nESCALATE_FEE, "nEscalationPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nEscalationsTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()
        '   get total transmissions
        cSQL = "Select COUNT(TOC.TRANSMISSION_OUTCOME_ID) AS transmissionCount " & _
                "From TRANSMISSION_OUTCOME TOC, TRANSMISSION_OUTCOME_STEP TOS " & _
                "Where TOS.STATUS = 'PROCESSED' " & _
                "AND TOC.TRANSMISSION_OUTCOME_ID = TOS.TRANSMISSION_OUTCOME_ID " & _
                "AND TOS.TRANSMISSION_TYPE_ID = 1 " & _
                "AND TOC.CALL_ID IN (Select DISTINCT CALL_ID From CALL " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID & ")"
        oCmd.CommandText = cSQL
        nTotalTransmissions = 0
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            nTotalTransmissions = CInt(oReader.GetValue(oReader.GetOrdinal("transmissionCount")))
        End If
        oReader.Close()

        '   get faxed pages
        cSQL = "Select NVL(Sum(PAGE_COUNT),0) AS pageCount From TRANSMISSION_OUTCOME_STEP " & _
            "Where TRANSMISSION_TYPE_ID = 1 " & _
            "AND STATUS = 'PROCESSED' " & _
            "AND TRANSMISSION_OUTCOME_STEP.TRANSMISSION_SEQ_STEP_ID IN (SELECT TRANSMISSION_SEQ_STEP_ID + 0 " & _
            "From TRANSMISSION_SEQ_STEP Where Exists (Select 'X' " & _
            "From ROUTING_PLAN RP, CALL " & _
            "Where CLIENT_HRCY_STEP_ID = " & cAHS_ID & " AND CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
            "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
            "AND RP.ACCNT_HRCY_STEP_ID = CALL.ACCNT_HRCY_STEP_ID + 0 " & _
            "AND RP.ROUTING_PLAN_ID = TRANSMISSION_SEQ_STEP.ROUTING_PLAN_ID + 0))"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            nTotalFaxedPages = CInt(oReader.GetValue(oReader.GetOrdinal("pageCount")))
        Else
            nTotalFaxedPages = 0
        End If
        oReader.Close()
        nFreeCount = getFaxFreeCount(oConn, cAHS_ID)
        If (nTotalTransmissions / nTotalClaimsReceived) - nFreeCount <= 0 Then
            nCountDifference = 0
        Else
            nCountDifference = (nTotalTransmissions / nTotalClaimsReceived) - nFreeCount
        End If
        ' oParamFld = oRpt.DataDefinition.FormulaFields.Item("nFaxedPagesTotal")
        'oParamFld.Text = CStr((nTotalFaxedPages / nTotalTransmissions) * nCountDifference * nTotalClaimsReceived)
        nFreeCount = getFaxFreeCount(oConn, cAHS_ID) * nTotalClaimsReceived
        getProcessingFees(oConn, oRpt, cAHS_ID, nFAX_FEE, "nFaxedPagesPrice")
        getFaxDescription(oConn, oRpt, cAHS_ID)

        '   get printed pages
        cSQL = "Select Sum(PAGE_COUNT) AS pageCount From TRANSMISSION_OUTCOME_STEP " & _
            "Where TRANSMISSION_TYPE_ID = 2 " & _
            "AND STATUS = 'PROCESSED' " & _
            "AND TRANSMISSION_OUTCOME_STEP.TRANSMISSION_SEQ_STEP_ID IN (SELECT TRANSMISSION_SEQ_STEP_ID + 0 " & _
            "From TRANSMISSION_SEQ_STEP Where Exists (Select 'X' " & _
            "From ROUTING_PLAN RP, CALL " & _
            "Where CLIENT_HRCY_STEP_ID = " & cAHS_ID & " AND CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
            "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
            "AND RP.ACCNT_HRCY_STEP_ID = CALL.ACCNT_HRCY_STEP_ID + 0 " & _
            "AND RP.ROUTING_PLAN_ID = TRANSMISSION_SEQ_STEP.ROUTING_PLAN_ID + 0))"
        'oRS = oConn.Execute(cSQL)
        'If Not oRS.EOF Then
        'getProcessingFees(oConn, oRpt, cAHS_ID, nPRINT_FEE, "nPrintPrice")
        'oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCallPrint")
        'oParamFld.Text = CStr(oRS.Fields("pageCount").Value)
        'End If
        'oRS.Close()
        '   Vendor referral
        cSQL = "Select count(*) as nTotal From CALL C, CALL_CLAIM CC, CALL_VENDOR_REFERRAL CVR " & _
            "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
            "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
            "AND C.CALL_ID = CC.CALL_ID " & _
            "AND CC.CALL_CLAIM_ID = CVR.CALL_CLAIM_ID " & _
            "AND CVR.REFERRAL_ACCEPTED = 'Y' "
        cSQL = cSQL & "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()

        getTieredFees(oConn, oRpt, cAHS_ID, "C", nSERVICE_FEE, nTotalClaimsCalls)
        getTieredFees(oConn, oRpt, cAHS_ID, "F", nSERVICE_FEE, nTotalClaimsFaxes)
        getTieredFees(oConn, oRpt, cAHS_ID, "N", nSERVICE_FEE, nTotalClaimsInt)
        oConn.Close()
        oConn = Nothing

    End Sub

    Private Sub getTieredFees(ByVal oConn As OracleConnection, _
                ByVal oRpt As AFFBillingSummaryTiered, _
                ByVal cAHS_ID As String, _
                ByVal cCallType As String, _
                ByVal nFeeTypeId As Integer, _
                ByVal nTotalClaims As Integer)

        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim cTierFROM, cTierTO, cTier As String
        Dim x As Integer
        Dim cPriceCalls As String = "nTier~PriceC"
        Dim cPriceFax As String = "nTier~PriceF"
        Dim cPriceInt As String = "nTier~PriceI"
        Dim cPrice As String
        Dim cTotalPriceCalls As String = "nTier~CallDollars"
        Dim cTotalPriceFax As String = "nTier~FaxDollars"
        Dim cTotalPriceInt As String = "nTier~IntDollars"
        Dim cTotalPrice As String
        Dim nBCRColNo As Integer
        Dim nECRColNo As Integer
        Dim nFeeAmntColNo As Integer
        Dim nClaimsFrom, nClaimsTo As Integer

        cTierFROM = "nTier~FROM"
        cTierTO = "nTier~TO"
        x = 1
        cSQL = "Select DISTINCT BEGIN_CALL_RANGE , END_CALL_RANGE ,FEE_AMOUNT  From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                " AND CALL_TYPE = '" & cCallType & "'" & _
                " AND FEE_TYPE_ID = " & nFeeTypeId & _
                " AND (REASON_CODE = '' OR REASON_CODE = '0' OR REASON_CODE IS NULL) " & _
                " ORDER BY BEGIN_CALL_RANGE"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                cTier = cTierFROM
                cTier = Replace(cTier, "~", CStr(x))
                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cTier)
                nBCRColNo = oReader.GetOrdinal("BEGIN_CALL_RANGE")
                nECRColNo = oReader.GetOrdinal("END_CALL_RANGE")
                nClaimsFrom = CInt(oReader.GetValue(nBCRColNo))
                oParamFld.Text = CStr(nClaimsFrom)

                If CInt(oReader.GetValue(nECRColNo)) = 0 Then
                    '   last tier
                    nClaimsTo = 999999
                Else
                    nClaimsTo = CInt(oReader.GetValue(nECRColNo))
                End If
                cTier = cTierTO
                cTier = Replace(cTier, "~", CStr(x))
                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cTier)
                oParamFld.Text = CStr(nClaimsTo)
                '
                If cCallType = "C" Then
                    cPrice = cPriceCalls
                    cTotalPrice = cTotalPriceCalls
                ElseIf cCallType = "F" Then
                    cPrice = cPriceFax
                    cTotalPrice = cTotalPriceFax

                ElseIf cCallType = "N" Then
                    cPrice = cPriceInt
                    cTotalPrice = cTotalPriceInt
                End If
                cPrice = Replace(cPrice, "~", CStr(x))
                cTotalPrice = Replace(cTotalPrice, "~", CStr(x))
                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cPrice)
                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("FEE_AMOUNT")))

                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cTotalPrice)
                nFeeAmntColNo = oReader.GetOrdinal("FEE_AMOUNT")
                If nTotalClaims > nClaimsTo Then
                    oParamFld.Text = CStr(CDbl(oReader.GetValue(nFeeAmntColNo)) * nClaimsTo)
                Else

                    oParamFld.Text = CStr(CDbl(oReader.GetValue(nFeeAmntColNo)) * (nTotalClaims - nClaimsFrom + CInt(IIf((nClaimsFrom = 1), 1, 0))))

                    Exit Do
                End If
                x = x + 1
                If x > 5 Then
                    Exit Do
                End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()

    End Sub

    Private Sub getProcessingFees(ByVal oConn As OracleConnection, _
                            ByVal oRpt As AFFBillingSummaryTiered, _
                            ByVal cAHS_ID As String, _
                            ByVal nFeeTypeId As Integer, _
                            ByVal cFormulaName As String)
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                "AND FEE_TYPE_ID = " & nFeeTypeId
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item(cFormulaName)
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("FEE_AMOUNT")))
        End If
        oReader.Close()

    End Sub

    Private Sub getFaxDescription(ByVal oConn As OracleConnection, _
                        ByVal oRpt As AFFBillingSummaryTiered, _
                        ByVal cAHS_ID As String)
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        cSQL = "Select DESCRIPTION From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                "AND FEE_TYPE_ID = 2"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("cFaxNote")
            oParamFld.Text = "'" & CStr(oReader.GetValue(oReader.GetOrdinal("DESCRIPTION"))) & "'"
        End If
        oReader.Close()
    End Sub

    Function getFaxFreeCount(ByVal oConn As OracleConnection, _
                            ByVal cAHS_ID As String) As Integer
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        getFaxFreeCount = 0
        cSQL = "Select FREE_COUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                "AND FEE_TYPE_ID = 2"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            getFaxFreeCount = CInt(oReader.GetValue(oReader.GetOrdinal("FREE_COUNT")))
        End If
        oReader.Close()

    End Function

    Private Sub getFees(ByVal oConn As OracleConnection, _
            ByVal oRpt As AFFBillingSummaryTiered, _
            ByVal cAHS_ID As String, _
            ByVal cLOB As String, _
            ByVal cCallType As String, _
            ByVal nFeeTypeId As Integer, _
            ByVal cFormulaName As String)
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                " AND CALL_TYPE = '" & cCallType & "'" & _
                "AND LOB_CD = '" & cLOB & "' " & _
                "AND FEE_TYPE_ID = " & nFeeTypeId
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item(cFormulaName)
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("FEE_AMOUNT")))
        End If
        oReader.Close()

    End Sub

End Module
