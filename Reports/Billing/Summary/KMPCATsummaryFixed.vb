'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST � JPRI-0970 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	KMP
'Script Date: 11/07/2005	By: Subhankar Sarkar & Avra Banerjee
'Requirement		: 	KMP CAT REPORT
'*/
'---------------------------------------------------------------------------------------------------------------------->

'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST � TPAL-0139
'FNS DESIGNER
'Client			:	ALL
'Date: 05/16/2011	By: Syed Waqas Ahmed Shah
'Requirement	: 	The hardcoded DB name should be changed to get DB name from session
'*/
'---------------------------------------------------------------------------------------------------------------------->
Option Explicit On 
Option Strict On

Imports System.Data.OracleClient

Module KMPCATsummaryFixed

    Sub getClaimsKMPCAT(ByRef oRpt As KMPBillingSummaryCATFixed, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String)
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

        Const nSERVICE_FEE As Integer = 1
        Const nFAX_FEE As Integer = 2
        Const nTEMP_FEE As Integer = 3
        Const nESCALATE_FEE As Integer = 4
        Const nVENDOR_REFERRAL_FEE As Integer = 5
        Const nPRINT_FEE As Integer = 6
        Const nCAT_FEE As Integer = 8

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
                "From CALL ,CALL_CLAIM  Where CALL.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
                "AND CALL.STATUS = 'COMPLETED' " & _
                "AND CALL.LOB_CD <> 'INF' " & _
                "AND CALL_CLAIM.CAT_FLG = 'Y' " & _
                "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID

        oConn.ConnectionString = CStr(HttpContext.Current.Session("ConnectionString")).Replace("DSN=", "Data Source=")
        oConn.Open()
        oCmd.CommandText = "ALTER SESSION SET NLS_DATE_FORMAT = 'DD-MON-YYYY HH:MI:SS'"
        oCmd.Connection = oConn
        oCmd.ExecuteNonQuery()

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            nTotalClaimsReceived = CInt(oReader.GetValue(oReader.GetOrdinal("totalCalls")))
        End If
        oReader.Close()

        '   get INF claims
        '  cSQL = "Select count(*) as nTotal From CALL " & _
        '        "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '       "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '      "AND STATUS = 'COMPLETED' " & _
        '      "AND LOB_CD = 'INF' " & _
        '      "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
        '      "Group by LOB_CD"

        '    oCmd.CommandText = cSQL
        '   oReader = oCmd.ExecuteReader()
        '   nTotalINFClaims = 0
        '   oReader.Read()
        '   If oReader.HasRows Then
        '   oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFTotal")
        '  getFees(oConn, oRpt, cAHS_ID, "INF", "I", nSERVICE_FEE, "nINFPriceC")

        '  nTotalINFClaims = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        '  oParamFld.Text = CStr(nTotalINFClaims)
        ' End If
        ' oReader.Close()
        '
        '   get call claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
                "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID(+) " & _
                "AND CALL_CLAIM.CAT_FLG = 'Y' " & _
                "AND (CALL_CALLER.CALLER_TYPE IS NULL OR (SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) <> 'F')) " & _
                "AND (CALL_CLAIM.CLAIM_TYPE IS NULL OR (SUBSTR(CALL_CLAIM.CLAIM_TYPE, 1, 1) <> 'N'))" & _
                "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
                " Group by LOB_CD"

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUCalls")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "C", nCAT_FEE, "nPAUPriceC")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLICalls")
                        getFees(oConn, oRpt, cAHS_ID, "PLI", "C", nSERVICE_FEE, "nPLIPriceC")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRCalls")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "C", nCAT_FEE, "nPPRPriceC")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUCalls")
                        getFees(oConn, oRpt, cAHS_ID, "CAU", "C", nSERVICE_FEE, "nCAUPriceC")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLICalls")
                        getFees(oConn, oRpt, cAHS_ID, "CLI", "C", nSERVICE_FEE, "nCLIPriceC")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPRCalls")
                        getFees(oConn, oRpt, cAHS_ID, "CPR", "C", nSERVICE_FEE, "nCPRPriceC")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORCalls")
                        getFees(oConn, oRpt, cAHS_ID, "WOR", "C", nSERVICE_FEE, "nWORPriceC")
                    Case "CRI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRICalls")
                        getFees(oConn, oRpt, cAHS_ID, "CRI", "C", nSERVICE_FEE, "nCRIPriceC")
                    Case "INF"
                        lIsINFO = True
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                If Not lIsINFO Then
                    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        '   get faxed claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CALL_CLAIM.CAT_FLG = 'Y' " & _
                "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
                "AND CALL_CLAIM.CALL_ID = CALL.CALL_ID + 0 " & _
                "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'F' " & _
                "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
                " Group by LOB_CD"

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "F", nSERVICE_FEE, "nPAUPriceF")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLIFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "PLI", "F", nSERVICE_FEE, "nPLIPriceF")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "F", nSERVICE_FEE, "nPPRPriceF")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "CAU", "F", nSERVICE_FEE, "nCAUPriceF")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLIFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "CLI", "F", nSERVICE_FEE, "nCLIPriceF")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPRFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "CPR", "F", nSERVICE_FEE, "nCPRPriceF")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "WOR", "F", nSERVICE_FEE, "nWORPriceF")
                    Case "CRI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRIFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "CRI", "F", nSERVICE_FEE, "nCRIPriceF")
                    Case "INF"
                        lIsINFO = True
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                If Not lIsINFO Then
                    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        '   get internet claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND CALL.STATUS = 'COMPLETED' " & _
             "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
             "AND CALL_CLAIM.CAT_FLG = 'Y' " & _
             "AND SUBSTR(CALL_CLAIM.CLAIM_TYPE, 1, 1) = 'N' " & _
             "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & _
             " Group by LOB_CD"

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUInternet")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "N", nSERVICE_FEE, "nPAUPriceI")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLIInternet")
                        getFees(oConn, oRpt, cAHS_ID, "PLI", "N", nSERVICE_FEE, "nPLIPriceI")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRInternet")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "N", nSERVICE_FEE, "nPPRPriceI")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUInternet")
                        getFees(oConn, oRpt, cAHS_ID, "CAU", "N", nSERVICE_FEE, "nCAUPriceI")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLIInternet")
                        getFees(oConn, oRpt, cAHS_ID, "CLI", "N", nSERVICE_FEE, "nCLIPriceI")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPRInternet")
                        getFees(oConn, oRpt, cAHS_ID, "CPR", "N", nSERVICE_FEE, "nCPRPriceI")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORInternet")
                        getFees(oConn, oRpt, cAHS_ID, "WOR", "N", nSERVICE_FEE, "nWORPriceI")
                    Case "CRI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRIInternet")
                        getFees(oConn, oRpt, cAHS_ID, "CRI", "N", nSERVICE_FEE, "nCRIPriceI")
                    Case "INF"
                        lIsINFO = True
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                If Not lIsINFO Then
                    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        '   get temped claims
        '  cSQL = "Select count(*) as nTotal From CALL, CALL_CLAIM " & _
        '      "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '    "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '    "AND STATUS = 'COMPLETED' " & _
        '    "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
        '    "AND CALL_CLAIM.CAT_FLG = 'Y' " & _
        '    "AND TEMPEDPOLICY_FLG = 'Y' " & _
        '    "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID


        ' oCmd.CommandText = cSQL
        ' oReader = oCmd.ExecuteReader()
        ' oReader.Read()

        ' If oReader.HasRows Then
        '   getProcessingFees(oConn, oRpt, cAHS_ID, nTEMP_FEE, "nTempPrice")
        ' oParamFld = oRpt.DataDefinition.FormulaFields.Item("nTempTotal")
        ' oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        ' End If
        'oReader.Close()
        '***************************************************
        '   get escalations
        cSQL = "Select count(*) as nTotal From CALL, ESCALATION_OUTCOME ,CALL_CLAIM  " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND CALL.STATUS = 'COMPLETED' " & _
                "AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID " & _
                "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
                "AND CALL_CLAIM.CAT_FLG = 'Y' " & _
                "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
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
                "AND TOC.CALL_ID IN (Select DISTINCT CALL.CALL_ID From CALL ,CALL_CLAIM  " & _
                "Where CALL.CALL_ID = CALL_CLAIM.CALL_ID AND CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_CLAIM.CAT_FLG = 'Y' " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & ")"

        oCmd.CommandText = cSQL
        nTotalTransmissions = 0
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            nTotalTransmissions = CInt(oReader.GetValue(oReader.GetOrdinal("transmissionCount")))
        End If
        oReader.Close()
        '   get faxed pages
        cSQL = "Select Sum(PAGE_COUNT) AS pageCount From TRANSMISSION_OUTCOME_STEP " & _
            "Where TRANSMISSION_TYPE_ID = 1 " & _
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
        'TotalFaxedPages = CInt(oRS.Fields("pageCount").Value)
        'Else
        '    nTotalFaxedPages = 0
        'End If
        'oRS.Close()
        'nFreeCount = getFaxFreeCount(oConn, cAHS_ID)
        'If (nTotalTransmissions / nTotalClaimsReceived) - nFreeCount <= 0 Then
        'nCountDifference = 0
        'Else
        '    nCountDifference = (nTotalTransmissions / nTotalClaimsReceived) - nFreeCount
        'End If
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nFaxedPagesTotal")
        'oParamFld.Text = CStr((nTotalFaxedPages / nTotalTransmissions) * nCountDifference * nTotalClaimsReceived)
        'nFreeCount = getFaxFreeCount(oConn, cAHS_ID) * nTotalClaimsReceived
        'getProcessingFees(oConn, oRpt, cAHS_ID, nFAX_FEE, "nFaxedPagesPrice")
        'getFaxDescription(oConn, oRpt, cAHS_ID)

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

        '  nINFFreePercentage = getFreeINFOPercent(oConn, cAHS_ID)
        ' nINFFreeCalls = CInt(nTotalClaimsReceived * (nINFFreePercentage / 100))
        '  If nTotalINFClaims > nINFFreeCalls Then
        '   oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFCalls2Bill")
        '  oParamFld.Text = CStr(nTotalINFClaims - nINFFreeCalls)
        '' End If
        '  oConn.Close()
    End Sub

    Sub getFees(ByVal oConn As OracleConnection, _
                ByVal oRpt As KMPBillingSummaryCATFixed, _
                ByVal cAHS_ID As String, _
                ByVal cLOB As String, _
                ByVal cCallType As String, _
                ByVal nFeeTypeId As Integer, _
                ByVal cFormulaName As String)
        Dim cSQL As String
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                " AND CALL_TYPE = '" & cCallType & "'" & _
                "AND LOB_CD = '" & cLOB & "' " & _
                "AND FEE_TYPE_ID = " & nFeeTypeId
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item(cFormulaName)
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("FEE_AMOUNT")))
        End If
        oReader.Close()

    End Sub

    Sub getProcessingFees(ByVal oConn As OracleConnection, _
                            ByVal oRpt As KMPBillingSummaryCATFixed, _
                            ByVal cAHS_ID As String, _
                            ByVal nFeeTypeId As Integer, _
                            ByVal cFormulaName As String)
        Dim cSQL As String
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                "AND FEE_TYPE_ID = " & nFeeTypeId
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item(cFormulaName)
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("FEE_AMOUNT")))
        End If
        oReader.Close()

    End Sub

    Sub getFaxDescription(ByVal oConn As OracleConnection, _
                        ByVal oRpt As KMPBillingSummaryCATFixed, _
                        ByVal cAHS_ID As String)
        Dim cSQL As String
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select DESCRIPTION From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                "AND FEE_TYPE_ID = 2"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
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
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getFaxFreeCount = CInt(oReader.GetValue(oReader.GetOrdinal("FREE_COUNT")))
        End If
        oReader.Close()

    End Function

    ' Function getFreeINFOPercent(ByVal oConn As OracleConnection, _
    '  ByVal cAHS_ID As String) As Integer
    ' Dim cSQL As String
    '  Dim oCmd As New OracleCommand()
    ' Dim oReader As OracleDataReader
    ' Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

    '  getFreeINFOPercent = 0
    '  cSQL = "Select FREE_PERCENTAGE From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
    '   "AND LOB_CD = 'INF'"
    ' oCmd.Connection = oConn
    '  oCmd.CommandText = cSQL
    ' oReader = oCmd.ExecuteReader()
    ' oReader.Read()
    ' If oReader.HasRows Then
    ' getFreeINFOPercent = CInt(oReader.GetValue(oReader.GetOrdinal("FREE_PERCENTAGE")))
    ' End If
    ' oReader.Close()

    ' End Function

End Module
