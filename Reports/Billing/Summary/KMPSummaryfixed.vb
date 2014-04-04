'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – JPRI-0856 
' FNS DESIGNER    
' Client			:	KMP
' Object			:	KMPSummaryfixed.vb
' Script Date: 06/08/2005		Script By: Shweta Vidyarthi
' Work Request/ILog #	:	JPRI-0856
' Requirement		: 	Vendor Referral query is changed to make Vendore referral same as Billing Details report
'--------------------------------------------------------------------------------------------------------------------*/

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

Module KMPsummaryFixed

    Sub getClaimsFixedPricingKMP(ByRef oRpt As KMPBillingSummaryFixed, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String, ByVal lIsCCE As Boolean)
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
        Dim nEmailCount As String

        Const nSERVICE_FEE As Integer = 1
        Const nFAX_FEE As Integer = 2
        Const nTEMP_FEE As Integer = 3
        Const nESCALATE_FEE As Integer = 4
        Const nVENDOR_REFERRAL_FEE As Integer = 5
        Const nPRINT_FEE As Integer = 6
        'line added for issue id MROU-3282 ************************
        Const nCHOICEPOINT_REFERRAL_FEE As Integer = 9

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
                "AND LOB_CD <> 'INF' "
        If lIsCCE Then
            cSQL = cSQL & "AND ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If

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
        cSQL = "Select count(*) as nTotal From CALL " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND LOB_CD = 'INF' "
        If lIsCCE Then
            cSQL = cSQL & "AND ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        cSQL = cSQL & " Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalINFClaims = 0
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFTotal")
            getFees(oConn, oRpt, cAHS_ID, "INF", "I", nSERVICE_FEE, "nINFPriceC")

            nTotalINFClaims = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            oParamFld.Text = CStr(nTotalINFClaims)
        End If
        oReader.Close()
        '
        '   get call claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID(+) " & _
             "AND nvl(CALL_claim.CAt_flg,'N') <> 'Y' " & _
             "AND (CALL_CALLER.CALLER_TYPE IS NULL OR (CALL_CALLER.CALLER_TYPE <> 'FA')) " & _
             "AND (CALL_CLAIM.CLAIM_TYPE IS NULL OR (SUBSTR(CALL_CLAIM.CLAIM_TYPE, 1, 1) <> 'N'))"


        If lIsCCE Then
            cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        cSQL = cSQL & " Group by LOB_CD"

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUCalls")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "C", nSERVICE_FEE, "nPAUPriceC")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLICalls")
                        getFees(oConn, oRpt, cAHS_ID, "PLI", "C", nSERVICE_FEE, "nPLIPriceC")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRCalls")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "C", nSERVICE_FEE, "nPPRPriceC")
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
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER , CALL_CLAIM " & _
               "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
               "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
               "AND STATUS = 'COMPLETED' " & _
               "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
               "AND NVL(CALL_CLAIM.CAT_FLG,'N') <> 'Y' " & _
               "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID(+) " & _
               "AND CALL_CALLER.CALLER_TYPE = 'FA' "


        If lIsCCE Then
            cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        cSQL = cSQL & " Group by LOB_CD"
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
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND nvl(CALL_claim.CAt_flg,'N') <> 'Y' " & _
             "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
            "AND SUBSTR(CALL_CLAIM.CLAIM_TYPE, 1, 1) = 'N' "


        If lIsCCE Then
            cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        cSQL = cSQL & " Group by LOB_CD"
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

        '----------------- Added By Sutapa on 23rd Mar 06--------------------------------------------------------------------*/
        '   get email claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER , CALL_CLAIM " & _
               "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
               "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
               "AND STATUS = 'COMPLETED' " & _
               "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
               "AND NVL(CALL_CLAIM.CAT_FLG,'N') <> 'Y' " & _
               "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID(+) " & _
               "AND CALL_CALLER.CALLER_TYPE = 'EML' "


        If lIsCCE Then
            cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        cSQL = cSQL & " Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                nEmailCount = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUEmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "PAU", nEmailCount, "nPAUEmailDollars")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLIEmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "PLI", nEmailCount, "nPLIEmailDollars")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPREmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "PPR", nEmailCount, "nPPREmailDollars")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUEmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "CAU", nEmailCount, "nCAUEmailDollars")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLIEmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "CLI", nEmailCount, "nCLIEmailDollars")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPREmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "CPR", nEmailCount, "nCPREmailDollars")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWOREmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "WOR", nEmailCount, "nWOREmailDollars")
                    Case "CRI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRIEmail")
                        getEmailFees(oConn, oRpt, cAHS_ID, "CRI", nEmailCount, "nCRIEmailDollars")
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
        '--------------------------------------------------------------------------------------------------------------------*/


        '   get temped claims
        '--------------------------------------------------------------------------------------------------------------------*/
        'BEGIN		{JPRI-0856}	SQL changed for temped calls
        '--------------------------------------------------------------------------------------------------------------------*/

        'cSQL = "Select count(*) as nTotal From CALL, CALL_CLAIM " & _
        '        "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '     "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '     "AND STATUS = 'COMPLETED' " & _
        '    "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
        '    "AND TEMPEDPOLICY_FLG = 'Y' "
        cSQL = "Select count(*) as nTotal From CALL, CALL_CLAIM " & _
               "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
               "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
               "AND STATUS = 'COMPLETED' " & _
               "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
               "AND TEMPEDPOLICY_FLG = 'Y' "






        '"AND nvl(Call_claim.CAt_flg,'N') <> 'Y' " & _
        '--------------------------------------------------------------------------------------------------------------------*/
        'END		{JPRI-0856}	SQL changed for temped calls
        '--------------------------------------------------------------------------------------------------------------------*/

        If lIsCCE Then
            cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nTEMP_FEE, "nTempPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nTempTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()
        '   get escalations
        '--------------------------------------------------------------------------------------------------------------------*/
        'BEGIN		{JPRI-0856}	SQL changed for Escalations
        '--------------------------------------------------------------------------------------------------------------------*/
        'cSQL = "Select count(*) as nTotal From CALL, ESCALATION_OUTCOME " & _
        '        "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '     "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '     "AND CALL.STATUS = 'COMPLETED' " & _
        '    "AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID "
        cSQL = "Select count(*) as nTotal From CALL, ESCALATION_OUTCOME, CALL_CLAIM " & _
               "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
               "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
               "AND CALL.STATUS = 'COMPLETED' " & _
               "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
               "AND nvl(Call_claim.CAt_flg,'N') <> 'Y' " & _
               "AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID "



        '--------------------------------------------------------------------------------------------------------------------*/
        'END		{JPRI-0856}	SQL changed for Escalations
        '--------------------------------------------------------------------------------------------------------------------*/

        If lIsCCE Then
            cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
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
          "AND TOC.CALL_ID IN (Select DISTINCT CALL_ID From CALL " & _
          "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
          "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
          "AND STATUS = 'COMPLETED' "
        If lIsCCE Then
            cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID & ")"
        Else
            cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & ")"
        End If
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

        nINFFreePercentage = getFreeINFOPercent(oConn, cAHS_ID)
        nINFFreeCalls = CInt(nTotalClaimsReceived * (nINFFreePercentage / 100))
        If nTotalINFClaims > nINFFreeCalls Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFCalls2Bill")
            oParamFld.Text = CStr(nTotalINFClaims - nINFFreeCalls)
        End If


        '*********************************************************************************************

        '   Vendor referral
        '--------------------------------------------------------------------------------------------------------------------*/
        'BEGIN		{JPRI-0856}	SQL changed for vendor referral
        '--------------------------------------------------------------------------------------------------------------------*/

        'cSQL = "Select LOB_CD,count(*) as nTotal From CALL C, CALL_CLAIM CC, CALL_ASI CASI, X_CALL_ASI XCASI " & _
        '    "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '    "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '    "AND C.CALL_ID = CC.CALL_ID " & _
        '    "AND C.CALL_ID = CASI.CALL_ID " & _
        '    "AND CASI.CALL_ASI_ID = XCASI.CALL_ASI_ID " & _
        '    "AND XCASI.FIELD LIKE 'ACCEPTED_MITIGATION_FLG%' " & _
        '    "AND C.STATUS = 'COMPLETED' " & _
        '    "AND XCASI.VALUE ='Y'"
        'If lIsCCE Then
        '    cSQL = cSQL & "AND C.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        'Else
        '    cSQL = cSQL & "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        'End If
        'cSQL = cSQL & " Group by LOB_CD"
        'If oReader.HasRows Then
        '        Dim str1LOB As String
        '        str1LOB = CStr(oReader.GetValue(oReader.GetOrdinal("LOB_CD")))
        '                getProcessingFeeskmp(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPrice", str1LOB)
        '                oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotal")
        '                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        '                'following has been added for JPRI -0856
        'End If
        'oReader.Close()

        'cSQL = "Select C.LOB_CD,count(*) as nTotal From CALL C, CALL_CLAIM CC, CALL_VENDOR_REFERRAL CVR " & _
        '    "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '    "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '    "AND C.CALL_ID = CC.CALL_ID " & _
        '    "AND CC.CALL_CLAIM_ID = CVR.CALL_CLAIM_ID " & _
        '    "AND nvl(CC.CAt_flg,'N') <> 'Y' " & _
        '    "AND CVR.REFERRAL_ACCEPTED = 'Y' "


        'If lIsCCE Then
        '    cSQL = cSQL & "AND C.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        'Else
        '    cSQL = cSQL & "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        'End If
        'cSQL = cSQL & " Group by C.LOB_CD"

        'oCmd.CommandText = cSQL
        'oReader = oCmd.ExecuteReader()
        'oReader.Read()
        'If oReader.HasRows Then
        '    Do
        '        Dim str1LOB As String
        '        str1LOB = CStr(oReader.GetValue(oReader.GetOrdinal("LOB_CD")))
        '        Select Case str1LOB
        '            Case "PPR"
        '                getProcessingFeeskmp(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPrice", str1LOB)
        '                oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotal")
        '                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        '            Case "PAU"
        '                getProcessingFeeskmp(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPricePAU", str1LOB)
        '                oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotalPAU")
        '                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        '        End Select
        '    Loop Until Not oReader.Read()
        'End If
        'oReader.Close()


        cSQL = "Select LOB_CD,count(*) as nTotal From CALL C, CALL_CLAIM CC, CALL_ASI CASI, X_CALL_ASI XCASI " & _
            "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
            "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
            "AND C.CALL_ID = CC.CALL_ID " & _
            "AND C.CALL_ID = CASI.CALL_ID " & _
            "AND CASI.CALL_ASI_ID = XCASI.CALL_ASI_ID " & _
            "AND XCASI.FIELD LIKE 'ACCEPTED_MITIGATION_FLG%' " & _
            "AND C.STATUS = 'COMPLETED' " & _
            "AND XCASI.VALUE ='Y'"


        If lIsCCE Then
            cSQL = cSQL & "AND C.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        cSQL = cSQL & " Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Dim str1LOB As String
            str1LOB = CStr(oReader.GetValue(oReader.GetOrdinal("LOB_CD")))
            getProcessingFeeskmp(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPrice", str1LOB)
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()

        '************88888888888888*************'***************************
        '   Vendor referral
        cSQL = "Select LOB_CD, count(*) as nTotalPAU From CALL C, CALL_CLAIM CC, CALL_VENDOR_REFERRAL CVR " & _
            "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
            "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
            "AND C.CALL_ID = CC.CALL_ID " & _
            "AND CC.CALL_CLAIM_ID = CVR.CALL_CLAIM_ID " & _
            "AND CVR.REFERRAL_ACCEPTED = 'Y' " & _
            " AND C.STATUS = 'COMPLETED'" & _
            "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        cSQL = cSQL & " Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Dim str1LOB As String
            str1LOB = CStr(oReader.GetValue(oReader.GetOrdinal("LOB_CD")))
            getProcessingFeeskmp(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPricePAU", str1LOB)
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotalPAU")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotalPAU")))
        End If
        oReader.Close()

        'code added for issue id MROU-3282 ************************
        '   Choice Point referral 
        cSQL = "Select count(*) as nTotal From CALL C, CALL_CLAIM CC " & _
            "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
            "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
            "AND C.CALL_ID = CC.CALL_ID " & _
            "AND CC.CHOICEPNT_ROUTE_FLG = 'Y' " & _
            "AND C.STATUS = 'COMPLETED'" & _
            "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nCHOICEPOINT_REFERRAL_FEE, "nCPPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()
        '********************************************************
        '--------------------------------------------------------------------------------------------------------------------*/
        'END		{JPRI-0856}	SQL changed for vendor referral
        '--------------------------------------------------------------------------------------------------------------------*/


        '************88888888888888*************'***************************
        '   Vendor referral
        '--------------------------------------------------------------------------------------------------------------------*/
        'BEGIN		{JPRI-0856}	SQL changed for vendor referral as for PAU code has been added above
        '--------------------------------------------------------------------------------------------------------------------*/
        'cSQL = "Select LOB_CD, count(*) as nTotalPAU From CALL C, CALL_CLAIM CC, CALL_VENDOR_REFERRAL CVR " & _
        '    "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '    "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '    "AND C.CALL_ID = CC.CALL_ID " & _
        '    "AND CC.CALL_CLAIM_ID = CVR.CALL_CLAIM_ID " & _
        '    "AND CVR.REFERRAL_ACCEPTED = 'Y' " & _
        '    " AND C.STATUS = 'COMPLETED'" & _
        '    "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        'cSQL = cSQL & " Group by LOB_CD"
        'oCmd.CommandText = cSQL
        'oReader = oCmd.ExecuteReader()
        'oReader.Read()
        'If oReader.HasRows Then
        '    Dim str1LOB As String
        '    str1LOB = CStr(oReader.GetValue(oReader.GetOrdinal("LOB_CD")))
        '    getProcessingFeeskmp(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPricePAU", str1LOB)
        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotalPAU")
        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotalPAU")))
        'End If
        'oReader.Close()
        '--------------------------------------------------------------------------------------------------------------------*/
        'END		{JPRI-0856}	SQL changed for vendor referral
        '--------------------------------------------------------------------------------------------------------------------*/
        oConn.Close()
    End Sub

    Sub getFees(ByVal oConn As OracleConnection, _
                ByVal oRpt As KMPBillingSummaryFixed, _
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

    Sub getEmailFees(ByVal oConn As OracleConnection, _
                ByVal oRpt As KMPBillingSummaryFixed, _
                ByVal cAHS_ID As String, _
                ByVal cLOB As String, _
                ByVal cCount As String, _
                ByVal cFormulaName As String)

        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim nCount As Double
        Dim nLowCount As Double
        Dim nMidCount As Double
        Dim nHighCount As Double
        Dim nLowCost As Decimal
        Dim nMidCost As Decimal
        Dim nHighCost As Decimal
        Dim nTotalCost As Decimal

        nCount = CDbl(cCount)
        nLowCount = 0
        nMidCount = 0
        nHighCount = 0
        nLowCost = 0
        nMidCost = 0
        nHighCost = 0

        If (nCount <= 2000) Then
            nLowCount = nCount
            nLowCost = CDec(nLowCount * 9.95)

        ElseIf (nCount > 2000 And nCount <= 5000) Then
            nLowCount = 2000
            nMidCount = nCount - nLowCount
            nLowCost = CDec(nLowCount * 9.95)
            nMidCost = CDec(nMidCount * 9.15)

        ElseIf (nCount > 5000) Then
            nLowCount = 2000
            nMidCount = 3000
            nHighCount = nCount - nLowCount - nMidCount
            nLowCost = CDec(nLowCount * 9.95)
            nMidCost = CDec(nMidCount * 9.15)
            nHighCost = CDec(nHighCount * 8.85)
        End If

        nTotalCost = nLowCost + nMidCost + nHighCost
        oParamFld = oRpt.DataDefinition.FormulaFields.Item(cFormulaName)
        oParamFld.Text = CStr(nTotalCost)
    End Sub


    Sub getProcessingFees(ByVal oConn As OracleConnection, _
                            ByVal oRpt As KMPBillingSummaryFixed, _
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
    Sub getProcessingFeeskmp(ByVal oConn As OracleConnection, _
                            ByVal oRpt As KMPBillingSummaryFixed, _
                            ByVal cAHS_ID As String, _
                            ByVal nFeeTypeId As Integer, _
                            ByVal cFormulaName As String, ByVal cLOB As String)

        Dim cSQL As String
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID
        cSQL = cSQL & " and LOB_CD ='" & cLOB & "'"
        cSQL = cSQL & "  And FEE_TYPE_ID = " & nFeeTypeId
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
                        ByVal oRpt As KMPBillingSummaryFixed, _
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

    Function getFreeINFOPercent(ByVal oConn As OracleConnection, _
                            ByVal cAHS_ID As String) As Integer
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        getFreeINFOPercent = 0
        cSQL = "Select FREE_PERCENTAGE From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                "AND LOB_CD = 'INF'"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getFreeINFOPercent = CInt(oReader.GetValue(oReader.GetOrdinal("FREE_PERCENTAGE")))
        End If
        oReader.Close()

    End Function

End Module
