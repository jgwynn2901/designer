Option Explicit On 
Option Strict On

Imports System.Data.OracleClient

Module AMESummary

    Sub getAMEClaimsFixedPricing(ByRef oRpt As AMEBillingSummary, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String, ByVal lIsCCE As Boolean)
        Dim cSQL As String
        Dim oConn As New OracleConnection
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oReaderAME As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim cClientNameParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
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
        Dim strLOB_AME As String
        Dim SqlStrBuilderAME As System.Text.StringBuilder


        Const nEMAIL_FEE As Integer = 1
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
        cClientNameParamFld = oRpt.DataDefinition.FormulaFields.Item("rClientName")
        cClientNameParamFld.Text = "'" & cClient & "'"

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

        ' Get the connection string from the ASP connection that was passed through 
        ' on the default.aspx page
        ' Change the DSN to be the .Net accepted Data Source
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

        'get INF claims
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
            '***********************
            'This code has been commented out because, INF Calls are not longer shown/calculated separately 
            'INF Calls no longer have call_type set to I
            '***********************
            'oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFTotal")
            'getFees(oConn, oRpt, cAHS_ID, "INF", "I", nSERVICE_FEE, "nINFPriceC")

            nTotalINFClaims = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFTotal")
            oParamFld.Text = CStr(nTotalINFClaims)
        End If
        oReader.Close()

        '   get call claims
        cSQL = "Select LOB_CD, count(distinct CALL.CALL_ID) as nTotal From CALL, CALL_CALLER, USERS " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND CALL.STATUS = 'COMPLETED' " & _
                "AND CALL.CALL_ID = CALL_CALLER.CALL_ID (+) " & _
                "AND (CALL_CALLER.CALLER_TYPE IS NULL " & _
                "OR CALL_CALLER.CALLER_TYPE not in ('FAX','NET','INT','EMAIL','EML', 'IFTCO'))" & _
                "AND USERS.USER_ID = CALL.USER_ID " & _
                "AND USERS.SITE_ID = 5 "

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
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRCalls")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "C", nSERVICE_FEE, "nPPRPriceC")
                    Case "INF"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFCalls")
                        getFees(oConn, oRpt, cAHS_ID, "INF", "C", nSERVICE_FEE, "nINFPriceC")
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            Loop Until Not oReader.Read
        End If
        oReader.Close()


        '------------------------------------------------------------------------------------
        'Get Email Claims

        SqlStrBuilderAME = New System.Text.StringBuilder("")
        SqlStrBuilderAME.Append(" Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER, CALL_CLAIM ")
        SqlStrBuilderAME.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderAME.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderAME.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilderAME.Append(" AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 ")
        SqlStrBuilderAME.Append(" AND call.CALL_ID = call_claim.CALL_ID  ")
        SqlStrBuilderAME.Append(" AND (CALL_CALLER.CALLER_TYPE = 'EML' or CALL_CALLER.CALLER_TYPE = 'EMAIL')")
        SqlStrBuilderAME.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID)
        SqlStrBuilderAME.Append(" Group by LOB_CD")
        oCmd.CommandText = SqlStrBuilderAME.ToString
        oReaderAME = oCmd.ExecuteReader()
        oReaderAME.Read()

        If oReaderAME.HasRows Then
            Do
                strLOB_AME = CStr(oReaderAME.GetValue(oReaderAME.GetOrdinal("LOB_CD")))
                Select Case strLOB_AME
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUEmails")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "E", nEMAIL_FEE, "nPAUPriceE")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPREmails")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "E", nEMAIL_FEE, "nPPRPriceE")
                    Case "INF"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFEmails")
                        getFees(oConn, oRpt, cAHS_ID, "INF", "E", nEMAIL_FEE, "nINFPriceE")
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                oParamFld.Text = CStr(oReaderAME.GetValue(oReaderAME.GetOrdinal("nTotal")))
            Loop Until Not oReaderAME.Read

        End If
        oReaderAME.Close()

        '-------------------------------------------------------------------------------------
        '   get faxed claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
                "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'F' "
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
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "F", nSERVICE_FEE, "nPAUPriceF")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "F", nSERVICE_FEE, "nPPRPriceF")
                    Case "INF"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "INF", "F", nSERVICE_FEE, "nINFPriceF")
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            Loop Until Not oReader.Read
        End If
        oReader.Close()

        '-------------------------------------------------------------------------------------
        '   get over flow claims
        cSQL = "Select LOB_CD, count(distinct CALL.CALL_ID) as nTotal From CALL, CALL_CALLER, USERS " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND CALL.STATUS = 'COMPLETED' " & _
                "AND CALL.CALL_ID = CALL_CALLER.CALL_ID (+) " & _
                "AND (CALL_CALLER.CALLER_TYPE IS NULL " & _
                "OR CALL_CALLER.CALLER_TYPE not in ('FAX','NET','INT','EMAIL','EML', 'IFTCO'))" & _
                "AND USERS.USER_ID = CALL.USER_ID " & _
                "AND USERS.SITE_ID IN (1, 2) "
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
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUOverFlow")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "O", nSERVICE_FEE, "nPAUPriceO")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPROverFlow")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "O", nSERVICE_FEE, "nPPRPriceO")
                    Case "INF"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFOverFlow")
                        getFees(oConn, oRpt, cAHS_ID, "INF", "O", nSERVICE_FEE, "nINFPriceO")
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            Loop Until Not oReader.Read
        End If
        oReader.Close()

        '   get internet claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
                "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
                "AND (CALL_CALLER.CALLER_TYPE = 'NET' OR CALL_CALLER.CALLER_TYPE = 'INT') "
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
                Select Case CType(oReader.GetValue(oReader.GetOrdinal("LOB_CD")), String)
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUInternet")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "N", nSERVICE_FEE, "nPAUPriceI")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPRInternet")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "N", nSERVICE_FEE, "nPPRPriceI")
                    Case "INF"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFInternet")
                        getFees(oConn, oRpt, cAHS_ID, "INF", "N", nSERVICE_FEE, "nINFPriceI")
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select
                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            Loop Until Not oReader.Read
        End If
        oReader.Close()

        '   get temped claims
        cSQL = "Select count(distinct(CALL.CALL_ID)) as nTotal From CALL, CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
                "AND TEMPEDPOLICY_FLG = 'Y' "
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
        'cSQL = "Select count(distinct(CALL.CALL_ID)) as nTotal From CALL, ESCALATION_OUTCOME " & _
        '        "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '        "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '        "AND CALL.STATUS = 'COMPLETED' " & _
        '        "AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID "
        'If lIsCCE Then
        '    cSQL = cSQL & "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        'Else
        '    cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        'End If

        cSQL = "SELECT COUNT(distinct(CALL1.CALL_ID)) AS nTotal, SUM(FEE1.FEE_AMOUNT) AS nTotalFee " & _
                "FROM CALL CALL1, ESCALATION_OUTCOME, USERS USERS1, CALL_CALLER CALL_CALLER1, FEE FEE1 " & _
                "WHERE CALL1.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL1.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND CALL1.STATUS = 'COMPLETED' " & _
                "AND CALL1.CALL_ID = ESCALATION_OUTCOME.CALL_ID " & _
                "AND CALL1.USER_ID = USERS1.USER_ID " & _
                "AND CALL1.CALL_ID = CALL_CALLER1.CALL_ID (+) " & _
                "AND FEE1.LOB_CD = CALL1.LOB_CD " & _
                "AND FEE1.FEE_TYPE_ID = 4 " & _
                "AND FEE1.CALL_TYPE = DECODE(CALL1.lob_cd, " & _
                                            "'INF', DECODE(USERS1.SITE_ID, " & _
                                                            "1, 'O', " & _
                                                            "2, 'O', " & _
                                                            "'I' " & _
                                                            "), " & _
                                                    "DECODE(SUBSTR(CALL_CALLER1.Caller_Type, 1, 1), " & _
                                                            "'F', 'F', " & _
                                                            "DECODE(CALL_CALLER1.Caller_Type, " & _
                                                                    "'EML', 'E', " & _
                                                                    "'EMAIL', 'E', " & _
                                                                    "'N', 'N', " & _
                                                                    "'NET', 'N', " & _
                                                                    "'INT', 'N', " & _
                                                                    "DECODE(USERS1.SITE_ID, " & _
                                                                            "1, 'O', " & _
                                                                            "2, 'O', " & _
                                                                            "'C' " & _
                                                                    ") " & _
                                                            ") " & _
                                                    ") " & _
                                            ") " & _
                "AND CALL1.CLIENT_HRCY_STEP_ID = " & cAHS_ID & " " & _
                "AND FEE1.ACCNT_HRCY_STEP_ID = CALL1.ACCNT_HRCY_STEP_ID "

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nESCALATE_FEE, "O", "nEscalationPrice")

            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nEscalationsTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))

            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nEscalationDollarsGrandTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotalFee")))
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
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nFaxedPagesTotal")

        cSQL = " Select COUNT(transmission_outcome.transmission_outcome_id) AS Printcount " & _
               "FROM transmission_outcome,transmission_outcome_step, CALL " & _
               "WHERE transmission_outcome.CALL_ID = CALL.CALL_ID " & _
               "AND transmission_outcome.transmission_outcome_id = transmission_outcome_step.transmission_outcome_id " & _
               "AND (transmission_outcome_step.status = 'PROCESSED' " & _
               "AND transmission_outcome_step.status <> 'FAILED') " & _
               "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & " AND CALL.ACCNT_HRCY_STEP_ID IN " & _
               "(SELECT accnt_hrcy_step_id " & _
                "FROM account_hierarchy_step START WITH accnt_hrcy_step_id =  " & cAHS_ID & _
                " CONNECT BY parent_node_id = PRIOR accnt_hrcy_step_id) " & _
                "AND CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND transmission_outcome.RESUBMITTED_FLG <> 'Y' " & _
                "AND transmission_outcome_step.RESUBMITTED_FLG <>'Y' " & _
                "AND transmission_outcome_step.transmission_type_id = 2"

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nPRINT_FEE, "nPrintedPagesPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPrintedPagesTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("Printcount")))
        End If
        oReader.Close()

        nINFFreePercentage = getFreeINFOPercent(oConn, cAHS_ID)
        nINFFreeCalls = CInt(nTotalClaimsReceived * (nINFFreePercentage / 100))
        If nTotalINFClaims > nINFFreeCalls Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFCalls2Bill")
            oParamFld.Text = CStr(nTotalINFClaims - nINFFreeCalls)
        End If

        '   Vendor referral
        cSQL = "Select count(distinct(C.CALL_ID)) as nTotal From CALL C, CALL_CLAIM CC, CALL_VENDOR_REFERRAL CVR " & _
            "Where C.CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
            "AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
            "AND C.CALL_ID = CC.CALL_ID " & _
            "AND CC.CALL_CLAIM_ID = CVR.CALL_CLAIM_ID " & _
            "AND CVR.REFERRAL_ACCEPTED = 'Y' "
        If lIsCCE Then
            cSQL = cSQL & "AND C.ACCNT_HRCY_STEP_ID = " & cAHS_ID
        Else
            cSQL = cSQL & "AND C.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        End If
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nVENDOR_REFERRAL_FEE, "nVRPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVRTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()

        oConn.Close()
    End Sub

    Sub getFees(ByVal oConn As OracleConnection, _
                ByVal oRpt As AMEBillingSummary, _
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
                            ByVal oRpt As AMEBillingSummary, _
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

    Sub getProcessingFees(ByVal oConn As OracleConnection, _
                            ByVal oRpt As AMEBillingSummary, _
                            ByVal cAHS_ID As String, _
                            ByVal nFeeTypeId As Integer, _
                            ByVal cCallType As String, _
                            ByVal cFormulaName As String)
        Dim cSQL As String
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                "AND FEE_TYPE_ID = " & nFeeTypeId & " " & _
                "AND CALL_TYPE = '" & cCallType & "'"
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
                        ByVal oRpt As AMEBillingSummary, _
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
