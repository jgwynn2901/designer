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

Module CnlServiceFeeReport
    Const nLOB_CD As String = "CALL.LOB_CD in('ALI','APD','CRG','INF')"
    Const nLOB_CDNoInf As String = "CALL.LOB_CD in('ALI','APD','CRG')"
    Sub getCnlBillingSummaryReport(ByRef oRpt As CnlBillingSummary, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String)
        Dim cSQL As String
        Dim oConn As New OracleConnection
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim nAPDEscTotal As Int32
        Dim nCRGEscTotal As Int32
        Dim nALIEscTotal As Int32


        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim callClaimID As Int32

        Dim oParamFld1 As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oParamFld2 As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oParamFld3 As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

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
                "AND LOB_CD <> 'INF' "
        cSQL = cSQL & "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID

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
        cSQL = cSQL & "AND CLIENT_HRCY_STEP_ID = " & cAHS_ID
        cSQL = cSQL & " Group by LOB_CD"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalINFClaims = 0
        oReader.Read()
        If oReader.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFCalls")
            getFees(oConn, oRpt, cAHS_ID, "INF", "I", nSERVICE_FEE, "nINFCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))

            nTotalINFClaims = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
            oParamFld.Text = CStr(nTotalINFClaims)
        End If
        oReader.Close()
        '----------------------------------------------
        '
        '   get call claims
        cSQL = "Select LOB_CD, count(*) as nTotal,Call_Claim.CALL_CLAIM_ID From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID(+) " & _
             "AND (CALL_CALLER.CALLER_TYPE IS NULL OR (SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) <> 'F')) " & _
             "AND (CALL_CLAIM.CLAIM_TYPE IS NULL OR (SUBSTR(CALL_CLAIM.CLAIM_TYPE, 1, 1) <> 'N'))"
        cSQL = cSQL & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        'cSQL = cSQL & " AND " & nLOB_CD
        cSQL = cSQL & " Group by LOB_CD,CALL_CLAIM_ID"

        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalClaimsCalls = 0
        oReader.Read()


        If oReader.HasRows Then
            Do
                lIsINFO = False

                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "APC"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nAPCCalls")
                        oParamFld1 = oRpt.DataDefinition.FormulaFields.Item("nAPDCalls")
                        oParamFld2 = oRpt.DataDefinition.FormulaFields.Item("nCRGCalls")
                        oParamFld3 = oRpt.DataDefinition.FormulaFields.Item("nALICalls")

                        getFees(oConn, oRpt, cAHS_ID, "APC", "C", nSERVICE_FEE, "nAPCCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        callClaimID = CInt(oReader.GetValue(oReader.GetOrdinal("CALL_CLAIM_ID")))
                        oParamFld.Text = CStr(CInt(oParamFld.Text) + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        oParamFld1.Text = CStr(CInt(oParamFld1.Text) + getAPDCalls(oConn, callClaimID))
                        oParamFld2.Text = CStr(CInt(oParamFld2.Text) + getCRGCalls(oConn, callClaimID))
                        oParamFld3.Text = CStr(CInt(oParamFld3.Text) + getALICalls(oConn, callClaimID))

                        'nTotalClaimsCalls = nTotalClaimsCalls + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))

                        nTotalClaimsCalls = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        'Case "APD"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nAPDCalls")
                        '    getFees(oConn, oRpt, cAHS_ID, "APD", "C", nSERVICE_FEE, "nAPDCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))

                        'Case "CRG"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRGCalls")
                        '    getFees(oConn, oRpt, cAHS_ID, "CRG", "C", nSERVICE_FEE, "nCRGCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        'Case "ALI"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nALICalls")
                        '    getFees(oConn, oRpt, cAHS_ID, "ALI", "C", nSERVICE_FEE, "nALICharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(CInt(oParamFld.Text) + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                    Case "INF"
                        lIsINFO = True
                        'Case Else
                        'oParamFld = Nothing '   force an error if not found
                End Select

                'If Not lIsINFO Then

                '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                '    nTotalClaimsCalls = nTotalClaimsCalls + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                '    nTotalClaimsCalls = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                'End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        'get faxed claims
        cSQL = "Select LOB_CD,CALL_CLAIM_ID,count(*) as nTotal From CALL, CALL_CALLER,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND CALL_CLAIM.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'F' "
        cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        ' cSQL = cSQL & " AND " & nLOB_CD
        cSQL = cSQL & " Group by LOB_CD,CALL_CLAIM_ID"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalClaimsFaxes = 0
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "APC"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nAPCFaxes")
                        oParamFld1 = oRpt.DataDefinition.FormulaFields.Item("nAPDFaxes")
                        oParamFld2 = oRpt.DataDefinition.FormulaFields.Item("nCRGFaxes")
                        oParamFld3 = oRpt.DataDefinition.FormulaFields.Item("nALIFaxes")

                        getFees(oConn, oRpt, cAHS_ID, "APC", "C", nFAX_FEE, "nAPCCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        callClaimID = CInt(oReader.GetValue(oReader.GetOrdinal("CALL_CLAIM_ID")))
                        oParamFld.Text = CStr(CInt(oParamFld.Text) + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        oParamFld1.Text = CStr(CInt(oParamFld1.Text) + getAPDFaxes(oConn, callClaimID))
                        oParamFld2.Text = CStr(CInt(oParamFld2.Text) + getCRGFaxes(oConn, callClaimID))
                        oParamFld3.Text = CStr(CInt(oParamFld3.Text) + getALIFaxes(oConn, callClaimID))
                        ' oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        nTotalClaimsFaxes = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        'Case "APD"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nAPDFaxes")
                        '    getFees(oConn, oRpt, cAHS_ID, "APD", "C", nSERVICE_FEE, "nAPDCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        'Case "CRG"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRGFaxes")
                        '    getFees(oConn, oRpt, cAHS_ID, "CRG", "C", nSERVICE_FEE, "nCRGCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                    Case "ALI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nALIFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "ALI", "C", nSERVICE_FEE, "nALICharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                    Case "INF"
                        lIsINFO = True
                        'Case Else
                        '    oParamFld = Nothing '   force an error if not found
                End Select
                'If Not lIsINFO Then
                '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                '    nTotalClaimsFaxes = nTotalClaimsFaxes + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                '    nTotalClaimsFaxes = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                'End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        '   get internet claims
        cSQL = "Select LOB_CD, count(*) as nTotal,Call_Claim.CALL_CLAIM_ID From CALL, CALL_CALLER ,CALL_CLAIM " & _
             " Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             " AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             " AND STATUS = 'COMPLETED' " & _
             " AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             " AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
             " AND SUBSTR(CALL_CLAIM.CLAIM_TYPE, 1, 1) = 'N' "
        cSQL = cSQL & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        '  cSQL = cSQL & " AND " & nLOB_CD
        cSQL = cSQL & " Group by LOB_CD,CALL_CLAIM_ID"
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        nTotalClaimsInt = 0
        oReader.Read()
        If oReader.HasRows Then
            Do
                lIsINFO = False
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "APC"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nAPCInternet")
                        oParamFld1 = oRpt.DataDefinition.FormulaFields.Item("nAPDInternet")
                        oParamFld2 = oRpt.DataDefinition.FormulaFields.Item("nCRGInternet")
                        oParamFld3 = oRpt.DataDefinition.FormulaFields.Item("nALIInternet")

                        getFees(oConn, oRpt, cAHS_ID, "APC", "C", nSERVICE_FEE, "nAPCCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        callClaimID = CInt(oReader.GetValue(oReader.GetOrdinal("CALL_CLAIM_ID")))
                        oParamFld.Text = CStr(CInt(oParamFld.Text) + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        oParamFld1.Text = CStr(CInt(oParamFld1.Text) + getAPDCalls(oConn, callClaimID))
                        oParamFld2.Text = CStr(CInt(oParamFld2.Text) + getCRGCalls(oConn, callClaimID))
                        oParamFld3.Text = CStr(CInt(oParamFld3.Text) + getALICalls(oConn, callClaimID))

                        '    'nTotalClaimsInt = nTotalClaimsInt + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        nTotalClaimsInt = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        'Case "APD"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nAPDInternet")
                        '    getFees(oConn, oRpt, cAHS_ID, "APD", "C", nSERVICE_FEE, "nAPDCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        'Case "CRG"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRGInternet")
                        '    getFees(oConn, oRpt, cAHS_ID, "CRG", "C", nSERVICE_FEE, "nCRGCharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        'Case "ALI"
                        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nALIInternet")
                        '    getFees(oConn, oRpt, cAHS_ID, "ALI", "C", nSERVICE_FEE, "nALICharges", CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                        '    oParamFld.Text = CStr(CInt(oParamFld.Text) + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal"))))
                    Case "INF"
                        lIsINFO = True
                        'Case Else
                        '    oParamFld = Nothing '   force an error if not found
                End Select
                'If Not lIsINFO Then
                '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                '    'nTotalClaimsInt = nTotalClaimsInt + CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                '    nTotalClaimsInt = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                'End If
            Loop Until Not oReader.Read
        End If
        oReader.Close()

        ''   get temped claims
        cSQL = "Select count(*) as nTotal From CALL, CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
                "AND TEMPEDPOLICY_FLG = 'Y' " & _
                "AND " & nLOB_CD
        cSQL = cSQL & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nTEMP_FEE, "nTempPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nTempTotal")
            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()
        ''   get escalations
        cSQL = "Select COUNT(*) as  nAPDTotal from  ESCALATION_OUTCOME,CALL_APD,CALL_CLAIM,CALL " & _
                "where CALL_CLAIM.CALL_ID = ESCALATION_OUTCOME.CALL_ID" & _
                " AND CALL_CLAIM.CALL_CLAIM_ID=CALL_APD.CALL_CLAIM_ID" & _
                " AND CALL.CALL_ID =CALL_CLAIM.CALL_ID" & _
                " AND CALL.CALL_ID=ESCALATION_OUTCOME.CALL_ID" & _
                " AND  CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                " AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                " AND CALL.STATUS = 'COMPLETED' " & _
                " AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID" & _
                " AND Coverage_FLG = 'Y'"
        cSQL = cSQL & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()

        nAPDEscTotal = Convert.ToInt32(oReader.GetValue(oReader.GetOrdinal("nAPDTotal")))
        oReader.Close()

        cSQL = "Select COUNT(*) as  nALITotal from  ESCALATION_OUTCOME,CALL_ALI,CALL_CLAIM,CALL " & _
                "where CALL_CLAIM.CALL_ID = ESCALATION_OUTCOME.CALL_ID" & _
                " AND CALL_CLAIM.CALL_CLAIM_ID=CALL_ALI.CALL_CLAIM_ID" & _
                " AND CALL.CALL_ID =CALL_CLAIM.CALL_ID" & _
                " AND CALL.CALL_ID=ESCALATION_OUTCOME.CALL_ID" & _
                " AND  CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                " AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                " AND CALL.STATUS = 'COMPLETED' " & _
                " AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID" & _
                " AND Coverage_FLG = 'Y'"
        cSQL = cSQL & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()

        nALIEscTotal = Convert.ToInt32(oReader.GetValue(oReader.GetOrdinal("nALITotal")))

        oReader.Close()


        cSQL = "Select COUNT(*) as  nCRGTotal from  ESCALATION_OUTCOME,CALL_CRG,CALL_CLAIM,CALL " & _
                "where CALL_CLAIM.CALL_ID = ESCALATION_OUTCOME.CALL_ID" & _
                " AND CALL_CLAIM.CALL_CLAIM_ID=CALL_CRG.CALL_CLAIM_ID" & _
                " AND CALL.CALL_ID =CALL_CLAIM.CALL_ID" & _
                " AND CALL.CALL_ID=ESCALATION_OUTCOME.CALL_ID" & _
                " AND  CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                " AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                " AND CALL.STATUS = 'COMPLETED' " & _
                " AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID" & _
                " AND Coverage_FLG = 'Y'"
        cSQL = cSQL & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()

        nCRGEscTotal = Convert.ToInt32(oReader.GetValue(oReader.GetOrdinal("nCRGTotal")))

        getProcessingFees(oConn, oRpt, cAHS_ID, nESCALATE_FEE, "nEscalationPrice")
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nEscalationsTotal")
        oParamFld.Text = CStr(nAPDEscTotal + nALIEscTotal + nCRGEscTotal) 'CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        'End If
        oReader.Close()

        'Services Cost
        'cSQL = "Select count(*) as nTotal From CALL " & _
        '        "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
        '        "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
        '        "AND CALL.STATUS = 'COMPLETED' " & _
        '        "AND " & nLOB_CD & " "
        'cSQL = cSQL & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID

        'oCmd.CommandText = cSQL
        'oReader = oCmd.ExecuteReader()
        'oReader.Read()
        'If oReader.HasRows Then
        '    getProcessingFees(oConn, oRpt, cAHS_ID, nSERVICE_FEE, "nServicePrice")
        '    oParamFld = oRpt.DataDefinition.FormulaFields.Item("nServiceTotal")
        '    oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        'End If
        'oReader.Close()
        ' Service Cost

        ''   get total transmissions
        cSQL = "Select COUNT(TOC.TRANSMISSION_OUTCOME_ID) AS transmissionCount " & _
                "From TRANSMISSION_OUTCOME TOC, TRANSMISSION_OUTCOME_STEP TOS " & _
                "Where TOS.STATUS = 'PROCESSED' " & _
                "AND TOC.TRANSMISSION_OUTCOME_ID = TOS.TRANSMISSION_OUTCOME_ID " & _
                "AND TOS.TRANSMISSION_TYPE_ID = 1 " & _
                "AND TOC.CALL_ID IN (Select DISTINCT CALL_ID From CALL " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
                "AND STATUS = 'COMPLETED' "
        cSQL = cSQL & "AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS_ID & ")"
        oCmd.CommandText = cSQL
        nTotalTransmissions = 0
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            nTotalTransmissions = CInt(oReader.GetValue(oReader.GetOrdinal("transmissionCount")))
        End If
        oReader.Close()

        ''   get faxed pages
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
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            nTotalFaxedPages = CInt(IIf(IsDBNull(oReader.GetValue(oReader.GetOrdinal("pageCount"))), 0, oReader.GetValue(oReader.GetOrdinal("pageCount"))))
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

        getCNLTieredFees(oConn, oRpt, cAHS_ID, "C", nSERVICE_FEE, nTotalClaimsCalls)
        getCNLTieredFees(oConn, oRpt, cAHS_ID, "F", nSERVICE_FEE, nTotalClaimsFaxes)
        getCNLTieredFees(oConn, oRpt, cAHS_ID, "N", nSERVICE_FEE, nTotalClaimsInt)
        oConn.Close()
    End Sub
    Private Sub getCNLTieredFees(ByVal oConn As OracleConnection, _
               ByVal oRpt As CnlBillingSummary, _
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
                " AND REASON_CODE = 'N' and LOB_CD='APC'" & _
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
                cPrice = Replace(cPrice, "~", CStr(1))
                cTotalPrice = Replace(cTotalPrice, "~", CStr(1))
                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cPrice)
                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("FEE_AMOUNT")))

                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cTotalPrice)
                nFeeAmntColNo = oReader.GetOrdinal("FEE_AMOUNT")
                If nTotalClaims > nClaimsTo Then
                    oParamFld.Text = CStr(CDbl(oReader.GetValue(nFeeAmntColNo)) * nClaimsTo)
                Else
                    oParamFld.Text = CStr(CDbl(oReader.GetValue(nFeeAmntColNo)) * (nTotalClaims - nClaimsFrom + CInt(IIf(nClaimsFrom = 1, 1, 0))))
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

    Sub getFees(ByVal oConn As OracleConnection, _
                ByVal oRpt As CnlBillingSummary, _
                ByVal cAHS_ID As String, _
                ByVal cLOB As String, _
                ByVal cCallType As String, _
                ByVal nFeeTypeId As Integer, _
                ByVal cFormulaName As String, _
                ByVal nTotal As Integer)
        Dim cSQL As String
        Dim nTotalPrice As Int32
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select FEE_AMOUNT as nAmount From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
                " AND CALL_TYPE = '" & cCallType & "'" & _
                " AND LOB_CD = '" & cLOB & "' " & _
                " AND FEE_TYPE_ID = " & nFeeTypeId
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then

            'nTotalPrice = CInt(oReader.GetValue(oReader.GetOrdinal("nAmount"))) * nTotal
            nTotalPrice = CInt(oReader.GetValue(oReader.GetOrdinal("nAmount")))

            oParamFld = oRpt.DataDefinition.FormulaFields.Item(cFormulaName)
            oParamFld.Text = CStr(nTotalPrice)
        End If
        oReader.Close()

    End Sub
    Sub getProcessingFees(ByVal oConn As OracleConnection, _
                            ByVal oRpt As CnlBillingSummary, _
                            ByVal cAHS_ID As String, _
                            ByVal nFeeTypeId As Integer, _
                            ByVal cFormulaName As String)
        Dim cSQL As String
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID & _
               " AND FEE_TYPE_ID = " & nFeeTypeId
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
                        ByVal oRpt As CnlBillingSummary, _
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
    'APD calls
    Function getAPDFaxes(ByVal oConn As OracleConnection, ByVal cClaimNo As Int32) As Int32
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        getAPDFaxes = 0
        cSQL = "Select Count(*) as TotalAPDFaxes from Call_APD where Call_claim_ID = '" & cClaimNo & "' and Coverage_Flg='Y'"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getAPDFaxes = CInt(oReader.GetValue(oReader.GetOrdinal("TotalAPDFaxes")))
        End If
        oReader.Close()
    End Function
    'APD Faxes
    Function getCRGFaxes(ByVal oConn As OracleConnection, ByVal cClaimNo As Int32) As Int32
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        getCRGFaxes = 0
        cSQL = "Select Count(*) as TotalCRGFaxes from Call_CRG where Call_claim_ID = '" & cClaimNo & "' and Coverage_Flg='Y'"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getCRGFaxes = CInt(oReader.GetValue(oReader.GetOrdinal("TotalCRGFaxes")))
        End If
        oReader.Close()
    End Function
    Function getALIFaxes(ByVal oConn As OracleConnection, ByVal cClaimNo As Int32) As Int32
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        getALIFaxes = 0
        cSQL = "Select Count(*) as TotalALICalls from Call_ALI where Call_claim_ID = '" & cClaimNo & "' and Coverage_Flg='Y'"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getALIFaxes = CInt(oReader.GetValue(oReader.GetOrdinal("TotalALICalls")))
        End If
        oReader.Close()
    End Function
    'APD calls
    Function getAPDCalls(ByVal oConn As OracleConnection, ByVal cClaimNo As Int32) As Int32
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        getAPDCalls = 0
        cSQL = "Select Count(*) as TotalAPDCalls from Call_APD where Call_claim_ID = '" & cClaimNo & "' and Coverage_Flg='Y'"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getAPDCalls = CInt(oReader.GetValue(oReader.GetOrdinal("TotalAPDCalls")))
        End If
        oReader.Close()
    End Function
    'CRG Calls
    Function getCRGCalls(ByVal oConn As OracleConnection, ByVal cClaimNo As Int32) As Int32
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        getCRGCalls = 0
        cSQL = "Select Count(*) as TotalCRGCalls from Call_CRG where Call_claim_ID = '" & cClaimNo & "' and Coverage_Flg='Y'"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getCRGCalls = CInt(oReader.GetValue(oReader.GetOrdinal("TotalCRGCalls")))
        End If
        oReader.Close()
    End Function
    'ALI Calls
    Function getALICalls(ByVal oConn As OracleConnection, ByVal cClaimNo As Int32) As Int32
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition

        getALICalls = 0
        cSQL = "Select Count(*) as TotalALICalls from Call_ALI where Call_claim_ID = '" & cClaimNo & "' and Coverage_Flg='Y'"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            getALICalls = CInt(oReader.GetValue(oReader.GetOrdinal("TotalALICalls")))
        End If
        oReader.Close()
    End Function
End Module
