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
Imports System.Text.StringBuilder
Module GBSsummary
    Sub getClaimsGBS(ByRef oRpt As GBSBillingSummaryFixed, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String)
        Dim cSQL As String
        Dim oConn As New OracleConnection
        Dim oCmd As New OracleCommand
        Dim oReaderWMI As OracleDataReader
        Dim oReaderGBS As OracleDataReader
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
        Dim strLOB_WMI As String
        Dim strLOB_GBS As String
        Dim saveTotalWMI As String
        Dim arrTotalWMI As String
        Dim sTotalWMI As String
        Dim strCompLOBTotal As String
        Dim nTotalWMI As Integer
        Dim nrTotalGBS As Integer
        Dim strReturnPrmt As String
        Dim SqlStrBuilderWMI As System.Text.StringBuilder
        Dim SqlStrBuilderGBS As System.Text.StringBuilder

    
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
        '--wmi client
        Dim strClient_ID As Int32
        Dim cAHS_ID_WMI As Int32
        strClient_ID = 46
        cAHS_ID_WMI = 12159325

        oConn.ConnectionString = CStr(HttpContext.Current.Session("ConnectionString")).Replace("DSN=", "Data Source=")
        oConn.Open()
        oCmd.CommandText = "ALTER SESSION SET NLS_DATE_FORMAT = 'DD-MON-YYYY HH:MI:SS'"
        oCmd.Connection = oConn
        oCmd.ExecuteNonQuery()

        '----WMI WMI
        '  oCmd.CommandText = SqlStrBuilderWMI.ToString
        ' oReaderWMI = oCmd.ExecuteReader()
        ' oReaderWMI.Read()

        '  Dim nTotalClaimsReceivedWMI As Integer
        ' If oReaderWMI.HasRows Then
        'nTo'talClaimsReceivedWMI = CInt(oReaderWMI.GetValue(oReaderWMI.GetOrdinal("totalCalls")))
        ' End If
        '  oReaderWMI.Close()
        '-------GBS BGS BGS  

        '********ALL GBS --ALL GBS --ALL GBS**get total number of claims (excluding INF calls)
        '   and exclud the WMI 003000 and Sunston Hotel Prorerty 001405** 
        'ALL GBS --ALL GBS --ALL GBS
        '**********************************************************************
         SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select COUNT(call.call_id) AS totalCalls ")
        SqlStrBuilderGBS.Append(" From CALL, call_account, call_claim")
        SqlStrBuilderGBS.Append(" Where call_account.CALL_CLAIM_ID = call_claim.CALL_CLAIM_ID")
        SqlStrBuilderGBS.Append("  and call.CALL_ID = call_claim.CALL_ID ")
        SqlStrBuilderGBS.Append(" and CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND LOB_CD <> 'INF' ")
        SqlStrBuilderGBS.Append(" and (call_account.LOCATION_CODE <> '003000' and call_account.LOCATION_CODE <> '001406')")
        SqlStrBuilderGBS.Append(" AND CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" AND call_claim.INFO_CALL_FLG <> 'Y'")
        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()

        Dim nTotalClaimsReceivedGBS As Integer
        If oReaderGBS.HasRows Then
            nTotalClaimsReceived = CInt(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("totalCalls")))
        End If

        oReaderGBS.Close()
        'TOTAL WITHOUT WMI
        ' If nTotalClaimsReceivedGBS >= nTotalClaimsReceivedWMI Then
        '     nTotalClaimsReceived = nTotalClaimsReceivedGBS - nTotalClaimsReceivedWMI
        '    Else
        '       nTotalClaimsReceived = nTotalClaimsReceivedGBS
        '   End If

        '****** ***************************************II ******************************************
        '*******************************************get INF claims ***********************************8

        '-GBS--GBS--GBS-GBS--GBS--GBS************* get INF claims ***-GBS-GBS-GBS--GBS--GBS--GBS-***************************
        ''   get INF claims
        'SqlStrBuilderGBS = New System.Text.StringBuilder("")
        'SqlStrBuilderGBS.Append("Select count(*) as nTotal from CALL C,CALL_ACCOUNT CA, CALL_CLAIM CCL ")
        'SqlStrBuilderGBS.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        'SqlStrBuilderGBS.Append(" AND C.CALL_ID(+) = CCL.CALL_ID ")
        'SqlStrBuilderGBS.Append(" AND CA.CALL_CLAIM_ID = CCL.CALL_CLAIM_ID ")
        'SqlStrBuilderGBS.Append(" AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        'SqlStrBuilderGBS.Append(" AND C.STATUS = 'COMPLETED' ")
        'SqlStrBuilderGBS.Append(" AND C.LOB_CD = 'INF' ")
        'SqlStrBuilderGBS.Append(" AND (CA.LOCATION_CODE <> '003000' and CA.LOCATION_CODE <> '001406')")
        'SqlStrBuilderGBS.Append(" AND C.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        'SqlStrBuilderGBS.Append(" AND CCL.INFO_CALL_FLG <> 'Y'")
        'SqlStrBuilderGBS.Append(" Group by LOB_CD")


        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append("Select count(*) as nTotal from CALL C,X_CALL_CLAIM XCC, CALL_CLAIM CCL ")
        SqlStrBuilderGBS.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND C.CALL_ID(+) = CCL.CALL_ID ")
        SqlStrBuilderGBS.Append(" AND xcc.CALL_CLAIM_ID = CCL.CALL_CLAIM_ID ")
        SqlStrBuilderGBS.Append(" AND C.CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND C.STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND C.LOB_CD = 'INF' ")
        SqlStrBuilderGBS.Append(" AND CCL.INFO_CALL_FLG <> 'Y'")
        SqlStrBuilderGBS.Append(" AND C.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" AND  xcc.FIELD='SUN_INFO_FLG' and xcc.VALUE ='N'")
        SqlStrBuilderGBS.Append(" Group by LOB_CD")


        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()

        nTotalINFClaims = 0

        If oReaderGBS.HasRows Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFTotal")
            getFees(oConn, oRpt, cAHS_ID, "INF", "I", nSERVICE_FEE, "nINFPriceC")
            nTotalINFClaims = CInt(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))
            oParamFld.Text = CStr(nTotalINFClaims)
        End If
        oReaderGBS.Read()


        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        '***** III ****************************III*********************************
        '********************************** get call claims *******************************


        'GBS GBS GBS ******and exclud the WMI 003000 and Sunston Hotel Prorerty 001405**
        '**********  get call claims ********-gbs*********************************************
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER,CALL_CLAIM, call_account ")
        SqlStrBuilderGBS.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "')")
        SqlStrBuilderGBS.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND CALL_CALLER.CALL_ID(+) = CALL.CALL_ID ")
        SqlStrBuilderGBS.Append(" AND call.CALL_ID = call_claim.CALL_ID  ")
        SqlStrBuilderGBS.Append(" and call_account.CALL_CLAIM_ID(+) = CALL_CLAIM.CALL_CLAIM_ID ")
        SqlStrBuilderGBS.Append(" AND (CALL_CALLER.CALLER_TYPE IS NULL OR (CALL_CALLER.CALLER_TYPE IS NOT NULL AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) <> 'F' AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) <> 'E' AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) <> 'N'))")
        SqlStrBuilderGBS.Append(" and (call_account.LOCATION_CODE <> '003000' and call_account.LOCATION_CODE <> '001406')")
        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" Group by LOB_CD")

        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()


        If oReaderGBS.HasRows Then
            Do
                lIsINFO = False

                strLOB_GBS = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("LOB_CD")))
                Select Case strLOB_GBS
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

                If Not lIsINFO Then ' If Not lIsINFO Then

                    'strReturnPrmt = getMatchLOB(strLOB_GBS, sTotalWMI, nrTotalGBS)
                    oParamFld.Text = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))
                End If  'If Not lIsINFO Then

            Loop Until Not oReaderGBS.Read
        End If
        oReaderGBS.Close()


        'Get Email Claims 

        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER, CALL_CLAIM, call_account ")
        SqlStrBuilderGBS.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND CALL_CALLER.CALL_ID(+) = CALL.CALL_ID")
        SqlStrBuilderGBS.Append(" AND call.CALL_ID = call_claim.CALL_ID  ")
        SqlStrBuilderGBS.Append(" and call_account.CALL_CLAIM_ID(+) = CALL_CLAIM.CALL_CLAIM_ID ")
        SqlStrBuilderGBS.Append(" AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'E' ")
        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" and (call_account.LOCATION_CODE <> '003000' and call_account.LOCATION_CODE <> '001406')")
        SqlStrBuilderGBS.Append(" Group by LOB_CD")
        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()

        If oReaderGBS.HasRows Then
            Do
                lIsINFO = False
                strLOB_GBS = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("LOB_CD")))
                Select Case strLOB_GBS
                    Case "PAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPAUEmails")
                        getFees(oConn, oRpt, cAHS_ID, "PAU", "E", nSERVICE_FEE, "nPAUPriceE")
                    Case "PLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLIEmails")
                        getFees(oConn, oRpt, cAHS_ID, "PLI", "E", nSERVICE_FEE, "nPLIPriceE")
                    Case "PPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPPREmailes")
                        getFees(oConn, oRpt, cAHS_ID, "PPR", "E", nSERVICE_FEE, "nPPRPriceE")
                    Case "CAU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCAUEmails")
                        getFees(oConn, oRpt, cAHS_ID, "CAU", "E", nSERVICE_FEE, "nCAUPriceE")
                    Case "CLI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCLIEmails")
                        getFees(oConn, oRpt, cAHS_ID, "CLI", "E", nSERVICE_FEE, "nCLIPriceE")
                    Case "CPR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCPREmails")
                        getFees(oConn, oRpt, cAHS_ID, "CPR", "E", nSERVICE_FEE, "nCPRPriceE")
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWOREmails")
                        getFees(oConn, oRpt, cAHS_ID, "WOR", "E", nSERVICE_FEE, "nWORPriceE")
                    Case "CRI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nCRIEmails")
                        getFees(oConn, oRpt, cAHS_ID, "CRI", "E", nSERVICE_FEE, "nCRIPriceE")
                    Case "INF"
                        lIsINFO = True
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select

                If Not lIsINFO Then ' If Not lIsINFO Then
                    oParamFld.Text = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))
                End If  'If Not lIsINFO Then

            Loop Until Not oReaderGBS.Read

        End If
        oReaderGBS.Close()

        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        '*******************************************IV ****************************************************** 
        '**************************************get faxed claims**********************************************



        ' GBS GBS GBS  **************** get faxed claims ********* GBS GBS GBS
        '**and exclud the WMI 003000 and Sunston Hotel Prorerty 001405**
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER, CALL_CLAIM, call_account ")
        SqlStrBuilderGBS.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND CALL_CALLER.CALL_ID(+) = CALL.CALL_ID")
        SqlStrBuilderGBS.Append(" AND call.CALL_ID = call_claim.CALL_ID  ")
        SqlStrBuilderGBS.Append(" and call_account.CALL_CLAIM_ID(+) = CALL_CLAIM.CALL_CLAIM_ID ")
        SqlStrBuilderGBS.Append(" AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'F' ")
        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" and (call_account.LOCATION_CODE <> '003000' and call_account.LOCATION_CODE <> '001406')")
        SqlStrBuilderGBS.Append(" Group by LOB_CD")
        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()

        If oReaderGBS.HasRows Then
            Do
                lIsINFO = False
                strLOB_GBS = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("LOB_CD")))
                Select Case strLOB_GBS
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

                If Not lIsINFO Then ' If Not lIsINFO Then
                    oParamFld.Text = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))
                End If  'If Not lIsINFO Then

            Loop Until Not oReaderGBS.Read

        End If
        oReaderGBS.Close()



        '***-GBS  GBS GBS*********** get internet claims****-GBS  GBS GBS-*******
        '**and exclud the WMI 003000 and Sunston Hotel Prorerty 001405**
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM, call_account ")
        SqlStrBuilderGBS.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_CALLER.CALL_ID (+)= CALL.CALL_ID")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_ID = CALL_CLAIM.CALL_ID ")
        SqlStrBuilderGBS.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" and call_account.CALL_CLAIM_ID(+) = CALL_CLAIM.CALL_CLAIM_ID")
        'SqlStrBuilderGBS.Append(" AND SUBSTR(CALL_CLAIM.CLAIM_TYPE, 1, 1) = 'N' ")
        SqlStrBuilderGBS.Append(" AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) = 'N' ")
        SqlStrBuilderGBS.Append(" and (call_account.LOCATION_CODE <> '003000' and call_account.LOCATION_CODE <> '001406')")
        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" Group by LOB_CD")
        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()

        If oReaderGBS.HasRows Then
            Do
                lIsINFO = False
                strLOB_GBS = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("LOB_CD")))
                Select Case strLOB_GBS
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

                If Not lIsINFO Then ' If Not lIsINFO Then
                    oParamFld.Text = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))
                End If  'If Not lIsINFO Then
            Loop Until Not oReaderGBS.Read

        End If
        oReaderGBS.Close()
        'oReaderWMI.Close()

        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

        '******************************************  VI *************************************************** 
        '***************************************get temped claims ******************************************

        '***-GBS  GBS GBS*********** get temped claims****-GBS  GBS GBS-*******
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select count(*) as nTotal From CALL, CALL_CLAIM,CALL_ACCOUNT ")
        SqlStrBuilderGBS.Append(" Where CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_ID = CALL_CLAIM.CALL_ID ")
        SqlStrBuilderGBS.Append(" AND CALL_CLAIM.CALL_CLAIM_ID = CALL_ACCOUNT.call_claim_id ")
        'SqlStrBuilderGBS.Append(" AND TEMPEDPOLICY_FLG = 'Y' ")
        SqlStrBuilderGBS.Append(" AND CALL_CLAIM.TEMPEDLOCATION_FLG = 'Y'") 'Modified as per work request NBAR-3294

        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" AND call_account.LOCATION_CODE <> '003000'")
        SqlStrBuilderGBS.Append(" AND call_account.LOCATION_CODE <> '001406'")
        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()

        If oReaderGBS.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nTEMP_FEE, "nTempPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nTempTotal")
            oParamFld.Text = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))
        End If
        oReaderGBS.Close()



        '  '********GBS GBS GBS ********  get escalations ********** GBS GBS GBS
        '**and exclud the WMI 003000 and Sunston Hotel Prorerty 001405**
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select count(*) as nTotal From CALL, ESCALATION_OUTCOME , CALL_CLAIM, CALL_ACCOUNT ")
        SqlStrBuilderGBS.Append(" Where CALL.CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND CALL.STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID ")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_ID = CALL_CLAIM.CALL_ID ")
        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" AND CALL_CLAIM.CALL_CLAIM_ID = CALL_ACCOUNT.CALL_CLAIM_ID ")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID ")
        SqlStrBuilderGBS.Append(" and (call_account.LOCATION_CODE <> '003000' and call_account.LOCATION_CODE <> '001406')")
        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()
        If oReaderGBS.HasRows Then
            getProcessingFees(oConn, oRpt, cAHS_ID, nESCALATE_FEE, "nEscalationPrice")
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nEscalationsTotal")
            oParamFld.Text = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))
        End If


        oReaderGBS.Read()


        '********GBS GBS GBS ******** get total transmissions ********** GBS GBS GBS
        '**and exclud the WMI 003000 and Sunston Hotel Prorerty 001405**
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select COUNT(TOC.TRANSMISSION_OUTCOME_ID) AS transmissionCount ")
        SqlStrBuilderGBS.Append(" From TRANSMISSION_OUTCOME TOC, TRANSMISSION_OUTCOME_STEP TOS ")
        SqlStrBuilderGBS.Append(" Where TOS.STATUS = 'PROCESSED' ")
        SqlStrBuilderGBS.Append(" AND TOC.TRANSMISSION_OUTCOME_ID = TOS.TRANSMISSION_OUTCOME_ID ")
        SqlStrBuilderGBS.Append(" AND TOS.TRANSMISSION_TYPE_ID = 1 ")
        SqlStrBuilderGBS.Append(" AND TOC.CALL_ID IN (Select DISTINCT CALL.CALL_ID From CALL, call_claim, call_account ")
        SqlStrBuilderGBS.Append(" Where CALL.CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND CALL.STATUS = 'COMPLETED' ")
        SqlStrBuilderGBS.Append(" AND CALL.CALL_ID = call_claim.CALL_ID")
        SqlStrBuilderGBS.Append(" and (call_account.LOCATION_CODE <> '003000' and call_account.LOCATION_CODE <> '001406')")
        SqlStrBuilderGBS.Append(" AND call_account.CALL_CLAIM_ID = call_claim.CALL_CLAIM_ID")
        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID & ") ")
        oCmd.CommandText = SqlStrBuilderGBS.ToString
        oReaderGBS = oCmd.ExecuteReader()
        oReaderGBS.Read()
        If oReaderGBS.HasRows Then
            nTotalTransmissions = CInt(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("transmissionCount")))
        End If
        oReaderGBS.Read()

        oReaderGBS.Close()


        '********GBS GBS GBS ******** get faxed pages ********** GBS GBS GBS
        '   get faxed pages
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select Sum(PAGE_COUNT) AS pageCount From TRANSMISSION_OUTCOME_STEP ")
        SqlStrBuilderGBS.Append(" Where TRANSMISSION_TYPE_ID = 1 ")
        SqlStrBuilderGBS.Append(" AND STATUS = 'PROCESSED' ")
        SqlStrBuilderGBS.Append(" AND TRANSMISSION_OUTCOME_STEP.TRANSMISSION_SEQ_STEP_ID IN (SELECT TRANSMISSION_SEQ_STEP_ID + 0 ")
        SqlStrBuilderGBS.Append(" From TRANSMISSION_SEQ_STEP Where Exists (Select 'X' ")
        SqlStrBuilderGBS.Append(" From ROUTING_PLAN RP, CALL ")
        SqlStrBuilderGBS.Append(" Where CLIENT_HRCY_STEP_ID = " & strClient_ID & " AND CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND RP.ACCNT_HRCY_STEP_ID = CALL.ACCNT_HRCY_STEP_ID + 0 ")
        SqlStrBuilderGBS.Append(" AND RP.ROUTING_PLAN_ID = TRANSMISSION_SEQ_STEP.ROUTING_PLAN_ID + 0))")
        'oCmd.CommandText = SqlStrBuilderGBS.ToString
        'oReaderGBS = oCmd.ExecuteReader()
        ' oReaderGBS.Read()
        'Dim IntPagerGBS As Integer
        ' IntPagerGBS = 0

        ' getProcessingFees(oConn, oRpt, cAHS_ID, nFAX_FEE, "nFaxedPagesTotal")
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nFaxedPagesTotal")
        ' IntPagerGBS = CInt((oReaderGBS.GetValue(oReaderGBS.GetOrdinal("pageCount"))))

        '  If IntPagerGBS >= IntPagerWMI Then
        '  oParamFld.Text = CStr(IntPagerGBS - IntPagerWMI)
        '  Else
        '  oParamFld.Text = CStr(IntPagerGBS)
        '  End If



        ' oCmd.CommandText = SqlStrBuilderWMI.ToString
        ' oReaderWMI = oCmd.ExecuteReader()
        ' oReaderWMI.Read()
        ' Dim intPrintWMI As Integer
        ' intPrintWMI = 0
        ' intPrintWMI = CInt(oReaderWMI.GetValue(oReaderWMI.GetOrdinal("Printcount")))


        '********GBS GBS GBS ********  prints ********** GBS GBS GBS
        '   get printed pages
        '***********************&&&&&&&&&&--prints--&&&&&&&&&&&&&&*****************
        SqlStrBuilderGBS = New System.Text.StringBuilder("")
        SqlStrBuilderGBS.Append(" Select COUNT(transmission_outcome.transmission_outcome_id) AS Printcount ")
        SqlStrBuilderGBS.Append(" FROM transmission_outcome,transmission_outcome_step, CALL, call_claim, Call_account  ")
        SqlStrBuilderGBS.Append(" WHERE transmission_outcome.CALL_ID = CALL.CALL_ID ")
        SqlStrBuilderGBS.Append(" and  CALL.CALL_ID = call_claim.CALL_ID ")
        SqlStrBuilderGBS.Append(" and  Call_account.CALL_CLAIM_ID = call_claim.CALL_CLAIM_ID ")
        SqlStrBuilderGBS.Append(" AND transmission_outcome.transmission_outcome_id = transmission_outcome_step.transmission_outcome_id ")
        SqlStrBuilderGBS.Append(" AND (transmission_outcome_step.status = 'PROCESSED' AND transmission_outcome_step.status <> 'FAILED') ")
        SqlStrBuilderGBS.Append(" AND CALL.CLIENT_HRCY_STEP_ID = " & strClient_ID)
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilderGBS.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilderGBS.Append(" AND transmission_outcome.RESUBMITTED_FLG <> 'Y' ")
        SqlStrBuilderGBS.Append(" AND transmission_outcome_step.RESUBMITTED_FLG <>'Y' ")
        SqlStrBuilderGBS.Append(" AND transmission_outcome_step.transmission_type_id = 2 ")
        SqlStrBuilderGBS.Append(" and call_account.LOCATION_CODE <>'003000' ")
        SqlStrBuilderGBS.Append(" and call_account.LOCATION_CODE <>'001406' ")


        ' oCmd.CommandText = SqlStrBuilderGBS.ToString
        ' oReaderGBS = oCmd.ExecuteReader()
        ' oReaderGBS.Read()
        ' Dim intPrintGBS As Integer
        ' intPrintGBS = 0
        '   getProcessingFees(oConn, oRpt, cAHS_ID, nPRINT_FEE, "nPrintedPagesPrice")
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPrintedPagesTotal")
        '  intPrintGBS = CInt(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("Printcount")))
        '  If intPrintGBS >= intPrintWMI Then
        ' oParamFld.Text = CStr(intPrintGBS - intPrintWMI)
        '    Else
        'oParamFld.Text = CStr(intPrintGBS)

        ' End If

        'oReaderGBS.Close()

        nINFFreePercentage = getFreeINFOPercent(oConn, cAHS_ID)
        nINFFreeCalls = CInt(nTotalClaimsReceived * (nINFFreePercentage / 100))
        If nTotalINFClaims > nINFFreeCalls Then
            oParamFld = oRpt.DataDefinition.FormulaFields.Item("nINFCalls2Bill")
            oParamFld.Text = CStr(nTotalINFClaims - nINFFreeCalls)
        End If
        oConn.Close()
    End Sub

    '===================================================================================================
    Sub getFees(ByVal oConn As OracleConnection, _
                 ByVal oRpt As GBSBillingSummaryFixed, _
                 ByVal cAHS_ID As String, _
                 ByVal cLOB As String, _
                 ByVal cCallType As String, _
                 ByVal nFeeTypeId As Integer, _
                 ByVal cFormulaName As String)
        Dim cSQL As String
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader


        cSQL = "Select FEE_AMOUNT From FEE Where ACCNT_HRCY_STEP_ID = " & cAHS_ID
        cSQL = cSQL + " AND CALL_TYPE = '" & cCallType & "'"
        cSQL = cSQL + " AND LOB_CD = '" & cLOB & "' "
        cSQL = cSQL + " AND FEE_TYPE_ID = " & nFeeTypeId

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
                            ByVal oRpt As GBSBillingSummaryFixed, _
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
                        ByVal oRpt As BillingSummaryFixed, _
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

    Function getMatchLOB(ByVal strLOB As String, ByVal sTotalArray As String, ByVal inReader As Integer) As String
        Dim saveLOB As String = ""
        Dim strCompLOBTotal As String
        Dim nTotal1 As Integer
        Dim nTotal As Integer
        Dim strParameteText As String
        Dim n As Integer
        Dim pos As Integer
        Dim arrCList As String
        Dim arrSplitString() As String = Split(sTotalArray, ":")

        saveLOB = strLOB
        'if lob was found in WMI THEN LOOK FOR IT 
        If InStr(sTotalArray, strLOB) > 0 Then
            'THEN SPLIT AND CREAT THE ARRAY SEARCH FOR LOB 
            For n = 1 To UBound(arrSplitString)
                arrCList = arrSplitString(n)
                arrCList = arrCList.Replace(":", "")
                If InStr(arrCList, strLOB) > 0 Then
                    strCompLOBTotal = Right(arrCList, Len(arrCList) - Len(strLOB))
                    nTotal1 = inReader 'CInt(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal"))) '- CInt(strCompLOBTotal)

                    If nTotal1 >= CInt(strCompLOBTotal) Then 'nTotalGBS1 >= CInt(strCompLOBTotal)
                        nTotal = nTotal1 - CInt(strCompLOBTotal)
                    Else
                        nTotal = CInt(nTotal)
                    End If 'nTotalGBS1 >= CInt(strCompLOBTotal)

                End If

            Next n ''THEN SPLIT AND CREAT THE ARRAY SEARCH FOR LOB 
            strParameteText = CStr(nTotal)
            ' oParamFld.Text = CStr(nTotalGBS)
        Else
            strParameteText = CStr(inReader)
            ' oParamFld.Text = CStr(oReaderGBS.GetValue(oReaderGBS.GetOrdinal("nTotal")))

        End If

        Return strParameteText

    End Function



End Module
