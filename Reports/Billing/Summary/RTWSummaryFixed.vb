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

Module RTWSummaryFixed

    Sub getRTWClaimsFixedPricing(ByRef oRpt As RTWBillingSummaryFixed, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String, ByVal lIsCCE As Boolean)
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
        '----------------------------------------------
        '
        '   get call claims
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID(+) " & _
             "AND (CALL_CALLER.CALLER_TYPE IS NULL OR (SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) IN ('C','I','O','E'))) "

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

                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))

                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORCalls")
                        getFees(oConn, oRpt, cAHS_ID, "WOR", "C", nSERVICE_FEE, "nWORPriceC")
                        If Not lIsINFO Then
                            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        End If
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select

            Loop Until Not oReader.Read
        End If
        oReader.Close()

        'get faxed claims
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
                lIsINFO = False
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORFaxes")
                        getFees(oConn, oRpt, cAHS_ID, "WOR", "F", nSERVICE_FEE, "nWORPriceF")
                        If Not lIsINFO Then
                            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        End If
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select

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
             "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) ='N' "
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
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORInternet")
                        getFees(oConn, oRpt, cAHS_ID, "WOR", "N", nSERVICE_FEE, "nWORPriceI")
                        If Not lIsINFO Then
                            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        End If
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select

            Loop Until Not oReader.Read
        End If
        oReader.Close()

        'Mail fees 
        cSQL = "Select LOB_CD, count(*) as nTotal From CALL, CALL_CALLER ,CALL_CLAIM " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND STATUS = 'COMPLETED' " & _
             "AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
             "AND CALL.CALL_ID = CALL_CLAIM.CALL_ID " & _
             "AND SUBSTR(CALL_CALLER.CALLER_TYPE, 1, 1) ='M' "
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
                Select Case oReader.GetValue(oReader.GetOrdinal("LOB_CD"))
                    Case "WOR"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWORMails")
                        getFeesForMail(oConn, oRpt, cAHS_ID, "WOR", "M", nSERVICE_FEE, "nWORPriceM")
                        If Not lIsINFO Then
                            oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                        End If
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select

            Loop Until Not oReader.Read
        End If
        oReader.Close()
        cSQL = "Select count(*) as nTotal From CALL, ESCALATION_OUTCOME " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
             "AND CALL.STATUS = 'COMPLETED' " & _
            "AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID "
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
        oConn.Close()
    End Sub

    Sub getFees(ByVal oConn As OracleConnection, _
                ByVal oRpt As RTWBillingSummaryFixed, _
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
    Sub getFeesForMail(ByVal oConn As OracleConnection, _
                    ByVal oRpt As RTWBillingSummaryFixed, _
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
                " AND LOB_CD = '" & cLOB & "' "
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
                            ByVal oRpt As RTWBillingSummaryFixed, _
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
                        ByVal oRpt As RTWBillingSummaryFixed, _
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

