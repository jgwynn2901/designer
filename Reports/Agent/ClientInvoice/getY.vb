'
'FNS DESIGNER/FNS CLAIMCAPTURE   
'Client					:	CCE
'Object					:	Default.asp   
'Script Date: 01/19/2007		Script By: Narayan.Ramachandran
'Work Request/ILog #	:	MROU-2726
'Requirement			: Changing the date format as the new search criteria the in the start date and end date
' The End Date criteria has been inttroduced.
'*/
'/*--------------------------------------------------------------------------------------------------------------------*/




Option Explicit On 
Option Strict On

Imports System.Data.OracleClient

Module getY

    Function planY(ByVal cAHS As String, _
                    ByVal cReportDate As String, _
                    ByVal cReportEndDate As String, _
                    ByVal oConn As OracleConnection, _
                    ByVal cAccountName As String, _
                    ByVal oRpt As AgentY, _
                    ByVal lFirstMonth As Boolean) As Boolean

        Const UNITSPERCLAIM As Integer = 4
        Const UNITSPERINFOCALL As Integer = 1
        Const FEEPERSERVICEUNIT As Single = 3.96
        Const ANNUALFEE As Integer = 850
        Const FREESERVICEUNITSPERMONTH As Integer = 20
        Const ESCALATIONFEE As Single = 7.5

        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim dRepStart As Date
        Dim dRepEnd As Date
        Dim cStart, cEnd As String
        Dim nTotalINFClaims As Integer
        Dim nTotalClaimsReceived As Integer
        Dim nTotalEscalations As Integer
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oValueFld As CrystalDecisions.CrystalReports.Engine.FieldObject

        dRepStart = CDate(cReportDate)
        dRepEnd = CDate(cReportEndDate) 'DateAdd(DateInterval.Month, 1, dRepStart)
        cStart = UCase$(Format(dRepStart, "dd-MMM-yyyy"))
        cEnd = UCase$(Format(dRepEnd, "dd-MMM-yyyy"))
        dRepEnd = DateAdd(DateInterval.Day, 1, dRepEnd)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("cPeriod")
        oParamFld.Text = "'" & cStart & "'" ' "'" & MonthName(Month(dRepStart)) & " " & Day(dRepStart) & "-" & Day(dRepEnd) & ", " & Year(dRepStart) & "'"
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("cAccountName")
        oParamFld.Text = "'" & cAccountName & "'"

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("cInvoiceNo")
        oParamFld.Text = "'" & cAHS & Format(Month(dRepStart), "00") & Right(Format(Year(dRepStart), "00"), 2) & "'"

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nClaimCalls")
        nTotalClaimsReceived = getClaims(cAHS, oConn, cStart, cEnd)
        oParamFld.Text = CStr(nTotalClaimsReceived)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nSUperClaim")
        oParamFld.Text = CStr(UNITSPERCLAIM)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nSUperInfoCall")
        oParamFld.Text = CStr(UNITSPERINFOCALL)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nFeePerSU")
        oParamFld.Text = CStr(FEEPERSERVICEUNIT)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nYearlyFee")
        oParamFld.Text = CStr(ANNUALFEE)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nInfoCalls")
        nTotalINFClaims = getINFCalls(cAHS, oConn, cStart, cEnd)
        oParamFld.Text = CStr(nTotalINFClaims)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nSUperMonth")
        If lFirstMonth Then
            oParamFld.Text = CStr(0)
        Else
            oParamFld.Text = CStr(FREESERVICEUNITSPERMONTH)
        End If
        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nEscalationsEach")
        oParamFld.Text = CStr(ESCALATIONFEE)

        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nQtyEscalations")
        nTotalEscalations = getEscalations(cAHS, oConn, cStart, cEnd)
        oParamFld.Text = CStr(nTotalEscalations)

        If (nTotalClaimsReceived + nTotalINFClaims + nTotalEscalations) <> 0 Then
            planY = True
        Else
            planY = False
        End If
    End Function

    Function getClaims(ByVal cAHS As String, _
                    ByVal oConn As OracleConnection, _
                    ByVal cStartDate As String, _
                    ByVal cEndDate As String) As Integer
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        '   get total number of claims (excluding INF calls)
        cSQL = "Select COUNT(call.call_id) AS totalCalls " & _
                "From CALL Where CALL_START_TIME >= TO_DATE('" & cStartDate & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEndDate & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND LOB_CD <> 'INF' " & _
                "AND ACCNT_HRCY_STEP_ID = " & cAHS
        oCmd.CommandText = cSQL
        oCmd.Connection = oConn
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        getClaims = 0
        If oReader.HasRows Then
            getClaims = CInt(oReader.GetValue(oReader.GetOrdinal("totalCalls")))
        End If
        oReader.Close()
        oReader.Dispose()
        oCmd.Dispose()
    End Function

    Function getINFCalls(ByVal cAHS As String, _
                            ByVal oConn As OracleConnection, _
                            ByVal cStartDate As String, _
                            ByVal cEndDate As String) As Integer
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        '   get INF claims
        cSQL = "Select count(*) as nTotal From CALL " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStartDate & "') " & _
                "AND CALL_START_TIME < TO_DATE('" & cEndDate & "') " & _
                "AND STATUS = 'COMPLETED' " & _
                "AND LOB_CD = 'INF' " & _
                "AND ACCNT_HRCY_STEP_ID = " & cAHS
        oCmd.CommandText = cSQL
        oCmd.Connection = oConn
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        getINFCalls = 0
        If oReader.HasRows Then
            getINFCalls = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()
        oReader.Dispose()
        oCmd.Dispose()
    End Function


    Function getEscalations(ByVal cAHS As String, _
                    ByVal oConn As OracleConnection, _
                    ByVal cStartDate As String, _
                    ByVal cEndDate As String) As Integer
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        '   get escalations
        cSQL = "Select count(*) as nTotal From CALL, ESCALATION_OUTCOME " & _
                "Where CALL_START_TIME >= TO_DATE('" & cStartDate & "') " & _
             "AND CALL_START_TIME < TO_DATE('" & cEndDate & "') " & _
             "AND CALL.STATUS = 'COMPLETED' " & _
            "AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID " & _
            "AND ACCNT_HRCY_STEP_ID = " & cAHS
        oCmd.CommandText = cSQL
        oCmd.Connection = oConn
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        getEscalations = 0
        If oReader.HasRows Then
            getEscalations = CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))
        End If
        oReader.Close()
        oReader.Dispose()
        oCmd.Dispose()
    End Function

End Module
