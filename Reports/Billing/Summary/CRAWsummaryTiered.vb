Option Explicit On 
Option Strict On

Imports System.Data.OracleClient
Imports System.Text.StringBuilder

Module CRAWsummaryTiered

    Sub getClaimsTieredCRAW(ByRef oRpt As CRAWBillingSummaryTiered, ByVal strUserSite As String, ByVal cAHS_ID As String, ByVal cClient As String, ByVal cReportStartDate As String, ByVal cReportEndDate As String)
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
        Dim nTotalClaimsCalls As Integer
        Dim nTotalClaimsFaxes As Integer
        Dim nTotalClaimsInt As Integer
        Dim intUserSite As Integer
        Dim nClaimsCount As Integer

        Dim SqlStrBuilder As System.Text.StringBuilder

        Const nSERVICE_FEE As Integer = 1
        Const nFAX_FEE As Integer = 2
        Const nFlagASP As Integer = 5

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
        SqlStrBuilder = New System.Text.StringBuilder("")
        SqlStrBuilder.Append(" Select COUNT(call.call_id) AS totalCalls ")
        SqlStrBuilder.Append(" From CALL, CALL_CALLER, CALL_ACCOUNT,CALL_CLAIM, USERS ")
        SqlStrBuilder.Append(" Where CALL_CLAIM.CALL_CLAIM_ID = CALL_ACCOUNT.CALL_CLAIM_ID(+)")
        SqlStrBuilder.Append(" AND CALL.call_id = CALL_CLAIM.CALL_ID ")
        SqlStrBuilder.Append(" AND USERS.USER_ID = call.USER_ID ")
        SqlStrBuilder.Append(" AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 ")
        SqlStrBuilder.Append(" and  CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilder.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilder.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilder.Append(" AND CALL.LOB_CD <> 'INF' ")
        SqlStrBuilder.Append(" AND CALL_ACCOUNT.COVERAGE_CODE  <> 'GI' ")
        If strUserSite = "ASP" Then
            SqlStrBuilder.Append(" AND users.site_id = " & nFlagASP)
        Else
            SqlStrBuilder.Append(" AND users.site_id <> " & nFlagASP)
        End If
        SqlStrBuilder.Append(" AND CLIENT_HRCY_STEP_ID = " & cAHS_ID)

        oConn.ConnectionString = "Data Source=CRAWP; user id=fnsowner;password=ctown_prod;Persist Security info=False;"

        oConn.Open()
        oCmd.CommandText = "ALTER SESSION SET NLS_DATE_FORMAT = 'DD-MON-YYYY HH:MI:SS'"
        oCmd.Connection = oConn
        oCmd.ExecuteNonQuery()
        oCmd.CommandText = SqlStrBuilder.ToString
        oReader = oCmd.ExecuteReader(CommandBehavior.SingleRow)
        oReader.Read()
        If oReader.HasRows Then
            nTotalClaimsReceived = CInt(oReader.GetValue(oReader.GetOrdinal("totalCalls")))
        End If
        oReader.Close()

        '   get call claims
        SqlStrBuilder = New System.Text.StringBuilder("")
        SqlStrBuilder.Append(" Select CALL_ACCOUNT.COVERAGE_CODE AS COVERAGE_CODE, count(*) as nTotal ")
        SqlStrBuilder.Append(" From CALL, CALL_CALLER, CALL_ACCOUNT,CALL_CLAIM, USERS ")
        SqlStrBuilder.Append(" Where CALL_CLAIM.CALL_CLAIM_ID = CALL_ACCOUNT.CALL_CLAIM_ID(+) ")
        SqlStrBuilder.Append(" AND CALL.call_id = CALL_CLAIM.CALL_ID ")
        SqlStrBuilder.Append(" AND USERS.USER_ID = call.USER_ID ")
        SqlStrBuilder.Append(" AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 ")
        SqlStrBuilder.Append(" AND CALL_START_TIME >= TO_DATE('" & cStart & "') ")
        SqlStrBuilder.Append(" AND CALL_START_TIME < TO_DATE('" & cEnd & "') ")
        SqlStrBuilder.Append(" AND STATUS = 'COMPLETED' ")
        SqlStrBuilder.Append(" AND CALL.LOB_CD <> 'INF' ")
        SqlStrBuilder.Append(" AND CALL_ACCOUNT.COVERAGE_CODE  <> 'GI' ")
        SqlStrBuilder.Append(" AND CLIENT_HRCY_STEP_ID = " & cAHS_ID)
        If strUserSite = "ASP" Then
            SqlStrBuilder.Append(" AND users.site_id = 5 ")
        Else
            SqlStrBuilder.Append(" AND users.site_id <> 5 ")
        End If
        SqlStrBuilder.Append(" Group by CALL_ACCOUNT.COVERAGE_CODE")
        oCmd.CommandText = SqlStrBuilder.ToString
        oReader = oCmd.ExecuteReader()
        nTotalClaimsCalls = 0
        oReader.Read()
        If oReader.HasRows Then
            Do

                Select Case oReader.GetValue(oReader.GetOrdinal("COVERAGE_CODE"))
                    Case "AL"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nALCalls")
                    Case "VC"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVCCalls")
                    Case "VD"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nVDCalls")
                    Case "HE"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nHECalls")
                    Case "PL"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nPLCalls")
                    Case "GL"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nGLCalls")
                    Case "SU"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nSUCalls")
                    Case "WC"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nWCCalls")
                    Case "DI"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nDICalls")
                    Case "DM"
                        oParamFld = oRpt.DataDefinition.FormulaFields.Item("nDMCalls")
                    Case Else
                        oParamFld = Nothing '   force an error if not found
                End Select

                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("nTotal")))
                nTotalClaimsCalls = nTotalClaimsCalls + CInt(oParamFld.Text) 'CInt(oReader.GetValue(oReader.GetOrdinal("nTotal")))

            Loop Until Not oReader.Read
        End If
        oReader.Close()

        getTieredFees(oConn, oRpt, cAHS_ID, "C", nSERVICE_FEE, nTotalClaimsCalls, strUserSite)

    End Sub




    Private Sub getTieredFees(ByVal oConn As OracleConnection, _
                    ByVal oRpt As CRAWBillingSummaryTiered, _
                    ByVal cAHS_ID As String, _
                    ByVal cCallType As String, _
                    ByVal nFeeTypeId As Integer, _
                    ByVal nTotalClaims As Integer, ByVal nAspClient As String)

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
        Dim vTotalByTier As String
        Dim cTotalByTier As String = "nTotalTier~"
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
          "  AND SUBSTR(REASON_CODE ,1,3) = '" & nAspClient & "'" & _
         " ORDER BY BEGIN_CALL_RANGE"

        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                cTier = cTierFROM
                cTier = cTier.Replace("~", CStr(x))
                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cTier)
                nBCRColNo = oReader.GetOrdinal("BEGIN_CALL_RANGE")
                nECRColNo = oReader.GetOrdinal("END_CALL_RANGE")
                nClaimsFrom = CInt(oReader.GetValue(nBCRColNo))
                oParamFld.Text = CStr(nClaimsFrom)

                If CInt(oReader.GetValue(nECRColNo)) = 999999 Then
                    '   last tier
                    nClaimsTo = 999999
                Else
                    nClaimsTo = CInt(oReader.GetValue(nECRColNo))

                End If
                cTier = cTierTO
                cTier = cTier.Replace("~", CStr(x))
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
                cPrice = cPrice.Replace("~", CStr(x))
                cTotalPrice = cTotalPrice.Replace("~", CStr(x))
                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cPrice)
                oParamFld.Text = CStr(oReader.GetValue(oReader.GetOrdinal("FEE_AMOUNT")))


                vTotalByTier = cTotalByTier
                vTotalByTier = vTotalByTier.Replace("~", CStr(x))


                oParamFld = oRpt.DataDefinition.FormulaFields.Item(cTotalPrice)
                nFeeAmntColNo = oReader.GetOrdinal("FEE_AMOUNT")
                If nTotalClaims > nClaimsTo Then
                    oParamFld.Text = CStr(CDbl(oReader.GetValue(nFeeAmntColNo)) * nClaimsTo)

                    oParamFld = oRpt.DataDefinition.FormulaFields.Item(vTotalByTier)
                    oParamFld.Text = CStr(nClaimsTo)
                Else
                    oParamFld.Text = CStr(CDbl(oReader.GetValue(nFeeAmntColNo)) * (nTotalClaims - nClaimsFrom + 1))

                    oParamFld = oRpt.DataDefinition.FormulaFields.Item(vTotalByTier)
                    oParamFld.Text = CStr(nTotalClaims - nClaimsFrom + 1)
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


End Module
