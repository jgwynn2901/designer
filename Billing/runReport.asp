<!--#include file="..\lib\genericSQL.asp"-->
<%
Response.Buffer = False
Response.Expires = -1

const FNSDesigner = "FNSDesigner"

dim cAHS, cStartDate, cSP, cSQL, oExcel, cTmpFile, cDownloadLocation

nTimeOut = Server.ScriptTimeout
Server.ScriptTimeout = 600	'	10 min

cAHS = Request.QueryString("AHS")
cStartDate = UCase(Request.QueryString("DATE"))
cStartDate = "1-" & left(cStartDate,3) & "-" & right(cStartDate,4)
deletePreviousReport
Set oExcel = Server.CreateObject("ExcelClass.XLSClass")
oExcel.cBackground = "#d6cfbd"
oExcel.writeMsg "Retrieving data from database"
if cAHS = 72 then
	cSP = "{call billingReportMarriot.ProcessCallInfo('" 
else
	cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
end if
cSP = cSP & cStartDate & "')}"
Conn.Execute cSP
if cAHS = 71 or 75 then	'	Fremont
	doFremont
	genFremontXLS
else
	genXLS
end if
'
'	update history table
cSQL = "INSERT INTO BILLING_HISTORY (MMM_YYYY,CREATED_BY,CREATED_ON,FILENAME,FILE_PATH,SERVER_NAME,AHS_ID) " & _
		"VALUES('" & UCase(Request.QueryString("DATE")) & "','" & _
		Session("NAME") & "','" & _
		now & "','" & _
		cTmpFile & "','" & _
		cDownloadLocation & "','" & _
		" '," & _
		cAHS & ")"
Conn.Execute(cSQL)
Set oExcel = Nothing
Server.ScriptTimeout = nTimeOut
'
'
sub deletePreviousReport
dim cSQL, oRS, nBillID

cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
do while not oRS.eof
	nBillID = CInt(oRS.fields(0))
	cSQL = "DELETE From BILLING_DETAIL Where BILLING_ID=" & nBillID
	Conn.Execute cSQL
	cSQL = "DELETE From BILLING Where BILLING_ID=" & nBillID
	Conn.Execute cSQL
	oRS.moveNext
loop	
oRS.close
set oRS = nothing
end sub

sub genXLS
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, nAdtlFees, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim cCmpParent

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATE")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
'
dRepDate = cDate("1-" & left(cStartDate,3) & "-" & right(cStartDate,4))
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
cBillID = oRS.fields(0)
oRS.close
'
cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cStartDate & "-" & cTime & ".xls"
oExcel.cDestinationFileName = cTmpFile
oExcel.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
oExcel.cExcelRangeName = "ODBCRange"
oExcel.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
oExcel.cDownloadLocation = cDownloadLocation
oExcel.openXLS
oExcel.writeMsg "Generating spreadsheet"
oExcel.writeCell "Account", cCustName
oExcel.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Svc_Fee,Adl_Fee,Total"
Do While Not oRS.EOF
	if isNull(oRS.Fields("PARENT_NAME").Value) then
		cParent = ""
	else
		cParent = oRS.Fields("PARENT_NAME").Value
	end if
	nCalls = 0
	nTotSvcFee = 0
	nTotAdnFee = 0
	nTotalFee = 0
	do while Not oRS.EOF 
		if isNull(oRS.Fields("PARENT_NAME").Value) then
			cCmpParent = ""
		else
			cCmpParent = oRS.Fields("PARENT_NAME").Value
		end if
		if cParent = cCmpParent then
			nCalls = nCalls + 1
			nAdtlFees = CSng(oRS.Fields("TOTAL_FAX_FEE").Value)+CSng(oRS.Fields("TEMP_FEE").Value)+CInt(oRS.Fields("ESCALATE_FEE").Value)+CSng(oRS.Fields("VENDOR_FEE").Value)
			nTotAdnFee = nTotAdnFee + nAdtlFees
			nTotSvcFee = nTotSvcFee + CSng(oRS.Fields("SERVICE_FEE").Value)
			nTotalFee = nTotalFee + CSng(oRS.Fields("SERVICE_FEE").Value) + nAdtlFees
			cValues =	"'" & removeSngQuote(oRS.Fields("PARENT_NAME").Value) & "'," & _
				CStr(oRS.Fields("Call_ID").Value) & ",'" & _
				oRS.Fields("Call_Type").Value & "','" & _
				oRS.Fields("CALLSTATUS").Value & "','"
				if isNull(oRS.Fields("LOSS_DATE").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & oRS.Fields("LOSS_DATE").Value
				end if
				cValues = cValues & "','" & _
				CStr(oRS.Fields("CALL_END_TIME").Value) & "','"
				if isNull(oRS.Fields("CLAIM_NUMBER").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & oRS.Fields("CLAIM_NUMBER").Value
				end if
				cValues = cValues & "','" & _
				oRS.Fields("LOB_CD").Value & "','"
				if isNull(oRS.Fields("POLICY_NUMBER").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & oRS.Fields("POLICY_NUMBER").Value
				end if
				cValues = cValues & "','" & _
				removeSngQuote(oRS.Fields("ACCOUNT_NAME").Value) & "','"
				redim aNameParts(1)
				aNameParts(0) = oRS.Fields("CALLER_FIRST_NAME").Value
				aNameParts(1) = oRS.Fields("CALLER_LAST_NAME").Value
				cValues = cValues & getName(aNameParts) & "','"
				redim aNameParts(1)
				aNameParts(0) = oRS.Fields("EMPLOYEE_FIRST_NAME").Value
				aNameParts(1) = oRS.Fields("EMPLOYEE_LAST_NAME").Value
				cValues = cValues & getName(aNameParts) & "'," & _
				FormatNumber(oRS.Fields("SERVICE_FEE").Value,1) & "," & _
				FormatNumber(nAdtlFees,1) & "," & _
				FormatNumber(CSng(oRS.Fields("SERVICE_FEE").Value) + nAdtlFees,1)
			oExcel.addRow cFields, cValues
			oRS.MoveNext
		else
			exit do
		end if
	loop
	cValues = "'Total'," & nCalls & ",'','','','','','','','','',''," & nTotSvcFee & "," & nTotAdnFee & "," & nTotalFee
	oExcel.addRow cFields, cValues
	cValues = "'','','','','','','','','','','','','','',''"
	oExcel.addRow cFields, cValues
Loop
oExcel.closeXLS
oExcel.sendFile
oRS.Close
Set oRS = Nothing
end sub

function getName(aParts)
dim x,y

getName = ""
if isArray(aParts) then
	x = uBound(aParts)
	for y=lBound(aParts) to x
		if not isNull(aParts(y)) then
			getName = getName & Trim(removeSngQuote(aParts(y))) & " "
		end if
	next
end if
end function

function removeSngQuote(cString)

if isNull(cString) then
	removeSngQuote = ""
else
	removeSngQuote = Trim(cString)
	if InStr(1,cString,"'",1) then
		removeSngQuote = replace(removeSngQuote,"'","''")
	end if
end if	
end function

' *********************************************************************
'	FREMONT
' *********************************************************************
sub doFremont
Dim cSQL
Dim oRS
Dim nCallCounter
Dim nFaxCounter
Dim nCallTotal
Dim nFaxTotal
Dim nPricePerCall
Dim nMonth

oExcel.writeMsg "Calculating averages for " & Request.QueryString("CUSTNAME")
nMonth = Month(cStartDate)
cSQL = "Select BILLING_DETAIL.CALL_TYPE,BILLING_DETAIL.STATUS From BILLING_DETAIL Where BILLING_DETAIL.CLIENT_NODE_ID = " & cAHS & " And BILLING_DETAIL.STATUS = 'ACTIVE'"
nCallCounter = 0
nFaxCounter = 0
nCallTotal = 0
nFaxTotal = 0

Set oRS = Conn.Execute(cSQL)
Do While Not oRS.EOF
    If oRS.Fields("CALL_TYPE").Value = "C" Then
        nCallCounter = nCallCounter + 1
    ElseIf oRS.Fields("CALL_TYPE").Value = "F" Then
        nFaxCounter = nFaxCounter + 1
    End If
    oRS.MoveNext
Loop
nCallTotal = nCallCounter * 18.5
nFaxTotal = nFaxCounter * 12.75
nPricePerCall = (nCallTotal + nFaxTotal) / (nCallCounter + nFaxCounter)
If nPricePerCall > 16 Then
    nPricePerCall = 16
End If
cSQL = "UPDATE BILLING_DETAIL SET BILLING_DETAIL.SERVICE_FEE = 0 Where BILLING_DETAIL.CLIENT_NODE_ID = " & cAHS & " And BILLING_DETAIL.STATUS = 'ACTIVE'"
Conn.Execute cSQL
cSQL = "UPDATE BILLING_DETAIL SET BILLING_DETAIL.SERVICE_FEE = " & nPricePerCall & " Where BILLING_DETAIL.CLIENT_NODE_ID = " & cAHS & " And BILLING_DETAIL.STATUS = 'ACTIVE' And BILLING_DETAIL.CALL_TYPE<>'I'"
Conn.Execute cSQL
oRS.Close
end sub

sub genFremontXLS
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, nAdtlFees, cValues, cTime
dim aNameParts(), nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nBranchNo, nNextBranchNo

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATE")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
'
dRepDate = cDate("1-" & left(cStartDate,3) & "-" & right(cStartDate,4))
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
cBillID = oRS.fields(0)
oRS.close
'
cSQL = "SELECT BILLING_DETAIL.*, " & _
		"CALL_BRANCH.BRANCH_NUMBER, " & _
		"CALL_BRANCH.BRANCH_OFFICE_NUMBER, " & _
		"CALL_BRANCH.BRANCH_OFFICE_NAME " & _
			"FROM CALL_CLAIM, " & _		
			"CALL_BRANCH, " & _
			"BILLING_DETAIL " & _
			"WHERE CALL_CLAIM.CALL_ID = BILLING_DETAIL.CALL_ID + 0 " & _
			"AND CALL_CLAIM.CALL_CLAIM_ID = CALL_BRANCH.CALL_CLAIM_ID (+) " & _
			"AND BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' " & _
			"Order by CALL_BRANCH.BRANCH_NUMBER DESC, CALL_TYPE, CALL_END_TIME"
			
set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cStartDate & "-" & cTime & ".xls"
oExcel.cDestinationFileName = cTmpFile
oExcel.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Fremont.xls"
oExcel.cExcelRangeName = "ODBCRange"
oExcel.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
oExcel.cDownloadLocation = cDownloadLocation
oExcel.openXLS
oExcel.writeMsg "Generating spreadsheet"
oExcel.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
cFields = "Branch_Name,Branch_No,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Svc_Fee,Adl_Fee,Total"
Do While Not oRS.EOF
	if isNull(oRS.Fields("BRANCH_NUMBER").Value) then
		nBranchNo = ""
	else
		nBranchNo = oRS.Fields("BRANCH_NUMBER").Value
	end if
	nCalls = 0
	nTotSvcFee = 0
	nTotAdnFee = 0
	nTotalFee = 0
	do while Not oRS.EOF 
		if isNull(oRS.Fields("BRANCH_NUMBER").Value) then
			nNextBranchNo = ""
		else
			nNextBranchNo = oRS.Fields("BRANCH_NUMBER").Value
		end if
		if nBranchNo = nNextBranchNo then
			nCalls = nCalls + 1
			nAdtlFees = CSng(oRS.Fields("TOTAL_FAX_FEE").Value)+CSng(oRS.Fields("TEMP_FEE").Value)+CInt(oRS.Fields("ESCALATE_FEE").Value)+CSng(oRS.Fields("VENDOR_FEE").Value)
			nTotAdnFee = nTotAdnFee + nAdtlFees
			nTotSvcFee = nTotSvcFee + CSng(oRS.Fields("SERVICE_FEE").Value)
			nTotalFee = nTotalFee + CSng(oRS.Fields("SERVICE_FEE").Value) + nAdtlFees
			cValues =	"'" & removeSngQuote(oRS.Fields("BRANCH_OFFICE_NAME").Value) & "','" & _
				oRS.Fields("BRANCH_NUMBER").Value & "'," & _
				CStr(oRS.Fields("Call_ID").Value) & ",'" & _
				oRS.Fields("Call_Type").Value & "','" & _
				oRS.Fields("CALLSTATUS").Value & "','"
				if isNull(oRS.Fields("LOSS_DATE").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & oRS.Fields("LOSS_DATE").Value
				end if
				cValues = cValues & "','" & _
				CStr(oRS.Fields("CALL_END_TIME").Value) & "','"
				if isNull(oRS.Fields("CLAIM_NUMBER").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & oRS.Fields("CLAIM_NUMBER").Value
				end if
				cValues = cValues & "','" & _
				oRS.Fields("LOB_CD").Value & "','"
				if isNull(oRS.Fields("POLICY_NUMBER").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & oRS.Fields("POLICY_NUMBER").Value
				end if
				cValues = cValues & "','" & _
				removeSngQuote(oRS.Fields("ACCOUNT_NAME").Value) & "','"
				redim aNameParts(1)
				aNameParts(0) = oRS.Fields("CALLER_FIRST_NAME").Value
				aNameParts(1) = oRS.Fields("CALLER_LAST_NAME").Value
				cValues = cValues & getName(aNameParts) & "','"
				redim aNameParts(1)
				aNameParts(0) = oRS.Fields("EMPLOYEE_FIRST_NAME").Value
				aNameParts(1) = oRS.Fields("EMPLOYEE_LAST_NAME").Value
				cValues = cValues & getName(aNameParts) & "'," & _
				FormatNumber(oRS.Fields("SERVICE_FEE").Value,1) & "," & _
				FormatNumber(nAdtlFees,1) & "," & _
				FormatNumber(CSng(oRS.Fields("SERVICE_FEE").Value) + nAdtlFees,1)
			oExcel.addRow cFields, cValues
			oRS.MoveNext
		else
			exit do
		end if
	loop
	cValues = "'Total',''," & nCalls & ",'','','','','','','','','',''," & nTotSvcFee & "," & nTotAdnFee & "," & nTotalFee
	oExcel.addRow cFields, cValues
	cValues = "'','','','','','','','','','','','','','','',''"
	oExcel.addRow cFields, cValues
Loop
oExcel.closeXLS
oExcel.sendFile
oRS.Close
Set oRS = Nothing
end sub

%>
