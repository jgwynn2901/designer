<%

'/*--------------------------------------------------------------------------------------------------------------------*/
'/* WORK REQUEST – JPRI-0828 [TDS/SOW Document # if Exists : FNS_TDS_JPRI-0820 v1.1.doc]

'FNS DESIGNER/FNS CLAIMCAPTURE
'Client					:	CCE
'Object					:	runReport.asp
'Script Date: 11/29/2004		Script By: Avra Banerjee
'Work Request/ILog #	:	JPRI-0828
'Requirement			: 	Add a column for “Caller Type” to the report
'							sent in a separate attachment.  The name can be put after the last column.
'							This is the Claim Capture Express Detail Billing Report
'							(for ALL CCE clients-detail is not available for specific clients within CCE).
'							It is run from the “Billing Reports” menu on FNSDESIGNER-Production.
'							Although the request is for Unitrin Personal Lines under the CCE account,
'							this change will affect all CCE clients.


'Change Line # 			:	275-285, 312-314, 414-415, 431, 433, 444, 446, 458, 460
'*/
'/*--------------------------------------------------------------------------------------------------------------------*/

'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR JPRI-0919 [TDS/SOW Document # if Exists : FNS_TDS_WR JPRI-0919.doc v1.1.doc]

'FNS CLAIMCAPTURE
'Client						: ARG
'Object						: runReport.asp
'Script Date				: 05/18/2005
'Script By					: Sandip Dey
'Work Request/ILog #		: JPRI-0919
'Requirement				: Add a new field - ACCOUNT NAME - in an existing report - Billing Report for ARG
'Change#					: CHANGE# 1,2,3,4 in the function genARG
'--------------------------------------------------------------------------------------------------------------------*/

'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR JPRI-0851 [TDS/SOW Document # NA]

'FNS DESIGNER
'Client						: ARG
'Object						: runReport.asp
'Script Date				: 16/08/2005
'Script By					: Subhankar Sarkar
'Work Request/ILog #		: JPRI-0851
'Requirement				: ARG Billing Report needs the Account Name field to be between Branch and Call Number
'Change#					: Column position changes is made in BillingArg.xls file. Corresponding changes are
'							  made in genARG function
'--------------------------------------------------------------------------------------------------------------------*/


'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR LPIC-0063 [TDS/SOW Document # if Exists : FNS_TDS_WR LPIC-0063.doc v1.1.doc]

'FNS CLAIMCAPTURE
'Client						: CNL
'Object						: runReport.asp
'Script Date				: 07/22/2005
'Updated On					: 11/07/2005
'Script By					: R.Narayan
'Work Request/ILog #		: LPIC-0063
'Reqirement	                '1)No pricing info for APC in the detail reports need to be shown.
							'2)Sub-total to be Based On grouped Call No

'--------------------------------------------------------------------------------------------------------------------*/
'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR ERCH-0017

'FNS CLAIMCAPTURE
'Client						: ALM
'Object						: runReport.asp
'Script Date				: 07/22/2005
'Script By					: Sutapa Majumdar & Subhankar Sarkar
'Work Request/ILog #		: ERCH-0017


'--------------------------------------------------------------------------------------------------------------------*/


%>
<%
Response.Buffer = False
Response.Expires = -1

'if request.QueryString("SCHEDULE") = "Yes" then 'bypasses login , passes it here
 'Session("ConnectionString") = "DSN=FNSPRODUCTION;UID=FNSOWNER;PWD=CTOWN_PROD"
 'Session("NAME") = "ADMINISTRATOR"
'Session("PASSWORD") = "CELEBRATION2003"
'end if

'/*--------------------------------------------------------------------------------------------------------------------*/
'/* WORK REQUEST – JPRI-0828 [TDS/SOW Document # if Exists : FNS_TDS_JPRI-0820 v1.1.doc]

'FNS DESIGNER/FNS CLAIMCAPTURE
'Client					:	CCE
'Object					:	runReport.asp
'Script Date			: 11/29/2004		Script By: Avra Banerjee
'Work Request/ILog #	:	JPRI-0828
'Requirement			: 	Add a column for “Caller Type” to the report
'							sent in a separate attachment.  The name can be put after the last column.
'							This is the Claim Capture Express Detail Billing Report
'							(for ALL CCE clients-detail is not available for specific clients within CCE).
'							It is run from the “Billing Reports” menu on FNSDESIGNER-Production.
'							Although the request is for Unitrin Personal Lines under the CCE account,
'							this change will affect all CCE clients.


'Change Line # 			:	275-285, 312-314, 414-415, 431, 433, 444, 446, 458, 460
'*/
'/*--------------------------------------------------------------------------------------------------------------------*/

'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR JPRI-0919 [TDS/SOW Document # if Exists : FNS_TDS_WR JPRI-0919.doc v1.1.doc]

'FNS CLAIMCAPTURE
'Client						: ARG
'Object						: runReport.asp
'Script Date				: 05/18/2005
'Script By					: Sandip Dey
'Work Request/ILog #		: JPRI-0919
'Requirement				: Add a new field - ACCOUNT NAME - in an existing report - Billing Report for ARG
'Change#					: CHANGE# 1,2,3,4 in the function genARG
'--------------------------------------------------------------------------------------------------------------------*/

'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR JMAR-0422

'FNS CLAIMCAPTURE
'Client						: ALM
'Object						: runReport.asp
'Script Date				: 05/09/2007
'Script By					: Prashant Shekhar
'Work Request/ILog #		: JMAR-0422
'Requirement				: Add a new field - BRANCH_OFFICE_NAME - in an existing report - Billing Report for ALM
'Change						: CHANGE in the function genALMxls to inlcude the Branch_Office _Name column.
'--------------------------------------------------------------------------------------------------------------------*/

'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR NBAR-5128

'FNS FNSDESIGNER
'Client						: CCE
'Object						: runReport.asp
'Script Date				: 05/12/2010
'Script By					: Sohail Iqbal
'Work Request/ILog #		: NBAR-5128
'Requirement				: The Insured State is required to populate on the FNS Designer Billing Detail Report for CCE.
'Change						: The genCCE function in the runReport.asp page will be modified to read the insured state from
'							: the billing detail table and insert it into the Excel report template in the Insured_State field.
'--------------------------------------------------------------------------------------------------------------------*/

'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR KFAB-6227

'FNS FNSDESIGNER
'Client						: AFF
'Object						: runReport.asp
'Script Date				: 12/07/2010
'Script By					: Syed Waqas Ahmed Shah
'Work Request/ILog #		: KFAB-6227
'Requirement				: Remove the dependency of AFF on CCE
'							  Remove Rep_Pro_Claim_Trained and Entered_During_Business_hours from Billing Detail Report Template
'Change						: genAFF will be created to remove the dependency of AFF on CCE. and to remove the two columns
'							: mentioned above.
'--------------------------------------------------------------------------------------------------------------------*/

'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – WR TPAL-0146

'FNS FNSDESIGNER
'Client						: TOW
'Object						: runReport.asp
'Script Date				: 02/21/2012
'Script By					: Syed Waqas Ahmed Shah
'Work Request/ILog #		: TPAL-0146
'Requirement				: Tower group billing fees / reports setup
'Change						: "genXLS_TOWASP" will be created for TOWER ASP
'--------------------------------------------------------------------------------------------------------------------*/


%>
<!--#include file="..\..\lib\genericSQL.asp"-->
<!--#include file="billing.inc"-->
<%

'Conn.Close
'Conn.Open "DSN=FNSP;UID=FNSOWNER;PWD=CTOWN_PROD;SERVER=FNSP"

const FNSDesigner = "FNSDesigner"
dim cAHS, cStartDate, cEndDate, cSP, cSQL, oExcel, cTmpFile, cDownloadLocation,cParServ,cCompServ,cCustname
dim lWithError, lIsAgentBill ,confFlag
dim cCCE, dRepStart, dRepEnd, cPer, cEmployeeSSN

nTimeOut = Server.ScriptTimeout
Server.ScriptTimeout = 7200	'	120 min

if Application("lExecutingBillingReport") then
	Response.redirect "inUse.htm"
else
	Application.Lock
	Application("lExecutingBillingReport") = true
	Application.UnLock
end if

cAHS = Request.QueryString("AHS")
cCustname= Request.QueryString("CUSTNAME")

dRepStart = CDate(Request.QueryString("DATEFROM"))
dRepEnd = CDate(Request.QueryString("DATETO"))
dRepEnd = dateadd("d",1,dRepEnd)
cStart = day(dRepStart ) & "-" & MonthName(month(dRepStart ),true) & "-" & year(dRepStart )
cEnd = day(dRepEnd) & "-" & MonthName(month(dRepEnd),true) & "-" & year(dRepEnd)
lIsAgentBill = InStr(1, Request.QueryString, "AgtIntBill", 1) <> 0
cCCE = Request.QueryString("CUSTCODE")
cEmployeeSSN = "#########"

Conn.Execute "ALTER SESSION SET NLS_DATE_FORMAT = 'DD-MON-YYYY HH:MI:SS'"
Set oExcel = Server.CreateObject("ExcelClass.XLSClass")

with oExcel
	.cBackground = "#d6cfbd"
	.writeMsg "Retrieving data from database"
end with

if lIsAgentBill then
	doAgentBilling Conn, oExcel
else

	deletePreviousReport

	select case Clng(cAHS)
	    case SELNo
			cSP = "{call billingreport_SEL1.ProcessCallInfo('" & cAHS & "', '"
		case SRSNo
			cSP = "{call billingreport_SRS.ProcessCallInfo('" & cAHS & "', '"
	    case CHBNo
			cSP = "{call billingreport_CHB.ProcessCallInfo('" & cAHS & "', '"
		case MARNo
			cSP = "{call billingReportMarriot.ProcessCallInfo('"
		case CCENo
			cSP = "{call billingReportCCE.ProcessCallInfo('"
		case AFFNo
			cSP = "{call billingReportAFF.ProcessCallInfo('"
		'KFAB-6227
		case AFFMNo
			cSP = "{call billingreportAFFM.ProcessCallInfo('" & cAHS & "', '"
	    case WMANo ' WASTE MANAGER GBS
	        cAHS = "12159325"
	         cSP = "{call BILLINGREPORT_WMA.ProcessCallInfo('"
	    case  SHPRNo  ' "Sunstone Hotel Property" 14504629  GBS
	         cAHS = "14504629"
	         cSP = "{call BILLINGREPORT_SHPR.ProcessCallInfo('"
        case CVGNo 'COMMON OR VIRG
	       cSP = "{call BILLINGREPORT_COMOFVIRG.ProcessCallInfo_COMVIRG('" & cAHS & "','"
	    case MCDNo
			cSP = "{call billingReportMAC.ProcessCallInfo('"
		case WIGNo, MGCNoReg, MGCNoIRC 'MGC consts have the same values
			cSP = "{call billingReport_WIG_MGC.ProcessCallInfo('" & cAHS & "', '"
		case AIMNo
			cSP = "{call billingReport_AIM.ProcessCallInfo('" & cAHS & "', '"
		case NBICNo
			cSP = "{call billingReport_NBIC.ProcessCallInfo('" & cAHS & "', '"
		case ONBNo
		   cSP = "{call billingReport_ONB.ProcessCallInfo('" & cAHS & "','"
		case CSGNoCall ,CSGNoOnline
		    cSP = "{call billingReport_CISG.ProcessCallInfo('" & cAHS & "', '"
		case AIKNo
			 cSP = "{call billingReportAIK.ProcessCallInfo('"
	    case RDCNo
		    cSP = "{call billingReport_RDC.ProcessCallInfo('" & cAHS & "', '"
	    case CRWNoASP,CRWNoFNS
	        cSP = "{call BILLINGREPORT_CRAW.ProcessCallInfo('" & cAHS & "', '"
	    case FGNo
	       cSP = "{call BILLINGREPORT_FRG.ProcessCallInfo('" & cAHS & "', '"
	    case CSAANo
	       cSP = "{call BILLINGREPORT_CSAA.ProcessCallInfoCSAA('" & cAHS & "', '"
		case AAANo
	       cSP = "{call BILLINGREPORT_AAA.ProcessCallInfoAAA('" & cAHS & "', '"
	    case KMPNo
	      cSP = "{call BILLINGREPORT_KMP.ProcessCallInfo('" & cAHS & "', '"
	    case ARGNo
	    cSP = "{call BillingReportARG.ProcessCallInfo('"
	    case GBSNo 'GBS
			cSP = "{call billingReport_GBS.ProcessCallInfo('" & cAHS & "', '"
	    case SENNo 'SEN
			cSP = "{call billingReportSEN.ProcessCallInfo('" & cAHS & "', '"
		case CNLNo ' Cnl LPIC 0063 07th Nov
			'if cCustName="Canal3-in-1" then
				cSP = "{call billingReportCNL.ProcessCallInfo('" & cAHS & "', '"
			'else
			'	cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
			'end if
		case HMLNo ' HML 18 Oct
			cSP = "{call billingReport_HML.ProcessCallInfo('" & cAHS & "', '"
        'Added for MCAS-0479
        case RTWNo ' RTW 8 March 2006
			cSP = "{call billingReportRTW.ProcessCallInfo('" & cAHS & "', '"
		case AMCNo
			cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
		case ALMNo ' ALM 05 Nov
			cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
		case PMCONo
			cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
		'MROU-3021
		case TGCNo, TGCNoASP
			cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
		'For MROU-3087
		case EVRNo
 			cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
		'FOR MROU-3549
		case SAFNo
 			cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"

		case ACENo
			cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
		case ESISNo
			cSP = "{call billingReport_ESIS.ProcessCallInfo('" & cAHS & "', '"
		case AMENo
			cSP = "{call billingReport_AME.ProcessCallInfo('" & cAHS & "', '"
		case SEANo
			cSP = "{call billingReport_SEA.ProcessCallInfo('" & cAHS & "', '"
	    case else
			 IF cCCE = "CCE" THEN
				cSP = "{call billingReportCCE.ProcessCallInfo('"
			 ELSE
				cSP = "{call billingReport.ProcessCallInfo('" & cAHS & "', '"
			 END IF
	end select
	cSP = cSP & cStart & "','" & cEnd & "')}"

	Conn.Execute cSP


	select case Clng(cAHS)
		case FRENo, FMTNo	'	Fremont
			doFremont
			genFremontXLS
		case MCDNo
			lWithError = genMAC
		case CCENo
			lWithError = genCCE  '--11
		case AFFNo
			lWithError = genCCE_AFF  '--21675233
		case AFFMNo
			lWithError = genAFF  '--550
	    case CHBNo
			lWithError = genCHB
		case WIGNo
		    lWithError= genWIG	'--104
		case MGCNo
	        lWithError= genMGC '--103
		case ONBNo
	        lWithError= genONB '--20
		case CSGNoCall,CSGNoOnline
	        lWithError = genCSG
		case RDCNo
			lWithError = genRDC '-25
        case AIKNo
			lWithError = genAIK
		case WMANo
	      	lWithError = genWMA
		case CRWNoASP,CRWNoFNS
	      	lWithError = genCRW
		case  KMPNo
	     	lWithError = genKMP
	    case  CSAANo
	     	lWithError = genCSAA
		case  AAANo
	     	lWithError = genAAA
		case  CIRNo
			lWithError = genCIR

		case  SELNo
	     	lWithError = genSEL

        case  SRSNo
	     	lWithError = genSRS


        'Added for JPRI-0941

		case AMCNo
		lWithError = genAmcXLS
		case PMCONo
			lWithError = genXLS
		'MROU-3021
		case TGCNo
			lWithError = genXLS
		'TPAL-0146
		case TGCNoASP
			lWithError = genTOWASPXLS
		'added for MROU-3087
		case EVRNo
			lWithError = genXLS
		'added for MROU-3549
		case SAFNo
			lWithError = genXLS
		case ACENo
			lWithError = genACEXLS
		case ESISNo
			lWithError = genESISXLS
		case ARGNo
			lWithError = genARG
		case SENNo
			lWithError = genSEN
		case RTWNo
			lWithError = genRTWXLS

		case GBSNo
			lWithError = genGBSXLS

		case CNLNo 'LPIC-0063 07/11/2005
			'if cCustName="Canal3-in-1" then
					lWithError = genCnlXLS	'LPIC-0063
			'else
			'	    lithError = genCnlOldXLS
			'end if
		case HMLNo 'HML Implementation 24/10/2005
			lWithError = genHML
		case ALMNo
			lWithError = genALMXLS
		case UNINo
			lWithError = genUNIXLS
		case AMENo
			lWithError = genAMEXLS
		case else
			IF cCCE = "CCE" THEN
				lWithError = genCCE
			ELSE
				lWithError = genXLS				
			END IF

	end select
	'
	if not lWithError then
		'	update history table
		cPer=MonthName(month(CDate(Request.QueryString("DATEFROM"))),true) & CStr(Year(Request.QueryString("DATEFROM")))
		cSQL = "INSERT INTO BILLING_HISTORY (MMM_YYYY,CREATED_BY,CREATED_ON,FILENAME,FILE_PATH,SERVER_NAME,AHS_ID) " & _
				"VALUES('" & cPer & "','" & _
				Session("NAME") & "','" & _
				now & "','" & _
				cTmpFile & "','" & _
				cDownloadLocation & "','" & _
				" '," & _
				cAHS & ")"
		with Conn
			.Execute(cSQL)
		end with
	end if
end if
Conn.Close
Set Conn = Nothing
Set oExcel = Nothing
Server.ScriptTimeout = nTimeOut
Application.Lock
Application("lExecutingBillingReport") = false
Application.UnLock

with response
	.Write "<script language=""JScript"">" & vbCRLF
	'.write "parent.document.location.href = ""default.htm""" & vbCRLF
	.Write "</script>" & vbCRLF
end with
'
'

' *********************************************************************
'	DELETE PREVIOUS REPORT
' *********************************************************************
sub deletePreviousReport
dim cSQL, oRS, nBillID, oRS1, oRS2, oRS3
cSQL = ""
cSQL = cSQL & " SELECT BILLING_ID From BILLING"
IF cCCE = "CCE" THEN
	cSQL = cSQL & " Where accnt_hrcy_step_id IN (" & cAHS & ",11)"
ELSE
	cSQL = cSQL & " Where accnt_hrcy_step_id IN (" & cAHS & ")"
END IF

set oRS = Conn.Execute(cSQL)
with oRS
	do while not .eof
		nBillID = CStr(.fields(0))
		cSQL = "DELETE From BILLING_DETAIL Where BILLING_ID IN (" & nBillID & ")"
		Conn.Execute cSQL
		cSQL = "DELETE From BILLING Where BILLING_ID IN (" & nBillID & ")"
		Conn.Execute cSQL
		.moveNext
	loop
	.close
end with
set oRS = nothing
end sub

'**********************************************************************

'*****************************GEN RTW**********************************

function genRTWXLS
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,cSQL12,oRS1,csubType,cCallerType,cSubSQL1,cFeeTypeID
dim dRepDate, cBillID, cCustName, cCustCode,cInputType
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee ,ccalltype
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serviceMail,cfaxvalue,cCall_Type,cMasterClient


cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
       if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
          cCustName = "Managed Care Innovations."
       else
          cCustName = Request.QueryString("CUSTNAME")
       end if

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"

'Response.write(cSQL)
'Response.end

set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingRTW.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Input_Type,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Master_Client,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		cfaxvalue= 0
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if 'end of is Null
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = 0 'CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)

                   'Get the Input type/caller type
				on error  resume next
				   cSQL1 =  " Select CALL_CALLER.CALLER_TYPE From CALL, CALL_CALLER " & _
		            		" WHERE STATUS = 'COMPLETED' " & _
							" AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
					        " AND CALL.CALL_ID=" & .Fields("Call_ID").Value
					cSQL1 = cSQL1 & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS
							set oRS1 = Conn.Execute(cSQL1)
					with oRS1
						 cCallerType = .fields(0)
						.close
					end with

               ' Response.write(cCallerType)
			   'Get the Master_Client Information
			    cMasterClient = ""
				cSQL1 =  "  select Name from call_account where  " & _
		            		" call_claim_id = (select call_claim_id from call_claim,call_caller  " & _
							" where call_caller.call_id =call_claim.call_id and " & _
							" call_caller.call_id =" & .Fields("Call_ID").Value & ")"

					set oRS1 = Conn.Execute(cSQL1)
					with oRS1
						 cMasterClient = Trim(.fields(0))
						.close
					end with

				'End

				if (isnull(cCallerType) or cCallerType="" ) then
				     'default value
                     cInputType = "C"
					 nTotSvcFee=0
                     'cCallerType= cInputType
			    else
                     cInputType=cCallerType
				end if  'isNull

				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					 cInputType & "','"  & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if 'is Null

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if 'if err.number

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if 'is null

					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if 'is null

					cValues = cValues & "','"
					'Master Client
					if isNull(cMasterClient) then
						cValues = cValues & ""
					else
						cValues = cValues & cMasterClient
					end if 'is null

					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value

					'Get the Service Fee Type

					if (cCallerType="F" or cCallerType="NET" or cCallerType="I" or cCallerType="A" or cCallerType="C" or cCallerType="E" or cCallerType="F" or cCallerType="M" or cCallerType="OTH" or cCallerType="AGT") Then
					     csubType = cCallerType

						 if ( cCallerType="NET" ) Then
						   csubType = "N"
					     end if

						if (cCallerType= "F" or cCallerType="M") Then
                              select case cCallerType
								  case "F"
								  cFeeTypeID = 2
								  case "M"
								  cFeeTypeID = 10
								  case else
								  cFeeTypeID = 4
							  end select
						else
                              cFeeTypeID = 1
                        end if

						if (cCallerType="E" or cCallerType="AGT" or cCallerType="OTH" or cCallerType="C") Then
                           csubType = "C"
						end if

						cSQL12 = " AND FEE_TYPE_ID=" & cFeeTypeID

                       if ( cCallerType="F" or cCallerType ="NET" or cCallerType="I"  or cCallerType="E" or cCallerType="AGT" or cCallerType="OTH" or cCallerType="C") then
						cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
								" WHERE ACCNT_HRCY_STEP_ID = " & cAHS  & _
								" AND CALL_TYPE='" & csubType & "'"

					   else
                        cSQL1 = cSQL1 & cSubSQL2
					   end if 'F/Net/OATH ...



					   if (cCallerType="M" ) then
                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND CALL_TYPE = 'M'"
					   end if 'M

						set oRS1 = Conn.Execute(cSQL1)
						with oRS1
							 cfaxvalue = CSng(.fields(0))
							.close
						end with

                    End if

          if (cCallerType <> "F" and cCallerType <> "NET" and cCallerType="I" and cCallerType <> "A" and cCallerType <>"C" and cCallerType <> "E" and cCallerType <>"F" and cCallerType <>"M" and cCallerType <>"AGT" and cCallerType <>"OTH") Then
						cfaxvalue = Csng(.Fields("SERVICE_FEE").Value)
          End if

                    'Check for informational call
					 ccalltype = .Fields("Call_Type").Value
                     if (ccalltype="I") then
                          cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
								" WHERE ACCNT_HRCY_STEP_ID = " & cAHS  & _
								" AND CALL_TYPE='" & ccalltype & "'"
						set oRS1 = Conn.Execute(cSQL1)
						with oRS1
							 cfaxvalue = CSng(.fields(0))
							.close
						end with
					 end if

					nTotSvcFee = nTotSvcFee + cfaxvalue
                    nTotalFee = nTotalFee + cfaxvalue + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee

					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(cfaxvalue) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(cfaxvalue + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if

		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			'cValues = "'','','','','','','','','','','','','','','','','','','','',''"
			'oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
on error resume next
genXLS = lErrorTriggered
end function

'-------------------------------------------------------------------------------------



'CHB
'----------------------------------------------------------------------------------------


function genCHB
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee,cSQL1,oRS1,cCallerType
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee,n
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls,serval
serval="0"

n=0


cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
       if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
          cCustName = "Managed Care Innovations."
       else
          cCustName = Request.QueryString("CUSTNAME")
       end if

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

'response.Write(cSQL)
'response.end
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingCHB.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Input_Type,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee

				cSQL1 =  " Select CALL_CALLER.CALLER_TYPE From CALL, CALL_CALLER " & _
		            		" WHERE STATUS = 'COMPLETED' " & _
							" AND CALL_CALLER.CALL_ID = CALL.CALL_ID + 0 " & _
					        " AND CALL.CALL_ID=" & .Fields("Call_ID").Value
					cSQL1 = cSQL1 & " AND CALL.CLIENT_HRCY_STEP_ID = " & cAHS
							set oRS1 = Conn.Execute(cSQL1)
					with oRS1
						 cCallerType = .fields(0)
						.close
					end with




					if (.Fields("Call_Type").Value="M" ) then
                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND CALL_TYPE = 'M'"
					   end if 'M

			       set oRS1 = Conn.Execute(cSQL1)
						with oRS1
							 serval = .fields(0)
							.close
						end with



				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					 cCallerType & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if .Fields("Call_Type").Value="C" then
				        n=n+1
					end if


					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"

					if (n<5000 ) and (.Fields("Call_Type").Value <> "M")  then
				     serval= cSng(FormatNumber(.Fields("SERVICE_FEE").Value))
				    end if
				    if ( n>=5000 and n<10000 ) and (.Fields("Call_Type").Value <> "M") then
				      serval="16.50"
				    end if
				    if (n>10000) and (.Fields("Call_Type").Value <> "M")  then
				      serval="14.50"
				    end if

				    if (isNull(serval)) then
				     serval ="0"
				    end if


				    cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(cSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genXLS = lErrorTriggered
end function


function genSEL
	dim cAHS, cStartDate, cSP, oRS, cSQL ,cSQL1 ,oRS1 , oRS2
	dim dRepDate, cBillID, cCustName, cCustCode
	dim cFields, cValues, cTime
	dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
	dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee,nTransmissionfees
	dim cCmpParent
	dim lErrorTriggered
	dim nGrandTotal
	dim nTotalNoCalls
	dim strTempLOB,strBillingDetailsID,strLocation,serval,serval1
	serval="0"
	serval1="0"

	cAHS = Request.QueryString("AHS")
	cStartDate = Request.QueryString("DATEFROM")
	cCustCode = Request.QueryString("CUSTCODE")
	cCustName = Request.QueryString("CUSTNAME")

	lErrorTriggered = false
	nGrandTotal = 0
	nTotalNoCalls = 0
	'
		if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
			cCustName = "Managed Care Innovations."
		else
			cCustName = Request.QueryString("CUSTNAME")
		end if

	dRepDate = cDate(cStartDate)
	cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
	set oRS = Conn.Execute(cSQL)
	with oRS
		cBillID = .fields(0)
		.close
	end with
	cTime = CStr(FormatDateTime( Time, vbShortTime))
	cTime = Replace(cTime, ":", "")
	cTmpFile = cCustCode & "-" & cTime & ".xls"
	'
		
	cSQL = "SELECT BD.*,S.NAME " & _
				"FROM BILLING_DETAIL BD " & _
				"INNER JOIN SITE S on S.SITE_ID = BD.SITE_ID  " & _				
				"WHERE (BD.BILLING_ID = " & cBillID & _
				" AND BD.STATUS='ACTIVE') " & _
				" AND BD.CALLSTATUS IN ('COMPLETED', 'PENDED', 'PEND-RESOLVED', 'PEND-ABORTED') " & _
				"Order by BD.CLIENT_NAME,BD.PARENT_NAME, BD.CALL_TYPE, BD.LOB_CD, BD.CALL_END_TIME"
				'response.write(cSQL)
				'response.end
	set oRS = Conn.Execute(cSQL)
	'
	
	cTime = CStr(FormatDateTime( Time, vbShortTime))
	cTime = Replace(cTime, ":", "")
	cTmpFile = cCustCode & "-" & cTime & ".xls"

	with oExcel
		.cDestinationFileName = cTmpFile
		.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingSEL.xls"
		.cExcelRangeName = "ODBCRange"
		.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
		'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
		cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
		.cDownloadLocation = cDownloadLocation
		.openXLS
		.writeMsg "Generating spreadsheet"
		.writeCell "Account", cCustName
		.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
		'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
	end with
	writePeriod(oExcel)
	cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Loss_State,Caller_Name,Employee_Name,Employee_SSN,Site_Id,Site_Name,Service,Total_Fax,Temp,Escalate,Vendor,Print,Total"
	with oRS
		Do While Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cParent = ""
			else
				cParent = .Fields("PARENT_NAME").Value
			end if
			nCalls = 0
			nTotSvcFee = 0
			nTotalFee = 0
			do while Not .EOF
				if isNull(.Fields("PARENT_NAME").Value) then
					cCmpParent = ""
				else
					cCmpParent = .Fields("PARENT_NAME").Value
				end if

				 'response.Write(.Fields("Call_Type").Value)

				'if (trim(.Fields("Call_Type").Value)="F" ) then

				if cParent = cCmpParent then
					nCalls = nCalls + 1
					cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND CALL_TYPE = 'F'" &_
							  " AND FEE_TYPE_ID='2'"

					'response.end
					  ' end if
					'response.Write(.Fields("Call_Type").Value)
					'response.end
					if (trim(.Fields("Call_Type").Value)="F" ) then
			          set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						do while Not .EOF
							 serval = .Fields("FEE_AMOUNT").Value
						oRS1.moveNext
					    loop
							oRS1.close
						end with
						nTotalFaxFee = CSng(serval)
						nTotSvcFee=nTotSvcFee+CSng(serval)
						else
						nTotalFaxFee =  CSng(.Fields("TOTAL_FAX_FEE").Value)
						nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
						'response.Write(.Fields("SERVICE_FEE").Value)
						'response.end
					end if

					nTempFee = CSng(.Fields("TEMP_FEE").Value)
					nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
					nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
					if isNull(.Fields("DRIVEN_FEE").Value) then
					nPrintFee=0
					else
					nPrintFee=Csng(.Fields("DRIVEN_FEE").Value)
					end if
					nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrintFee
					cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
						CStr(.Fields("Call_ID").Value) & ",'" & _
						.Fields("Call_Type").Value & "','" & _
						.Fields("CALLSTATUS").Value & "','"
						'on error resume next
						if isNull(.Fields("LOSS_DATE").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("LOSS_DATE").Value
						end if

						if err.number <> 0 then
							if err.number = -2147217887 then
								writeError .Fields("Call_ID").Value
							end if
							lErrorTriggered = true
							exit do
						end if

						cValues = cValues & "','" & _
						CStr(.Fields("CALL_END_TIME").Value) & "','"
						if isNull(.Fields("CLAIM_NUMBER").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("CLAIM_NUMBER").Value
						end if
						cValues = cValues & "','" & _
						.Fields("LOB_CD").Value & "','"
						if isNull(.Fields("POLICY_NUMBER").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("POLICY_NUMBER").Value
						end if
						cValues = cValues & "','" & _
						removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
						
												
						'*****************************
						' Modified on 25/10/2005

						strTempLOB=.Fields("LOB_CD").Value
						strBillingDetailsID=.Fields("BILLING_DETAIL_ID").Value
						if(strTempLOB="WOR")then
							cSQL = "SELECT BENEFIT_STATE FROM CALL_CLAIM CLM,BILLING_DETAIL BDL"
							cSQL = cSQL & " WHERE BDL.CALL_ID= CLM.CALL_ID"
							cSQL = cSQL & " AND BDL.CLIENT_NODE_ID=650"
							cSQL = cSQL & " AND BDL.BILLING_DETAIL_ID=" & strBillingDetailsID
						end if
						if(strTempLOB="CAU" or strTempLOB="CPR" or strTempLOB="CLI" or strTempLOB="PAU" ) then

							cSQL = "SELECT ADDRESS_STATE FROM CALL_LOSS_LOCATION cll ,CALL_CLAIM cml ,BILLING_DETAIL          bdl"
							cSQL = cSQL & " WHERE cll.CALL_CLAIM_ID=cml.CALL_CLAIM_ID"
							cSQL = cSQL & " AND cml.CALL_ID = bdl.CALL_ID"
							cSQL = cSQL & " AND bdl.CLIENT_NODE_ID = 650"
							cSQL = cSQL & " AND bdl.BILLING_DETAIL_ID =" & strBillingDetailsID
						end if

						set oRS2 = Conn.Execute(cSQL)

						with oRS2
						do while Not .EOF
							 strLocation = .fields(0)
						oRS2.moveNext
					    loop
							oRS2.close
						end with


						cValues = cValues & strLocation & "','"
						'******************************
						redim aNameParts(1)
						aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
						aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
						cValues = cValues & getName(aNameParts) & "','"
						redim aNameParts(1)
						aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
						aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
						cValues = cValues & getName(aNameParts) & "','"
						cValues = cValues & cEmployeeSSN & "','"
						
						cValues = cValues & oRS.Fields("SITE_ID").Value & "','"
						cValues = cValues & oRS.Fields("NAME").Value & "','"
																		

						if (trim(.Fields("Call_Type").Value)="F" ) then
						cValues = cValues & FormatNumber(CSng(serval)) & "','"
						'serval1=serval
						else
						cValues = cValues & FormatNumber(.Fields("SERVICE_FEE").Value) & "','"
						'serval=.Fields("SERVICE_FEE").Value
						'FormatNumber(.Fields("SERVICE_FEE").Value) & "','" &
						end if

						cValues = cValues & FormatNumber(nTotalFaxFee) & "','" & _
						FormatNumber(nTempFee) & "','" & _
						FormatNumber(nEscalateFee) & "','" & _
						FormatNumber(nVendorFee) & "','" & _
						FormatNumber(nPrintFee) & "','"  &_
						FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrintFee) & "'"

					oExcel.addRow cFields, cValues
					.MoveNext
				else
					exit do
				end if
			loop
			if not lErrorTriggered then
				cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','','" & FormatCurrency(nTotalFee) & "'"
				oExcel.addRow cFields, cValues
				cValues = "'','','','','','','','','','','','','','','','','','','','','','',''"
				oExcel.addRow cFields, cValues
			else
				exit do
			end if
			nGrandTotal = nGrandTotal + nTotalFee
			nTotalNoCalls = nTotalNoCalls + nCalls
		Loop
		.Close
	end with
	cValues = "'','','','','','','','','','','','','','','','','','','','','','',''"
	oExcel.addRow cFields, cValues
	cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
	oExcel.addRow cFields, cValues
	oExcel.closeXLS
	if not lErrorTriggered then
		oExcel.sendFile
	end if
	Set oRS = Nothing
	Set oRS1 = Nothing
	Set oRS2 = Nothing
	genSEL = lErrorTriggered
end function



'SRS------------------------------


function genSRS
	dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
	dim dRepDate, cBillID, cCustName, cCustCode
	dim cFields, cValues, cTime
	dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
	dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee,nTransmissionfees
	dim cCmpParent
	dim lErrorTriggered
	dim nGrandTotal
	dim nTotalNoCalls
	dim strTempLOB,strBillingDetailsID,strLocation,serval,serval1
	serval="0"
	serval1="0"

	cAHS = Request.QueryString("AHS")
	cStartDate = Request.QueryString("DATEFROM")
	cCustCode = Request.QueryString("CUSTCODE")
	cCustName = Request.QueryString("CUSTNAME")

	lErrorTriggered = false
	nGrandTotal = 0
	nTotalNoCalls = 0
	'
		if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
			cCustName = "Managed Care Innovations."
		else
			cCustName = Request.QueryString("CUSTNAME")
		end if

	dRepDate = cDate(cStartDate)
	cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
	set oRS = Conn.Execute(cSQL)
	with oRS
		cBillID = .fields(0)
		.close
	end with
	cTime = CStr(FormatDateTime( Time, vbShortTime))
	cTime = Replace(cTime, ":", "")
	cTmpFile = cCustCode & "-" & cTime & ".xls"
	'
	cSQL = "SELECT BILLING_DETAIL.* " & _
				"FROM BILLING_DETAIL " & _
				"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
				" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
				" AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED' " & _
				"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
				'response.write(cSQL)
				'response.end
	set oRS = Conn.Execute(cSQL)
	'
	cTime = CStr(FormatDateTime( Time, vbShortTime))
	cTime = Replace(cTime, ":", "")
	cTmpFile = cCustCode & "-" & cTime & ".xls"

	with oExcel
		.cDestinationFileName = cTmpFile
		.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingSRS.xls"
		.cExcelRangeName = "ODBCRange"
		.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
		'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
		cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
		.cDownloadLocation = cDownloadLocation
		.openXLS
		.writeMsg "Generating spreadsheet"
		.writeCell "Account", cCustName
		.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
		'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
	end with
	writePeriod(oExcel)
	cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Loss_State,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Print,Total"
	with oRS
		Do While Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cParent = ""
			else
				cParent = .Fields("PARENT_NAME").Value
			end if
			nCalls = 0
			nTotSvcFee = 0
			nTotalFee = 0
			do while Not .EOF
				if isNull(.Fields("PARENT_NAME").Value) then
					cCmpParent = ""
				else
					cCmpParent = .Fields("PARENT_NAME").Value
				end if

				 'response.Write(.Fields("Call_Type").Value)

				'if (trim(.Fields("Call_Type").Value)="F" ) then

				if cParent = cCmpParent then
					nCalls = nCalls + 1
					cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND CALL_TYPE = 'F'" &_
							  " AND FEE_TYPE_ID='2'"

					'response.end
					  ' end if
					'response.Write(.Fields("Call_Type").Value)
					'response.end
					if (trim(.Fields("Call_Type").Value)="F" ) then
			          set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						do while Not .EOF
							 serval = .Fields("FEE_AMOUNT").Value
						oRS1.moveNext
					    loop
							oRS1.close
						end with
						nTotalFaxFee = CSng(serval)
						nTotSvcFee=nTotSvcFee+CSng(serval)
						else
						nTotalFaxFee =  CSng(.Fields("TOTAL_FAX_FEE").Value)
						nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
						'response.Write(.Fields("SERVICE_FEE").Value)
						'response.end
					end if

					nTempFee = CSng(.Fields("TEMP_FEE").Value)
					nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
					nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
					if isNull(.Fields("DRIVEN_FEE").Value) then
					nPrintFee=0
					else
					nPrintFee=Csng(.Fields("DRIVEN_FEE").Value)
					end if
					nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrintFee
					cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
						CStr(.Fields("Call_ID").Value) & ",'" & _
						.Fields("Call_Type").Value & "','" & _
						.Fields("CALLSTATUS").Value & "','"
						'on error resume next
						if isNull(.Fields("LOSS_DATE").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("LOSS_DATE").Value
						end if

						if err.number <> 0 then
							if err.number = -2147217887 then
								writeError .Fields("Call_ID").Value
							end if
							lErrorTriggered = true
							exit do
						end if

						cValues = cValues & "','" & _
						CStr(.Fields("CALL_END_TIME").Value) & "','"
						if isNull(.Fields("CLAIM_NUMBER").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("CLAIM_NUMBER").Value
						end if
						cValues = cValues & "','" & _
						.Fields("LOB_CD").Value & "','"
						if isNull(.Fields("POLICY_NUMBER").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("POLICY_NUMBER").Value
						end if
						cValues = cValues & "','" & _
						removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"

						'*****************************
						' Modified on 25/10/2005

						strTempLOB=.Fields("LOB_CD").Value
						strBillingDetailsID=.Fields("BILLING_DETAIL_ID").Value
						if(strTempLOB="WOR")then
							cSQL = "SELECT BENEFIT_STATE FROM CALL_CLAIM CLM,BILLING_DETAIL BDL"
							cSQL = cSQL & " WHERE BDL.CALL_ID= CLM.CALL_ID"
							cSQL = cSQL & " AND BDL.CLIENT_NODE_ID=600"
							cSQL = cSQL & " AND BDL.BILLING_DETAIL_ID=" & strBillingDetailsID
						end if
						if(strTempLOB="CAU" or strTempLOB="CPR" or strTempLOB="CLI" or strTempLOB="PAU" ) then

							cSQL = "SELECT ADDRESS_STATE FROM CALL_LOSS_LOCATION cll ,CALL_CLAIM cml ,BILLING_DETAIL          bdl"
							cSQL = cSQL & " WHERE cll.CALL_CLAIM_ID=cml.CALL_CLAIM_ID"
							cSQL = cSQL & " AND cml.CALL_ID = bdl.CALL_ID"
							cSQL = cSQL & " AND bdl.CLIENT_NODE_ID = 650"
							cSQL = cSQL & " AND bdl.BILLING_DETAIL_ID =" & strBillingDetailsID
						end if

						set oRS = Conn.Execute(cSQL)

						with oRS
						do while Not .EOF
							 strLocation = .fields(0)
						oRS.moveNext
					    loop
							oRS.close
						end with


						cValues = cValues & strLocation & "','"
						'******************************
						redim aNameParts(1)
						aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
						aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
						cValues = cValues & getName(aNameParts) & "','"
						redim aNameParts(1)
						aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
						aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
						cValues = cValues & getName(aNameParts) & "','"
						cValues = cValues & cEmployeeSSN & "','"

						if (trim(.Fields("Call_Type").Value)="F" ) then
						cValues = cValues & FormatNumber(CSng(serval)) & "','"
						'serval1=serval
						else
						cValues = cValues & FormatNumber(.Fields("SERVICE_FEE").Value) & "','"
						'serval=.Fields("SERVICE_FEE").Value
						'FormatNumber(.Fields("SERVICE_FEE").Value) & "','" &
						end if

						cValues = cValues & FormatNumber(nTotalFaxFee) & "','" & _
						FormatNumber(nTempFee) & "','" & _
						FormatNumber(nEscalateFee) & "','" & _
						FormatNumber(nVendorFee) & "','" & _
						FormatNumber(nPrintFee) & "','"  &_
						FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrintFee) & "'"

					oExcel.addRow cFields, cValues
					.MoveNext
				else
					exit do
				end if
			loop
			if not lErrorTriggered then
				cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','','" & FormatCurrency(nTotalFee) & "'"
				oExcel.addRow cFields, cValues
				cValues = "'','','','','','','','','','','','','','','','','','','','',''"
				oExcel.addRow cFields, cValues
			else
				exit do
			end if
			nGrandTotal = nGrandTotal + nTotalFee
			nTotalNoCalls = nTotalNoCalls + nCalls
		Loop
		.Close
	end with
	cValues = "'','','','','','','','','','','','','','','','','','','','',''"
	oExcel.addRow cFields, cValues
	cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
	oExcel.addRow cFields, cValues
	oExcel.closeXLS
	if not lErrorTriggered then
		oExcel.sendFile
	end if
	Set oRS = Nothing
	genSRS = lErrorTriggered
end function


'-------------------------------------

' *********************************************************************
'	CCE
' *********************************************************************

function genCCE
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim cBranchName, cCmpBranch
dim nCallsByBranch, nTotalByBranch
dim nFeeSum


cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0


'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=11"
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

'/*--------------------------------------------------------------------------------------------------------------------*/
'//BEGIN		{JPRI-0828}	Altered SQL Statement to fetch CALL_CALLER.CALLER_TYPE
'/*--------------------------------------------------------------------------------------------------------------------*/
cSQL = ""
cSQL = cSQL & " SELECT BILLING_DETAIL.*,CALL_CALLER.CALLER_TYPE"
cSQL = cSQL & " FROM BILLING_DETAIL,CALL_CALLER "
cSQL = cSQL & " WHERE BILLING_DETAIL.CALL_ID = CALL_CALLER.CALL_ID"
cSQL = cSQL & " AND (BILLING_DETAIL.BILLING_ID = " & cBillID

cSQL = cSQL & " AND BILLING_DETAIL.STATUS='ACTIVE') "
IF cAHS <> 11 THEN
	cSQL = cSQL & " AND PARENT_NODE_ID IN ( " & cAHS & ")"
END IF
cSQL = cSQL & " Order by PARENT_NAME, BRANCH_NAME, CLIENT_NAME, CALL_TYPE"
'/*--------------------------------------------------------------------------------------------------------------------*/
'//END			{JPRI-0828}	Altered SQL Statement to fetch CALL_CALLER.CALLER_TYPE
'/*--------------------------------------------------------------------------------------------------------------------*/

set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingCCE.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)

'/*--------------------------------------------------------------------------------------------------------------------*/
'//BEGIN		{JPRI-0828}	Altered cFields to accomodate value for CALL_CALLER.CALLER_TYPE at the end column
'/*--------------------------------------------------------------------------------------------------------------------*/
cFields = "Branch,Account,Call_No,T,Status,Caller_Type,Loss_Dt,Call_Dt,Claim_No,LOB," & _
		  "Policy_No,Risk_Location,Insured_State,Insured,Caller_Name,Employee_Name," & _
		  "Service,Total_Fax,Temp,Escalate,Vendor,Total"
'/*--------------------------------------------------------------------------------------------------------------------*/
'//END			{JPRI-0828}	Altered cFields to accomodate value for CALL_CALLER.CALLER_TYPE at the end column
'/*--------------------------------------------------------------------------------------------------------------------*/

with oRS
	Do While Not .EOF
		if lErrorTriggered then
			exit do
		end if
		cParent = .Fields("PARENT_NAME").Value
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while not .eof
			if lErrorTriggered then
				exit do
			end if
			if cParent = .Fields("PARENT_NAME").Value then
				' brake by branch
				if isNull(.Fields("BRANCH_NAME").Value) then
					cBranchName = ""
				else
					cBranchName = .Fields("BRANCH_NAME").Value
				end if
				nCallsByBranch = 0
				nTotalByBranch = 0
				do while not .eof
					if cParent = .Fields("PARENT_NAME").Value then
						if isNull(.Fields("BRANCH_NAME").Value) then
							cCmpBranch = ""
						else
							cCmpBranch = .Fields("BRANCH_NAME").Value
						end	if

						if cBranchName = cCmpBranch then
							nCallsByBranch = nCallsByBranch + 1
							nCalls = nCalls + 1
							nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
							nTempFee = CSng(.Fields("TEMP_FEE").Value)
							nEscalateFee = CSng(.Fields("ESCALATE_FEE").Value)
							nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
							nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
							nFeeSum = CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
							nTotalByBranch = nTotalByBranch + nFeeSum
							nTotalFee = nTotalFee + nFeeSum
							cValues =	"'" & .Fields("BRANCH_NAME").Value & "','" & _
								removeSngQuote(.Fields("PARENT_NAME").Value) & "','" & _
								CStr(.Fields("Call_ID").Value) & "','" & _
								.Fields("Call_Type").Value & "','" & _
								.Fields("CALLSTATUS").Value & "','" &_
							   	removeSngQuote(.Fields("Caller_Type").Value) & "','"

							on error resume next
							if isNull(.Fields("LOSS_DATE").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("LOSS_DATE").Value
							end if

							if err.number <> 0 then
								if err.number = -2147217887 then
									writeError .Fields("Call_ID").Value
								end if
								lErrorTriggered = true
								exit do
							end if

							cValues = cValues & "','" & _
								CStr(.Fields("CALL_END_TIME").Value) & "','"
							if isNull(.Fields("CLAIM_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("CLAIM_NUMBER").Value
							end if
							cValues = cValues & "','" & _
								.Fields("LOB_CD").Value & "','"
							if isNull(.Fields("POLICY_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & removeSngQuote(.Fields("POLICY_NUMBER").Value)
							end if
							cValues = cValues & "','" & _
								removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
							cValues = cValues & .Fields("INSURED_STATE").Value & "','"
							cValues = cValues & removeSngQuote(.Fields("INSURED_NAME").Value) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
							aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
							aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','" & _
								FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
								FormatNumber(nTotalFaxFee) & "','" & _
								FormatNumber(nTempFee) & "','" & _
								FormatNumber(nEscalateFee) & "','" & _
								FormatNumber(nVendorFee) & "','" & _
								FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"


								oExcel.addRow cFields, cValues

							if err.number <> 0 then
								writeErrText
								lErrorTriggered = true
								exit do
							end if
							on error goto 0
							.MoveNext
						else
							if not lErrorTriggered then
								cValues = "'Branch SubTotal',''," & nCallsByBranch & ",'','','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotalByBranch) & "'"
							    oExcel.addRow cFields, cValues
								cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
								oExcel.addRow cFields, cValues
							end if
							exit do
						end if
					else
						exit do
					end if
				loop
			else
				if not lErrorTriggered then
					cValues = "'Account SubTotal',''," & nCalls & ",'','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
					oExcel.addRow cFields, cValues
					cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
					oExcel.addRow cFields, cValues
				end if
				exit do
			end if
		loop
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with

cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genCCE = lErrorTriggered
end function

' *********************************************************************
'	CCE Affirmative >>accnt_hrcy_step_id  = 21675233
' *********************************************************************

function genCCE_AFF
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim cBranchName, cCmpBranch
dim nCallsByBranch, nTotalByBranch
dim nFeeSum


cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0


'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id = " & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

'/*--------------------------------------------------------------------------------------------------------------------*/
'//BEGIN		{JPRI-0828}	Altered SQL Statement to fetch CALL_CALLER.CALLER_TYPE
'/*--------------------------------------------------------------------------------------------------------------------*/
cSQL = ""
cSQL = cSQL & " SELECT BILLING_DETAIL.*,CALL_CALLER.CALLER_TYPE"
cSQL = cSQL & " FROM BILLING_DETAIL,CALL_CALLER "
cSQL = cSQL & " WHERE BILLING_DETAIL.CALL_ID = CALL_CALLER.CALL_ID"
cSQL = cSQL & " AND (BILLING_DETAIL.BILLING_ID = " & cBillID

cSQL = cSQL & " AND BILLING_DETAIL.STATUS='ACTIVE') "
IF cAHS <> 11 THEN
	cSQL = cSQL & " AND CLIENT_NODE_ID IN ( " & cAHS & ")"
END IF
cSQL = cSQL & " Order by PARENT_NAME, BRANCH_NAME, CLIENT_NAME, CALL_TYPE"
'/*--------------------------------------------------------------------------------------------------------------------*/
'//END			{JPRI-0828}	Altered SQL Statement to fetch CALL_CALLER.CALLER_TYPE
'/*--------------------------------------------------------------------------------------------------------------------*/

set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingCCE_AFF.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)

'/*--------------------------------------------------------------------------------------------------------------------*/
'//BEGIN		{JPRI-0828}	Altered cFields to accomodate value for CALL_CALLER.CALLER_TYPE at the end column
'/*--------------------------------------------------------------------------------------------------------------------*/
cFields = "Branch,Account,Call_No,Rep_Pro_Claim_Trained,Entered_During_Business_Hours,T,Status,Caller_Type,Loss_Dt," & _
		  "Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Insured_State,Insured,Caller_Name,Employee_Name," & _
		  "Service,Total_Fax,Temp,Escalate,Vendor,Total"
'/*--------------------------------------------------------------------------------------------------------------------*/
'//END			{JPRI-0828}	Altered cFields to accomodate value for CALL_CALLER.CALLER_TYPE at the end column
'/*--------------------------------------------------------------------------------------------------------------------*/

with oRS
	Do While Not .EOF
		if lErrorTriggered then
			exit do
		end if
		cParent = .Fields("PARENT_NAME").Value
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while not .eof
			if lErrorTriggered then
				exit do
			end if
			if cParent = .Fields("PARENT_NAME").Value then
				' brake by branch
				if isNull(.Fields("BRANCH_NAME").Value) then
					cBranchName = ""
				else
					cBranchName = .Fields("BRANCH_NAME").Value
				end if
				nCallsByBranch = 0
				nTotalByBranch = 0
				do while not .eof
					if cParent = .Fields("PARENT_NAME").Value then
						if isNull(.Fields("BRANCH_NAME").Value) then
							cCmpBranch = ""
						else
							cCmpBranch = .Fields("BRANCH_NAME").Value
						end	if

						if cBranchName = cCmpBranch then
							nCallsByBranch = nCallsByBranch + 1
							nCalls = nCalls + 1
							nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
							nTempFee = CSng(.Fields("TEMP_FEE").Value)
							nEscalateFee = CSng(.Fields("ESCALATE_FEE").Value)
							nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
							nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
							nFeeSum = CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
							nTotalByBranch = nTotalByBranch + nFeeSum
							nTotalFee = nTotalFee + nFeeSum
							cValues =	"'" & .Fields("BRANCH_NAME").Value & "','" & _
								removeSngQuote(.Fields("PARENT_NAME").Value) & "','" & _
								CStr(.Fields("Call_ID").Value) & "','" & _
								  .Fields("TRAINED_FOR_PRO_CLAIM").Value & "','" & _
								  .Fields("FILING_BTWN_HRS_ABOVE").Value & "','" & _
								.Fields("Call_Type").Value & "','" & _
								.Fields("CALLSTATUS").Value & "','" &_

							   	removeSngQuote(.Fields("Caller_Type").Value) & "','"

							on error resume next
							if isNull(.Fields("LOSS_DATE").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("LOSS_DATE").Value
							end if

							if err.number <> 0 then
								if err.number = -2147217887 then
									writeError .Fields("Call_ID").Value
								end if
								lErrorTriggered = true
								exit do
							end if

							cValues = cValues & "','" & _
								CStr(.Fields("CALL_END_TIME").Value) & "','"
							if isNull(.Fields("CLAIM_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("CLAIM_NUMBER").Value
							end if
							cValues = cValues & "','" & _
								.Fields("LOB_CD").Value & "','"
							if isNull(.Fields("POLICY_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & removeSngQuote(.Fields("POLICY_NUMBER").Value)
							end if
							cValues = cValues & "','" & _
								removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
							cValues = cValues & .Fields("INSURED_STATE").Value & "','"
							cValues = cValues & removeSngQuote(.Fields("INSURED_NAME").Value) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
							aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
							aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','" & _
								FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
								FormatNumber(nTotalFaxFee) & "','" & _
								FormatNumber(nTempFee) & "','" & _
								FormatNumber(nEscalateFee) & "','" & _
								FormatNumber(nVendorFee) & "','" & _
								FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"


								oExcel.addRow cFields, cValues

							if err.number <> 0 then
								writeErrText
								lErrorTriggered = true
								exit do
							end if
							on error goto 0
							.MoveNext
						else
							if not lErrorTriggered then
								cValues = "'Branch SubTotal',''," & nCallsByBranch & ",'','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotalByBranch) & "'"
							    oExcel.addRow cFields, cValues
								cValues = "'','','','','','','','','','','','','','','','','','','','','','','',''"
								oExcel.addRow cFields, cValues
							end if
							exit do
						end if
					else
						exit do
					end if
				loop
			else
				if not lErrorTriggered then
					cValues = "'Account SubTotal',''," & nCalls & ",'','','','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
					oExcel.addRow cFields, cValues
					cValues = "'','','','','','','','','','','','','','','','','','','','','','','',''"
					oExcel.addRow cFields, cValues
				end if
				exit do
			end if
		loop
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with

cValues = "'','','','','','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genCCE_AFF = lErrorTriggered
end function



' *********************************************************************
'	Affirmative >>accnt_hrcy_step_id  = 550 (KFAB-6227)
' *********************************************************************

function genAFF
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim cBranchName, cCmpBranch
dim nCallsByBranch, nTotalByBranch
dim nFeeSum


cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id = " & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with

cSQL = ""
cSQL = cSQL & " SELECT BILLING_DETAIL.*,CALL_CLAIM.CLAIM_TYPE_DESC, "
cSQL = cSQL & " CALL_CLAIM.TEMPEDPOLICY_FLG "
cSQL = cSQL & " FROM BILLING_DETAIL,CALL_CLAIM "
cSQL = cSQL & " WHERE BILLING_DETAIL.CALL_ID = CALL_CLAIM.CALL_ID"
cSQL = cSQL & " AND (BILLING_DETAIL.BILLING_ID = " & cBillID
cSQL = cSQL & " AND CALLSTATUS = 'COMPLETED' "
cSQL = cSQL & " AND LOB_CD = 'PAU' "
cSQL = cSQL & " AND BILLING_DETAIL.STATUS='ACTIVE') "
cSQL = cSQL & " Order by PARENT_NAME, BRANCH_NAME, CLIENT_NAME, CALL_TYPE"

set oRS = Conn.Execute(cSQL)

cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingAFF.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
end with
writePeriod(oExcel)

'/*--------------------------------------------------------------------------------------------------------------------*/
'//BEGIN		{KFAB-6227}	Altered cFields to accomodate value for CALL_CLAIM.TEMPEDPOLICY_FLG after the STATUS column
'/*--------------------------------------------------------------------------------------------------------------------*/
cFields = "Branch,Account,Call_No,T,Status,Temped_Call,Caller_Type,Loss_Dt," & _
		  "Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Insured_State,Insured,Caller_Name,Employee_Name," & _
		  "Service,Total_Fax,Temp,Escalate,Vendor,Total"
'/*--------------------------------------------------------------------------------------------------------------------*/
'//END			{KFAB-6227}	Altered cFields to accomodate value for CALL_CLAIM.TEMPEDPOLICY_FLG after the STATUS column
'/*--------------------------------------------------------------------------------------------------------------------*/

with oRS
	Do While Not .EOF
		if lErrorTriggered then
			exit do
		end if
		cParent = .Fields("PARENT_NAME").Value
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while not .eof
			if lErrorTriggered then
				exit do
			end if
			if cParent = .Fields("PARENT_NAME").Value then
				' brake by branch
				if isNull(.Fields("BRANCH_NAME").Value) then
					cBranchName = ""
				else
					cBranchName = .Fields("BRANCH_NAME").Value
				end if
				nCallsByBranch = 0
				nTotalByBranch = 0
				do while not .eof
					if cParent = .Fields("PARENT_NAME").Value then
						if isNull(.Fields("BRANCH_NAME").Value) then
							cCmpBranch = ""
						else
							cCmpBranch = .Fields("BRANCH_NAME").Value
						end	if

						if cBranchName = cCmpBranch then
							nCallsByBranch = nCallsByBranch + 1
							nCalls = nCalls + 1
							nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
							nTempFee = CSng(.Fields("TEMP_FEE").Value)
							nEscalateFee = CSng(.Fields("ESCALATE_FEE").Value)
							nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
							nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
							nFeeSum = CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
							nTotalByBranch = nTotalByBranch + nFeeSum
							nTotalFee = nTotalFee + nFeeSum
							cValues =	"'" & .Fields("BRANCH_NAME").Value & "','" & _
								removeSngQuote(.Fields("PARENT_NAME").Value) & "','" & _
								CStr(.Fields("Call_ID").Value) & "','" & _
								 .Fields("Call_Type").Value & "','" & _
								.Fields("CALLSTATUS").Value & "','" &_
								.Fields("TEMPEDPOLICY_FLG").Value & "','" &_
							   	removeSngQuote(.Fields("CLAIM_TYPE_DESC").Value) & "','"

							on error resume next
							if isNull(.Fields("LOSS_DATE").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("LOSS_DATE").Value
							end if

							if err.number <> 0 then
								if err.number = -2147217887 then
									writeError .Fields("Call_ID").Value
								end if
								lErrorTriggered = true
								exit do
							end if

							cValues = cValues & "','" & _
								CStr(.Fields("CALL_END_TIME").Value) & "','"
							if isNull(.Fields("CLAIM_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("CLAIM_NUMBER").Value
							end if
							cValues = cValues & "','" & _
								.Fields("LOB_CD").Value & "','"
							if isNull(.Fields("POLICY_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & removeSngQuote(.Fields("POLICY_NUMBER").Value)
							end if
							cValues = cValues & "','" & _
								removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
							cValues = cValues & .Fields("INSURED_STATE").Value & "','"
							cValues = cValues & removeSngQuote(.Fields("INSURED_NAME").Value) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
							aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
							aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','" & _
								FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
								FormatNumber(nTotalFaxFee) & "','" & _
								FormatNumber(nTempFee) & "','" & _
								FormatNumber(nEscalateFee) & "','" & _
								FormatNumber(nVendorFee) & "','" & _
								FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"


								oExcel.addRow cFields, cValues

							if err.number <> 0 then
								writeErrText
								lErrorTriggered = true
								exit do
							end if
							on error goto 0
							.MoveNext
						else
							if not lErrorTriggered then
								cValues = "'Branch SubTotal',''," & nCallsByBranch & ",'','','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotalByBranch) & "'"
							    oExcel.addRow cFields, cValues
								cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
								oExcel.addRow cFields, cValues
							end if
							exit do
						end if
					else
						exit do
					end if
				loop
			else
				if not lErrorTriggered then
					cValues = "'Account SubTotal',''," & nCalls & ",'','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
					oExcel.addRow cFields, cValues
					cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
					oExcel.addRow cFields, cValues
				end if
				exit do
			end if
		loop
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with

cValues = "'','','','','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genAFF = lErrorTriggered
end function


'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
'LPIC 0063
function genCnlOldXLS
dim cAHS, cStartDate, cSP, oRS, cSQL

dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime,cCallID,CLastCallID
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee,nServiceFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
nTotSvcFee=0
cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
CLastCallID=0
'
       if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
          cCustName = "Managed Care Innovations."
       else
          cCustName = Request.QueryString("CUSTNAME")
       end if

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS


set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with



'LPIC-0063 Modified SQL Dated :21/10/2005

cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"

set oRS = Conn.Execute(cSQL)

cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0

		do while Not .EOF

			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
				nCalls=nCalls+1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nServiceFee=CSng(.Fields("SERVICE_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)




				nTotalFee = nTotalFee + CSng(nServiceFee) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if


					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(nServiceFee) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(nServiceFee) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"


				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genCnlOldXLS = lErrorTriggered
end function
'-------------------------------------------------------------------------------------
'LPIC-0063
function genCnlXLS
dim cAHS, cStartDate, cSP, oRS, cSQL

dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime,cCallID,CLastCallID
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee,nServiceFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
nTotSvcFee=0
cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
CLastCallID=0
'
       if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
          cCustName = "Managed Care Innovations."
       else
          cCustName = Request.QueryString("CUSTNAME")
       end if

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS




set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with



'LPIC-0063 Modified SQL Dated :21/10/2005
cSQL = " SELECT bl.billing_detail_id,bl.billing_id,bl.status,bl.parent_node_id,bl.client_node_id,bl.account_name,"
cSQL = cSQL &  "                            bl.parent_name,bl.client_name,bl.call_id,bl.call_type,bl.loss_date,"
cSQL = cSQL &  "                            call_ali.claim_number,'ALI' as LOB_CD,bl.policy_number,"
cSQL = cSQL &  " 						   bl.insured_name,bl.caller_last_name,bl.caller_first_name,bl.caller_mi,"
cSQL = cSQL &  "                            bl.employee_last_name,bl.employee_first_name,bl.employee_mi,bl.employee_ssn,"
cSQL = cSQL &  " 						   bl.call_end_time,"
cSQL = cSQL &  "                            bl.service_fee,bl.total_fax_fee,bl.temp_fee,bl.escalate_fee,"
cSQL = cSQL &  " 						   bl.vendor_fee,bl.callstatus,bl.driven_fee "
cSQL = cSQL &  " 			FROM BILLING_DETAIL bl ,call_ali,CALL,CALL_CLAIM"
cSQL = cSQL &  " 			WHERE bl.BILLING_ID = "& cBillID &""
cSQL = cSQL &  " 			AND bl.STATUS='ACTIVE' AND bl.CALLSTATUS = 'COMPLETED' "
cSQL = cSQL &  " 			and call_ali.CALL_CLAIM_ID=CALL_CLAIM.CALL_CLAIM_ID"
cSQL = cSQL &  " 			and CALL.CALL_ID=bl.CALL_ID"
cSQL = cSQL &  " 			And CALL.CALL_ID=Call_claim.CALL_ID"
cSQL = cSQL &  " 			and call_ali.COVERAGE_FLG='Y' and bl.LOB_CD <>'INF'"
cSQL = cSQL &  " 		"
cSQL = cSQL &  " union ALL "
cSQL = cSQL &  "  SELECT bl.billing_detail_id,bl.billing_id,bl.status,bl.parent_node_id,bl.client_node_id,bl.account_name,"
cSQL = cSQL &  "                            bl.parent_name,bl.client_name,bl.call_id,bl.call_type,bl.loss_date,"
cSQL = cSQL &  "                            call_apd.claim_number,'APD' as LOB_CD,bl.policy_number,"
cSQL = cSQL &  " 						   bl.insured_name,bl.caller_last_name,bl.caller_first_name,bl.caller_mi,"
cSQL = cSQL &  "                            bl.employee_last_name,bl.employee_first_name,bl.employee_mi,bl.employee_ssn,"
cSQL = cSQL &  " 						   bl.call_end_time,"
cSQL = cSQL &  "                            bl.service_fee,bl.total_fax_fee,bl.temp_fee,bl.escalate_fee,"
cSQL = cSQL &  " 						   bl.vendor_fee,bl.callstatus,bl.driven_fee "
cSQL = cSQL &  " 			FROM BILLING_DETAIL bl ,call_apd,CALL,CALL_CLAIM"
cSQL = cSQL &  " 			WHERE bl.BILLING_ID = "& cBillID &""
cSQL = cSQL &  " 			AND bl.STATUS='ACTIVE' AND bl.CALLSTATUS = 'COMPLETED' "
cSQL = cSQL &  " 			and call_apd.CALL_CLAIM_ID=CALL_CLAIM.CALL_CLAIM_ID"
cSQL = cSQL &  " 			and CALL.CALL_ID=bl.CALL_ID"
cSQL = cSQL &  " 			And CALL.CALL_ID=Call_claim.CALL_ID"
cSQL = cSQL &  " 			and call_apd.COVERAGE_FLG='Y' and bl.LOB_CD <>'INF'"
cSQL = cSQL &  " union ALL "
cSQL = cSQL &  " SELECT bl.billing_detail_id,bl.billing_id,bl.status,bl.parent_node_id,bl.client_node_id,bl.account_name,"
cSQL = cSQL &  "                            bl.parent_name,bl.client_name,bl.call_id,bl.call_type,bl.loss_date,"
cSQL = cSQL &  "                            call_crg.claim_number,'CRG' as LOB_CD,bl.policy_number,"
cSQL = cSQL &  " 						   bl.insured_name,bl.caller_last_name,bl.caller_first_name,bl.caller_mi,"
cSQL = cSQL &  "                            bl.employee_last_name,bl.employee_first_name,bl.employee_mi,bl.employee_ssn,"
cSQL = cSQL &  " 						   bl.call_end_time,"
cSQL = cSQL &  "                            bl.service_fee,bl.total_fax_fee,bl.temp_fee,bl.escalate_fee,"
cSQL = cSQL &  " 						   bl.vendor_fee,bl.callstatus,bl.driven_fee "
cSQL = cSQL &  " 			FROM BILLING_DETAIL bl ,call_crg,CALL,CALL_CLAIM"
cSQL = cSQL &  " 			WHERE bl.BILLING_ID = "& cBillID &""
cSQL = cSQL &  " 			AND bl.STATUS='ACTIVE' AND bl.CALLSTATUS = 'COMPLETED' "
cSQL = cSQL &  " 			and call_crg.CALL_CLAIM_ID=CALL_CLAIM.CALL_CLAIM_ID"
cSQL = cSQL &  " 			and CALL.CALL_ID=bl.CALL_ID"
cSQL = cSQL &  " 			And CALL.CALL_ID=Call_claim.CALL_ID and bl.LOB_CD <>'INF'"
cSQL = cSQL &  " 			"
cSQL = cSQL &  " 			and call_crg.COVERAGE_FLG='Y'"
cSQL = cSQL &  " union ALL "
cSQL = cSQL &  " SELECT bl.billing_detail_id,bl.billing_id,bl.status,bl.parent_node_id,bl.client_node_id,bl.account_name,"
cSQL = cSQL &  "                           bl.parent_name,bl.client_name,bl.call_id,bl.call_type,bl.loss_date,"
cSQL = cSQL &  "                           bl.claim_number,'INF' as LOB_CD,bl.policy_number,"
cSQL = cSQL &  " 						   bl.insured_name,bl.caller_last_name,bl.caller_first_name,bl.caller_mi,"
cSQL = cSQL &  "                           bl.employee_last_name,bl.employee_first_name,bl.employee_mi,bl.employee_ssn,"
cSQL = cSQL &  " 						   bl.call_end_time,"
cSQL = cSQL &  "                           bl.service_fee,bl.total_fax_fee,bl.temp_fee,bl.escalate_fee,"
cSQL = cSQL &  " 						   bl.vendor_fee,bl.callstatus,bl.driven_fee "
cSQL = cSQL &  " 						   FROM BILLING_DETAIL bl "
cSQL = cSQL &  " 						   WHERE bl.BILLING_ID = "& cBillID &""
cSQL = cSQL &  " 						   AND bl.STATUS='ACTIVE' AND bl.CALLSTATUS = 'COMPLETED' and bl.LOB_CD='INF' "
cSQL = cSQL &  " 						   Order by CALL_ID "


'cSQL = "SELECT BILLING_DETAIL.*, " & _
			'"FROM BILLING_DETAIL " & _
			'"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			'" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED' and BILLING_DETAIL.LOB_CD in('INF','APC')) " & _
			'"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"

set oRS = Conn.Execute(cSQL)

cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingCNL.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Call_Tm,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0

		do while Not .EOF
			nTotSvcFee =0
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then

				' Edited as On  25 Aug 2005
				'--------------------------
				cCallID=CSng(.Fields("CALL_ID").Value)
				if(cCallID<>CLastCallID) then
					nCalls=nCalls+1
					CLastCallID=cCallID
				end if


					nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
					nTempFee = CSng(.Fields("TEMP_FEE").Value)
					nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
					nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
					nServiceFee=CSng(.Fields("SERVICE_FEE").Value)
					nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)




				nTotalFee = nTotalFee + CSng(nServiceFee) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if
					cTempDate = CStr(formatdatetime(.Fields("CALL_END_TIME").Value,vbshortdate))
					cTempTime = CStr(formatdatetime(.Fields("CALL_END_TIME").Value,vblongtime))

					cValues = cValues & "','" & cTempDate & "','" & cTempTime & "','"

					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if


					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(nServiceFee) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(nServiceFee) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"


				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			nTotSvcFee =0
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genCnlXLS = lErrorTriggered
end function

'-------------------------------------------------------------------------------------
function genCSAA
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval

serval="0"



cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
       if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
          cCustName = "Managed Care Innovations."
       else
          cCustName = Request.QueryString("CUSTNAME")
       end if

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'response.write(cSQL)
'response.end
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		serval="0"
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
			    serval="0"
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				serval = CSng(.Fields("SERVICE_FEE").Value)
				'response.write(serval)

				'Email Type ..
				if (mid((.Fields("Call_Type").Value ),1,1)="E") then
                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND FEE_TYPE_ID = 10 "

			       set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						     Do While Not oRS1.EOF
							 serval = .fields(0)
							 oRS1.MoveNext
				             loop
						     .close
						end with
               end if

                nTotSvcFee=nTotSvcFee + CSng(serval)
				nTotalFee = nTotalFee + CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if




					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genCSAA = lErrorTriggered
end function

'......................................................
'--BCAB-0586 - AAAReports---------------------------------------------------------------
function genAAA
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval

serval="0"

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
cCustName = Request.QueryString("CUSTNAME")
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with

cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'response.write(cSQL)
'response.end
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		serval="0"
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
			    serval="0"
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				serval = CSng(.Fields("SERVICE_FEE").Value)
				'response.write(serval)

				'Email Type ..
				if (mid((.Fields("Call_Type").Value ),1,1)="E") then
                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND FEE_TYPE_ID = 10 "

			       set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						     Do While Not oRS1.EOF
							 serval = .fields(0)
							 oRS1.MoveNext
				             loop
						     .close
						end with
               end if

                nTotSvcFee=nTotSvcFee + CSng(serval)
				nTotalFee = nTotalFee + CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if




					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genAAA = lErrorTriggered
end function

'......................................................

'-------------------------------------------------------------------------------------

function genXLS
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval

serval="0"



cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'response.write(cSQL)
'response.end
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		'Response.Write cParent & "<br>1"
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		serval="0"
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
			    serval="0"
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				serval = CSng(.Fields("SERVICE_FEE").Value)
				'response.write(serval)


				if (mid((.Fields("Call_Type").Value ),1,1)="E") then
                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND FEE_TYPE_ID = 8 "
					   'M

					   'response.write(cSQL1)
					   'response.end
				'if not cSQL1="" then
			       set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						     Do While Not oRS1.EOF
							 serval = .fields(0)
							 oRS1.MoveNext
				             loop
						     .close
						end with
               end if
		'response.write "test" & .Fields("PARENT_NAME").Value
		'response.end
                nTotSvcFee=nTotSvcFee + CSng(serval)
				nTotalFee = nTotalFee + CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if




					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & removeSngQuote(.Fields("POLICY_NUMBER").Value)
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genXLS = lErrorTriggered
end function


'-----------------------------------------------NBAR - 3295-----------------------------------------------

function genGBSXLS
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval

serval="0"

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

'cSQL = "SELECT BILLING_DETAIL.*, CALL_CALLER.CALLER_TYPE as CallerType, CALL_CLAIM.CLAIM_TYPE as T " & _
'			"FROM BILLING_DETAIL, CALL_CALLER, CALL_CLAIM " & _
'			"WHERE (BILLING_DETAIL.CALL_ID = CALL_CALLER.CALL_ID & _
'			" AND BILLING_DETAIL.CALL_ID = CALL_CLAIM.CALL_ID " & _
'			" AND BILLING_DETAIL.BILLING_ID = " & cBillID & _
'			" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
'			" Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"


cSQL = "SELECT BILLING_DETAIL.*, CALL_CALLER.CALLER_TYPE as CallerType, CALL_CLAIM.CLAIM_TYPE as T " & _
			"FROM BILLING_DETAIL, CALL_CALLER, CALL_CLAIM " & _
			"WHERE (BILLING_DETAIL.CALL_ID = CALL_CALLER.CALL_ID AND BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.CALL_ID = CALL_CLAIM.CALL_ID AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"

set oRS = Conn.Execute(cSQL)
'response.write(cSQL)
'response.end
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingGBS.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
'cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
cFields = "Account,Call_No,T,CallerType,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		serval="0"
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
			    serval="0"
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				serval = CSng(.Fields("SERVICE_FEE").Value)
				'response.write(serval)


				if (mid((.Fields("Call_Type").Value ),1,1)="E") then
                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND FEE_TYPE_ID = 8 "

				'if not cSQL1="" then
			       set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						     Do While Not oRS1.EOF
							 serval = .fields(0)
							 oRS1.MoveNext
				             loop
						     .close
						end with
               end if
 'Response.Write .Fields("Call_ID").Value
 'Response.End

                nTotSvcFee=nTotSvcFee + CSng(serval)
				nTotalFee = nTotalFee + CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("T").Value & "','" & _
					.Fields("CallerType").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

'Response.Write cValues
'Response.end
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genGBSXLS = lErrorTriggered
end function


'----------------------------------------- END OF NBAR - 3295 --------------------------



'......................................................PMAC 0782---------------



function genACEXLS
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval

serval="0"



cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'
cSQL = "SELECT BILLING_DETAIL.*, CALL_CLAIM.CLAIM_TYPE as ClaimType " & _
			"FROM BILLING_DETAIL, CALL_CLAIM " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.CALL_ID = CALL_CLAIM.CALL_ID AND BILLING_DETAIL.STATUS='ACTIVE' " & _
			" AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)

'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingACE.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
'cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
cFields = "Account,Call_No,Claim_Type,T,Status,Program_Type,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Benefit_State,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		serval="0"
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
			    serval="0"
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				serval = CSng(.Fields("SERVICE_FEE").Value)
				'response.write(serval)


				if (mid((.Fields("Call_Type").Value ),1,1)="E") then
                    cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
					" WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
					" AND FEE_TYPE_ID = 8 "

			       set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						     Do While Not oRS1.EOF
							 serval = .fields(0)
							 oRS1.MoveNext
				             loop
						     .close
						end with
               end if

				dim program_type

				cSQLProgram_Type = "select getprogramtypeace(" & .Fields("CALL_ID").Value &_
				                   " ) from dual"


				set cSQLProgram_Type = Conn.Execute(cSQLProgram_Type)
				with cSQLProgram_Type
					if isNull(.Fields(0).Value) then
						program_type = ""
					else
						program_type =  .fields(0)
					end if
				end with

			   'MCAS-0824 -------------get the value for Record_Only field
			    dim ClaimTypeFlg, ClaimTypeDesc
			    ClaimTypeFlg = ""
			    ClaimTypeDesc = ""

			    ClaimTypeFlg = .Fields("ClaimType").Value

			    if ClaimTypeFlg = "R" then
					ClaimTypeDesc = "Record Only"
				else
					ClaimTypeDesc = "Claim"
				end if
                '--------------------------------------
				cValues = cValues & "','" & _

                nTotSvcFee=nTotSvcFee + CSng(serval)
				nTotalFee = nTotalFee + CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					ClaimTypeDesc & "','" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','" & _
					program_type & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if




					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','"

					dim cBenefit_State_SQL,oRS_Benefit_State
					cBenefit_State_SQL = ""
					cBenefit_State_SQL =  " select benefit_state from call_claim where call_claim.call_id = " & .Fields("CALL_ID").Value

					set oRS_Benefit_State = Conn.Execute(cBenefit_State_SQL)
					with oRS_Benefit_State
						if isNull(.Fields(0).Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .fields(0)
						end if
					end with
					cValues = cValues & "','" & _

					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genACEXLS = lErrorTriggered
end function

'......................................................


'......................................................
'............added for  KFAB-0042

function genESISXLS
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
dim dRepDate, cBillID, cCustName, cCustCode,Division
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval

serval="0"

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

'CALL_TYPE should come under "T" column on this report
'Input_Type should capture CALLER:CALLER_TYPE.
'CLAIM:CALLER_TYPE = "FAX" then this is a fax claim. Fax fee will be charged.
'CLAIM:CALLER_TYPE = "EML" then this is a email claim. Email fee will be charged.
'CLAIM:CALLER_TYPE = "INT" then this is a internet claim. Internet fee will be charged.
'For other CLAIM:CALLER_TYPE the claims are telephonic. Phone fee will be charged.

cSQL = "SELECT distinct b.*, u.name as user_name, c.call_start_time as CallStartTime, c.call_end_time as CallEndTime, " & _
		"cc.TEMPEDPOLICY_FLG, cc.claim_type as ClaimType, " & _
		"DECODE(cca.caller_Type, 'IFTCO', 'TRANS', cca.caller_Type) as Input_Type, " & _
		"cb.BRANCH_NUMBER, cb.BRANCH_OFFICE_NAME, cll.ADDRESS_STATE  as LossState " & _
		"FROM BILLING_DETAIL b, call c, call_caller cca, call_claim cc, users u, call_loss_location cll, " & _
		"(select cc.call_id,cb.branch_number,cb.branch_office_name " & _
		"from call_claim cc,call_branch cb " & _
		"where cc.call_claim_id = cb.call_claim_id(+) )cb " & _
		"WHERE (b.BILLING_ID = " & cBillID & _
		" AND b.STATUS = 'ACTIVE' AND b.CALLSTATUS = 'COMPLETED' and b.call_id = c.call_id AND c.CALL_ID = cca.CALL_ID(+) " & _
		" AND c.CALL_ID = cc.CALL_ID AND c.user_id = u.user_id AND cc.CALL_ID = cb.CALL_ID(+)) AND cc.CALL_CLAIM_ID = cll.CALL_CLAIM_ID(+) " & _
		"Order by b.CLIENT_NAME,b.PARENT_NAME,b.CALL_TYPE,b.LOB_CD,b.CALL_END_TIME"

'response.write(cSQL)
'response.end


set oRS = Conn.Execute(cSQL)
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingESIS.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
'-----------------Loss state
'cFields = "Account,Rep,Claim_Type,Input_Type,Branch_Number,Branch_Name,Call_No,T,Status,Division,Loss_Dt,Call_Start_Dt,Call_End_Dt,Duration,Claim_No,LOB,Policy_No,Employee_Insured,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
cFields = "Account,Rep,Claim_Type,Input_Type,Branch_Number,Branch_Name,Call_No,T,Status,Division,Loss_Dt,Loss_State,Call_Start_Dt,Call_End_Dt,Duration,Claim_No,LOB,Policy_No,Employee_Insured,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"

'cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"

with oRS
	Do While Not .EOF


		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		serval=0
		do while Not .EOF


			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if

			if cParent = cCmpParent then
			    serval="0"
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				serval = CSng(.Fields("SERVICE_FEE").Value)
				'response.write(serval)

				''''''''''RBEG-0002'''''''''
				if (.Fields("Input_Type").Value = "EML" or .Fields("Input_Type").Value = "EMAIL") then

					if (.Fields("LOB_CD").Value <> "INF") then

                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND CALL_TYPE = 'E' "

					'if not cSQL1="" then
					set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						     Do While Not oRS1.EOF
							 serval = .fields(0)
							 oRS1.MoveNext
				             loop
						     .close
						end with
					end if
               end if

               if (.Fields("Input_Type").Value = "NET" or .Fields("Input_Type").Value = "INT" ) then

					if (.Fields("LOB_CD").Value <> "INF") then

                              cSQL2 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND CALL_TYPE = 'N' "

					'if not cSQL2="" then
					set oRS2 = Conn.Execute(cSQL2)
						with oRS2
						     Do While Not oRS2.EOF
							 serval = .fields(0)
							 oRS2.MoveNext
				             loop
						     .close
						end with
					end if
               end if

               ''''''''''''''RBEG-0002''''''''''''''''''


				'Get Divison Value
				Division = "ESIS"
				'cSQL3= " select getdivision(" & .Fields("Call_ID").Value & ") from dual"
				'cSQL3= " select Value from X_CALL_CLAIM where call_claim_ID= '"& .Fields("Call_CLAIM_ID").Value &"' and FIELD = 'PROGRAM_TYPE'"
				'set oRS3 = Conn.Execute(cSQL3)
				'with oRS3
					'if Not oRS3.EOF then
						'Division = .fields(0)

					'end if
					'.close
				'end with

			    'END

				'GetProgramType
				dim ProgramType

				'Get account name
				dim Call_Status
				dim DivisionName
				DivisionName=division

				Call_Status = .Fields("TEMPEDPOLICY_FLG").Value

				'not  able to test as there is not enough data, so it is hardcoded for the time

				'DivisionName = "ESIS"

			   'Get Duration Value
				dim Duration
				Duration = Round(DateDiff("s",.Fields("CallStartTime").Value,.Fields("CallEndTime").Value) / 60)
			   'END

			   'MCAS-0824 -------------get the value for Record_Only field
			    dim ClaimTypeFlg, ClaimTypeDesc
			    ClaimTypeFlg = ""
			    ClaimTypeDesc = ""

			    ClaimTypeFlg = .Fields("ClaimType").Value

			    if ClaimTypeFlg = "R" then
					ClaimTypeDesc = "Record Only"
				else
					ClaimTypeDesc = "Claim"
				end if
                '--------------------------------------
                nTotSvcFee=nTotSvcFee + CSng(serval)
				nTotalFee = nTotalFee + CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "','" & _
					removeSngQuote(.Fields("USER_NAME").Value) & "','" & _
					ClaimTypeDesc & "','" & _
					.Fields("INPUT_TYPE").Value & "','"

					if isNull(.Fields("BRANCH_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("BRANCH_NUMBER").Value
					end if

					cValues = cValues & "','"

					if isNull(.Fields("BRANCH_OFFICE_NAME").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("BRANCH_OFFICE_NAME").Value
					end if

					cValues = cValues & "','" & _
					.Fields("CALL_ID").Value & "','" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','" & _
					DivisionName & "','"

					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					'------ Added LossState for JMAR-0511
					cValues = cValues & "','" & _
					.Fields("LossState").Value & "','" & _
					.Fields("CallStartTime").Value & "','" & _
					.Fields("CallEndTime").Value & "','" & _
					Duration & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("INSURED_NAME").Value) & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
					'response.write cValues
					'response.end

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop

		if not lErrorTriggered then
			''cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			''cValues = "'','','','','','','','','','','','','','','','','','',''"
			cValues = "'','','','','','','','','','','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
		'Response.End
	Loop
	.Close
end with
'cValues = "'','','','','','','','','','','','','','','','','','',''"
cValues = "'','','','','','','','','','','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
'cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genESISXLS = lErrorTriggered
end function


'......................................................

'Added for JPRI-0941

function genAmcXLS
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls




cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
       if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
          cCustName = "Managed Care Innovations."
       else
          cCustName = Request.QueryString("CUSTNAME")
       end if

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with

cSQL = " select * "
cSQL = cSQL &  " from ( "
cSQL = cSQL &  " SELECT "
cSQL =cSQL &  "   bd.*,b.BRANCH_OFFICE_NAME"
cSQL = cSQL &  " FROM "
cSQL = cSQL &  "   BILLING_DETAIL bd,call_Branch b,call_claim c"
cSQL = cSQL &  " WHERE"
cSQL = cSQL &  "    c.call_claim_id=b.call_claim_id(+)"
cSQL = cSQL &  "    and c.CALL_ID=bd.CALL_ID "
cSQL = cSQL &  "   And (bd.BILLING_ID = "& cBillID &" AND"
cSQL = cSQL &  "   bd.STATUS='ACTIVE' AND"
cSQL = cSQL &  "   bd.CALLSTATUS = 'COMPLETED'))"
cSQL = cSQL &  "   Order by LOB_CD desc,BRANCH_OFFICE_NAME,PARENT_NAME,CLIENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"



set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingAMC.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Branch_Office_Name,Account,Call_No,Type,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"

with oRS
	Do While Not .EOF
		if isNull(.Fields("branch_office_name").Value) then
			cParent = ""
		else
			cParent = .Fields("branch_office_name").Value

		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("branch_office_name").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("branch_office_name").Value

			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues = "'" & .Fields("BRANCH_OFFICE_NAME").Value & "','"
				cValues =	cValues & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then

			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"

			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else

			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"

oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genAmcXLS = lErrorTriggered
end function

'*********************************
'sen
'********************************
function genSEN
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL2,cSQL1,cSQL3,cSQL4,cSQL5,oRS1,oRS3,oRS4,oRS5
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim cClaim_Number ,cClaim_Number_Tmp
dim cRep_Tmp,nCount,nCount1,nCount2,loopCount,cCallType_Tmp
dim testCount,morethanonerecord
dim x,y,z,s,p,q

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
cSQL3=""
cSQL4=""
loopCount = 0
nCount2=0
cCustName = Request.QueryString("CUSTNAME")

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with

dim oBaseRs
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
nCount=0
nCount1=0
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingSEN.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
		'Count total number of specific calls beforehand as per 01/03/2006 discussion
		testCount = 0
		if (loopCount = 0 ) then
                    cSQL3 = " SELECT COUNT(*) as CallCount FROM BILLING_DETAIL WHERE " & vbnewline & _
                            " SUBSTR(TRIM(NAME),1,1)<>'$' " & vbnewline & _
                            " AND TRIM(CALL_TYPE)='C' " & vbnewline & _
                            " AND BILLING_ID = " & cBillID & vbnewline
                            '" AND TRIM(CALLSTATUS)<>'HOLD-RESOLVED' "
                    set oRS3 = Conn.Execute(cSQL3)
                         'with oRS3
							Do While Not oRS3.EOF
								nCount = cInt(Trim(oRS3.fields("CallCount").Value))
								testCount= nCount
								oRS3.MoveNext
							loop
						 oRS3.Close
						 'end with


					'Response.Write("First SQL value is " & testCount)

					cSQL4 = " SELECT COUNT(*) as CallCount FROM BILLING_DETAIL WHERE " & vbnewline & _
                            " SUBSTR(TRIM(NAME),1,1)='$' " & vbnewline & _
                            " AND TRIM(CALL_TYPE)='C' " & vbnewline & _
                            " AND TRIM(CALLSTATUS)='COMPLETED' " & vbnewline & _
                            " AND TRIM(REPORT_TYPE)='I' " & vbnewline & _
                            " AND BILLING_ID = " & cBillID

                    set oRS4 = Conn.Execute(cSQL4)
                         'with oRS4
							Do While Not oRS4.EOF
								nCount1 = cInt(Trim(oRS4.fields("CallCount").Value))
								testCount= nCount1
								oRS4.MoveNext
							loop
						 oRS4.Close
						 'end with

					'Response.Write("Second SQL value is " & testCount)

					cSQL5 = " SELECT COUNT(*) as CallCount FROM BILLING_DETAIL WHERE " & vbnewline & _
                            " SUBSTR(TRIM(NAME),1,1)='$' " & vbnewline & _
                            " AND TRIM(CALL_TYPE)='C' " & vbnewline & _
                            " AND TRIM(CALLSTATUS)='COMPLETED' " & vbnewline & _
                            " AND BILLING_ID = " & cBillID & vbnewline & _
                            " AND (TRIM(REPORT_TYPE) <>'I' " & vbnewline & _
                            " OR REPORT_TYPE is null) "


                    set oRS5 = Conn.Execute(cSQL5)
                         'with oRS5
							Do While Not oRS5.EOF
								nCount2 = cInt(Trim(oRS5.fields("CallCount").Value))
								testCount= nCount2
								oRS5.MoveNext
							loop
						 oRS5.Close
						 'end with
					'Response.Write("Third SQL value is " & testCount)
					loopCount = loopCount + 1
           End if
           'End as per 01/03/2006 discussion



cFields = "Account,Rep,Input_type,Transportation,Branch_Number,Undrwrtg_cmpny,Account_CC,Call_No,T,Status,Division,Loss_Dt,Call_Dt,Call_end,Duration,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Print,Total"
	dim cSelectQry
	cSelectQry = "bd.CLIENT_NAME,  bd.PARENT_NAME,  bd.TOTAL_FAX_FEE," & vbnewline & _
				"bd.TEMP_FEE,  bd.ESCALATE_FEE,  bd.DRIVEN_FEE, " & vbnewline & _
				"bd.NAME,  bd.REPORT_TYPE,  bd.BRANCH_NUMBER, " & vbnewline & _
				"bd.UNDRWRTG_CMPNY,  bd.SEC_CD,  bd.CALL_ID, " & vbnewline & _
				"bd.CALL_TYPE,  bd.CALLSTATUS,  bd.POLICY_NUMBER, " & vbnewline & _
				"bd.LOSS_DATE,  bd.CALL_START_TIME,  bd.CALL_END_TIME, " & vbnewline & _
				"bd.DURATION,  bd.CLAIM_NUMBER,  bd.LOB_CD, " & vbnewline & _
				"bd.ACCOUNT_NAME,  bd.CALLER_FIRST_NAME,  bd.CALLER_LAST_NAME, " & vbnewline & _
				"bd.EMPLOYEE_FIRST_NAME,  bd.EMPLOYEE_LAST_NAME,  bd.EMPLOYEE_SSN, " & vbNewLine & _
				"NVL(TRANSPORTATION, 'N') AS TRANSPORTATION "
	cSQL = ""
	cSQL1=""
	cSQL = "SELECT " & cSelectQry & vbnewline & _
				" FROM BILLING_DETAIL bd " & vbnewline & _
				" WHERE (bd.BILLING_ID = " & cBillID & vbnewline & _
				" AND bd.STATUS='ACTIVE') " & vbnewline & _
				" ORDER BY bd.CLIENT_NAME, bd.PARENT_NAME, bd.CALL_TYPE, bd.LOB_CD, bd.CALL_END_TIME"

	set oRS = Conn.Execute(cSQL)

	with oRS
		Do While Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cParent = ""
			else
				cParent = .Fields("PARENT_NAME").Value
			end if

			nCalls = 0
			nTotSvcFee = 0
			nTotalFee = 0

			do while Not .EOF
				if isNull(.Fields("PARENT_NAME").Value) then
					cCmpParent = ""
				else
					cCmpParent = .Fields("PARENT_NAME").Value
				end if

		if cParent = cCmpParent then
			nCalls = nCalls + 1
			nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
			nTempFee = CSng(.Fields("TEMP_FEE").Value)
			nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
			nDrivenFee = CSng(.Fields("DRIVEN_FEE").Value)
			'nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
			'nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nDrivenFee
			cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
				"'" & .Fields("name").Value & "'," & _
				"'" & .Fields("report_type").Value & "'," & _
				"'" & .Fields("transportation").Value & "'," & _
				"'" & .Fields("branch_number").Value & "'," & _
				"'" & .Fields("undrwrtg_cmpny").Value & "'," & _
				"'" & .Fields("sec_cd").Value & "','" & _
				CStr(.Fields("Call_ID").Value) & "','" & _
				.Fields("Call_Type").Value & "','" & _
				.Fields("CALLSTATUS").Value & "','"

				'----------- Modified on 9th NOV------------
				cType=trim(.Fields("Call_Type").Value)
				cAccnt_CC=trim(.Fields("sec_cd").Value)
				cUndrwrtg_cmpny=trim(.Fields("undrwrtg_cmpny").Value)

				if isNull(.Fields("POLICY_NUMBER").Value) then
					cPolicy_No =  ""
				else
					cPolicy_No = trim(.Fields("POLICY_NUMBER").Value)
				end if

				if cType="I" then
					cDivision = "INFO"
				elseif cUndrwrtg_cmpny="045" or cUndrwrtg_cmpny="048" or cUndrwrtg_cmpny="45" or cUndrwrtg_cmpny="48"then
					cDivision = "PS"
				elseif cUndrwrtg_cmpny="088" or cUndrwrtg_cmpny="88" then
					cDivision = "SSDO"
				elseif cAccnt_CC="J" or instr(cPolicy_No,"90")=1 or instr(cPolicy_No,"91")=1 then
					cDivision = "NA"
				elseif (left(cPolicy_No,1) <> "0") and len(cPolicy_No)=16 then
					cDivision = "SBP"
				else
					cDivision = "OTHER"

				end if
				cValues = cValues  & cDivision & "','"

				'-------------------------------------------
				'on error resume next
				if isNull(.Fields("LOSS_DATE").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & .Fields("LOSS_DATE").Value
				end if

				if err.number <> 0 then
					if err.number = -2147217887 then
						writeError .Fields("Call_ID").Value
					end if
					lErrorTriggered = true
					exit do
				end if
				cValues = cValues & "','" & _
				CStr(.Fields("CALL_start_TIME").Value) & "','" &_
				CStr(.Fields("CALL_end_TIME").Value) & "','" &_
				CStr(.Fields("duration").Value) & "','"
				if isNull(.Fields("CLAIM_NUMBER").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & .Fields("CLAIM_NUMBER").Value
				end if
				cValues = cValues & "','" & _
				.Fields("LOB_CD").Value & "','"
				if isNull(.Fields("POLICY_NUMBER").Value) then
					cValues = cValues & ""
				else
					cValues = cValues & .Fields("POLICY_NUMBER").Value
				end if
				cValues = cValues & "','" & _
				removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
				redim aNameParts(1)
				aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
				aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
				cValues = cValues & getName(aNameParts) & "','"
				redim aNameParts(1)
				aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
				aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value

				'---------------Modified on 17th Feb----------------
				cInputType=trim(.Fields("report_type").Value)
				cRep= trim(.Fields("name").Value)
				cStatus=trim(.Fields("CALLSTATUS").Value)
				nServiceFee=0

				cSQL2=""
				cClaim_Number = trim(.Fields("CLAIM_NUMBER").Value)
				cCallType_Tmp=""
				cRep_Tmp = ""
				cClaim_Number_Tmp = ""
				morethanonerecord=0
				cSQL1 = "SELECT NAME"  & vbnewline & _
				" FROM BILLING_DETAIL " & vbnewline & _
				" WHERE BILLING_ID = " & cBillID & vbnewline & _
				" AND STATUS='ACTIVE' " & vbnewline & _
				" AND CALLSTATUS='HOLD-RESOLVED' " & vbnewline & _
				" AND CALL_TYPE='C'" & _
				" AND CLAIM_NUMBER='" & cClaim_Number & "'"


				if (cStatus="COMPLETED") then
					 set oRS1 = Conn.Execute(cSQL1)
					 'With oRS1
						Do While Not oRS1.EOF
											cCallType_Tmp="C"
											cClaim_Number_Tmp = cClaim_Number
											cRep_Tmp =  oRS1.fields("NAME").Value
											morethanonerecord=1
						oRs1.MoveNext
						Loop
				     oRS1.Close
				     'End With

			    End if

                nServiceFee=0

				if (left(cRep, 1)<>"$") then
					if cType="I" then
						nServiceFee=3
					elseif cType="C" then
							if nCount>2501 and nCount<5001 then
								nServiceFee=15
							elseif nCount>=5001 then
								nServiceFee=14
							else
								nServiceFee=16
							end if
					end if
				elseif (Left(cRep,1)= "$") then
					if (cType="I") then
						nServiceFee=0
					elseif (cType="C") then
						if (morethanonerecord=0) then
							if (cStatus ="COMPLETED") then
							   if (cInputType <> "I" or cInputType = "" or IsNull(cInputType)) then
									if nCount2>1 and nCount2<7501 then
										nServiceFee=5
									end if
									if nCount2>=7501 and nCount2<10001 then
										nServiceFee=4.5
									end if
									if nCount2>=10001 then
										nServiceFee=4
									end if
								else
									if (nCount1>1 and nCount1<7501) then
										nServiceFee=5
									end if
									if (nCount1>=7501 and nCount1<10001) then
										nServiceFee=4.5
									end if
									if (nCount1>=10001) then
										nServiceFee=4
									end if
								end if
							else
								nServiceFee=0
						    end if
						elseif (morethanonerecord=1) then
							if (cInputType = "I" And cStatus ="COMPLETED") then
								if (nCount1>1 and nCount1<7501) then
									nServiceFee=5
								end if
								if (nCount1>=7501 and nCount1<10001) then
									nServiceFee=4.5
								end if
								if (nCount1>=10001) then
									nServiceFee=4
								end if
							elseif (cInputType = "I" And cStatus ="HOLD-RESOLVED") then
								nServiceFee=0
							elseif (cInputType <> "I" and cStatus="COMPLETED") then
							    if nCount2>1 and nCount2 < 7501 then
							       nServiceFee=5
							    end if
							    if nCount2>=7501 and nCount2 < 10001 then
								   nServiceFee = 4.5
								end if
								if nCount2 >= 10001 then
								   nServiceFee = 4
								end if
							end if
							'*************************
						end if 'more than one record=no
					end if
				end if


				nTotSvcFee = nTotSvcFee + nServiceFee
				nTotalFee = nTotalFee + nServiceFee + nTotalFaxFee + nTempFee + nEscalateFee + nDrivenFee
				'-----------------------------------------------------

				cValues = cValues & getName(aNameParts) & "','"
				cValues = cValues & cEmployeeSSN & "','" & _
				FormatNumber(nServiceFee) & "','" & _
				FormatNumber(nTotalFaxFee) & "','" & _
				FormatNumber(nTempFee) & "','" & _
				FormatNumber(nEscalateFee) & "','" & _
				FormatNumber(nDrivenFee) & "','" & _
				FormatNumber(nServiceFee + nTotalFaxFee + nTempFee + nEscalateFee + nDrivenFee) & "'"
				oExcel.addRow cFields, cValues
			.MoveNext
		else
			exit do
		end if
	loop
	if not lErrorTriggered then
		cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
		oExcel.addRow cFields, cValues
		cValues=""
		'cValues = "'','','','','','','','','','','','','','','','','','','','','','','','','','','',''"
		'oExcel.addRow cFields, cValues
	else
		exit do
	end if
			nGrandTotal = nGrandTotal + nTotalFee
			nTotalNoCalls = nTotalNoCalls + nCalls
		Loop
		.Close
	end with


cValues = "'','','','','','','','','','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.writeMsg "Report Generation is Completed."
	oExcel.sendFile
end if
Set oRS1 = Nothing
Set oRS = Nothing
genSEN = lErrorTriggered
end function






' *********************************************************************
'	McDonalds
' *********************************************************************
function genMAC
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim cCarrierName
dim cBranchName, cCmpBranch

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'
cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
			"Order by BRANCH_NAME"
set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\McDonalds.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Branch,Carrier,Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("BRANCH_NAME").Value) then
			cBranchName = ""
		else
			cBranchName = removeSngQuote(.Fields("BRANCH_NAME").Value)
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("BRANCH_NAME").Value) then
				cCmpBranch = ""
			else
				cCmpBranch = removeSngQuote(.Fields("BRANCH_NAME").Value)
			end if
			if cBranchName = cCmpBranch then
				nCalls = nCalls + 1
				cCarrierName = removeSngQuote(.Fields("CARRIER_NAME").Value)
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & cBranchName & "'," & _
					"'" & cCarrierName & "'," & _
					"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				response.Write ("Error. Call ID: " & CStr(.Fields("Call_ID").Value))
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & _
				"'',''," & nCalls & ",'','','','','','','','','','','" & FormatCurrency(nTotSvcFee,2) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','',''," & nTotalNoCalls & ",'','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genMAC = lErrorTriggered
end function


sub writeError(nCallID)
with response
	.Write "<script language=""JScript"">" & vbCRLF
	.Write "w = window.open("""",""error_window"",""width=500,height=300,toolbar=no,location=no,directories=no,status=no,menubar=no"")" & vbCRLF
	.Write "w.document.write(""<HTML>"");" & vbCRLF
	.Write "w.document.write(""<BODY bgColor=tomato>"");" & vbCRLF
	.Write "w.document.write(""<CENTER>"");" & vbCRLF
	.Write "w.document.write(""<H2><font size=4 face=Verdana, Tahoma, Arial>Incorrect Data Detected. The Report cannot continue:</font></H2>"");" & vbCRLF
	.Write "w.document.write(""<H2><font size=3 face=Verdana, Tahoma, Arial>Please correct the 'Loss Date' and rerun the report.</font></H2>"");" & vbCRLF
	.Write "w.document.write(""<TABLE border=1 cellpadding=10 bgcolor=#dddddd><TR><TD>"");" & vbCRLF
	.Write "w.document.write(""<font size=3 face=Verdana, Tahoma, Arial>Call ID = " & nCallID & "</font>"");" & vbCRLF
	.Write "w.document.write(""</TD></TR></TABLE>"");" & vbCRLF
	.Write "w.document.write(""<FORM id=form1 name=form1><INPUT type=button value=\"" OK \"" onclick=self.close()></FORM>"");" & vbCRLF
	.Write "w.document.write(""</CENTER>"");" & vbCRLF
	.Write "w.document.write(""</BODY>"");" & vbCRLF
	.Write "w.document.write(""</HTML>"");" & vbCRLF
	.Write "</script>" & vbCRLF
end with
end sub

sub writeErrText()
with response
	  .Write "<script language=""JScript"">" & vbCRLF
	  .Write "w = window.open("""",""error_window"",""width=500,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no"")" & vbCRLF
	  .Write "w.document.write(""<HTML>"");" & vbCRLF
	  .Write "w.document.write(""<BODY bgColor=tomato>"");" & vbCRLF
	  .Write "w.document.write(""<CENTER>"");" & vbCRLF
	  .Write "w.document.write(""<H2><font size=4 face=Verdana, Tahoma, Arial>An error ocurred. The Report cannot continue:</font></H2>"");" & vbCRLF
	  .Write "w.document.write(""<TABLE border=1 cellpadding=10 bgcolor=#dddddd><TR><TD>"");" & vbCRLF
	  .Write "w.document.write(""<font size=3 face=Verdana, Tahoma, Arial>Error number = " & err.number & "</font></TD></TR>"");" & vbCRLF
	  .Write "w.document.write(""<TR><TD><font size=3 face=Verdana, Tahoma, Arial>Error description = " & err.Description & "</font>"");" & vbCRLF
	  .Write "w.document.write(""</TD></TR></TABLE>"");" & vbCRLF
	  .Write "w.document.write(""<FORM id=form1 name=form1><INPUT type=button value=\"" OK \"" onclick=self.close()></FORM>"");" & vbCRLF
	  .Write "w.document.write(""</CENTER>"");" & vbCRLF
	  .Write "w.document.write(""</BODY>"");" & vbCRLF
	  .Write "w.document.write(""</HTML>"");" & vbCRLF
	  .Write "</script>" & vbCRLF
end with
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
with oRS
	Do While Not .EOF
		If .Fields("CALL_TYPE").Value = "C" Then
	        nCallCounter = nCallCounter + 1
	    ElseIf .Fields("CALL_TYPE").Value = "F" Then
			nFaxCounter = nFaxCounter + 1
		End If
		.MoveNext
	Loop
	.Close
end with
set oRS = nothing
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
end sub

sub genFremontXLS
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), nCalls, nTotSvcFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim nBranchNo, nNextBranchNo
dim nGrandTotal
dim nTotalNoCalls

with Request
	cAHS = .QueryString("AHS")
	cStartDate = .QueryString("DATEFROM")
	cCustCode = .QueryString("CUSTCODE")
	cCustName =.QueryString("CUSTNAME")
end with
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'
cSQL = "SELECT BILLING_DETAIL.*, " & _
		"CALL_BRANCH.BRANCH_NUMBER, " & _
		"CALL_BRANCH.BRANCH_OFFICE_NUMBER, " & _
		"CALL_BRANCH.BRANCH_OFFICE_NAME " & _
			"FROM CALL_CLAIM, " & _
			"CALL_BRANCH, " & _
			"BILLING_DETAIL " & _
			"WHERE BILLING_DETAIL.CALL_ID = CALL_CLAIM.CALL_ID (+) " & _
			"AND CALL_CLAIM.CALL_CLAIM_ID = CALL_BRANCH.CALL_CLAIM_ID (+) " & _
			"AND BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' " & _
			"Order by CALL_BRANCH.BRANCH_NUMBER, CALL_TYPE, CALL_END_TIME"

set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Fremont.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Branch_Name,Branch_No,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("BRANCH_NUMBER").Value) then
			nBranchNo = ""
		else
			nBranchNo = .Fields("BRANCH_NUMBER").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("BRANCH_NUMBER").Value) then
				nNextBranchNo = ""
			else
				nNextBranchNo = .Fields("BRANCH_NUMBER").Value
			end if
			if nBranchNo = nNextBranchNo then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("BRANCH_OFFICE_NAME").Value) & "','" & _
					.Fields("BRANCH_NUMBER").Value & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if
					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
		cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
		oExcel.addRow cFields, cValues
		cValues = "'','','','','','','','','','','','','','','','','','',''"
		oExcel.addRow cFields, cValues
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
with oExcel
	.closeXLS
	.sendFile
end with
Set oRS = Nothing
end sub

'==============================================================================
'====WIG
'=============================================================================
function genWIG
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls , cParServfee ,cCompServ,cConfirmFlag
cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")


lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)

   with oRS
	  cBillID = .fields(0)
	   .close
   end with

'
'
dim  ordField
  cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
	"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
	   "Order by SERVICE_FEE DESC,CLIENT_NAME ,PARENT_NAME,CALL_TYPE, LOB_CD, CALL_END_TIME"


set oRS = Conn.Execute(cSQL)
'DESC
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	 'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	  cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS

	Do While Not .EOF' ---ckop by the parent name
	                 cParent            =  .Fields("CLIENT_NAME").Value
		             cParServfee        =  .Fields("SERVICE_FEE").Value

		             if ISNULL(.Fields("SERVICE_FEE").Value)then
			            cParServfee = 0
			         else
				         cParServfee  = CSng(.Fields("SERVICE_FEE").Value)
			         end if
        nCalls = 0
	    nTotSvcFee = 0
	    nTotalFee = 0

	 do while Not .EOF
	     if cParent=.Fields("CLIENT_NAME").Value then ' cParent= .Fields("PARENT_NAME").Value then

			       'brake by servicfee
			        if ISNULL(.Fields("SERVICE_FEE").Value ) then
			            cCompServ= 0
			          else
				        cCompServ = CSng(.Fields("SERVICE_FEE").Value)
				     end if
				end if
			 'END IF
		if cParServfee= cCompServ  then
			    nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = CInt(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)

		        nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("CLIENT_NAME").Value ) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
      end if
'end if
  loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genWIG = lErrorTriggered
end function



'==============================================================================
'====MGC
'=============================================================================
function genMGC
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee,sNameReport
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls , cParServfee ,cCompServ,cConfirmFlag ,parConfirmFlag, sClient

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")


lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)

   with oRS
	  cBillID = .fields(0)
	   .close
   end with

'
'
dim  ordField
  cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
	"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
	   "Order by CONFIRM_FLG ASC, SERVICE_FEE DESC, CLIENT_NAME ,LOB_CD,PARENT_NAME, CALL_TYPE, CALL_END_TIME"


set oRS = Conn.Execute(cSQL)
'DESC
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)

end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS

	Do While Not .EOF' ---ckop by the parent name
	            '***************************************************
	             sClient   = "" & .Fields("CLIENT_NODE_ID").Value
	                      if sClient =103 then
	                     ' Chek if flaf exist in MGC (103),then confirm name of the subreport
	          	              parConfirmFlag       = "" &.Fields("CONFIRM_FLG").Value
		                       if parConfirmFlag = "N" then
		                          sNameReport = "Full First Report Charges"
		                        end if
		                        if parConfirmFlag = "Y" then
		                          sNameReport = "IRC Charge"
		                        end if
		                      end if
		            '***************************************
		             cParent            =  .Fields("CLIENT_NAME").Value
		             cParServfee        =  .Fields("SERVICE_FEE").Value


		             if ISNULL(.Fields("SERVICE_FEE").Value)then
			            cParServfee = 0
			         else
				         cParServfee  = CSng(.Fields("SERVICE_FEE").Value)
			         end if
        nCalls = 0
	    nTotSvcFee = 0
	    nTotalFee = 0

	 do while Not .EOF
	     if cParent=.Fields("CLIENT_NAME").Value then ' cParent= .Fields("PARENT_NAME").Value then
	              'brake by Confirmation flg
	               cConfirmFlag  = "" &.Fields("CONFIRM_FLG").Value

			       'brake by servicfee
			        if ISNULL(.Fields("SERVICE_FEE").Value ) then
			            cCompServ= 0
			          else
				        cCompServ = CSng(.Fields("SERVICE_FEE").Value)
				     end if
				end if
			 'END IF
	if	parConfirmFlag = cConfirmFlag and  cParServfee= cCompServ then  ' ====chek aconfirmation flag
	   nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = CInt(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
			    nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("CLIENT_NAME").Value ) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
      end if
  loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genWIG = lErrorTriggered
end function

'==============================================================================
'====ONB
'=============================================================================
function genONB
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls , cParServfee ,cCompServ,cConfirmFlag
cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")


lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)

   with oRS
	  cBillID = .fields(0)
	   .close
   end with

'
'
dim  ordField
  cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
	"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
	   "Order by SERVICE_FEE DESC,CLIENT_NAME ,PARENT_NAME,CALL_TYPE, LOB_CD, CALL_END_TIME"


set oRS = Conn.Execute(cSQL)
'DESC
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	'.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingONB.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)

end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Service,Total_Fax,Temp,Escalate,Vendor,PTC_Net_referral,Drivein_Fee,Total"
with oRS

	Do While Not .EOF' ---ckop by the parent name
	                 cParent            =  .Fields("CLIENT_NAME").Value
		             cParServfee        =  .Fields("SERVICE_FEE").Value

		             if ISNULL(.Fields("SERVICE_FEE").Value)then
			            cParServfee = 0
			         else
				         cParServfee  = CSng(.Fields("SERVICE_FEE").Value)
			         end if
        nCalls = 0
	    nTotSvcFee = 0
	    nTotalFee = 0

	 do while Not .EOF
	     if cParent=.Fields("CLIENT_NAME").Value then ' cParent= .Fields("PARENT_NAME").Value then

			       'brake by servicfee
			        if ISNULL(.Fields("SERVICE_FEE").Value ) then
			            cCompServ= 0
			          else
				        cCompServ = CSng(.Fields("SERVICE_FEE").Value)
				     end if
				end if
			 'END IF
		if cParServfee= cCompServ  then
			    nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = CSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTriage = CSng(.Fields("Triage").Value)
			    nDriven_Fee = CSng(.Fields("DRIVEN_FEE").Value)

		        nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nTriage + nDriven_Fee
				cValues =	"'" & removeSngQuote(.Fields("CLIENT_NAME").Value ) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(nTriage) & "','" & _
					FormatNumber(nDriven_Fee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nTriage + nDriven_Fee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
      end if
'end if
  loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genONB = lErrorTriggered
end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	AGENT BILLING
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub doAgentBilling(oConn, oExcel)
dim cStartDate, oRS, oRS1, cSQL
dim cRepDate, cServType, cAgentName, cPaymentMethod
dim cDownloadLocation, cValues, cAHS
dim dRepStart, dRepEnd, cStart, cEnd
dim nTotalUnits, nTotalUnits4Claims, nTotalUnits4INF , nTotalDollars
dim nGrandTotal, cMonth
Dim dCreatedDate
Dim lFirstMonth
Dim nFreeMonthlyUnits, nTotalEscalations, nTotalAgents, x

'cStartDate = Request.QueryString("DATEFROM")
'---------------------------------------
' ILOG issue MROU-2726
'Modified By R.Narayan
dRepStart = CDate(Request.QueryString("DATEFROM"))
dRepEnd = CDate(Request.QueryString("DATETO"))
'cRepDate = ucase("1-" & left(cStartDate,3) & "-" & right(cStartDate,4))
'dRepStart = CDate(cRepDate)
dRepEnd = DateAdd("m", 1, dRepStart)
cStart = day(dRepStart ) & "-" & MonthName(month(dRepStart ),true) & "-" & year(dRepStart )
cEnd = day(dRepEnd) & "-" & MonthName(month(dRepEnd),true) & "-" & year(dRepEnd)
dRepEnd = DateAdd("d", -1, dRepEnd)
''----------------------------------------------

'
cTmpFile = "AgentBilling" & cStartDate & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\AgentBilling.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account_ID,Invoice_No,Agent_Name,Escalations,Service_Units,Service_Plan,Payment_Method,Amount_to_Bill"

'cSQL = "Select count(*) AS nTot From ACCOUNT_HIERARCHY_STEP Where PARENT_NODE_ID = 23"
'set oRS = oConn.Execute(cSQL)
'nTotalAgents = oRS.Fields("nTot").value
cSQL = "Select * From ACCOUNT_HIERARCHY_STEP Where PARENT_NODE_ID = 23"
set oRS = oConn.Execute(cSQL)
x = 0
do while not oRS.eof
	x = x + 1
	If not isnull(oRS.Fields("AGENT_BILLING_METHOD").value) Then
		cServType = UCase(oRS.Fields("AGENT_BILLING_METHOD").value)
	End If
	if not isnull(oRS.Fields("NAME").value) Then
		cAgentName = oRS.Fields("NAME").value
	End If
	if not isnull(oRS.Fields("AGENT_PAYMENT_TYPE").value) Then
		cPaymentMethod = oRS.Fields("AGENT_PAYMENT_TYPE").value
		if cPaymentMethod = "CREDIT" then
			cPaymentMethod = "Credit Card"
		end if
	End If
	dCreatedDate = CDate(oRS.Fields("CREATED_DT").value)
	If Month(dRepStart) = Month(dCreatedDate) And Year(dRepStart) = Year(dCreatedDate) Then
		lFirstMonth = True
		nFreeMonthlyUnits = 0
	Else
		lFirstMonth = False
		nFreeMonthlyUnits = 20
	End If

	cAHS = oRS.Fields("ACCNT_HRCY_STEP_ID").value
	'   get total number of claims (excluding INF calls)
    cSQL = "Select COUNT(call.call_id) AS totalCalls " & _
			"From CALL Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
			"AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
			"AND STATUS = 'COMPLETED' " & _
			"AND LOB_CD <> 'INF' " & _
			"AND ACCNT_HRCY_STEP_ID = " & cAHS
	set oRS1 = oConn.Execute(cSQL)
	nTotalUnits4Claims = clng(oRS1.Fields("totalCalls").value) * 4
	'   get INF claims
	cSQL = "Select count(*) as nTotal From CALL " & _
			"Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
			"AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
			"AND STATUS = 'COMPLETED' " & _
			"AND LOB_CD = 'INF' " & _
			"AND ACCNT_HRCY_STEP_ID = " & cAHS
	set oRS1 = oConn.Execute(cSQL)
	nTotalUnits4INF = clng(oRS1.Fields("nTotal").value)
	'   get escalations
	cSQL = "Select count(*) as nTotal From CALL, ESCALATION_OUTCOME " & _
			"Where CALL_START_TIME >= TO_DATE('" & cStart & "') " & _
			"AND CALL_START_TIME < TO_DATE('" & cEnd & "') " & _
			"AND CALL.STATUS = 'COMPLETED' " & _
            "AND CALL.CALL_ID = ESCALATION_OUTCOME.CALL_ID " & _
            "AND CALL.ACCNT_HRCY_STEP_ID = " & cAHS
	set oRS1 = oConn.Execute(cSQL)
	nTotalEscalations = clng(oRS1.Fields("nTotal").value) * 7.5
	'
	nTotalDollars = 0.0
	If cServType = "Y" Then
		'	****************** YEARLY PLAN **************************
		cServType = "Yearly"
		nTotalUnits = (nTotalUnits4INF + nTotalUnits4Claims)
		if nTotalUnits - nFreeMonthlyUnits > 0 then
			nTotalDollars = cdbl((nTotalUnits - nFreeMonthlyUnits) * 3.96)
		end if
	elseif cServType = "M" Then
		cServType = "Monthly"
		nTotalUnits = nTotalUnits4INF + nTotalUnits4Claims
		if nTotalUnits > 0 then
			nTotalDollars = cdbl(nTotalUnits * 4.25)
		end if
	end if
	nTotalDollars = nTotalDollars + nTotalEscalations
	nGrandTotal = nGrandTotal + nTotalDollars
	cValues =	"'" & cAHS & "','"
	if Month(dRepStart) < 10 then
		cMonth = "0" & Month(dRepStart)
	else
		cMonth = Month(dRepStart)
	end if
	cValues = cValues & cAHS & cMonth & Right(cstr(Year(dRepStart)), 2) & "','" & _
				cAgentName & "','" & _
				FormatNumber(nTotalEscalations) & "','" & _
				FormatNumber(nTotalUnits) & "','" & _
				cServType & "','" & _
				cPaymentMethod & "','" & _
				FormatNumber(nTotalDollars) & "'"
	oExcel.addRow cFields, cValues
	oRS.MoveNext
loop
cValues = "'','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
with oExcel
	.closeXLS
	.sendFile
end with
Set oRS = Nothing
end sub

'CISG*****************************
'****************************************************************************************************************

function genCSG
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee ,cAccountName ,nAccountName
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent ,pConfirmationFlg ,cConfirmationFlg
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with

 if cCustCode = "CSG" then
     confFlag = "N"
 elseif cCustCode = "CSGL" then
      confFlag = "Y"
 end if
'
cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' " & _
			" AND BILLING_DETAIL.CONFIRM_FLG='" & confFlag & "') " & _
			"Order by CLIENT_NAME,ACCOUNT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"



set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		'if isNull(.Fields("PARENT_NAME").Value) then
			'cParent = ""
		'else
			'cParent = .Fields("PARENT_NAME").Value
		'end if
	if	isNull(.Fields("ACCOUNT_NAME").Value)  then
	      cAccountName = ""
	else
	     cAccountName = .Fields("ACCOUNT_NAME").Value
	 end if

		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
 do while Not .EOF
			'if isNull(.Fields("PARENT_NAME").Value) then
				'cCmpParent = ""
			'else
				'cCmpParent = .Fields("PARENT_NAME").Value
			'end if

		     if	isNull(.Fields("ACCOUNT_NAME").Value)  then
	              nAccountName = ""
	          else
	              nAccountName = .Fields("ACCOUNT_NAME").Value
	        end if

     ' if (cParent = cCmpParent) or (cAccountName = nAccountName)  then
      if cAccountName = nAccountName  then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					  if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					  else
						cValues = cValues & .Fields("LOSS_DATE").Value
					  end if

					 if err.number <> 0 then
						  if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						  end if
						 lErrorTriggered = true
						 exit do
					  end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
		 else
			 exit do
	    end if

  loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
 Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genCSG = lErrorTriggered
end function

' *********************************************************************
'	AIK
' *********************************************************************
function genAIK
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'
cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)

cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingAIK.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Site_ID,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "',"
					cValues = cValues & .Fields("SITE_ID").Value & ",'" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genAIK = lErrorTriggered
end function
'*************************** genRDC *********************************
function genRDC
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'
cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD,INSURED_NAME, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("INSURED_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genRDC = lErrorTriggered
end function

'*********************************** WMA ***Waste mnegement********************************
function genWMA
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
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
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingWMA.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
                           '------  for print  use the DRIVEN_FEE column in table Billing_details
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Print,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0


		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nPrint = CSng(.Fields("DRIVEN_FEE").Value) 'use the existing column

				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrint
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee,2) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(nPrint) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrint) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genWMA = lErrorTriggered
end function
'**********************************CRAWFORD*********************************************************8*****************************************8*********
function genCRW
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim nClientFlag

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'                     'if "Crawford ASP Claims"  siteId = 5 then claim was taken ASP Crawford client
                      'if "Crawford After Hours" everything else  it is claim was taken after hours from FNS
                          cCustName = Request.QueryString("CUSTNAME")
                       if cCustName = "Crawford ASP Claims" then
                           nClientFlag  = "A"
                        else
                           nClientFlag  = "F"
                       end if


      dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'
               cSQL = "SELECT BILLING_DETAIL.* " & _
			   "FROM BILLING_DETAIL " & _
			   "WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			   " AND BILLING_DETAIL.STATUS='ACTIVE'" & _
			   " AND CONFIRM_FLG ='" & nClientFlag & "') " & _
			   "Order by BRANCH_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"

set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingCRW.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,LocationCode,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,Coverage_Code,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("BRANCH_NAME").Value) then
			cParent = ""
	     else
			cParent = .Fields("BRANCH_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("BRANCH_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("BRANCH_NAME").Value
			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				'nTempFee = CSng(.Fields("TEMP_FEE").Value)
				'nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				'nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee '
				 cValues = "'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "','" &  .Fields("BRANCH_NAME").Value  & "'," & _
				      CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing

genCRW = lErrorTriggered
end function
'

'****************************************** KEMP***************************
function genKMP
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,cSQL2,oRS1,oRS2
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls
dim nTotSvcFee
dim nTotAdnFee
dim nTotalFee
dim nTotalFaxFee
dim nVendorFee
dim nEscalateFee
dim nTempFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval,serval1,serval2

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0.0
nTotalNoCalls = 0
'
   cCustName = Request.QueryString("CUSTNAME")

  dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'
cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD,CONFIRM_FLG, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'response.write(cSQL)

cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"
with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingKMP.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left( ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Cat_Flag,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0.0
		nTotalFee = 0.0
		nVendorFee=0.0
		do while Not .EOF
		    nVendorFee=0.0
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
			    if (.Fields("CONFIRM_FLG").Value)="N" then
				nCalls = nCalls + 1
				end if



				if (.Fields("CALL_TYPE").Value)	= "I" then
				serval=1.55
				else

				serval=CSng(.Fields("SERVICE_FEE").Value)
				end if
				if ((.Fields("Call_Type").Value)="F") then
				nTotalFaxFee = serval
				'nTotSvcFee=0
				else
				nTotalFaxFee=0
				nTotSvcFee = nTotSvcFee + serval
				end if
				'nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				if (.Fields("CONFIRM_FLG").Value)="N" then
				if (.Fields("LOB_CD").Value)="PAU"  then
				cSQL1 = "Select  count(*) as nTotalPAU From CALL C, CALL_CLAIM CC, CALL_VENDOR_REFERRAL CVR " & _
                       " Where C.CALL_ID = CC.CALL_ID " & _
                       " AND CC.CALL_CLAIM_ID = CVR.CALL_CLAIM_ID " & _
                       " AND CVR.REFERRAL_ACCEPTED = 'Y' " & _
                       " AND C.CALL_ID=" & .Fields("Call_ID").Value &_
                       " AND C.STATUS = 'COMPLETED'" & _
                       " AND  C.LOB_CD = 'PAU'" & _
                       " AND C.CLIENT_HRCY_STEP_ID = " & cAHS



               'if (trim(.Fields("Call_Type").Value)="F" ) then
			          set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						do while Not .EOF
							 serval1 = .Fields(0).Value
							 'response.Write(serval1)
							 'response.end
						oRS1.moveNext
					    loop
							oRS1.close
						end with
				if (Csng(serval1)=1) then
				nVendorFee=8
				else
				nVendorFee=0
				end if
				end if

				if (.Fields("LOB_CD").Value)<>"PAU" then
				cSQL2=" Select count(*) as nTotal From CALL C, CALL_CLAIM CC, CALL_ASI CASI, X_CALL_ASI XCASI " & _
                      " Where  C.CALL_ID = CC.CALL_ID " & _
                      " AND C.CALL_ID = CASI.CALL_ID " & _
                      " AND CASI.CALL_ASI_ID = XCASI.CALL_ASI_ID " & _
                      " AND XCASI.FIELD LIKE 'ACCEPTED_MITIGATION_FLG%' " & _
                      " AND C.STATUS = 'COMPLETED' " & _
                      " AND not C.LOB_CD = 'PAU'" & _
                      " AND C.CALL_ID=" & .Fields("Call_ID").Value &_
                      " AND XCASI.VALUE ='Y'" &_
                      " AND C.CLIENT_HRCY_STEP_ID = " & cAHS


              'response.Write(cSQL2)

               'if (trim(.Fields("Call_Type").Value)="F" ) then
			          set oRS2 = Conn.Execute(cSQL2)
						with oRS2
						do while Not .EOF
							 serval2 = .Fields(0).Value
						oRS2.moveNext
					    loop
							oRS2.close
						end with
				if (Csng(serval2)=1)then
				nVendorFee=5
				else
				nVendorFee=0
				end if
				end if

				end if

				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				'nVendorFee = CSng(.Fields("VENDOR_FEE").Value)

				'nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				nTotalFee = nTotalFee + serval + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & .Fields("CONFIRM_FLG").Value & "','"
					if (trim(.Fields("Call_Type").Value)="F" ) then
					cValues = cValues & FormatNumber(0) & "','"
						'serval1=serval
					else
					cValues = cValues & FormatNumber(serval) & "','"
						'serval=.Fields("SERVICE_FEE").Value
						'FormatNumber(.Fields("SERVICE_FEE").Value) & "','" &
					end if

					cValues = cValues & FormatNumber(nTotalFaxFee) & "','" & _

					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(serval + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
			'.Fields("CONFIRM_FLG").Value & "','"
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genKMP = lErrorTriggered
end function

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&888888 CIR cONECTUCA &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
function genCIR
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'




dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
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
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingWMA.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
                           '------  for print  use the DRIVEN_FEE column in table Billing_details
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Print,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nPrint = CSng(.Fields("DRIVEN_FEE").Value) 'use the existing column

				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrint
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(nPrint) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrint) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genCIR = lErrorTriggered
end function

' *********************************************************************
'	ARG
' *********************************************************************
function genARG
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim cBranchName, cCmpBranch
dim nCallsByBranch, nTotalByBranch


cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'


dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)


with oRS
	cBillID = .fields(0)
	.close
end with
'
cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE') " & _
			"Order by BRANCH_NAME, PARENT_NAME, CLIENT_NAME, CALL_TYPE"


set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingARG.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	 'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	 cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)

'---------------------------------------------------------------------------------------------------------------------->
'-- BEGIN		{JPRI-0851}		CHANGE# 1
'				Following line commented and next line added to add a new column
'				- Account and Brach fields get interchanged their respoective positions.
'---------------------------------------------------------------------------------------------------------------------->
'cFields = "Account,Branch,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"

'---------------------------------------------------------------------------------------------------------------------->
'-- END			{JPRI-0851}
'---------------------------------------------------------------------------------------------------------------------->

'---------------------------------------------------------------------------------------------------------------------->
'-- BEGIN		{JPRI-0919}		CHANGE# 1
'				Following line commented and next line added to add a new column
'				- ACCOUNT NAME - to the existing Billing Report for ARG
'---------------------------------------------------------------------------------------------------------------------->
'cFields = "Branch,Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"

cFields = "Account,Branch,Account_Number,Call_No,T,Account_Name,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
'---------------------------------------------------------------------------------------------------------------------->
'-- END			{JPRI-0919}
'---------------------------------------------------------------------------------------------------------------------->
with oRS
		do while not .eof
			if lErrorTriggered then
				exit do
			end if
				' brake by branch
				if isNull(.Fields("BRANCH_NAME").Value) then
					cBranchName = ""
				else
					cBranchName = .Fields("BRANCH_NAME").Value
				end if
				nCallsByBranch = 0
				nTotalByBranch = 0
				do while not .eof
						if isNull(.Fields("BRANCH_NAME").Value) then
							cCmpBranch = ""
						else
							cCmpBranch = .Fields("BRANCH_NAME").Value
						end	if
						if cBranchName = cCmpBranch then
							nCallsByBranch = nCallsByBranch + 1
							nCalls = nCalls + 1
							nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
							nTempFee = CSng(.Fields("TEMP_FEE").Value)
							nEscalateFee = CSng(.Fields("ESCALATE_FEE").Value)
							nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
							nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
							nFeeSum = CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
							nTotalByBranch = nTotalByBranch + nFeeSum
							nTotalFee = nTotalFee + nFeeSum

							'---------------------------------------------------------------------------------------------------------------------->
							'-- BEGIN		{JPRI-0851}			CHANGE# 2
							'				PARENT_NAME and BRANCH_NAME change their positions in cValues field.
							'Previous values are commented out below:
						    '
							'---------------------------------------------------------------------------------------------------------------------->
							'cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "','" & removeSngQuote(.Fields("BRANCH_NAME").Value) & "'," & _
							'	CStr(.Fields("Call_ID").Value) & ",'" & _
							'	.Fields("Call_Type").Value & "','" & _
							'	.Fields("CALLSTATUS").Value & "','"
							'---------------------------------------------------------------------------------------------------------------------->
							'-- END		{JPRI-0851}
							'---------------------------------------------------------------------------------------------------------------------->
							'---------------------------------------------------------------------------------------------------------------------->
							'-- BEGIN		{JPRI-0919}			CHANGE# 2
							'				Following lines commented and next few lines added to add a new column
							'				- ACCOUNT NAME - to the existing Billing Report for ARG
							'---------------------------------------------------------------------------------------------------------------------->

							'cValues =	"'" & removeSngQuote(.Fields("BRANCH_NAME").Value) & "','" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
							'	CStr(.Fields("Call_ID").Value) & ",'" & _
							'	.Fields("Call_Type").Value & "','" & _
							'	.Fields("CALLSTATUS").Value & "','"

							cValues =	"'" & removeSngQuote(.Fields("BRANCH_NAME").Value) & "','" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
								CStr(.Fields("Call_ID").Value) & ",'" & _
								.Fields("Call_Type").Value & "','"

							if isNull(.Fields("PARENT_NODE_ID").Value) then
								cValues = cValues & "','"
							else
								cValues = cValues & .Fields("PARENT_NODE_ID").Value & "','"
							end if

							if isNull(.Fields("ACCOUNT_NAME").Value) then
								cValues = cValues & ""
							else
								cValues =	cValues & removeSngQuote(.Fields("ACCOUNT_NAME").Value)
							end if

							cValues =	cValues & "','" & .Fields("CALLSTATUS").Value & "','"

							'---------------------------------------------------------------------------------------------------------------------->
							'-- END		{JPRI-0919}
							'---------------------------------------------------------------------------------------------------------------------->

							on error resume next
							if isNull(.Fields("LOSS_DATE").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("LOSS_DATE").Value
							end if

							if err.number <> 0 then
								if err.number = -2147217887 then
									writeError .Fields("Call_ID").Value
								end if
								lErrorTriggered = true
								exit do
							end if

							cValues = cValues & "','" & _
								CStr(.Fields("CALL_END_TIME").Value) & "','"
							if isNull(.Fields("CLAIM_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & .Fields("CLAIM_NUMBER").Value
							end if
							cValues = cValues & "','" & _
								.Fields("LOB_CD").Value & "','"
							if isNull(.Fields("POLICY_NUMBER").Value) then
								cValues = cValues & ""
							else
								cValues = cValues & removeSngQuote(.Fields("POLICY_NUMBER").Value)
							end if
							cValues = cValues & "','" & _
								removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
							aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','"
							redim aNameParts(1)
							aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
							aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
							cValues = cValues & getName(aNameParts) & "','"
							cValues = cValues & cEmployeeSSN & "','" & _
								FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
								FormatNumber(nTotalFaxFee) & "','" & _
								FormatNumber(nTempFee) & "','" & _
								FormatNumber(nEscalateFee) & "','" & _
								FormatNumber(nVendorFee) & "','" & _
								FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
							oExcel.addRow cFields, cValues
							if err.number <> 0 then
								writeErrText
								lErrorTriggered = true
								exit do
							end if
							on error goto 0
							.MoveNext
						else
							exit do
						end if
				loop
				if not lErrorTriggered then
					'---------------------------------------------------------------------------------------------------------------------->
					'-- BEGIN		{JPRI-0919}			CHANGE# 3
					'---------------------------------------------------------------------------------------------------------------------->
					'cValues = "'Branch SubTotal'," & nCallsByBranch & ",'','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotalByBranch) & "'"
					cValues = "'Branch SubTotal'," & nCallsByBranch & ",'','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotalByBranch) & "'"
					'---------------------------------------------------------------------------------------------------------------------->
					'-- END		{JPRI-0919}
					'---------------------------------------------------------------------------------------------------------------------->
					oExcel.addRow cFields, cValues
					'---------------------------------------------------------------------------------------------------------------------->
					'-- BEGIN		{JPRI-0919}			CHANGE# 4
					'---------------------------------------------------------------------------------------------------------------------->
					'cValues = "'','','','','','','','','','','','','','','','','','','',''"
					cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
					'---------------------------------------------------------------------------------------------------------------------->
					'-- END		{JPRI-0919}
					'---------------------------------------------------------------------------------------------------------------------->
					oExcel.addRow cFields, cValues
				end if
		loop
	.Close
end with
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genARG = lErrorTriggered
end function

function genHML
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
nCalls = 0
nTotSvcFee = 0
nTotalFee = 0
'
cCustName = Request.QueryString("CUSTNAME")

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
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
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("CLIENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
	Loop
	.Close
end with

cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
oExcel.addRow cFields, cValues

oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genHML = lErrorTriggered
end function
'******
' Added for UniSource April 17 2006
'******
function genUniXLS
dim cAHS, cStartDate, cSP, oRS, cSQL
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls




cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
STOP
       if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
          cCustName = "Managed Care Innovations."
       else
          cCustName = Request.QueryString("CUSTNAME")
       end if

dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with

cSQL = " select * "
cSQL = cSQL &  " from ( "
cSQL = cSQL &  " SELECT "
cSQL =cSQL &  "   bd.*,b.BRANCH_OFFICE_NAME"
cSQL = cSQL &  " FROM "
cSQL = cSQL &  "   BILLING_DETAIL bd,call_Branch b,call_claim c"
cSQL = cSQL &  " WHERE"
cSQL = cSQL &  "    c.call_claim_id=b.call_claim_id(+)"
cSQL = cSQL &  "    and c.CALL_ID=bd.CALL_ID "
cSQL = cSQL &  "   And (bd.BILLING_ID = "& cBillID &" AND"
cSQL = cSQL &  "   bd.STATUS='ACTIVE' AND"
cSQL = cSQL &  "   bd.CALLSTATUS = 'COMPLETED'))"
cSQL = cSQL &  "   Order by LOB_CD desc,BRANCH_OFFICE_NAME,PARENT_NAME,CLIENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"



set oRS = Conn.Execute(cSQL)
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingAMC.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Branch_Office_Name,Account,Call_No,Type,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"

with oRS
	Do While Not .EOF

		if isNull(.Fields("branch_office_name").Value) then
			cParent = ""
		else
			cParent = .Fields("branch_office_name").Value

		end if
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		do while Not .EOF
			if isNull(.Fields("branch_office_name").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("branch_office_name").Value

			end if
			if cParent = cCmpParent then
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
				nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues = "'" & .Fields("BRANCH_OFFICE_NAME").Value & "','"
				cValues =	cValues & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','" & _
					FormatNumber(.Fields("SERVICE_FEE").Value) & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"
				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then

			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"

			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else

			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"

oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genUNIXLS = lErrorTriggered
end function

'......................................................

'*******************************'************************************************************'********************
'ALM
'********************************
function genALMXLS
	dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
	dim dRepDate, cBillID, cCustName, cCustCode
	dim cFields, cValues, cTime
	dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
	dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee,nTransmissionfees
	dim cCmpParent
	dim lErrorTriggered
	dim nGrandTotal
	dim nTotalNoCalls
	dim strTempLOB,strBillingDetailsID,strLocation,serval,serval1
	serval="0"
	serval1="0"

	cAHS = Request.QueryString("AHS")
	cStartDate = Request.QueryString("DATEFROM")
	cCustCode = Request.QueryString("CUSTCODE")
	cCustName = Request.QueryString("CUSTNAME")

	lErrorTriggered = false
	nGrandTotal = 0
	nTotalNoCalls = 0
	'
		if  cAHS = "12801314" and cCustName = "COMMONWEALTH OF VIRGINIA" then
			cCustName = "Managed Care Innovations."
		else
			cCustName = "HANOVER"'Request.QueryString("CUSTNAME")
		end if

	dRepDate = cDate(cStartDate)
	cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
	set oRS = Conn.Execute(cSQL)
	with oRS
		cBillID = .fields(0)
		.close
	end with

	cTime = CStr(FormatDateTime( Time, vbShortTime))
	cTime = Replace(cTime, ":", "")
	cTmpFile = cCustCode & "-" & cTime & ".xls"

	cSQL = "SELECT b.*,cb.BRANCH_OFFICE_NAME,cc.call_claim_id " & _
				"FROM BILLING_DETAIL b, call_claim cc, call_branch cb " & _
				"WHERE (b.BILLING_ID = " & cBillID & _
				" AND b.STATUS='ACTIVE' " & _
				" AND b.CALLSTATUS = 'COMPLETED' and b.call_id = cc.call_id  " & _
				"AND cc.CALL_CLAIM_ID = cb.CALL_CLAIM_ID) " &_
				"Order by b.CLIENT_NAME,b.PARENT_NAME, b.CALL_TYPE,b.LOB_CD, b.CALL_END_TIME"
	set oRS = Conn.Execute(cSQL)
	'
	cTime = CStr(FormatDateTime( Time, vbShortTime))
	cTime = Replace(cTime, ":", "")
	cTmpFile = cCustCode & "-" & cTime & ".xls"

	with oExcel
		.cDestinationFileName = cTmpFile
		.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingALM.xls"
		.cExcelRangeName = "ODBCRange"
		.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
		'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
		cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
		.cDownloadLocation = cDownloadLocation
		.openXLS
		.writeMsg "Generating spreadsheet"
		.writeCell "Account", cCustName
		.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
		'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
	end with
	writePeriod(oExcel)
	cFields = "Account,Branch_Office_Name,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Loss_State,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Print,Total"
	with oRS
		Do While Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cParent = ""
			else
				cParent = .Fields("PARENT_NAME").Value
			end if
			nCalls = 0
			nTotSvcFee = 0
			nTotalFee = 0

			do while Not .EOF
				if isNull(.Fields("PARENT_NAME").Value) then
					cCmpParent = ""
				else
					cCmpParent = .Fields("PARENT_NAME").Value
				end if


				if cParent = cCmpParent then
					nCalls = nCalls + 1
					cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND CALL_TYPE = 'F'" &_
							  " AND FEE_TYPE_ID='2'"


				if isNull(.Fields("Branch_Office_Name").Value) then
					cCmpParent = ""
				else
					cCmpParent = .Fields("Branch_Office_Name").Value
				end if
					'response.end
					  ' end if
					'response.Write(.Fields("Call_Type").Value)
					'response.end
					if (trim(.Fields("Call_Type").Value)="F" ) then
			          set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						do while Not .EOF
							 serval = .Fields("FEE_AMOUNT").Value
						oRS1.moveNext
					    loop
							oRS1.close
						end with
						nTotalFaxFee = CSng(serval)
						nTotSvcFee=nTotSvcFee+CSng(serval)
						else
						nTotalFaxFee =  CSng(.Fields("TOTAL_FAX_FEE").Value)
						nTotSvcFee = nTotSvcFee + CSng(.Fields("SERVICE_FEE").Value)
						'response.Write(.Fields("SERVICE_FEE").Value)
						'response.end
					end if

					nTempFee = CSng(.Fields("TEMP_FEE").Value)
					nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
					nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
					nPrintFee=Csng(.Fields("DRIVEN_FEE").Value)
					nTotalFee = nTotalFee + CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee+nPrintFee
					cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
								"'" & removeSngQuote(.Fields("Branch_Office_Name").Value) & "',"  &_
						CStr(.Fields("Call_ID").Value) & ",'" & _
						.Fields("Call_Type").Value & "','" & _
						.Fields("CALLSTATUS").Value & "','"
						'on error resume next
						if isNull(.Fields("LOSS_DATE").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("LOSS_DATE").Value
						end if


						if err.number <> 0 then
							if err.number = -2147217887 then
								writeError .Fields("Call_ID").Value
							end if
							lErrorTriggered = true
							exit do
						end if

						cValues = cValues & "','" & _
						CStr(.Fields("CALL_END_TIME").Value) & "','"
						if isNull(.Fields("CLAIM_NUMBER").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("CLAIM_NUMBER").Value
						end if
						cValues = cValues & "','" & _
						.Fields("LOB_CD").Value & "','"
						if isNull(.Fields("POLICY_NUMBER").Value) then
							cValues = cValues & ""
						else
							cValues = cValues & .Fields("POLICY_NUMBER").Value
						end if
						cValues = cValues & "','" & _
						removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"

						'*****************************
						' Modified on 25/10/2005

						strTempLOB=.Fields("LOB_CD").Value
						strBillingDetailsID=.Fields("BILLING_DETAIL_ID").Value
						if(strTempLOB="WOR")then
							cSQL = "SELECT BENEFIT_STATE FROM CALL_CLAIM CLM,BILLING_DETAIL BDL"
							cSQL = cSQL & " WHERE BDL.CALL_ID= CLM.CALL_ID"
							cSQL = cSQL & " AND BDL.CLIENT_NODE_ID=650"
							cSQL = cSQL & " AND BDL.BILLING_DETAIL_ID=" & strBillingDetailsID
						end if
						if(strTempLOB="CAU" or strTempLOB="CPR" or strTempLOB="CLI" or strTempLOB="PAU" ) then

							cSQL = "SELECT ADDRESS_STATE FROM CALL_LOSS_LOCATION cll ,CALL_CLAIM cml ,BILLING_DETAIL          bdl"
							cSQL = cSQL & " WHERE cll.CALL_CLAIM_ID=cml.CALL_CLAIM_ID"
							cSQL = cSQL & " AND cml.CALL_ID = bdl.CALL_ID"
							cSQL = cSQL & " AND bdl.CLIENT_NODE_ID = 650"
							cSQL = cSQL & " AND bdl.BILLING_DETAIL_ID =" & strBillingDetailsID
						end if

                        if (strTempLOB="INF") then
							cSQL = " (select decode(upper(x_call_asi.FIELD),'INSURED_STATE',x_call_asi.VALUE) as loss_state "
							cSQL = cSQL & " from x_call_asi,call_asi,billing_detail "
							cSQL = cSQL & " where call_asi.call_asi_id = x_call_asi.call_asi_id"
							cSQL = cSQL & " AND call_asi.CALL_ID = billing_detail.CALL_ID"
							cSQL = cSQL & " AND billing_detail.CLIENT_NODE_ID = 650"
							cSQL = cSQL & " AND billing_detail.BILLING_DETAIL_ID = " & strBillingDetailsID
							cSQL = cSQL & " AND x_call_asi.FIELD = 'INSURED_STATE')"
						end if

						set oRS = Conn.Execute(cSQL)
						with oRS
						Do while not .EOF
							strLocation = .fields(0)
						.moveNext
						loop
						.close
						end with
						cValues = cValues & strLocation & "','"
						'******************************
						redim aNameParts(1)
						aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
						aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
						cValues = cValues & getName(aNameParts) & "','"
						redim aNameParts(1)
						aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
						aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
						cValues = cValues & getName(aNameParts) & "','"
						cValues = cValues & cEmployeeSSN & "','"

						if (trim(.Fields("Call_Type").Value)="F" ) then
						cValues = cValues & FormatNumber(CSng(serval)) & "','"
						'serval1=serval
						else
						cValues = cValues & FormatNumber(.Fields("SERVICE_FEE").Value) & "','"
						'serval=.Fields("SERVICE_FEE").Value
						'FormatNumber(.Fields("SERVICE_FEE").Value) & "','" &
						end if

						cValues = cValues & FormatNumber(nTotalFaxFee) & "','" & _
						FormatNumber(nTempFee) & "','" & _
						FormatNumber(nEscalateFee) & "','" & _
						FormatNumber(nVendorFee) & "','" & _
						FormatNumber(nPrintFee) & "','"  &_
						FormatNumber(CSng(.Fields("SERVICE_FEE").Value) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee + nPrintFee) & "'"

					oExcel.addRow cFields, cValues
					.MoveNext
				else
					exit do
				end if
			loop
			if not lErrorTriggered then
				cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','','" & FormatCurrency(nTotalFee) & "'"
				oExcel.addRow cFields, cValues
				cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
				oExcel.addRow cFields, cValues
			else
				exit do
			end if
			nGrandTotal = nGrandTotal + nTotalFee
			nTotalNoCalls = nTotalNoCalls + nCalls
		Loop
		.Close
	end with
	cValues = "'','','','','','','','','','','','','','','','','','','','','',''"
	oExcel.addRow cFields, cValues
	cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
	oExcel.addRow cFields, cValues
	oExcel.closeXLS
	if not lErrorTriggered then
		oExcel.sendFile
	end if
	Set oRS = Nothing
	genALMXLS = lErrorTriggered
end function
'*******************************
'END ALM
'*******************************


'*******************************
' AME
'*******************************
function genAMEXLS
	dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
	dim dRepDate, cBillID, cCustName, cCustCode,Division
	dim cFields, cValues, cTime
	dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
	dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
	dim cCmpParent
	dim lErrorTriggered
	dim nGrandTotal
	dim nTotalNoCalls
	dim serval

	serval="0"

	cAHS = Request.QueryString("AHS")
	cStartDate = Request.QueryString("DATEFROM")
	cCustCode = Request.QueryString("CUSTCODE")
	cCustName = Request.QueryString("CUSTNAME")
	lErrorTriggered = false
	nGrandTotal = 0
	nTotalNoCalls = 0

	dRepDate = cDate(cStartDate)
	cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
	set oRS = Conn.Execute(cSQL)
	with oRS
		cBillID = .fields(0)
		.close
	end with

	cSQL = "SELECT distinct b.*, u.name as user_name, c.call_start_time as CallStartTime, c.call_end_time as CallEndTime, " & _
			"cc.TEMPEDPOLICY_FLG, cc.claim_type as ClaimType, " & _
			"cca.caller_Type as Input_Type, " & _
			"cb.BRANCH_NUMBER, cb.BRANCH_OFFICE_NAME, cll.ADDRESS_STATE  as LossState " & _
			"FROM BILLING_DETAIL b, call c, call_caller cca, call_claim cc, users_view u, call_loss_location cll, " & _
			"(select cc.call_id,cb.branch_number,cb.branch_office_name " & _
			"from call_claim cc,call_branch cb " & _
			"where cc.call_claim_id = cb.call_claim_id(+) )cb " & _
			"WHERE (b.BILLING_ID = " & cBillID & _
			" AND b.STATUS = 'ACTIVE' AND b.CALLSTATUS = 'COMPLETED' and b.call_id = c.call_id AND c.CALL_ID = cca.CALL_ID(+) " & _
			" AND c.CALL_ID = cc.CALL_ID AND c.user_id = u.user_id AND cc.CALL_ID = cb.CALL_ID(+)) AND cc.CALL_CLAIM_ID = cll.CALL_CLAIM_ID(+) " & _
			"Order by b.CLIENT_NAME,b.PARENT_NAME,b.CALL_TYPE,b.LOB_CD,b.CALL_END_TIME"

	set oRS = Conn.Execute(cSQL)

	cTime = CStr(FormatDateTime( Time, vbShortTime))
	cTime = Replace(cTime, ":", "")
	cTmpFile = cCustCode & "-" & cTime & ".xls"

	with oExcel
		.cDestinationFileName = cTmpFile
		.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\BillingAME.xls"
		.cExcelRangeName = "ODBCRange"
		.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
		cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
		.cDownloadLocation = cDownloadLocation
		.openXLS
		.writeMsg "Generating spreadsheet"
		.writeCell "Account", cCustName
		.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	end with
	writePeriod(oExcel)

	cFields = "Account,Rep,Claim_Type,Input_Type,Branch_Number,Branch_Name,Call_No,T,Status,Division,Loss_Dt,Loss_State,Call_Start_Dt,Call_End_Dt,Duration,Claim_No,LOB,Policy_No,Employee_Insured,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"

	with oRS
		Do While Not .EOF

			if isNull(.Fields("PARENT_NAME").Value) then
				cParent = ""
			else
				cParent = .Fields("PARENT_NAME").Value
			end if
			nCalls = 0
			nTotSvcFee = 0
			nTotalFee = 0
			serval=0

			Do While Not .EOF

				if isNull(.Fields("PARENT_NAME").Value) then
					cCmpParent = ""
				else
					cCmpParent = .Fields("PARENT_NAME").Value
				end if

				if cParent = cCmpParent then
					serval="0"
					nCalls = nCalls + 1
					nTotalFaxFee = CDbl(.Fields("TOTAL_FAX_FEE").Value)
					nTempFee = CDbl(.Fields("TEMP_FEE").Value)
					nEscalateFee = CDbl(.Fields("ESCALATE_FEE").Value)
					nVendorFee = CDbl(.Fields("VENDOR_FEE").Value)
					serval = CDbl(.Fields("SERVICE_FEE").Value)

					if (.Fields("Input_Type").Value = "EML" or .Fields("Input_Type").Value = "EMAIL") then
						'if (.Fields("LOB_CD").Value <> "INF") then
								cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
										" WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
										" AND CALL_TYPE = 'E' "
							set oRS1 = Conn.Execute(cSQL1)
							with oRS1
								Do While Not oRS1.EOF
								serval = .fields(0)
								oRS1.MoveNext
								loop
								.close
							end with
						'end if
					end if

					if (.Fields("Input_Type").Value = "NET" or .Fields("Input_Type").Value = "INT" ) then
						'if (.Fields("LOB_CD").Value <> "INF") then
							cSQL2 = " Select FEE_AMOUNT FROM FEE " & _
							" WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							" AND CALL_TYPE = 'N' "

							set oRS2 = Conn.Execute(cSQL2)
							with oRS2
								Do While Not oRS2.EOF
								serval = .fields(0)
								oRS2.MoveNext
								loop
								.close
							end with
						'end if
					end if

					'Get Divison Value
					Division = "Ameriprise"

					'GetProgramType
					dim ProgramType

					'Get account name
					dim Call_Status
					dim DivisionName
					DivisionName=division

					Call_Status = .Fields("TEMPEDPOLICY_FLG").Value

					dim Duration
					Duration = Round(DateDiff("s",.Fields("CallStartTime").Value,.Fields("CallEndTime").Value) / 60)

					dim ClaimTypeFlg, ClaimTypeDesc
					ClaimTypeFlg = ""
					ClaimTypeDesc = ""

					ClaimTypeFlg = .Fields("ClaimType").Value

					if ClaimTypeFlg = "R" then
						ClaimTypeDesc = "Record Only"
					else
						ClaimTypeDesc = "Claim"
					end if

					nTotSvcFee=nTotSvcFee + CDbl(serval)
					nTotalFee = nTotalFee + CDbl(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
					cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "','" & _
								removeSngQuote(.Fields("USER_NAME").Value) & "','" & _
								ClaimTypeDesc & "','" & _
								.Fields("INPUT_TYPE").Value & "','"

					if isNull(.Fields("BRANCH_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("BRANCH_NUMBER").Value
					end if

					cValues = cValues & "','"

					if isNull(.Fields("BRANCH_OFFICE_NAME").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("BRANCH_OFFICE_NAME").Value
					end if

					cValues =	cValues & "','" & _
								.Fields("CALL_ID").Value & "','" & _
								.Fields("Call_Type").Value & "','" & _
								.Fields("CALLSTATUS").Value & "','" & _
								DivisionName & "','"

					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if

					cValues =	cValues & "','" & _
								.Fields("LossState").Value & "','" & _
								.Fields("CallStartTime").Value & "','" & _
								.Fields("CallEndTime").Value & "','" & _
								Duration & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if

					cValues =	cValues & "','" & _
								.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("POLICY_NUMBER").Value
					end if

					cValues =	cValues & "','" & _
								removeSngQuote(.Fields("INSURED_NAME").Value) & "','" & _
								removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues =	cValues & serval & "','" & _
								FormatNumber(nTotalFaxFee) & "','" & _
								FormatNumber(nTempFee) & "','" & _
								FormatNumber(nEscalateFee) & "','" & _
								FormatNumber(nVendorFee) & "','" & _
								FormatNumber(CDbl(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

					oExcel.addRow cFields, cValues
					.MoveNext
				else
					exit do
				end if
			loop

			if not lErrorTriggered then
				cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
				oExcel.addRow cFields, cValues
				cValues = "'','','','','','','','','','','','','','','','','','','','','','','','','','','','',''"
				oExcel.addRow cFields, cValues
			else
				exit do
			end if
			nGrandTotal = nGrandTotal + nTotalFee
			nTotalNoCalls = nTotalNoCalls + nCalls
		Loop
		.Close
	end with
	cValues = "'','','','','','','','','','','','','','','','','','','','','','','','','','','','',''"
	oExcel.addRow cFields, cValues
	cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
	oExcel.addRow cFields, cValues
	oExcel.closeXLS
	if not lErrorTriggered then
		oExcel.sendFile
	end if
	Set oRS = Nothing
	genAMEXLS = lErrorTriggered
end function
'*******************************
' END AME
'*******************************

'*****************************************************
' TPAL-0146 Tower group billing fees / reports setup
'*****************************************************

function genTOWASPXLS
dim cAHS, cStartDate, cSP, oRS, cSQL,cSQL1,oRS1
dim dRepDate, cBillID, cCustName, cCustCode
dim cFields, cValues, cTime
dim aNameParts(), cParent, nCalls, nTotSvcFee, nTotAdnFee, nTotalFee
dim nTotalFaxFee, nTempFee, nEscalateFee, nVendorFee
dim cCmpParent
dim lErrorTriggered
dim nGrandTotal
dim nTotalNoCalls
dim serval

serval="0"



cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATEFROM")
cCustCode = Request.QueryString("CUSTCODE")
cCustName = Request.QueryString("CUSTNAME")
lErrorTriggered = false
nGrandTotal = 0
nTotalNoCalls = 0
'
dRepDate = cDate(cStartDate)
cSQL = "SELECT BILLING_ID From BILLING Where accnt_hrcy_step_id=" & cAHS
set oRS = Conn.Execute(cSQL)
with oRS
	cBillID = .fields(0)
	.close
end with
'

cSQL = "SELECT BILLING_DETAIL.* " & _
			"FROM BILLING_DETAIL " & _
			"WHERE (BILLING_DETAIL.BILLING_ID = " & cBillID & _
			" AND BILLING_DETAIL.STATUS='ACTIVE' AND BILLING_DETAIL.CALLSTATUS = 'COMPLETED' " & _
			" AND BILLING_DETAIL.SITE_ID=1) " & _
			"Order by CLIENT_NAME,PARENT_NAME, CALL_TYPE, LOB_CD, CALL_END_TIME"
set oRS = Conn.Execute(cSQL)
'response.write(cSQL)
'response.end
'
cTime = CStr(FormatDateTime( Time, vbShortTime))
cTime = Replace(cTime, ":", "")
cTmpFile = cCustCode & "-" & cTime & ".xls"

with oExcel
	.cDestinationFileName = cTmpFile
	.cSourceFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Billing\Billing.xls"
	.cExcelRangeName = "ODBCRange"
	.cDownloadDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "Reports\Download\"
	'cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & FNSDesigner & "/Reports/Download"
	cDownloadLocation = "http://" & Request.ServerVariables("SERVER_NAME") & "/Reports/Download"
	.cDownloadLocation = cDownloadLocation
	.openXLS
	.writeMsg "Generating spreadsheet"
	.writeCell "Account", cCustName
	.writeCell "DbName", Left(ConnectionString, InStr(1, ConnectionString, ";", vbTextCompare) - 1 )
	'.writeCell "Period", Left(cStartDate,3) & " " & right(cStartDate,4)
end with
writePeriod(oExcel)
cFields = "Account,Call_No,T,Status,Loss_Dt,Call_Dt,Claim_No,LOB,Policy_No,Risk_Location,Caller_Name,Employee_Name,Employee_SSN,Service,Total_Fax,Temp,Escalate,Vendor,Total"
with oRS
	Do While Not .EOF
		if isNull(.Fields("PARENT_NAME").Value) then
			cParent = ""
		else
			cParent = .Fields("PARENT_NAME").Value
		end if
		'Response.Write cParent & "<br>1"
		nCalls = 0
		nTotSvcFee = 0
		nTotalFee = 0
		serval="0"
		do while Not .EOF
			if isNull(.Fields("PARENT_NAME").Value) then
				cCmpParent = ""
			else
				cCmpParent = .Fields("PARENT_NAME").Value
			end if
			if cParent = cCmpParent then
			    serval="0"
				nCalls = nCalls + 1
				nTotalFaxFee = CSng(.Fields("TOTAL_FAX_FEE").Value)
				nTempFee = CSng(.Fields("TEMP_FEE").Value)
				nEscalateFee = cSng(.Fields("ESCALATE_FEE").Value)
				nVendorFee = CSng(.Fields("VENDOR_FEE").Value)
				serval = CSng(.Fields("SERVICE_FEE").Value)
				'response.write(serval)


				if (mid((.Fields("Call_Type").Value ),1,1)="E") then
                              cSQL1 = " Select FEE_AMOUNT FROM FEE " & _
							  " WHERE ACCNT_HRCY_STEP_ID = " & cAHS & _
							  " AND FEE_TYPE_ID = 8 "
					   'M

					   'response.write(cSQL1)
					   'response.end
				'if not cSQL1="" then
			       set oRS1 = Conn.Execute(cSQL1)
						with oRS1
						     Do While Not oRS1.EOF
							 serval = .fields(0)
							 oRS1.MoveNext
				             loop
						     .close
						end with
               end if
		'response.write "test" & .Fields("PARENT_NAME").Value
		'response.end
                nTotSvcFee=nTotSvcFee + CSng(serval)
				nTotalFee = nTotalFee + CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee
				cValues =	"'" & removeSngQuote(.Fields("PARENT_NAME").Value) & "'," & _
					CStr(.Fields("Call_ID").Value) & ",'" & _
					.Fields("Call_Type").Value & "','" & _
					.Fields("CALLSTATUS").Value & "','"
					on error resume next
					if isNull(.Fields("LOSS_DATE").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("LOSS_DATE").Value
					end if

					if err.number <> 0 then
						if err.number = -2147217887 then
							writeError .Fields("Call_ID").Value
						end if
						lErrorTriggered = true
						exit do
					end if




					cValues = cValues & "','" & _
					CStr(.Fields("CALL_END_TIME").Value) & "','"
					if isNull(.Fields("CLAIM_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & .Fields("CLAIM_NUMBER").Value
					end if
					cValues = cValues & "','" & _
					.Fields("LOB_CD").Value & "','"
					if isNull(.Fields("POLICY_NUMBER").Value) then
						cValues = cValues & ""
					else
						cValues = cValues & removeSngQuote(.Fields("POLICY_NUMBER").Value)
					end if
					cValues = cValues & "','" & _
					removeSngQuote(.Fields("ACCOUNT_NAME").Value) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("CALLER_FIRST_NAME").Value
					aNameParts(1) = .Fields("CALLER_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					redim aNameParts(1)
					aNameParts(0) = .Fields("EMPLOYEE_FIRST_NAME").Value
					aNameParts(1) = .Fields("EMPLOYEE_LAST_NAME").Value
					cValues = cValues & getName(aNameParts) & "','"
					cValues = cValues & cEmployeeSSN & "','"
					cValues = cValues & serval & "','" & _
					FormatNumber(nTotalFaxFee) & "','" & _
					FormatNumber(nTempFee) & "','" & _
					FormatNumber(nEscalateFee) & "','" & _
					FormatNumber(nVendorFee) & "','" & _
					FormatNumber(CSng(serval) + nTotalFaxFee + nTempFee + nEscalateFee + nVendorFee) & "'"

				oExcel.addRow cFields, cValues
				.MoveNext
			else
				exit do
			end if
		loop
		if not lErrorTriggered then
			cValues = "'SubTotal'," & nCalls & ",'','','','','','','','','','','','" & FormatCurrency(nTotSvcFee) & "','','','','','" & FormatCurrency(nTotalFee) & "'"
			oExcel.addRow cFields, cValues
			cValues = "'','','','','','','','','','','','','','','','','','',''"
			oExcel.addRow cFields, cValues
		else
			exit do
		end if
		nGrandTotal = nGrandTotal + nTotalFee
		nTotalNoCalls = nTotalNoCalls + nCalls
	Loop
	.Close
end with
cValues = "'','','','','','','','','','','','','','','','','','',''"
oExcel.addRow cFields, cValues
cValues = "'Grand Total','" & nTotalNoCalls & "','','','','','','','','','','','','','','','','','" & FormatCurrency(nGrandTotal) & "'"
oExcel.addRow cFields, cValues
oExcel.closeXLS
if not lErrorTriggered then
	oExcel.sendFile
end if
Set oRS = Nothing
genTOWASPXLS = lErrorTriggered
end function
'*******************************
' END TOW ASP
'*******************************

sub writePeriod(oExcel)
dim dRepStart, dRepEnd
dRepStart = CDate(Request.QueryString("DATEFROM"))
dRepEnd = CDate(Request.QueryString("DATETO"))
oExcel.writeCell "Period", CStr(dRepStart) & " to " & CStr(dRepEnd)
end sub

%>
