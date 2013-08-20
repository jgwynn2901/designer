<!--#include file="..\lib\common.inc"-->
<%
dim oConn, cSQL, oRS
dim cName, cType, cTitle, cPhone, cFax, cEMail, cDesc

cType = "'" & Replace(Request.Form("CNT_TYPE"), "'", "''") & "' "
cName = "'" & Replace(Request.Form("CNT_NAME"), "'", "''") & "' "
cTitle = "'" & Replace(Request.Form("CNT_TITLE"), "'", "''") & "' "
cPhone = "'" & Request.Form("CNT_PHONE") & "' "
cFax = "'" & Request.Form("CNT_FAX") & "' "
cEMail = "'" & Replace(Request.Form("CNT_EMAIL"), "'", "''") & "' "
cDesc = "'" & Replace(Request.Form("CNT_DESC"), "'", "''") & "'"

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING

if Request.Form("ContactID") <> "" then
	' edit mode
	cSQL = "UPDATE CONTACT SET TYPE=" & cType & ",NAME=" & cName & ",TITLE=" & cTitle & ",PHONE=" & cPhone & ",FAX=" & cFax & ",EMAIL=" & cEMail & ",DESCRIPTION=" & cDesc
	cSQL = cSQL & " WHERE CONTACT_ID=" & Request.Form("ContactID")
else
	cSQL = "{call Designer.GetValidSeq('CONTACT', 'CONTACT_ID', {resultset 1, outResult})}"
	Set oRS = oConn.Execute(cSQL)
	
	cSQL = "INSERT INTO CONTACT (CONTACT_ID, ACCNT_HRCY_STEP_ID, TYPE, NAME, TITLE, PHONE, FAX, EMAIL, DESCRIPTION) VALUES ("
	cSQL = cSQL & oRS("outResult") & ", "
	cSQL = cSQL & Request.Form("AHSID") & ", "
	cSQL = cSQL & cType & ","
	cSQL = cSQL & cName & ","
	cSQL = cSQL & cTitle & ","
	cSQL = cSQL & cPhone & ","
	cSQL = cSQL & cFax & ","
	cSQL = cSQL & cEMail & ","
	cSQL = cSQL & cDesc & ")"
	oRS.close
end if	
oConn.Execute(cSQL)
oConn.Close
set oConn = nothing
%>
