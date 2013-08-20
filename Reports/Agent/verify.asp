<!--#include file="..\..\lib\genericSQL.asp"-->
<%
dim cSQL, oRS
dim lMissingBM, lMissingEmail, cAgentName, lWithError
dim cAHS

cAHS = Request.QueryString("AHS")
if cAHS = "23" then
	cSQL = "Select * From ACCOUNT_HIERARCHY_STEP Where PARENT_NODE_ID = 23 AND ACTIVE_STATUS='ACTIVE'"
else
	cSQL = "Select * From ACCOUNT_HIERARCHY_STEP Where ACCNT_HRCY_STEP_ID = " & cAHS
end if	
set oRS = Conn.Execute(cSQL)
lWithError = false
with response
	.Write "<body BGCOLOR=""#d6cfbd"">" & vbCRLF
	.Write "<div align=""center"">" & vbCRLF
	.Write "<TABLE border=""1"" width=""80%"">" & vbCRLF
	.Write "<tr><td colspan=""3""><font face=""Tahoma"" size=""4"">Verifying Agents</font></td></tr>" & vbCRLF
	.Write "<tr><td width=""60%""><font face=""Tahoma"" size=""3""><b>Agent</b></font></td>" & vbCRLF
	.Write "<td width=""20%"" align=""center""><font face=""Tahoma"" size=""3""><b>Service Type</b></font></td>" & vbCRLF
	.Write "<td width=""20%"" align=""center""><font face=""Tahoma"" size=""3""><b>Email</b></font></td></tr>" & vbCRLF
end with
Do While Not oRS.EOF
	lMissingBM = false
	lMissingEmail = false
	cAgentName = oRS.Fields("NAME").Value
	If Not IsNull(oRS.Fields("AGENT_BILLING_METHOD").Value) Then
		if len(oRS.Fields("AGENT_BILLING_METHOD").Value) = 0 then
			lMissingBM = true
		end if
	else
		lMissingBM = true
	End If
	If Not IsNull(oRS.Fields("EMAIL_ADDRESS").Value) Then
		if len(oRS.Fields("EMAIL_ADDRESS").Value) = 0 then
			lMissingEmail = true
		end if
	else
		lMissingEmail = true
	End If
	if lMissingEmail or lMissingBM then
		lWithError = true
		response.Write "<tr><td width=""60%""><font face=""Tahoma"" size=""2"">" & cAgentName & "</font></td>" & vbCRLF
		if lMissingBM then
			response.Write "<td width=""20%"" align=""center""><font face=""Tahoma"" size=""2"">missing</font></td>" & vbCRLF
		else
			response.Write "<td width=""20%"" align=""center""><font face=""Tahoma"" size=""2"">ok</font></td>" & vbCRLF
		end if
		if lMissingEmail then
			response.Write "<td width=""20%"" align=""center""><font face=""Tahoma"" size=""2"">missing</font></td></tr>" & vbCRLF
		else
			response.Write "<td width=""20%"" align=""center""><font face=""Tahoma"" size=""2"">ok</font></td></tr>" & vbCRLF
		end if
	end if
	oRS.movenext
loop
response.Write "</table></div>" & vbCRLF
if lWithError then
	response.Write "<div style=""{position:relative;top:35;}"">" & vbCRLF
	response.Write "<font face=""Tahoma"" size=""3"">Please fix the error(s) and try again.</font>" & vbCRLF
	response.Write "</div>" & vbCRLF
end if
response.Write "</body>" & vbCRLF
Conn.close
set Conn = nothing
set oRS = nothing
if not lWithError then
	response.redirect "../Billing/ss.asp?" & request.QueryString & "&CUSTNAME=" & server.HTMLEncode( cAgentName )
end if	
%>
