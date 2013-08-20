<%	Response.Expires = 0
	On Error Resume Next

%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\CheckLogicExpression.inc"-->

<HTML>
<HEAD>
<SCRIPT LANGUAGE="VBScript">

<%
dim lErrors, lValid
dim cRID, cRIDText, cRIDType, cRIDComments

dim cRIDLang, lUserId

dim oConn, oRSRule
dim cSQL, cTemp
dim cWarning, cError

lErrors = True
lValid = false
		
cRID	= Request.Form("RID")
cRIDText = Request.Form("TxtRuleText")
cRIDType = Request.Form("TxtRuleType")
cRIDComments = Request.Form("TxtCommentsText")
	
cRIDText = Replace(cRIDText,"'","''")
cRIDText = Replace(cRIDText,CHR(13)," ")
cRIDText = Replace(cRIDText,CHR(10)," ")

cRIDLang= "01"			' pended...
	
If cRID = "NEW" Then 
	cRID = ""
end if
if ucase(cRIDType) = "ROUTING" Then
	lValid = CheckLogicExpression(cRIDText,"VBScript")
elseif request.form("ValidClient") = "True" then
	lValid = true
else
	lValid = false
end if

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING

cTemp = Replace(cRIDText,"[]","")
cSQL = "{call Designer_2.CheckAttrNamesInExpression(" &_
				"'" & cTemp & "','~', {resultset  2, cStatusMsg,nStatusCode})}"
Set oRS = oConn.Execute(cSQL)			
if CStr(oRS("nStatusCode")) <> "0" then
	cWarning = oRS("cStatusMsg") & VBCRLF & CheckADOErrors(oConn,"Rule CheckAttrNamesInExpression")
else
	lValid = true	
End If
oRS.Close

lUserId = Session("SecurityObj").m_UserID
If cRID = "" Then
	cSQL = "{call Designer_RULES.SP_ADD_RULE('"& cRIDType & "', '"& cRIDText & "',"&  lUserId & ", '"& cRIDLang & "', '"& cRIDComments & "', {resultset  1, outResult})}"
Else	
	cSQL = "{call Designer_RULES.SP_UPDATE_RULE('"& cRID & "','"& cRIDType & "', '"& cRIDText & "',"&  lUserId & ", '"& cRIDLang &"', '"& cRIDComments & "', {resultset  1, outResult})}"
End If

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly, adCmdText
if CStr(oRS("outResult"))="0" then
	lErrors = true
else
	If cRID = "" Then 
		cRID = oRS("outResult") 
		%>parent.frames("WORKAREA").SetRID(<%=cRID%>)<%
	end if
	lErrors = false
end if
cError = CheckADOErrors(oConn,"Rules Save" )
oRS.Close
Set oRS = Nothing
oConn.Close
Set oConn = Nothing
If lErrors Or cError <> "" Then 
		LogStatusGroupBegin
		LogStatus S_ERROR, cError, "RULES", "", 0, ""
		LogStatusGroupEnd
%>		
parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.")	
parent.frames("WORKAREA").SetDirty
parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true)		
<%	
ElseIf not lValid Or cWarning <> "" Then
		LogStatusGroupBegin
		If not lValid Then 
			LogStatus S_WARNING, "Rule syntax is not valid.", "RULES", "", 0, ""
		end if
		If cWarning <> "" Then 
			LogStatus S_WARNING, cWarning, "RULES", "", 0, ""
		end if
		LogStatusGroupEnd
%>
parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Update successful with warnings, check status report.")	
parent.frames("WORKAREA").ClearDirty
parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true)		
<%
	Else
		LogStatusGroupBegin
		LogStatusGroupEnd
%>
parent.frames("WORKAREA").UpdateStatus("Update successful.")			
parent.frames("WORKAREA").ClearDirty
parent.frames("WORKAREA").SetStatusInfoAvailableFlag(false)

<%	End If 
			
%>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>


