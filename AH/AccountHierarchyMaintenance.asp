<!--#include file="..\lib\common.inc"-->
<% Response.expires=0 %>
<%
Function Swap (InData)
	If InData = "" Then
		Swap = "null"
	Else
		Swap = InData
	End If
End Function

Function SwapFlg(InData)
	If InData = "on" Then
		SwapFlg = "Y"
	Else
		SwapFlg = "N"
	End If
End Function

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	
If Request.QueryString("AHSID") = "NEW" Then
	PARENT_NODE_ID = Request.QueryString("PARENT_AHSID")
End If
If Request.QueryString("ACTION") = "SAVE" AND Request.QueryString("AHSID") = "NEW" Then
	QSQL = ""
	QSQL = QSQL & "{call Designer.GetValidSeq('ACCOUNT_HIERARCHY_STEP', 'ACCNT_HRCY_STEP_ID', {resultset 1, outResult})}"
	Set RSNextID = Conn.Execute(QSQL)

	SQLINSERT = ""
	SQLINSERT = SQLINSERT & "INSERT INTO ACCOUNT_HIERARCHY_STEP ("
	SQLINSERT = SQLINSERT & "ACCNT_HRCY_STEP_ID, "
	SQLINSERT = SQLINSERT & "NODE_TYPE_ID,NAME,TYPE, FNS_CLIENT_CD,"
	SQLINSERT = SQLINSERT & "ADDRESS_1,ADDRESS_2,ADDRESS_3,COUNTRY, "
	SQLINSERT = SQLINSERT & "CITY, STATE, ZIP, FIPS, CLIENT_NODE_ID,PEER_NODE_ID, "
	SQLINSERT = SQLINSERT & "PARENT_NODE_ID, COUNTY, PHONE,FAX,FEIN,SIC,SUID, "
	SQLINSERT = SQLINSERT & "NATURE_OF_BUSINESS, LOCATION_CODE,AUTO_ESCALATE, "
	SQLINSERT = SQLINSERT & "ESCALATION_CALLBACK_NUM, CREATED_DT,UPLOAD_KEY,"
	SQLINSERT = SQLINSERT & "ACTIVE_STATUS, STATUS_DATE, POLICY_SEARCH_ID) VALUES ( "
	SQLINSERT = SQLINSERT & RSNextID("outResult") & ", "
	SQLINSERT = SQLINSERT & Swap(Request.Form("NODE_TYPE_ID"))  & ", "
	SQLINSERT = SQLINSERT & "'" & Request.Form("NAME")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("TYPE")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("FNS_CLIENT_CD")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("ADDRESS_1")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("ADDRESS_2")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("ADDRESS_3")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("COUNTRY")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("CITY")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("STATE")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("ZIP")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("FIPS")  & "', "
	SQLINSERT = SQLINSERT & Swap(Request.Form("CLIENT_NODE_ID"))  & ", "
	SQLINSERT = SQLINSERT & Swap(Request.Form("PEER_NODE_ID"))  & ", "
	SQLINSERT = SQLINSERT & Swap(Request.Form("PARENT_NODE_ID"))  & ", "
	SQLINSERT = SQLINSERT & "'" & Request.Form("COUNTY")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("PHONE")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("FAX")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("FEIN")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("SIC")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("SUID")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("NATURE_OF_BUSINESS")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("LOCATION_CODE")  & "', "
	SQLINSERT = SQLINSERT & "'" & SwapFlg(Request.Form("AUTO_ESCALATE"))  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("ESCALATION_CALLBACK_NUM")  & "', "
	SQLINSERT = SQLINSERT & "TO_DATE('" & date() & "', 'MM/DD/YY'),"
	SQLINSERT = SQLINSERT & "'" & Request.Form("UPLOAD_KEY")  & "', "
	SQLINSERT = SQLINSERT & "'" & Request.Form("ACTIVE_STATUS")  & "', "
	SQLINSERT = SQLINSERT & "TO_DATE('" & date() & "', 'MM/DD/YY'),"
	SQLINSERT = SQLINSERT & Swap(Request.Form("POLICY_SEARCH_ID"))  & ") "
	Set RSIn = Conn.Execute(SQLINSERT)
	Response.Redirect "AccountHierarchyMaintenance.asp?AHSID=" & RSNextID("outResult") & "&STATUS=SAVE"
End If

If Request.QueryString("ACTION") = "SAVE" AND Request.QueryString("AHSID") <> "NEW" Then
	SQLSAVE = ""
	SQLSAVE = SQLSAVE & "UPDATE ACCOUNT_HIERARCHY_STEP SET "
	SQLSAVE = SQLSAVE & "NODE_TYPE_ID=" & Swap(Request.Form("NODE_TYPE_ID")) & ", "
	SQLSAVE = SQLSAVE & "NAME='" & Replace(Request.Form("NAME"), "'", "''") & "', "
	SQLSAVE = SQLSAVE & "TYPE='" & Replace(Request.Form("TYPE"), "'", "''") & "', "
	SQLSAVE = SQLSAVE & "FNS_CLIENT_CD='" &Replace( Request.Form("FNS_CLIENT_CD"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "ADDRESS_1='" & Replace(Request.Form("ADDRESS_1"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "ADDRESS_2='" & Replace(Request.Form("ADDRESS_2"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "ADDRESS_3='" & Replace(Request.Form("ADDRESS_3"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "COUNTRY='" & Replace(Request.Form("COUNTRY"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "CITY='" & Replace(Request.Form("CITY"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "STATE='" & Replace(Request.Form("STATE"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "ZIP='" & Replace(Request.Form("ZIP"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "FIPS='" & Replace(Request.Form("FIPS"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "CLIENT_NODE_ID=" & Swap(Request.Form("CLIENT_NODE_ID")) & ","
	SQLSAVE = SQLSAVE & "PEER_NODE_ID=" & Swap(Request.Form("PEER_NODE_ID")) & ","
	SQLSAVE = SQLSAVE & "PARENT_NODE_ID=" & Swap(Request.Form("PARENT_NODE_ID")) & ","
	SQLSAVE = SQLSAVE & "COUNTY='" & Replace(Request.Form("COUNTY"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "PHONE='" & Replace(Request.Form("PHONE"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "FAX='" & Replace(Request.Form("FAX"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "FEIN='" & Replace(Request.Form("FEIN"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "SIC='" & Replace(Request.Form("SIC"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "SUID='" & Replace(Request.Form("SUID"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "NATURE_OF_BUSINESS='" & Replace(Request.Form("NATURE_OF_BUSINESS"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "LOCATION_CODE='" & Replace(Request.Form("LOCATION_CODE"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "AUTO_ESCALATE='" & SwapFlg(Request.Form("AUTO_ESCALATE")) & "',"
	SQLSAVE = SQLSAVE & "ESCALATION_CALLBACK_NUM='" & Replace(Request.Form("ESCALATION_CALLBACK_NUM"), "'", "''") & "',"
	'SQLSAVE = SQLSAVE & "NODE_LEVEL='" & Request.Form("NODE_LEVEL") & "',"
	SQLSAVE = SQLSAVE & "MODIFIED_DT= TO_DATE('" & date() & "', 'MM/DD/YY') ,"
	SQLSAVE = SQLSAVE & "UPLOAD_KEY='" & Replace(Request.Form("UPLOAD_KEY"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "ACTIVE_STATUS='" & Replace(Request.Form("ACTIVE_STATUS"), "'", "''") & "',"
	SQLSAVE = SQLSAVE & "STATUS_DATE= TO_DATE('" & date() & "', 'MM/DD/YY') ,"
	SQLSAVE = SQLSAVE & "POLICY_SEARCH_ID=" & Swap(Replace(Request.Form("POLICY_SEARCH_ID"), "'", "''")) 
	SQLSAVE = SQLSAVE & " WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	Set RSSave = Conn.Execute(SQLSAVE)
	Response.Redirect "AccountHierarchyMaintenance.asp?AHSID=" & Request.QueryString("AHSID") & "&STATUS=SAVE"
End If
	
If Request.QueryString("AHSID") <> "" AND Request.QueryString("AHSID") <> "NEW" Then
	SQL = ""
	SQL = SQL & "SELECT TO_CHAR(CREATED_DT, 'MM/DD/YYYY') As CREATED_DT, "
	SQL = SQL & "TO_CHAR(MODIFIED_DT, 'MM/DD/YYYY') As MODIFIED_DT, "
	SQL = SQL & "NODE_TYPE_ID, PARENT_NODE_ID, CLIENT_NODE_ID, NAME, PEER_NODE_ID, "
	SQL = SQL & "AUTO_ESCALATE, FNS_CLIENT_CD, ADDRESS_1, ADDRESS_2, ADDRESS_3, "
	SQL = SQL & "COUNTRY, CITY, STATE, ZIP, FIPS, COUNTY, PHONE, FAX, FEIN, SIC, SUID, "
	SQL = SQL & "NATURE_OF_BUSINESS, NODE_LEVEL, LOCATION_CODE, ESCALATION_CALLBACK_NUM, "
	SQL = SQL & "UPLOAD_KEY, ACTIVE_STATUS, TO_CHAR(STATUS_DATE, 'MM/DD/YYYY') As STATUS_DATE, POLICY_SEARCH_ID, TYPE "
	SQL = SQL & "FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	Set RS = Conn.Execute(SQL)
        
NODE_TYPE_ID = RS("NODE_TYPE_ID")                 
PARENT_NODE_ID = RS("PARENT_NODE_ID")               
CLIENT_NODE_ID = RS("CLIENT_NODE_ID")        
NAME = RS("NAME")                      
PEER_NODE_ID = RS("PEER_NODE_ID")                 
AUTO_ESCALATE = RS("AUTO_ESCALATE")                
FNS_CLIENT_CD = RS("FNS_CLIENT_CD")               
ADDRESS_1 = RS("ADDRESS_1")                    
ADDRESS_2 = RS("ADDRESS_2")                    
ADDRESS_3 = RS("ADDRESS_3")                    
COUNTRY = RS("COUNTRY")                        
CITY = RS("CITY")                         
STATE = RS("STATE")                        
ZIP = RS("ZIP")                           
FIPS = RS("FIPS")                         
COUNTY = RS("COUNTY")                        
PHONE = RS("PHONE")                       
FAX = RS("FAX")                          
FEIN = RS("FEIN")                         
SIC = RS("SIC")                         
SUID = RS("SUID")                       
NATURE_OF_BUSINESS = RS("NATURE_OF_BUSINESS")            
NODE_LEVEL = RS("NODE_LEVEL")                 
LOCATION_CODE = RS("LOCATION_CODE")             
ESCALATION_CALLBACK_NUM = RS("ESCALATION_CALLBACK_NUM")     
UPLOAD_KEY = RS("UPLOAD_KEY")                  
ACTIVE_STATUS = RS("ACTIVE_STATUS")               
STATUS_DATE = RS("STATUS_DATE")                 
POLICY_SEARCH_ID = RS("POLICY_SEARCH_ID")             
CREATED_DT = RS("CREATED_DT")                    
MODIFIED_DT = RS("MODIFIED_DT")
ATYPE = RS("TYPE")

End If
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function BtnSave_onclick() {
var strmsg
strmsg = ""
if (document.all.ACTIVE_STATUS.value == "")
	{
	strmsg = strmsg + "Active Status is a required field."
	}
if (strmsg == "")
	{
	document.all.FrmSave.submit()
	}
	else
	{
	alert (strmsg)
	}
}

function window_onload() {
	document.all.STATE.value = "<%= STATE %>"
	document.all.NODE_TYPE_ID.value = "<%= NODE_TYPE_ID %>"
	document.all.ACTIVE_STATUS.value = "<%= ACTIVE_STATUS %>"
	<% If AUTO_ESCALATE = "Y" Then  %>
		document.all.AUTO_ESCALATE.checked = true
	<% End If %>
<%	 Select Case Request.QueryString("STATUS")
			Case "SAVE" %>
				document.all.SPANSTATUS.innerHTML = ": Saved"
<%		Case "NEW" %>
				document.all.SPANSTATUS.innerHTML = ": New"
<%	    Case Else %>
			document.all.SPANSTATUS.innerHTML = ": Ready"
	<% End Select %>
}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=window EVENT=onload>
<!--
 window_onload()
//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub BtnGrfxBack_Onclick()
<% If Request.QueryString("AHSID") <> "NEW" Then %>
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
<% Else %>
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("PARENT_AHSID") %>"
<% End If %>
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  rightmargin=0 leftmargin=0 bottommargin=0 topmargin=0 BGCOLOR='<%=BODYBGCOLOR%>' >
<!--#include file="..\lib\NavBack.inc"-->
<FORM NAME=FrmSave METHOD=POST ACTION="AccountHierarchyMaintenance.asp?ACTION=SAVE&AHSID=<%= Request.QueryString("AHSID") %>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp;&#187; Business Entity Details </TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1 WIDTH=100%></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>


<TABLE  cellspacing=0 cellpadding=0>
<TR>
<TD CLASS=LABEL><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" Title=""></TD>
<TD CLASS=LABEL ALIGN=LEFT><SPAN ID=SPANSTATUS STYLE="COLOR:#006699" CLASS=LABEL></SPAN></TD>
</TR>
</TABLE>
<TABLE  cellspacing=0 cellpadding=0>
<TR>
<TD CLASS=LABEL WIDTH="70%" ALIGN=LEFT><LABEL STYLE="COLOR:black" CLASS=LABEL><NOBR>Created: <%= CREATED_DT %></LABEL></TD>
<TD CLASS=LABEL ALIGN=RIGHT><LABEL STYLE="COLOR:black" CLASS=LABEL><NOBR>Updated: <%= MODIFIED_DT %></LABEL></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL COLSPAN=2>Name:<BR><INPUT TYPE=TEXT NAME=NAME CLASS=LABEL SIZE=60 MAXLENGTH=80 VALUE="<%= NAME %>"></TD>
<TD CLASS=LABEL>FNS Client Code:<BR><INPUT TYPE=TEXT NAME=FNS_CLIENT_CD CLASS=LABEL SIZE=3 MAXLENGTH=3 VALUE="<%= FNS_CLIENT_CD %>"></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Address 1:<BR><INPUT TYPE=TEXT NAME=ADDRESS_1 CLASS=LABEL SIZE=45 MAXLENGTH=45 VALUE="<%= ADDRESS_1 %>"></TD>
<TD CLASS=LABEL>Phone:<BR><INPUT TYPE=TEXT NAME=PHONE CLASS=LABEL SIZE=20 MAXLENGTH=20 VALUE="<%= PHONE %>"></TD>
<TD CLASS=LABEL>Fax:<BR><INPUT TYPE=TEXT NAME=FAX CLASS=LABEL SIZE=20 MAXLENGTH=20 VALUE="<%= FAX %>"></TD>
</TR>
<TR>
<TD CLASS=LABEL>Address 2:<BR><INPUT TYPE=TEXT NAME=ADDRESS_2 CLASS=LABEL SIZE=45 MAXLENGTH=45 VALUE="<%= ADDRESS_2 %>"></TD>
<TD CLASS=LABEL>FEIN:<BR><INPUT TYPE=TEXT NAME=FEIN CLASS=LABEL SIZE=20 MAXLENGTH=20 VALUE="<%= FEIN %>"></TD>
<TD CLASS=LABEL>SIC:<BR><INPUT TYPE=TEXT NAME=SIC CLASS=LABEL SIZE=20 MAXLENGTH=6 VALUE="<%= SIC %>"></TD>
</TR>
<TR>
<TD CLASS=LABEL>Address 3:<BR><INPUT TYPE=TEXT NAME=ADDRESS_3 CLASS=LABEL SIZE=45 MAXLENGTH=45 VALUE="<%= ADDRESS_3 %>"></TD>
<TD CLASS=LABEL Colspan=2>Nature of Business:<BR><INPUT TYPE=TEXT NAME=NATURE_OF_BUSINESS CLASS=LABEL SIZE=45 MAXLENGTH=30 VALUE="<%= NATURE_OF_BUSINESS %>"></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>City:<BR><INPUT TYPE=TEXT NAME=CITY CLASS=LABEL SIZE=30 MAXLENGTH=30 VALUE="<%= CITY %>"></TD>
<TD CLASS=LABEL>State:<BR>
<SELECT NAME=STATE CLASS=LABEL>
<OPTION VALUE="">
<!--#include file="..\lib\States.asp"-->
</SELECT></TD>
<TD CLASS=LABEL>Zip:<BR><INPUT TYPE=TEXT NAME=ZIP CLASS=LABEL SIZE=10 MAXLENGTH=20 VALUE="<%= ZIP %>"></TD>
<TD CLASS=LABEL>County:<BR><INPUT TYPE=TEXT NAME=COUNTY CLASS=LABEL SIZE=10 MAXLENGTH=30 VALUE="<%= COUNTY %>"></TD>
<TD CLASS=LABEL>Country:<BR><INPUT TYPE=TEXT NAME=COUNTRY CLASS=LABEL SIZE=22 MAXLENGTH=80 VALUE="<%= COUNTRY %>"></TD>

</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>SUID:<BR><INPUT TYPE=TEXT NAME=SUID CLASS=LABEL SIZE=20 MAXLENGTH=20 VALUE="<%= SUID %>"></TD>
<TD CLASS=LABEL>FIPS:<BR><INPUT TYPE=TEXT NAME=FIPS CLASS=LABEL SIZE=5 MAXLENGTH=5 VALUE="<%= FIPS %>"></TD>
<TD CLASS=LABEL>Location Code:<BR><INPUT TYPE=TEXT NAME=LOCATION_CODE CLASS=LABEL SIZE=30 MAXLENGTH=30 VALUE="<%= LOCATION_CODE %>"></TD>
<TD CLASS=LABEL>Node Type:<BR>
<SELECT NAME=NODE_TYPE_ID CLASS=LABEL STYLE="WIDTH:145">
<OPTION VALUE="">
<%
SQLNODE = ""
SQLNODE = SQLNODE & "SELECT * FROM NODE_TYPE ORDER BY NAME"
Set RS2 = Conn.Execute(SQLNODE)
Do WHile Not RS2.EOF
%>
<OPTION VALUE="<%= RS2("NODE_TYPE_ID") %>"><%= RS2("NAME") %>
<% 
RS2.MoveNext
Loop
RS2.Close
%>
</SELECT>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2>Type:<BR><INPUT TYPE=TEXT NAME=TYPE CLASS=LABEL SIZE=30 MAXLENGTH=30 VALUE="<%= ATYPE %>"></TD>
<TD CLASS=LABEL>Escalation Call Back #:<BR><INPUT TYPE=TEXT NAME=ESCALATION_CALLBACK_NUM CLASS=LABEL SIZE=30 MAXLENGTH=10 VALUE="<%= ESCALATION_CALLBACK_NUM %>"></TD>
<TD CLASS=LABEL VALIGN=BOTTOM><INPUT TYPE=CHECKBOX NAME=AUTO_ESCALATE CLASS=LABEL>Auto Escalate?</TD>
</TR>
<TR>

<TD CLASS=LABEL COLSPAN=2>Policy Search ID:<BR><INPUT TYPE=TEXT NAME=POLICY_SEARCH_ID CLASS=LABEL SIZE=30 MAXLENGTH=30 VALUE="<%= POLICY_SEARCH_ID %>"></TD>
<TD CLASS=LABEL >Active Status:<BR>
<SELECT NAME=ACTIVE_STATUS STYLE="WIDTH:100" CLASS=LABEL>
<OPTION VALUE="ACTIVE">Active
<OPTION VALUE="COMBINED">Combined
<OPTION VALUE="DEACTIVATED">Deactivated
</SELECT>
</TD>
<!--<TD CLASS=LABEL>Status Date:<BR><INPUT TYPE=TEXT NAME=STATUS_DATE CLASS=LABEL SIZE=25 MAXLENGTH=20 VALUE="<%= STATUS_DATE %>"></TD>-->
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=8>Upload Key:<BR><INPUT TYPE=TEXT NAME=UPLOAD_KEY CLASS=LABEL SIZE=93 MAXLENGTH=255 VALUE="<%= UPLOAD_KEY %>"></TD>
</TR>
</TABLE>
<TABLE WIDTH="100%">
<TR>
<TD CLASS=LABEL>Client Node ID:<BR>
<IMG SRC="../Images/Attach.gif" ID=BtnATTACHCLIENTNODE STYLE="CURSOR:HAND" ALT="Attach Client Node">
<IMG SRC="../Images/Detach.gif" ID=BtnDETACHCLIENTNODE STYLE="CURSOR:HAND" ALT="Detach Client Node">
<INPUT TYPE=TEXT READONLY STYLE="BACKGROUND-COLOR:SILVER" NAME=CLIENT_NODE_ID CLASS=LABEL SIZE=10 MAXLENGTH=10 VALUE="<%= CLIENT_NODE_ID %>"></TD>
<TD CLASS=LABEL COLSPAN=2>Peer Node ID:<BR>
<IMG SRC="../Images/Attach.gif" ID=BtnATTACHPEERNODE STYLE="CURSOR:HAND" ALT="Attach Peer Node">
<IMG SRC="../Images/Detach.gif" ID=BtnDETACHPEERNODE STYLE="CURSOR:HAND" ALT="Detach Peer Node">
<INPUT TYPE=TEXT READONLY STYLE="BACKGROUND-COLOR:SILVER" NAME=PEER_NODE_ID CLASS=LABEL SIZE=10 MAXLENGTH=10 VALUE="<%= PEER_NODE_ID %>"></TD>
<TD CLASS=LABEL>Parent Node ID:<BR>
<IMG SRC="../Images/Attach.gif" ID=BtnATTACHPARENTNODE STYLE="CURSOR:HAND" ALT="Attach Parent Node">
<IMG SRC="../Images/Detach.gif" ID=BtnDETACHPARENTNODE STYLE="CURSOR:HAND" ALT="Detach Parent Node">
<INPUT TYPE=TEXT READONLY STYLE="BACKGROUND-COLOR:SILVER" NAME=PARENT_NODE_ID CLASS=LABEL SIZE=10 MAXLENGTH=10 VALUE="<%= PARENT_NODE_ID %>"></TD>
</TR>
</TABLE>
<BR><BR>&nbsp;
<BUTTON CLASS=STDBUTTON NAME=BtnSave LANGUAGE=javascript onclick="return BtnSave_onclick()"><U>S</U>ave</BUTTON>
&nbsp;
<BUTTON CLASS=STDBUTTON NAME=BtnCancel onclick="BtnGrfxBack_Onclick()"><U>C</U>ancel</BUTTON>
</FORM>
</BODY>
</HTML>
