<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<%
Response.Expires=0 
Response.AddHeader  "Pragma", "no-cache"
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "CF_SQLStatement.asp"
End If
If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"

If Request.querystring("STATUS")<>"TRUE" Then
	Session("StatusMsg") = ""
	Session("StatusNum") = ""
End If

If Len(Request.QueryString("FRAMEID")) < 1 OR IsNumeric(Request.QueryString("FRAMEID")) = False Then
	Session("ErrorMessage") = "On page " &  Request.ServerVariables("SCRIPT_NAME") & " QueryString FRAMEID was Null or Not Numeric"
	Response.Redirect "..\directerror.asp"
End If

If Request.QueryString("ACTION") <> "SAVE" Then
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "SELECT SQLSelect, SQLFrom, SQLWhere, SQLOrderBy, TYPE FROM FRAME WHERE FRAME_ID =" & Request.QueryString("FRAMEID")
	Set RS = Conn.Execute(SQLST)
	If RS.EOF or isnull(RS) Then
		Session("ErrorMessage") = "Statement = " & SQLST & " ----- returned no records" & vbCrlf
		Response.redirect	 "..\directerror.asp"
	End If
%>
<html>
<head>
<object ID="ClipboardAgent" <%GetClipboardCLSID()%> width="1" height="1">
<param NAME="MaxPropertiesStringLength" VALUE="1000">
<param NAME="MaxPropertyNameLength" VALUE="50">
<param NAME="MaxPropertyValueLength" VALUE="200">
<param NAME="NameValueDelimiter" VALUE="#">
<param NAME="PropertyItemDelimiter" VALUE="|">
<param NAME="PrivateClipboardFormatName" VALUE="CF_TEXT">
</object>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->

Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub

Sub BtnSave_onclick
		FrmFrame.Submit()
End Sub
-->
</script>
<script LANGUAGE="JavaScript">
function CanDocUnloadNow()
{
	if (false == confirm("Data has changed. Leave page without saving?"))
		return false;
	else
		return true;
}
</script>
</head>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
function CAttrSearchObj()
{
	this.selected = "";
	this.AID = "";
	this.AIDName = "";
	this.AIDCaption = "";
	this.AIDInputType = "";
}	
function CCopyObj()
{
	this.name = "";
}
var SearchObj = new CAttrSearchObj();
var CopyObj = new CCopyObj();

function BtnWhere_onclick() {
//lret = window.showModalDialog("SQLHelp.asp?SQL=WHERE", null,  "dialogWidth=320px; dialogHeight=300px; center=yes");
}
function BtnSelect_onclick() {
lret = window.showModalDialog("../Attribute/AttributeMaintenance.asp?SEARCHONLY=TRUE&COPYPASTE=TRUE", SearchObj,  "dialogWidth=520px; dialogHeight=500px; center=yes");
			lret = ClipboardAgent.ClearAllProperties();
			lret = ClipboardAgent.PropertiesString = "~" + SearchObj.AIDName + "~";
			lret = ClipboardAgent.SetPropertiesToClipboard();
}
function BtnFrom_onclick() {
lret = window.showModalDialog("SQLHelpModal.asp?SQL=FROM", CopyObj,  "dialogWidth=650px; dialogHeight=425px; center=yes");
			lret = ClipboardAgent.ClearAllProperties();
			lret = ClipboardAgent.PropertiesString = CopyObj.name;
			lret = ClipboardAgent.SetPropertiesToClipboard();
}
function BtnOrderby_onclick() {
//lret = window.showModalDialog("SQLHelp.asp?SQL=ORDER", null,  "dialogWidth=320px; dialogHeight=300px; center=yes");
}
function BtnClear_onclick() {
	document.all.SQLSelect.value = ""
	document.all.SQLFrom.value = ""
	document.all.SQLOrderBy.value = ""
	document.all.SQLWhere.value = ""
}
//-->
</script>
</head>
<body BGCOLOR="#d6cfbd" topmargin="0" leftmargin="0" CanDocUnloadNowInf="NO" ScreenMode="<%= MODE %>">
<form NAME="FrmFrame" ACTION="CF_SQLStatement.ASP?ACTION=SAVE&amp;FRAMEID=<%= Request.QueryString("FRAMEID") %>" METHOD="POST">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» SQL Data
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;SQL Data.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<input TYPE="HIDDEN" NAME="TYPE" VALUE="<%= RS("TYPE") %>">
<table CLASS="LABEL">
<tr>
<td CLASS="LABEL">SQL Select:<br><textarea CLASS="LABEL" NAME="SQLSelect" COLS="120" ROWS="5" OnChange="SetDirty" <TEXTAREA <% If MODE="RO" Then Response.Write(" DISABLED STYLE='BACKGROUND-COLOR:SILVER' ") %> OnKeyPress="SetDirty"><%= RS("SQLSelect") %></textarea>
</td>
</tr>
<tr>
<td CLASS="LABEL">SQLFrom:<br><textarea CLASS="LABEL" NAME="SQLFrom" COLS="120" ROWS="5" <TEXTAREA <% If MODE="RO" Then Response.Write(" DISABLED STYLE='BACKGROUND-COLOR:SILVER' ") %> OnChange="SetDirty" OnKeyPress="SetDirty"><%= RS("SQLFrom") %></textarea>
</td>
</tr>
<tr>
<td CLASS="LABEL">SQLWhere:<br><textarea <% If MODE="RO" Then Response.Write(" DISABLED STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS="LABEL" NAME="SQLWhere" COLS="120" ROWS="5" OnChange="SetDirty" OnKeyPress="SetDirty"><%= RS("SQLWhere") %></textarea></td>
</tr>
<tr>
<td CLASS="LABEL">SQLOrderBy:<br><textarea <% If MODE="RO" Then Response.Write(" DISABLED STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS="LABEL" NAME="SQLOrderBy" COLS="120" ROWS="5" OnChange="SetDirty" OnKeyPress="SetDirty"><%= RS("SQLOrderBy") %></textarea></td>
</tr>
<tr>
<td CLASS="LABEL"><button CLASS="StdButton" <% If MODE="RO" Then Response.Write(" DISABLED ") %> Name="BtnSave" ACCESSKEY="S"><u>S</u>ave</button>
&nbsp;<button NAME="BtnClear" CLASS="StdButton" <% If MODE="RO" Then Response.Write(" DISABLED ") %> ACCESSKEY="L" LANGUAGE="javascript" onclick="return BtnClear_onclick()">C<u>l</u>ear</button>
&nbsp;<button NAME="BtnSelect" CLASS="StdButton" STYLE="CURSOR:HAND;WIDTH:150" LANGUAGE="javascript" onclick="return BtnSelect_onclick()" title="Review Attribute Names">Copy Attribute Name</button>
&nbsp;<button NAME="BtnFrom" CLASS="StdButton" STYLE="CURSOR:HAND;WIDTH:150" LANGUAGE="javascript" onclick="return BtnFrom_onclick()" title="Review Table Names">Copy Table Name</button>
</td>
</tr>
</table>
<table>
<tr>
<% If Session("StatusNum") <> "" AND Request.querystring("STATUS")="TRUE" Then %>
	<td CLASS="LABEL"><img SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER="0" TITLE="Status Report" NOWRAP VALIGN="BOTTOM" NAME="BtnStatus" ID="BtnStatus" WIDTH="16" HEIGHT="16"></td>
	<td CLASS="LABEL">
	<% If Session("StatusNum") <> "0" Then %>
		<font COLOR="MAROON">Saved! SQL statment may be incorrect please check your syntax.<br><%= Session("StatusMsg")%> </font></td>
	<% Else %>
		<font COLOR="#006699">Saved! <br><%= Session("StatusMsg")%> </font></td>
	<% End If 
End If 
%>
<% If Session("StatusNum") = "" AND Request.querystring("STATUS")="TRUE" Then %>
	<td CLASS="LABEL"><img SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER="0" TITLE="Status Report" NOWRAP VALIGN="BOTTOM" NAME="BtnStatus" ID="BtnStatus" WIDTH="16" HEIGHT="16"></td>
	<td CLASS="LABEL"><font COLOR="#006699">Saved! </font></td>
<% End If 
Session("StatusMsg") = ""
Session("StatusNum") = ""
%>
</tr>
</table>

</form>
</body>
</html>
<% Else
on error resume next
SQLST = ""
SQLST = SQLST & "UPDATE FRAME SET "
SQLST = SQLST & "TYPE='" & Request.Form("TYPE") & "'"
	WhereCls = WhereCls & ", SQLSelect='" & Replace(Request.Form("SQLSelect"), "'", "''") & "'"
	WhereCls = WhereCls & ", SQLFrom='" & Replace(Request.Form("SQLFrom"), "'", "''") & "'"
	WhereCls = WhereCls & ", SQLWhere='" & Replace(Request.Form("SQLWhere"), "'", "''") & "'"
	WhereCls = WhereCls & ", SQLOrderBy='" & Replace(Request.Form("SQLOrderBy"), "'", "''") & "'"
	WhereCls = WhereCls & " WHERE FRAME_ID=" & Request.QueryString("FRAMEID")
	
	TestSQLSt = ""
	TestSQLSt = TestSQLSt & Replace(Request.Form("SQLSelect"), "'", "''") & " " & Replace(Request.Form("SQLFrom"), "'", "''") & " "
	TestSQLSt = TestSQLSt & Replace(Request.Form("SQLWhere"), "'", "''") & " " & Replace(Request.Form("SQLOrderBy"), "'", "''")
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	If Trim(TestSQLSt) <> "" Then
		QSQL = ""
		QSQL = QSQL & "{call Designer_2.CheckSQLExpression('" & TestSQLSt & "','~', '1' ,{resultset 1, StatusMsg, StatusNum})}"
		Set RS2 = Conn.Execute(QSQL)
		If 	RS2("StatusNum") <> "0" Then
			Session("StatusMsg") = RS2("StatusMsg")
			Session("StatusNum") = RS2("StatusNum")
		End If
	End If
	Set RS = Conn.Execute(SQLST & WhereCls)
	Response.Redirect "CF_SQLStatement.asp?STATUS=TRUE&FRAMEID=" & Request.QueryString("FRAMEID")
End If
%>

