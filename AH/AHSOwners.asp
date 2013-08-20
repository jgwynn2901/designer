<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%  Response.Expires=0

    dim oConn, oRS, cSQL, cSQL1
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.MaxRecords = MAXRECORDCOUNT
	
	cSQL = " SELECT O.* , AHSO.ACTIVE_START_DT, ACTIVE_END_DT " &_
	       "   FROM OWNER O, AHS_OWNER AHSO " &_
		   "  WHERE O.OWNER_ID           = AHSO.OWNER_ID " &_
		   "    AND AHSO.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID") &_
		   "  ORDER BY O.NAME_LAST"
	
	oRS.Open cSQL, CONNECT_STRING, adOpenStatic, adLockReadOnly, adCmdText

%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->
Sub BtnGrfxBack_Onclick()
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Sub


-->
</script>

<script LANGUAGE="jscript">
function AHSSearchObj()
{
	this.ahsid = "";
}
var oAHS = new AHSSearchObj();
</script>

</head>
<body BGCOLOR="<%= BODYBGCOLOR %>" rightmargin="0" bottommargin="0" leftmargin="0" topmargin="0">
<!--#include file="..\lib\NavBack.inc"-->
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Owners &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table CELLSPACING="0" CELLPADDING="0" WIDTH="300" BORDER="0" STYLE="BACKGROUND-COLOR:Seashell">
<tr><td CLASS="LABEL"><br></td></tr>
<tr>
<td CLASS="LABEL"><b>AHS ID: </b><%=Request.QueryString("AHSID")%></td>
</tr>
<tr><td CLASS="LABEL"><br></td></tr>
</table>

<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'90%';width:'100%'">
<!--<object data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE&amp;HIDEREFRESH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEREMOVE=TRUE&amp;HIDEPASTE=TRUE&amp;HIDEEDIT=TRUE&amp;HIDENEW=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="UserBtnControl" type="text/x-scriptlet"></object>-->
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%">
<thead CLASS="ResultHeader">
<tr>
<td CLASS="ResultHeader">ID</td>
<td CLASS="ResultHeader">Name Last</td>
<td CLASS="ResultHeader">Name First</td>
<td CLASS="ResultHeader">Title</td>
<td CLASS="ResultHeader">Work Phone</td>
<td CLASS="ResultHeader">Fax</td>
<td CLASS="ResultHeader">E-Mail</td>
<td CLASS="ResultHeader">Active Start Dt</td>
<td CLASS="ResultHeader">Active End Dt</td>
</tr>
</thead>
	<tbody ID="TableRows">
<% If oRS.EOF AND oRS.BOF Then %>
<tr ID="FieldRow" CLASS="RESULTROW">
<td CLASS="LABEL" COLSPAN="8" ALIGN=CENTER>No owners found</td>
</tr>
<%
	Else
	Do While Not oRS.EOF
%>
<tr ID="FieldRow" CLASS="RESULTROW" OnClick="Javascript:multiselect(this);" CNT="<%= oRS("OWNER_ID") %>" >
<td CLASS="ResultCell"><%= renderCell(oRS("OWNER_ID")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("NAME_LAST")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("NAME_FIRST")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("NAME_TITLE")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("PHONE_WORK")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("PHONE_FAX")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("EMAIL")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("ACTIVE_START_DT")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("ACTIVE_END_DT")) %></td>
</tr>
<% 
oRS.MoveNext
Loop
oRS.Close
End If
%>
</tbody>
</table>
</div>
</fieldset>

</body>
</html>
