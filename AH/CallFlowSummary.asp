<!--#include file="..\lib\common.inc"-->
<%= Response.Expires=0 %>
<%
	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = 30
	ConnectionString = CONNECT_STRING
	SQL = ""
	SQL = SQL & "SELECT ACCOUNT_CALLFLOW.*, CALLFLOW.* FROM ACCOUNT_CALLFLOW, CALLFLOW WHERE ACCOUNT_CALLFLOW.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID") 
	SQL = SQL & " AND ACCOUNT_CALLFLOW.CALLFLOW_ID = CALLFLOW.CALLFLOW_ID "
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="Javascript">
<!--
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("CALLFLOWID");
}
-->
</SCRIPT>
<!-- #include file="..\lib\CFBtnControl.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../CallFlow/CallFlowEditor.asp?CALLFLOW=" & dblhighlight(Document.all.tblResult.rows(i))
	end if
End Function

-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>

<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<OBJECT data="../Scriptlets/ObjButtons.htm" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=CFBtnControl type=text/x-scriptlet></OBJECT>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:80;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 frame=void rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>Call Flow ID</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Name</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Description</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>LOB</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
		<tr ID="FieldRow" CLASS="ResultRow" CALLFLOWID="<%= RS("CALLFLOW_ID") %>" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL><%= RS("CALLFLOW_ID") %></td>
			<td NOWRAP CLASS=LABEL><%= RS("NAME") %></td>
			<td NOWRAP CLASS=LABEL><%= RS("DESCRIPTION") %></td>
			<td NOWRAP CLASS=LABEL><%= RS("LOB_CD") %></td>
		</tr>
<%
RS.MoveNext
Loop
RS.Close
%>
	</tbody>
</table>
</DIV>
</FIELDSET>

</BODY>
</HTML>
