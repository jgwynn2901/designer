<!--#include file="..\lib\common.inc"-->
<%= Response.Expires=0 %>
<%
	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = 30
	ConnectionString = CONNECT_STRING
	SQL = ""
	SQL = SQL & "SELECT * FROM POLICY WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="Javascript">
<!--
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	GetSelection( objRow.getAttribute("DYNKEY") )
}
-->
</SCRIPT>
<!-- #include file="..\lib\PolicyBtnControl.inc" -->
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<OBJECT data="../Scriptlets/ObjButtons.htm" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="PolicyBtnControl" Name="PolicyBtnControl" type=text/x-scriptlet></OBJECT>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:90;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 frame=void rules=all ID="tblFields" name="tblFields" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>Policy Number</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>LOB</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>Description</div></td>
		</tr>
	</thead>
	
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>	
		<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
			<td NOWRAP CLASS=LABEL ><%= RS("POLICY_NUMBER") %></td>
			<td NOWRAP CLASS=LABEL><%= RS("LOB_CD") %></td>
			<td NOWRAP CLASS=LABEL><%= RS("POLICY_DESC") %></td>
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
