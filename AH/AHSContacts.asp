<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<% Response.Expires=0

	dim oConn, oRS, cSQL
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.MaxRecords = MAXRECORDCOUNT

	if Request.Querystring("DEL") <> "" then
		cSQL = "DELETE FROM CONTACT WHERE CONTACT_ID=" & Request.querystring("DEL")
		oConn.Execute (cSQL )
		oConn.Close 
		set oConn = nothing
	end if
	cSQL = "SELECT * FROM CONTACT WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID") & " ORDER BY NAME"
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

Function getSelectedID
	dim nIdx
	
	nIdx = getselectedindex(document.all.tblResult)
	If nIdx <> -1 Then
		getSelectedID = CLng(document.all.tblResult.rows(nIdx).getAttribute("CNT"))
	Else
		getSelectedID = 0
	End If
End Function

-->
</script>

<script LANGUAGE="jscript">
function AHSSearchObj()
{
	this.ahsid = "";
}
var oAHS = new AHSSearchObj();
</script>

<script LANGUAGE="JavaScript" FOR="UserBtnControl" EVENT="onscriptletevent (event, obj)">
	var nNtx, cURL, nOpt;
	
	switch (event)
	{
		case "EDITBUTTONCLICK":
			nNtx = getSelectedID();
			if (nNtx == 0)
				alert ("Please select a Contact to edit.")
			else
				{
				window.showModalDialog ('AddEditContact.asp?EDIT=' + nNtx + '&AHSID=<%= Request.QueryString("AHSID") %>', null, 'dialogWidth=500px; dialogHeight=350px; center=yes');
				self.location.href = "AHSContacts.asp?AHSID=<%= Request.QueryString("AHSID") %>"
				}
			break;
		case "NEWBUTTONCLICK":
			window.showModalDialog ("AddEditContact.asp?AHSID=<%= Request.QueryString("AHSID") %>", 0, "dialogWidth:500px; dialogHeight:350px; center:yes");
			self.location.href = "AHSContacts.asp?AHSID=<%= Request.QueryString("AHSID") %>"
			break;
			
		case "REMOVEBUTTONCLICK":
			nNtx = getSelectedID();
			if (nNtx == 0)
				alert ("Please select a Contact to delete.");
			else
				if (confirm("Are you sure you want to delete this Contact?"))
					self.location.href = "AHSContacts.asp?DEL=" + nNtx + "&AHSID=<%= Request.QueryString("AHSID") %>";
			break;
	
		default:
				alert("NOT A FEATURE");
			break;
	}
</script>
</head>
<body BGCOLOR="<%= BODYBGCOLOR %>" rightmargin="0" bottommargin="0" leftmargin="0" topmargin="0">
<!--#include file="..\lib\NavBack.inc"-->
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Contacts &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<object data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE&amp;HIDEREFRESH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="UserBtnControl" type="text/x-scriptlet"></object>
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%">
<thead CLASS="ResultHeader">
<tr>
<td CLASS="ResultHeader">ID</td>
<td CLASS="ResultHeader">Type</td>
<td CLASS="ResultHeader">Name</td>
<td CLASS="ResultHeader">Title</td>
<td CLASS="ResultHeader">Phone</td>
<td CLASS="ResultHeader">Fax</td>
<td CLASS="ResultHeader">E-Mail</td>
<td CLASS="ResultHeader">Description</td>
</tr>
</thead>
	<tbody ID="TableRows">
<% If oRS.EOF AND oRS.BOF Then %>
<tr ID="FieldRow" CLASS="RESULTROW">
<td CLASS="LABEL" COLSPAN="8" ALIGN=CENTER>No Contacts found</td>
</tr>
<%
	Else
	Do While Not oRS.EOF
%>
<tr ID="FieldRow" CLASS="RESULTROW" OnClick="Javascript:multiselect(this);" CNT="<%= oRS("CONTACT_ID") %>">
<td CLASS="ResultCell"><%= renderCell(oRS("CONTACT_ID")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("TYPE")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("NAME")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("TITLE")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("PHONE")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("FAX")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("EMAIL")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("DESCRIPTION")) %></td>
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
