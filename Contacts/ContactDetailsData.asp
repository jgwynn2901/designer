<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
	cAHSID = Request.QueryString("AHSID")
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Contact Details Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script language="javascript">

var g_StatusInfoAvailable = false;

function CContactDataObj()
{
	this.Selected = false;
}

var ContactDataObj = new CContactDataObj();
</script>

<script LANGUAGE="JScript" FOR="UserBtnControl" EVENT="onscriptletevent (event, obj)">
	var nNtx, cURL, nOpt;

	switch (event)
	{
		case "EDITBUTTONCLICK":
			nNtx = getSelectedCID();
			if (nNtx == 0)
				alert ("Please select a Contact to edit.")
			else
				{
				ContactDataObj.Selected = false;
				strURL = "ContactsModal.asp?CID=" + nNtx + "&AHSID=<%=Request.QueryString("AHSID")%>";
				window.showModalDialog (strURL, ContactDataObj, "center:yes;status:no;help:no");
				if (ContactDataObj.Selected == true)
					Refresh();
				}
			break;
		case "NEWBUTTONCLICK":
			ContactDataObj.Selected = false;
			strURL = "ContactsModal.asp?CID=NEW" + "&AHSID=<%=Request.QueryString("AHSID")%>";
			window.showModalDialog (strURL, ContactDataObj, "center:yes;status:no;help:no");
			if (ContactDataObj.Selected == true)
				Refresh();
			break;
			
		case "REMOVEBUTTONCLICK":
			nNtx = getSelectedCID();
			if (nNtx != 0)
				{
				self.location.href = "ContactSave.asp?DELETE=" + nNtx + "&AHSID=<%=cAHSID%>"
				}
			else
				alert ("Please select a Contact to delete.");
			break;
	}
</script>

<script language="VBScript">
sub window_onload
'stop
end sub

Sub Refresh
	cAHSID = <%=Request.QueryString("AHSID")%>
	self.location.href = "ContactDetailsData.asp?AHSID=" & cAHSID
End Sub

Sub BtnGrfxBack_Onclick()
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Sub

Function GetSelectedCID
	dim idx
	
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedCID = document.all.tblFields.rows(idx).getAttribute("CID")
	Else
		GetSelectedCID = ""
	End If
End Function
</script>
<!--#include file="..\lib\tablecommon.inc"-->
</head>
<body BGCOLOR="<%= BODYBGCOLOR %>" rightmargin="0" bottommargin="0" leftmargin="0" topmargin="0">
<!--#include file="..\lib\NavBack.inc"-->
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Contacts &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table CELLSPACING="0" CELLPADDING="0" WIDTH="300" BORDER="0" STYLE="BACKGROUND-COLOR:Seashell" ID="Table3">
<tr><td CLASS="LABEL"><br></td></tr>
<tr>
<td CLASS="LABEL"><b>AHS ID: </b><%=Request.QueryString("AHSID")%></td>
</tr>
<tr><td CLASS="LABEL"><br></td></tr>
</table>
<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'90%';width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE&amp;HIDEREFRESH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="UserBtnControl" type="text/x-scriptlet" VIEWASTEXT></object>
<table cellPadding="2" rules=all  cellSpacing="0" scrolling="auto" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
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

<%
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL = "SELECT * FROM CONTACT WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID") & " ORDER BY NAME"
	Set oRS = oConn.Execute(cSQL)
	If oRS.EOF Then %>
<tr ID="FieldRow" CLASS="RESULTROW">
<td CLASS="LABEL" COLSPAN="8" ALIGN=CENTER>No Contacts found</td>
</tr>
<%
	Else	
	Do While Not oRS.EOF
%>
	<tr ID="FieldRow" CLASS="RESULTROW" OnClick="Javascript:multiselect(this);" CID="<%= oRS("CONTACT_ID") %>">
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
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	End If
%>
</tbody>
</table>
</fieldset>
</BODY>
</HTML>


