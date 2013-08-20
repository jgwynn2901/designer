<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<% Response.Expires=0 
	Set RS = Server.CreateObject("ADODB.RecordSet")
	Set RS0 = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	
If Request.QueryString("DELETE") = "TRUE" Then
	SQL2 = ""
	SQL2 = SQL2 & "DELETE FROM ATTRIBUTE_OVERRIDE WHERE ATTRIBUTEOVERRIDE_ID=" & Request.QueryString("ATTRIBUTEOVERRIDE_ID")
	RS.Open SQL2, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
	Response.Redirect "Attribute_OVerride_Details.asp?ATTR_INSTANCE_ID=" & Request.QueryString("ATTR_INSTANCE_ID")
End If

SQL = ""
SQL = SQL & "SELECT * FROM ATTRIBUTE_OVERRIDE WHERE ATTR_INSTANCE_ID=" & Request.QueryString("ATTR_INSTANCE_ID") & " ORDER BY SEQUENCE"
RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText

%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub BtnGrfxBack_Onclick()
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Sub

<!--#include file="..\lib\Help.asp"-->

-->
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" FOR="UserBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "EDITBUTTONCLICK":
		idx = getselectedindex(document.all.tblResult)
		if (idx != "-1")
		{
		OVERRIDEID = document.all.tblResult.rows(idx).getAttribute("OVERRIDEID")
		self.location.href = "Attribute_Override.asp?ATTR_INSTANCE_ID=<%= Request.QueryString("ATTR_INSTANCE_ID") %>&ATTRIBUTEOVERRIDE_ID=" + OVERRIDEID
		}
			break;
		case "ATTACHBUTTONCLICK":
				idx = getselectedindex(document.all.tblResult)
				if (idx != "-1")
				{
				OVERRIDEID = document.all.tblResult.rows(idx).getAttribute("OVERRIDEID")
					lret = confirm ("Are you sure you want to delete this over ride?")
					if (lret == true)
					{
					self.location.href = "Attribute_Override_Details.asp?DELETE=TRUE&ATTR_INSTANCE_ID=<%= Request.QueryString("ATTR_INSTANCE_ID") %>&ATTRIBUTEOVERRIDE_ID="+ OVERRIDEID
					}
				}
			break;
		case "REMOVEBUTTONCLICK":
				self.location.href = "Attribute_Override.asp?ATTR_INSTANCE_ID=<%= Request.QueryString("ATTR_INSTANCE_ID") %>&ATTRIBUTEOVERRIDE_ID=NEW"
				break;
		default:
				alert("ERROR: NOT A FEATURE");
			break;
	}
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript">
function AltNameObj()
{
	this.altname = "";
}
var CAltName = new AltNameObj();
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="90" HEIGHT="10"><nobr>&nbsp;» Attribute Override Listing
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset ID="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'90%';width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEATTACH=FALSE&ATTACHCAPTION=Delete&HIDEREFRESH=TRUE&REMOVECAPTION=New&HIDENEW=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=UserBtnControl type=text/x-scriptlet></OBJECT>
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
<thead CLASS="ResultHeader">
<TR>
<TD CLASS=ResultHeader>ID</TD>
<TD CLASS=ResultHeader>Sequence</TD>
<TD CLASS=ResultHeader>Property Name</TD>
<TD CLASS=ResultHeader>Override Rule</TD>
</TR>
</THEAD>
	<tbody ID="TableRows">
<% If RS.EOF AND RS.BOF Then %>
<TR ID="FieldRow" CLASS=RESULTROW OVERRIDEID=''>
<TD CLASS=LABEL COLSPAN=5>No attribute overrides found.</TD>
</TR>
<%
	Else
	Do While Not RS.EOF
		SQL = "SELECT * FROM RULES WHERE RULE_ID=" & RS.Fields("OVERRIDE_RULE_ID").Value 
		RS0.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
%>
<TR ID="FieldRow" CLASS=RESULTROW OnClick="Javascript:multiselect(this);" OVERRIDEID='<%= RS("ATTRIBUTEOVERRIDE_ID") %>'>
<TD CLASS=ResultCell><%= renderCell(RS("ATTRIBUTEOVERRIDE_ID")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS("SEQUENCE")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS("PROPERTY_NAME")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS0("RULE_TEXT")) %></TD>
</TR>
<% 
		RS0.Close 
		RS.MoveNext
Loop
RS.Close
End If
%>
</tbody>
</TABLE>
</DIV>
</fieldset>
</BODY>
</HTML>
