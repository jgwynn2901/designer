<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<% Response.Expires=0 %>
<%

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	
If Request.QueryString("DELETE") = "TRUE" Then
	SQL2 = ""
	SQL2 = SQL2 & "DELETE FROM ALTERNATE_NAME WHERE ALTERNATE_NAME_ID=" & Request.QueryString("ID")
	RS.Open SQL2, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
End If

SQL = ""
SQL = SQL & "SELECT * FROM ALTERNATE_NAME WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID") & " ORDER BY NAME"
RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText

%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
<!--#include file="..\lib\Help.asp"-->
Sub BtnGrfxBack_Onclick()
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Sub

Function GetSelectedName
	dim idx
	idx = CInt(getselectedindex(document.all.tblResult))
	If idx <> -1 Then
		GetSelectedName = document.all.tblResult.rows(idx).getAttribute("ALTNAME")
	Else
		GetSelectedName = ""
	End If
End Function

Function MsgDelete (ID)
Lret = msgbox ("Are you sure you want to delete Alternate Name: " & ID, 1, "FNSDesigner")
if Lret = "1" Then
	self.location.href = "AHS_ALT_Name.asp?AHSID=<%= Request.Querystring("AHSID") %>&DELETE=TRUE&ID=" & ID
End If
End Function
-->
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" FOR="UserBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "ATTACHBUTTONCLICK":
				lret = GetSelectedName();
				if (lret == "")
				{
				alert ("Please select a row")
				}
				else
				{
				tst = MsgDelete(lret);
				}
			break;
		case "REMOVEBUTTONCLICK":
				CAltName.altname = "";
				window.showModalDialog ("Add_Alt_Name.asp?AHSID=<%= Request.QueryString("AHSID") %>", CAltName, "dialogWidth=500px; dialogHeight=180px; center=yes");
				self.location.href = "AHS_ALT_Name.asp?AHSID=<%= Request.QueryString("AHSID") %>"
			break;
		default:
				alert("NOT A FEATURE");
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
<!--#include file="..\lib\NavBack.inc"-->
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Alternate Names &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
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
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEATTACH=FALSE&ATTACHCAPTION=Delete&HIDEREFRESH=TRUE&REMOVECAPTION=New&HIDENEW=TRUE&HIDEEDIT=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=UserBtnControl type=text/x-scriptlet></OBJECT>
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
<thead CLASS="ResultHeader">
<TR>
<TD CLASS=ResultHeader>ID</TD>
<TD CLASS=ResultHeader>Name</TD>
</TR>
</THEAD>
	<tbody ID="TableRows">
<% If RS.EOF AND RS.BOF Then %>
<TR ID="FieldRow" CLASS=RESULTROW ALTNAME=''>
<TD CLASS=LABEL COLSPAN=5>No Alternate Names Found</TD>
</TR>
<%
	Else
	Do While Not RS.EOF
%>
<TR ID="FieldRow" CLASS=RESULTROW OnClick="Javascript:multiselect(this);" ALTNAME='<%= RS("ALTERNATE_NAME_ID") %>'>
<TD CLASS=ResultCell><%= RS("ALTERNATE_NAME_ID") %></TD>
<TD CLASS=ResultCell><%= RS("NAME") %></TD>
</TR>
<% 
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
