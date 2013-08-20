<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<% Response.Expires=0

	dim oConn, oRS, cSQL
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.MaxRecords = MAXRECORDCOUNT

	if Request.Querystring("DEL") <> "" then
		cSQL = "DELETE FROM CRA_BRANCH_TYPES WHERE CRA_BRANCH_TYPES_ID=" & Request.querystring("DEL")
		oConn.Execute (cSQL )
		oConn.Close 
		set oConn = nothing
	end if
	cSQL = "Select CBT.*, CC.*, TERRCD_A.TERRITORY_CD PR_TERR, TERRCD_B.TERRITORY_CD SE_TERR From CRA_BRANCH_TYPES CBT, COVERAGE_CODE CC, CRA_TERRITORY_CODE TERRCD_A, CRA_TERRITORY_CODE TERRCD_B Where CBT.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	cSQL = cSQL & " AND CBT.PRIMARY_TERRITORY_ID = TERRCD_A.TERRITORY_ID "
	cSQL = cSQL & " AND CBT.SECONDARY_TERRITORY_ID = TERRCD_B.TERRITORY_ID(+) "
	cSQL = cSQL & " AND CBT.COVERAGE_CODE_ID = CC.COVERAGE_CODE_ID"
	oRS.Open cSQL, CONNECT_STRING, adOpenStatic, adLockReadOnly, adCmdText
	
function decodeZIPLoc(cCode)
select case cCode
	case "LL"
		decodeZIPLoc = "Loss Location"
	case "RL"
		decodeZIPLoc = "Risk Location"
	case "EL"
		decodeZIPLoc = "Employee's"
	case "VL"
		decodeZIPLoc = "Vehicle Location"
	case else
		decodeZIPLoc = renderCell(cCode)
end select
end function	
%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub BtnGrfxBack_Onclick()
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Sub

Function getSelectedID
	dim nIdx
	
	nIdx = getselectedindex(document.all.tblResult)
	If nIdx <> -1 Then
		getSelectedID = CLng(document.all.tblResult.rows(nIdx).getAttribute("CBT"))
	Else
		getSelectedID = 0
	End If
End Function

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
				alert ("Please select a Branch Type to edit.")
			else
				{
				window.showModalDialog ('CRA_AddEditBT.asp?EDIT=' + nNtx + '&AHSID=<%= Request.QueryString("AHSID") %>', null, 'dialogWidth=500px; dialogHeight=350px; center=yes');
				self.location.href = "CRA_BranchTy.asp?AHSID=<%= Request.QueryString("AHSID") %>"
				}
			break;
		case "NEWBUTTONCLICK":
			window.showModalDialog ("CRA_AddEditBT.asp?AHSID=<%= Request.QueryString("AHSID") %>", 0, "dialogWidth:500px; dialogHeight:450px; center:yes");
			self.location.href = "CRA_BranchTy.asp?AHSID=<%= Request.QueryString("AHSID") %>"
			break;
			
		case "REMOVEBUTTONCLICK":
			nNtx = getSelectedID();
			if (nNtx == 0)
				alert ("Please select a Branch Type to delete.");
			else
				if (confirm("Are you sure you want to delete this Branch Type?"))
					self.location.href = "CRA_BranchTy.asp?DEL=" + nNtx + "&AHSID=<%= Request.QueryString("AHSID") %>";
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Crawford Branch Types  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<td CLASS="ResultHeader">Coverage Code</td>
<td CLASS="ResultHeader">Primary Terr. Code</td>
<td CLASS="ResultHeader">Secondary Terr. Code</td>
<td CLASS="ResultHeader">Zip Code Location</td>
<td CLASS="ResultHeader">Branch Override Rule</td>
</tr>
</thead>
	<tbody ID="TableRows">
<% If oRS.EOF AND oRS.BOF Then %>
<tr ID="FieldRow" CLASS="RESULTROW">
<td CLASS="LABEL" COLSPAN="8" ALIGN=CENTER>No Crawford Branch Types found</td>
</tr>
<%
	Else
	Do While Not oRS.EOF
%>
<tr ID="FieldRow" CLASS="RESULTROW" OnClick="Javascript:multiselect(this);" CBT="<%= oRS("CRA_BRANCH_TYPES_ID") %>">
<td CLASS="ResultCell"><%= renderCell(oRS("CRA_BRANCH_TYPES_ID")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("COVERAGE_CODE")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("PR_TERR")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("SE_TERR")) %></td>
<td CLASS="ResultCell"><%= decodeZIPLoc(oRS("ZIPCODE_LOCATION")) %></td>
<td CLASS="ResultCell"><%= renderCell(oRS("BRANCH_OVERRIDE_RULE_ID")) %></td>
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
