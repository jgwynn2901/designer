<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<% Response.Expires=0

	dim oConn, oRS, cSQL, cCCID, cStates, cSC, lisFirst, cStateList
	dim oCmd, cAHSID, oRS0, cLOB2Excl
	
	cAHSID = Request.QueryString("AHSID")
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.MaxRecords = MAXRECORDCOUNT

	if Request.Querystring("DEL") <> "" then
		cCCID = Request.Querystring("DEL")
		cSQL = "Delete From BENEFIT_STATE Where COVERAGE_CODE_ID=" & cCCID
		oConn.Execute cSQL
		cSQL = "Delete From SPECIAL_COMMENTS Where COVERAGE_CODE_ID=" & cCCID
		oConn.Execute cSQL
		cSQL = "Delete From BENEFIT_STATE Where COVERAGE_CODE_ID=" & cCCID
		oConn.Execute cSQL
	end if

	cSQL = "Select distinct LOB_CD, COVERAGE_CODE From COVERAGE_CODE Where ACCNT_HRCY_STEP_ID=" & cAHSID
	set oRS0 = oConn.Execute( cSQL )
	cLOB2Excl = ""
	do while not oRS0.eof
		if len(cLOB2Excl) <> 0 then
			cLOB2Excl = cLOB2Excl & ","
		end if
		cLOB2Excl = cLOB2Excl & oRS0("LOB_CD") & "|" & oRS0("COVERAGE_CODE")
		oRS0.movenext
	loop
	oRS0.close
	set oRS0 = nothing
	cSQL = "Select CC.COVERAGE_CODE_ID, CC.LOB_CD, CC.COVERAGE_CODE, SC.COMMENTS, BS.STATE "
	cSQL = cSQL & "From BENEFIT_STATE BS, COVERAGE_CODE CC, SPECIAL_COMMENTS SC "
	cSQL = cSQL & "Where CC.ACCNT_HRCY_STEP_ID=" & cAHSID
	cSQL = cSQL & " AND CC.COVERAGE_CODE_ID = BS.COVERAGE_CODE_ID(+) "
	cSQL = cSQL & "AND CC.COVERAGE_CODE_ID = SC.COVERAGE_CODE_ID(+) "
	cSQL = cSQL & "Order BY LOB_CD, COVERAGE_CODE, STATE"
	set oRS = oConn.Execute( cSQL )
%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
sub window_onload
'stop
end sub

Sub BtnGrfxBack_Onclick()
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Sub

Function getSelectedAttr(cAttribute)
	dim nIdx
	
	nIdx = getselectedindex(document.all.tblResult)
	If nIdx <> -1 Then
		getSelectedAttr = document.all.tblResult.rows(nIdx).getAttribute(cAttribute)
	Else
		getSelectedAttr = ""
	End If
End Function

</script>

<script LANGUAGE="JavaScript" FOR="UserBtnControl" EVENT="onscriptletevent (event, obj)">
	var cNtx, cURL, cLOB, cCC, cStates, cSC;
	
	switch (event)
	{
		case "EDITBUTTONCLICK":
			cNtx = getSelectedAttr("CC_ID");
			cLOB = getSelectedAttr("LOB");
			cCC = getSelectedAttr("CC");
			cStates = getSelectedAttr("STATES");
			cSC = getSelectedAttr("SC");
			if (cNtx == "")
				alert ("Please select a Coverage Type to edit.")
			else
				{
				window.showModalDialog ('CRA_AddEditCT.asp?EDIT=' + cNtx + '&AHSID=<%= Request.QueryString("AHSID") %>' + '&LOB=' + cLOB + "&CC=" + cCC + "&STATES=" + cStates + "&SC=" + cSC, null, 'dialogWidth=500px; dialogHeight=510px; center=yes');
				self.location.href = "CRA_CoverTy.asp?AHSID=<%= Request.QueryString("AHSID") %>"
				}
			break;
		case "NEWBUTTONCLICK":
			window.showModalDialog ("CRA_AddEditCT.asp?AHSID=<%= Request.QueryString("AHSID") %>" + "&EXCL=" + "<%=cLOB2Excl%>", 0, "dialogWidth:500px; dialogHeight:510px; center:yes");
			self.location.href = "CRA_CoverTy.asp?AHSID=<%= Request.QueryString("AHSID") %>"
			break;
			
		case "REMOVEBUTTONCLICK":
			cNtx = getSelectedAttr("CC_ID");
			if (cNtx == "")
				alert ("Please select a Coverage Type to delete.");
			else
				if (confirm("Are you sure you want to delete this Coverage Type?"))
					self.location.href = "CRA_CoverTy.asp?DEL=" + cNtx + "&AHSID=<%= Request.QueryString("AHSID") %>";
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Crawford Coverage Types  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<object data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE&amp;HIDEREFRESH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="UserBtnControl" type="text/x-scriptlet" VIEWASTEXT></object>
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%">
<thead CLASS="ResultHeader">
<tr>
<td CLASS="ResultHeader">ID</td>
<td CLASS="ResultHeader">LOB</td>
<td CLASS="ResultHeader">Coverage Code</td>
<td CLASS="ResultHeader">States</td>
<td CLASS="ResultHeader">Special Comments</td>
</tr>
</thead>
	<tbody ID="TableRows">
<% If oRS.EOF Then %>
<tr ID="FieldRow" CLASS="RESULTROW">
<td CLASS="LABEL" COLSPAN="8" ALIGN=CENTER>No Crawford Coverage Types found</td>
</tr>
<%
	Else
	Do While Not oRS.EOF
		cCCID = oRS("COVERAGE_CODE_ID")
		cLOB_CD = oRS("LOB_CD")
		cCovCode = oRS("COVERAGE_CODE")
		cStates = ""
		cStateList = ""
		lisFirst = true
		cSC = oRS("COMMENTS")
		'
		Do While Not oRS.EOF 
			if cint(oRS("COVERAGE_CODE_ID")) <> cint(cCCID) then
				exit do
			end if
			if not lisFirst then
				cStates = cStates & ", "
				cStateList = cStateList & "||"
			else
				lisFirst = false
			end if
			cStates = cStates & oRS("STATE")
			cStateList = cStateList & oRS("STATE")
			oRS.MoveNext 
		loop
%>
		<tr ID="FieldRow" CLASS="RESULTROW" OnClick="Javascript:multiselect(this);" CC_ID="<%=cCCID%>" LOB="<%=cLOB_CD%>" CC="<%=cCovCode%>" STATES="<%=cStateList%>" SC="<%=cSC%>">
		<td CLASS="ResultCell"><%= renderCell(cCCID) %></td>
		<td CLASS="ResultCell"><%= renderCell(cLOB_CD) %></td>
		<td CLASS="ResultCell"><%= renderCell(cCovCode) %></td>
		<td CLASS="ResultCell"><%= renderCell(cStates) %></td>
		<td CLASS="ResultCell"><%= renderCell(cSC) %></td>
</tr>
<% 
	Loop
	oRS.Close
End If
oConn.Close 
%>
</tbody>
</table>
</div>
</fieldset>
</body>
</html>
