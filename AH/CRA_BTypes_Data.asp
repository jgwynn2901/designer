<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
dim cAHSID, cBranchTypeID

cAHSID = Request.QueryString("AHSID") 
cBranchTypeID = Request.QueryString("EDIT")
if not isEmpty(Request.QueryString("SAVE")) then
	'	save
	response.Redirect "CRA_BTypes_Save.asp?" & Request.QueryString
end if
%>
<HTML>
<HEAD>
	<META name="VI60_defaultClientScript" content="VBScript">
	<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language=jscript>
function CRuleSearchObj()
{
	this.RID = "";
	this.Selected = false;
}

var RuleSearchObj = new CRuleSearchObj();

function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}
</script>

<SCRIPT ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub ExeSave()
	dim cHRef

	cHRef= "CRA_BTypes_Data.asp?AHSID=<%=cAHSID%>&BT_ID=<%=cBranchTypeID%>&SAVE="
	document.location.href = cHRef & "&Cov_Code=" & document.all.Cov_Code.value & "&PrTerr_Code=" & document.all.PrTerr_Code.value & "&SecTerr_Code=" & document.all.SecTerr_Code.value & "&ZIPLoc_Code=" & document.all.ZIP_Code.value & "&BOR=" & document.all.lblRuleID.innerText
End Sub

Sub AttachRule
	dim cRID
	
	cRID = document.all.lblRuleID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	RuleSearchObj.RID = cRID
	RuleSearchObj.Selected = false

	If cRID = "0" Then 
		cRID = "NEW"
	end if
	
	If cRID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit sub
	End If
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_CALLFLOW&RID=" & cRID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog strURL, RuleSearchObj, "dialogWidth=500px; center=yes"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> cRID then
			document.body.setAttribute "ScreenDirty", "YES"	
			document.all.lblRuleID.innerText = RuleSearchObj.RID
		end if
	End If

End Sub

Sub Detach
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		document.all.lblRuleID.innerText = "0"
	end if
End Sub

</SCRIPT>
</HEAD>
<%
dim oRS, cSQL, oConn
dim nCovCodeID, nPriTerrCodeID, nSecTerrCodeID, cZIPCodeLocation, cBranchOvrRuleID

cDesc = ""
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
if Request.QueryString("EDIT") <> "" then
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.MaxRecords = MAXRECORDCOUNT
	cSQL = "SELECT * FROM CRA_BRANCH_TYPES WHERE CRA_BRANCH_TYPES_ID=" & Request.QueryString("EDIT")
	oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly, adCmdText
	nCovCodeID = oRS.Fields("COVERAGE_CODE_ID").Value 
	nPriTerrCodeID = oRS.Fields("PRIMARY_TERRITORY_ID").Value 
	nSecTerrCodeID = oRS.Fields("SECONDARY_TERRITORY_ID").Value 
	cZIPCodeLocation = oRS.Fields("ZIPCODE_LOCATION").Value
	cBranchOvrRuleID = oRS.Fields("BRANCH_OVERRIDE_RULE_ID").Value
	oRS.Close
end if
%>
	<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="#d6cfbd" ScreenDirty="NO" ScreenMode="RW">
			<table class="Label" ID="Table3">
				<tr>
					<td VALIGN="CENTER" WIDTH="5">
						<img ID="StatusRpt" SRC="../images/StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
					</td>
					<td width="485">
						:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
					</td>
				</tr>
			</table>
			<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0" id="TblControls">
				<tr>
					<td>
						<table class="LABEL" ID="Table4">
							<tr>
								<td>A.H.S. ID:&nbsp;<span id="spanAHSID"><%=cAHSID%></span></td>
							</tr>
						</table>
						&nbsp
						<table ID="Table1">
							<tr>
								<td CLASS="LABEL" width="131"><font size="1">Coverage Code:</font></td>
								<td CLASS="LABEL">
									<select NAME="Cov_Code" CLASS="LABEL" STYLE="WIDTH:100" ScrnBtn="TRUE" ID="Select3">
									<%
									cSQL = "Select * From COVERAGE_CODE Where ACCNT_HRCY_STEP_ID = " & cAHSID
									Set oRS = oConn.Execute(cSQL)
									Do WHile Not oRS.EOF
										%>
										<option VALUE="<%=oRS.Fields("COVERAGE_CODE_ID").Value%>"><%=oRS.Fields("COVERAGE_CODE").Value%>
										<% 
										oRS.MoveNext
									Loop
									oRS.Close
									%>
									</select>
								</td>
							</tr>
						</table>
						&nbsp;
						<table width="380" ID="Table2">
							<tr>
								<td CLASS="LABEL" width="131"><font size="1">Branch Territories:</font></td>
								<td CLASS="LABEL" width="133">Primary:<br>
									<select NAME="PrTerr_Code" STYLE="WIDTH:100" ScrnBtn="TRUE" CLASS="LABEL" ID="Select5">
									<%
									cSQL = "Select TERRITORY_ID, TERRITORY_CD From CRA_TERRITORY_CODE"
									Set oRS = oConn.Execute(cSQL)
									Do WHile Not oRS.EOF
									%>
										<option VALUE="<%=oRS.Fields("TERRITORY_ID").Value%>"><%=oRS.Fields("TERRITORY_CD").Value%>
									<% 
										oRS.MoveNext
									Loop
									%>
									</select>
								</td>
								<td CLASS="LABEL" width="102">Secondary:<br>
									<select NAME="SecTerr_Code" STYLE="WIDTH:100" ScrnBtn="TRUE" CLASS="LABEL" ID="Select4">
										<option VALUE="">
										<% 
										oRS.MoveFirst 
										Do WHile Not oRS.EOF
										%>
											<option VALUE="<%=oRS.Fields("TERRITORY_ID").Value%>"><%=oRS.Fields("TERRITORY_CD").Value%>
										<% 
											oRS.MoveNext
										Loop
										oRS.Close
										set oRS = nothing
										oConn.Close 
										set oConn = nothing
										%>
									</select>
								</td>
							</tr>
						</table>
						&nbsp;
						<table width="380" ID="Table5">
							<tr>
								<td CLASS="LABEL" width="131"><font size="1">ZIPCode to Use:</font></td>
								<td CLASS="LABEL">
									<select NAME="ZIP_Code" CLASS="LABEL" STYLE="WIDTH:120" ScrnBtn="TRUE" ID="Select1">
										<option VALUE="LL">Loss Location</option>
										<option VALUE="RL">Risk Location</option>
										<option VALUE="EL">Employee's</option>
										<option VALUE="VL">Vehicle Location</option>
									</select> 
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<p></p>
			<table width="380" ID="Table6" height="21">
			<tr>
				<td CLASS="LABEL" width="258" height="17"><font size="1">Branch Assignment Override Rule ID:</font></td>
				<td CLASS="LABEL" style="border-style: inset; border-width: 1" width="47" height="17" align="right" id="lblRuleID">0</td>
				<td CLASS="LABEL" width="35" height="17" align="right">
					<img NAME="BtnAttachRule" STYLE="cursor:hand" SRC="..\images\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule"></td>
				<td CLASS="LABEL" width="20" height="17">
					<img NAME="BtnDetachRule" STYLE="cursor:hand" SRC="..\images\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach"></td>
			</tr>
			</table>
			
<script language=vbscript>
Sub window_onload
<%
if Request.QueryString("EDIT") <> "" then
%>
	SelectOption document.all.Cov_Code,"<%=nCovCodeID%>"
	SelectOption document.all.PrTerr_Code,"<%=nPriTerrCodeID%>"
	SelectOption document.all.SecTerr_Code,"<%=nSecTerrCodeID%>"
	SelectOption document.all.ZIP_Code,"<%=cZIPCodeLocation%>"
	<%
	if not isNull(cBranchOvrRuleID) then
	%>
	document.all.lblRuleID.innerText = "<%=cBranchOvrRuleID%>"
	<%
	end if
	%>
<%
end if
%>	
End Sub
</script>

	</body>
</HTML>
