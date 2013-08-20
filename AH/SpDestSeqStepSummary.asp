<!-- #include file="..\lib\common.inc"-->
<!-- #include file="..\lib\security.inc"-->

<%	
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	If HasModifyPrivilege("FNSD_SPECIFIC_DESTINATION", SECURITYPRIV) <> True Then MODE = "RO"
	Dim Mode
	if Request.QueryString("MODE") <> "" then MODE = Request.QueryString("MODE")
	Mode = Request.QueryString("MODE")
	IF Request.QueryString("SPDEST_ID") <> "NEW" Then
		SET Conn = Server.CreateObject("ADODB.Connection")
		SET rs_SeqStep = Server.CreateObject("ADODB.RecordSet")
		ConnectionString = CONNECT_STRING
		S_Query = "Select * From SPECIFIC_DESTN_SEQ_STEP WHERE SPECIFIC_DESTINATION_ID = " & Trim(Request.QueryString("SPDEST_ID"))& " Order By SEQUENCE"
		Conn.Open ConnectionString
		rs_SeqStep.Open s_Query, Conn, adOpenStatic
	END IF
	IF Request.QueryString("Action") = "Delete" and IsNumeric(Request.QueryString("Delete_ID")) Then
		SET Conn = Server.CreateObject("ADODB.Connection")
		SET rs_SeqStep = Server.CreateObject("ADODB.RecordSet")
		ConnectionString = CONNECT_STRING
		S_DeleteQuery = "Delete From SPECIFIC_DESTN_SEQ_STEP WHERE SPECIFIC_DESTN_SEQ_STEP_ID = " & Request.QueryString("Delete_ID")
		S_Query = "Select * From SPECIFIC_DESTN_SEQ_STEP WHERE SPECIFIC_DESTINATION_ID = " & Trim(Request.QueryString("SPDEST_ID")) & " Order By SEQUENCE"
		Conn.Open ConnectionString
		SET rs_Delete = Conn.Execute(S_DeleteQuery)
		rs_SeqStep.Open s_Query, Conn, adOpenStatic
	End IF
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Link Rel="StyleSheet" Type="text/css" Href="..\FNSDESIGN.CSS">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="JavaScript" FOR="SeqStepBtnControl" EVENT="onscriptletevent (event, obj)">
<!--
   switch (event)
	{
	case "EDITBUTTONCLICK":
       	EditClick();
		break;

	case "NEWBUTTONCLICK":
		NewClick();
		break;

	case "REMOVEBUTTONCLICK":
		DeleteClick();
		break;

	default:
		break;
}
//-->
</SCRIPT>
<SCRIPT LANGUAGE="Javascript">
<!--
function dblclick( objRow )
{
EditClick();
}
function GetHightLightedSeqStep_ID( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("SDSSID");
}
function GetHightLightedDestination_ID( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("SDID");
}
//-->
</SCRIPT>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Function EditClick()
	i = getselectedindex( Document.all.tbl_SeqStep )
	if 0 < i then
		s_URL = "SpDestSeqStep.asp?MODE=" & "<%=Mode%>" & "&SeqStep_ID=" & GetHightLightedSeqStep_ID(Document.all.tbl_SeqStep.rows(i)) & "&SDID=" & GetHightLightedDestination_ID(Document.all.tbl_SeqStep.rows(i))
		l_Ret = window.showModalDialog(s_URL,Null,"dialogWidth:450px;dialogHeight:300px;center")
''		l_Ret = window.open(s_URL)
		self.location.href = "SpDestSeqStepSummary.asp?MODE=" & "<%=Mode%>" & "&SPDEST_ID=" & "<%=Request.QueryString("SPDEST_ID")%>"
	end if
End Function

Function NewClick()
	If "<%=Mode%>" = "RO" Then
		MsgBox "Browse (Read Only) Mode Is Enforced.", 0, "FNSDesigner"
		Exit Function
	End If
	l_Ret = window.showModalDialog("SpDestSeqStep.asp?Mode=" & "<%=Mode%>" & "&SeqStep_ID=NEW" & "&SDID=" & "<%=Request.QueryString("SPDEST_ID")%>",Null,"dialogWidth:450px;dialogHeight:300px;center")
	self.location.href = "SpDestSeqStepSummary.asp?Mode=" & "<%=Mode%>" & "&SPDEST_ID=" & "<%=Request.QueryString("SPDEST_ID")%>"
End Function

Function DeleteClick()
	If "<%=Mode%>" = "RO" Then
		MsgBox "Browse (Read Only) Mode Is Enforced.", 0, "FNSDesigner"
		Exit Function
	End If
	IF tbl_SeqStep.rows.length <= 2 Then
		MsgBox "Specific Destination Must Have At Least One Sequence Step Record Attached.", ,"FNSDesigner"
		Exit Function
	End IF
	i = getselectedindex( Document.all.tbl_SeqStep )
	if 0 < i then
		self.location.href = "SpDestSeqStepSummary.asp?SPDEST_ID=" & "<%=Request.QueryString("SPDEST_ID")%>" & "&Action=Delete&Delete_ID=" & GetHightLightedSeqStep_ID(Document.all.tbl_SeqStep.rows(i))
	End If
End Function
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">

<TABLE CELLPADDING="0" CELLSPACING="0" >
  <tr><td colspan="2" HEIGHT="4"></td></tr>
  <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Specific Destination Sequence Step&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp</td>
      <td HEIGHT="5" ALIGN="LEFT">
  <Table CELLPADDING="0" CELLSPACING="0">
      <tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
          <td WIDTH="175" HEIGHT="8"></td></tr>
  </Table></td></tr>
  <tr><td colspan="2" HEIGHT="1"></td></tr>
</TABLE>
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&HIDEATTACH=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=SeqStepBtnControl type=text/x-scriptlet></OBJECT>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL>Record Count is <%= rs_SeqStep.RecordCount %></SPAN>
 <Fieldset ID="SpDestSeqStep" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'0';width:'100%'">
  <DIV align="LEFT" id="SeqStep_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
  <TABLE cellPadding=2 cellSpacing=0 rules=all ID="tbl_SeqStep" name="tbl_SeqStep"  Width="100%">
	 <THead CLASS="ResultHeader">
		<tr align="left" >
			<td class=thd><div id="SDSSID"><NOBR>S.D.S.S ID</div></td>
			<td class=thd><div id="SEQ"><NOBR>Seq. #</div></td>
			<td class=thd><div id="RETRY"><NOBR>Retry #</div></td>
			<td class=thd><div id="RETRYWAITTIME"><NOBR>Re. Wait Time</div></td>
			<td class=thd><div id="DESTINATION"><NOBR>Destination</div></td>
			<td class=thd><div id="TRANSID"><NOBR>Trans. ID</div></td>
		</tr>
	</THead>
	<TBody ID="TableRows">
<% Do While Not rs_SeqStep.EOF  %>
		<tr ID="rowData" CLASS="ResultRow" SDSSID="<%=rs_SeqStep("SPECIFIC_DESTN_SEQ_STEP_ID")%>" SDID="<%=rs_SeqStep("SPECIFIC_DESTINATION_ID")%>" SEQ="<%=rs_SeqStep("SEQUENCE")%>" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);">
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs_SeqStep("SPECIFIC_DESTN_SEQ_STEP_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs_SeqStep("SEQUENCE")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs_SeqStep("RETRY_COUNT")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs_SeqStep("RETRY_WAIT_TIME")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs_SeqStep("DESTINATION_STRING")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs_SeqStep("TRANSMISSION_TYPE_ID")) %></td>
		</tr>
<%
rs_SeqStep.MoveNext
Loop
rs_SeqStep.Close
SET rs_SeqStep = NOTHING
%>
	</TBody>
</TABLE>
</DIV>
</Fieldset>
</BODY>
</HTML>
