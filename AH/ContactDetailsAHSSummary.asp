<!-- #include file="..\lib\common.inc"-->
<!-- #include file="..\lib\security.inc"-->

<%	
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	Dim Mode, ClientCode
	
	ClientCode = Request.QueryString("ClientCode")

	if Request.QueryString("MODE") <> "" then MODE = Request.QueryString("MODE")
	
	IF Request.QueryString("COID") <> "NEW" Then
		SET Conn = Server.CreateObject("ADODB.Connection")
		SET rs = Server.CreateObject("ADODB.RecordSet")
		ConnectionString = CONNECT_STRING
		S_Query = "Select * From AHS_CONTACT WHERE CONTACT_ID =" &  Request.QueryString("COID")
		
		Conn.Open ConnectionString
		rs.Open s_Query, Conn, adOpenStatic
	END IF
	IF Request.QueryString("Action") = "Delete" and IsNumeric(Request.QueryString("Delete_ID")) Then
		SET Conn = Server.CreateObject("ADODB.Connection")
		SET rs_SeqStep = Server.CreateObject("ADODB.RecordSet")
		ConnectionString = CONNECT_STRING
		S_DeleteQuery = "Delete From AHS_CONTACT WHERE AHS_CONTACT_ID = " & Request.QueryString("Delete_ID")

		'***************************************
		' DMS: 4/11/2000
		' Rather than firing a select to refresh the screen, redirect the screen to itself
		'S_Query = "Select * From AHS_CONTACT WHERE CONTACT_ID =" &  Trim(Request.QueryString("COID"))
		'rs_SeqStep.Open S_Query, Conn, adOpenStatic
		'***************************************
		Conn.Open ConnectionString
		SET rs_Delete = Conn.Execute(S_DeleteQuery)
		response.redirect("ContactDetailsaHSSummary.asp?MODE=" & Request.QueryString("MODE") & "&COID=" & Request.QueryString("COID"))
		
	End IF
	'response.write(S_Query)
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
function GetHightLightedAhsContact_ID( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("AHSCID");
}
//-->
</SCRIPT>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Function EditClick()
	i = getselectedindex( Document.all.tbl_SeqStep )
	if 0 < i then
		s_URL = "AHSContact-f.asp" & "?MODE=" & "EDIT" & "&AHSCID=" & GetHightLightedAhsContact_ID(Document.all.tbl_SeqStep.rows(i)) & "&COID=" &  "<%=Request.QueryString("COID")%>" 
		l_Ret = window.showModalDialog(s_URL,Null,"dialogWidth:450px;dialogHeight:300px;center")
		
		self.location.href = "ContactDetailsAHSSummary.asp?Mode=" & "<%=Mode%>" & "&COID=" & "<%=Request.QueryString("COID")%>" 
	end if
End Function

Function NewClick()

	If "<%=Mode%>" = "RO" Then
		MsgBox "Browse (Read Only) Mode Is Enforced.", 0, "FNSDesigner"
		Exit Function
	End If
	s_URL = "AHSContact-f.asp" & "?MODE=" & "NEW" & "&AHSCID=NEW" & "&COID=" &  "<%=Request.QueryString("COID")%>" 
	l_Ret = window.showModalDialog(s_URL,Null,"dialogWidth:450px;dialogHeight:300px;center")
	
	self.location.href = "ContactDetailsAHSSummary.asp?Mode=" & "<%=Mode%>" & "&COID=" & "<%=Request.QueryString("COID")%>" 
End Function

Function DeleteClick()
	If "<%=Mode%>" = "RO" Then
		MsgBox "Browse (Read Only) Mode Is Enforced.", 0, "FNSDesigner"
		Exit Function
	End If
	
	i = getselectedindex( Document.all.tbl_SeqStep )
	if 0 < i then
	     s_url = "ContactDetailsAHSSummary.asp?Mode=" & "<%=Mode%>" & "&Action=Delete&Delete_ID=" & GetHightLightedAhsContact_ID(Document.all.tbl_SeqStep.rows(i)) &  "&COID=" & "<%=Request.QueryString("COID")%>" 
		 self.location.href = s_url
	End If
End Function
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">

<TABLE CELLPADDING="0" CELLSPACING="0" >
  <tr><td colspan="2" HEIGHT="4"></td></tr>
  <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» AHS Contact &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp</td>
      <td HEIGHT="5" ALIGN="LEFT">
  <Table CELLPADDING="0" CELLSPACING="0">
      <tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
          <td WIDTH="175" HEIGHT="8"></td></tr>
  </Table></td></tr>
  <tr><td colspan="2" HEIGHT="1"></td></tr>
</TABLE>
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&HIDEATTACH=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=SeqStepBtnControl type=text/x-scriptlet></OBJECT>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL>Record Count is <%= rs.RecordCount %></SPAN>
 <Fieldset ID="SpDestSeqStep" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'0';width:'100%'">
  <DIV align="LEFT" id="SeqStep_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
  <TABLE cellPadding=2 cellSpacing=0 rules=all ID="tbl_SeqStep" name="tbl_SeqStep"  Width="100%">
	 <THead CLASS="ResultHeader">
		<tr align="left" >
			<td class=thd><div id="AHSCID"><NOBR>AHS ContactID</div></td>
			<td class=thd><div id="COID"><NOBR>Contact Id</div></td>
			<td class=thd><div id="AHSID"><NOBR>AHS ID</div></td>
			<td class=thd><div id="ASTDT"><NOBR>Active Start Dt</div></td>
			<td class=thd><div id="AEDDT"><NOBR>Active End Dt</div></td>
	
		</tr>
	</THead>
	<TBody ID="TableRows">
<% Do While Not rs.EOF  %>
		<tr ID="rowData" CLASS="ResultRow" AHSCID="<%=rs("AHS_CONTACT_ID")%>" COID="<%=rs("CONTACT_ID")%>" AHSID="<%=rs("ACCNT_HRCY_STEP_ID")%>" AEDDT="<%=rs("ACTIVE_END_DT")%>" ASTDT="<%=rs("ACTIVE_START_DT")%>" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);">
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs("AHS_CONTACT_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs("CONTACT_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs("ACCNT_HRCY_STEP_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs("ACTIVE_START_DT")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(rs("ACTIVE_END_DT")) %></td>
		

		</tr>
<%
rs.MoveNext
Loop
rs.Close
SET rs = NOTHING
%>
	</TBody>
</TABLE>
</DIV>
</Fieldset>
</BODY>
</HTML>
