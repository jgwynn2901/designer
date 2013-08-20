<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->  
<!--#include file="..\lib\AHSTree.inc"--> 
<% Response.Expires=0 
   Response.Buffer = True
   On Error Resume Next
%>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

If Request.QueryString("DELETED") <> "" Then
  
	SQLDEL=  "DELETE FROM FRAUD_DETECTION_RULE WHERE FRAUD_DETECTION_TYPE_ID = " & Request.QueryString("DELETED")
	Set RS = Conn.Execute(SQLDEL)
	SQLDEL1=  "DELETE FROM FRAUD_DETECTION_TYPE WHERE FRAUD_DETECTION_TYPE_ID = " & Request.QueryString("DELETED")
	Set RS = Conn.Execute(SQLDEL1)
	strError = CheckADOErrors(Conn,"DELETE" )
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing
	
	If strError = "" Then 
		Response.Redirect "FraudDetecSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If

End If

Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.MaxRecords = MAXRECORDCOUNT
ConnectionString = CONNECT_STRING
			
cSQL = "SELECT * FROM FRAUD_DETECTION_TYPE FDT WHERE " &_
		 "FDT.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	
oRS.Open cSQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="Javascript">
function dblclick( objRow )
{
	EditClick()
}
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("FDTID");
}
</SCRIPT>
<!-- #include file="..\lib\BRBtnControl.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../FraudDetection/FDMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&FDTID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function NewClick()
    parent.frames.window.location = "../FraudDetection/FDMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&FDTID=NEW&DETAILONLY=TRUE&AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then

		FDTID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to delete Fraud Detection Type ID:" & FDTID & " for Account ID:<%=Request.QueryString("AHSID")%>?", 1, "FNSDesigner")
		If lret = 1 Then
	       self.location.href = "FraudDetecSummary.asp?DELETED=" & FDTID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
		
	end if
End Function

Sub window_onload
<% If oRS.RecordCount = MAXRECORDCOUNT Then %>
	StatusSpan.innerHTML = "<%= MSG_MAXRECORDS %>"
<% Else %>
	StatusSpan.innerHTML = "Record Count is <%= oRS.RecordCount %>"
<% End If %>	

<%If strError <> "" Then %>
	MsgBox ("<%=strError%>")
<% End If %>

End Sub
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
PARAMS = "&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=TRUE&HIDESEARCH=TRUE"
%>
<OBJECT id="BRBtnControl" type=text/x-scriptlet data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0">
	</OBJECT>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Branch_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0  rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div><NOBR>FDT ID</div></td>
			<td class=thd><div><NOBR>LOB</div></td>
			<td class=thd><div><NOBR>Description</div></td>
			<td class=thd><div><NOBR>Threshold</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not oRS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" FDTID='<%= oRS("FRAUD_DETECTION_TYPE_ID") %>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(oRS("FRAUD_DETECTION_TYPE_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(oRS("LOB_CD")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(oRS("DESCRIPTION")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(oRS("THRESHOLD")) %></td>			
	</tr>
<% 
oRS.MoveNext
Loop
oRS.Close
Set oRS = Nothing
Conn.Close
Set Conn = Nothing
%>
</tbody>
</table>
</DIV>
</FIELDSET>
</SCRIPT>
</BODY>
</HTML>
