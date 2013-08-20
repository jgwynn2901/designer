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
    Dim SQLDEL,SQLDEL1,RS
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
		
If Request.QueryString("DELETED") <> "" Then
  
	SQLDEL1=  "DELETE FROM ACCOUNT_TIP_LIST WHERE ACCOUNT_TIP_ID = " & Request.QueryString("DELETED")	
	Set RS = Conn.Execute(SQLDEL1)
	SQLDEL=  "DELETE FROM ACCOUNT_TIP WHERE ACCOUNT_TIP_ID = " & Request.QueryString("DELETED")
	Set RS = Conn.Execute(SQLDEL)	
	strError = CheckADOErrors(Conn,"DELETE" )
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing
	
	If strError = "" Then 
		Response.Redirect "TipsSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If
End If

Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.MaxRecords = MAXRECORDCOUNT
ConnectionString = CONNECT_STRING
			
cSQL = "SELECT AT.ACCOUNT_TIP_ID,AT.LOB_CD,AT.DESCRIPTION FROM ACCOUNT_TIP AT WHERE " &_
		 "AT.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	
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
	EditClick();
}
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("ATID");
}

function getAttributeValue( objRow,StrAttrib )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute(StrAttrib);
}
</SCRIPT>
<!-- #include file="..\lib\BRBtnControl.inc" -->
<script id="clientEventHandlersVBS" type="text/vbscript" language="vbscript">
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../Tips/TipsMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&ATID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function NewClick()    
    parent.frames.window.location = "../Tips/TipsMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&ATID=NEW&DETAILONLY=TRUE&AHSID=<%= Request.QueryString("AHSID") %>"
End Function


Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then		
		LOB = getAttributeValue(Document.all.tblResult.rows(i),"LOB_CD")
		ATID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to delete Account Tips for LOB: " & LOB & " and A.H.S ID: <%=Request.QueryString("AHSID")%>?", 1, "FNSDesigner")
		If lret = 1 Then
	       self.location.href = "TipsSummary.asp?DELETED=" & ATID & "&AHSID=<%= Request.QueryString("AHSID") %>"
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

</script>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0 >
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
PARAMS = "&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=TRUE&HIDESEARCH=TRUE"
%>
<OBJECT id=BRBtnControl style="LEFT: 0px; WIDTH: 100%; HEIGHT: 23px" type=text/x-scriptlet data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE<%=PARAMS%>">
	</OBJECT>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Branch_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0  rules=all ID="tblResult" name="tblResult" width=100%>
    <thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div><NOBR>TIP ID</div></td>
			<td class=thd><div><NOBR>LOB</div></td>
			<td class=thd><div><NOBR>Description</div></td>			
		</tr>
    </thead>
	<tbody ID="TableRows">
<% Do While Not oRS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" ATID='<%= oRS("ACCOUNT_TIP_ID") %>' LOB_CD='<%= oRS("LOB_CD") %>'>
			<td nowrap class=ResultCell><%= renderCell(oRS("ACCOUNT_TIP_ID")) %></td>
			<td nowrap class=ResultCell><%= renderCell(oRS("LOB_CD")) %></td>
			<td nowrap class=ResultCell><%= renderCell(oRS("DESCRIPTION")) %></td>			
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
</BODY>
</HTML>