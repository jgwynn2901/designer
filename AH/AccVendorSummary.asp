<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->  
<!--#include file="..\lib\AHSTree.inc"--> 
<% 
Response.Expires=0 
Response.Buffer = True

dim cSQL, oRS, oConn, cAHSID
dim nID2Delete, cLOB, nST

cAHSID = trim(Request.QueryString("AHSID"))
nID2Delete = CInt(Request.QueryString("DELETED"))
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING

if nID2Delete <> 0 then
	cSQL = "SELECT * FROM ACCOUNT_VENDOR " & _
			"WHERE ACCOUNT_VENDOR_ID = " & nID2Delete
	Set oRS = oConn.Execute(cSQL)
	cLOB = oRS("LOB")
	nST = oRS("SERVICE_TYPE_ID")
	oRS.close
	cSQL = "DELETE FROM ACCOUNT_VENDOR " & _
			"WHERE ACCNT_HRCY_STEP_ID = " & cAHSID & _
			" AND SERVICE_TYPE_ID = " & nST & _
			" AND LOB = '" & cLOB & "'"
	oConn.Execute cSQL
end if	
cSQL = "SELECT distinct AV.*, ST.TYPE, CM.NAME FROM ACCOUNT_VENDOR AV, SERVICE_TYPE ST, CONTACT_METHOD CM " & _
		"WHERE AV.ACCNT_HRCY_STEP_ID = " & cAHSID & _
		" AND AV.SERVICE_TYPE_ID = ST.SERVICE_TYPE_ID " & _
		"AND AV.CONTACT_METHOD_ID = CM.CONTACT_METHOD_ID"
cSQL = "SELECT distinct AV.*, ST.TYPE FROM ACCOUNT_VENDOR AV, SERVICE_TYPE ST " & _
		"WHERE AV.ACCNT_HRCY_STEP_ID = " & cAHSID & _
		" AND AV.SERVICE_TYPE_ID = ST.SERVICE_TYPE_ID ORDER BY " & _
		"AV.SERVICE_TYPE_ID, LOB"

'cSQL = "SELECT distinct SERVICE_TYPE_ID, LOB FROM ACCOUNT_VENDOR " & _
		'"WHERE ACCNT_HRCY_STEP_ID = " & cAHSID 
Set oRS = oConn.Execute(cSQL)
if not oRS.eof then
	RS_LOB = oRS("LOB")
	RS_ST = oRS("SERVICE_TYPE_ID")
end if
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="Javascript">
<!--
function dblclick( objRow )
{
	EditClick()
}
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("AVID");
}

function CBranchAssignTypeSearchObj()
{
	this.BATID = "";
	this.Selected = "";
}
var BranchAssignTypeSearchObj = new CBranchAssignTypeSearchObj();
-->
</SCRIPT>
<!-- #include file="..\lib\BRBtnControl.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "AccVendorMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&AVID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function NewClick()
parent.frames.window.location = "AccVendorMaintenance.asp?CONTEXTTYPE=DRILLIN&AHSID=<%= Request.QueryString("AHSID") %>&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&AVID=NEW"
End Function

Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		AVID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to delete Account/Vendor # " & AVID & "?", 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "AccVendorSummary.asp?DELETED=" & AVID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	end if
End Function

Function RefreshClick()
	self.location.href = "AHBranchSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Sub window_onload
<% If strInclude <> "" Then %>
	FilterSpan.innerHTML = "<IMG SRC='..\images\filter2.gif'></IMG>"
<%	Else %>	
	FilterSpan.innerHTML = ""
<%	End If%>


<%If strError <> "" Then %>
	MsgBox ("<%=strError%>")
<% End If %>


End Sub
-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
	
PARAMS = "&HIDEEDIT=false"
PARAMS = PARAMS & "&HIDENEW=false&HIDECOPY=TRUE&HIDEPASTE=TRUE"
PARAMS = PARAMS & "&HIDEREMOVE=false&HIDEREFRESH=TRUE&HIDESEARCH=TRUE"
%>
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="BRBtnControl" type=text/x-scriptlet></OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Branch_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0  rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div><NOBR>LOB</div></td>
			<td class=thd><div><NOBR>Service Type</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% 
cLOB = ""
nST = 0
Do While Not oRS.EOF 
	if Cint(oRS("SERVICE_TYPE_ID")) <> nST or oRS("LOB") <> cLOB then
%>	
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" AVID='<%=oRS("ACCOUNT_VENDOR_ID")%>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(oRS("LOB")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(oRS("TYPE")) %></td>
	</tr>
<% 
		nST = Cint(oRS("SERVICE_TYPE_ID"))
		cLOB = oRS("LOB")
	end if
	oRS.MoveNext
Loop
oRS.Close
oConn.close
set oConn = nothing
set oRS = nothing
%>
</tbody>
</table>
</DIV>
</FIELDSET>
</SCRIPT>
</BODY>
</HTML>
