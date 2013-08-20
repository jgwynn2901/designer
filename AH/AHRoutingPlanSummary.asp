<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\AHSTree.inc"--> 

<% 
Response.Expires=0
Response.Buffer = true
Response.AddHeader  "Pragma", "no-cache"

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

'If Request.QueryString("DELETED") <> "" Then
'	SQLDEL = ""
'	SQLDEL = SQLDEL & "{call Designer.DeleteRoutingPlan(" & Request.QueryString("DELETED") & ")}"
'	Set RS = Conn.Execute(SQLDEL)
'	Response.Redirect "AHRoutingPlanSummary.asp?AHSID=" & Request.QueryString("AHSID")
'End If

If Request.QueryString("COPY") = "TRUE" Then 
	SQLCOPY = ""
	SQLCOPY = SQLCOPY & "{call designer.CopyRoutingPlan("& Request.QueryString("RPID") & ", "& Request.QueryString("AHSID") &", {resultset 1, outRoutingPlanID})}"
	Set RS2 = Conn.Execute(SQLCOPY)
	Response.Redirect "AHRoutingPlanSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

If Request.QueryString("MULTICOPY") = "TRUE" Then 

	ids = Request.QueryString("MUTLIRPIDS")

	Dim TokenArray
	TokenArray = Split(ids,",",-1,1)

	lastIndex = UBound(TokenArray)
	
	Dim RszRtn
	for j=0 to LastIndex
		SQLCOPY = ""
		SQLCOPY = SQLCOPY & "{call designer.CopyRoutingPlan("& TokenArray(j) & ", "& Request.QueryString("AHSID") &", {resultset 1, outRoutingPlanID})}"

		Set RszRtn = Conn.Execute(SQLCOPY)
		RszRtn.Close
		Set RszRtn = Nothing
	Next
	Response.Redirect "AHRoutingPlanSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

If Request.QueryString("CLEARFILTER") <> "" Then
	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_RPFILTER"
	Response.Redirect "AHRoutingPlanSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	SQL = ""
	SQL = SQL & "SELECT ROUTING_PLAN_ID, LOB_CD, DESCRIPTION FROM ROUTING_PLAN WHERE ENABLED_FLG <> 'N' And ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	
	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_RPFILTER", "MUSTINCLUDE")	
	
	If Request.form("txtRPID") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.form("txtRPID")
		Else
			strInclude  = Request.form("txtRPID")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_RPFILTER", "MUSTINCLUDE", strInclude
	End If
	
	if strInclude <> "" then SQL = SQL & " AND ROUTING_PLAN_ID IN (" & strInclude & ") "

	SQL = SQL & " ORDER BY LOB_CD, DESCRIPTION "
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
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
	return objRow.getAttribute("ROUTINGPLANID");
}

function getname( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("RPNAME");
}

function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHRoutingPlanSummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}

-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../RoutingPlan/RoutingPlanSummary-f.asp?ROUTING_PLAN_ID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function CopyClick()
	ClipboardAgent.ClearAllProperties()
	
	SelectesItems = getmultipleindex(Document.all.tblResult, "ROUTING_PLAN_ID")
	
	Dim TokenArray
	TokenArray = Split(SelectesItems,"||",-1,1)
	
	count = UBound(TokenArray) + 1
	
	If count = 1 Then
		i = getselectedindex( Document.all.tblResult )
		if 0 < i then
			ClipboardAgent.AddProperty "ROUTING_PLAN_ID", "RPID=" & dblhighlight(Document.all.tblResult.rows(i))
			ClipboardAgent.AddProperty "ROUTING_PLAN_ID_TEXT", getname(Document.all.tblResult.rows(i))
			ClipboardAgent.SetPropertiesToClipboard
		end if
	ElseIf count > 1 Then
			ClipboardAgent.AddProperty "MULTI_ROUTING_PLAN_IDS", join(TokenArray, ",")
			ClipboardAgent.AddProperty "MULTI_ROUTING_PLAN_IDS_TEXT", CStr(count) & " Routing plans (" & join(TokenArray, ",") & ")"
			ClipboardAgent.SetPropertiesToClipboard
	End If
	
End Function

Function PasteClick()
	If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("ROUTING_PLAN_ID") <> "" Then
		lret = msgbox ("Are you sure you want to paste Routing Plan """ & ClipboardAgent.GetProperty("ROUTING_PLAN_ID_TEXT") & """ (" & ClipboardAgent.GetProperty("ROUTING_PLAN_ID") & ") ?" , 1, "FNSDesigner") 
		If lret = "1" Then 
			self.location.href = "AHRoutingPlanSummary.asp?COPY=TRUE&" & ClipboardAgent.GetProperty("ROUTING_PLAN_ID") & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	ElseIf ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("MULTI_ROUTING_PLAN_IDS") <> "" Then
		lret = msgbox ("Are you sure you want to paste " & ClipboardAgent.GetProperty("MULTI_ROUTING_PLAN_IDS_TEXT") & "?" , 1, "FNSDesigner") 
		If lret = "1" Then 
			self.location.href = "AHRoutingPlanSummary.asp?MULTICOPY=TRUE&MUTLIRPIDS=" & ClipboardAgent.GetProperty("MULTI_ROUTING_PLAN_IDS") & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If		
	Else
		MsgBox "No data to paste!", 0, "FNSDesigner"
	End If
End Function

Function NewClick()
	parent.frames.window.location = "../RoutingPlan/RoutingPlanSummary-f.asp?ROUTING_PLAN_ID=NEW&AHSID=<%= Request.QueryString("AHSID") %>"
End Function

'Function RemoveClick()
'	i = getselectedindex( Document.all.tblResult )
'	if 0 < i then
'		RPID = dblhighlight(Document.all.tblResult.rows(i))
'		lret = MsgBox ("Are you sure you want to remove Routing Plan ID:" & RPID, 1, "FNSDesigner")
'		If lret = 1 Then
'			self.location.href = "AHRoutingPlanSummary.asp?DELETED=" & RPID & "&AHSID=<%= Request.QueryString("AHSID") %>"
'		End If
'	end if
'End Function

Function SearchClick()
	lret = ""
	strURL = ""
	SearchObj.multiselected = ""
	strURL = "../RoutingPlan/RoutingPlanSearchModal.asp?CONTAINERTYPE=MODAL&AHSID=<%= Request.QueryString("AHSID") %>"
	lret = window.showModalDialog(strURL  ,SearchObj ,"dialogWidth:580px;dialogHeight:550px;center")
	if SearchObj.multiselected <> "" Then
		document.all.rpid.action = "AHRoutingPlanSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
		document.all.txtRPID.value = SearchObj.multiselected
		document.all.rpid.submit 
		'self.location.href = "AHRoutingPlanSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & SearchObj.multiselected
	End If
End Function

'Function AttachClick()
'
'End Function

Function RefreshClick()
	self.location.href = "AHRoutingPlanSummary.asp?<%= Request.Querystring %>"
End Function

Sub window_onload
<% If RS.RecordCount = MAXRECORDCOUNT Then %>
	StatusSpan.innerHTML = "<%= MSG_MAXRECORDS %>"
<% Else %>
	StatusSpan.innerHTML = "Record Count is <%= RS.RecordCount %>"
<% End If %>	

<% If strInclude <> "" Then %>
	FilterSpan.innerHTML = "<IMG SRC='..\images\filter2.gif'></IMG>"
<%	Else %>	
	FilterSpan.innerHTML = ""
<%	End If%>

	ClipboardAgent.GetPropertiesFromClipboard
End Sub
-->
</SCRIPT>
<SCRIPT LANGUAGE=JavaScript>
<!--
function CRPSearchObj()
{
	this.routing_plan_id = "";
	this.ahsid = "";
	this.multiselected = "";
}
var SearchObj = new CRPSearchObj();
//-->
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" FOR="RPBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
	case "EDITBUTTONCLICK":
       	EditClick()
		break;

	case "NEWBUTTONCLICK":
		NewClick()
		break;

 	case "COPYBUTTONCLICK":
		CopyClick()
		break;

	case "PASTEBUTTONCLICK":
		PasteClick()
		break;

	case "SEARCHBUTTONCLICK":
		SearchClick()
		break;

//	case "REMOVEBUTTONCLICK":
//		RemoveClick()
//		break;
	
	case "REFRESHBUTTONCLICK":
		RefreshClick()
		break;
		
//	case "ATTACHBUTTONCLICK":
//		AttachClick()
//		break;
	
	default:
		break;
}
 
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
	
	PARAMS = ""
	If HasModifyPrivilege("FNSD_ROUTING_PLAN","") <> True Then	PARAMS = PARAMS & "&HIDEEDIT=TRUE"
	If HasAddPrivilege("FNSD_ROUTING_PLAN","") <> True Then	PARAMS = PARAMS & "&HIDENEW=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE"
''	If HasDeletePrivilege("FNSD_ROUTING_PLAN","") <> True Then	PARAMS = PARAMS & "&HIDEREMOVE=TRUE"
%>
<OBJECT VIEWASTEXT data="../Scriptlets/ObjButtons.asp?SEARCHCAPTION=Filter&HIDEATTACH=TRUE&HIDEREMOVE=TRUE<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="RPBtnControl" type=text/x-scriptlet></OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="ROUTING_PLAN_ID"><NOBR>R.P. ID</div></td>
			<td class=thd><div id="LOB_CD"><NOBR>LOB</div></td>
			<td class=thd><div id="DESCRIPTION"><NOBR>Desc</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" RPNAME= "<%= Mid(RS("DESCRIPTION"),1,25) %>" ROUTINGPLANID='<%= RS("ROUTING_PLAN_ID") %>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("ROUTING_PLAN_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("LOB_CD")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("DESCRIPTION")) %></td>
	</tr>
<% 
RS.MoveNext
Loop
RS.Close
%>
</tbody>
</table>
</DIV>
</FIELDSET>
</SCRIPT>
<form id="rpid" action="AHRoutingPlanSummary.asp" method= "post">
<input type=hidden name="txtRPID">
</form>

<OBJECT VIEWASTEXT ID="ClipboardAgent" 
<%GetClipboardCLSID()%>
width=1 height=1>
<PARAM NAME="MaxPropertiesStringLength" VALUE="1000">
<PARAM NAME="MaxPropertyNameLength" VALUE="50">
<PARAM NAME="MaxPropertyValueLength" VALUE="200">
<PARAM NAME="NameValueDelimiter" VALUE="#">
<PARAM NAME="PropertyItemDelimiter" VALUE="|">
<PARAM NAME="PrivateClipboardFormatName" VALUE="CF_FNSDESIGNER">
</OBJECT>
</BODY>
</HTML>
