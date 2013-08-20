<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\AHSTree.inc"--> 
<% Response.Expires=0 
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

If Request.QueryString("DELETED") <> "" Then
	SQLDEL = ""
	SQLDEL = SQLDEL & "DELETE FROM FEE WHERE FEE_ID=" & Request.QueryString("DELETED")
	Set RS = Conn.Execute(SQLDEL)
	Response.Redirect "AHBillingSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

'If Request.QueryString("COPY") = "TRUE" Then 
'	SQLCOPY = ""
'	SQLCOPY = SQLCOPY & "{call designer.CopyRoutingPlan("& Request.QueryString("RPID") & ", "& Request.QueryString("AHSID") &", {resultset 1, outRoutingPlanID})}"
'	Set RS2 = Conn.Execute(SQLCOPY)
'	Response.Redirect "AHRoutingPlanSummary.asp?AHSID=" & Request.QueryString("AHSID")
'End If

If Request.QueryString("CLEARFILTER") <> "" Then
	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_BILLINGFILTER"
	Response.Redirect "AHBillingSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	SQL = ""
	SQL = SQL & "SELECT FEE.FEE_ID, FEE.LOB_CD, FEE.FEE_AMOUNT, FEE.DESCRIPTION, FEE_TYPE.NAME FROM FEE, FEE_TYPE WHERE ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID") & " AND "
	SQL = SQL & "FEE.FEE_TYPE_ID = FEE_TYPE.FEE_TYPE_ID "
	
	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_BILLINGFILTER", "MUSTINCLUDE")	
	
	If Request.QueryString("MultiSelected") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.QueryString("multiselected")
		Else
			strInclude  = Request.QueryString("multiselected")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_BILLINGFILTER", "MUSTINCLUDE", strInclude
	End If
	
	if strInclude <> "" then SQL = SQL & " AND FEE_ID IN (" & strInclude & ") "
	
	SQL = SQL & " ORDER BY FEE_ID"
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="Javascript">
<!--
function dblclick ( objRow )
{
	EditClick()
}
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("FEEID");
}
function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHBillingSummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}
-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../Billing/BillingMaintenance.asp?CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&BID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function CopyClick()
msgbox "Not Available"
	'ClipboardAgent.ClearAllProperties()
	'i = getselectedindex( Document.all.tblResult )
	'if 0 < i then
		'ClipboardAgent.AddProperty "FEE_ID", "FEE_ID=" & dblhighlight(Document.all.tblResult.rows(i))
		'ClipboardAgent.SetPropertiesToClipboard
	'end if
End Function

Function PasteClick()
msgbox "Not Available"
	'If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("ROUTING_PLAN_ID") <> "" Then
	'lret = msgbox ("Are you sure you want to paste Routing Plan ID " & ClipboardAgent.GetProperty("ROUTING_PLAN_ID") & "?" , 1, "FNSDesigner") 
		'If lret = "1" Then 
			'self.location.href = "AHRoutingPlanSummary.asp?COPY=TRUE&" & ClipboardAgent.GetProperty("ROUTING_PLAN_ID") & "&AHSID=<%= Request.QueryString("AHSID") %>"
		'End If
	'Else
		'MsgBox "No data to paste!", 0, "FNSDesigner"
	'End If
End Function

Function NewClick()
	parent.frames.window.location = "../Billing/BillingMaintenance.asp?AHSID=<%= Request.QueryString("AHSID") %>&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&BID=NEW"
End Function

Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		BID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to remove Billing ID:" & BID, 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "AHBillingSummary.asp?DELETED=" & BID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	end if
End Function

Function SearchClick()
	lret = ""
	strURL = ""
	SearchObj.multiselected = ""
	strURL = "../Billing/BillingMaintenance.asp?CONTAINERTYPE=MODAL&SEARCHONLY=TRUE&AHSID=<%= Request.QueryString("AHSID") %>&SearchAHSID=<%= Request.QueryString("AHSID") %>"
	lret = window.showModalDialog(strURL  ,SearchObj ,"center")
	
	if SearchObj.BillingID <> "" Then
		multi = Replace(SearchObj.BillingID,"||",",")
		self.location.href = "AHBillingSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & multi
	End If
End Function

Function AttachClick()
	msgbox "Not ImplementedYet"
End Function

Function RefreshClick()
	self.location.href = "AHBillingSummary.asp?AHSID=<%= Request.Querystring("AHSID") %>&REFRESH=TRUE"
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
function BSearchObj()
{
	this.BillingID = "";
	this.ahsid = "";
	this.multiselected = "";
}
var SearchObj = new BSearchObj();
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

	case "REMOVEBUTTONCLICK":
		RemoveClick()
		break;
	
	case "REFRESHBUTTONCLICK":
		RefreshClick()
		break;
		
	case "ATTACHBUTTONCLICK":
		AttachClick()
		break;
	
	default:
		break;
}
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
	
	PARAMS = ""
	If HasModifyPrivilege("FNSD_FEE","") <> True Then	PARAMS = PARAMS & "&HIDEEDIT=TRUE"
	If HasAddPrivilege("FNSD_FEE","") <> True Then	PARAMS = PARAMS & "&HIDENEW=TRUE"
	If HasDeletePrivilege("FNSD_FEE","") <> True Then	PARAMS = PARAMS & "&HIDEREMOVE=TRUE"
%>
<OBJECT data="../Scriptlets/ObjButtons.asp?SEARCHCAPTION=Filter&HIDEATTACH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="RPBtnControl" type=text/x-scriptlet></OBJECT>

<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>Fee ID</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>LOB</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Fee Type</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>Amount</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>Description</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" FEEID='<%= RS("FEE_ID") %>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("FEE_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("LOB_CD")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("NAME")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(FormatCurrency(RS("FEE_AMOUNT"))) %></td>
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
<OBJECT ID="ClipboardAgent" 
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
