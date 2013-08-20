<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\AHSTree.inc"--> 

<%Response.Expires=0

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	

If Request.QueryString("DELETED") <> "" Then
	SQLDEL = ""
'	SQLDEL = "{call Designer.DeletePolicy(" & Request.QueryString("DELETED") & ")}"
'	Set RS = Conn.Execute(SQLDEL)
	Response.Redirect "AHPolicySummary.asp?AHSID=" & Request.QueryString("AHSID")
End If	

If Request.QueryString("COPY") = "TRUE" Then 
	SQLCOPY = ""
	SQLCOPY = "{call designer.CopyPolicy("& Request.QueryString("PID") & ", "& Request.QueryString("AHSID") &", {resultset 1, outPolicyID})}"
	Set RS2 = Conn.Execute(SQLCOPY)
	RS2.Close
	Conn.Close
	Response.Redirect "AHPolicySummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

If Request.QueryString("CLEARFILTER") <> "" Then
	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_POLICYFILTER"
	Response.Redirect "AHPolicySummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	SQL = ""

	'**************************************************
	' DMS: 2/17/00 Changed the SQL to grab the columns 
	'              LOB_CD and ACCNT_HRCY_STEP_ID from 
	'              AHS_POLICY as the columns have been removed from 
	'              the POLICY table.
	'**************************************************

	SQL = "SELECT P.*, AHSP.*, " &_
		  " (SELECT  VALUE FROM POLICY_EXTENSION WHERE POLICY_EXTENSION.POLICY_ID = P.Policy_ID " &_
		  " AND NAME = 'CLAIM:POLICY:CONTRACT_NUMBER') AS CONTRACT_NO " &_
	      "  FROM POLICY P, AHS_POLICY AHSP" &_
		  " WHERE P.POLICY_ID             = AHSP.POLICY_ID " &_
		  "   AND AHSP.ACCNT_HRCY_STEP_ID = " & Request.QueryString("AHSID")

	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_POLICYFILTER", "MUSTINCLUDE")	
	
	If Request.QueryString("MultiSelected") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.QueryString("multiselected")
		Else
			strInclude  = Request.QueryString("multiselected")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_POLICYFILTER", "MUSTINCLUDE", strInclude
	End If
	
	if strInclude <> "" then SQL = SQL & " AND POLICY_ID IN (" & strInclude & ") "
	
	SQL = SQL & "  ORDER BY AHSP.LOB_CD, P.EFFECTIVE_DATE DESC"
	
	RS.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblFields )
	if 0 < i then
		parent.frames.window.location = "../Policy/PolicyMaintenance.asp?CONTAINERTYPE=FRAMEWORK&PID=" & dblhighlight(Document.all.tblFields.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>&MODE=RW&DETAILONLY=TRUE"
	end if
End Function

Function CopyClick()
	ClipboardAgent.ClearAllProperties()
	i = getselectedindex( Document.all.tblFields )
	if 0 < i then
		ClipboardAgent.AddProperty "POLICY_ID", "PID=" & dblhighlight(Document.all.tblFields.rows(i))
		ClipboardAgent.SetPropertiesToClipboard
	end if 
End Function

Function PasteClick()
	If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("POLICY_ID") <> "" Then
	lret = msgbox ("Are you sure you want to paste Policy ID " & ClipboardAgent.GetProperty("POLICY_ID") & "?" , 1, "FNSDesigner") 
		If lret = "1" Then '
			self.location.href = "AHPolicySummary.asp?COPY=TRUE&" & ClipboardAgent.GetProperty("POLICY_ID") & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	Else
		MsgBox "No data to paste!", 0, "FNSDesigner"
	End If
End Function

Function NewClick()
	parent.frames.window.location = "../Policy/PolicyMaintenance.asp?CONTAINERTYPE=FRAMEWORK&PID=NEW&AHSID=<%= Request.QueryString("AHSID") %>&MODE=RW&DETAILONLY=TRUE"
End Function

Function RemoveClick()
	MsgBox "Delete is not implemented for Policy.",0,"FNSNetDesigner"
	Exit Function

	i = getselectedindex( Document.all.tblFields )
	if 0 < i then
		PID = dblhighlight(Document.all.tblFields.rows(i))
		lret = MsgBox ("Are you sure you want to remove Policy ID:" & PID, 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "AHPolicySummary.asp?DELETED=" & PID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	end if
End Function

Function SearchClick()
	lret = ""
	strURL = ""
	SearchObj.PID = ""
	strURL = "../Policy/PolicyMaintenance.asp?CONTAINERTYPE=MODAL&SEARCHONLY=TRUE&SearchAHSID=<%= Request.QueryString("AHSID") %>"
	lret = window.showModalDialog(strURL  ,SearchObj ,"center")
	if SearchObj.PID <> "" Then
		multi = Replace(SearchObj.PID,"||",",")
		self.location.href = "AHPolicySummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & multi
	End If
End Function

Function AttachClick()

End Function

Function RefreshClick()
	self.location.href = "AHPolicySummary.asp?<%= Request.Querystring %>"
End Function

Sub window_onload
<% If RS.RecordCount = MAXRECORDCOUNT Then %>
	StatusSpan.innerHTML = "<%= MSG_MAXRECORDS %>"
<% Else %>
	StatusSpan.innerHTML = "Record Count is <%= RS.RecordCount %>"
<% End If %>	
	ClipboardAgent.GetPropertiesFromClipboard
<% If strInclude <> "" Then %>
	FilterSpan.innerHTML = "<IMG SRC='..\images\filter2.gif'></IMG>"
<%	Else %>	
	FilterSpan.innerHTML = ""
<%	End If%>
End Sub
-->
</SCRIPT>
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
	return objRow.getAttribute("PID")
}
function BSearchObj()
{
	this.PID = "";
	this.Selected = "";
}
function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHPolicySummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}

var SearchObj = new BSearchObj();
-->
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" FOR="PolicyBtnControl" EVENT="onscriptletevent (event, obj)">
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
	If HasModifyPrivilege("FNSD_POLICY","") <> True Then	PARAMS = PARAMS & "&HIDEEDIT=TRUE"
	If HasAddPrivilege("FNSD_POLICY","") <> True Then	PARAMS = PARAMS & "&HIDENEW=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE"
	If HasDeletePrivilege("FNSD_POLICY","") <> True Then	PARAMS = PARAMS & "&HIDEREMOVE=TRUE"
%>
<OBJECT data="../Scriptlets/ObjButtons.asp?SEARCHCAPTION=Filter&HIDEATTACH=TRUE<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="PolicyBtnControl" Name="PolicyBtnControl" type=text/x-scriptlet></OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblFields" name="tblFields" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>Policy ID</div></td>
			<td class=thd><div id="NAME_HEAD"><NOBR>Policy Number</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>LOB</div></td>
			<td class=thd><div id="CONTRACTNUMBER_HEAD"><NOBR>Contract Number</div></td>
			<td class=thd><div id="EFFECTIVEDATE_HEAD"><NOBR>Effective Date</div></td>
			<td class=thd><div id="EXPIRATIONDATE_HEAD"><NOBR>Expiration Date</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>Description</div></td>
		</tr>
	</thead>
	
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>	
		<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" PID='<%= RS("POLICY_ID") %>'>
			<td NOWRAP CLASS=ResultCell ><%= renderCell(RS("POLICY_ID")) %></td>
			<td NOWRAP CLASS=ResultCell ><%= renderCell(RS("POLICY_NUMBER")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("LOB_CD")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("CONTRACT_NO")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("EFFECTIVE_DATE")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("EXPIRATION_DATE")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("POLICY_DESC")) %></td>
		</tr>
<% 
RS.MoveNext
Loop
RS.Close
Conn.Close

%>
</tbody>
</table>
</DIV>
</FIELDSET>
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
