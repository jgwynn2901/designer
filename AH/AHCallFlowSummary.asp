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
		
Function NextPkey( TableName, ColName )
	NextSQL = ""
	NextSQL = "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function

If Request.QueryString("ATTACH") = "TRUE" Then
	NextID = NextPkey( "ACCOUNT_CALLFLOW", "ACCOUNTCALLFLOW_ID")
	ON ERROR RESUME NEXT
	SQL = ""
	SQL = SQL & "INSERT INTO ACCOUNT_CALLFLOW (ACCOUNTCALLFLOW_ID, CALLFLOW_ID, "	
	SQL = SQL & "ACCNT_HRCY_STEP_ID, LOB_CD, CALL_START_FLG ) VALUES ( "
	SQL = SQL & NextID & ", "
	SQL = SQL & Request.QueryString("ATTACHCALLFLOW_ID")  & ", "
	SQL = SQL & Request.QueryString("AHSID")  & ", "
	SQL = SQL & "'???', "
	SQL = SQL & " 'Y' ) "
	Set Update = Conn.Execute(SQL)
	IF Conn.Errors.Count > 0 Then
	    s_AttachErrorMsg = Server.URLEncode("<FONT COLOR='RED'>Attach Error:</FONT> " & Mid(Conn.Errors(0).Description, Instr(1, Conn.Errors(0).Description, ":") +2))
		response.redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID") + "&ErrMsg=" & s_AttachErrorMsg
	ELSE
		response.redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If
End if
	
If Request.QueryString("DELETED") <> "" AND ISNumeric(Request.QueryString("DELETED")) Then
	SQLDEL = "" 
	SQLDEL = "DELETE FROM ACCOUNT_CALLFLOW WHERE ACCOUNTCALLFLOW_ID=" & Request.QueryString("DELETED")
	Set RSDel = Conn.Execute(SQLDEL)
	Response.Redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

IF Request.QueryString("COPY_OPERATION") = "SingleCopy" THEN
	If Request.QueryString("COPYCFID") <> "" AND IsNumeric(Request.QueryString("COPYCFID")) Then
		SQLCOPY = "" 
		SQLCOPY = "{call Designer.AddAcctCallFlowAndCopyCallFlow(" & Request.QueryString("COPYCFID") & ", " & Request.QueryString("COPYACFID") & ", " & Request.QueryString("AHSID") & ", '???', {resultset 1, outCallFlowId, outAcctCallFlowId})}"
		Set RSCopy = Conn.Execute(SQLCOPY)
		if RSCopy.Fields("outCallFlowId") = "0" OR RSCopy.Fields("outAcctCallFlowId") = "0" then
			s_Error = server.URLEncode("<FONT COLOR='RED'>Error Pasting:</FONT> CFID(" & Request.QueryString("COPYCFID") & ")")
			Response.Redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID") + "&ErrMsg=" & s_Error
		else
			Response.Redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID")
		end if
		RSCopy.Close
		Set RSCopy = Nothing
	End If
ELSEIF Request.QueryString("COPY_OPERATION") = "MultiCopy" THEN
	Dim a_CF_Array, a_ACF_Array, s_CopyErrorMsg
	a_CF_Array = Split(Request.QueryString("MULTICOPYCFIDS"), ",", -1, 1)
	a_ACF_Array = Split(Request.QueryString("MULTICOPYACFIDS"), ",", -1, 1)
	i_Loop = UBound(a_CF_Array)
		For i = 0 to i_Loop
			IF i < 10 THEN
				s_Counter = "??" + CStr(i)
			ELSEIF i > 9 AND i < 100 THEN
				s_Counter = "?" + CStr(i)
			ELSE
				s_Counter = CStr(i)
			END IF
			SQLCOPY = ""
			SQLCOPY = "{call Designer.AddAcctCallFlowAndCopyCallFlow(" & a_CF_Array(i) & ", " & a_ACF_Array(i) & ", " & Request.QueryString("AHSID") & ", '" & s_Counter & "', {resultset 1, outCallFlowId, outAcctCallFlowId})}"
			Set RSCopy = Conn.Execute(SQLCOPY)
			IF RSCopy.Fields("outCallFlowId") = "0" OR RSCopy.Fields("outAcctCallFlowId") = "0" THEN
				If s_CopyErrorMsg = "" Then
					s_CopyErrorMsg = "<FONT COLOR='RED'>Error Pasting:</FONT> CFIDs (" & a_CF_Array(i) & ")"
				Else
					s_CopyErrorMsg = s_CopyErrorMsg & ", (" & a_CF_Array(i) & ")"
				End IF
				RSCopy.Close
				Set RSCopy = Nothing
			END IF	
		Next
	IF s_CopyErrorMsg = "" THEN
		Response.Redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID")
	ELSE
		s_Error = server.URLEncode(s_CopyErrorMsg & ".  LOB Conflict." )
		Response.Redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID") + "&ErrMsg=" & s_Error
	END IF
END IF
If Request.QueryString("CLEARFILTER") <> "" Then
	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_CFFILTER"
	Response.Redirect "AHCallFlowSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	SQL = ""
	SQL = SQL & "SELECT ACCOUNT_CALLFLOW.*, CALLFLOW.* FROM ACCOUNT_CALLFLOW, CALLFLOW WHERE ACCOUNT_CALLFLOW.ACCNT_HRCY_STEP_ID= " & Request.QueryString("AHSID")
	
	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_CFFILTER", "MUSTINCLUDE")	

	If Request.QueryString("MultiSelected") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.QueryString("multiselected")
		Else
			strInclude  = Request.QueryString("multiselected")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_CFFILTER", "MUSTINCLUDE", strInclude
	End If

	if strInclude <> "" then SQL = SQL & " AND CALLFLOW.CALLFLOW_ID IN (" & strInclude & ") "

	SQL = SQL & " AND ACCOUNT_CALLFLOW.CALLFLOW_ID = CALLFLOW.CALLFLOW_ID ORDER BY CALLFLOW.CALLFLOW_ID"
	
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
	return objRow.getAttribute("CALLFLOWID");
}
function GetAccountCallflow_id( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("ACCOUNTCALLFLOWID");
}

function GetCallflowName( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("CFNAME");
}

function CRPSearchObj()
{
	this.routing_plan_id = "";
	this.ahsid = "";
	this.multiselected = "";
}
var SearchObj = new CRPSearchObj();

function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHCallFlowSummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}
-->
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" FOR="CFBtnControl" EVENT="onscriptletevent (event, obj)">
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
		RefreshClick();
		break;
	case "ATTACHBUTTONCLICK":
		AttachClick();
		break;
	default:
		break;
}
   
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		'Dim curURL
		'curURL  = "../AH/AHCallFlowSummary.asp?CONTAINERCONTEXT=DRILLIN&CONTAINERTYPE=FRAMEWORK&AHSID=<%= Request.QueryString%>"
		'SetInfoForNavigateBack(curURL)
		
		parent.frames.window.location = "../CallFlow/CallFlowEditor.asp?ACCOUNTCALLFLOW_ID=" & GetAccountCallflow_id(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>&CALLFLOW_ID=" & dblhighlight(Document.all.tblResult.rows(i))
		'parent.frames.window.location = "../CallFlow/CallFlowContainer.asp?CALLFLOW_ID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function CopyClick()
	ClipboardAgent.ClearAllProperties()
	CF_SelectedItems = getmultipleindex(Document.all.tblResult, "CALLFLOW_ID")
	IF Instr(1, CF_SelectedItems, "||") THEN
		Dim j
		For j = 1 to Document.all.tblResult.rows.Length -1
			If ( Document.all.tblResult.rows(j).className = "ResultSelectRow" ) Then
				if CF_IDs = "" then
					CF_IDs = CF_IDs + dblhighlight(Document.all.tblResult.rows(j))
					ACF_IDs = ACF_IDs + GetAccountCallflow_id(Document.all.tblResult.rows(j))
					CF_NAMEs = CF_NAMEs + GetCallflowName(Document.all.tblResult.rows(j))
				else
					CF_IDs = CF_IDs + "," + dblhighlight(Document.all.tblResult.rows(j))
					ACF_IDs = ACF_IDs + "," + GetAccountCallflow_id(Document.all.tblResult.rows(j))
					CF_NAMEs = CF_NAMEs + "," + GetCallflowName(Document.all.tblResult.rows(j))
				end if
			End IF
		Next
		Array_CF_IDs = Split(CF_IDs, ",", -1, 1)
		Array_ACF_IDs = Split(ACF_IDs, ",", -1, 1)
		Arrray_CF_NAMEs = Split(CF_NAMEs, ",", -1, 1)
		ClipboardAgent.AddProperty "MULTI_CALLFLOW_IDS", Join(Array_CF_IDs, ",")
		ClipboardAgent.AddProperty "MULTI_ACCOUNTCALLFLOW_IDS", Join(Array_ACF_IDs, ",")
		ClipboardAgent.AddProperty "MULTI_CALLFLOW_NAMES", Join(Arrray_CF_NAMEs, ",")
	ELSE
		i = getselectedindex( Document.all.tblResult )
		if i <> -1 then
			ClipboardAgent.AddProperty "CALLFLOW_ID", dblhighlight(Document.all.tblResult.rows(i))
			ClipboardAgent.AddProperty "ACCOUNTCALLFLOW_ID", GetAccountCallflow_id(Document.all.tblResult.rows(i))
			ClipboardAgent.AddProperty "CALLFLOW_NAME", GetCallflowName(Document.all.tblResult.rows(i))
		end if
	END IF
	ClipboardAgent.SetPropertiesToClipboard 
	'MsgBox "CALLFLOW_ID=" & ClipboardAgent.GetProperty("CALLFLOW_ID") & VBCRLF & "ACCOUNTCALLFLOW_ID" & ClipboardAgent.GetProperty("ACCOUNTCALLFLOW_ID")
End Function

Function PasteClick()
	IF ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("CALLFLOW_ID") <> "" THEN
		lret = MsgBox ("Are you sure you want to copy CallFlow """ & ClipboardAgent.GetProperty("CALLFLOW_NAME")& """ (" & ClipboardAgent.GetProperty("CALLFLOW_ID")& ") ?", 1, "FNSDesigner" ) 
		If lret = "1" Then
			Url = ""
			Url = Url & "AHCallFlowSummary.asp?COPY_OPERATION=SingleCopy&COPYCFID=" & ClipboardAgent.GetProperty("CALLFLOW_ID") 
			Url = Url & "&COPYACFID=" & ClipboardAgent.GetProperty("ACCOUNTCALLFLOW_ID") 
			Url = Url & "&AHSID=<%= Request.QueryString("AHSID") %>" 
			self.location.href = Url
		End If
	ELSEIF ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("MULTI_CALLFLOW_IDS") <> "" THEN
		lret = MsgBox ("Are you sure you want to copy CallFlows """ & ClipboardAgent.GetProperty("MULTI_CALLFLOW_NAMES")& """ (" & ClipboardAgent.GetProperty("MULTI_CALLFLOW_IDS")& ") ?", 1, "FNSDesigner" ) 
		If lret = "1" Then
			Url = ""
			Url = Url & "AHCallFlowSummary.asp?COPY_OPERATION=MultiCopy&MULTICOPYCFIDS=" & ClipboardAgent.GetProperty("MULTI_CALLFLOW_IDS") 
			Url = Url & "&MULTICOPYACFIDS=" & ClipboardAgent.GetProperty("MULTI_ACCOUNTCALLFLOW_IDS") 
			Url = Url & "&AHSID=<%= Request.QueryString("AHSID") %>" 
			self.location.href = Url
		End If
	ELSE
		MsgBox "No data to paste!", 0, "FNSDesigner"
	END IF
End Function

Function NewClick()
	parent.frames.window.location = "../CallFlow/CallFlowEditor.asp?CALLFLOW_ID=NEW&AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		ACFID = GetAccountCallflow_id(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to remove Account Call Flow ID:" & ACFID & "?", 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "AHCallflowSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&DELETED=" & ACFID
		End If
	end if
End Function

Function RefreshClick()
	self.location.href = "AHCallFLowSummary.asp?<%= Request.QueryString %>"
End Function


Function AttachClick()
	lret = ""
	strURL = ""
	SearchObj.multiselected = ""
	strURL = "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=MODAL&LAUNCHER=SEARCH"
	lret = window.showModalDialog(strURL  ,SearchObj ,"dialogWidth:625px;dialogHeight:550px;center")
	if SearchObj.multiselected <> "" Then
		self.location.href = "AHCallFlowSummary.asp?ATTACH=TRUE&AHSID=<%= Request.QueryString("AHSID") %>&ATTACHCALLFLOW_ID=" & SearchObj.multiselected
	End If
End Function

Function SearchClick()
	lret = ""
	strURL = ""
	SearchObj.multiselected = ""
	strURL = "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=MODAL&AHSID=<%= Request.QueryString("AHSID") %>"
	lret = window.showModalDialog(strURL  ,SearchObj ,"dialogWidth:550px;dialogHeight:550px;center")
	if SearchObj.multiselected <> "" Then
		self.location.href = "AHCallFlowSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & SearchObj.multiselected
	End If
End Function

Sub window_onload
<% IF Request.QueryString("ErrMsg") <> "" Then %>
		StatusSpan.innerHTML = "<%= Request.QueryString("ErrMsg") %>"
<% ELSE %>
	<% If RS.RecordCount = MAXRECORDCOUNT Then %>
			StatusSpan.innerHTML = "<%= MSG_MAXRECORDS %>"
	<% Else %>
			StatusSpan.innerHTML = "Record Count is <%= RS.RecordCount %>"
	<% End If %>	
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
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
	
	PARAMS = ""
	If HasModifyPrivilege("FNSD_CALLFLOW","") <> True Then	PARAMS = PARAMS & "&HIDEEDIT=TRUE"
	If HasAddPrivilege("FNSD_CALLFLOW","") <> True Then	PARAMS = PARAMS & "&HIDENEW=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEATTACH=TRUE"
	If HasDeletePrivilege("FNSD_CALLFLOW","") <> True Then	PARAMS = PARAMS & "&HIDEREMOVE=TRUE"
%>
<OBJECT data="../Scriptlets/ObjButtons.asp?SEARCHCAPTION=Filter<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=CFBtnControl type=text/x-scriptlet></OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>A.C.F ID</div></td>
			<td class=thd><div id="NAME_HEAD"><NOBR>C.F. ID</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>LOB</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Name</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Description</div></td>
			
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
		<tr ID="FieldRow" CLASS="ResultRow" CFNAME="<%= RS("NAME") %>" ACCOUNTCALLFLOWID='<%= RS("ACCOUNTCALLFLOW_ID") %>' CALLFLOWID='<%= RS("CALLFLOW_ID") %>' OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);">
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("ACCOUNTCALLFLOW_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("CALLFLOW_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("LOB_CD")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("NAME")) %></td>
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
