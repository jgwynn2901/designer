<%
'***************************************************************
'This fle is called from the hierarchy tree, when selecting a node, 
'and displays information related to Mailboxes.
'
'$History: AHMailboxSummary.asp $                       
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:38p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/AH
'* Hartford SRS: Initial revision
'***************************************************************
%>
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
'	get next primary key
Function NextPkey( TableName, ColName )
	NextSQL = "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("NextID") 
End Function

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

If Request.QueryString("INSERT") <> "" Then

	SQLINS = "INSERT INTO MAILBOX_ASSIGNMENT_TYPE(MAILBOX_ASSIGNMENT_TYPE_ID,ACCNT_HRCY_STEP_ID) VALUES (" 
	SQLINS = SQLINS & NextPkey("MAILBOX_ASSIGNMENT_TYPE","MAILBOX_ASSIGNMENT_TYPE_ID") & "," & Request.QueryString("INSERT") & ")"
	Set RS = Conn.Execute(SQLINS)
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing
	Response.Redirect "AHMailboxSummary.asp?AHSID=" & Request.QueryString("AHSID") 
End If
If Request.QueryString("DELETED") <> "" Then
	on error resume next
	Conn.BeginTrans 
	SQLDEL = "DELETE FROM MAILBOX_ASSIGNMENT_RULE WHERE MAILBOX_ASSIGNMENT_TYPE_ID = " & Request.QueryString("DELETED") 
	Conn.Execute(SQLDEL)
	if err.number = 0 then
		SQLDEL = "DELETE FROM MAILBOX_ASSIGNMENT_TYPE WHERE MAILBOX_ASSIGNMENT_TYPE_ID = " & Request.QueryString("DELETED") 
		Set RS = Conn.Execute(SQLDEL)
	end if

	'Trilok - Audit Delete Changes.
	SQLDEL = " UPDATE MAILBOX_ASSIGNMENT_TYPE_AUDIT set LAST_MODIFIED_BY =" & Session("SecurityObj").m_UserId & " where ACTION='DELETED' and MAILBOX_ASSIGNMENT_TYPE_ID ='" & Request.QueryString("DELETED") & "'"
	Conn.Execute(SQLDEL)

	if err.number <> 0 then
		Conn.RollbackTrans 
	else
		Conn.CommitTrans 
	end if
	strError = CheckADOErrors(Conn,"DELETE" )
	Conn.Close
	Set Conn = Nothing
	
	If strError = "" Then 
		Response.Redirect "AHMailboxSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If

End If
If Request.QueryString("COPY") <> "" Then 
	SQLCOPY = "{call Designer.SP_COPY_BRANCHASSIGNMENTTYPE(" & Request.QueryString("COPY") 
	SQLCOPY = SQLCOPY & " ," & Request.QueryString("AHSID") & ",{resultset  1, outResult})}"
	Set RS = Conn.Execute(SQLCOPY)

	strError = CheckADOErrors(Conn,"COPY" )

	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing

	If strError = "" Then 
		Response.Redirect "AHMailboxSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If

End If

If Request.QueryString("CLEARFILTER") <> "" Then
	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_MAFILTER"
	Response.Redirect "AHMailboxSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	
	
			
	SQL = "SELECT * FROM MAILBOX_ASSIGNMENT_TYPE, RULES  WHERE MAILBOX_ASSIGNMENT_TYPE.RULE_ID = RULES.RULE_ID(+) AND " &_
		 "MAILBOX_ASSIGNMENT_TYPE.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	
	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_MAFILTER", "MUSTINCLUDE")	
	
	If Request.QueryString("MultiSelected") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.QueryString("multiselected")
		Else
			strInclude  = Request.QueryString("multiselected")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_MAFILTER", "MUSTINCLUDE", strInclude
	End If
	
	if strInclude <> "" then SQL = SQL & " AND MAILBOX_ASSIGNMENT_TYPE.MAILBOX_ASSIGNMENT_TYPE_ID IN (" & strInclude & ") "

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
	return objRow.getAttribute("MATID");
}
function getname( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("MRNAME");
}
function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHMailboxSummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}

function CMailboxAssignTypeSearchObj()
{
	this.MATID = "";
	this.Selected = "";
}
var MailboxAssignTypeSearchObj = new CMailboxAssignTypeSearchObj();
-->
</SCRIPT>
<!-- #include file="..\lib\BRBtnControl.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../MailboxAssignment/MailboxAssignTypeMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&MATID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function CopyClick()
	ClipboardAgent.ClearAllProperties()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		ClipboardAgent.AddProperty "MAILBOX_ASSIGNMENT_TYPE_ID", dblhighlight(Document.all.tblResult.rows(i))
		ClipboardAgent.AddProperty "BRANCH_TEXT", getname(Document.all.tblResult.rows(i))
		ClipboardAgent.SetPropertiesToClipboard
	end if
End Function

Function PasteClick()
	If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("BRANCHASSIGNMENTTYPE_ID") <> "" Then
	lret = msgbox ("Are you sure you want to make a copy of Branch Assignment Type """ & ClipboardAgent.GetProperty("BRANCH_TEXT") & """ (" & ClipboardAgent.GetProperty("BRANCHASSIGNMENTTYPE_ID") &  ") for Account ID:<%=Request.QueryString("AHSID")%>?" , 1, "FNSDesigner") 
		If lret = "1" Then 
			self.location.href = "AHMailboxSummary.asp?COPY=" & ClipboardAgent.GetProperty("BRANCHASSIGNMENTTYPE_ID") & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	Else
		MsgBox "No data to paste!", 0, "FNSDesigner"
	End If
End Function

Function NewClick()
	parent.frames.window.location = "../MailboxAssignment/MailboxAssignTypeMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&MATID=NEW&DETAILONLY=TRUE&AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		MATID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to delete Mailbox Assignment Type ID:" & MATID & " for Account ID:<%=Request.QueryString("AHSID")%>?", 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "AHMailboxSummary.asp?DELETED=" & MATID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	end if
End Function

Function SearchClick()
	lret = ""
	strURL = ""
	MailboxAssignTypeSearchObj.MATID = ""
	strURL = "../MailboxAssignment/MailboxAssignTypeMaintenance.asp?CONTAINERTYPE=MODAL&MODE=RO&SearchAHSID=<%= Request.QueryString("AHSID") %>"
	lret = window.showModalDialog(strURL, MailboxAssignTypeSearchObj ,"center")
	if MailboxAssignTypeSearchObj.MATID <> "" Then
		multi = Replace(MailboxAssignTypeSearchObj.MATID,"||",",")
		self.location.href = "AHMailboxSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & multi
	End If
End Function


Function RefreshClick()
	self.location.href = "AHMailboxSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
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
	
	PARAMS = "&HIDECOPY=TRUE&HIDEPASTE=TRUE"
	If HasModifyPrivilege("FNSD_MAILBOX_ASSIGNMENT","") <> True Then	PARAMS = PARAMS & "&HIDEEDIT=TRUE"
	If HasAddPrivilege("FNSD_MAILBOX_ASSIGNMENT","") <> True Then	PARAMS = PARAMS & "&HIDENEW=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE"
	If HasDeletePrivilege("FNSD_MAILBOX_ASSIGNMENT","") <> True Then	PARAMS = PARAMS & "&HIDEREMOVE=TRUE"
%>
<OBJECT data="../Scriptlets/ObjButtons.asp?SEARCHCAPTION=Filter&HIDEATTACH=TRUE<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="BRBtnControl" type=text/x-scriptlet></OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Branch_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0  rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div><NOBR>M.A.T. ID</div></td>
			<td class=thd><div><NOBR>Description</div></td>
			<td class=thd><div><NOBR>Rule Text</div></td>
			<td class=thd><div><NOBR>Rule ID</div></td>			
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" MRNAME="<%= Mid(RS("DESCRIPTION"),1,25) %>" MATID='<%= RS("MAILBOX_ASSIGNMENT_TYPE_ID") %>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("MAILBOX_ASSIGNMENT_TYPE_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("DESCRIPTION")) %></td>
			<td  TITLE="<%=ReplaceQuotesInText(renderCell(RS("RULE_TEXT")))%>" NOWRAP CLASS="ResultCell"><%=TruncateText(renderCell(RS("RULE_TEXT")),25)%></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("RULE_ID")) %></td>			
	</tr>
<% 
RS.MoveNext
Loop
RS.Close
Set RS = Nothing
Conn.Close
Set Conn = Nothing
%>
</tbody>
</table>
</DIV>
</FIELDSET>
</SCRIPT>

<OBJECT ID="ClipboardAgent" VIEWASTEXT
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
