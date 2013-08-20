<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\AHSTree.inc"--> 
<% Response.Expires=0 
   Response.Buffer = True
   On Error Resume Next
%>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

If Request.QueryString("INSERT") <> "" Then
	InsertSQL = ""
		DEPTID = CLng(NextPkey("DEPARTMENT_CODES","DEPARTMENT_CODES_ID"))
		If NewDEPTID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "DEPARTMENT_CODES", "DEPARTMENT_CODES_ID", NewDEPTID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
		RS.Close
		Conn.Close
	    END IF
	'SQLINS = "{call Designer.AddAccountUser('" & Request.QueryString("AHSID") & "', '" &_
	'		 Request.QueryString("INSERT") & "', {resultset  2, StatusMsg,StatusNum})}"
	'Set RS = Conn.Execute(SQLINS)			
	'if CStr(RS("StatusNum")) <> "0" then
	'	strError = RS("StatusMsg")
	'End If
	'RS.Close
	'Conn.Close		
	If strError = "" Then	
		Response.Redirect "AHLocationDeptSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If
End If

If Request.QueryString("DELETED") <> "" Then

	SQLDEL = "DELETE FROM DEPARTMENT_CODES WHERE DEPARTMENT_CODES_ID = " & Request.QueryString("DELETED")
	SQLDEL = SQLDEL & " AND ACCNT_HRCY_STEP_ID = " & Request.QueryString("AHSID")
	Set RS = Conn.Execute(SQLDEL)
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing
	Response.Redirect "AHLocationDeptSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

'If Request.QueryString("COPY") <> "" Then 

'	SQLCOPY = "INSERT INTO ACCOUNT_USER (ACCNT_HRCY_STEP_ID, USER_ID) VALUES ("
'	SQLCOPY = SQLCOPY & Request.QueryString("AHSID") & " , " & Request.QueryString("COPY") & ")"
'	Set RS = Conn.Execute(SQLCOPY)
'	strError = CheckADOErrors(Conn,"COPY" )
'
'	Set RS = Nothing
'	Conn.Close
'	Set Conn = Nothing

'	If strError = "" Then 
'		Response.Redirect "AHAccountUserSummary.asp?AHSID=" & Request.QueryString("AHSID")
'	End If

'End If

'If Request.QueryString("CLEARFILTER") <> "" Then
'	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_USERFILTER"
'	Response.Redirect "AHAccountUserSummary.asp?AHSID=" & Request.QueryString("AHSID")
'End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	SQL = "SELECT * FROM DEPARTMENT_CODES DC " &_
		"WHERE DC.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")

	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_USERFILTER", "MUSTINCLUDE")	
	
	If Request.QueryString("MultiSelected") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.QueryString("multiselected")
		Else
			strInclude  = Request.QueryString("multiselected")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_USERFILTER", "MUSTINCLUDE", strInclude
	End If
	
	if strInclude <> "" then SQL = SQL & " AND DEPARTMENT_CODES.DEPARTMENT_CODES_ID IN (" & strInclude & ") "
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="Javascript">
<!--
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("DEPTID");
}
function getname( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("USERNAME");
}
function CDeptSearchObj()
{
	this.DEPTID = "";
	this.DEPTIDName = "";
	this.Selected = "";
}
var DeptSearchObj = new CDeptSearchObj();
function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHLocationDeptSummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}


-->
</SCRIPT>
<!-- #include file="..\lib\AUBtnControl.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../Department/DepartmentMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&DEPTID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function



Function NewClick()
	parent.frames.window.location = "../Department/DepartmentMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DEPTID=NEW&DETAILONLY=TRUE&AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Function RemoveClick()

	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		DEPTID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to detach Department ID:" & DEPTID & " from Account ID:<%=Request.QueryString("AHSID")%>?", 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "AHLocationDeptSummary.asp?DELETED=" & DEPTID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	end if
End Function

Function SearchClick()
	lret = ""
	strURL = ""
	DeptSearchObj.DEPTID = ""
	strURL = "../Department/DepartmentMaintenance.asp?CONTAINERTYPE=MODAL&SEARCHONLY=TRUE"
	lret = window.showModalDialog(strURL  ,UserSearchObj ,"center")
	if UserSearchObj.UID <> "" Then
		self.location.href = "AHLocationDeptSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & DeptSearchObj.DEPTID
	End If
End Function


Function RefreshClick()
	self.location.href = "AHLocationDeptSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
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

<% If strError <> "" Then %>
	MsgBox ("<%=strError%>")
<% End If %>


End Sub

</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
PARAMS = "&HIDEEDIT=false"
PARAMS = PARAMS & "&HIDENEW=false&HIDECOPY=TRUE&HIDEPASTE=TRUE"
PARAMS = PARAMS & "&HIDEREMOVE=false&HIDEREFRESH=TRUE&HIDESEARCH=TRUE"
%>
<OBJECT id=AUBtnControl style="LEFT: 0px; WIDTH: 100%; HEIGHT: 23px" type=text/x-scriptlet data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE&amp;SEARCHCAPTION=Filter<%=PARAMS%>">
	</OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="ID_HEAD"><NOBR>Dept ID</div></td>
			<td class=thd><div id="Name_HEAD"><NOBR>Dept Name</div></td>
			<td class=thd><div id="Code_HEAD"><NOBR>Dept Code</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  DEPTID='<%= RS("DEPARTMENT_CODES_ID") %>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("DEPARTMENT_CODES_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("DEPARTMENT_NAME")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("DEPARTMENT_CODE")) %></td>
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

<OBJECT ID="ClipboardAgent" <%GetClipboardCLSID()%> width=1 height=1 VIEWASTEXT>
<PARAM NAME="MaxPropertiesStringLength" VALUE="1000">
<PARAM NAME="MaxPropertyNameLength" VALUE="50">
<PARAM NAME="MaxPropertyValueLength" VALUE="200">
<PARAM NAME="NameValueDelimiter" VALUE="#">
<PARAM NAME="PropertyItemDelimiter" VALUE="|">
<PARAM NAME="PrivateClipboardFormatName" VALUE="CF_FNSDESIGNER">
</OBJECT>
</BODY>
</HTML>
