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
	SQLINS = "{call Designer.AddAccountUser('" & Request.QueryString("AHSID") & "', '" &_
			 Request.QueryString("INSERT") & "', {resultset  2, StatusMsg,StatusNum})}"
	Set RS = Conn.Execute(SQLINS)			
	if CStr(RS("StatusNum")) <> "0" then
		strError = RS("StatusMsg")
	End If
	RS.Close
	Conn.Close		
	If strError = "" Then	
		Response.Redirect "AHAccountUserSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If
End If

If Request.QueryString("DELETED") <> "" Then

	SQLDEL = "DELETE FROM ACCOUNT_USER WHERE USER_ID = " & Request.QueryString("DELETED")
	SQLDEL = SQLDEL & " AND ACCNT_HRCY_STEP_ID = " & Request.QueryString("AHSID")
	Set RS = Conn.Execute(SQLDEL)
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing
	Response.Redirect "AHAccountUserSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

If Request.QueryString("COPY") <> "" Then 

	SQLCOPY = "INSERT INTO ACCOUNT_USER (ACCNT_HRCY_STEP_ID, USER_ID) VALUES ("
	SQLCOPY = SQLCOPY & Request.QueryString("AHSID") & " , " & Request.QueryString("COPY") & ")"
	Set RS = Conn.Execute(SQLCOPY)
	strError = CheckADOErrors(Conn,"COPY" )

	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing

	If strError = "" Then 
		Response.Redirect "AHAccountUserSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If

End If

If Request.QueryString("CLEARFILTER") <> "" Then
	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_USERFILTER"
	Response.Redirect "AHAccountUserSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	SQL = "SELECT * FROM USERS_SITE_VIEW, ACCOUNT_USER " &_
		"WHERE ACCOUNT_USER.USER_ID = USERS_SITE_VIEW.USER_ID AND ACCOUNT_USER.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")

	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_USERFILTER", "MUSTINCLUDE")	
	
	If Request.QueryString("MultiSelected") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.QueryString("multiselected")
		Else
			strInclude  = Request.QueryString("multiselected")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_USERFILTER", "MUSTINCLUDE", strInclude
	End If
	
	if strInclude <> "" then SQL = SQL & " AND ACCOUNT_USER.USER_ID IN (" & strInclude & ") "
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
	return objRow.getAttribute("UID");
}
function getname( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("USERNAME");
}
function CUserSearchObj()
{
	this.UID = "";
	this.UIDName = "";
	this.Selected = "";
}
var UserSearchObj = new CUserSearchObj();
function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHAccountUserSummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}


-->
</SCRIPT>
<!-- #include file="..\lib\AUBtnControl.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../Users/UserMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&UID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function CopyClick()
	ClipboardAgent.ClearAllProperties()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		ClipboardAgent.AddProperty "USER_ID", dblhighlight(Document.all.tblResult.rows(i))
		ClipboardAgent.AddProperty "USER_TEXT", getname(Document.all.tblResult.rows(i))
		ClipboardAgent.SetPropertiesToClipboard
	end if
End Function

Function PasteClick()
	If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("USER_ID") <> "" Then
	lret = msgbox ("Are you sure you want to attach User """ & ClipboardAgent.GetProperty("USER_TEXT") & """ (" & ClipboardAgent.GetProperty("USER_ID") & ") to Account ID:<%=Request.QueryString("AHSID")%>?" , 1, "FNSDesigner") 
		If lret = "1" Then 
			self.location.href = "AHAccountUserSummary.asp?COPY=" & ClipboardAgent.GetProperty("USER_ID") & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	Else
		MsgBox "No data to paste!", 0, "FNSDesigner"
	End If
End Function

Function NewClick()
	parent.frames.window.location = "../Users/UserMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&UID=NEW&DETAILONLY=TRUE&AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		UID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to detach User ID:" & UID & " from Account ID:<%=Request.QueryString("AHSID")%>?", 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "AHAccountUserSummary.asp?DELETED=" & UID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
	end if
End Function

Function SearchClick()
	lret = ""
	strURL = ""
	UserSearchObj.UID = ""
	strURL = "../Users/UserMaintenance.asp?CONTAINERTYPE=MODAL&SEARCHONLY=TRUE"
	lret = window.showModalDialog(strURL  ,UserSearchObj ,"center")
	if UserSearchObj.UID <> "" Then
		self.location.href = "AHAccountUserSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & UserSearchObj.UID
	End If
End Function


Function AttachClick
	UserSearchObj.Selected = false

	strURL = "../Users/UserMaintenance.asp?CONTAINERTYPE=MODAL&SEARCHONLY=TRUE"
	showModalDialog  strURL,UserSearchObj ,"center"

	If UserSearchObj.Selected = true And UserSearchObj.UID <> "" Then	
		self.location.href = "AHAccountUserSummary.asp?INSERT=" & UserSearchObj.UID & "&AHSID=<%= Request.QueryString("AHSID") %>"
	End If
End Function

Function RefreshClick()
	self.location.href = "AHAccountUserSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
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
-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
	PARAMS = ""
	If HasAutomaticSecurityPrivilege() = False Then
		If HasModifyPrivilege("FNSD_USERS","") <> True Then	PARAMS = PARAMS & "&HIDEEDIT=TRUE"
		If HasAddPrivilege("FNSD_USERS","") <> True Then	PARAMS = PARAMS & "&HIDENEW=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE"
		If HasDeletePrivilege("FNSD_USERS","") <> True Then	PARAMS = PARAMS & "&HIDEREMOVE=TRUE"
	End If
%>
<OBJECT data="../Scriptlets/ObjButtons.asp?REMOVECAPTION=Detach&SEARCHCAPTION=Filter<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="AUBtnControl" type=text/x-scriptlet></OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div id="NAME_HEAD"><NOBR>User ID</div></td>
			<td class=thd><div id="PHONE_HEAD"><NOBR>Name</div></td>
			<td class=thd><div id="EXTENSION_HEAD"><NOBR>Site</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" USERNAME="<%= RS("NAME") %>" UID='<%= RS("USER_ID") %>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("USER_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("NAME")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("SITE_NAME")) %></td>
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
