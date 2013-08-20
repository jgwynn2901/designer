<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\CheckSharedCallFlow.inc"-->
<!--#include file="..\lib\RefCountRptinc.asp"-->
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\CommonError.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<%
Response.Expires=0
Response.AddHeader  "Pragma", "no-cache"

If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then   
	Session("NAME") = ""
	Response.Redirect "CallFlowMaintenanceEditor.asp"
End If
If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"
If HasDeletePrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then DELETE = "RO"
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	
Function Swap(InData)
	If InData <> "" Then
		Swap = InData
	Else
		Swap = "null"
	End If
End Function

Function NextPkey( TableName, ColName )
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function

Function SwapFlg(InData)
	If InData = "on" Then
		SwapFlg = "Y"
	Else
		SwapFlg = "N"
	End If
End Function

If Request.QueryString("COPYERR") <> "" then
	'cErr = Request.QueryString("cErr")
	LogStatusGroupBegin
	LogStatus S_ERROR, "Paste Frame operation failed", "FRAME_ORDER", "", 0, ""
	LogStatusGroupEnd
end if
'else
If Request.QueryString("DELETECALLFLOW") <> "" AND Isnumeric(Request.QueryString("DELETECALLFLOW")) Then 
	SQLDELETE = ""
	SQLDELETE = SQLDELETE & "{ call Designer.DeleteCallFlow(" & Request.QueryString("DELETECALLFLOW") & ")}"
	Set RSDel = Conn.Execute(SQLDELETE)
	Response.Redirect "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=FRAMEWORK"
End If

If Request.QueryString("COPYASNEW") <> "" Then
	SQL = ""
	SQL = "{call Designer.CopyFrame(" &  Request.QueryString("FRAMEID") & " ,{resultset 1, outFrameId,StatusMsg,StatusNum})}"
	Set RSCopy = Conn.Execute(SQL)
	SQL = ""
	SQL = SQL & "{call Designer.CopyFrameOrder(" &  Request.QueryString("FRAMEID") & ", "
	SQL = SQL & Request.QueryString("CALLFLOW_ID") & "," 
	SQL = SQL & RSCopy("outFrameId") & ","
	SQL = SQL & Request.QueryString("CALLFLOW_ID") & ",{resultset 1, StatusMsg, StatusNum})}"
	Set RS=Conn.Execute(SQL)
	SQLDelete = ""
	SQLDelete = SQLDelete & "DELETE FROM FRAME_ORDER WHERE CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID") & " AND "
	SQLDelete = SQLDelete & "FRAME_ID=" & Request.QueryString("FRAMEID")
	Set RSDel = Conn.Execute(SQLDelete)
	Response.Redirect "CallFlowMaintenanceEditor.asp?CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID")
End If

If Request.QueryString("DETACH") <> "" Then
	SQLDelete = ""
	SQLDelete = SQLDelete & "DELETE FROM FRAME_ORDER WHERE CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID") & " AND "
	SQLDelete = SQLDelete & "FRAME_ID=" & Request.QueryString("FRAMEID")
	Set RSDel = Conn.Execute(SQLDelete)
	Response.Redirect "CallFlowMaintenanceEditor.asp?CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID")
End If

If Request.QueryString("ATTACH") <> "" Then
	SQL2 = ""
	SQL2 = SQL2 & "INSERT INTO FRAME_ORDER (FRAME_ID, CALLFLOW_ID" 
	SQL2 = SQL2 & ", SEQUENCE "
	SQL2 = SQL2 & ",TITLE, ATTRIBUTE_PREFIX, ENABLEDRULE_ID, VALIDRULE_ID, "
	SQL2 = SQL2 & "MODAL_FLG, ENTRY_ACTION_ID, ACTION_ID, HELPSTRING, "
	SQL2 = SQL2 & "DESCRIPTION, TYPE, SQLSELECT, SQLFROM, SQLWHERE, SQLORDERBY, "
	SQL2 = SQL2 & "MAXPAGERESULTROWS, ONEROWAUTOSELECT_FLG "
	SQL2 = SQL2 & ") VALUES ("
	SQL2 = SQL2 & Request.QueryString("FRAMEID") & ", " 
	SQL2 = SQL2 & Request.QueryString("CALLFLOW_ID") 
	SQL2 = SQL2 & ",0"
	SQL2 = SQL2 & ",'-999999999', '-999999999', -999999999, -999999999, "
	SQL2 = SQL2 & "'U', -999999999, -999999999, '-999999999', "
	SQL2 = SQL2 & "'-999999999', '-999999999', '-999999999', '-999999999', '-999999999', '-999999999', "
	SQL2 = SQL2 & "-999999999, 'U' "
	SQL2 = SQL2 & ")" 
	Set RS=Conn.Execute(SQL2)
	Response.Redirect "CallFlowMaintenanceEditor.asp?CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID")
End If

If Request.QueryString("COPY") <> "" Then
	'SQL = ""
	'SQL = "{call Designer.CopyFrame(" &  Request.QueryString("FRAMEID") & " ,{resultset 1, outFrameId, StatusMsg, StatusNum})}"
	'Set RSCopy = Conn.Execute(SQL)

	SQL = ""
	SQL = SQL & "{call Designer.CopyFrameOrder(" &  Request.QueryString("FRAMEID") & ", "
	SQL = SQL & Request.QueryString("FROM_CALLFLOW_ID") & "," 
	'SQL = SQL & RSCopy("outFrameId") & ","
	SQL = SQL & Request.QueryString("FRAMEID") & ","
	SQL = SQL & Request.QueryString("CALLFLOW_ID") & ",{resultset 1, StatusMsg, StatusNum})}"
	Set RS=Conn.Execute(SQL)
	if cint(RS.Fields("StatusNum")) <> 0 then
		cErr = RS.Fields("StatusMsg")
		Response.Redirect "CallFlowMaintenanceEditor.asp?COPYERR=Y&CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID")	
	else
		Response.Redirect "CallFlowMaintenanceEditor.asp?CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID")	
	End If
End If

If Request.QueryString("ACTION") = "SAVE" AND  Request.QueryString("CALLFLOW_ID") = "NEW" Then
	NextID = NextPkey("CALLFLOW", "CALLFLOW_ID")
	SQLIN = ""
	SQLIN = SQLIN & "INSERT INTO CALLFLOW (CALLFLOW_ID, CALLTYPE_ID, CAT_FLG, "
	SQLIN = SQLIN & "NAME, DESCRIPTION) VALUES ("
	SQLIN = SQLIN & NextID & ", "
	SQLIN = SQLIN & "0" & ", "
	SQLIN = SQLIN & "'" & SwapFlg(Request.Form("CAT_FLG")) & "', "
	SQLIN = SQLIN & "'" & Request.Form("NAME") & "', "
	SQLIN = SQLIN & "'" & Request.Form("DESCRIPTION") & "') "
	Set RSIN = Conn.Execute(SQLIN)
	Response.Redirect "CallFlowMaintenanceEditor.asp?STATUS=SAVED&CALLFLOW_ID=" & NextID
End If

If Request.QueryString("ACTION") = "SAVE" AND  Request.QueryString("CALLFLOW_ID") <> "NEW" Then
	SQLUP = ""
	SQLUP = SQLUP & "UPDATE CALLFLOW SET "
	SQLUP = SQLUP & "CAT_FLG='" & SwapFlg(Request.Form("CAT_FLG")) & "', "
	SQLUP = SQLUP & "NAME='" & Replace(Request.Form("NAME"), """", """""") & "', "
	SQLUP = SQLUP & "DESCRIPTION='" & Replace(Request.Form("DESCRIPTION"), """", """""") & "' WHERE "
	SQLUP = SQLUP & "CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID")
	Set RSUP = Conn.Execute(SQLUP)
	Response.Redirect "CallFlowMaintenanceEditor.asp?STATUS=SAVED&CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID")
End If

If Request.QueryString("CALLFLOW_ID") <> "NEW" Then
	SQL = ""
	SQL = "SELECT * FROM CALLFLOW WHERE CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID") 
	Set RSSelect = Conn.Execute(SQL)
	CALLFLOW_ID = RSSelect("CALLFLOW_ID")
	CAT_FLG = RSSelect("CAT_FLG")
	NAME = ReplaceQuotesInText(RSSelect("NAME"))
	DESCRIPTION = ReplaceQuotesInText(RSSelect("DESCRIPTION"))
	SharedCallFlowCount = CheckSHaredCallFlow(CALLFLOW_ID,True, True, 2, False, False, 0) 
Else
	CALLFLOW_ID = "NEW"
End If
'end if
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Call Flow Editor</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

<!--#include file="..\lib\Help.asp"-->

dim aFrames()

Sub window_onload
	ClipboardAgent.GetPropertiesFromClipboard
<%
dim x

x = 0
If Request.QueryString("CALLFLOW_ID") <> "NEW" Then
	StrSQL = ""
	StrSQL = StrSQL & "SELECT FRAME_ORDER.SEQUENCE, FRAME_ORDER.FRAME_ID, FRAME.NAME FROM FRAME_ORDER, FRAME WHERE CALLFLOW_ID=" & CALLFLOW_ID & " AND "	
	StrSQL = StrSQL & "FRAME.FRAME_ID = FRAME_ORDER.FRAME_ID ORDER BY FRAME_ORDER.SEQUENCE"
	Set RS = Conn.Execute(StrSQL)
%>
NodeX = TreeView1.AddNode ("",1 , "CFID=<%= CALLFLOW_ID %>", "CALLFLOW", "Call Flow: <%= NAME %>, <%= CALLFLOW_ID %>", "FOLDER", "FOLDERSEL")	
<% Do While Not RS.EOF %>
	redim preserve aFrames(<%=x%>)
	aFrames(<%=x%>) = "<%=RS.fields("FRAME_ID") %>"
	<%x = x + 1%>
	NodeX = TreeView1.AddNode ("CFID=<%= CALLFLOW_ID %>", 4, "FRAMEID=<%= RS("FRAME_ID") %>" , "FRAME", "(<%= RS("SEQUENCE") %>) Frame: <%= RS("NAME") %>, <%= RS("FRAME_ID") %>", "FOLDER", "FOLDERSEL")	
<%
RS.movenext
loop
RS.close
%>

<% Else %>
	SpanStatus2.innerHTML = "New Call Flow"
	NodeX = TreeView1.AddNode ("",1 , "CFID=<%= CALLFLOW_ID %>", "FRAME", "New Call Flow", "FOLDER", "FOLDERSEL")	
<% End If %>
<% If Request.QueryString("CALLFLOW_ID") <> "NEW" Then %>
lret = TreeView1.AddMenuItem("FRAME", "&Visual Editor", ErrStr)
lret = TreeView1.AddMenuItem("CALLFLOW", "&Visual Editor", ErrStr)
<% If MODE <> "RO" Then %>
lret = TreeView1.AddMenuItem("FRAME", "-", ErrStr)
lret = TreeView1.AddMenuItem("FRAME", "&Copy As New", ErrStr)
lret = TreeView1.AddMenuItem("CALLFLOW", "&Attach Frame", ErrStr)
lret = TreeView1.AddMenuItem("FRAME", "&Attach Frame", ErrStr)
lret = TreeView1.AddMenuItem("FRAME", "Copy &Frame", ErrStr)
lret = TreeView1.AddMenuItem("FRAME", "&Detach Frame", ErrStr)
<% End If
End If %>

If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("FRAME_ID") <> "" and "<%=Request.QueryString("CALLFLOW_ID")%>" <> "NEW" Then
	lret = TreeView1.AddMenuItem("CALLFLOW", "&Paste Frame", ErrStr)
	lret = TreeView1.AddMenuItem("FRAME", "&Paste Frame", ErrStr)
End If

TreeView1.ExpandNode("CFID=<%= CALLFLOW_ID %>")
	<% If CAT_FLG = "Y" Then %>
		document.all.CAT_FLG.checked = true
	<% End If %>
<% If Request.QueryString("STATUS") = "SAVED" Then %>
	SpanStatus2.innerHTML = "Saved"
<% End If %>

<% If SharedCallFlowCount > 1 Then %>
	SpanStatus.innerHTML = "Warning! "
	SpanStatus2.innerhtml = "Shared count is greater than 1"
	<%	If CInt(SharedCallFlowCount) = CInt(Application("MaximumSharedCount")) Then %>
			SpanSharedCount.innerHTML = "<%=SharedCallFlowCount%>" & "<Font size=1 Color='Maroon'>+</Font>"
	<%	Else %>
			SpanSharedCount.innerHTML = "<%=SharedCallFlowCount%>"
	<%  End If 
   End if 

If Request.QueryString("COPYERR") <> "" then %>
	SpanTreeStatus.innerhtml = "Error! Paste Frame operation failed."
	SpanTreeStatus.style.color = "Red"
<%end if%>
End Sub


Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
dim nTop, x, lFound

Select Case MenuItem
	Case "&Visual Editor"
		SpanTreeStatus.innerHTML = "Visual Editor changes may not be reflected until refresh"
		SpanTreeStatus.style.color = "Maroon"
		If NodeType = "CALLFLOW" Then
			Key = NodeKey
		Else
			Key = NodeKey & "&CFID=<%= CALLFLOW_ID %>"
		End If
		If TreeView1.Nodes(TreeView1.selected).children < 1 AND NodeType = "CALLFLOW" Then
			Call LaunchNewCFEditor(Key)
		Else
			Call LaunchCFEditor(Key)
		End If
	Case "Copy &Frame"
		ClipboardAgent.ClearAllProperties()
		lret = TreeView1.AddMenuItem("FRAME", "&Paste Frame", ErrStr)
		lret = TreeView1.AddMenuItem("CALLFLOW", "&Paste Frame", ErrStr)	
		ClipboardAgent.AddProperty "FRAME_ORDER_DATA", "GOT DATA"
		ClipboardAgent.AddProperty "FRAME_ID", NodeKey
		ClipboardAgent.AddProperty "FROM_CALLFLOW_ID", "<%=Request.QueryString("CALLFLOW_ID")%>"
		ClipboardAgent.SetPropertiesToClipboard
	Case "&Paste Frame"
		ClipboardAgent.GetPropertiesFromClipboard
		If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("FRAME_ORDER_DATA") <> "" Then
			lret = MsgBox ("Are you sure you wish to paste a frame in this call flow", 1, "FNSDesigner")
			if lret = "1" Then
				self.location.href = "CallFlowMaintenanceEditor.asp?COPY=TRUE&" & ClipboardAgent.GetProperty("FRAME_ID") & "&FROM_CALLFLOW_ID=" & ClipboardAgent.GetProperty("FROM_CALLFLOW_ID") &"&CALLFLOW_ID=<%=Request.QueryString("CALLFLOW_ID")%>"
			End If
		Else
			Msgbox "Nothing to paste", 0, "FNSDesigner"
		End If
	Case "&Attach Frame"
		strURL = "FrameSearchModal.asp"
		showModalDialog  strURL  ,FrameObj ,"dialogWidth:450px;dialogHeight:450px;center"
		If FrameObj.FrameID <> "" Then
			nTop = uBound(aFrames)
			lFound = false
			for x=0 to nTop	'	option base 0
				if FrameObj.FrameID = aFrames(x) then
					lFound = true
					exit for
				end if
			next
			if lFound then
				msgbox "The frame " & FrameObj.FrameID & " is already in the Callflow.",vbexclamation,"FNSDesigner"
			else
				self.location.href = "CallFlowMaintenanceEditor.asp?ATTACH=TRUE&FRAMEID=" & FrameObj.FrameID  & "&CALLFLOW_ID=<%= CALLFLOW_ID %>"
			end if
		End If
	Case "&Detach Frame"
			lret = msgbox("Are you sure you want to detach this frame?", 1, "FNSDesigner")
			if lret = "1" Then
				self.location.href = "CallFlowMaintenanceEditor.asp?DETACH=TRUE&" & NodeKey & "&CALLFLOW_ID=<%= CALLFLOW_ID %>"
			End If
	Case "&Copy As New"
		lret = msgbox ("Are you sure you want to copy this frame: " & NodeText & Chr(13) & "Copying this frame will create a new unique instance of this frame" & VbCrlf & "and the current frame will be detached.", 1, "FNSNet")
		if lret = "1" Then
			self.location.href = "CallFlowMaintenanceEditor.asp?COPYASNEW=TRUE&" & NodeKey & "&CALLFLOW_ID=<%= CALLFLOW_ID %>"
		End If
End Select
End Sub

Function Handles(Obj, Title)
	If InStr(1, top.frames("TOP").location.href, "Toppane.asp") <> 0 Then
		lret = top.frames("TOP").SetHandle(Obj, Title)
	End If
End Function

Sub BtnGrfxBack_Onclick()
	self.location.href = "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=FRAMEWORK"
End Sub

Sub BtnClear_onclick
	document.all.NAME.value = ""
	document.all.DESCRIPTION.value = ""
	document.all.CAT_FLG.checked = false
End Sub

Sub BtnNew_onclick
	self.location.href = "CallFlowMaintenanceEditor.asp?CALLFLOW_ID=NEW"
End Sub

Sub BtnSave_onclick
strErr = ""
If document.all.NAME.value = "" Then
	strErr = strErr & "Name is a required field" & VBCRLF
End If
If document.all.Description.value = "" Then
	strErr = strErr & "Description is a required field" & VBCRLF
End If
If strErr = "" Then 
	document.all.FrmSave.submit()
Else
	msgbox strErr, 0 , "FNSDesigner"
End If
End Sub

Sub BtnRefresh_onclick
	self.location.href = "CallFLowMaintenanceEditor.asp?CALLFLOW_ID=<%= CALLFLOW_ID %>"
End Sub

Sub StatusRpt_onclick
	If CLng(<%=SharedCallFlowCount%>) > 1 Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other details reported", 0 , "FNSNet"
	End If
End Sub

Sub RefCountRpt_onclick
	lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedCallFLow=True&ID=<%= CALLFLOW_ID %>", Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
End Sub

Sub BtnDelete_onclick
lret = ""
lret = window.showModalDialog("Areyousure.asp?COUNT=<%=SharedCallFlowCount%>", null, " center=yes")
if lret = "DELETE" Then
	self.location.href = "CallFLowMaintenanceEditor.asp?DELETECALLFLOW=<%= CALLFLOW_ID %>"
end if
End Sub

-->
</script>
<script LANGUAGE="JavaScript">
function LaunchCFEditor(key) {
Url = "CallFlow-f.asp?" + key
var VisEditorObj = window.open(Url, null, "height=600,width=800,status=no,toolbar=no,menubar=no,location=no,resizable=yes,top=0,left=0");
lret = Handles( VisEditorObj, "CALLFLOW");
VisEditorObj.focus()
}
function LaunchNewCFEditor(key) {
Url = "NewFrame.asp?" + key
var VisEditorObj2 = window.open(Url, null, "height=600,width=800,status=no,toolbar=no,menubar=no,location=no,resizable=yes,top=0,left=0");
lret = Handles( VisEditorObj2, "CALLFLOW");
VisEditorObj2.focus()
}
function CanDocUnloadNow(){
	if (false == confirm("Data has changed. Leave page without saving?"))
		return false;
	else
		return true;
}
function CRuleSearchObj(){
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CFrameSearchObj(){
	this.FrameID = "";
}
function CRPSearchObj(){
	this.routing_plan_id = "";
	this.ahsid = "";
	this.multiselected = "";
}
var SearchObj = new CRPSearchObj();
var FrameObj = new CFrameSearchObj();
var RuleSearchObj = new CRuleSearchObj();
</script>
</head>
<body BGCOLOR="#d6cfbd" topmargin="0" rightmargin="0" leftmargin="0" bottommargin="0">
<!--#include file="..\lib\NavBack.inc"-->
<form NAME="FrmSave" ACTION="CallFlowMaintenanceEditor.asp?ACTION=SAVE&amp;CALLFLOW_ID=<%= CALLFLOW_ID %>" METHOD="POST">
<table WIDTH="98%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Call Flow
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table cellpadding="0" cellspacing="0">
<td WIDTH="14" CLASS="LABEL">
<img ID="RefCountRpt" SRC="..\images\RefCount.gif" STYLE="CURSOR:HAND" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Reference Count">
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="10">
:<span id="SpanSharedCount" CLASS="LABEL"><%=SharedCallFlowCount%></span>
</td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" STYLE="CURSOR:HAND" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<font COLOR="MAROON">
<span ID="SpanStatus" CLASS="LABEL" STYLE="COLOR:#FF0000"></span>
<span ID="SpanStatus2" CLASS="LABEL" STYLE="COLOR:#006699">Ready</span></font>
</td>
</table>
<table WIDTH="100%" cellpadding="2" cellspacing="0"><tr><td>
<table BORDER="0" cellpadding="0" cellspacing="0">
<tr>
<td CLASS="LABEL">Call Flow ID:(<%= CALLFLOW_ID %>)</td>
</tr>
<tr>
<td CLASS="LABEL">Name:<br><input TYPE="TEXT" SIZE="40" <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>CLASS="LABEL" NAME="NAME" VALUE="<%= NAME %>"></td>
<td CLASS="LABEL" VALIGN="BOTTOM"><input TYPE="CHECKBOX" id="CAT_FLG" name="CAT_FLG" <% If MODE="RO" Then Response.write(" DISABLED ") %>>Catastrophe:</td>
</tr>
<tr>
<td CLASS="LABEL" COLSPAN="2">Description:<br><input TYPE="TEXT" <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> SIZE="70" CLASS="LABEL" id="Description" name="Description" VALUE="<%= DESCRIPTION %>"></td>
</tr>
<tr>
</table>
</td><td VALIGN="TOP" ALIGN="RIGHT">
<table>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE="RO" Then Response.write(" DISABLED ") %> NAME="BtnSave" ACCESSKEY="S"><u>S</u>ave</button>
</tr>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE="RO" Then Response.write(" DISABLED ") %> NAME="BtnNew" ACCESSKEY="N"><u>N</u>ew</button>
</tr>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE="RO" Then Response.write(" DISABLED ") %> NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button>
</tr>
<% If CALLFLOW_ID <> "NEW" Then %>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If DELETE="RO" Then Response.write(" DISABLED ") %> NAME="BtnDelete" ACCESSKEY="D"><u>D</u>elete</button>
</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Call Flow Frames
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<img SRC="..\IMAGES\Refresh.gif" STYLE="CURSOR:HAND" ALIGN="LEFT" ALT="Refresh" ID="BtnRefresh">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" STYLE="CURSOR:HAND" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
: <span ID="SpanTreeStatus" CLASS="LABEL" STYLE="COLOR:#006699">Ready</span>
<br>
<object CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0">
	<param NAME="LPKPath" VALUE="LPKfilename.LPK">
</object>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="57%">
<param NAME="ShowTips" VALUE="False">
</object>
<object ID="ClipboardAgent" <%GetClipboardCLSID()%> width="1" height="1">
<param NAME="MaxPropertiesStringLength" VALUE="1000">
<param NAME="MaxPropertyNameLength" VALUE="50">
<param NAME="MaxPropertyValueLength" VALUE="200">
<param NAME="NameValueDelimiter" VALUE="#">
<param NAME="PropertyItemDelimiter" VALUE="|">
<param NAME="PrivateClipboardFormatName" VALUE="CF_FNSDESIGNER">
</object>