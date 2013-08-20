<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<% Response.Expires=0 
on error resume next
If HasViewPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "RoutingPlanSummaryTree.asp"
End If
If HasModifyPrivilege("FNSD_ROUTING_PLAN",SECURITYPRIV) <> True Then MODE = "RO"

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	WarnColor = "#006699"
	strWarning = "Ready"
If Request.QueryString("ROUTING_PLAN_ID") = "NEW" Then
SQL = ""
SQL = SQL & "INSERT INTO ROUTING_PLAN (ROUTING_PLAN_ID, "
SQL = SQL & ""
End If

If Request.QueryString("ROUTING_PLAN_ID") <> "NEW" Then 
	
	SQL = ""
	SQL = SQL & "SELECT ROUTING_PLAN.*, TRANSMISSION_SEQ_STEP.*, TRANSMISSION_TYPE.NAME "
	SQL = SQL & "FROM ROUTING_PLAN, TRANSMISSION_SEQ_STEP, TRANSMISSION_TYPE "
	SQL = SQL & "WHERE ROUTING_PLAN.ROUTING_PLAN_ID=" & Request.QueryString("routing_plan_id") & " AND "
	SQL = SQL & "ROUTING_PLAN.ROUTING_PLAN_ID = TRANSMISSION_SEQ_STEP.ROUTING_PLAN_ID (+) AND "
	SQL = SQL & "TRANSMISSION_SEQ_STEP.TRANSMISSION_TYPE_ID = TRANSMISSION_TYPE.TRANSMISSION_TYPE_ID (+) ORDER BY TRANSMISSION_SEQ_STEP.SEQUENCE"
	Set RS = Conn.Execute(SQL)
	If RS.EOF AND RS.BOF Then
		Session("ErrorMessage") = "SQL Statement " & SQl & "<BR><BR>Returned no records"
		Response.Redirect "directerror.asp"
	End if
	TransCount = 0
			
	Do While RS("SEQUENCE") <> "1"
		StrWarning = "Warning: Invalid Transmission Step sequence"	
		WarnColor = "Maroon"
		RS.MoveNext
		If RS.EOF Then
			exit do
		End If
	Loop
	RS.MoveFirst
End If
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->
Sub window_onload
<% If Request.QueryString("ROUTING_PLAN_ID") <> "NEW" Then %>
NodeX = TreeView1.AddNode ("",1 , "RPID=<%= Cstr(Request.QueryString("ROUTING_PLAN_ID")) %>", "ROOT", "Routing Plan: <%= Replace(RS("DESCRIPTION"), """","""""") %>,  <%= Cstr(Request.QueryString("ROUTING_PLAN_ID")) %>", "FOLDER", "FOLDERSEL")
<%
If RS("TRANSMISSION_SEQ_STEP_ID") <> "" Then
 Do While Not RS.EOF %>
	NodeX = TreeView1.AddNode ("RPID=<%= Request.QueryString("ROUTING_PLAN_ID") %>", 4 , "TRANSMISSION_SEQ_STEP_ID=<%= RS("TRANSMISSION_SEQ_STEP_ID") %>", "TRANSMISSION", "(<%= RS("SEQUENCE") %>) Transmission Sequence Step: <%= Replace(RS("DESTINATION_STRING"),"""","""""") %>, <%= Replace(RS("NAME"),"""","""""") %> , <%= RS("TRANSMISSION_SEQ_STEP_ID") %>" ,"TRANSMISSION", "TRANSMISSIONSEL" )
<%	
	SQL4 = "SELECT * FROM OUTPUT_SUBJECT_BODY WHERE TRANSMISSION_SEQ_STEP_ID=" & RS("TRANSMISSION_SEQ_STEP_ID")
	Set RS4 = Conn.Execute(SQL4)
	Do While Not RS4.EOF
%>
	NodeX = TreeView1.AddNode ("TRANSMISSION_SEQ_STEP_ID=<%= RS("TRANSMISSION_SEQ_STEP_ID") %>", 4 , "OSBID=<%= RS4("OUTPUT_SUBJECT_BODY_ID") %>&TRANSMISSION_SEQ_STEP_ID=<%= RS("TRANSMISSION_SEQ_STEP_ID")%>", "SUBJECT_BODY", "Subject File: <%= RS4("SUBJECT_FILE_NAME") %>;  Body File: <%= RS4("BODY_FILE_NAME") %>,  <%= RS4("OUTPUT_SUBJECT_BODY_ID") %>" ,"PAGE", "PAGESEL" )
<%
	RS4.MoveNext
	Loop
	RS4.Close
%>
<%
	SQL2 = ""
	SQL2 = SQL2 & "SELECT * FROM OUTPUT_ITEM WHERE TRANSMISSION_SEQ_STEP_ID = " & RS("TRANSMISSION_SEQ_STEP_ID") & "ORDER BY SEQUENCE "
	Set RS2 = Conn.Execute(SQL2)
	Do While Not RS2.EOF
		cOUTPUTDEF_ID = RS2("OUTPUTDEF_ID")
%>
	NodeX = TreeView1.AddNode ("TRANSMISSION_SEQ_STEP_ID=<%= RS("TRANSMISSION_SEQ_STEP_ID") %>", 4 , "OUTPUT_ITEM_ID=<%= RS2("OUTPUT_ITEM_ID") %>", "ITEM", "(<%= RS2("SEQUENCE") %>) Output Item: <%= RS2("OUTPUT_ITEM_ID") %>" ,"OUTPUTITEM", "OUTPUTITEMSEL" )	
<%
	SQL3 = ""
	SQL3 = SQL3 & "SELECT * FROM OUTPUT_DEFINITION WHERE OUTPUTDEF_ID=" & RS2("OUTPUTDEF_ID")
	Set RS3 = Conn.Execute(SQL3)
	Do While Not RS3.EOF
%>
NodeX = TreeView1.AddNode ("OUTPUT_ITEM_ID=<%= RS2("OUTPUT_ITEM_ID") %>", 4 , "ODID=<%= RS3("OUTPUTDEF_ID") %>&OUTPUT_ITEM_ID=<%= RS2("OUTPUT_ITEM_ID") %>", "DEFINITION", "Output Definition: <%= Replace(RS3("NAME"),"""","""""") %>, <%= RS3("OUTPUTDEF_ID") %>" ,"OUTPUTDEFINITION", "OUTPUTDEFINITIONSEL" )
<%
SQL4 = "SELECT * FROM OUTPUT_PAGE WHERE OUTPUTDEF_ID=" & RS3("OUTPUTDEF_ID") & " ORDER BY PAGE_NUMBER"
Set RS4 = Conn.Execute(SQL4)
Do While Not RS4.EOF
%>
NodeX = TreeView1.AddNode ("ODID=<%= RS3("OUTPUTDEF_ID") %>&OUTPUT_ITEM_ID=<%= RS2("OUTPUT_ITEM_ID") %>", 4 , "OUTPUT_ITEM_ID=<%= RS2("OUTPUT_ITEM_ID") %>&ODID=<%= RS3("OUTPUTDEF_ID") %>&OPID=<%= RS4("OUTPUT_PAGE_ID") %>", "PAGE", "Output Page: <%= RS4("NAME") %>, <%= RS4("OUTPUT_PAGE_ID") %>" ,"PAGE", "PAGESEL" )
<%
RS4.MoveNext
Loop
RS4.Close
SQL4 = "SELECT * FROM OUTPUT_FILE WHERE OUTPUTDEF_ID=" & RS3("OUTPUTDEF_ID")
Set RS4 = Conn.Execute(SQL4)
Do While Not RS4.EOF
%>
NodeX = TreeView1.AddNode ("ODID=<%= RS3("OUTPUTDEF_ID") %>&OUTPUT_ITEM_ID=<%= RS2("OUTPUT_ITEM_ID") %>", 4 , "OUTPUT_ITEM_ID=<%= RS2("OUTPUT_ITEM_ID") %>&OUTPUTDEF_ID=<%= RS3("OUTPUTDEF_ID") %>&OFID=<%= RS4("OUTPUT_FILE_ID") %>", "FILE", "Output File: <%= RS4("OUTPUT_FILE_NAME") %>, <%= RS4("OUTPUT_FILE_ID") %>" ,"PAGE", "PAGESEL" )
<%
RS4.MoveNext
Loop
RS4.Close
RS3.MoveNext
Loop
RS3.Close
%>

<%
RS2.MoveNext
Loop
RS2.Close
%>

<%
RS.MoveNext
Loop
RS.Close
%>
<% End If %>
	ClipboardAgent.GetPropertiesFromClipboard
	<% If MODE <> "RO" Then %>
	If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("TRANSMISSION_SEQ_STEP_ID") <> "" Then
		lret = TreeView1.AddMenuItem("ROOT", "&Paste Transmission Sequence Step", ErrStr)
	End If
	<% End If %>
	
	<% If MODE <> "RO" Then %>
	lret = TreeView1.AddMenuItem("ROOT", "&New Transmission Sequence Step", ErrStr)
	lret = TreeView1.AddMenuItem("ROOT", "&Copy Routing Plan", ErrStr)
	lret = TreeView1.AddMenuItem("DEFINITION", "New &File Output", ErrStr)
	lret = TreeView1.AddMenuItem("TRANSMISSION", "Attach &Email Subject/Body Templates", ErrStr)
	<% End If %>
	lret = TreeView1.AddMenuItem("FILE", "&File Output Properties", ErrStr)
	lret = TreeView1.AddMenuItem("FILE", "&Delete File Output", ErrStr)
	lret = TreeView1.AddMenuItem("SUBJECT_BODY", "&Email Subject/Body Properties", ErrStr)
	lret = TreeView1.AddMenuItem("SUBJECT_BODY", "&Delete Email Subject/Body Templates", ErrStr)
	lret = TreeView1.AddMenuItem("PAGE", "&Visual Editor", ErrStr)
	lret = TreeView1.AddMenuItem("DEFINITION", "&Visual Editor", ErrStr)
	<% If MODE <> "RO" Then %>
	lret = TreeView1.AddMenuItem("TRANSMISSION", "&Copy This Transmission Sequence Step", ErrStr)
	lret = TreeView1.AddMenuItem("TRANSMISSION", "&Create Output Item", ErrStr)
	<% End If %>
	lret = TreeView1.AddMenuItem("TRANSMISSION", "&Transmission Properties", ErrStr)
	<% If MODE <> "RO" Then %>
	lret = TreeView1.AddMenuItem("TRANSMISSION", "-", ErrStr)
	lret = TreeView1.AddMenuItem("TRANSMISSION", "&Delete This Transmission Sequence Step", ErrStr)
	<% End If %>
	lret = TreeView1.AddMenuItem("ITEM", "&Output Item Properties", ErrStr)
	<% If MODE <> "RO" Then %>	
	lret = TreeView1.AddMenuItem("ITEM", "-", ErrStr)
	lret = TreeView1.AddMenuItem("ITEM", "&Delete Output Item", ErrStr)
	<% End If %>
	
	<% If Request.QueryString("EXPAND") <> "" Then %>
		TreeView1.ExpandNode("<%= Request.QueryString("EXPAND") %>")
	<% Else %>
		TreeView1.ExpandNode("RPID=<%= Cstr(Request.QueryString("ROUTING_PLAN_ID")) %>")
	<% End If %>
<% Else %>
	NodeX = TreeView1.AddNode ("",1 , "STEP=1", "NEWROOT", "New Routing Plan: ", "FOLDER", "FOLDERSEL")
<% End If %>
	
	document.all.StatusSpan.innerhtml = "<%= StrWarning %>"
	document.all.StatusSpan.style.color = "<%= WarnColor %>"
End Sub

Dim VeditorObj

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
	Select Case MenuItem
		Case "&New Transmission Sequence Step"
			lret = showModalDialog ("TransmissionSeqmodal.asp?STATUS=NEW&ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>"  , null, "dialogWidth:450px;dialogHeight:500px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
		Case "&Transmission Properties"
			lret = showModalDialog ("TransmissionSeqmodal.asp?STATUS=UPDATE&ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&" & NodeKey   , null, "dialogWidth:450px;dialogHeight:500px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=" & NodeKey
		Case "&Visual Editor"
			'If Not IsObject(VeditorObj) Then
			StatusSpan.innerHTML = "Visual Editor changes may not be reflected until refresh"
			StatusSpan.style.color = "maroon"
				Call LaunchODEditor(NodeKey)
				'Set VeditorObj = Window.open  ("OutputDefinitionEditor-f.asp?" & NodeKey, null, "height=500,width=750,status=no,toolbar=no,menubar=no,location=no")
			'Else
			'	VeditorObj.focus()
			'End If
		Case "&Delete This Transmission Sequence Step"
			If TreeView1.Nodes(TreeView1.selected).children < 1 Then 
				lret = msgbox ("Are you sure you want to delete this Transmission Sequence Step?", 1, "FNSDesigner")
				if lret = "1" Then
					Parent.Frames("hiddenpage").location.href = "SaveTransmission.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&STATUS=DELETE&" & NodeKey 
				End If
			Else
				MsgBox "Output Items must be removed before deleting the Transmission Sequence Step", 0, "FNSDesigner"
			End If
		Case "&Delete Output Item"
			lret = msgbox ("Are you sure you want to delete this Output Item?" & Vbcrlf & "The corresponding output definition will be detached from this routing plan.", 1, "FNSDesigner")
			if lret = "1" Then
				Parent.Frames("hiddenpage").location.href = "SaveTransmission.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&STATUS=DELETE_OUTPUT_ITEM&" & NodeKey 
			End If
		Case "&Create Output Item"
			lret = showModalDialog ("OutPutItemModal.asp?STATUS=NEW&ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&" & NodeKey   , null, "dialogWidth:450px;dialogHeight:450px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
		Case "&Copy This Transmission Sequence Step"
			ClipboardAgent.ClearAllProperties()
			lret = TreeView1.AddMenuItem("ROOT", "&Paste Transmission Sequence Step", ErrStr)
			ClipboardAgent.AddProperty "TRANSMISSION_SEQ_STEP_ID", NodeKey
			ClipboardAgent.SetPropertiesToClipboard
		Case "&Paste Transmission Sequence Step"
			ClipboardAgent.GetPropertiesFromClipboard
			'msgbox ClipboardAgent.GetProperty("TRANSMISSION_SEQ_STEP_ID")
			If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("TRANSMISSION_SEQ_STEP_ID") <> "" Then
				lret = MsgBox ("Are you sure you wish to paste a Transmission Sequence Step and " & VbCrlf & "along with it's attached output items, definitions, and pages ?", 1, "FNSDesigner")
				if lret = "1" Then
					Parent.Frames("hiddenpage").location.href = "SaveTransmission.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&STATUS=COPY&" & ClipboardAgent.GetProperty("TRANSMISSION_SEQ_STEP_ID")
				End If
			Else
				Msgbox "Nothing to paste", 0, "FNSDesigner"
			End If
		Case "&Output Item Properties"
			lret = showModalDialog ("OutPutItemModal.asp?STATUS=UPDATE&ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&" & NodeKey   , null, "dialogWidth:450px;dialogHeight:450px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=" & NodeKey
		Case "&Copy Routing Plan"
			ClipboardAgent.ClearAllProperties()
			ClipboardAgent.AddProperty "ROUTING_PLAN_ID", NodeKey
			ClipboardAgent.SetPropertiesToClipboard
		case "New &File Output"
			lret = showModalDialog ("FileOutputModal.asp?STATUS=NEW&" & NodeKey  , null, "dialogWidth:530px;dialogHeight:320px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
		case "&File Output Properties"
			lret = showModalDialog ("FileOutputModal.asp?STATUS=UPDATE&" & NodeKey, null, "dialogWidth:530px;dialogHeight:320px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
		case "&Delete File Output"
			lret = msgbox ("Are you sure you want to delete this File Output definition?" & Vbcrlf, 1, "FNSDesigner")
			if lret = "1" Then
				Parent.Frames("hiddenpage").location.href = "FileOutputSave.asp?STATUS=DELETE&" & NodeKey 
				self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
			End If
		case "Attach &Email Subject/Body Templates"
			lret = showModalDialog ("SubjBodyModal.asp?STATUS=NEW&" & NodeKey  , null, "dialogWidth:530px;dialogHeight:320px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
		case "&Email Subject/Body Properties"
			lret = showModalDialog ("SubjBodyModal.asp?STATUS=UPDATE&" & NodeKey, null, "dialogWidth:530px;dialogHeight:320px")
			self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
		case "&Delete Email Subject/Body Templates"
			lret = msgbox ("Are you sure you want to delete this Email Subject/Body definition?" & Vbcrlf, 1, "FNSDesigner")
			if lret = "1" Then
				Parent.Frames("hiddenpage").location.href = "SubjBodySave.asp?STATUS=DELETE&" & NodeKey 
				self.location.href = "RoutingPlanSummaryTree.asp?ROUTING_PLAN_ID=<%= Request.querystring("ROUTING_PLAN_ID") %>&EXPAND=STEP=1"
			End If
		Case Else
	End Select
End Sub

Function Handles(Obj, Title)
If InStr(1, top.frames("TOP").location.href, "Toppane.asp") <> 0 Then
	lret = top.frames("TOP").SetHandle(Obj, Title)
End If
End Function

Sub BtnRefresh_onclick
	self.location.href = "RoutingPlanSummaryTree.asp?<%= Request.QueryString %>"
End Sub
-->
</script>
<script LANGUAGE="JavaScript">
function LaunchODEditor(key) {
Url = "OutputDefinitionEditor-f.asp?AHSID=<%= Request.QueryString("AHSID")%>&" + key
var VisEditorObj = window.open(Url, null, "height=500,width=750,status=no,toolbar=no,menubar=no,location=no,resizable=yes,top=0,left=0");
lret = Handles( VisEditorObj, "OUTPUT");
VisEditorObj.focus()
}
</script>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" rightmargin="0" leftmargin="0" bottommargin="0">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<input TYPE="HIDDEN" NAME="WARNINGS">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Routing Plan Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8">
</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18" VALIGN="BOTTOM" NOWRAP>
<img SRC="..\IMAGES\Refresh.gif" STYLE="CURSOR:HAND" ALIGN="LEFT" ALT="Refresh" ID="BtnRefresh">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
:<nobr><span ID="StatusSpan" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<object CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0" 1>
	<param NAME="LPKPath" VALUE="LPKfilename.LPK">
</object>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="80%">
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
</body>
</html>
