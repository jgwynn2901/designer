<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\CheckSharedCallFlow.inc"-->
<!--#include file="..\lib\RefCountRptinc.asp"-->
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"--> 
<!--#include file="..\lib\CommonError.inc"-->
<%
SharedCallFlowCount = 0
Response.Expires=0
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "CallFlowEditor.asp"
End If
If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"
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
	LogStatusGroupBegin
	LogStatus S_ERROR, "Paste Frame operation failed", "FRAME_ORDER", "", 0, ""
	LogStatusGroupEnd
end if
If Request.QueryString("REASSIGN") = "TRUE" Then
	SQL = ""
	SQL = SQL & "UPDATE ACCOUNT_CALLFLOW SET CALLFLOW_ID=" & Request.QueryString("RECALLFLOW_ID")
	SQL = SQL & " WHERE ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID")
	Set RSReassign = Conn.Execute(SQL)
	Response.Redirect "CallFlowEditor.asp?ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID") & "&AHSID=" & Request.QueryString("AHSID") & "&CALLFLOW_ID=" & Request.QueryString("CFID")
End If

If Request.QueryString("COPYASNEW") <> "" Then
	SQL = ""
	SQL = "{call Designer.CopyFrame(" &  Request.QueryString("FRAMEID") & " ,{resultset 1, outFrameId,StatusMsg,StatusNum})}"
	Set RSCopy = Conn.Execute(SQL)
	If RSCopy("StatusNum") <> "0" Then
		errmsg = RSCopy("StatusMsg")
	Else
		SQL = ""
		SQL = SQL & "{call Designer.CopyFrameOrder(" &  Request.QueryString("FRAMEID") & ", "
		SQL = SQL & Request.QueryString("CFID") & "," 
		SQL = SQL & RSCopy("outFrameId") & ","
		SQL = SQL & Request.QueryString("CFID") & ",{resultset 1, StatusMsg, StatusNum})}"
		Set RS=Conn.Execute(SQL)
		SQLDelete = ""
		SQLDelete = SQLDelete & "DELETE FROM FRAME_ORDER WHERE CALLFLOW_ID=" & Request.QueryString("CFID") & " AND "
		SQLDelete = SQLDelete & "FRAME_ID=" & Request.QueryString("FRAMEID")
		Set RSDel = Conn.Execute(SQLDelete)
	End If
	Response.Redirect "CallFlowEditor.asp?ERRORMSG=" & errmsg & "&ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID") & "&AHSID=" & Request.QueryString("AHSID") & "&CALLFLOW_ID=" & Request.QueryString("CFID")
End If

If Request.QueryString("DETACH") <> "" Then
	SQLDelete = ""
	SQLDelete = SQLDelete & "DELETE FROM FRAME_ORDER WHERE CALLFLOW_ID=" & Request.QueryString("CFID") & " AND "
	SQLDelete = SQLDelete & "FRAME_ID=" & Request.QueryString("FRAMEID")
	Set RSDel = Conn.Execute(SQLDelete)
	Response.Redirect "CallFlowEditor.asp?ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID") & "&AHSID=" & Request.QueryString("AHSID") & "&CALLFLOW_ID=" & Request.QueryString("CFID")
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
	SQL2 = SQL2 & Request.QueryString("CFID") 
	SQL2 = SQL2 & ",0"
	SQL2 = SQL2 & ",'-999999999', '-999999999', -999999999, -999999999, "
	SQL2 = SQL2 & "'U', -999999999, -999999999, '-999999999', "
	SQL2 = SQL2 & "'-999999999', '-999999999', '-999999999', '-999999999', '-999999999', '-999999999', "
	SQL2 = SQL2 & "-999999999, 'U' "
	SQL2 = SQL2 & ")" 
	Set RS=Conn.Execute(SQL2)
	Response.Redirect "CallFlowEditor.asp?ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID") & "&AHSID=" & Request.QueryString("AHSID") & "&CALLFLOW_ID=" & Request.QueryString("CFID")
End If

If Request.QueryString("COPY") <> "" Then
	SQL = ""
	SQL = SQL & "{call Designer.CopyFrameOrder(" &  Request.QueryString("FRAMEID") & ", "
	SQL = SQL & Request.QueryString("FROM_CALLFLOW_ID") & "," 
	SQL = SQL & Request.QueryString("FRAMEID") & ","
	SQL = SQL & Request.QueryString("CALLFLOW_ID") & ",{resultset 1, StatusMsg, StatusNum})}"
	Set RS=Conn.Execute(SQL)
	if cint(RS.Fields("StatusNum")) <> 0 then
		cErr = RS.Fields("StatusMsg")
		Response.Redirect "CallFlowEditor.asp?COPYERR=Y&CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID") & "&ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID")
	else
		Response.Redirect "CallFlowEditor.asp?CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID") & "&ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID")
	End If
End If

IF Request.QueryString("ACTION") = "SAVE" AND  Request.QueryString("CALLFLOW_ID") = "NEW" THEN
	ON ERROR RESUME NEXT
	SQLInsert = ""
	SQLInsert = SQLInsert & "{call Designer.AddORUPDATECALLFLOWANDACCTCF(NULL, 0, '" & SwapFlg(Request.Form("CAT_FLG")) & "', "
	SQLInsert = SQLInsert & "'" & REPLACE(Replace(Request.Form("NAME"), """", """"), "'", "''") & "', "
	SQLInsert = SQLInsert & "'" & Replace(Replace(Request.Form("DESCRIPTION"), """", """"), "'", "''" ) & "', "
	SQLInsert = SQLInsert & Request.Form("ACCNT_HRCY_STEP_ID") & ", '" & Request.Form("LOB_CD") & "', "
'***************************************
'	CALLFLOW_TYPE FNSNET IS HARDCODED (for now)
'***************************************
	
	SQLInsert = SQLInsert & "'" & SwapFlg(Request.Form("CALL_START_FLG")) & "', " & Swap(Request.Form("VALIDRULE_ID")) & ",'FNSNET'"
	SQLInsert = SQLInsert & ", {resultset 1, outCF_ID, outACF_ACF_ID, StatusMsg, StatusNum})}"
	Set RS=Conn.Execute(SQLInsert)
	IF CStr(RS("StatusNum")) <> "0" THEN
		s_InsertErrorMsg = Server.URLEncode(Mid(RS.FIELDS("StatusMsg"), (Instr(1,RS.FIELDS("StatusMsg"), ":") +2)))
		Response.Redirect "CallFlowEditor.asp?STATUS=" & s_InsertErrorMsg & "&CALLFLOW_ID=NEW&AHSID=" & Request.QueryString("AHSID")
	ELSE
		Response.Redirect "CallFlowEditor.asp?STATUS=SAVED&CALLFLOW_ID=" & RS.FIELDS("outCF_ID") & "&ACCOUNTCALLFLOW_ID=" & RS.FIELDS("outACF_ACF_ID") & "&AHSID=" & Request.QueryString("AHSID")
	END IF
	RS.CLOSE
	SET RS = NOTHING
END IF

If Request.QueryString("ACTION") = "SAVE" AND  Request.QueryString("CALLFLOW_ID") <> "NEW" Then
	ON ERROR RESUME NEXT
	SQLUpdate = ""
	SQLUpdate = SQLUpdate & "{call Designer.AddORUPDATECALLFLOWANDACCTCF(" & Request.QueryString("CALLFLOW_ID") & ", 0, '" & SwapFlg(Request.Form("CAT_FLG")) & "', "
	SQLUpdate = SQLUpdate & "'" & REPLACE(Replace(Request.Form("NAME"), """", """"), "'", "''") & "', "
	SQLUpdate = SQLUpdate & "'" & Replace(Replace(Request.Form("DESCRIPTION"), """", """"), "'", "''" ) & "', "
	SQLUpdate = SQLUpdate & Request.Form("ACCNT_HRCY_STEP_ID") & ", '" & Request.Form("LOB_CD") & "', "
	SQLUpdate = SQLUpdate & "'" & SwapFlg(Request.Form("CALL_START_FLG")) & "', " & Swap(Request.Form("VALIDRULE_ID")) & ",'FNSNET'"
	SQLUpdate = SQLUpdate & ", {resultset 1, outCF_ID, outACF_ACF_ID, StatusMsg, StatusNum})}"
	Set RS=Conn.Execute(SQLUpdate)
	IF CStr(RS("StatusNum")) <> "0" THEN
		s_UpdateErrorMsg = Server.URLEncode(Mid(RS.FIELDS("StatusMsg"), (Instr(1,RS.FIELDS("StatusMsg"), ":") +2)))
		Response.Redirect "CallFlowEditor.asp?STATUS=" & s_UpdateErrorMsg & "&CALLFLOW_ID=" & Request.QueryString("CALLFLOW_ID") & "&ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID") & "&AHSID=" & Request.QueryString("AHSID")
	ELSE
		Response.Redirect "CallFlowEditor.asp?STATUS=SAVED&CALLFLOW_ID=" & RS.FIELDS("outCF_ID") & "&ACCOUNTCALLFLOW_ID=" & Request.QueryString("ACCOUNTCALLFLOW_ID") & "&AHSID=" & Request.QueryString("AHSID")
	END IF
	RS.CLOSE
	SET RS = NOTHING
END IF
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Call Flow Editor</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->

Sub window_onload
	ClipboardAgent.GetPropertiesFromClipboard
<%
If Request.QueryString("CALLFLOW_ID") <> "NEW" Then
	StrSQL2 = ""
	StrSQL2 = StrSQL2 & "SELECT ACCOUNT_CALLFLOW.*, RULES.RULE_TEXT FROM ACCOUNT_CALLFLOW, RULES WHERE ACCOUNT_CALLFLOW.ACCOUNTCALLFLOW_ID =" & Request.QueryString("ACCOUNTCALLFLOW_ID") & " AND "
	StrSQL2 = StrSQL2 & " ACCOUNT_CALLFLOW.VALIDRULE_ID = RULES.RULE_ID(+)"
	Set RS2 = Conn.Execute(StrSQL2)
	StrSQL3 = ""
	StrSQL3 = StrSQL3 & "SELECT * FROM CALLFLOW WHERE CALLFLOW_ID=" & RS2("CALLFLOW_ID")
	Set RS3 = Conn.Execute(StrSQL3)
	VALIDRULE_ID = Trim(RS2("VALIDRULE_ID"))
	If Not IsNull(RS2("RULE_TEXT")) Then
		VALIDRULE_TEXT = RS2("RULE_TEXT")
		FULL_VALIDRULE_TEXT = RS2("RULE_TEXT") 
	Else
		VALIDRULE_TEXT = ""
		FULL_VALIDRULE_TEXT = ""
	End If
	If Len(VALIDRULE_TEXT) > 50 Then
		VALIDRULE_TEXT = Mid(VALIDRULE_TEXT, 1, 50) & "..."
	End If
	LOB_CD = RS2("LOB_CD")
	CALL_START_FLG = RS2("CALL_START_FLG")
	NAME = Replace(RS3("NAME"), """", "&quot;")  ''NAME = RS3("NAME")
	DESCRIPTION = Replace(RS3("DESCRIPTION"), """", "&quot;") ''DESCRIPTION = RS3("DESCRIPTION")
	CAT_FLG = RS3("CAT_FLG")
	ACCNT_HRCY_STEP_ID = RS2("ACCNT_HRCY_STEP_ID")
	ACCOUNTCALLFLOW_ID = RS2("ACCOUNTCALLFLOW_ID")
	CALLFLOW_ID= RS3("CALLFLOW_ID")
	If Request.QueryString("CALLFLOW_ID") <> "NEW" Then
		SharedCallFlowCount = CheckSHaredCallFlow(CALLFLOW_ID,True, True, 2, False, False, 0) 
	End If
	StrSQL = ""
	StrSQL = StrSQL & "SELECT FRAME_ORDER.SEQUENCE, FRAME_ORDER.FRAME_ID, FRAME.NAME FROM FRAME_ORDER, FRAME WHERE CALLFLOW_ID=" & RS3("CALLFLOW_ID") & " AND "	
	StrSQL = StrSQL & "FRAME.FRAME_ID = FRAME_ORDER.FRAME_ID ORDER BY FRAME_ORDER.SEQUENCE"
	Set RS = Conn.Execute(StrSQL)
%>
	NodeX = TreeView1.AddNode ("",1 , "CFID=<%= CALLFLOW_ID %>", "CALLFLOW", "Call Flow: <%= Replace(RS3("NAME"),"""","""""")%>, <%= RS3("CALLFLOW_ID") %>", "FOLDER", "FOLDERSEL")	
	<% Do While Not RS.EOF %>
	NodeX = TreeView1.AddNode ("CFID=<%= CALLFLOW_ID %>", 4, "FRAMEID=<%= RS("FRAME_ID") %>" , "FRAME", "(<%= RS("SEQUENCE") %>) Frame: <%= RS("NAME") %>, <%= RS("FRAME_ID") %>", "FOLDER", "FOLDERSEL")
<%
	RS.movenext
	loop
	RS.close
%>
	lret = TreeView1.AddMenuItem("FRAME", "&Visual Editor", ErrStr)
	lret = TreeView1.AddMenuItem("CALLFLOW", "&Visual Editor", ErrStr)

	<% If MODE <> "RO" Then %>
	lret = TreeView1.AddMenuItem("FRAME", "-", ErrStr)
	lret = TreeView1.AddMenuItem("FRAME", "&Copy As New", ErrStr)
	lret = TreeView1.AddMenuItem("FRAME", "&Attach Frame", ErrStr)
	lret = TreeView1.AddMenuItem("CALLFLOW", "&Attach Frame", ErrStr)
	lret = TreeView1.AddMenuItem("FRAME", "Copy &Frame", ErrStr)
	lret = TreeView1.AddMenuItem("FRAME", "&Detach Frame", ErrStr)
	<% End If %>
	If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("FRAME_ID") <> "" and "<%=Request.QueryString("CALLFLOW_ID")%>" <> "NEW" Then
		lret = TreeView1.AddMenuItem("CALLFLOW", "&Paste Frame", ErrStr)
		lret = TreeView1.AddMenuItem("FRAME", "&Paste Frame", ErrStr)
	End If
	document.all.VALIDRULE_ID_TEXT.innerhtml = "<%= Replace(VALIDRULE_TEXT, """", "&quot;") %>"
	VALIDRULE_ID_TEXT.title = "<%= Replace(FULL_VALIDRULE_TEXT, """", """""")  %>"
	document.all.LOB_CD.value = "<%= LOB_CD %>"
	<% If CALL_START_FLG = "Y" Then %>
		document.all.CALL_START_FLG.checked = true
	<% End If %>
	<% If CAT_FLG = "Y" Then %>
		document.all.CAT_FLG.checked = true
	<% End If %>
<% 
Else
%>
	NodeX = TreeView1.AddNode ("",1 , "CFID=<%= CALLFLOW_ID %>", "FRAME", "New Call Flow", "FOLDER", "FOLDERSEL")	
	<% If Request.QueryString("AHSID") <> "" Then %>
		document.all.accnt_hrcy_step_id.value = "<%= Request.QueryString("AHSID") %>"
	<%Else%>
		document.all.accnt_hrcy_step_id.value = "1"
	<% End iF %>
<% 
End If
%>	

<% If Request.QueryString("STATUS") = "SAVED" Then %>
	SpanStatus2.innerHTML = "Saved"
<% Elseif instr(1, Request.QueryString("STATUS"), "FNSOWNER.") Then %>
	SpanStatus.innerHTML = "Save Error: "
	SpanStatus2.innerHTML = Replace("<%= Request.QueryString("STATUS") %>", "FNSOWNER.", "")
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

If Request.QueryString("COPYERR") <> "" then
%>
	SpanTreeStatus.innerhtml = "<Font color='RED'>Error: </Font>Paste Frame operation failed."
<%end if%>

	TreeView1.ExpandNode("CFID=<%= CALLFLOW_ID %>")
<% If Request.QueryString("ERRORMSG") <> "" Then %>
		SpanTreeStatus.innerHTML = "<%= Request.QueryString("ERRORMSG") %>"
<% End If %>	
End Sub


Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
Select Case MenuItem
	Case "&Visual Editor"
		SpanTreeStatus.innerHTML = "Visual Editor changes may not be reflected until refresh"
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
		ClipboardAgent.AddProperty "FRAME_TEXT", NodeText
		ClipboardAgent.AddProperty "FROM_CALLFLOW_ID", "<%= CALLFLOW_ID %>"
		ClipboardAgent.SetPropertiesToClipboard
		
	Case "&Paste Frame"
		ClipboardAgent.GetPropertiesFromClipboard
		If ClipboardAgent.IsClipboardDataAvailable = True AND ClipboardAgent.GetProperty("FRAME_ORDER_DATA") <> "" Then
			lret = MsgBox ("Are you sure you wish to paste frame """ & ClipboardAgent.GetProperty("FRAME_TEXT") & """ in this call flow", 1, "FNSDesigner")
			if lret = "1" Then
				self.location.href = "CallFlowEditor.asp?COPY=TRUE&" & ClipboardAgent.GetProperty("FRAME_ID") & "&FROM_CALLFLOW_ID=" & ClipboardAgent.GetProperty("FROM_CALLFLOW_ID") & "&CALLFLOW_ID=<%=Request.QueryString("CALLFLOW_ID")%>&ACCOUNTCALLFLOW_ID=<%=Request.QueryString("ACCOUNTCALLFLOW_ID")%>"
				'self.location.href = "CallFlowEditor.asp?COPY=TRUE&" & "FROM_CALLFLOW_ID=" & ClipboardAgent.GetProperty("FROM_CALLFLOW_ID")  & "&" & ClipboardAgent.GetProperty("FRAME_ID") & "&AHSID=<%= Request.QueryString("AHSID") %>&CFID=<%= CALLFLOW_ID %>&ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID %>"
			End If
		Else
			Msgbox "Nothing to paste", 0, "FNSDesigner"
		End If
	Case "&Attach Frame"
		strURL = "FrameSearchModal.asp"
		showModalDialog  strURL  ,FrameObj ,"dialogWidth:450px;dialogHeight:450px;center"
		If FrameObj.FrameID <> "" Then
			self.location.href = "CallFlowEditor.asp?ATTACH=TRUE&FRAMEID=" & FrameObj.FrameID  & "&CFID=<%= CALLFLOW_ID %>&AHSID=<%= Request.QueryString("AHSID") %>&ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID %>"
		End If
	Case "&Detach Frame"
			lret = msgbox("Are you sure you want to detach this frame? """ & NodeText & """", 1, "FNSDesigner")
			if lret = "1" Then
				self.location.href = "CallFlowEditor.asp?DETACH=TRUE&" & NodeKey & "&AHSID=<%= Request.QueryString("AHSID") %>&CFID=<%= CALLFLOW_ID %>&ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID %>"
			End If
	Case "&Copy As New"
		lret = msgbox ("Are you sure you want to copy this frame: " & NodeText & Chr(13) & "Copying this frame will create a new unique instance of this frame" & VbCrlf & "and the current frame will be detached.", 1, "FNSNet")
		if lret = "1" Then
			self.location.href = "CallFlowEditor.asp?COPYASNEW=TRUE&" & NodeKey & "&AHSID=<%= Request.QueryString("AHSID") %>&CFID=<%= CALLFLOW_ID %>&ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID%>"
		End If
End Select
End Sub

Function Handles(Obj, Title)
	If InStr(1, top.frames("TOP").location.href, "Toppane.asp") <> 0 Then
		lret = top.frames("TOP").SetHandle(Obj, Title)
	End If
End Function

Function AttachRule (ID, SPANID)
MODE = document.body.getAttribute("ScreenMode")
RID = ID.value
RuleSearchObj.RID = RID
RuleSearchObj.RIDText = SPANID.innerhtml
RuleSearchObj.Selected = false

If RID = "" Then RID = "NEW"

	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If

strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID
showModalDialog  strURL  ,RuleSearchObj ,"dialogWidth:450px;dialogHeight:450px;center"
	
If RuleSearchObj.Selected = true Then
	If RuleSearchObj.RID <> ID.value then
		ID.value = RuleSearchObj.RID
	end if
	If len(RuleSearchObj.RIDText) > 50 Then
		SPANID.innerhtml = Mid(RuleSearchObj.RIDText, 1, 50) & "..."
	Else
		SPANID.innerhtml = RuleSearchObj.RIDText
	End If
	SPANID.Title = RuleSearchObj.RIDText
ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
	SPANID.innerhtml = RuleSearchObj.RIDText
	SPANID.Title = RuleSearchObj.RIDText
End If
End Function

Function DetachRule(ID, SPANID)
<% If MODE = "RO" Then Response.write(" Exit Function ") %>
	ID.value = ""
	SPANID.innerhtml = ""
	SPANID.Title = ""
End Function

Sub BtnNew_onclick
	self.location.href = "CallFlowEditor.asp?CALLFLOW_ID=NEW&ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID %>&AHSID=<%= Request.QueryString("AHSID") %>"
End Sub

Sub BtnClear_onclick
	document.all.NAME.value = ""
	document.all.DESCRIPTION.value = ""
	document.all.VALIDRULE_ID.value = ""
	document.all.VALIDRULE_ID_TEXT.innerhtml = ""
	document.all.LOB_CD.value = ""
	document.all.CALL_START_FLG.checked = false
	document.all.CAT_FLG.checked = false
End Sub

Sub BtnSave_onclick
	StrError = ""
	If document.all.NAME.value = "" Then
		StrError = StrError & "Name is required field" & VbCrlf
	End If
	If document.all.DESCRIPTION.value = "" Then
		StrError = StrError & "Description is a required field" & VbCrlf
	End If
	If document.all.LOB_CD.value = "" Then
		StrError = StrError & "LOB is a required field" & VbCrlf
	End If
	If StrError = "" Then
		FrmSave.Submit()
	Else
		MsgBox StrError, 0 , "FNSDesigner"
	End If
End Sub

Sub BtnRefresh_onclick
	self.location.href = "CallFlowEditor.asp?CALLFLOW_ID=<%= Request.QueryString("CALLFLOW_ID") %>&AHSID=<%= Request.QueryString("AHSID") %>&ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID %>"
End Sub

Sub BtnGrfxBack_Onclick()
<% If Request.QueryString("AHSID") <> "" Then %>
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&DROPDOWN=CALLFLOW"
<% Else %>
	self.location.href = "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=FRAMEWORK"
<% End If %>
End Sub

Sub BtnAttachCF_onclick
<% If MODE="RO" Then Response.write(" Exit Sub ") %>
	lret = ""
	strURL = ""
	SearchObj.multiselected = ""
	strURL = "../CallFlow/CallFlowSearchModal.asp?CONTAINERTYPE=MODAL&LAUNCHER=SEARCH"
	lret = window.showModalDialog(strURL  ,SearchObj ,"dialogWidth:625px;dialogHeight:550px;center")
	if SearchObj.multiselected <> "" Then
		self.location.href = "CallFlowEditor.asp?REASSIGN=TRUE&ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID %>&AHSID=<%= Request.QueryString("AHSID") %>&RECALLFLOW_ID=" & SearchObj.multiselected
	End If
End Sub

Sub StatusRpt_onclick
	If CLng(<%=SharedCallFlowCount%>) > 1 Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other details reported", 0 , "FNSNet"
	End If
End Sub

Sub RefCountRpt_onclick
<% If Request.QueryString("CALLFLOW_ID") <> "NEW" Then %>
	lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedCallFLow=True&ID=<%= CALLFLOW_ID %>", Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
<% End If %>
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
function CRPSearchObj()
{
	this.routing_plan_id = "";
	this.ahsid = "";
	this.multiselected = "";
}
var SearchObj = new CRPSearchObj();
var FrameObj = new CFrameSearchObj();
var RuleSearchObj = new CRuleSearchObj();
</script>
</head>
<body bgcolor="<%= BODYBGCOLOR %>" topmargin="0" rightmargin="0" leftmargin="0" bottommargin="0" ScreenMode="<%= MODE %>">
<!--#include file="..\lib\NavBack.inc"-->
<form NAME="FrmSave" ACTION="CallFlowEditor.asp?ACTION=SAVE&amp;ACCOUNTCALLFLOW_ID=<%= ACCOUNTCALLFLOW_ID %>&amp;CALLFLOW_ID=<%= Request.QueryString("CALLFLOW_ID") %>&amp;AHSID=<%= Request.QueryString("AHSID") %>" METHOD="POST">
<table WIDTH="98%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Call Flow
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Edit Call Flow.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<td WIDTH="14">
<img ID="BtnAttachCF" SRC="..\images\Attach.gif" STYLE="CURSOR:HAND" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Re-Assign Call Flow">
</td>
<td WIDTH="14">
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
<td CLASS="LABEL">Name:<br><input TYPE="TEXT" SIZE="40" <% If MODE = "RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> CLASS="LABEL" NAME="NAME" VALUE="<%= NAME %>"></td>
<td CLASS="LABEL" VALIGN="BOTTOM"><input TYPE="CHECKBOX" id="CAT_FLG" <% If MODE = "RO" Then Response.write(" DISABLED ") %> name="CAT_FLG">Catastrophe:</td>
</tr>
<tr>
<td CLASS="LABEL" COLSPAN="2">Description:<br><input TYPE="TEXT" SIZE="70" <% If MODE = "RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER'  ") %> CLASS="LABEL" id="Description" name="Description" VALUE="<%= DESCRIPTION %>"></td>
</tr>
<tr>
</table>
</td><td VALIGN="TOP" ALIGN="RIGHT">
<table>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE = "RO" Then Response.write(" DISABLED ") %> NAME="BtnSave" ACCESSKEY="S"><u>S</u>ave</button>
</tr>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE = "RO" Then Response.write(" DISABLED ") %> <% If Request.QueryString("CALLFLOW_ID") = "NEW" Then Response.write(" DISABLED ") %> NAME="BtnNew" ACCESSKEY="N"><u>N</u>ew</button>
</tr>
<tr>
<td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE = "RO" Then Response.write(" DISABLED ") %> NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button>
</tr>
</table>
</td></tr></table>
<table WIDTH="98%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Account Call Flow
&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Edit Call Flow.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table cellspacing="1" cellpadding="0">
<tr>
<td CLASS="LABEL">LOB:<br>
<select NAME="LOB_CD" CLASS="LABEL" <% If MODE = "RO" Then Response.write(" DISABLED ") %>>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "SELECT * FROM LOB WHERE LOB_CD IS NOT NULL"
	Set RS4 = Conn.Execute(SQLST)
Do While Not RS4.EOF
%>
<option VALUE="<%= RS4("LOB_CD") %>"><%= RS4("LOB_NAME") %>
<%
RS4.MoveNext
Loop
RS4.CLose
%>
</select></td>
<td CLASS="LABEL" ALIGN="LEFT" VALIGN="MIDDLE">
A.H. Step ID:<br><input READONLY TYPE="TEXT" CLASS="LABEL" NAME="ACCNT_HRCY_STEP_ID" STYLE="BACKGROUND-COLOR:SILVER" VALUE="<%= ACCNT_HRCY_STEP_ID %>" SIZE="10">
</td>
<!--<TD valign=bottom ALIGN=LEFT><IMG src="..\Images\attach.GIF" TITLE="Attach Account Hierarchy Step" STYLE="CURSOR:HAND" align=absbottom></TD>-->
<td CLASS="LABEL" VALIGN="BOTTOM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input TYPE="CHECKBOX" id="CALL_START_FLG" <% If MODE = "RO" Then Response.write(" DISABLED ") %> name="CALL_START_FLG">Call Start Flag:</td>
</tr>
</table>
<table>
<tr>
<td CLASS="LABEL" NOWRAP>
<img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEnabledRule" TITLE="Attach Rule" OnClick="AttachRule VALIDRULE_ID, VALIDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
<img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEnabledRule" TITLE="Detach Rule" OnClick="DetachRule VALIDRULE_ID, VALIDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
Valid Rule: 
<span ID="VALIDRULE_ID_TEXT" CLASS="LABEL" TITLE="<%= FULL_VALIDRULE_TEXT %>">
</span><input TYPE="HIDDEN" NAME="VALIDRULE_ID" VALUE="<%= Trim(VALIDRULE_ID) %>">
</td>
</tr></table>
</form>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10">&nbsp;» Call Flow Frames
&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Edit Call Flow.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<object ID="TreeView1" <%GetTreeCLSID()%> Width="100%" Height="47%">
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
