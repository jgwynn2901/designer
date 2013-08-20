<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\NavigateBack.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "CallFlowSearchModal.asp"
End If
If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"

%>
<% IncludeNavigateBackJS(Request.QueryString("CONTAINERCONTEXT")) %>

<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim routing_plan_id, ahsid 

Sub BtnSelect_onclick
	set TableObj = document.frames("TabFrame").document.frames("WORKAREA")
	<% If Request.QueryString("CONTAINERTYPE") <> "FRAMEWORK" AND Request.QueryString("LAUNCHER") <> "SEARCH" Then %>
		RPID = TableObj.getmultipleindex( TableObj.Document.all.tblResult )
	<% Else %>
		RPID = TableObj.getselectedindex( TableObj.Document.all.tblResult ) 
	<% End If %>
	if RPID <> "" AND RPID <> "-0.1" then
		<% if IncludeNavigateBackInvoke(Request.QueryString("CONTAINERCONTEXT")) = true then %>
			Dim curURL
			curURL  = "../CallFlow/CallFlowSearchModal.asp?CONTAINERCONTEXT=DRILLIN&CONTAINERTYPE=FRAMEWORK"
			SetInfoForNavigateBack(curURL)
		<% end if %>
		<% If Request.QueryString("CONTAINERTYPE") ="FRAMEWORK" Then %>	
			self.location.href = "CallFlowMaintenanceEditor.asp?CALLFLOW_ID=" & RPID
		<% Else %>
			window.dialogArguments.multiselected = RPID
			window.close
		<% End If %>
	Else
		msgbox "Please Select a Call Flow !", 0, "FNSDesigner"
	End If
End Sub

Sub BtnCancel_onclick
	window.close
End Sub

Sub window_onload
<% If Request.QueryString("CONTAINERTYPE") <> "FRAMEWORK" Then %>
	routing_plan_id = window.dialogArguments.routing_plan_id
	ahsid = window.dialogArguments.ahsid
	TabFrame.document.location = "CallFlowSearch-f.asp?CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>&ahsid=<%= Request.QueryString("AHSID") %>"
<% End If %>
End Sub

Sub BtnSelectAll_onclick
<% If Request.QueryString("CONTAINERTYPE") <> "FRAMEWORK" Then %>
set TableObj = document.frames("TabFrame").document.frames("WORKAREA")
For i = 1 to TableObj.Document.all.tblResult.rows.length-1 step 1
	TableObj.Document.all.tblResult.rows(i).classname = "ResultSelectRow"
Next
<% End If %>
End Sub

Sub BtnNew_onclick
<% If Request.QueryString("AHSID") <> "" Then %>
	self.location.href = "CallFlowEditor.asp?CALLFLOW_ID=NEW"
<% Else %>
	self.location.href = "CallFlowMaintenanceEditor.asp?CALLFLOW_ID=NEW"
<% End If %>
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=0 topmargin=0 bottommargin=0 rightmargin=0 BGCOLOR='<%=BODYBGCOLOR%>' >
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="100%" HEIGHT="90%" SRC="
<% If Request.QueryString("CONTAINERTYPE") = "FRAMEWORK" Then %>
CallFlowSearch-f.asp?CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>
<% End If %>">
</iframe><BR>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnSelect>Select</BUTTON></TD>
<% If Request.QueryString("CONTAINERTYPE") <> "MODAL" Then %>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON <% If MODE="RO" Then Response.Write(" DISABLED ") %> NAME=BtnNew>New</BUTTON></TD>
<% End If %>
<% If Request.QueryString("LAUNCHER") <> "SEARCH" Then 
If Request.QueryString("CONTAINERTYPE") = "MODAL" Then %>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnSelectAll>Highlight All</BUTTON></TD>
<%
End If
 End If %>
<% If Request.QueryString("CONTAINERTYPE") = "MODAL" Then %>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnCancel>Cancel</BUTTON></TD>
<% End If %>
</TR>
</BODY>
</HTML>
