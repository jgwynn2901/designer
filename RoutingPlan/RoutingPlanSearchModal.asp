<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\NavigateBack.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<%	IncludeNavigateBackJS(Request.QueryString("CONTAINERCONTEXT")) %>

<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim routing_plan_id, ahsid 
<% IF Request.QueryString("SELECTONLY") = "TRUE" THEN %>
		Dim RoutingPlanSearchObj
		Set RoutingPlanSearchObj = Window.dialogArguments
<%	END IF %>
	
Sub BtnSelect_onclick
	set TableObj = document.frames("TabFrame").document.frames("WORKAREA")
	<% If Request.QueryString("CONTAINERTYPE") <> "FRAMEWORK" Then %>
			RPID = TableObj.getmultipleindex( TableObj.Document.all.tblResult )
	<% Else 
			if Request.QueryString("SELECTONLY") = "TRUE" then %>
				RoutingPlanSearchObj.RPID = TableObj.getmultipleindex(TableObj.Document.all.tblResult)
				RoutingPlanSearchObj.RPDesc = TableObj.GetMultipleIndexDesc(TableObj.Document.all.tblResult)
				if RoutingPlanSearchObj.RPID = "" then
					MsgBox "No Routing Plan Is Selected."
					Exit Sub
				else
					RoutingPlanSearchObj.RPSelected = True
				End if
				window.close
				Exit Sub
		<%	else %>
				RPID = TableObj.getselectedindex( TableObj.Document.all.tblResult ) 
		<%	end if
		End If %>
	if RPID <> "" AND RPID <> "-1" then
		<% if IncludeNavigateBackInvoke(Request.QueryString("CONTAINERCONTEXT")) = true then %>
			Dim curURL
			curURL  = "../RoutingPlan/RoutingPlanSearchModal.asp?CONTAINERCONTEXT=DRILLIN&CONTAINERTYPE=FRAMEWORK"
			SetInfoForNavigateBack(curURL)
		<% end if %>
		<% If Request.QueryString("CONTAINERTYPE") ="FRAMEWORK" Then %>		

			document.frames("TabFrame").location.href = "RoutingPlanSummary-f.asp?ROUTING_PLAN_ID=" & RPID
			document.all.BtnSelect.disabled = true
		<% Else %>
			window.dialogArguments.multiselected = RPID
			window.close
		<% End If %>
	Else
		msgbox "Please select a Routing Plan"
	End If
End Sub

Sub BtnCancel_onclick
	window.close
End Sub

Sub window_onload
<% If Request.QueryString("CONTAINERTYPE") <> "FRAMEWORK" Then %>
	routing_plan_id = window.dialogArguments.routing_plan_id
	ahsid = window.dialogArguments.ahsid
	TabFrame.document.location = "RoutingPlanSearch-f.asp?CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>&ahsid=<%= Request.QueryString("AHSID") %>"
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

-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=0 topmargin=0 bottommargin=0 rightmargin=0 BGCOLOR='<%=BODYBGCOLOR%>' >
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="100%" HEIGHT="90%" SRC="
<% If Request.QueryString("CONTAINERTYPE") = "FRAMEWORK" Then %>
RoutingPlanSearch-f.asp?CONTAINERTYPE=<%= Request.QueryString("CONTAINERTYPE") %>
<% End If %>">
</iframe><BR>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnSelect>Select</BUTTON></TD>
<% If Request.QueryString("CONTAINERTYPE") <> "FRAMEWORK" Then %>
	<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnSelectAll>Highlight All</BUTTON></TD>
	<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnCancel>Cancel</BUTTON></TD>
<% Elseif Request.QueryString("CONTAINERTYPE") = "FRAMEWORK" AND Request.QueryString("SELECTONLY") = "TRUE" Then %>
	<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME=BtnCancel>Cancel</BUTTON></TD>
<% End If %>
</TR>
</BODY>
</HTML>
