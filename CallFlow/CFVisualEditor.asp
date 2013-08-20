<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\CheckSharedFrame.inc"-->
<!--#include file="..\lib\CheckSharedCallFlow.inc"-->
<!--#include file="..\lib\RefCountRptinc.asp"-->
<!--#include file="..\lib\Security.inc"-->
<%
Response.Expires = 0
If HasViewPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then   
	Session("NAME") = ""
	Response.Redirect "CF_VisualEditor.asp"
End If
If HasModifyPrivilege("FNSD_CALLFLOW",SECURITYPRIV) <> True Then MODE = "RO"

If IsNull(Session("CallFlowRS")) Then
	Session("ErrorMessage") = "Session has Expired"
	Response.Redirect "..\directerror.asp"
End If

If (Request.QueryString("CFID") = "" OR  Not IsNumeric(Request.QueryString("CFID"))) AND Request.QueryString("NAVIGATE") = "" Then
		Session("ErrorMessage") = "On Page CFVisualEditor QueryStrings CFID and NAVIGATE were not numeric or null"
		Response.redirect	 "..\directerror.asp"
End If
 
If Request.QueryString("CFID") <> "" Then 
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST2 = ""
	Session("FrameCount") = ""
	
	If Session("RSBookMark") = "" Then
		Session("RSBookMark") = 0
	End If
	If Request.QueryString("NAVIGATE") = "" Then
		Session("RSBookMark") = 1
	End If
	SQLST = SQLST & "SELECT CALLFLOW.NAME As CFNAME, FRAME_ORDER.FRAME_ID, FRAME_ORDER.CALLFLOW_ID, FRAME_ORDER.SEQUENCE, FRAME.NAME FROM CALLFLOW, FRAME_ORDER,	FRAME WHERE FRAME_ORDER.CALLFLOW_ID =" & Request.QueryString("CFID") & " AND "
	SQLST = SQLST & "CALLFLOW.CALLFLOW_ID = FRAME_ORDER.CALLFLOW_ID AND FRAME_ORDER.FRAME_ID = FRAME.FRAME_ID ORDER BY SEQUENCE"
	
	SQLST2 = SQLST2 & "SELECT COUNT(*) As FrameCount FROM FRAME_ORDER,	FRAME WHERE FRAME_ORDER.CALLFLOW_ID =" & Request.QueryString("CFID") & "AND FRAME_ORDER.FRAME_ID = FRAME.FRAME_ID  ORDER BY SEQUENCE"

	Set RS = Conn.Execute(SQLST)
	Set RSList = Conn.Execute(SQLST)
	Set RS2 = Conn.Execute(SQLST2)
	If Clng(RS2("FrameCount")) =0 Then
		Response.Redirect "NewFrame.asp?CFID=" & Request.QueryString("CFID") & "&NEW=TRUE"
	End If
	Session("CallFlowName") = RS("CFNAME")
	If RS.EOF or isnull(RS) Then
		Session("ErrorMessage") = "Statement = " & SQLST & " ----- returned no records" & vbCrlf
		Response.redirect	 "..\directerror.asp"
	Else
		Session("FrameCount") = RS2("FrameCount")
		FRAMEID = RS("FRAME_ID")
	End If	
	
	If Request.QueryString("FRAMEID") <> "" Then
		Do While Clng(Request.QueryString("FRAMEID")) <> Clng(RS("FRAME_ID"))
			RS.MoveNext
			Session("RSBookMark") = Session("RSBookMark") + 1
			If RS.EOF Then 
				RS.MoveFirst
				Session("RSBookMark") = 1
				Statusmsg = "Frame ID: " & Request.QueryString("FRAMEID") & " may not be part of this call flow any longer."
				Exit Do
			End If
		Loop
			FRAMEID = RS("FRAME_ID")
	End If

End If

If 	Request.QueryString("NAVIGATE") <> "" Then
	Select Case Request.QueryString("NAVIGATE")
		Case "next"
			RS.Move(Session("RSBookMark"))
			If Not RS.EOF Then
				Session("RSBookMark") = Session("RSBookMark") + 1
				FRAMEID = RS("FRAME_ID")
			Else
				RS.MoveFirst
				Session("RSBookMark") = 1
				FRAMEID = RS("FRAME_ID")
			End If
		Case "previous"
			RS.moveFirst
			If (Session("RSBookMark")-1) >= 1 Then
				RS.move(Session("RSBookMark")-2)
				Session("RSBookMark") = Session("RSBookMark") - 1
				FRAMEID = RS("FRAME_ID")
			Else
				RS.move(Clng(Session("FrameCount"))-1)
				Session("RSBookMark") = Cint(Session("FrameCount"))
				FRAMEID = RS("FRAME_ID")
			End If
		Case Else
	End Select
End If
LogStatusGroupBegin()
SharedCount = CheckSharedFrame(RS("FRAME_ID"), True, False, 2, False, False, 0)
SharedFrameCount = CheckSHaredCallFlow(Request.QueryString("CFID"),True, False, 2, False, False, 0) 
LogStatusGroupEnd()	
	
	
	
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE>Call Flow Frame Visual Editor</TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub Window_OnLoad
	SELECTEDTAB = ""
'	top.window.document.title = "Call Flow Editor"
	FRAMELIST.value = "<%= RS("FRAME_ID") %>"
	If Parent.frames("WORKAREA").location.href <> "about:BLANK" Then
	Select Case parent.frames("WORKAREA").TheWindow.TabsControl.GetCurrentTabNum()
		Case "1"
			SELECTEDTAB = "General"
		Case "2"
			SELECTEDTAB = "Layout"
		Case "3"
			SELECTEDTAB = "Rules"
		Case "4"
			SELECTEDTAB = "SQL Data"
		Case "5"
			SELECTEDTAB = "Listing"
		Case "6"
			SELECTEDTAB = "Frame Order"
		Case Else
			SELECTEDTAB = ""
	End Select
	End If
	Parent.frames("WORKAREA").location.href = "CF_TabWindow.asp?CFID=<%=RS("CALLFLOW_ID") %>&FRAMEID=<%= FRAMEID %>&ACTIVETAB=" & SELECTEDTAB
	<% If SharedCount > 1 Then %>
		document.all.StatusSpan.InnerHTML = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
		<%	If CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
				document.all.SpanFrameSharedCount.innerHTML = "<%=SharedCount%>" & "<Font size=1 color='Maroon'>+</Font>"
		<%	Else %>
				document.all.SpanFrameSharedCount.innerHTML = "<%=SharedCount%>"
		<%	End If
		End If %>
	<% If SharedFrameCount > 1 Then %>
			document.all.StatusSpan2.InnerHTML = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
		<%	If CInt(SharedFrameCount) = CInt(Application("MaximumSharedCount")) Then %>
				document.all.SpanCallFlowSharedCount.innerHTML = "<%= SharedFrameCount %>" & "<Font size=1 color='Maroon'>+</Font>"
		<%	Else %>
				document.all.SpanCallFlowSharedCount.innerHTML = "<%= SharedFrameCount %>"
		<%	End If
		End If 
		If Statusmsg <> "" Then %>
		lret = msgbox("<%= Statusmsg %>",0,"FNSDesigner")
	<% End if 
		  Statusmsg = "" %>
		  
<%
	strConnName= "Customized Connection"
	For Each Key in Application.Contents
		If Application.Contents(Key) =Session("ConnectionString") Then
			strConnName = Key
			Exit For
		End If
	Next
%>
	top.window.document.Title = "Call Flow Editor (<%=strConnName%>)"
			  
End Sub

Sub BtnNext_onclick
If Instr(1, parent.frames("WORKAREA").location, "NewFrame.asp") = 0 Then
	lret = parent.frames("WORKAREA").TheWindow.CanActiveFrameUnloadNow()
	If lret = True Then
		self.location.href = "CFVisualEditor.asp?NAVIGATE=next&CFID=<%=RS("CALLFLOW_ID") %>"
	End If
Else
	MsgBox "Cancel or save new frame before navigating", 0 , "FNSDesigner"
End If
End Sub

Sub BtnPrevious_onclick
	lret = parent.frames("WORKAREA").TheWindow.CanActiveFrameUnloadNow()
	If lret = True Then
		self.location.href = "CFVisualEditor.asp?NAVIGATE=previous&CFID=<%=RS("CALLFLOW_ID") %>"	
	End If
End Sub

Sub BtnNewFrame_onclick
<% If MODE="RO" Then Response.write (" Exit Sub ") %>
	Parent.frames("WORKAREA").location.href = "NewFrame.asp?CFID=<%=RS("CALLFLOW_ID") %>"
End Sub

Sub BtnDetachFrame_onclick
<% If MODE="RO" Then Response.write (" Exit Sub ") %>
	lret = msgbox ("Are you sure you want to detach frame: " & FRAMELIST(FRAMELIST.selectedIndex).text, 1, "FNSNet")
	if lret = 1 Then
		Parent.frames("HIDDENPAGE").location.href = "DetachFrame.asp?FRAMEID=<%= FRAMEID %>&CFID=<%=RS("CALLFLOW_ID") %>"
	End If
End Sub

Sub BtnAttachFrame_onclick
<% If MODE="RO" Then Response.write (" Exit Sub ") %>
	strURL = "FrameSearchModal.asp"
	showModalDialog  strURL  ,FrameObj ,"dialogWidth:450px;dialogHeight:450px;center"
	If FrameObj.FrameID <> "" Then
		Parent.frames("HIDDENPAGE").location.href = "AttachFrame.asp?SEQUENCE=<%= RS("SEQUENCE") %>&FRAMEID=" & FrameObj.FrameID  & "&CFID=<%=RS("CALLFLOW_ID") %>"
	End If
End Sub

Sub BtnCopyFrame_onclick
<% If MODE="RO" Then Response.write (" Exit Sub ") %>
	lret = msgbox ("Are you sure you want to copy this frame: " & FRAMELIST(FRAMELIST.selectedIndex).text & Chr(13) & "Copying this frame will create a new unique instance of this frame" & VbCrlf & "and the current frame will be detached.", 1, "FNSNet")
	if lret = 1 Then
		Parent.frames("HIDDENPAGE").location.href = "CopyFrame.asp?FRAMEID=<%= FRAMEID %>&CFID=<%=RS("CALLFLOW_ID") %>"
	End If
End Sub

Sub BtnStatus_onclick
	If CLng(<%=SharedCount%>) > 1 Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other details reported", 0 , "FNSNet"
	End If
End Sub

Sub BtnRefCount_onclick
	lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedFrame=True&ID=<%= RS("FRAME_ID") %>", Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
End Sub

Sub FRAMELIST_onchange
If FRAMELIST.value <> "" Then
	self.location.href = "CFVisualEditor.asp?CFID=<%=RS("CALLFLOW_ID") %>&FRAMEID="	& document.all.FRAMELIST.Value
End If
End Sub

Sub BtnRefCount2_onclick
	lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedCallFlow=True&ID=<%= RS("CALLFLOW_ID") %>", Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
End Sub

Sub BtnStatus2_onclick
	If CLng(<%=SharedFrameCount%>) > 1 Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other details reported", 0 , "FNSNet"
	End If
End Sub

-->
</SCRIPT>
<SCRIPT LANGUAGE=JavaScript>
<!--
function CFrameSearchObj()
{
	this.FrameID = "";
}
var FrameObj = new CFrameSearchObj();
//-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR=#d6cfbd topmargin=0 rightmargin=0 leftmargin=0 bottommargin=0>
<TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=1>
<TR>
<TD COLSPAN=10>
<TABLE>
<TR>
<TD><IMG SRC="../IMAGES/Attach.GIF" BORDER=0 STYLE="CURSOR:HAND" TITLE="Attach existing frame to this Call Flow" ID=BtnAttachFrame name=BtnAttachFrame></TD>
<TD><IMG SRC="../IMAGES/NewFrame.GIF" BORDER=0 STYLE="CURSOR:HAND" TITLE="New Frame" ID=BtnNewFrame name=BtnNewFrame></TD>
<TD CLASS=LABEL><NOBR>Call Flow Name: (<%= Request.QueryString("CFID") %>)</TD>
<TD CLASS=LABEL><NOBR><LABEL CLASS=LABEL><%= RS("CFNAME") %></LABEL></TD>
<TD CLASS=LABEL valign=MIDDLE><NOBR><IMG SRC="..\images\RefCount.gif" STYLE="CURSOR:HAND" ID=BtnRefCount2 align=absmiddle TITLE="Shared Count">
<TD CLASS=LABEL valign=MIDDLE>
:<span id="SpanCallFlowSharedCount"><%= SharedFrameCount %></span></TD>
<TD CLASS=LABEL><IMG SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER=0 TITLE="Status Report" NOWRAP NAME=BtnStatus2 ID=BtnStatus2 align=absmiddle>
<SPAN ID=StatusSpan2 STYLE="COLOR:">&nbsp;</SPAN>
</TD>
</TR>
</TABLE>
<TABLE BORDER=0>
<TR>
<TD CLASS=GrpLabelLine colspan=22 HEIGHT=1></TD>
</TR>
<TR>
<TD><IMG SRC="../IMAGES/Detach.gif" BORDER=0 STYLE="CURSOR:HAND" TITLE="Detach this frame from the current Call Flow" ID=BtnDetachFrame name=BtnDetachFrame></TD>
<TD><IMG SRC="../IMAGES/MakeCopy.GIF" BORDER=0 STYLE="CURSOR:HAND" TITLE="Copy frame as new" ID=BtnCopyFrame name=BtnCopyFrame></TD>
</TD>
<TD CLASS=LABEL>Frame Name: (<%= RS("FRAME_ID") %>)
<SELECT NAME="FRAMELIST" CLASS=LABEL >
<% 
Do While Not RSList.EOF %>
<OPTION VALUE="<%= RSList("FRAME_ID") %>"><%= RSList("NAME") %>
<%
RSList.MoveNext
loop
%>
</SELECT>
</TD> 
<TD CLASS=LABEL ALIGN=LEFT valign=BOTTOM><NOBR><IMG SRC="..\images\RefCount.gif" STYLE="CURSOR:HAND" ID=BtnRefCount align=absmiddle TITLE="Shared Count"></TD>
<TD CLASS=LABEL ALIGN=LEFT valign=BOTTOM>
:<span id="SpanFrameSharedCount"><%=SharedCount%></span></TD>
<TD CLASS=LABEL ALIGN=LEFT valign=BOTTOM><IMG SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER=0 TITLE="Status Report" NOWRAP align=absmiddle NAME=BtnStatus ID=BtnStatus>
<SPAN ID=StatusSpan STYLE="COLOR:">&nbsp;</SPAN>
 </TD>
 </TR>
</TABLE>
</TD>
<TD ALIGN=RIGHT VAlIGN=TOP>
<TABLE CELLSPACING=2 CELLPADDING=0> 
<TR>
<TD CLASS=LABEL><IMG SRC="../IMAGES/NAV2a.gif" STYLE="CURSOR:HAND" TITLE="Previous Frame" BORDER=1 name=BtnPrevious></TD>
<TD CLASS=LABEL STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:BLACK" height=0><NOBR><%= Clng(Session("RSBookMark")) %> of <%= Session("FrameCount") %></TD>
<TD CLASS=LABEL><IMG SRC="../IMAGES/NAV1a.gif" STYLE="CURSOR:HAND" TITLE="Next Frame" BORDER=1  name=BtnNext></TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>
</BODY>
</HTML>
