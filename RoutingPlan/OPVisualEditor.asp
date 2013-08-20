<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\CheckSharedOD.inc"-->
<%
Dim SharedCount, SharedCountText
	SharedCount = 0
	SharedCountText = "Ready"
If Request.QueryString("ODID") = "" AND Request.QueryString("NAVIGATE") = "" Then
	Session("ErrorMessage") = "On Page OPVisualEditor: Querystring ODID was null or not numeric"
	Response.Redirect "../directerror.asp"
End If

If Request.QueryString("ODID") <> ""  Then

SharedCount = CheckSharedOD(CLng(Request.QueryString("ODID")), True, True, 1, False, False, 0)

If Not Isnumeric(Request.QueryString("ODID")) Then
	Session("ErrorMessage") = "On Page OPVisualEditor: Querystring OPID was null or not numeric"
	Response.Redirect "../directerror.asp"
End If
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST2 = ""
	Session("RSCount") = ""
	Set Session("RSPages") = nothing
		
	If Session("RSIndex")  = "" Then
		Session("RSIndex")  = 0
	End If
	If Request.QueryString("NAVIGATE") = "" Then
		Session("RSIndex") = 1
	End If
	
	SQLST2 = SQLSST2 & "SELECT COUNT(*) AS PageCount FROM OUTPUT_PAGE WHERE OUTPUTDEF_ID=" & Request.QueryString("ODID")
	SQLST = SQLST & "SELECT OUTPUT_PAGE_ID,PAGE_NUMBER, NAME, ORIENTATION  FROM OUTPUT_PAGE WHERE OUTPUTDEF_ID=" & Request.QueryString("ODID") & " ORDER BY PAGE_NUMBER"
	Set RS = Conn.Execute(SQLST)
	Set RSList = Conn.Execute(SQLST)
	Set RS2 = Conn.Execute(SQLST2)
	If RS.EOF AND RS.BOF Then
		Session("ErrorMessage") = "On Page OPVisualEditor the SQL Statement returned no records"
		Response.Redirect "../directerror.asp"
	End If
	Session("RSCount") = RS2("PageCount")
	Set Session("RSPages") = RS
	OUTPUT_PAGE_ID = RS("OUTPUT_PAGE_ID")
	If Request.QueryString("OPID") <> "" Then
		Do While Clng(Request.QueryString("OPID")) <> Clng(RS("OUTPUT_PAGE_ID"))
			Session("RSPages").MoveNext
			Session("RSIndex") = Session("RSIndex") + 1
		Loop 
	   OUTPUT_PAGE_ID = Session("RSPages")("OUTPUT_PAGE_ID")
	End If
End If

If Request.QueryString("NAVIGATE") <> "" Then
Select Case Request.QueryString("NAVIGATE")
		Case "NEXT"
		RS.Move(Session("RSIndex"))
		If Not Session("RSPages").EOF Then
			Session("RSIndex") = Session("RSIndex") + 1
			OUTPUT_PAGE_ID = RS("OUTPUT_PAGE_ID")
		Else
			RS.MoveFirst
			Session("RSIndex") = 1
			OUTPUT_PAGE_ID = RS("OUTPUT_PAGE_ID")
		End If
		Case "PREVIOUS"
			RS.MoveFirst
			If (Session("RSIndex")-1) >= 1 Then
				RS.Move(Session("RSIndex")-1)
				Session("RSIndex") = Session("RSIndex") - 1
				OUTPUT_PAGE_ID = RS("OUTPUT_PAGE_ID")
			Else
				RS.Move(Cint(Session("RSCount"))-1)
				Session("RSIndex") = Cint(Session("RSCount"))
				OUTPUT_PAGE_ID = RS("OUTPUT_PAGE_ID")
			End If
		Case Else
	End Select
End If
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	SELECTEDTAB = ""
'	top.window.document.Title = "Output Definition Editor"
	PAGELIST.value = "<%= RS("OUTPUT_PAGE_ID") %>"
	If Parent.frames("WORKAREA").location.href <> "about:blank" Then
		Select Case parent.frames("WORKAREA").TheWindow.TabsControl.GetCurrentTabNum()
			Case "1"
				SELECTEDTAB = "General"
			Case "2"
				SELECTEDTAB = "Layout"
			Case "3"
				SELECTEDTAB = "Listing"
			Case Else
				SELECTEDTAB = ""
		End Select
	End If
<%  IF CInt(SharedCount) > 1 THEN
		SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
		If CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
				document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>" & "<Font size=1 color='Maroon'>+</Font>"
	<%	Else %>
				document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>"
	<%	End If
	END IF %>
				
	Parent.frames("WORKAREA").location.href="OP_TabWindow.asp?AHSID=<%= Request.QueryString("AHSID") %>&ODID=<%= Request.QueryString("ODID") %>&OPID=<%= OUTPUT_PAGE_ID %>&ACTIVETAB=" & SELECTEDTAB
<%
	strConnName= "Customized Connection"
	For Each Key in Application.Contents
		If Application.Contents(Key) =Session("ConnectionString") Then
			strConnName = Key
			Exit For
		End If
	Next
%>
	top.window.document.Title = "Output Definition Editor (<%=strConnName%>)"
End Sub

Sub BACK_onclick
lret = parent.frames("WORKAREA").TheWindow.CanActiveFrameUnloadNow()
	If lret = True Then
		self.location.href = "OPVisualEditor.asp?NAVIGATE=PREVIOUS&ODID=<%= Request.QueryString("ODID") %>"
	End If
End Sub

Sub FORWARD_onclick
lret = parent.frames("WORKAREA").TheWindow.CanActiveFrameUnloadNow()
	If lret = True Then
		self.location.href = "OPVisualEditor.asp?NAVIGATE=NEXT&ODID=<%= Request.QueryString("ODID") %>"
		parent.frames("WORKAREA").TheWindow.SetActiveTab("LAYOUT")
	End If
End Sub

Sub PAGELIST_onchange
If PAGELIST.value <> "" Then
	self.location.href = "OPVisualEditor.asp?AHSID=<%= Request.QueryString("AHSID") %>&ODID=<%= Request.QueryString("ODID") %>&OPID="	& document.all.PAGELIST.Value 
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
	lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedOD=True&ID=<%=Request.QueryString("ODID")%>", Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
End Sub
-->
</SCRIPT>
</HEAD>
<BODY leftmargin=0 rightmargin=0 BGCOLOR='<%=BODYBGCOLOR%>'  topMargin=3>
<TABLE WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<TR>
<TD>
<TABLE>
<TR>
<TD CLASS=LABEL>Page Name:</TD>
<TD CLASS=LABEL>
<SELECT NAME=PAGELIST CLASS=LABEL>
<% Do While Not RSList.EOF %>
<OPTION VALUE="<%= RSList("OUTPUT_PAGE_ID") %>"><%= RSList("NAME") %>
<%
RSList.MoveNext
Loop
RSList.Close %>
</SELECT>
<SPAN ID=PAGENAMEID></SPAN>
</TD>
<TD CLASS=LABEL valign=MIDDLE><NOBR><IMG SRC="..\images\RefCount.gif" STYLE="CURSOR:HAND" ID=BtnRefCount align=absmiddle TITLE="Shared Count">
<TD CLASS=LABEL valign=MIDDLE>
:<span id="SpanSharedCount"><%=SharedCount%></span>
</TD>
<TD CLASS=LABEL><IMG SRC="../IMAGES/StatusRpt.gif" STYLE="CURSOR:HAND" BORDER=0 TITLE="Status Report" NOWRAP NAME=BtnStatus ID=BtnStatus align=absmiddle>
</TD>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=SharedCountText%></SPAN>
</TD>
</TR>
</TABLE>
</TD>
<TD ALIGN=RIGHT>
<TABLE CELLSPACING=1 CELLPADDING=0>
<TR ALIGN=RIGHT>
<TD><IMG SRC="../IMAGES/NAV2a.GIF" BORDER=1 STYLE="CURSOR:HAND" NAME="BACK"></TD>
<TD CLASS=LABEL STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:BLACK"><%= Session("RSIndex") %> of <%= Session("RSCount") %></TD>
<TD><IMG SRC="../IMAGES/NAV1a.GIF" BORDER=1 STYLE="CURSOR:HAND" NAME="FORWARD"></TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>
</BODY>
</HTML>