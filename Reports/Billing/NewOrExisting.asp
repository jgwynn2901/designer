<%@ Language=VBScript %>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
dim nSel, cMsg

cMsg = "There are already reports for this period." & chr(10) & chr(13) & _
		"Do you want to run a new report?"
nSel = MsgBox(cMsg, 35, "Reports for <%=Request.QueryString("CUSTNAME")%>" & cCustName)
if nSel = 6 then
	self.location.href = "runReport.asp?" & "<%=Request.QueryString%>"
elseif nSel = 7 then
	self.location.href = "SelectReport.asp?" & "<%=Request.QueryString%>"
else
	parent.frames("header").location.href = "SFR.asp"
end if
End Sub

-->
</SCRIPT>
</HEAD>
<body BGCOLOR="#d6cfbd">
</BODY>
</HTML>
