<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file=".\lib\common.inc"-->

<!--#include file=".\lib\StatusRptinc.asp"-->
<!--#include file=".\lib\RefCountRptinc.asp"-->
<!--#include file=".\lib\CheckSharedRule.inc"-->
<!--#include file=".\lib\CheckSharedAttribute.inc"-->
<!--#include file=".\lib\CheckSharedFrame.inc"-->
<!--#include file=".\lib\CheckSharedCallFlow.inc"-->
<!--#include file=".\lib\CheckSharedOutputDef.inc"-->

<!--#include file=".\lib\CheckSharedLUType.inc"-->
<%
Dim SharedCount

'SharedCount = CheckSharedRule(-132, True, True, 2, False, False, 0)
'SharedCount = CheckSharedRule(-121, True, True, 2, False, False, 0)
'SharedCount = CheckSharedAttribute(-4, True, True, 2, False, False, 0)
'SharedCount = CheckSharedFrame(-52, True, True, 2, False, False, 0)
'SharedCount = CheckSharedCallFlow(89, True, True, 2, False, False, 0)
'SharedCount = CheckSharedLUType(1, True, True, 2, False, False, 0)
SharedCount = CheckSharedOutputDef(50228, True, True, 2, False, False, 0)


Response.Expires = 0
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Status Report - FNS Net Designer</title>
<body BGCOLOR="#d6cfbd">
Test Status Report: <%=SharedCount%>
<BR>
<BUTTON ID=StatusRpt>Show Status Report</BUTTON>
<BR>
<BUTTON ID=RefCountRpt>Show Ref. Count Report</BUTTON>
<BR>
</body>
</html>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub StatusRpt_onclick()
	If CLng(<%=SharedCount%>) > 1 Then
		lret = window.showModalDialog (".\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
'		lret = window.Open (".\StatusRpt\StatusRpt.asp", "TEST",  "Width=580, Height=380,toolbar=no, location=no, center=yes")
	Else
		MsgBox "The shared count is only <%=SharedCount%>"
	End If
End Sub

Sub RefCountRpt_onclick()

	lret = window.showModalDialog (".\StatusRpt\RefCountRpt.asp?CheckSharedOutputDef=True&ID=50228", Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
	'lret = window.Open (".\StatusRpt\RefCountRpt.asp?CheckSharedAttribute=True&ID=-52", "TEST",  "Width=580, Height=400,toolbar=no,location=no,center=yes")

End Sub
</script>
