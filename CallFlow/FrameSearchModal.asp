<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Frame Search</title>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	<%
	if request.querystring("CONTAINERTYPE") <> "FRAMEWORK" then
	%>
	window.dialogArguments.FrameID = ""
	<%
	end if
	%>
End Sub

Sub BtnCancel_OnClick
	window.dialogArguments.FrameID = ""
	window.close
End Sub

Sub BtnView_OnClick
	dim cURL, cFrameID
	
	cFrameID =  document.frames("FrameSearchIFrame").document.frames("WORKAREA").GetFrame
	If cFrameID = -1 OR cFrameID="X" Then
		MsgBox "Please select a Frame.", 0, "FNSNetDesigner"
	Else
		cURL = "showFrame.asp?FRAMEID=" & cFrameID
		showModalDialog  cURL  ,,"dialogWidth:550px;dialogHeight:450px;center;border:thick;maximize:yes;help:no"
	end if
End Sub

Sub BtnDelete_OnClick
	dim cFrameID
	
	cFrameID =  document.frames("FrameSearchIFrame").document.frames("WORKAREA").GetFrame
	If cFrameID = -1 OR cFrameID="X" Then
		MsgBox "Please select a Frame to delete.", 0, "FNSNetDesigner"
	Else
		if msgbox("Are you sure you want to delete the Frame " & cFrameID & "?",vbYesNo,"FNSDesigner") = vbYes then
			document.frames("FrameSearchIFrame").document.frames("hiddenPage").location.replace("FrameDelete.asp?FRAMEID=" & cFrameID)
		end if
	end if
End Sub

Sub BtnSelect_OnClick
	FrameID =  document.frames("FrameSearchIFrame").document.frames("WORKAREA").GetFrame
	If FrameID = -1 OR FrameID="X" Then
		MsgBox "Please select a Frame.", 0, "FNSNetDesigner"
	Else
		window.dialogArguments.FrameID = FrameID
		window.close	
	End If
End Sub

Sub Document_OnKeyDown()
	If window.event.altKey Then
		KeyPress = Chr(window.event.keyCode)
		Select Case KeyPress
			case "C":
				document.frames("FrameSearchIFrame").document.frames("TOP").BtnSearch_onclick
			case "L":
				document.frames("FrameSearchIFrame").document.frames("TOP").BtnClear_onclick
		End Select
	End If
End Sub

</script>
</head>

<BODY  leftmargin=0 topmargin=0 bottommargin=0 rightmargin=0 BGCOLOR='<%=BODYBGCOLOR%>' >
<iframe id="FrameSearchIFrame" FRAMEBORDER="0" src="FrameSearch-f.asp" WIDTH="100%" HEIGHT="90%">
</iframe>
<table>
<tr>
<%
if Request.QueryString("FRAMEMAINT") = "" then
%>
	<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnSelect" ACCESSKEY="S"><u>S</u>elect</button></td>
	<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnCancel">Cancel</button></td>
<%
else
%>
	<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnView">View</button></td>
<%
	if HasViewPrivilege("FNSD_FRAME_DELETE","") then
	%>
		<td CLASS="LABEL"><button CLASS="StdButton" NAME="BtnDelete">Delete</button></td>
	<%
	end if
end if
%>
</tr>
</table>
</body>
</html>
