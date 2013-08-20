<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>


<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Validation Results:</title>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub BtnClose_OnClick
	window.close
End Sub

sub Window_OnLoad
End Sub
</script>

</head>
<body BGCOLOR="#d6cfbd">

<%

dim ScriptText
Dim ResText, ErrorText,DescrText, Text, Line, Column
Dim Source

Dim validator
Set validator = server.CreateObject("CEvaluate.CEval")
ResText= validator.EvaluateScript(Request.QueryString("RuleText"),ErrorText, DescrText, Text, Line, Column)

Dim NewText1, Highlighted, NewText2
Dim pos
pos=CInt(Column)

if Text<>"" AND pos >0 then
	NewText1=Left(Text, pos)
	dim lText
	lText=Len(Text)
	
	if lText>Pos then
		NewText2 = Right(Text,lText - pos-1)
		Highlighted=Mid(Text, pos+1,1)
	else
		Highlighted=" "
		NewText2=""
	end if
end if
%>

<div align="center"><br>
<%if ErrorText = "" then %>
result:<input type="text" CLASS="LABEL" NAME="ResText" value="<%=ResText%>" ID="RESULT">
&nbsp&nbsp&nbsp&nbsp<b>There's no compilation errors</b><br>
<%else %>
error code:<input type="text" CLASS="LABEL" NAME="ErrorText" size="4" value="<%=ErrorText%>" ID="ERROR">
at position:<input type="text" CLASS="LABEL" NAME="POSITION"  size="3" value="<%=Column%>" ID="POSITION">
<br><br>
<td>
<textarea class="LABEL" MAXLENGTH="600" name="DescrText" style="BACKGROUND-COLOR: antiquewhite overflow:hidden" 
readOnly cols="60" rows="4" ID="DescrText" ><%=DescrText%></textarea>
</td>
<br><br>
</div>
&nbsp&nbsp<%=NewText1%><STRONG><FONT style="BACKGROUND-COLOR: #99ff33" color="#000099"><%=Highlighted%></FONT></STRONG><%=NewText2%>
<br>
<% end if%>
<div align="center">
<br><br><button ID="BtnClose">Close</button>
</div>
</body
</html>



