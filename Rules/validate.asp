<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"

Dim oDictMap ' key =index, item=index
Dim labels
Set oDictMap	=	CreateObject("Scripting.Dictionary")
%>

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Validate Rules - FNS Net Designer</title>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub BtnClose_OnClick
	window.close
End Sub

Sub BtnValidate_OnClick

	dim strResult
	dim bRresult
	dim source
	dim ncount

	source = document.all.RuleText.value
	source = Replace(source, "??", """")
	
	nCount = CInt(document.all.Count.value)
	if (nCount=0) then
		strResult=source
	else
		bResult=ReplaceTextInTildas (ncount, source, strResult)
	end if	
	
	strResult=encodeURIComponent(strResult)
	lret = window.showModalDialog ("validateExe.asp?RuleText=" &  strResult, Null,  "dialogWidth=600px; dialogHeight=300px; center=yes")

End Sub

function ReplaceTextInTildas (byval count, byval StrSource, byref strRes)
	Dim i
	dim strOldTex
	dim strNewText
	
	strRes=strSource
		
	for i=0 to count-1
		strReplace=document.all("ARGUMENT" & i).value
		strFind="~" & document.all("ARGLABEL" & i).value & "~"
		strRes=Replace (strRes,strFind, strReplace)
	next
	ReplaceTextInTildas=true
end Function 'ReplaceTextInTildas
</script>

</head>
<body BGCOLOR="#d6cfbd">
<SCRIPT LANGUAGE="VBScript" RUNAT="Server"> 
	
	' set a new label to both dictionary objects
	function IncludeParsedText (byval LabelIndex, byval strText)
		if oDictMap.Exists(strText)=false then ' add one more key to map!
			oDictMap(strText)=oDictMap.Count+1
		end if
	end function

	function ParseAllTildas (byval StrSource)
	
	Dim ToContinue, nCount
	Dim StartPos, NextPos
	Dim ParsedText, tilda

	ToContinue=true
	StartPos=1
	tilda="~"
	nCount=0
				
	Do until ToContinue=false
	'find first tilda
		NextPos=InStr (StartPos, strSource, tilda)
		if NextPos>0 then ' found
			'find second tilda
			StartPos=NextPos+1
			NextPos=InStr (StartPos, strSource, tilda)

			if NextPos>0 then ' found
				ParsedText=Mid(strSource,StartPos,NextPos-StartPos)
				ncount=ncount+1
				StartPos=NextPos+1
				IncludeParsedText nCount, ParsedText
				Redim labels (oDictMap.Count)
				labels=oDictMap.Keys
			else
				ParseAllTildas=false
				exit function
			end if
		
		else
			ToContinue=false
		end if
		Loop
		ParseAllTildas=true

	end Function 'ParseAllTildas

	function GetLabel (byval index) 
	
		if index < oDictMap.count then
				GetLabel = labels(index)
		else
				GetLabel = ""
		end if
	end Function 


</script>
<%
dim res
dim text

text= Request.QueryString("RuleText")
Text = Replace (Text,"""","??")
res = ParseAllTildas (Request.QueryString("RuleText"))
%>


<div style="position:absolute;top:4;left:10";width:'100%'>
<input type="hidden"  NAME="RuleText" value="<%=text%>" ID="RuleText">
<input type="hidden" NAME="Count" value="<%=oDictMap.Count%>" ID="Count">
</div>
<div align="center">
<%
dim index
dim idlabel
dim idarg
if oDictMap.count <> 0 then 
	for index=0 to oDictMap.count-1
		idlabel="ARGLABEL" & CStr(index)
		idarg="ARGUMENT" & CStr(index)
%>

<br><tr>
<td><input type="text" name = "ARGLABEL" style="BACKGROUND-COLOR: #d6cfbd" size="60" readonly ID="<%=idlabel%>" value = <%=labels(index)%> > </td>
<td><input type="text" NAME="ARGUMENT" size="10" ID="<%=idarg%>"></td></tr>
<%	Next%>

<%else%>
<!-----<td>
<br><br> Rule to be validated:<br>
<textarea class="LABEL" MAXLENGTH="600" name="RuleText2" style="BACKGROUND-COLOR: #d6cfbd" overflow:hidden" 
readOnly cols="50" rows="15" ID="RuleText2" ><%=Text%></textarea></td>---->
<br> <i>There're no required input parameters.</i><br><br><br><br>
<%end if %>
<br>
<br>
<tr><td><button CLASS="StdButton" NAME="BtnValidate" ACCESSKEY="V"  ID="BtnValidate"><u>V</u>alidate</button></td>
<tr><td><button CLASS="StdButton" NAME="BtnClose" ACCESSKEY="V"  ID="BtnClose">Close</button></td>

</div>
</body>
</html>
