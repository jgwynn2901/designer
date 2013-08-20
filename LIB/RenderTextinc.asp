<%
Function TruncateText(inText,inChar)
	if not IsNull(inText) then
		If Len(inText) < inChar Then
			TruncateText = inText
		Else
			TruncateText = Mid ( inText, 1, inChar) & " ..."
		End If
	end if
End Function

Function ReplaceQuotesInText(inText)
	if not IsNull(inText) then
		ReplaceQuotesInText = Replace(inText,"""","&quot;")
	end if
End Function
%>