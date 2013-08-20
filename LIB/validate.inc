<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

'Form Field Validation Functions
'Author:   Jon Raynor
'Date:     3/26/2004


Function CheckLen(inputstring,expectedlen)

If len(Cstr(inputstring)) = expectedlen then
	CheckLen = True
	exit function
End if
	CheckLen = False

End Function

Function CheckLenRange(inputstring,min,max)

If len(Cstr(inputstring)) >= min and len(Cstr(inputstring)) <= max then
	CheckLenRange = True
	exit Function
End If
	CheckLenRange = False 

End Function

Function IsNumber(inputstring)

Dim i 

If IsStrEmpty(inputstring) Then
	IsNumber = False
	Exit Function
End If

If not IsNumeric(inputstring) Then 
	IsNumber = False
	Exit Function
End If

IsNumber = True

End Function

Function IsStrEmpty(inputstring)

If Len(inputstring) = 0 Then
	IsStrEmpty = True
	Exit Function
End If

IsStrEmpty = False

End Function

Function RemoveChars(inputstring,chartoremove)

Dim i
Dim pos
Dim parts

pos = InStr(inputstring, chartoremove)

If pos Then
    
    parts = Split(Cstr(inputstring), Cstr(chartoremove))
    inputstring = ""

    For i = 0 To UBound(parts)
        inputstring = inputstring & parts(i)
    Next

End If

RemoveChars = inputstring

End Function

</SCRIPT>
