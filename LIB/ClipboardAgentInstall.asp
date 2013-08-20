<!--#include file="ClipboardCLSID.inc"--> 
<HTML>
<HEAD>
<TITLE>ATL 3.0 test page for object ClipboardAgent</TITLE>
</HEAD>
<BODY>
<OBJECT ID="ClipboardAgent" 
<%GetClipboardCLSID("")%>
width=1 height=1>
<PARAM NAME="MaxPropertiesStringLength" VALUE="1000">
<PARAM NAME="MaxPropertyNameLength" VALUE="20">
<PARAM NAME="MaxPropertyValueLength" VALUE="200">
<PARAM NAME="NameValueDelimiter" VALUE=",">
<PARAM NAME="PropertyItemDelimiter" VALUE="|">

</OBJECT>
<BR>
<INPUT ID=SETPROP TYPE=BUTTON Value="Add Property (Name, Value):">
<INPUT ID=PROPNAME TYPE=TEXT><INPUT ID=PROPVAL TYPE=TEXT>
<BR>
<BR>
<INPUT ID=SETPROPSTR TYPE=BUTTON Value="Set Properties String to Clipboard">
<INPUT ID=GETPROPSTR TYPE=BUTTON Value="Get Properties String to Clipboard">
<BR>
Clipboard Data:&nbsp<SPAN ID=CLIPDATA></SPAN>
<BR>
<BR>
<INPUT ID=FMTNAME TYPE=BUTTON Value="Is FNS Clipboard Data Available?">
<BR>
<BR>


<INPUT ID=GETPROP TYPE=BUTTON Value="Get Property by Name">
<INPUT ID=GETPROPNAME TYPE=TEXT>
<BR>
</BODY>
</HTML>
<script LANGUAGE="VBSCRIPT">

Sub FMTNAME_onclick()
	Dim test
	test = ClipboardAgent.IsClipboardDataAvailable
	If test = True Then
		MsgBox "IsClipboardDataAvailable=TRUE"
	Else
		MsgBox "IsClipboardDataAvailable=FALSE"
	End If

End Sub

Sub SETPROPSTR_onclick()
	ClipboardAgent.SetPropertiesToClipboard
End Sub

Sub GETPROPSTR_onclick()
	ClipboardAgent.GetPropertiesFromClipboard
	CLIPDATA.innerHTML = ""
	CLIPDATA.innerHTML = ClipboardAgent.PropertiesString
End Sub

Sub SETPROP_onclick()
	ClipboardAgent.AddProperty PROPNAME.value, PROPVAL.value
End Sub

Sub GETPROP_onclick()
	MsgBox ClipboardAgent.GetProperty(GETPROPNAME.value)
End Sub
</script>