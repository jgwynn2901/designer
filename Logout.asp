<% Response.Expires=0 %>
<html> 
<head>
<title>FNSNet Designer Log Out</title>
<link rel="stylesheet" type="text/css" href="FNSDESIGN.css">
<script LANGUAGE="VBScript">
Sub CloseMe()
	close	
End sub

Sub IndicateWait()
	SPANID.innerText = 	SPANID.innerText & " . "
End sub

Sub window_onLoad()
	setInterval "IndicateWait",500
	setTimeout "CloseMe", 2000
End Sub
</script>
</head>
<body bgcolor="#C0C0C0" class="LABEL" ALIGN="CENTER">
<p>&nbsp;</p>
<big>
<p align="center">Logging Out FNSNet Designer<br><br>Please wait</p>
<SPAN ID="SPANID" style="position:absolute;top:63;left:200"></SPAN>

</big>
</body>
<%	Session.Abandon %>
</html>
