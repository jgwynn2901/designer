<%@ LANGUAGE="VBSCRIPT" %>
<%  Response.expires=0
	If Session("CurrentUser").CanUserDo("Account Maintenance","View") <> true  Then 
		Response.redirect "../AccessDenied.asp"
	End If
	

%>
<html>

<head>
<link rel="stylesheet" type="text/css" href="../fnsnet.css">
<title>Cache List Maintence </title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
</head>

<body background="" >

<table cellspacing=60 align=center>
	<tr>	
		<td><a href="AM_CallFlowCacheListmaintenance.asp"> Call Flow Cache List Maintenance </a></td>
	</tr>
	<tr>	
		<td><a href="AM_DataCachemaintenance.asp"> Data Cache List Maintenance </a></td>
	</tr>
</body>
</html>
