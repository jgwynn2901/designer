<%
If Not Session("SecurityObj").CheckDBConnection Then
	Session("NAME") = ""
End If
If Session("NAME") = "" Then 
	Response.Redirect "HTTP://" & Request.Servervariables("server_name") & Application("VirtualRoot") & "/Login.asp"
End If
%>