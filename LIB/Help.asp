Function LaunchHelp( URL )
	strURL = "HTTP://<%= Request.Servervariables("server_name") %>/FNSdesigner/Help/" & URL
	lret = window.showHelp(strURL)
End Function
