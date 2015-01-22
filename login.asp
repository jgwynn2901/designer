<!--#include file=lib/AHSTree.inc -->
<%
Function in_array(arrayElements,currentElement)
in_array = False
For i=0 To Ubound(arrayElements)
If Trim(arrayElements(i)) = Trim(currentElement) Then
in_array = True
Exit FOR
End If
Next
End Function
Response.Expires = 0

dim oConn, oRS

On Error Resume Next

BrowserObject = Request.ServerVariables("HTTP_USER_AGENT")
TokenArray = Split(BrowserObject," ",-1,1)
If ( TokenArray(2) <> "MSIE") Then
	Response.Redirect "BadBrowser.htm"
Else
	Version = Trim(Replace(TokenArray(3),";",""))
	'	extract version number (v6.0 contains a letter)
	cVer = ""
	for x=1 to len(Version)
		if isnumeric(mid(Version,x,1)) or mid(Version,x,1)="." then
			cVer = cVer & mid(Version,x,1)
		end if
	next
	If (csng(cVer) < 4) Then
		Response.Redirect "BadBrowser.htm"
	End If
End If

If Request.QueryString("ACTION") = "LOGIN" Then

	ConnectionString = CStr(Request.Form("ConnectString"))
	if ConnectionString = "" Then ConnectionString = CStr(Request.Form("CustomConnectString"))

	If IsEmpty(ConnectionString) Or CStr(ConnectionString) = "" Then
		Response.redirect "login.asp?Error=CONNSTR"
	End If

	If IsObject(Session("SecurityObj")) Then
		If Session("SecurityObj").IsLoggedOn() Then	Session("SecurityObj").LogOff()
	End If
		Response.cookies("UserName") = Request.Form("UNAME")
		Response.cookies("UserName").Expires = now + 7
		Response.cookies("UserDB") = Request.Form("ConnectString")
		Response.cookies("UserDB").Expires = now + 7

	Session("SecurityObj").m_ConnectionString = ConnectionString
	Session("SecurityObj").m_ShouldLog = false
	bHaveConn = Session("SecurityObj").CheckDBConnection		
	If bHaveConn Then
		Session("ConnectionString") = ConnectionString
		If Session("SecurityObj").Logon(Request("UNAME"),Request("PWD"),"") Then
			Session("NAME") = Request.Form("UNAME")
			Session("PASSWORD") = Request.Form("PWD")
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.Open ConnectionString
			SQL = ""
			SQL = SQL & "SELECT * FROM SETTING WHERE USER_ID=" & Session("SecurityObj").m_UserId
			Set oRS = oConn.Execute(SQL)
			Session("USERMAXRECORDS") = 50
			Session("USERTREELEVELS") = 2
			Session("USERTREECOUNT") = 10
			If IsObject(oRS) Then
				If Not oRS.EOF AND Not oRS.BOF Then

					Do While Not oRS.EOF

						currName = CStr(oRS("NAME"))

						Select Case currName
						case "DESIGNER_MAXRECORDS"
							Session("USERMAXRECORDS") = oRS("VALUE")
						case "FAVORITES_AHSID"
							If Isnull(oRS("VALUE")) Then
								Session("AHLIST") = ""
							Else
								Session("AHLIST") = oRS("VALUE")
								SetFilterByName "KEY=" & CStr(oRS("KEY_ID")), CStr(oRS("TYPE")), currName, CStr(oRS("VALUE"))
							End If
						case "EXP_FAVORITES_AHSID"
							If Isnull(oRS("VALUE")) Then
								Session("EXPANDLIST") = ""
							Else
								Session("EXPANDLIST") = oRS("VALUE")
								SetFilterByName "KEY=" & CStr(oRS("KEY_ID")), CStr(oRS("TYPE")), currName, CStr(oRS("VALUE"))
							End If
						case "USEWHERECLAUSE"
							SetFilterByName "KEY=" & CStr(oRS("KEY_ID")), CStr(oRS("TYPE")), currName, CStr(oRS("VALUE"))
						case "WHERECLAUSE"
							SetFilterByName "KEY=" & CStr(oRS("KEY_ID")), CStr(oRS("TYPE")), currName, CStr(oRS("VALUE"))
						case "MUSTINCLUDE"
							SetFilterByName "KEY=" & CStr(oRS("KEY_ID")), CStr(oRS("TYPE")), currName, CStr(oRS("VALUE"))
						case "MUSTEXCLUDE"
							SetFilterByName "KEY=" & CStr(oRS("KEY_ID")), CStr(oRS("TYPE")), currName, CStr(oRS("VALUE"))
						case "NODEDELIM"
							SetFilterByName "KEY=" & CStr(oRS("KEY_ID")), CStr(oRS("TYPE")), currName, CStr(oRS("VALUE"))
						case "DESIGNER_TREECOUNT"
						If IsNull(oRS("VALUE")) Then
							Session("USERTREECOUNT") = 10
						Else
							Session("USERTREECOUNT") = oRS("VALUE")
						End If
						case "DESIGNER_TREELEVEL"
							If IsNull(oRS("VALUE")) Then
								Session("USERTREELEVELS") = 2
							Else
								Session("USERTREELEVELS") = oRS("VALUE")
							End If
						case "LAYOUTCTLWIDTH"
							Session("LayoutCtlWidth") = oRS("VALUE")
						case "LAYOUTCTLHEIGHT"
							Session("LayoutCtlHeight") = oRS("VALUE")
						End Select

						oRS.MoveNext
					Loop
					oRS.close
					if Session("LayoutCtlWidth") = "" then
						Session("LayoutCtlWidth") = "51"
					end if
					if Session("LayoutCtlHeight") = "" then
						Session("LayoutCtlHeight") = "24"
					end if
				End If
			End If
			'cache ACCOUNT_USER records
			SQL = "SELECT ACCNT_HRCY_STEP_ID FROM ACCOUNT_USER WHERE USER_ID=" & Session("SecurityObj").m_UserId
			Set oRS = oConn.Execute(SQL)
			If IsObject(oRS) Then
				If Not oRS.EOF AND Not oRS.BOF Then
					Do While Not oRS.EOF
						If IsEmpty(Session("ACCOUNT_SECURITY")) Then
							Session("ACCOUNT_SECURITY") = CStr(oRS("ACCNT_HRCY_STEP_ID"))
						Else
							Session("ACCOUNT_SECURITY") = Session("ACCOUNT_SECURITY") & "," & CStr(oRS("ACCNT_HRCY_STEP_ID"))
						End If
						oRS.MoveNext
					Loop
				End If
			End If
			oRS.close
			set oRS = nothing
			oConn.Close
			set oConn = nothing
			Dim fnsEnvironments
			fnsEnvironments=Array("ANALYST", "FNSBA", "QA", "PREPRODUCTION", "PRODUCTION")
			
			Dim environment, ConStrAnalyst, ENVIRONMENT_ABBREVIATION
	        environment = CStr(Request.Form("environment"))	
	        SQL = "Select ENVIRONMENT_ABBREVIATION From DBConnection Where ENABLED = 'Y' AND ENVIRONMENT = '" & environment &"'"	
	        Set Connect = Server.CreateObject("ADODB.Connection")
	        Connect.Open "DSN=FNSANALYST;UID=FNSOWNER;PWD=CTOWN_DESIGNER"
	        Set Evn_RS = Connect.Execute(SQL)
	        IF Not Evn_RS.EOF AND Not Evn_RS.BOF THEN
		        Evn_RS.MoveFirst
		            Do While NOT Evn_RS.EOF		        
		                ENVIRONMENT_ABBREVIATION = Evn_RS("ENVIRONMENT_ABBREVIATION")
		                Evn_RS.MoveNext
		            Loop
	        END IF	
	        If 	ENVIRONMENT_ABBREVIATION <> "" Then	    
	            Session("ENVIRONMENT_ABBREVIATION") = ENVIRONMENT_ABBREVIATION
	        End If	
			
			If in_array(fnsEnvironments,environment) Then
				Session("isAsp")=false
			Else
				Session("isAsp")=true
			End If
			
			Evn_RS.close
			set Evn_RS = nothing
			Connect.Close
			set Connect = nothing
			
			Response.Redirect "Main-f.asp"
		Else
			Response.redirect "login.asp?Error=IP"
		End If
	Else
		Response.redirect "login.asp?Error=NOCONN"
	End If

Else
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>FNSNet Designer Login</title>
<link href="designer_login.css" rel="stylesheet" type="text/css">
<script LANGUAGE="javascript">
<!--
function checkInstallStatus()
{

	var objTreeCtrl =		document.all("TreeCtrl");
	var objClipboardCtrl =	document.all("ClipboardCtrl");
	var objLayoutCtrl =		document.all("LayoutCtrl");

	if ((objTreeCtrl == null) || (objTreeCtrl.readyState != 4))
	{
		STATUS.innerHTML = "<BR><B>The installation of the FNSNet Designer controls failed. Please contact your application administrator. (error on DynTree control)</B><BR><BR><A href='Troubleshoot.htm'>Click here</A> for toubleshooting instructions.";
		return false;
	}

	if ((objClipboardCtrl == null) || (objClipboardCtrl.readyState != 4))
	{
		STATUS.innerHTML = "<BR><B>The installation of the FNSNet Designer controls failed. Please contact your application administrator. (error on FNSClipboard control)</B><BR><BR><A href='Troubleshoot.htm'>Click here</A> for toubleshooting instructions.";
		return false;
	}

	if ((objLayoutCtrl == null) || (objLayoutCtrl.readyState != 4))
	{
		STATUS.innerHTML = "<BR><B>The installation of the FNSNet Designer controls failed. Please contact your application administrator. (error on FNS Layout control)</B><BR><BR><A href='Troubleshoot.htm'>Click here</A> for toubleshooting instructions.";
		return false;
	}
	else
	{
		LOGINVIEW.style.visibility = "visible";
		STATUS.innerHTML = "";
		return true;
	}
}
-->
</script>

<script LANGUAGE="vbscript">
<!--
dim bToggleConnect
bToggleConnect = true

Sub Btnlogin_onclick
document.all.environment.value = document.all.ConnectString.options(document.all.ConnectString.SelectedIndex).text
If document.all.UNAME.Value = "" OR document.all.PWD.Value = "" Then
	msgbox "User name and pasword are required." ,48 , "Designer login"
Else
	FrmLogin.Submit()
End If
End Sub

Sub btnClear_OnClick
	document.all.UNAME.Value = ""
	document.all.PWD.Value = ""
End Sub

Sub document_onkeydown
		select case window.event.keyCode
			case 13
				call Btnlogin_onclick
			case else:
		end select
End Sub
-->
</script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
if (window.dialogArguments != null)
{
	STATUS.innerHTML = "Designer Session has timed out!"
	alert ("This Designer Session has timed out, please exit the designer and login again")
	window.close()
}

if (top.frames.length > 0)
{
	top.frames.location.href = "login.asp"
}

	top.window.document.title = "FNSNet Designer"
	checkInstallStatus()
	document.all.PWD.focus()
<% If Request.Cookies("UserDB") <> "" Then %>document.all.ConnectString.value = "<%= Request.Cookies("UserDB") %>"
<% End If %>

}
//-->
</SCRIPT>
</head>
<body LANGUAGE=javascript onload="return window_onload()" bgcolor="#018B98">
<span ID="STATUS" CLASS=LABEL>Verifying / Installing FNSNet Designer components... Please wait.</span>
<span ID="OFFSCREEN" CLASS=LABEL style="position: absolute; left:-999; top:-999;">

<object classid="CLSID:DACBC157-5101-11D3-80B6-009027139D85" codeBase="./lib/FNSDControls.CAB#version=1,0,0,22" h id="TreeCtrl" style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabindex="-1" eight="1" width="1" VIEWASTEXT></object>

<object ID="ClipboardCtrl" tabindex="-1" WIDTH="1" HEIGHT="1" STYLE="visibility: hidden" CLASSID="CLSID:5A4655E0-DAFB-11D2-AFC8-0060082408D7" CODEBASE="./lib/FNSDControls.cab#Version=1,0,0,4" VIEWASTEXT>
</object>

<object ID="LayoutCtrl" tabindex="-1" WIDTH="1" HEIGHT="1" CLASSID="CLSID:13A2FD71-9ECC-11D2-AF73-0060082408D7" STYLE="visibility: hidden" CODEBASE="./lib/FNSDControls.cab#Version=1,1,1,0" VIEWASTEXT>
</object>
</span>
<div ID="LOGINVIEW" style="visibility: hidden;">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="95%" ID="Table3">
	<tr>
		<td align="center" valign="middle">
		<table width="538" border="0" cellspacing="0" cellpadding="0" ID="Table4">
			<tr>
				<td>
				<table border="0" cellspacing="0" cellpadding="0" ID="Table5">
					<tr valign="top">
						<td width="230" align="left">
						<img src="images/designer.gif" width="162" height="47"></td>
						<td bgcolor="#000000" width="1">
						<img src="images/spacer.gif" width="1" height="1"></td>
						<td width="307" align="right" valign="center">
						<span style='font-weight:bold;font-size:16px'>Innovation First Notice</span>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#000000">
				<img src="images/spacer.gif" width="1" height="1"></td>
			</tr>
			<tr>
				<td>
				<table border="0" cellspacing="0" cellpadding="0" ID="Table6">
					<tr valign="top">
						<td width="230" class="brandedColor1">
						<img src="images/spacer.gif" width="230" height="5"></td>
						<td bgcolor="#000000" width="1">
						<img src="images/spacer.gif" width="1" height="1"></td>
						<td width="307" class="brandedColor2">
						<img src="images/spacer.gif" width="307" height="5"></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#000000">
				<img src="images/spacer.gif" width="1" height="1"></td>
			</tr>
			<tr>
				<td align="left">
				<table border="0" cellspacing="0" cellpadding="0" ID="Table7">
					<tr>
						<td width="230" bgcolor="#532901" align="left" valign="bottom">
						<img src="images/orange_grad_login2.jpg" width="230" height="229"></td>
						<td bgcolor="#000000" width="1">
						<img src="images/spacer.gif" width="1" height="1"></td>
						<td width="307" align="right" valign="bottom" bgcolor="#018B98">
						<form NAME="FrmLogin" ACTION="login.asp?ACTION=LOGIN" METHOD="POST">
							<table border="0" cellspacing="3" cellpadding="0" width="95%" ID="Table8">
								<tr>
									<td>
									<table width="100%" border="0" cellspacing="2" cellpadding="0" ID="Table9">
										<tr>
											<td>User Name</td>
										</tr>
										<tr>
											<td>
											<input type="text" style="width:185" CLASS="fillin" name="UNAME" VALUE="<%= Request.Cookies("UserName")%>" maxlength="15" tabindex="1" ID="Text2">
											</td>
										</tr>
									</table>
									</td>
								</tr>
								<tr>
									<td>
									<table width="100%" border="0" cellspacing="2" cellpadding="0" ID="Table10">
										<tr>
											<td>Password</td>
										</tr>
										<tr>
											<td>
											<input type="password" style="width:185" CLASS="fillin" name="PWD" maxlength="15" tabindex="2" ID="Password2">
											</td>
										</tr>
									</table>
									</td>
								</tr>
								<tr>
									<td>
									<table width="100%" border="0" cellspacing="2" cellpadding="0" ID="Table11">
										<tr>
											<td>Database</td>
										</tr>
										<tr>
											<td>
											<select style="width:130"  class="fillin" id="Select1" name="ConnectString" tabindex="3">
											<%=Application("CONNECT_OPTIONS")%>
											</select>&nbsp;&nbsp;
											<input type="hidden" name ="environment" value="" />
											</td>
										</tr>
									</table>
									</td>
								</tr>
								<tr>
									<td>
									<table width="100%" border="0" cellpadding="0" cellspacing="2" ID="Table12">
										<tr>
											<td colspan="2">
											<img src="images/spacer.gif" width="1" height="1"></td>
										</tr>
										<tr>
											<td colspan="2">
											<hr size="1" noshade></td>
										</tr>
										<tr>
											<td colspan="2">
											<img src="images/spacer.gif" width="1" height="1"></td>
										</tr>
									</table>
									<table width="100%" border="0" cellpadding="0" cellspacing="3" ID="Table13">
										<tr align="left">
											<td>
											<input type="button" name="Btnlogin" value="Log On" class="button" tabindex="6">&nbsp;&nbsp;
											<input type="button" name="btnClear" value="Clear" class="button" tabindex="7">
											</td>
										</tr>
									</table>
									</td>
								</tr>
							</table>
						</form>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#000000">
				<img src="images/spacer.gif" width="1" height="1"></td>
			</tr>
		</table>
		<table width="538" border="0" cellspacing="0" cellpadding="0" ID="Table14">
			<tr>
				<td><img src="images/spacer.gif" width="20" height="20"></td>
			</tr>
			<tr>
				<td class="copyright">Copyright 1998-<%=YEAR(Now)%>, First Notice Systems, Inc. - All rights reserved. Version 2.75</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<% If Request.QueryString("Error") = "IP" Then %>
<center>
	<label CLASS="LABEL"><font COLOR="MAROON">User name or password not found.</font></label>
</center>
<% ElseIf Request.QueryString("Error") = "CONNSTR" Then%>
<center>
	<label CLASS="LABEL"><font COLOR="MAROON">Invalid connection string.</font></label>
</center>
<% ElseIf Request.QueryString("Error") = "NOCONN" Then%>
<center>
	<label CLASS="LABEL"><font COLOR="MAROON">Unable to establish a connection.</font></label>
</center>
<% End If %>
</body>
</span>
</html>
<% End If %>