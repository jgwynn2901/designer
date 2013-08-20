<%
	Response.Expires = 0
	Response.Buffer = true
%>

<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.UID.value = ""
	end if
End Sub

Function GetUID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetUID = document.all.tblFields.rows(idx).getAttribute("UID")
	Else
		GetUID = ""
	End If
End Function

Function GetUIDName
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("NAME").innerText
	End If
	GetUIDName = strText
End Function
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'95%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>User Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Site</div></td>
			<td class="thd"><div id><nobr>Active</div></td>
			<td class="thd"><div id><nobr>Creation Date</div></td>
			<td class="thd"><div id><nobr>Expiration Date</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
	dim RecCount
	RecCount = -1
	If Request.QueryString <> "" Then
	RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				UID = Request.QueryString("SearchUID") & "%"
				NAME = Request.QueryString("SearchName") & "%"
				'site is always exact
				SITE = Request.QueryString("SearchSite")
			Case "C"
				UID = "%" & Request.QueryString("SearchUID") & "%"
				NAME = "%" & Request.QueryString("SearchName") & "%"
				'site is always exact
				SITE = Request.QueryString("SearchSite")
			Case "E"
				UID = Request.QueryString("SearchUID")
				NAME = Request.QueryString("SearchName")
				'site is always exact
				SITE = Request.QueryString("SearchSite")
		End Select
		ACTIVE = Request.QueryString("SearchActive")
		If ACTIVE <> "" Then
			WHERECLS = WHERECLS & "UPPER(ACTIVE) = '" & UCASE(ACTIVE)  & "' AND "
		End If
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		else
			WHERECLS = WHERECLS & "NAME LIKE '%'"
		End If
		If Request.QueryString("SearchUID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "USER_ID LIKE '" & UID & "'"
		End If
		If Request.QueryString("SearchSite") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "SITE_ID = " & SITE 
		End If
		If Request.QueryString("SearchAHSID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID =" & Request.QueryString("SearchAHSID") 
		End If
		
		if WHERECLS <> "" then
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			WHERECLS = WHERECLS & " And REUSE = 'N'"
			If Request.QueryString("SearchAHSID") <> "" Then
				SQLST = "SELECT * FROM USERS_SITE_VIEW USV,ACCOUNT_USER AU WHERE "  & WHERECLS & " AND USV.USER_ID=AU.USER_ID ORDER BY NAME" 
			else			
				SQLST = "SELECT * FROM USERS_SITE_VIEW WHERE "  & WHERECLS & " ORDER BY NAME" 
			end if

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" UID=''>
	<td COLSPAN="6" NOWRAP CLASS="ResultCell" ID="USER_ID" align="center"><font size="2">No users found.</font></td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" UID="<%=RS("USER_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("USER_ID"))%></td>
<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("SITE_NAME"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACTIVE"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("PASSWORD_CREATION_DATE"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("PASSWORD_EXPIRATION_DATE"))%></td>


</tr>
<%
				RS.MoveNext
				Loop
			End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing

		End if
End If
%>

</tbody>
</table>
</div>
</fieldset>
<SCRIPT LANGUAGE="VBScript">
<%	If RecCount >= 0 Then %>
parent.parent.enableTab("Details")
parent.parent.enableTab("Groups")
parent.parent.enableTab("Permissions")
parent.parent.enableTab("Accounts")
parent.parent.enableTab("Locations")
multiselect document.all.tblFields.rows(1)
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=RecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=RecCount%>")
	end if		
end if
<%	End If %>
</SCRIPT>
</body>
</html>
