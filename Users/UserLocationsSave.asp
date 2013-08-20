<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT LANGUAGE="vbscript">
<%
dim cAction, cSQL, cSQLString, nUserID, cError, cName, oRS, lDuplicate

	On Error Resume Next
	cError = ""
	lDuplicate = false
	cAction = CStr(Request.Form("TxtAction"))
	cSQLString = Request.Form("TxtSaveData")	

	If cAction = "UPDATE" Then
		cSQL = BuildSQL(cSQLString,Chr(128), Chr(129), "UPDATE", "LOCATIONS_USER", "LOCATION_USER_ID", "")		 
		Set oRS = Conn.Execute(cSQL)
		cError = CheckADOErrors(Conn,"Users " & ACTION)
	Elseif cAction = "INSERT" Then
		'cName = uCase(Trim(Request.Form("txtName")))
		'cSQL = "Select * from Users where Upper(Name) = '" & cName & "'"
		'Set oRS = Conn.Execute(cSQL)
		'if not oRS.eof then
			'lDuplicate = true
			'oRS.close
			'%>
			'msgbox "User '" & "<%=cName%>" & "' already exists.", vbExclamation, "FNSDesigner"
			'<%
		'else
			'oRS.close
			nUserID = CLng(NextPkey("LOCATIONS_USER","LOCATION_USER_ID"))
			If nUserID > 0 Then
				cSQL = BuildSQL(cSQLString, Chr(128), Chr(129), "INSERT", "LOCATIONS_USER", "LOCATION_USER_ID", nUserID)		 
				Conn.Execute(cSQL)
				cError = CheckADOErrors(Conn,"LOCATIONS_USER " & ACTION)
				If cError = "" Then 
					%>
					parent.frames("WORKAREA").UpdateLUD("<%=nUserID%>")
					<%	
				end if
			Else
				cError = "Unable to obtain next primary key for LOCATIONS_USER table."
			End If
		end if
	End If
	Conn.Close
	
if not lDuplicate then
	If cError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, cError, "LOCATIONS_USER", "", 0, ""
		LogStatusGroupEnd
	%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.")
		parent.frames("WORKAREA").SetDirty()
		parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true)		
	<%
	Else
		LogStatusGroupBegin
		LogStatusGroupEnd 
	%>
		parent.frames("WORKAREA").UpdateStatus("Update successful.")
		parent.frames("WORKAREA").ClearDirty()
		parent.frames("WORKAREA").SetStatusInfoAvailableFlag(false)
	<%
	End If
end if	
%>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
