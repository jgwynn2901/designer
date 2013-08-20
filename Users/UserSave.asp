<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT LANGUAGE="vbscript">
<%
	dim cAction, cSQL, cSQLString, cUserID, cError, cName, oRS, cReuse
	dim aRecycle(4), y, lNewKey

	On Error Resume Next
	cError = ""
	cAction = CStr(Request.Form("TxtAction"))
	cSQLString = Request.Form("TxtSaveData")
	cReuse = Request.Form("TxtReuse")
	cUserID = Request.Form("UID")
	aRecycle(0) = "Delete From ACCESSPERMISSIONS Where USER_ID = " & cUserID
	aRecycle(1) = "Delete From USER_GROUP Where USER_ID = " & cUserID
	aRecycle(2) = "Delete From SETTING Where USER_ID = " & cUserID
	aRecycle(3) = "Delete From ACCOUNT_USER Where USER_ID = " & cUserID
	aRecycle(4) = "Delete From SECURITY_LOG Where USER_ID = " & cUserID
	If cAction = "UPDATE" Then
		cSQL = BuildSQL(cSQLString,Chr(128), Chr(129), "UPDATE", "USERS", "USER_ID", "")
		Conn.BeginTrans
		Conn.Execute(cSQL)
		cError = CheckADOErrors(Conn, "Users " & ACTION)
		if cReuse = "Y" then
			y = 0
			do while len(cError) = 0 and y <= Ubound(aRecycle)
				Conn.Execute( aRecycle( y ) )
				cError = CheckADOErrors(Conn, "Users " & ACTION)
				y = y + 1
			loop
		end if
		If len(cError) <> 0 Then
			Conn.RollbackTrans
		else
			Conn.CommitTrans
		end if
	Elseif cAction = "INSERT" Then
		'	recycling has been disabled because USER_IDs are stored in calls
		'	try to find a recycled record
		'cSQL = "Select USER_ID From USERS Where REUSE = 'Y'"
		'Set oRS = Conn.Execute(cSQL)
		'if not oRS.eof then
			'nUserID = CLng(oRS.Fields("USER_ID").value)
			'lNewKey = false
		'else
			nUserID = CLng(NextPkey("USERS","USER_ID"))
			lNewKey = true
		'end if
		'oRS.close
		If nUserID > 0 Then
			if lNewKey then
				cSQL = BuildSQL(cSQLString, Chr(128), Chr(129), "INSERT", "USERS", "USER_ID", nUserID)
			else
				cSQL = BuildSQL(cSQLString, Chr(128), Chr(129), "UPDATE", "USERS", "USER_ID", nUserID)
			end if
			Conn.Execute(cSQL)
			cError = CheckADOErrors(Conn,"Users " & ACTION)
			If cError = "" Then 
				%>
				parent.frames("WORKAREA").UpdateUID("<%=nUserID%>")
				parent.frames("WORKAREA").enableTabs
				parent.frames("WORKAREA").lockUserName
				<%	
			end if
		Else
			cError = "Unable to obtain next primary key for USERS table."
		End If
	End If
	Conn.Close
	
	If cError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, cError, "USERS", "", 0, ""
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
%>
</SCRIPT>
</HEAD>
<BODY>
</BODY>
