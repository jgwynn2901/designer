<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<script type="text/vbscript" language="vbscript">
<%
	On Error Resume Next
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")	
	
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "ACCOUNT_TIP", "ACCOUNT_TIP_ID", "")		 		
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Account Tip " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		cAHSID = Request.Form("AHSID")
		cLOB = Request.Form("LOB_CD")
	    cSQL = "SELECT * FROM ACCOUNT_TIP AT " 
		cSQL = cSQL & "WHERE ACCNT_HRCY_STEP_ID = " & cAHSID 
		cSQL = cSQL & " AND LOB_CD = '" & cLOB & "'"
		set oRS = Conn.Execute(cSQL)
		lDuplicate = false
		if not oRS.eof then
			Response.write "msgbox ""The AHSID+LOB combination must be unique."", 48, ""FNSDesigner"""
			lDuplicate = true
		else
			NewATID = CLng(NextPkey("ACCOUNT_TIP","ACCOUNT_TIP_ID"))
			If NewATID > 0 Then
				InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "ACCOUNT_TIP", "ACCOUNT_TIP_ID", NewATID)
				'Response.Write InsertSQL
				'Response.End
				Set RSUpdate = Conn.Execute(InsertSQL)
				
				strError = CheckADOErrors(Conn,"Account Tip " & ACTION)
				If strError = "" Then Response.write("parent.frames(""WORKAREA"").UpdateATID """ & NewATID &  """")	
			Else
				strError = "Unable to obtain next primary key for ACCOUNT_TIP table."
			End If
		end if
	End If 
	Conn.Close
	set Conn = nothing
	set oRS = nothing
	if not lDuplicate then
		If strError <> "" Then	
			LogStatusGroupBegin
			LogStatus S_ERROR, strError, "ACCOUNT_TIP", "", 0, ""
			LogStatusGroupEnd %>
			parent.frames("WORKAREA").UpdateStatus "<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report."
			'parent.frames("WORKAREA").UpdateStatus strError
			parent.frames("WORKAREA").SetDirty
			parent.frames("WORKAREA").SetStatusInfoAvailableFlag true
	<%	Else
			LogStatusGroupBegin
			LogStatusGroupEnd %>
			'parent.frames("WORKAREA").SetSuccessfulQueryCount 1 
			parent.frames("WORKAREA").UpdateStatus "Update successful."
			parent.frames("WORKAREA").ClearDirty
			parent.frames("WORKAREA").SetStatusInfoAvailableFlag false
	<%	End If
	end if
 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
