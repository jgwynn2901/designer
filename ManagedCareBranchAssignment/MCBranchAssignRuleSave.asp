<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	On Error Resume Next
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")		
	
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "MC_BRANCHASSIGNMENTRULE", "MC_BRANCHASSIGNMENTRULE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Managed Care Branch Assignment Rule " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewMCBARID = CLng(NextPkey("MC_BRANCHASSIGNMENTRULE","MC_BRANCHASSIGNMENTRULE_ID"))
		If NewMCBARID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "MC_BRANCHASSIGNMENTRULE", "MC_BRANCHASSIGNMENTRULE_ID", NewMCBARID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Managed Care Branch Assignment Rule " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateMCBARID('" & NewMCBARID &  "');")
		Else
			strError = "Unable to obtain next primary key for MC_BRANCHASSIGNMENTRULE table."
		End If			
			
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "MC_BRANCHASSIGNMENTRULE", "MC_BRANCHASSIGNMENTRULE_ID", SQL_STRING)		 
		Set RSUpdate = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"Managed Care Branch Assignment Rule " & ACTION)
	End If 

	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "MC_BRANCHASSIGNMENTRULE", "", 0, ""
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames("WORKAREA").SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);		
<%		
	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("Update successful.");
		parent.frames("WORKAREA").ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);				
<%		
	End If

 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
