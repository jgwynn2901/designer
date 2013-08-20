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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "LU_CODE", "LU_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Lookup Code " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewLUCID = CLng(NextPkey("LU_CODE","LU_ID"))
		If NewLUCID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "LU_CODE", "LU_ID", NewLUCID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Lookup Code " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateLUCID('" & NewLUCID &  "');")	
		Else
			strError = "Unable to obtain next primary key for LU_CODE table."
		End If				
		
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "LU_CODE", "LU_ID", SQL_STRING)		 
		Set RSUpdate = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"Lookup Code " & ACTION)
	End If 

	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "LU_CODE", "", 0, ""
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
