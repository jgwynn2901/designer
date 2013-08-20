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

	If ACTION = "INSERT" Then			
		InsertSQL = ""
		InsertSQL = "INSERT INTO JURISDICTION_STATE (POLICY_ID, STATE) VALUES " & SQL_STRING
		Set RSUpdate = Conn.Execute(InsertSQL)
				
	ElseIf ACTION = "DELETE" Then
		DeleteSQL = "DELETE FROM JURISDICTION_STATE WHERE " & SQL_STRING
		Set RSUpdate = Conn.Execute(DeleteSQL)
		
	ElseIf ACTION = "INSERT_ALL" Then
		Conn.beginTrans
		Dim PID 		
		PID = Request.Form("PID")
		
		If PID <> "" Then
			//Remove all previous Jurisdiction States for selected Policy
			DeleteSQL = "DELETE FROM JURISDICTION_STATE Where POLICY_ID = " & PID
		End If
		
		Conn.Execute(DeleteSQL)
		
		 //Insert newly selected Juristic states
		InsertSQL = "INSERT ALL " & SQL_STRING & " SELECT count(*) FROM dual"					
		
		Set RSUpdate = Conn.Execute(InsertSQL)
		
		If Conn.Errors.count > 0 Then
			Conn.RollbackTrans
		Else
			Conn.CommitTrans
		End If
		
	End If 
	
	

	strError = CheckADOErrors(Conn,"Jurisdiction State " & ACTION)
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "JURISDICTION_STATE", "", 0, ""
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
