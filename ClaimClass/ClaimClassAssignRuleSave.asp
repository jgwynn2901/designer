<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%

'***************************************************************
'General purpose: Builds Update or insert stmt. for claim_class_assignment table
'
'$History: ClaimClassAssignRuleSave.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/10/06    Time: 10:59p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/ClaimClass
'* New Claim Class Assignment module: Search, Details etc.



	On Error Resume Next
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")	
	
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "CLAIM_CLASS_ASSIGNMENT", "CLAIM_CLASS_ASSIGNMENT_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Claim Class Assignment Rule " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewCARID = CLng(NextPkey("CLAIM_CLASS_ASSIGNMENT","CLAIM_CLASS_ASSIGNMENT_ID"))

		If NewCARID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "CLAIM_CLASS_ASSIGNMENT", "CLAIM_CLASS_ASSIGNMENT_ID", NewCARID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"Claim Class Assignment Rule " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateCARID('" & NewCARID &  "');")	
		Else
			strError = "Unable to obtain next primary key for CLAIM_CLASS_ASSIGNMENT table."
		End If			
		
	End If
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "CLAIM_CLASS_ASSIGNMENT", "", 0, ""
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames("WORKAREA").SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);		
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames("WORKAREA").UpdateStatus("Update successful.");
		parent.frames("WORKAREA").ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);
<%	End If
 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
