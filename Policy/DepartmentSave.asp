<%
'***************************************************************
'generates insert or update query for Department.
'
'$History: DepartmentSave.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 1:39p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Policy
'* Added Department Interface due to ESIS Project.  It allows User to
'* create Department record attached to the AHSID in PROD Designer. The
'* permission used is the same as for Branch.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 12:10p
'* Created in $/FNS_DESIGNER/Source/Designer/Policy
'* Added Department Interface due to the ESIS Project.  It allows user to
'* attach AHSID to the department record.  Also, it allows user to delete,
'* create a new record and Edit an record in PROD Designer.  Permission
'* setup is the same as for Branch.  


'***************************************************************
%>
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
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "DEPARTMENT_CODES", "DEPARTMENT_CODES_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"DEPARTMENT" & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewDEPTID = CLng(NextPkey("DEPARTMENT_CODES","DEPARTMENT_CODES_ID"))
		If NewDEPTID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "DEPARTMENT_CODES", "DEPARTMENT_CODES_ID", NewDEPTID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"DEPARTMENT_CODES" & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateDEPTID('" & NewDEPTID &  "');")	
		Else
			strError = "Unable to obtain next primary key for DEPARTMENT_CODES table."
		End If			

	End If

	Conn.Close
	
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "DEPARTMENT_CODES", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames('WORKAREA').SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').UpdateStatus('Update successful.');
		parent.frames('WORKAREA').ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		
<%	End If

 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
