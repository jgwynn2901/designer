<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%

'***************************************************
'  DMS: 2/17: Since the LOB_CD and ACCNT_HRCY_STEP_ID has been  
'             moved to AHS_POLICY table, 2 SQL's are required to update 
'             the information on the page.
'
'***************************************************

Dim ACTION, SQL_STRING1 , SQL_STRING2,SQL_STRING3, SQL_STRING4,UpdateSQL1, UpdateSQL2, UpdateSQL3
Dim RSinsert1, RSinsert2,RSinsert3, InsertSQL1, InsertSQL2,InsertSQL3, NewPID, RSUpdate, strError
Dim NewAHS_PID, RSUpdate1, RSUpdate2,RSUpdate3
Dim NewPolicy_Ext_PID 

    On Error Resume Next
	ACTION      = CStr(Request.Form("TxtAction"))
	SQL_STRING1 = Request.Form("TxtSaveData1")	
	SQL_STRING2 = Request.Form("TxtSaveData2")
	SQL_STRING3 = Request.Form("TxtSaveData3")	
	
	If ACTION = "UPDATE" Then
		UpdateSQL1   = ""
		UpdateSQL1   = BuildSQL(SQL_STRING1,Chr(128), Chr(129), "UPDATE", "POLICY", "POLICY_ID", "")		 
		
		Set RSinsert1 = Conn.Execute(UpdateSQL1)
		strError      = CheckADOErrors(Conn,"Policy " & ACTION)

        UpdateSQL2    = ""
        UpdateSQL2    = BuildSQL(SQL_STRING2,Chr(128), Chr(129), "UPDATE", "AHS_POLICY", "POLICY_ID", "")		 
		
		Set RSinsert2 = Conn.Execute(UpdateSQL2)
		strError      = CheckADOErrors(Conn,"Ahs_Policy " & ACTION)
		

		'**************************** RLOW-0169 Modification *****************************						
		SQL_STRING4 = "Select * from Policy_Extension where Policy_Id = " & Request.Form("PID")
		Set oRS = Conn.Execute(SQL_STRING4)
		Dim strCount
		strCount = 0
		Do While Not oRS.EOF
		strCount = strCount+1
		oRS.moveNext
		loop
		oRS.Close

		if strCount > 0 then
			'MMAI-0007
			'Prashant Shekhar			
			UpdateSQL3    = ""
			UpdateSQL3    = BuildSQL(SQL_STRING3,Chr(128), Chr(129), "UPDATE", "POLICY_EXTENSION", "POLICY_ID", "")		 

			Set RSinsert3 = Conn.Execute(UpdateSQL3)
			strError      = CheckADOErrors(Conn,"Policy_Extension " & ACTION)
		else
			 NewPolicy_Ext_PID = CLng(NextPkey("POLICY_EXTENSION","POLICY_EXTENSION_ID"))
			   
			If NewPolicy_Ext_PID > 0 Then
			   SQL_STRING3   = replace(SQL_STRING3, "NEW", Request.Form("PID"))
			   InsertSQL3    = BuildSQL(SQL_STRING3, Chr(128), Chr(129), "INSERT", "POLICY_EXTENSION", "POLICY_EXTENSION_ID", NewPolicy_Ext_PID)		 
			   Set RSUpdate3 = Conn.Execute(InsertSQL3)
			   strError      = CheckADOErrors(Conn,"Policy_Extension " & ACTION)
			Else
			   strError = strError & "Unable to obtain next primary key for POLICY_EXTENSION table."
			End if
		end if
		
       '**************************** RLOW-0169 Modification *****************************						 
	Elseif ACTION = "INSERT" Then

	'*****************************
	' DMS: 2/23/00 
	' 2 inserts are now required . The first insert is in the policy table
	' The second insert, which uses the PK generated for policy table,is the AHS_Policy tbl
	'*****************************
		
		InsertSQL1 = ""
		InsertSQL2 = ""
		InsertSQL3 = ""
		
		' Get the new PK for policy tbl.

		NewPID = CLng(NextPkey("POLICY","POLICY_ID"))
		If NewPID > 0 Then
			
			InsertSQL1    = BuildSQL(SQL_STRING1, Chr(128), Chr(129), "INSERT", "POLICY", "POLICY_ID", NewPID)		 
			Set RSUpdate1 = Conn.Execute(InsertSQL1)
			strError      = CheckADOErrors(Conn,"Policy " & ACTION)
			If strError   = "" then 
			   Response.write("parent.frames('WORKAREA').UpdatePID('" & NewPID &  "');")	
			   
			   NewAHS_PID = CLng(NextPkey("AHS_POLICY","AHS_POLICY_ID"))
			   
			   If NewAHS_PID > 0 Then
			      SQL_STRING2   = replace(SQL_STRING2, "NEW", NewPID)
				  InsertSQL2    = BuildSQL(SQL_STRING2, Chr(128), Chr(129), "INSERT", "AHS_POLICY", "AHS_POLICY_ID", NewAHS_PID)		 
				  Set RSUpdate2 = Conn.Execute(InsertSQL2)
		   	      strError      = CheckADOErrors(Conn,"Ahs_Policy " & ACTION)
			   Else
			      strError = strError & "Unable to obtain next primary key for AHS_POLICY table."
			   End if
			   
			   'MMAI-0007
			   'Prashant Shekhar
			   
			    NewPolicy_Ext_PID = CLng(NextPkey("POLICY_EXTENSION","POLICY_EXTENSION_ID"))
			   
			   If NewPolicy_Ext_PID > 0 Then
			      SQL_STRING3   = replace(SQL_STRING3, "NEW", NewPID)
				  InsertSQL3    = BuildSQL(SQL_STRING3, Chr(128), Chr(129), "INSERT", "POLICY_EXTENSION", "POLICY_EXTENSION_ID", NewPolicy_Ext_PID)		 
				  Set RSUpdate3 = Conn.Execute(InsertSQL3)
		   	      strError      = CheckADOErrors(Conn,"Policy_Extension " & ACTION)
			   Else
			      strError = strError & "Unable to obtain next primary key for POLICY_EXTENSION table."
			   End if
			   
			   
			End If
		Else
			strError = "Unable to obtain next primary key for POLICY table."
		End If				
		
	End If
	
	Conn.Close
	
	If strError <> "" Then
	    LogStatusGroupBegin
		LogStatus S_ERROR, strError, "POLICY", "", 0, ""
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
</html>
