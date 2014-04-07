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
		 lret = SplitString(cSQLString, Chr(128))
	    Set oComm=Server.CreateObject("ADODB.Command")
        Set oComm.ActiveConnection=Conn
        oComm.commandtype= 4	'	adCmdStoredProc 
        oComm.commandtext="SP_UPDATE_USER"
        oComm.Parameters.Refresh ' Fetching parameters list from SP. The order of parameters and their values must be synced
        nUserID = SplitString(lret(1), Chr(129))(2)
        oComm.Parameters(0) = nUserID		
        for i = 2 to UBOUND(lret) step 1
            If lret(i) <> "" Then
                lret2 = SplitString(lret(i), Chr(129))
                oComm.Parameters(i-1) = lret2(2)                
            End If
        Next
        oComm.execute()            	        
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
	     nUserID = CLng(NextPkey("USERS","USER_ID"))
		 lNewKey = true		
		 If nUserID > 0 Then
	        lret = SplitString(cSQLString, Chr(128))
		    Set oComm=Server.CreateObject("ADODB.Command")
            Set oComm.ActiveConnection=Conn
            oComm.commandtype= 4	'	adCmdStoredProc 
            oComm.commandtext="SP_INSERT_USER"
            oComm.Parameters.Refresh ' Fetching parameters list from SP. The order of parameters and their values must be synced
            oComm.Parameters(0) = nUserID			
            for i = 2 to UBOUND(lret) step 1
                If lret(i) <> "" Then
                    lret2 = SplitString(lret(i), Chr(129))
                    oComm.Parameters(i-1) = lret2(2)
                End If
            Next                     
            oComm.execute()            
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
