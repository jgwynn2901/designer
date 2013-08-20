<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	On Error Resume Next

	ACTION = Request.Form("TxtAction")
	NID = Request.Form("NID")
	SQL_STRING = Request.Form("TxtSaveData")		
	
	select case ACTION
		case "UPDATE" 
			UpdateSQL = ""
			UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "NETWORK", "NETWORK_ID", "")		 
			Set RSinsert = Conn.Execute(UpdateSQL)
			strError = CheckADOErrors(Conn,"Network " & ACTION)
		case "INSERT"
			InsertSQL = ""
			NewNID = CLng(NextPkey("NETWORK","NETWORK_ID"))
			If NewNID > 0 Then
				InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "NETWORK", "NETWORK_ID", NewNID)		 
				Set RSUpdate = Conn.Execute(InsertSQL)
				strError = CheckADOErrors(Conn,"Network " & ACTION)
				If strError = "" Then 
					Response.write("parent.frames('WORKAREA').UpdateNID('" & NewNID &  "');")
				end if
			Else
				strError = "Unable to obtain next primary key for VENDOR table."
			End If
		case "ADD VENDOR"
			cSQL = "Insert into VENDOR_NETWORK values (" & NID & "," & SQL_STRING & ")"
			Conn.Execute(cSQL)
			strError = CheckADOErrors(Conn,"Network " & ACTION)
		case "DELETE VENDOR"
			cSQL = "Delete from VENDOR_NETWORK where NETWORK_ID = " & NID & " and VENDOR_ID = " & SQL_STRING
			Conn.Execute(cSQL)
			strError = CheckADOErrors(Conn,"Network " & ACTION)
	end select
	Conn.Close
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "NETWORK", "", 0, ""
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
<%	End If %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
