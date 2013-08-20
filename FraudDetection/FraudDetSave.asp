<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT language=vbscript>
<%
	On Error Resume Next
	ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")	
	
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "FRAUD_DETECTION_TYPE", "FRAUD_DETECTION_TYPE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Fraud Detection Type " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		cAHSID = Request.Form("AHSID")
		cLOB = Request.Form("LOB_CD")
	    cSQL = "SELECT * FROM FRAUD_DETECTION_TYPE FDT " 
		cSQL = cSQL & "WHERE ACCNT_HRCY_STEP_ID = " & cAHSID 
		cSQL = cSQL & " AND LOB_CD = '" & cLOB & "'"
		set oRS = Conn.Execute(cSQL)
		lDuplicate = false
		if not oRS.eof then
			Response.write "msgbox ""The AHSID+LOB combination must be unique."", 48, ""FNSDesigner"""
			lDuplicate = true
		else
			NewFDTID = CLng(NextPkey("FRAUD_DETECTION_TYPE","FRAUD_DETECTION_TYPE_ID"))
			If NewFDTID > 0 Then
				InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "FRAUD_DETECTION_TYPE", "FRAUD_DETECTION_TYPE_ID", NewFDTID)
				Set RSUpdate = Conn.Execute(InsertSQL)
			
				strError = CheckADOErrors(Conn,"Fraud Detection Type " & ACTION)
				If strError = "" Then Response.write("parent.frames(""WORKAREA"").UpdateFDTID """ & NewFDTID &  """")	
			Else
				strError = "Unable to obtain next primary key for FRAUD_DETECTION_TYPE table."
			End If
		end if
	End If 
	Conn.Close
	set Conn = nothing
	set oRS = nothing
	if not lDuplicate then
		If strError <> "" Then	
			LogStatusGroupBegin
			LogStatus S_ERROR, strError, "FRAUD_DETECTION_TYPE", "", 0, ""
			LogStatusGroupEnd %>
			parent.frames("WORKAREA").UpdateStatus "<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report."
			parent.frames("WORKAREA").SetDirty
			parent.frames("WORKAREA").SetStatusInfoAvailableFlag true
	<%	Else
			LogStatusGroupBegin
			LogStatusGroupEnd %>
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
