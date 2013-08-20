<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT LANGUAGE="VBScript">
<%	
	ON ERROR RESUME NEXT
	ACTION = Request.Form("Txt_Operation")
	SQL_STRING = Request.Form("Txt_SQLString")
	Select Case ACTION
		Case "UPDATESpDestination"
			UpdateSQL = ""
			UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "SPECIFIC_DESTINATION", "SPECIFIC_DESTINATION_ID", "")		 
			Set rsSave = Conn.Execute(UpdateSQL)
			strError = CheckADOErrors(Conn,"SPECIFIC_DESTINATION " & ACTION)
			If strError <> "" Then
				LogStatusGroupBegin
				LogStatus S_ERROR, strError, "SpecificDestination", "", 0, ""
				LogStatusGroupEnd
			%>
				Parent.frames("WORKAREA").spanstatus.innerHTML = "<Font Color='Red'>Update Error! </Font>" & strError
			<%
			Else
			%>
				Parent.frames("WORKAREA").spanstatus.innerHTML = "Update Successfully"
			<%			
			End if
		Case "SaveNewBoth"
			Set rsSave = Conn.Execute(SQL_STRING)
			IF Conn.Errors.Count = 0 Then
				If rsSave("StatusNum") = "0" Then
			%>
					Parent.frames("WORKAREA").spanstatus.innerHTML = "Save Successfully"
					Parent.frames("WORKAREA").updateIFrameStatus("Save Successfully")
					Parent.frames("WORKAREA").span_SDID.innerHTML = "<%= rsSave("outNewDestination_ID") %>"
					Parent.frames("WORKAREA").SeqStepIFrame.document.all.spanSeqStepID.innerHTML = "<%= rsSave("outNewSeqStep_ID") %>"
			<%
					rsSave.close
				Else
			%>
					Parent.frames("WORKAREA").spanstatus.innerHTML = "<Font Color='Red'>Error: </Font>" & "<%= rs_Save("StatusMsg") %>"
					Parent.frames("WORKAREA").updateIFrameStatus("<Font Color='Red'>Save Error! </Font>")	
			<%
					rsSave.close
				End If
			ELSE
			%>
				Parent.frames("WORKAREA").spanstatus.innerHTML = "<Font Color='Red'>Error: </Font>" & "<%= Conn.Errors(1).Description %>"
				Parent.frames("WORKAREA").updateIFrameStatus("<Font Color='Red'>Save Error! </Font>")	
			<%
			End If			
		Case "SaveNewSeqStep"
			Set rsSave = Conn.Execute(SQL_STRING)
			IF Conn.Errors.Count = 0 THEN
				If rsSave("StatusNum") = "0" Then
				%>
					Parent.frames("WORKAREA").SpanStatusSeqStep.innerHTML = "Save Successfully"
					Parent.frames("WORKAREA").spanSeqStepID.innerHTML = "<%= rsSave("outNewSeqStep_ID") %>"
				<%
					rsSave.close
				Else
				%>
					Parent.frames("WORKAREA").SpanStatusSeqStep.innerHTML = "<Font Color='Red'>Error: </Font>" & "<%= rs_Save("StatusMsg") %>"
			<%
					rsSave.close
				End If
			Else
			%>
				Parent.frames("WORKAREA").SpanStatusSeqStep.innerHTML = "<Font Color='Red'>Error: </Font>" & "<%= Conn.Errors(1).Description %>"
			'''	Parent.frames("WORKAREA").spanSeqStepID.innerHTML = "<Font Color='Red'>Save Error! </Font>"
			<%
				rsSave.close			
			End If
		Case "UPDATESeqStep" 
			UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "SPECIFIC_DESTN_SEQ_STEP", "SPECIFIC_DESTN_SEQ_STEP_ID", "")		 
			Set rsSave = Conn.Execute(UpdateSQL)
			strError = CheckADOErrors(Conn,"SPECIFIC_DESTN_SEQ_STEP" & ACTION)
			If strError <> "" Then
				LogStatusGroupBegin
				LogStatus S_ERROR, strError, "SPECIFIC_DESTN_SEQ_STEP", "", 0, ""
				LogStatusGroupEnd
			%>
				Parent.frames("WORKAREA").document.all.SpanStatusSeqStep.innerHTML = "<Font Color='Red'>Update Error: </Font>" & strError
			<%
			Else
			%>
				Parent.frames("WORKAREA").document.all.SpanStatusSeqStep.innerHTML = "Update Successfully"
			<%			
			End if	
		Case "DELETE" 
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "SPECIFIC_DESTN_SEQ_STEP", "SPECIFIC_DESTN_SEQ_STEP_ID", SQL_STRING)		 
		Set rsSave = Conn.Execute(DeleteSQL)
		strError = CheckADOErrors(Conn,"SPECIFIC_DESTN_SEQ_STEP " & ACTION)
			If strError <> "" Then
				LogStatusGroupBegin
				LogStatus S_ERROR, strError, "SPECIFIC_DESTN_SEQ_STEP", "", 0, ""
				LogStatusGroupEnd
			%>
				Parent.frames("WORKAREA").spanstatus.innerHTML = "<Font Color='Red'>Delete Error!</Font>"
			<%
			Else
			%>
				Parent.frames("WORKAREA").spanstatus.innerHTML = "Delete Successfully"
			<%			
			End if
	
	END SELECT
	Conn.Close	
	
%>
</SCRIPT>
</HEAD>
</HTML>