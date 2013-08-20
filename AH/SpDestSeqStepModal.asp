<!-- #include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!-- #include file="..\lib\RenderTextinc.asp"-->
<!-- #include file="..\lib\security.inc"-->

<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	Dim ContainerType, Mode
	Dim s_SQL2, Conn, ConnectionString, rs, RS2, s_Select
	Dim s_SeqStepID,s_DestinationID,s_SeqNumber ,s_RetryCount ,s_RetryWaitTime ,s_DestinationStr,s_Transmission 
	Dim s_QAFax, s_QAPrinter, s_ProductionPrinter, s_QAEMAIL, s_QALEGACYEMAIL
	
	ContainerType = "Modal"
	Mode = Request.QueryString("MODE")

	If HasModifyPrivilege("FNSD_SPECIFIC_DESTINATION", SECURITYPRIV) <> True Then MODE = "RO"
	IF Request.QueryString("SeqStep_ID") <> "NEW" THEN
		SET Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = CONNECT_STRING
		s_Select = "Select * From SPECIFIC_DESTN_SEQ_STEP Where SPECIFIC_DESTN_SEQ_STEP_ID = " & Request.QueryString("SeqStep_ID")
		Conn.Open ConnectionString
		SET rs = Conn.Execute(s_Select)	
		s_SeqStepID = rs("SPECIFIC_DESTN_SEQ_STEP_ID")
		s_DestinationID = rs("SPECIFIC_DESTINATION_ID")
		s_SeqNumber = rs("SEQUENCE")
		s_RetryCount = rs("RETRY_COUNT")
		s_RetryWaitTime = rs("RETRY_WAIT_TIME")
		s_DestinationStr = ReplaceQuotesInText(rs("DESTINATION_STRING"))
		s_Transmission = rs("TRANSMISSION_TYPE_ID")
		rs.Close
		Conn.Close
		SET rs = NOTHING
		SET Conn = NOTHING
	
		Select Case s_Transmission
			Case "1"
				s_Transmission = "Fax"
			Case "2"
				s_Transmission = "Print"
			Case "3"
				s_Transmission = "EDI"
			Case "4"
				s_Transmission = "ICMS"
			Case "5"
				s_Transmission = "EDI UOF"
			Case "6"
				s_Transmission = "Email"
			Case "7"
				s_Transmission = "LegacyEmail"	
			Case "8"
				s_Transmission = "TIF FTP"	
			Case "9"
				s_Transmission = "EmailFile"	
			Case "10"
				s_Transmission = "EmailPDF"	
			Case "11"
				s_Transmission = "PDF FTP"	
			Case Else
				s_Transmission = ""
		End Select	
	ELSE
		s_SeqNumber = 1
		s_RetryCount = 3
		s_RetryWaitTime = 180
	END IF


%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Specifc Destination Sequence Step</title>
<link Rel="StyleSheet" Type="text/css" Href="..\FNSDESIGN.CSS">
<script Language="VBScript">
<!--#include file="..\lib\Help.asp"-->

Dim g_StatusInfoAvailable
	g_StatusInfoAvailable = false
	<% CALL getdefault() %>
	
	's_ProductionPrinter = "\\Cha0s00t\OPER_HP5SI_A"
	's_QAPrinter = "\\CHA0S2T\CHA4SI"
	's_QAFax = "6178862422"

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
	end if
end sub

Sub updateStatus(s_Msg)
	spanSeqStepID.innerHTML = s_Msg
End Sub

Function ExeSave
	s_SQL = ""
	ExeSave = false	
	If document.body.getAttribute("ScreenMode") = "RO" Then
			MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
			EXIT FUNCTION
	End If

	If Not f_ValidateScreenData Then EXIT FUNCTION
	s_Destination = ""
	s_AltDestination = ""
	If CStr(document.all.TxtTransmission.Value) = "1" Then
			s_AltDestination = "<%= s_QAFax %>"
	End If
	If CStr(document.all.TxtTransmission.Value) = "2" And document.all.TxtDestination.Value = "" Then
			s_Destination = "<%=s_ProductionPrinter%>"
			s_AltDestination = "<%=s_QAPrinter%>"
	ElseIf CStr(document.all.TxtTransmission.Value) = "2" And document.all.TxtDestination.Value <> "" Then
			s_Destination = document.all.TxtDestination.Value
			s_AltDestination = "<%=s_QAPrinter%>"
	End If
	If CStr(document.all.TxtTransmission.Value) = "7" Then
			s_AltDestination = "<%=s_QALEGACYEMAIL%>"
	ELSEIF CStr(document.all.TxtTransmission.Value) = "6" or _
			CStr(document.all.TxtTransmission.Value) = "9" or _
			CStr(document.all.TxtTransmission.Value) = "10" THEN
			s_AltDestination = "<%=s_QAEMAIL%>"
	End If
	IF document.all.spanSeqStepID.innerTEXT = "NEW" And document.all.spanSpDestID.innerTEXT <> "NEW" THEN
		s_SQL = s_SQL & "{call Designer.AddSpDestinationAndSeqStep(" & document.all.spanSpDestID.innerTEXT & ", "
		s_SQL = s_SQL & "NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, "
		s_SQL = s_SQL & document.all.TxtSeqNumber.value & ", " & f_RenderEntry(document.all.TxtRetryCount.value) & ", "
		s_SQL = s_SQL & f_RenderEntry(document.all.TxtRetryWaitTime.value) & ", " & f_RenderEntry(document.all.TxtDestination.Value) & ", " 
		s_SQL = S_SQL & f_RenderEntry(s_AltDestination) & ", " & document.all.TxtTransmission.value & ", "
		s_SQL = s_SQL & "{resultset 1, outNewDestination_ID, outNewSeqStep_ID, StatusMsg, StatusNum})}"
		document.all.Txt_SQLString.value = s_SQL
		document.all.Txt_Operation.value = "SaveNewSeqStep"
	ELSE	' Update
		s_SQL = s_SQL & "SPECIFIC_DESTN_SEQ_STEP_ID" & Chr(129) & document.all.spanSeqStepID.innerTEXT & Chr(129) & "0" & Chr(128)
		s_SQL = s_SQL & "SEQUENCE" & Chr(129) & document.all.TxtSeqNumber.value & Chr(129) & "0" & Chr(128)
		s_SQL = s_SQL & "RETRY_COUNT" & Chr(129) & f_RenderEntry(document.all.TxtRetryCount.value) & Chr(129) & "0" & Chr(128)
		s_SQL = s_SQL & "RETRY_WAIT_TIME" & Chr(129) & f_RenderEntry(document.all.TxtRetryWaitTime.value) & Chr(129) & "0" & Chr(128)
		s_SQL = s_SQL & "DESTINATION_STRING" & Chr(129) & document.all.TxtDestination.Value & Chr(129) & "1" & Chr(128)
		s_SQL = s_SQL & "ALT_DESTINATION_STRING" & Chr(129) & s_AltDestination & Chr(129) & "1" & Chr(128)
		s_SQL = s_SQL & "TRANSMISSION_TYPE_ID" & Chr(129) & document.all.TxtTransmission.Value & Chr(129) & "0" & Chr(128)
		document.all.Txt_SQLString.value = s_SQL
		document.all.Txt_Operation.value = "UPDATESeqStep"
		
	END IF

	document.all.frm_SeqStep.Target = "hiddenPage"
	document.body.setAttribute "ScreenDirty", "NO"
	document.all.frm_SeqStep.Submit()
	ExeSave = True
End Function
	
Sub BtnSaveSeqStep_OnClick
	b_Save = ExeSave
End Sub

Sub BtnNewSeqStep_OnClick
	document.all.spanSeqStepID.innerTEXT = "NEW"
	document.all.TxtSeqNumber.value = ""
	document.all.TxtRetryCount.value = ""
	document.all.TxtRetryWaitTime.value = ""
	document.all.TxtDestination.Value = ""
	document.all.TxtTransmission.value = ""
End Sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then 
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Function f_ValidateScreenData()
	f_ValidateScreenData = False
	s_ErrorMsg = ""
	If document.all.TxtSeqNumber.value = "" OR _
		Not IsNumeric(document.all.TxtSeqNumber.value) OR _
		Instr(1, document.all.TxtSeqNumber.value, "e") > 0 OR _
		Instr(1, document.all.TxtSeqNumber.value, ".") > 0 OR _
		Instr(1, document.all.TxtSeqNumber.value, "-") > 0 OR _
		document.all.TxtSeqNumber.value = "0" Then
			s_ErrorMsg = s_ErrorMsg & "Seq. Number is a Required Nermeric Field and Must be 1 or Higher Number." & VBCrLf
	End If
	If document.all.TxtRetryCount.value <> ""  Then
		if Not IsNumeric(document.all.TxtRetryCount.value) OR _
			Instr(1, document.all.TxtRetryCount.value, "e") > 0 OR _
			Instr(1, document.all.TxtRetryCount.value, ".") > 0 OR _
			Instr(1, document.all.TxtRetryCount.value, "-") > 0 OR _
			document.all.TxtRetryCount.value = "0" Then
			s_ErrorMsg = s_ErrorMsg & "Retry Count is a Numeric Field and Must be 1 or Higher Number." & VBCrLf
		end if
	End If
	If document.all.TxtRetryWaitTime.value <> "" Then
		if Not IsNumeric(document.all.TxtRetryWaitTime.value) OR _
			Instr(1, document.all.TxtRetryWaitTime.value, "e") > 0 OR _
			Instr(1, document.all.TxtRetryWaitTime.value, ".") > 0 OR _
			Instr(1, document.all.TxtRetryWaitTime.value, "-") > 0 OR _
			document.all.TxtRetryWaitTime.value = "0" Then
			s_ErrorMsg = s_ErrorMsg & "Retry Wait Time is a Nermeric Field and Must be 1 or Higher Number." & VBCrLf
		end if
	End If
	If document.all.TxtTransmission.value = "" Then
		s_ErrorMsg = s_ErrorMsg & "Transmission Type is a Required Field." & VBCrLf
	Else
		Select Case CStr(document.all.TxtTransmission.value)
			Case "1"
				IF document.all.TxtDestination.Value = "" OR _
					Not IsNumeric(document.all.TxtDestination.Value) OR _
					Instr(1, document.all.TxtDestination.Value, ".") > 0 OR _
					Instr(1, document.all.TxtDestination.Value, "e") > 0 OR _
					Instr(1, document.all.TxtDestination.Value, "-") > 0 OR _
					Len(document.all.TxtDestination.Value) < 10 Then
						s_ErrorMsg = s_ErrorMsg & "Please Provide a 10 Digit Fax Number in the Destination Field." & VBCrLf
				End If
			Case "2"
				IF document.all.TxtDestination.Value = "" Then document.all.TxtDestination.Value = "<%=s_ProductionPrinter%>"
				IF document.all.TxtDestination.Value <> "" And document.all.TxtDestination.Value <> "<%=s_ProductionPrinter%>" THEN
					If  MsgBox("The Default Printer Will Be Overrided by " & document.all.TxtDestination.Value & "." & VBCrLf & "YES to Continue, NO to Return to the Screen Panel?", vbYesNo, "FNSDesigner") = VBNo Then
							f_ValidateScreenData = False
							Exit Function
					End If
				End If	
		End Select
	End If
	If s_ErrorMsg <> "" Then
		MsgBox s_ErrorMsg, 0, "FNSDesigner"
		Exit Function					
	End If
	f_ValidateScreenData = True
End Function

Function f_RenderEntry(s_FieldValue)
		If s_FieldValue = "" Then
			f_RenderEntry = "NULL"
		ElseIf INSTR(1, s_FieldValue, "'") <> 0 OR INSTR(1, s_FieldValue, """") <> 0 Then
			f_RenderEntry = "'" & Replace(s_FieldValue, "'", "''") & "'"
		Else
			f_RenderEntry = "'" & s_FieldValue & "'"
		End If
End Function
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
  <tr><td colspan="2" HEIGHT="4"></td></tr>
  <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Specific Destination Sequence Step&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
  <td HEIGHT="5" ALIGN="LEFT">
							<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
									<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
									<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
										<td WIDTH="300" HEIGHT="8"></td></tr>
							</table></td></tr>
  
  <tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
  <tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
  <tr><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18"><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report"></td>
      <td WIDTH="385">:<span ID="SpanStatusSeqStep" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span></td></tr>
</table>
<br>
<% IF Request.QueryString("SDID") <> "NEW" THEN %>
<table class="LABEL">
  <tr><td width="305" nowrap>Specific Destination ID:&nbsp;<span ID="spanSpDestID" CLASS="LABEL"><%= Request.QueryString("SDID") %></span></td>
</table>
<br>
<% END IF %>
<span CLASS="LABEL">Specific Destination Sequence Step ID:&nbsp;<span id="spanSeqStepID"><%=Request.QueryString("SeqStep_ID")%></span>

<table CLASS="LABEL">
	<tr><td>	
		<table CLASS="LABEL" cellpadding="2" cellspacing="2">
			<tr><td CLASS="LABEL" width="100">Seq. Number:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtSeqNumber" VALUE="<%=s_SeqNumber%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
				<td CLASS="LABEL" width="100">Retry Count:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtRetryCount" VALUE="<%=s_RetryCount%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
				<td CLASS="LABEL" width="100">Retry Wait Time:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtRetryWaitTime" VALUE="<%=s_RetryWaitTime%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td></tr>
			<table CLASS="LABEL" cellpadding="2" cellspacing="2">
	  			<tr><td CLASS="LABEL">Destination:<br><input ScrnInput="TRUE" size="40" CLASS="LABEL" MAXLENGTH="255" TYPE="TEXT" NAME="TxtDestination" VALUE="<%=s_DestinationStr%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
				<td CLASS="LABEL">Transmission:<br><select ScrnBtn="TRUE" CLASS="LABEL" NAME="TxtTransmission" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"><%=GetControlDataHTML("TRANSMISSION_TYPE", "TRANSMISSION_TYPE_ID", "NAME", s_Transmission , true)%></select></td></tr>
			</table>
				<td VALIGN="TOP" ALIGN="RIGHT">
<% IF ContainerType <> "Modal" THEN
	If Request.QueryString("SDID") = "NEW" THEN %>
			<table ALIGN="RIGHT">	 
				<tr><td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE = "RO" Then Response.write(" DISABLED ") %> NAME="BtnSaveSeqStep" ACCESSKEY="S"><u>S</u>ave</button></tr>
				<tr><td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE = "RO" Then Response.write(" DISABLED ") %> <% If Request.QueryString("SDID") = "NEW" Then Response.write(" DISABLED ") %> NAME="BtnNewSeqStep" ACCESSKEY="N"><u>N</u>ew</button></tr>
				<tr><td CLASS="LABEL"><button CLASS="STDBUTTON" NAME="BtnCloseSeqStep" ACCESSKEY="L">C<u>l</u>ose</button></tr>
			</table>
	<% Else %>
			<table ALIGN="RIGHT">	 
				<tr><td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE = "RO" Then Response.write(" DISABLED ") %> NAME="BtnSaveSeqStep" ACCESSKEY="S"><u>S</u>ave</button></tr>
				<tr><td CLASS="LABEL"><button CLASS="STDBUTTON" NAME="BtnCloseSeqStep" ACCESSKEY="L">C<u>l</u>ose</button></tr>
			</table>	
	<% End If
END IF %>
					
		</table>
</table>
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<form Name="frm_SeqStep" ACTION="SpecificDestinationSave.asp" METHOD="POST" Target="hiddenPage">
	<input Type="Hidden" Name="Txt_SQLString">
	<input Type="Hidden" Name="Txt_Operation">
</form>
</body>
</html>
