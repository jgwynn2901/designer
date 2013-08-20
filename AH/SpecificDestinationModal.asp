<!--#include file="..\lib\common.inc"-->
<!-- #include file="..\lib\security.inc"-->
<!-- #include file="..\lib\RenderTextinc.asp"-->
<!-- #include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RefCountRptinc.asp"-->
<!--#include file="..\lib\ZIP.inc"-->
<!--#include file="..\lib\ControlData.inc"-->
<%	
	If HasModifyPrivilege("FNSD_SPECIFIC_DESTINATION", SECURITYPRIV) <> True Then MODE = "RO"
	DIM s_iFrameSRC, Mode, s_ACCNTID
	Dim s_QAFax, s_QAPrinter, s_ProductionPrinter, s_QAEMAIL, s_QALEGACYEMAIL
	
	Mode = Request.QueryString("MODE")
	IF Request.QueryString("AHSID") = Request.QueryString("CLIENT_NODE") THEN
		s_ACCNTID = Request.QueryString("AHSID")
	ELSE
		If (Request.QueryString("CLIENT_NODE") = "") Then
			s_ACCNTID = Request.QueryString("AHSID")
		Else
			s_ACCNTID = Request.QueryString("CLIENT_NODE")
		End If
	END IF
	
	IF Request.QueryString("SDID") <> "" Then
		IF Request.QueryString("SDID") <> "NEW" Then
			s_iFrameSRC = "SpDestSeqStepSummary.asp?Mode=" & Mode & "&SPDEST_ID=" & Request.QueryString("SDID")
			SET Conn = Server.CreateObject("ADODB.Connection")
			ConnectionString = CONNECT_STRING
			s_SQL1 = "Select * From SPECIFIC_DESTINATION Where SPECIFIC_DESTINATION_ID = " & Request.QueryString("SDID")
			Conn.Open ConnectionString
			SET RS1 = Conn.Execute(s_SQL1)
			s_SpDestID = RS1("SPECIFIC_DESTINATION_ID")
			s_LastName = ReplaceQuotesInText(RS1("NAME_LAST"))
			s_FirstName = ReplaceQuotesInText(RS1("NAME_FIRST"))
			s_MI = RS1("NAME_MI")
			s_Title = ReplaceQuotesInText(RS1("TITLE"))
			s_Address1 = ReplaceQuotesInText(RS1("ADDRESS1"))
			s_Address2 = ReplaceQuotesInText(RS1("ADDRESS2"))
			s_City = ReplaceQuotesInText(RS1("CITY"))
			s_State = RS1("STATE")
			s_Zip = RS1("ZIP")
			s_Phone = RS1("PHONE")
			s_AHSID = RS1("ACCNT_HRCY_STEP_ID")
			s_OSHAE_FORM = RS1("ALTERNATE_FORM_FLG")
			s_LOB = RS1("LOB")
			RS1.CLOSE
			SET RS1 = NOTHING
		ELSE	' New Sp. Destination & Seq Step
			s_iFrameSRC = "SpDestSeqStepModal.asp?Mode=" & Mode & "&SDID=" & Request.QueryString("SDID") & "&SeqStep_ID=NEW&ACCNT_ID=" & Request.QueryString("AHSID") & "&Name=" & Request.QueryString("NAME") & "&CLIENT_NODE=" & Request.QueryString("CLIENT_NODE")
		END IF
	end if
 



	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Specific Destination</title>
<link Rel="StyleSheet" Type="text/css" Href="..\FNSDESIGN.CSS">
<script LANGUAGE="VBScript">
<!--#include file="..\lib\Help.asp"-->
Dim g_StatusInfoAvailable
	g_StatusInfoAvailable = false
	<% Call getdefault() %>
	's_ProductionPrinter = "\\Cha0s00t\OPER_HP5SI_A"
	's_QAPrinter = "\\Cha0s2t\CHA4SI"
	's_QAFax = "6178862422"

sub selectOption(objSelect, strValue)
dim i

for i=0 to objSelect.length-1
	if trim(strValue) = trim(objSelect(i).value) then
		objSelect(i).selected = true
		exit for
	end if
next
end sub

function getLOBs()
dim i, cAnswer

cAnswer = ""
for i=0 to document.all.TxtLOB.length-1
	if document.all.TxtLOB(i).selected then
		if len(cAnswer) = 0 then
			cAnswer = document.all.TxtLOB(i).value
		else
			cAnswer = cAnswer & ";" & document.all.TxtLOB(i).value
		end if
	end if
next
getLOBs = cAnswer
end function

Sub window_onload
dim aLOBs, x
<%IF Request.QueryString("SDID") <> "" Then%>
	document.all.STATE.Value = "<%=s_State%>"
<%end if%>	
if "<%=s_OSHAE_FORM%>" = "Y" then
	document.all.OSHAE_FORM.checked = true
end if
<%
if len(s_LOB) <> 0 then
%>
	aLOBs = split("<%=s_LOB%>", ";")
	for x=0 to ubound(aLOBs)
		selectOption document.all.TxtLOB, aLOBs(x)
	next
<%	
end if
%>
End Sub

sub Control_OnChange
	document.body.setAttribute "ScreenDirty", "YES"	
end sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then 
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null, "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If
End Sub

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Function ExeSave
dim cFlag
	s_SQL = ""
	ExeSave = false
	If document.body.getAttribute("ScreenMode") = "RO" Then
			MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
			EXIT FUNCTION
	End If
	document.body.setAttribute "ScreenDirty", "YES"
	IF Not f_ValidateScreenData Then EXIT Function
	Set o_IFrame = SeqStepIFrame.document.all
	s_Destination = ""
	s_AltDestination = ""

	If CStr(o_IFrame.TxtTransmission.Value) = "1" Then
			s_AltDestination = "<%=s_QAFax %>"
	End If
	If CStr(o_IFrame.TxtTransmission.Value) = "2" And o_IFrame.TxtDestination.Value = "" Then
			s_Destination = "<%=s_ProductionPrinter%>"
			s_AltDestination = "<%=s_QAPrinter%>"
	ElseIf CStr(o_IFrame.TxtTransmission.Value) = "2" And o_IFrame.TxtDestination.Value <> "" Then
			s_Destination = o_IFrame.TxtDestination.Value
			s_AltDestination = "<%=s_QAPrinter%>"
	End If
	
	If CStr(o_IFrame.TxtTransmission.Value) = "7" Then
			s_AltDestination = "<%=s_QALEGACYEMAIL%>"
	ELSEIF CStr(o_IFrame.TxtTransmission.Value) = "6" or _
			CStr(o_IFrame.TxtTransmission.Value) = "9" or _
			CStr(o_IFrame.TxtTransmission.Value) = "10" THEN
			s_AltDestination = "<%=s_QAEMAIL%>"
	End If
	If document.All.OSHAE_FORM.Checked Then
		cFlag = "'Y'"
	Else
		cFlag = "'N'"
	end if
	s_SQL = s_SQL & "{call Designer.AddSpDestinationAndSeqStep(NULL, " & document.all.Txt_ACCNTID.Value & ", "
	s_SQL = s_SQL & f_RenderEntry(document.all.TxtNAME_FIRST.Value) & ", " & f_RenderEntry(document.all.TxtNAME_LAST.Value) & ", " & f_RenderEntry(document.all.TxtNAME_INITIAL.Value) & ", " 
	s_SQL = s_SQL & f_RenderEntry(document.all.TxtTITLE.Value) & ", " & f_RenderEntry(document.all.TxtADDRESS1.Value) & ", " & f_RenderEntry(document.all.TxtADDRESS2.Value) & ", " & f_RenderEntry(document.all.CITY.Value) & ", " & f_RenderEntry(document.all.STATE.Value) & ", " & f_RenderEntry(document.all.ZIP.Value) & ", " & f_RenderEntry(document.all.TxtPHONE.Value) & ", '"  & getLOBs & "', " & cFlag & ","
	s_SQL = s_SQL & o_IFrame.TxtSeqNumber.value & ", " & f_RenderEntry(o_IFrame.TxtRetryCount.value) & ", "
	s_SQL = s_SQL & f_RenderEntry(o_IFrame.TxtRetryWaitTime.value) & ", " & f_RenderEntry(o_IFrame.TxtDestination.Value) & ", " 
	s_SQL = S_SQL & f_RenderEntry(s_AltDestination) & ", " & o_IFrame.TxtTransmission.value & ", "
	s_SQL = s_SQL & "{resultset 1, outNewDestination_ID, outNewSeqStep_ID, StatusMsg, StatusNum})}"
	document.all.Txt_SQLString.value = s_SQL
	document.all.Txt_Operation.value = "SaveNewBoth"
	document.all.frm_Entry.Target = "hiddenPage"
	document.all.frm_Entry.Submit()
	Set o_IFrame = Nothing
	ExeSave = TRUE
End Function

Sub BtnSaveSpDestination_OnClick()
	IF ACCNT_ID.innerText = "NEW" THEN EXIT SUB
	document.body.setAttribute "ScreenDirty", "YES"
	s_ErrorMsg = ""
	If document.all.TxtNAME_LAST.Value = "" Then
		s_ErrorMsg = s_ErrorMsg & "Last Name is a Required Field." & VBCrLf
	End if
	If document.all.TxtNAME_First.Value = "" Then
		s_ErrorMsg = s_ErrorMsg & "First Name is a Required Field." & VBCrLf
	End if
	IF s_ErrorMsg <> "" THEN
		Msgbox s_ErrorMsg, 0, "FNSDesigner"
		EXIT SUB
	END IF

	s_SQL = ""
	s_SQL = s_SQL & "SPECIFIC_DESTINATION_ID"  & Chr(129) & document.all.Txt_DestinationID.Value & Chr(129) & "0" & Chr(128)
	s_SQL = s_SQL & "NAME_FIRST" & Chr(129) & document.all.TxtNAME_FIRST.Value& Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "NAME_LAST" & Chr(129) & document.all.TxtNAME_LAST.Value & Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "NAME_MI" & Chr(129) & document.all.TxtNAME_INITIAL.Value & Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "TITLE" & Chr(129) & document.all.TxtTITLE.Value& Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "ADDRESS1" & Chr(129) & document.all.TxtADDRESS1.Value & Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "ADDRESS2" & Chr(129) & document.all.TxtADDRESS2.Value & Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "CITY" & Chr(129) & document.all.CITY.Value & Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "STATE" & Chr(129) & document.all.STATE.Value & Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "ZIP" & Chr(129) & document.all.ZIP.Value & Chr(129) & "1" & Chr(128)
	s_SQL = s_SQL & "PHONE" & Chr(129) & document.all.TxtPhone.Value & Chr(129) & "1" & Chr(128)
	If document.All.OSHAE_FORM.Checked Then
		s_SQL = s_SQL & "ALTERNATE_FORM_FLG" & Chr(129) & "Y" & Chr(129) & "1" & Chr(128)
	Else
		s_SQL = s_SQL & "ALTERNATE_FORM_FLG" & Chr(129) & "N" & Chr(129) & "1" & Chr(128)
	End If
	s_SQL = s_SQL & "LOB" & Chr(129) & getLOBs & Chr(129) & "1" & Chr(128)
	document.all.Txt_SQLString.value = s_SQL
	document.all.Txt_Operation.value = "UPDATESpDestination"
	document.all.frm_Entry.Target = "hiddenPage"
	document.all.frm_Entry.Submit()
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function f_ValidateScreenData
	f_ValidateScreenData = False
	s_ErrorMsg = ""
	If document.all.Txt_ACCNTID.value = "" Then
		s_ErrorMsg = s_ErrorMsg & "AHSID Is Not In Place." & VBCrlf
	End If	
	If document.all.TxtNAME_LAST.Value = "" Then
		s_ErrorMsg = s_ErrorMsg & "Last Name is a Required Field." & VBCrLf
	End if
	If document.all.TxtNAME_First.Value = "" Then
		s_ErrorMsg = s_ErrorMsg & "First Name is a Required Field." & VBCrLf
	End if
Set o_IFrame = SeqStepIFrame.document.all
	IF o_IFrame.TxtSeqNumber.value = "" OR _
	   o_IFrame.TxtSeqNumber.value = "0" OR _
	   Instr(1, o_IFrame.TxtSeqNumber.value, "e") > 0 OR _
	   Instr(1, o_IFrame.TxtSeqNumber.value, ".") > 0 OR _
	   Instr(1, o_IFrame.TxtSeqNumber.value, "-") > 0 OR _
	   Not IsNumeric(o_IFrame.TxtSeqNumber.value) Then
			s_ErrorMsg = s_ErrorMsg & "Sequence is a Required Numeric Field and Must be 1 or Higher Number." & VBCrLf
	END IF
	IF o_IFrame.TxtRetryCount.value <> "" Then
		if Not IsNumeric(o_IFrame.TxtRetryCount.value) OR _
			Instr(1, o_IFrame.TxtRetryCount.value, "e") > 0 OR _
			Instr(1, o_IFrame.TxtRetryCount.value, ".") > 0 OR _
			Instr(1, o_IFrame.TxtRetryCount.value, "-") > 0 OR _
			o_IFrame.TxtRetryCount.value = "0" Then
			s_ErrorMsg = s_ErrorMsg & "Retry Count is a Numeric Field and Must be 1 or Higher Number." & VBCrLf
		end if
	END IF
	IF o_IFrame.TxtRetryWaitTime.value <> "" Then
		If Not IsNumeric(o_IFrame.TxtRetryWaitTime.value) OR _
			Instr(1, o_IFrame.TxtRetryWaitTime.value, "e") > 0 OR _
			Instr(1, o_IFrame.TxtRetryWaitTime.value, ".") > 0 OR _
			Instr(1, o_IFrame.TxtRetryWaitTime.value, "-") > 0 OR _
			o_IFrame.TxtRetryWaitTime.value = "0" Then
				s_ErrorMsg = s_ErrorMsg & "Retry Wait Time is a Numeric Field and Must be 1 or Higher Number." & VBCrLf
		End If
	END IF	
	IF o_IFrame.TxtTransmission.value = "" Then
		s_ErrorMsg = s_ErrorMsg & "Transmission Type is a Required Field." & VBCrLf
	Else
		Select Case CStr(o_IFrame.TxtTransmission.value)
			Case "1"
				IF o_IFrame.TxtDestination.value = "" OR _
					Not IsNumeric(o_IFrame.TxtDestination.value) OR _
					Instr(1, o_IFrame.TxtDestination.value, ".") > 0 OR _
					Instr(1, o_IFrame.TxtDestination.value, "e") > 0 OR _
					Instr(1, o_IFrame.TxtDestination.value, "-") > 0 OR _
					Len(o_IFrame.TxtDestination.value) < 10 Then
						s_ErrorMsg = s_ErrorMsg & "Please Provide a 10 Digit Fax Number in the Destination Field." & VBCrLf
				End If
			Case "2"
				IF o_IFrame.TxtDestination.Value = "" Then o_IFrame.TxtDestination.Value = "<%=s_ProductionPrinter%>"
				IF o_IFrame.TxtDestination.Value <> "" And o_IFrame.TxtDestination.Value <> "<%=s_ProductionPrinter%>" THEN
					If  MsgBox("The Default Printer Will Be Overrided by " & o_IFrame.TxtDestination.Value & "." & VBCrLf & "YES to Continue, NO to Return to the Screen Panel?", vbYesNo, "FNSDesigner") = VBNo Then
							f_ValidateScreenData = False
							Exit Function
					End If
				End If	
		End Select
	END IF
Set o_IFrame = Nothing
	IF s_ErrorMsg = "" Then
		 f_ValidateScreenData = True
	ELSE
		MSGBOX s_ErrorMsg, 0, "FNSDesigner"
	END IF
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

Sub updateStatus(s_Msg)
	document.all.SpanStatus.innerHTML = s_Msg
End Sub

Sub updateIFrameStatus(s_Msg)
	SeqStepIFrame.document.all.SpanStatusSeqStep.innerHTML = s_Msg
End Sub

Sub PostTo(strURL)
	frm_Entry.action = "AHSpecDestSearch-f.asp?MODE=<%=Request.Querystring("MODE")%>&FULLSEARCH"
	frm_Entry.method = "POST"
	frm_Entry.target = "_parent"	
	frm_Entry.submit
End Sub

</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
  <tr><td colspan="2" HEIGHT="4"></td></tr>
  <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Specific Destination&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
  <td HEIGHT="5" ALIGN="LEFT">
							<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
									<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
									<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
										<td WIDTH="300" HEIGHT="8"></td></tr>
							</table>
							</td></tr>
  <tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
  <tr><td colspan="2" HEIGHT="1"></td></tr>
</table><br>
<%IF Request.QueryString("SDID") <> "" Then%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
  <tr><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18"><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report"></td>
      <td WIDTH="385">:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span></td></tr>
</table>
<table class="LABEL">
  <tr><td width="305" nowrap>Specific Destination ID:&nbsp;<span ID="span_SDID" CLASS="LABEL"><%=Request.QueryString("SDID")%></span></td>
</table>  
<table class="LABEL">
  <tr><td width="305" nowrap>Account:&nbsp;<span ID="AHSID_TEXT" CLASS="LABEL"><%= ReplaceQuotesInText(Request.QueryString("NAME")) %></span></td>
	  <td>A.H.Step ID:&nbsp;<span ID="ACCNT_ID" CLASS="LABEL"><%=Request.QueryString("AHSID")%></span></td></tr>
</table>
<input Type="Hidden" Name="Txt_DestinationID" Value="<%=Request.QueryString("SDID")%>">
<input Type="Hidden" Name="Txt_ACCNTID" Value="<%=s_ACCNTID%>">
<table CLASS="LABEL" cellpadding="2" cellspacing="2">
<tr><td>
    <table>
      <tr><td CLASS="LABEL">Last Name:<br><input ScrnInput="TRUE" size="45" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" ID="TxtNAME_LAST" VALUE="<%=s_LastName%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	      <td CLASS="LABEL">First Name:<br><input ScrnInput="TRUE" size="30" CLASS="LABEL" MAXLENGTH="40" TYPE="TEXT" ID="TxtNAME_FIRST" VALUE="<%=s_FirstName%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	      <td CLASS="LABEL">MI:<br><input ScrnInput="TRUE" size="1" CLASS="LABEL" MAXLENGTH="1" TYPE="TEXT" ID="TxtNAME_INITIAL" VALUE="<%=s_MI%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td></tr>
	  <tr><td CLASS="LABEL">Title:<br><input ScrnInput="TRUE" size="45" CLASS="LABEL" MAXLENGTH="80" TYPE="TEXT" ID="TxtTITLE" VALUE="<%=s_Title%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
		  <td CLASS="LABEL" colspan="2">Phone:<br><input ScrnInput="TRUE" size="14" CLASS="LABEL" MAXLENGTH="14" TYPE="TEXT" ID="TxtPHONE" VALUE="<%=s_Phone%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td></tr>
	  <tr><td CLASS="LABEL" COLSPAN="2">Address 1:<br><input ScrnInput="TRUE" size="45" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" ID="TxtADDRESS1" VALUE="<%=s_Address1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
		<td CLASS="LABEL">&nbsp;</td>
		<td CLASS="LABEL">LOB:<br><select ScrnBtn="TRUE" NAME="TxtLOB" CLASS="LABEL" tabindex=3 ID="Select1" multiple size="5"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD","",false)%></select></td>
	  </tr>
	  <tr><td CLASS="LABEL" COLSPAN="4">Address 2:<br><input ScrnInput="TRUE" size="45" CLASS="LABEL" MAXLENGTH="45" TYPE="TEXT" ID="TxtADDRESS2" VALUE="<%=s_Address2%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td></tr>
	</table>
	<table>
		<tr>
      <td CLASS="LABEL" ALIGN="LEFT">Zip:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="10" size="10" TYPE="TEXT" NAME="ZIP" VALUE="<%=s_ZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	  <td CLASS="LABEL">City:<br><input CLASS="READONLY" READONLY TABINDEX=-1 MAXLENGTH="30" size="30" TYPE="TEXT" NAME="CITY" VALUE="<%=s_City%>" ></td>
      <td CLASS="LABEL">State:<br><input CLASS="READONLY" READONLY TABINDEX=-1 MAXLENGTH="3" size="3" TYPE="TEXT" NAME="STATE" VALUE="<%=s_State%>" ></td>
      <td CLASS="LABEL" width="76">&nbsp;</td>
      <td CLASS="LABEL" VALIGN="BOTTOM" align="right"><input TYPE="CHECKBOX" ScrnBtn="TRUE" NAME="OSHAE_FORM" CLASS="LABEL" OnCLick="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Checkbox1">OSHA Form</td>
      </tr>
	 </table>
<% IF Request.QueryString("SDID") <> "NEW" Then %>
	 <td VALIGN="TOP" ALIGN="RIGHT">
	 <table>
	 <tr><td CLASS="LABEL" ALIGN="RIGHT" Width="120"><button CLASS="STDBUTTON" <% If MODE = "RO" Then Response.write(" DISABLED ") %> NAME="BtnSaveSpDestination" ACCESSKEY="S"><u>S</u>ave</button></tr>
	 </table> </td>
<% END IF %>
</tr></td> 
</table>
<p></p>
<iframe FRAMEBORDER="0" WIDTH="100%" Height="225" ID="SeqStepIFrame" NAME="SeqStepIFrame" SRC="<%= s_iFrameSRC %>" Scrolling="no"></iframe>
<BR>
<hr>
<br>
<%else%>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Specific Destination selected.
</div>
<%end if%>
<form ID="frm_Entry" ACTION="SpecificDestinationSave.asp" METHOD="POST" Target="hiddenPage"> 
	<input Type="Hidden" Name="Txt_SQLString">
	<input Type="Hidden" Name="Txt_Operation">
</form>
</body>
</html>