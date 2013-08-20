<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%
 
    Response.Expires=0 
	AccountTextLen = 30	

	Dim SharedCount, SharedCountText, COID, ClientCode

	
	
	SharedCount = 0
	SharedCountText = "Ready"
	
	COID	= CStr(Request.QueryString("COID"))
	Mode    = Request.QueryString("MODE")
	ClientCode = left(getinstancename,3)

	If COID <> "" Then
		If COID = "NEW" Then 
			SharedCount = 0
		End If
	End If	
	
    If COID <> "" Then
	   If COID <> "NEW" then
				IF ClientCode = "SED" then
				s_iFrameSRC = "SEDContactDetailsAHSSummary.asp?" & "Mode=" & Mode &  "&COID=" &  COID & "&ClientCode=" & ClientCode
				else
	      s_iFrameSRC = "ContactDetailsAHSSummary.asp?" & "Mode=" & Mode &  "&COID=" &  COID & "&ClientCode=" & ClientCode
	      end if
		  Set Conn = Server.CreateObject("ADODB.Connection")
		  Conn.Open CONNECT_STRING

		  '**********************************
		  ' DMS: 3/20/00 Modified the select stmt. to 
		  ' accomodate schema changes
		  '**********************************
		  
		  SQLST = "SELECT C.*, " &_
		          "       AHSC.ACCNT_HRCY_STEP_ID, " &_ 
				  "       AHS.NAME ACCOUNT_NAME " &_
		          "  FROM CONTACT C, " &_
				  "       AHS_CONTACT AHSC, " &_
				  "       ACCOUNT_HIERARCHY_STEP AHS " &_
				  " WHERE C.CONTACT_ID            = AHSC.CONTACT_ID(+)  " &_
				  "   AND AHSC.ACCNT_HRCY_STEP_ID = AHS.ACCNT_HRCY_STEP_ID(+)  " &_
				  "   AND C.CONTACT_ID            = " & COID
		  Set RS = Conn.Execute(SQLST)
		  If Not RS.EOF then
			RSNAME               = ReplaceQuotesInText(RS("NAME"))
			RSTYPE               = ReplaceQuotesInText(RS("TYPE"))
			RSTITLE              = ReplaceQuotesInText(RS("TITLE"))
			RSADDRESSLINE1       = ReplaceQuotesInText(RS("ADDRESS_LINE1"))
			RSADDRESSLINE2       = ReplaceQuotesInText(RS("ADDRESS_LINE2"))
			RSCITY               = ReplaceQuotesInText(RS("CITY"))
			RSSTATE              = ReplaceQuotesInText(RS("STATE"))
			RSZIP                = ReplaceQuotesInText(RS("ZIP"))
			RSCOUNTRY            = ReplaceQuotesInText(RS("COUNTRY"))
			RSPHONE              = RS("PHONE")
			RSFAX                = RS("FAX")
			RSEMAIL              = RS("EMAIL")
			RSDESCRIPTION        = ReplaceQuotesInText(RS("DESCRIPTION"))
			RSACCNT_HRCY_STEP_ID = RS("ACCNT_HRCY_STEP_ID")
			RSACCOUNT_NAME       = ReplaceQuotesInText(RS("ACCOUNT_NAME"))
		 if ClientCode = "SED" then
			RSCELLPHONE					 = RS("CELL_PHONE")
			RSOTHERPHONE				 = RS("OTHER_PHONE")
			RSHOURTYPE  				 = RS("HOUR_TYPE")
			end if
			
		  end if	
		  RS.Close
		  Set RS = Nothing
		  Conn.Close
		  Set Conn = Nothing
	    end if	
    End If
	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Contact Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
var AHSSearchObj = new CAHSSearchObj();
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if COID <> "" then 
		  if ClientCode = "SED" then %>
		  document.all.txthourtype.value = "<%= RSHOURTYPE %>"
		  <%end if %>
			<% if SharedCount <= 1 then %>
			document.all.ChkEdit.checked = true
			ChkEdit_OnClick
<%	else %>
	SetStatusInfoAvailableFlag(true)
			document.all.ChkEdit.checked = false
			ChkEdit_OnClick
<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
		end if	
	end if 
%>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "CONTACTSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateCOID(inCOID)
	document.all.COID.value = inCOID
	document.all.spanCOID.innerText = inCOID
End Sub


Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then 
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Function GetCOID
	if document.all.COID.value <> "NEW" then
		GetCOID = document.all.COID.value
	else
		GetCOID = ""
	end if 
End Function

Function GetCOIDName
	GetCOIDName = document.all.TxtName.value
End Function

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

Function ValidateScreenData
	If  document.all.TxtName.value = "" then
		MsgBox "Name is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	ValidateScreenData = true
End Function

Function ExeCopy
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeCopy = bRet
		exit Function
	end if
	
	if document.all.COID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.COID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.COID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.COID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
	
		sResult = sResult & "CONTACT_ID"& Chr(129) & document.all.COID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "TYPE"& Chr(129) & document.all.TxtTYPE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TITLE"& Chr(129) & document.all.TxtTITLE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ADDRESS_LINE1"& Chr(129) & document.all.TxtAddressLine1.value & Chr(129) & "1" & Chr(128)
    sResult = sResult & "ADDRESS_LINE2"& Chr(129) & document.all.TxtAddressLine2.value & Chr(129) & "1" & Chr(128)
    sResult = sResult & "CITY"& Chr(129) & document.all.TxtCity.value & Chr(129) & "1" & Chr(128)
    sResult = sResult & "STATE"& Chr(129) & document.all.TxtState.value & Chr(129) & "1" & Chr(128)
    sResult = sResult & "ZIP"& Chr(129) & document.all.TxtZip.value & Chr(129) & "1" & Chr(128)
    sResult = sResult & "COUNTRY"& Chr(129) & document.all.TxtCountry.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE"& Chr(129) & document.all.TxtPHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAX"& Chr(129) & document.all.TxtFAX.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "EMAIL"& Chr(129) & document.all.TxtEMAIL.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDESCRIPTION.value & Chr(129) & "1" & Chr(128)
		
		'added for Sedgwick
		<%if left(getInstanceName,3) = "SED" then %>
		sResult = sResult & "CELL_PHONE"& Chr(129) & document.all.TxtCELLPHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "OTHER_PHONE"& Chr(129) & document.all.TxtOTHERPHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "HOUR_TYPE"& Chr(129) & document.all.TxtHourType.value & Chr(129) & "1" & Chr(128)
		<%end if %>
		
		' DMS: 3/22/00 Made changes to accomodate the new table ahs_contact 
		' Pass 1 for AHSID to update contact tble. Pass the actual AHSID to update ahs_contact table
		
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & 1 & Chr(129) & "1" & Chr(128)
		
		document.all.TxtSaveData.Value = sResult
        'document.all.AHSID.Value       = document.all.AHSID_ID.innerText
		document.all.FrmDetails.Submit()
		
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If
	ExeSave = bRet
End Function

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
	end if
end sub

sub SetScreenFieldsReadOnly(bReadOnly, strNewClass)

	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnInput") = "TRUE" then
			document.all(iCount).readOnly = bReadOnly
			document.all(iCount).className = strNewClass
		elseif document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then
			document.all(iCount).disabled = bReadOnly
		end if
	next

end sub

sub ChkEdit_OnClick
	document.all.ChkEdit.setAttribute "ScrnBtn","FALSE"
	
	if document.all.ChkEdit.checked = true then
		SetScreenFieldsReadOnly false,"LABEL"
		document.body.setAttribute "ScreenMode", "RW"		
	else
		SetScreenFieldsReadOnly true,"DISABLED"
		document.body.setAttribute "ScreenMode", "RO"
	end if
	document.all.ChkEdit.setAttribute "ScrnBtn","TRUE"
end sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub




Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Contact Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="CONTACTSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">
<INPUT TYPE="HIDDEN" NAME="AHSID" 

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchCOID" value="<%=Request.QueryString("SearchCOID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" name="SearchDescription" value="<%=Request.QueryString("SearchDescription")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" name="SearchsType" value="<%=Request.QueryString("SearchsType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="COID" value="<%=Request.QueryString("COID")%>" >
<input type="hidden" NAME="ClientCode" value="<%=Request.QueryString("ClientCode")%>" ID="Hidden1">

<%	
If COID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=SharedCountText%></SPAN>
</td>
<td>
<input  ScrnBtn="TRUE"  TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING=0 CELLSPACING=0 >
<tr>
<td>

<table class="LABEL">
<tr>
	<td>
	</td>
</tr>
</table>
<table class="LABEL">
	<tr>
	<td CLASS=LABEL COLSPAN=2>Contact ID:&nbsp<span id="spanCOID"><%=Request.QueryString("COID")%></span></td>
	</tr>
	<tr>
	<td CLASS=LABEL>Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=255 SIZE=30 TYPE="TEXT" NAME="TxtName" VALUE="<%=RSNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS=LABEL>Type:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=40 size=20 TYPE="TEXT" NAME="TxtType" VALUE="<%=RSTYPE%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS=LABEL>Title:<br><input ScrnInput="TRUE" class = "LABEL"  size=20 MAXLENGTH=20 TYPE="TEXT" NAME="TxtTitle" VALUE="<%=RSTITLE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
		<% if left(getInstanceName,3) = "SED" then%>
	<td class = label >Hour Type:<br> <select NAME="txtHourType"   CLASS="LABEL" ID="Select1" ONCHANGE="VBScript::Control_OnChange">
	<option value =""></option>
	<option value = "Normal(8:00AM TO 4:30PM EST, M-F)" >Normal(8:00AM TO 4:30PM EST, M-F)</option>
	<option value = "After(4:31PM TO 7:59AM EST, M-F,all day Saturday & Sunday)" >After(4:31PM TO 7:59AM EST, M-F,all day Saturday & Sunday)</option>
	<option value = "Both(8:00AM TO 7:59AM EST, M-F,all day Saturday & Sunday)" >Both(8:00AM TO 7:59AM EST, M-F,all day Saturday & Sunday)</option>
	
	</select>  </td>
	<%end if %>
	</tr>
	
	<tr>
	<td CLASS=LABEL>Address_line1:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=30 size=30 TYPE="TEXT" NAME="TxtAddressLine1" VALUE="<%=RSADDRESSLINE1%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS=LABEL>Address_line2:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=30 size=20 TYPE="TEXT" NAME="TxtAddressLine2" VALUE="<%=RSADDRESSLINE2%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td CLASS=LABEL>City:<br><input ScrnInput="TRUE" size=20 CLASS="LABEL" MAXLENGTH=30 TYPE="TEXT" NAME="TxtCity" VALUE="<%=RSCITY%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
    </tr>
    <tr>
    <td CLASS=LABEL>State:<br><input ScrnInput="TRUE" size=2 CLASS="LABEL" MAXLENGTH=3 TYPE="TEXT" NAME="TxtState" VALUE="<%=RSSTATE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>	
    <td CLASS=LABEL>Zip:<br><input ScrnInput="TRUE" size=15 CLASS="LABEL" MAXLENGTH=15 TYPE="TEXT" NAME="TxtZip" VALUE="<%=RSZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>	
    <td CLASS=LABEL>Country:<br><input ScrnInput="TRUE" size=20 CLASS="LABEL" MAXLENGTH=30 TYPE="TEXT" NAME="TxtCountry" VALUE="<%=RSCOUNTRY%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>	
	</tr>
	<tr>
	  <td CLASS=LABEL>Phone:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=14 TYPE="TEXT" NAME="TxtPhone" VALUE="<%=RSPHONE%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	  <td CLASS=LABEL>Fax:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=14 TYPE="TEXT" NAME="TxtFax" VALUE="<%=RSFAX%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	  <td CLASS=LABEL>Email:<br><input ScrnInput="TRUE" size=20 CLASS="LABEL" MAXLENGTH=255 TYPE="TEXT" NAME="TxtEMAIL" VALUE="<%=RSEMAIL%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</TR>
	</TABLE>
	<TABLE>
	<TR>
		<td CLASS=LABEL COLSPAN=2>Description:<br><input ScrnInput = "TRUE" size="86" TYPE="TEXT" MAXLENGTH=2000 NAME="TxtDESCRIPTION" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange" VALUE="<%= RSDESCRIPTION%>"></td>
		<% if ClientCode = "SED" then %>
	<td CLASS=LABEL>Cell Phone:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=14 TYPE="TEXT" NAME="TxtCellPhone" VALUE="<%=RSCellPhone%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text2"></td>
	  <td CLASS=LABEL>Other Phone:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=10 size=14 TYPE="TEXT" NAME="TxtOtherPhone" VALUE="<%=RSOtherPhone%>"  ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" ID="Text3"></td>
	<%end if%>
	</TR>
</TABLE>
<% if COID <> "NEW" then%>
<iframe FRAMEBORDER="0" WIDTH="100%" Height="225" ID="SeqStepIFrame" NAME="SeqStepIFrame" SRC="<%= s_iFrameSRC %>" Scrolling="no"></iframe>
<% end if%>
<% Else %>

<DIV style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Contact selected.
</DIV>

<% End If %>


</form>
</body>
</html>


