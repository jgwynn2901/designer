<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ControlData.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub BtnClear_onclick()
	document.all.SearchPID.value = ""
	document.all.SearchNumber.value = ""
	document.all.SearchAHSID.value = ""
	document.all.SearchCarrier.value = ""
	document.all.SearchTPADMIN.value = ""
	document.all.SearchAgent.value = ""
	document.all.SearchLOBCD.value = ""
	document.all.SearchMCTYPE.value = ""
	document.all.SearchSelfInsuredFlg.value = ""
	document.all.SearchEffective.value = ""
	document.all.SearchOriginalEffective.value = ""
	document.all.SearchExpiration.value = ""
	document.all.SearchCancellation.value = ""
	document.all.SearchCompanyCode.value = ""
End Sub

Sub BtnSearch_onclick()
	errstr = ""

	If Not CheckDate(document.all.SearchEffective.value) AND document.all.SearchEffective.value <> "" Then
		errstr = errstr & "Effective Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
	End If
	If Not CheckDate(document.all.SearchExpiration.value) AND document.all.SearchExpiration.value <> "" Then
		errstr = errstr & "Expiration Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
	End If
	If Not CheckDate(document.all.SearchOriginalEffective.value) AND document.all.SearchOriginalEffective.value <> "" Then
		errstr = errstr & "Original Effective Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
	End If
	If errstr = "" Then
		SPANSTATUS.innerHTML = "<%= MSG_SEARCH %>"
		FrmSearch.submit
	Else
		MsgBox errstr, 0 , "FNSNetDesigner"
	End If
	
End Sub

Function CheckDate( InDate )
	If Not IsDate(InDate) Then
		CheckDate = false
		Exit Function
	End If
	If Len(InDate) <> 10  Then
		CheckDate = false
		Exit Function
	End If
	If Not IsNumeric(Mid(InDate,1,2)) Then
		CheckDate = false
		Exit Function
	End If	
	If Mid(InDate,1,2) > 12 Then 
		CheckDate = false
		Exit Function
	End If
	CheckDate = true
End Function

Sub window_onload
	'document.all.SearchName.focus ' Timing Problem
	document.all.SearchType(0).checked = True
	UpdateStatus("Ready")	
	
<%	If Request.QueryString <> "" Then %>
<%		If CStr(Request.QueryString("SearchType")) = "B" Then	%>
			document.all.SearchType(0).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "C" Then	%>
			document.all.SearchType(1).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "E" Then	%>
			document.all.SearchType(2).checked = True
<%		End If 				

		If CStr(Request.QueryString("SearchLOBCD")) <> "" Then	%>
			SelectOption document.all.SearchLOBCD,"<%=CStr(Request.QueryString("SearchLOBCD"))%>"
<%		End If 

		If CStr(Request.QueryString("SearchMCTYPE")) <> "" Then	%>
			SelectOption document.all.SearchMCTYPE,"<%=CStr(Request.QueryString("SearchMCTYPE"))%>"
<%		End If 

		If CStr(Request.QueryString("SearchSelfInsuredFlg")) <> "" Then	%>
			SelectOption document.all.SearchSelfInsuredFlg,"<%=CStr(Request.QueryString("SearchSelfInsuredFlg"))%>"
<%		End If 
	End If %>	

	If document.all.SearchPID.value <> "" Or document.all.SearchNumber.value <> "" Or _
		document.all.SearchAHSID.value <> "" Or document.all.SearchTPADMIN.value <> "" Or _
		document.all.SearchCarrier.value <> ""  Or document.all.SearchAgent.value <> "" Or _
		document.all.SearchLOBCD.value <> ""  Or document.all.SearchSelfInsuredFlg.value <> "" Or _
		document.all.SearchEffective.value <> ""  Or document.all.SearchOriginalEffective.value <> "" Or _
		document.all.SearchExpiration.value <> ""  Or _
		document.all.SearchCompanyCode.value <> "" Or document.all.SearchMCTYPE.value <> ""  Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub

Sub PostTo(cURL)
	dim cPID, aArr0, cLOB

	cPID = Parent.frames("WORKAREA").GetPID
	cLOB = Parent.frames("WORKAREA").GetLOB
	aArr0 = Split(cPID, "||")
	If UBound(aArr0) >= 0 Then 
		document.all.PID.value = aArr0(0)
		document.all.LOB.value = cLOB
	Else		
		document.all.PID.value = ""
		document.all.LOB.value = ""
	End If

	FrmSearch.action = cURL
	FrmSearch.method = "GET"	
	FrmSearch.target = "_parent"	
	FrmSearch.submit
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub StatusRpt_OnClick
	MsgBox "No other detail status reported.",0,"FNSNetDesigner"		
End Sub
sub enable_exact()
   document.all.SearchType(2).checked  = true 
   document.all.SearchType(0).disabled  = true 
   document.all.SearchType(1).disabled  = true 
end sub

sub enable_begin()
   document.all.SearchType(0).checked  = true 
   document.all.SearchType(0).disabled  = false 
   document.all.SearchType(1).disabled  = false 
   
end sub
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Policy Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>

<form Name="FrmSearch" METHOD="GET" ACTION="PolicySearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="PID" value="<%=Request.QueryString("PID")%>">
<input type="hidden" NAME="LOB" value="<%=Request.QueryString("LOB")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
 <table>
	<tr></tr>
	<tr></tr>
	<tr nowrap>  
	<td CLASS="LABEL" colspan="2">Number:<br><input CLASS="LABEL" tabindex="1" size="30" TYPE="TEXT" NAME="SearchNumber" VALUE="<%=Request.QueryString("SearchNumber")%>"></td>
	</tr>				
 	<tr>	 			
	<td NOWRAP CLASS="LABEL">A.H.Step ID:<br><input CLASS="LABEL" size="12" tabindex="3" TYPE="TEXT" NAME="SearchAHSID" VALUE="<%=Request.QueryString("SearchAHSID")%>" onfocus="enable_exact()" onBlur="enable_begin()"></td>
	<td NOWRAP CLASS="LABEL">Carrier ID:<br><input CLASS="LABEL" size="12" tabindex="4" TYPE="TEXT" NAME="SearchCarrier" VALUE="<%=Request.QueryString("SearchCarrier")%>"></td>	
	<td NOWRAP CLASS="LABEL">Agent ID:<br><input CLASS="LABEL" size="12" tabindex="5" TYPE="TEXT" NAME="SearchAgent" VALUE="<%=Request.QueryString("SearchAgent")%>"></td>
    <td NOWRAP CLASS="LABEL">T P A  ID:<br><input CLASS="LABEL" size="12" tabindex="6" TYPE="TEXT" NAME="SearchTPADMIN" VALUE="<%=Request.QueryString("SearchTPADMIN")%>" ID="Text1"></td>	
	<td NOWRAP CLASS="LABEL" colspan="2">LOB:<br><select NAME="SearchLOBCD" CLASS="LABEL" tabindex="7"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD","",true)%></select></td>
	</tr>	
 	<tr>	 			
	<td NOWRAP CLASS="LABEL">Managed Care Type:<br><select NAME="SearchMCTYPE" CLASS="LABEL" tabindex="8">
	 <option VALUE>
	 <option VALUE="NEITHER">NEITHER
	 <option VALUE="CERTIFIED">CERTIFIED
	 <option VALUE="NOTCERTIFIED">NOTCERTIFIED</select></td>
	<td NOWRAP CLASS="LABEL" colspan="4">Self insured?<br><select CLASS="LABEL" tabindex="9" name="SearchSelfInsuredFlg"> 
				<option SELECTED value> <both> </option>
				<option value="Y">Yes</option>
			 	<option value="N">No</option>
				</select></td>
	</tr>					
	<tr>				 
	<td NOWRAP CLASS="LABEL">Effective Date:<br><input CLASS="LABEL" size="12" tabindex="10" TYPE="TEXT" NAME="SearchEffective" VALUE="<%=Request.QueryString("SearchEffective")%>"></td>
	<td NOWRAP CLASS="LABEL">Expiration Date:<br><input CLASS="LABEL" size="12" tabindex="11" TYPE="TEXT" NAME="SearchExpiration" VALUE="<%=Request.QueryString("SearchExpiration")%>"></td>
	</tr>				 
	<tr>				
	<td NOWRAP CLASS="LABEL">Policy ID:<br><input CLASS="LABEL" tabindex="13" size="12" TYPE="TEXT" NAME="SearchPID" VALUE="<%=Request.QueryString("SearchPID")%>"></td>
	<td NOWRAP CLASS="LABEL">Orig. Effect. Date:<br><input CLASS="LABEL" size="12" tabindex="14" TYPE="TEXT" NAME="SearchOriginalEffective" VALUE="<%=Request.QueryString("SearchOriginalEffective")%>"></td>
	<td NOWRAP CLASS="LABEL" colspan="2">Company Code:<br><input CLASS="LABEL" size="12" tabindex="16" TYPE="TEXT" NAME="SearchCompanyCode" VALUE="<%=Request.QueryString("SearchCompanyCode")%>"></td>
	</tr>				
	<tr>
	</tr>
 </table>
</td>			
<td VALIGN="TOP" rowspan="3">
 <table>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="20" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="21" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
 </table>
</td>	
</tr>
</table>
<table topmargin="0" bottommargin="0">
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex="17" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex="18" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex="19" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</table>
</form>
</body>
</html>
