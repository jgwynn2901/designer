<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%
Response.Expires=0 
	AccountTextLen = 30	

	Dim SharedCount, SharedCountText, XREFID
	SharedCount = 0
	SharedCountText = "Ready"
	
	XREFID= CStr(Request.QueryString("XREFID"))
  
	 If XREFID <> "" Then
		 If XREFID = "NEW" Then 
		    SharedCount = 0
		   End If
	  End If	
	
	
	
    If XREFID <> "" Then

	If XREFID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = ""
		
		'SQLST = "SELECT COVERAGECODE_CONVERSION_ID,ACCNT_HRCY_STEP_ID,"
		'SQLST = SQLST & "COVERAGE_CODE,VENDOR_DESIGNATOR,DESCRIPTION "
		'SQLST = SQLST & "FROM COVERAGECODE_CONVERSION "
		'SQLST = SQLST & "WHERE COVERAGECODE_CONVERSION_ID ="& XREFID
      SQLST = "SELECT COVERAGECODE_CONVERSION.*,ACCOUNT_HIERARCHY_STEP.NAME "
      SQLST = SQLST & "FROM ACCOUNT_HIERARCHY_STEP,COVERAGECODE_CONVERSION "
      SQLST = SQLST & "WHERE COVERAGECODE_CONVERSION.ACCNT_HRCY_STEP_ID = ACCOUNT_HIERARCHY_STEP.ACCNT_HRCY_STEP_ID(+)"
      SQLST = SQLST & "AND COVERAGECODE_CONVERSION_ID ="& XREFID
		
		 
		
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then
			RS_XREFID             = RS("COVERAGECODE_CONVERSION_ID")
			RS_AHSID              = RS("ACCNT_HRCY_STEP_ID")
			RS_AHSNAMEID          = RS("NAME")
			RS_COVERAGE_CODE      = RS("COVERAGE_CODE")
			RS_VENDOR_DESIGNATOR  = RS("VENDOR_DESIGNATOR")
			RS_DESCRIPTION        = ReplaceQuotesInText(RS("DESCRIPTION"))
			
	
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
<title>Greeting Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
var AHSSearchObj = new CAHSSearchObj();
var RuleSearchObj = new CRuleSearchObj();
var g_StatusInfoAvailable = false;
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if XREFID <> "" then %>
			<% if SharedCount <= 1 then %>
			 document.all.txtCoverageCode.value = "<%=RS_COVERAGE_CODE %>"
			 document.all.txtVendorDesignator.value = "<%=RS_VENDOR_DESIGNATOR%>"
			 document.all.txtDescription.value = "<%=RS_DESCRIPTION%>"
			
<%	else %>
	<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
			end if
		end if	
	end if 
%>

End Sub

Sub PostTo(strURL)
	FrmDetails.action = "CoverageSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateXREFID(inXREFID)

	document.all.XREFID.value = inXREFID
	document.all.spanXREFID.innerText = inXREFID
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

Function GetXREFID
	if document.all.XREFID.value <> "NEW" then
		GetXREFID = document.all.XREFID.value
	else
		GetXREFID = ""
	end if 
End Function

'Function GetXREFIDName
	'GetXREFIDName = document.all.txtCoverageCode.value
'End Function

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
errmsg = ""
	
	
	If  document.all.txtCoverageCode.value = "" then
		errmsg = errmsg &  "Coverage Code is a required field."
	end if
	If  document.all.AHSID_ID.innerText = "" then
		errmsg = errmsg &  "AHS ID is a required field."
	end if
	If  document.all.txtVendorDesignator.value = "" then
		errmsg = errmsg &  "Vendor Designator is a required field."
	end if
	If  document.all.txtDescription.value = "" then
		errmsg = errmsg &  "Description is a required field."
	end if
	If errmsg = "" Then
		ValidateScreenData = true
	Else
		msgbox errmsg , 0 , "FNSDesigner"
		ValidateScreenData = false
	End If
End Function

sub UpdateScreenOnDelete()
	document.all.XREFID.value = ""
	FrmDetails.action = "CoverageDetails.asp?STATUS=Delete successful."
	FrmDetails.target = "_self"
	FrmDetails.submit
end sub
Function ExeDelete
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeDelete = bRet
		exit Function
	end if
	
	if document.all.XREFID.value = "" then
		ExeDelete = false
		exit function
	end if

End Function
Function ExeSave
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.XREFID.value = "" then
		ExeSave = false
		exit function
	end if
If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
	 If document.all.XREFID.value = "NEW"  then
	  document.all.TxtAction.value = "INSERT"
	else
	    document.all.TxtAction.value = "UPDATE"
	 end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if 
		
	    sResult = sResult & "COVERAGECODE_CONVERSION_ID" & Chr(129) & document.all.XREFID.value & Chr(129) & "0" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID" & Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "COVERAGE_CODE" & Chr(129) & document.all.txtCoverageCode.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "VENDOR_DESIGNATOR" & Chr(129) & document.all.txtVendorDesignator.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.txtDescription.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
 Else   
		SpanStatus.innerHTML = "Nothing to Save"
End If
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
Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Function AttachRule (ID, SPANID, strTITLE)
	RID = ID.value
	MODE = document.body.getAttribute("ScreenMode")

	RuleSearchObj.RID = RID
	RuleSearchObj.RIDText = SPANID.title
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
			
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_COVERAGE_CODE_XREF&RID="& RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.value then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.value = RuleSearchObj.RID
		end if
		UpdateSpanText SPANID, RuleSearchObj.RIDText
	ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateSpanText SPANID, RuleSearchObj.RIDText
	End If

End Function

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.value = ""
		SPANID.innerText = ""
	end if
End Function

Function DetachAccount(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Function AttachAccount (ID, SPANID)
	AHSID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.AHSIDName = SPANID.title
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No account currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_COVERAGE_CODE_XREF&SELECTONLY=TRUE&AHSID=" &AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	showModalDialog  strURL  ,AHSSearchObj ,"center"

	'if Selected=true update everything, otherwise if AHSID is the same, update text in case of save
	If AHSSearchObj.Selected = true Then
		If AHSSearchObj.AHSID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = AHSSearchObj.AHSID
		end if
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	ElseIf ID.innerText = AHSSearchObj.AHSID And AHSSearchObj.AHSID<> "" Then
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	End If
End Function

<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
  <table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
    <tr><td colspan="2" HEIGHT="4"></td></tr>
    <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Coverage Code Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   <img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
    <td HEIGHT="5" ALIGN="LEFT">
      <table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
       <tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
       <tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
       <td WIDTH="300" HEIGHT="8"></td></tr>
      </table></td></tr>
   <tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
   <tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="CoverageSave.asp" TARGET="hiddenPage">
   <input TYPE="HIDDEN" NAME="TxtSaveData">
   <input TYPE="HIDDEN" NAME="TxtAction">
   <input TYPE="HIDDEN" NAME="TxtCount">
   

   
<% 'need to maintain these values in order to post back to the search tab %>

<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" name="SearchXREFID" value="<%=Request.QueryString("SearchXREFID")%>">
<input type="hidden" name="SearchCoverageCode" value="<%=Request.QueryString("SearchCoverageCode")%>">
<input type="hidden" name="SearchVendorDesignator" value="<%=Request.QueryString("SearchVendorDesignator")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="XREFID" value="<%=Request.QueryString("XREFID")%>">
<%	

If XREFID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
</tr>
</table>
<table class="LABEL">
<tr>
	<td>
	<img NAME="BtnAttachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<img NAME="BtnDetachAHSID" STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::DetachAccount AHSID_ID, AHSID_TEXT">
	</td>
	<!--<td width=305 nowrap>Account:<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=RS_AHSNAMEID%>" ></SPAN></td>-->
	<td width="305" nowrap>Account:<span ID="AHSID_TEXT" CLASS="LABEL"> <%=RS_AHSNAMEID%></span></td>
	<td>A.H.Step ID: <span ID="AHSID_ID" CLASS="LABEL"><%=RS_AHSID%></span></td>
	</tr>
</table>

<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<td><table class="LABEL">
<tr>
	<td CLASS="LABEL" COLSPAN="2">Coverage Code ID:&nbsp;<span id="spanXREFID"><%=Request.QueryString("XREFID")%></span></td>
	</tr>
	</table>
	<table>
	<tr>
	<tr nowrap>
    <td CLASS="LABEL">Coverage Code:<br>
	<input size="15" CLASS="LABEL" MAXLENGTH="15" TYPE="TEXT" NAME="txtCoverageCode" tabindex="1" VALUE="<%=RS_COVERAGE_CODE %>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td><img src="arrw19d.gif" WIDTH="31" HEIGHT="31"></td>
	<td CLASS="LABEL">Vendor Designator:<br>
	<select STYLE="WIDTH:75" NAME="txtVendorDesignator" CLASS="LABEL" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange">
		<option VALUE="RENT">RENT
		<option VALUE="TOW">TOW
		<option VALUE="BODY">BODY
		<option VALUE="TEST">TEST
		<option VALUE="WATER">WATER
		<option VALUE="PP">PERS
	</select>
	</td>
	</tr>
    <tr>
	<td CLASS="LABEL" colspan="3">Description:<br>
	<input size="100" CLASS="LABEL" MAXLENGTH="127" TYPE="TEXT" NAME="txtDescription" tabindex="3" VALUE="<%=RS_DESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	
	<tr>
	
    <tr>
	
 </table>
<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
<%=Request.QueryString("STATUS") & "<br>"%>
No Coverage Code selected.
</div>
<% End If %>
</form>
</body>
</html>


