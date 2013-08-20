<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->


<%	

Response.Expires = 0 
	Response.AddHeader  "Pragma", "no-cache"
	Response.Buffer = true
	
	OID = Request.QueryString("searchOID")
	'OID = Request.QueryString("OID")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Owner Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>

var g_StatusInfoAvailable = false;

</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 175;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 175;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub PostTo(strURL)
	FrmDetails.action = "OwnerSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub


Sub UpdateOID(inOID)
	document.all.OID.value = inOID
	document.all.spanOID.innerText = inOID
	
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

Function GetOID
	if document.all.OID.value <> "NEW" then
		GetOID = document.all.OID.value
	else
		GetOID = ""
	end if 
End Function

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Function f_CheckIsThisRequired
	IF CStr(document.all.getAttribute("IsThisRequired")) = "Y" Then
		f_CheckIsThisRequired = true
	ELSE
		f_CheckIsThisRequired = False
	END IF
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	' Nothing to validate as only owner_id is a required field restare all char allowing nulls.
	ValidateScreenData = true
End Function

Function GetSelectedAHSOID
	GetSelectedAHSOID = document.frames("DataFrame").GetSelectedAHSOID
End Function

Function GetSelectedAHSID
	GetSelectedAHSID = document.frames("DataFrame").GetSelectedAHSID
End Function

Sub ExeNewBranchRule
    
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.OID.value = "" Or document.all.OID.value = "NEW" Then
		Exit Sub
	End If

    <%If HasAddPrivilege("FNSD_OWNER","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to add AHS Owner.",0,"FNSNetDesigner"
		Exit Sub
    <%End If %>		
	dim OID, AHSOID, MODE
	AHSOID = "NEW"
	OID = document.all.OID.value
	MODE = document.body.getAttribute("ScreenMode")
    strURL = "AHSOwnerDetailsDataModal.asp?OID=" & OID & "&AHSOID=" & AHSOID & "&MODE=" & MODE 	
	showModalDialog  strURL ,"center"
	Refresh
End Sub

Sub Refresh
	OID = document.all.OID.value
	document.all.tags("IFRAME").item("DataFrame").src = "OwnerDetailsData.asp?SearchOID=" & OID
End Sub

Sub ExeEditBranchRule
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.OID.value = "" Or document.all.OID.value = "NEW" Then
		Exit Sub
	End If
	
    <%If HasModifyPrivilege("FNSD_OWNER","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to edit AHS Owner.",0,"FNSNetDesigner"
		Exit Sub
    <%End If %>		

	dim AHSOID, OID, AHSID
	AHSOID = GetSelectedAHSOID
	AHSID  = GetSelectedAHSID
	OID = document.all.OID.value
	
	If AHSOID <> "" Then
		strURL = "AHSOwnerDetailsDataModal.asp?OID=" & OID & "&AHSOID=" & AHSOID & "&AHSID=" & AHSID & "&MODE=" & MODE 	
		showModalDialog  strURL,"center"
		Refresh
	Else
		MsgBox "Please select a AHS Owner to Edit.", 0, "FNSNet Designer"		
	End If
	
End Sub

Sub ExeRemoveBranchRule
	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.OID.value = "" Or document.all.OID.value = "NEW" Then
		Exit Sub
	End If

	
    <%If HasDeletePrivilege("FNSD_OWNER","") <> True Then  %>		
		MsgBox "You do not have the appropriate security privileges to delete AHS Owner.",0,"FNSNetDesigner"
		Exit Sub
    <%End If %>		

	dim AHSOID, sResult
	AHSOID = GetSelectedAHSOID
	
	If AHSOID <> "" Then
		sResult = sResult &  AHSOID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		
		FrmDetails.action = "AHSOwnerSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select an AHS Owner to Remove.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub

Function InEditMode
	InEditMode = true
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		InEditMode = false
	End If
End Function



Function ExeSave
	If Not InEditMode Then
		ExeSave = false
		Exit Function
	End If

	If document.all.OID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
	if ValidateScreenData = false then 
		ExeSave = false
		exit function
	end if

	If document.all.OID.value = "NEW" then
		document.all.TxtAction.value = "INSERT"
	else
		document.all.TxtAction.value = "UPDATE"
	end if
	sResult = sResult & "OWNER_ID"      & Chr(129) & document.all.OID.value & Chr(129) & "0" & Chr(128)
	sResult = sResult & "NAME_TITLE"    & Chr(129) & document.all.TxtTitle.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "NAME_LAST"     & Chr(129) & document.all.TxtNameLast.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "NAME_FIRST"    & Chr(129) & document.all.TxtNameFirst.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESS_LINE1" & Chr(129) & document.all.TxtAddress1.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESS_LINE2" & Chr(129) & document.all.TxtAddress2.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESS_CITY"  & Chr(129) & document.all.TxtCity.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESS_STATE" & Chr(129) & document.all.TxtState.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESS_ZIP"   & Chr(129) & document.all.TxtZip.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESS_FIPS"  & Chr(129) & document.all.TxtFips.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESS_COUNTY" & Chr(129) & document.all.TxtCounty.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ADDRESs_COUNTRY" & Chr(129) & document.all.TxtCountry.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "PHONE_HOME"    & Chr(129) & document.all.TxtHomePhone.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "PHONE_WORK"    & Chr(129) & document.all.TxtWorkPhone.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "PHONE_FAX"     & Chr(129) & document.all.TxtFax.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "EMAIL"& Chr(129) & document.all.TxtEmail.value & Chr(129) & "1" & Chr(128)

	document.all.TxtSaveData.Value = sResult
	document.all.FrmDetails.action = "OwnerSave.asp"
	document.all.FrmDetails.method = "POST"
	document.all.FrmDetails.target = "hiddenPage"	
	document.all.FrmDetails.submit()
	bRet = true
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
<!--#include file="..\lib\Help.asp"-->
</script>
<!--#include file="..\lib\BABtnControl.inc"-->

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Owner Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="OwnerSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchOID"       value="<%=Request.QueryString("SearchOID")%>">
<input type="hidden" name="SearchTitle"     value="<%=Request.QueryString("SearchTitle")%>">
<input type="hidden" name="SearchNameLast"  value="<%=Request.QueryString("SearchNameLast")%>">
<input type="hidden" name="SearchNameFirst" value="<%=Request.QueryString("SearchNameFirst")%>">
<input type="hidden" name="SearchAdd1"      value="<%=Request.QueryString("SearchAdd1")%>">
<input type="hidden" name="SearchAdd2"      value="<%=Request.QueryString("SearchAdd2")%>">
<input type="hidden" name="SearchCity"      value="<%=Request.QueryString("SearchCity")%>">
<input type="hidden" name="SearchState"     value="<%=Request.QueryString("SearchState")%>">
<input type="hidden" name="SearchZip"       value="<%=Request.QueryString("SearchZip")%>">
<input type="hidden" name="SearchWphone"    value="<%=Request.QueryString("SearchWphone")%>">
<input type="hidden" name="SearchHphone"    value="<%=Request.QueryString("SearchHphone")%>">
<input type="hidden" name="SearchFax"       value="<%=Request.QueryString("SearchFax")%>">
<input type="hidden" NAME="MODE"            value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="OID"           value="<%=Request.QueryString("searchOID")%>">

<%	
    Dim OID
    OID	= CStr(Request.QueryString("searchOID"))
    If OID <> "" Then
	   If OID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM OWNER WHERE OWNER_ID = " & OID
				
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF Then
			TITLE       = RS("NAME_TITLE")
			NAME_FIRST  = ReplaceQuotesInText(RS("NAME_FIRST"))
			NAME_LAST   = ReplaceQuotesInText(RS("NAME_LAST"))
			ADD1        = RS("ADDRESS_LINE1")
			ADD2        = RS("ADDRESS_LINE2")
			CITY        = RS("ADDRESS_CITY")
			STATE       = RS("ADDRESS_STATE")
			ZIP         = RS("ADDRESS_ZIP")
			STATE       = RS("ADDRESS_STATE")
			WPHONE      = RS("PHONE_WORK")
			HPHONE      = RS("PHONE_HOME")
			FAX         = RS("PHONE_FAX")			
			
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>

<table CLASS="LABEL" >
<tr></tr>
<tr></tr>
<tr></tr>
<tr><td colspan=2>Owner ID:&nbsp;<span id="spanOID"><%=Request.QueryString("searchOID")%></span></td></tr>
<tr>
	<td >Title:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="20" TYPE="TEXT" NAME="TxtTitle" VALUE="<%=Title%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Name First:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="20" TYPE="TEXT" NAME="TxtNameFirst" VALUE="<%=name_first%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Name Last:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="25" TYPE="TEXT" NAME="TxtNameLast" VALUE="<%=name_last%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Address Line1:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="25" TYPE="TEXT" NAME="TxtAddress1" VALUE="<%=ADD1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Address Line2:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="25" TYPE="TEXT" NAME="TxtAddress2" VALUE="<%=ADD1%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
<tr>	
	<td >City:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="20" TYPE="TEXT" NAME="TxtCity" VALUE="<%=CITY%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >State:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="2" TYPE="TEXT" NAME="TxtState" VALUE="<%=STATE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Zip:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="10" TYPE="TEXT" NAME="TxtZip" VALUE="<%=ZIP%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
  	<td >FIPS:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="10" TYPE="TEXT" NAME="TxtFips" VALUE="<%=WPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >County:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="15" TYPE="TEXT" NAME="TxtCounty" VALUE="<%=HPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	
</tr>
<tr>
   	<td >Country:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="15" TYPE="TEXT" NAME="TxtCountry" VALUE="<%=FAX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Work Ph:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="15" TYPE="TEXT" NAME="TxtWorkPhone" VALUE="<%=WPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
    <td >Home Ph:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="15" TYPE="TEXT" NAME="TxtHomePhone" VALUE="<%=HPHONE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Fax:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="15" TYPE="TEXT" NAME="TxtFax" VALUE="<%=FAX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td >Email:<br><input ScrnInput="TRUE" MAXLENGTH="255" CLASS="LABEL" size="35" TYPE="TEXT" NAME="TxtEmail" VALUE="<%=FAX%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>




</table>

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr></tr>
<tr></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» AHS Owner</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>


<span class="Label" ID=SPANDATA>Retrieving...</span>
<fieldset id="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&HIDEATTACH=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:20;LEFT:0" id=BABtnControl type=text/x-scriptlet></OBJECT>
<iframe width=100% height=0 name="DataFrame" src="OwnerDetailsData.asp?<%=Request.QueryString%>">
</fieldset>

<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Owner selected.
</div>


<% End If %>

</form>
</body>
</html>


