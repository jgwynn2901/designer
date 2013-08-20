<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->

<%	Response.Expires=0 %>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Network Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
var g_StatusInfoAvailable = false;

function CVendorSearchObj()
{
	this.VID = "";
	this.Selected = false;
}
var VendorSearchObj = new CVendorSearchObj();

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
Sub window_onload
<%
	if CStr(Request.QueryString("MODE")) = "RO" then %>
		SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>

End Sub

Sub PostTo(strURL)
	FrmDetails.action = "NetworkSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateNID(inNID)
	document.all.NID.value = inNID
	document.all.spanNID.innerText = inNID
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

Function GetNID
	if document.all.NID.value <> "NEW" then
		GetNID = document.all.NID.value
	else
		GetNID = ""
	end if 
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
	dim cErrMsg

	cErrMsg = ""
	ValidateScreenData = true
	If document.all.TxtName.value = "" then
		cErrMsg = "Network name is a required field." & vbCRLF
	end if
	if len(cErrMsg) <> 0 then
		MsgBox cErrMsg, 0, "FNSDesigner"
		ValidateScreenData = false
	end if
End Function

Function ExeSave
	sResult = ""
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.NID.value = "" then
		ExeSave = false
		exit function
	end if
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.NID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "NETWORK_ID"& Chr(129) & document.all.spanNID.innerText & Chr(129) & "0" & Chr(128)

		sResult = sResult & "NAME"& Chr(129) & document.all.TxtName.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDescription.value & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If

	ExeSave = bRet
End Function

Sub Refresh
	NID = document.all.NID.value
	document.all.tags("IFRAME").item("DataFrame").src = "NetworkData.asp?NID=" & NID
End Sub

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


Sub ExeButtonsNew()

	If document.all.NID.value = "" Or document.all.NID.value = "NEW" Then
		Exit Sub
	End If

	VendorSearchObj.Selected = false
	strURL = "VendorMaintenance.asp?SECURITYPRIV=FNSD_NETWORKS&CONTAINERTYPE=MODAL"
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog strURL, VendorSearchObj, "center"

	If VendorSearchObj.Selected Then 
		sResult = VendorSearchObj.VID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "ADD VENDOR"
		FrmDetails.action = "NetworkSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	End If
End Sub

Sub ExeButtonsRemove
	If document.all.NID.value = "" Or document.all.NID.value = "NEW" Then
		Exit Sub
	End If

	dim VID, sResult
	VID = document.frames("DataFrame").GetSelectedVID
	
	If VID <> "" Then
		sResult = VID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE VENDOR"
		FrmDetails.action = "NetworkSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select a Vendor to remove.", 0, "FNSDesigner"		
	End If

	Exit Sub
End Sub


Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>

<SCRIPT LANGUAGE="JScript" FOR="NetworkControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "REMOVEBUTTONCLICK":
				ExeButtonsRemove();
			break;
		case "NEWBUTTONCLICK":
				ExeButtonsNew();
			break;
		default:
			break;
	}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Network Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="NetworkSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchNetworkID" value="<%=Request.QueryString("SearchNetworkID")%>">
<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="NID" value="<%=Request.QueryString("NID")%>">
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
<%	
dim oConn, oRS, cSQL, VID

NID = Request.QueryString("NID")
If NID <> "" Then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	If NID <> "NEW" then
		cSQL = "SELECT * FROM NETWORK WHERE NETWORK_ID = " & NID
		Set oRS = oConn.Execute(cSQL)
		If Not oRS.EOF then 
			RS_NAME = ReplaceQuotesInText(oRS("NAME"))
			RS_DESCRIPTION = ReplaceQuotesInText(oRS("DESCRIPTION"))
		end if	
		oRS.Close
		Set oRS = Nothing
	end if	
%>		
<table class="LABEL">
<tr>
<td COLSPAN="6">Network ID:&nbsp;<span id="spanNID"><%=Request.QueryString("NID")%></span></td>
</tr> 

<tr>
<td >Name:<br><input size="25" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtName" VALUE="<%=RS_NAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
<td>Description:<br><input ScrnInput="TRUE" size="80" CLASS="LABEL" MAXLENGTH="120" TYPE="TEXT" NAME="TxtDescription" VALUE="<%=RS_DESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
</tr>
</table>
<OBJECT data="../Scriptlets/ObjButtons.asp?NEWCAPTION=Add&HIDESEARCH=TRUE&HIDEEDIT=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=TRUE&HIDEATTACH=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=NetworkControl type=text/x-scriptlet></OBJECT>
<iframe FRAMEBORDER="0" width=100% height=71% name="DataFrame" src="NetworkData.asp?<%=Request.QueryString%>">

<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Network selected.
</div>
<% End If %>
</form>
</body>
</html>


