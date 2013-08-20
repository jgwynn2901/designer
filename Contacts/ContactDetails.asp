<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	
	Dim cCID, cAHSID, isRequired
	
	cCID =  Request.QueryString("CID")
	cAHSID = Request.QueryString("AHSID")
	
	s_DisplayMsg = "Ready"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Contact Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">
window.returnValue = false;

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

<!--#include file="..\lib\Help.asp"-->

dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub UpdateCID(inCID)
	document.all.CID.value = inCID
	document.all.spanCID.innerText = inCID
	if document.all.spanCID.innerText <> "NEW" then
			document.body.setAttribute "IsThisRequired", "N"
	End If
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

Function GetCID
	if document.all.CID.value <> "NEW" then
	    GetCID = document.all.CID.value
	else
		GetCID = ""
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
	If CStr(document.body.getAttribute("IsThisRequired")) = "Y" Then
		f_CheckIsThisRequired = true
	Else
		f_CheckIsThisRequired = False
	End if
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	If document.all.CNT_NAME.value = "" Then
		MsgBox "Name is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	ValidateScreenData = true
End Function

Function ExeSave
	sResult = ""
	bRet = false
		
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.CID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "CONTACT_ID " & Chr(129) & document.all.CID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID" & Chr(129) & document.all.AHSID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME" & Chr(129) & document.all.CNT_NAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONE" & Chr(129) & document.all.CNT_PHONE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TYPE" & Chr(129) & document.all.CNT_TYPE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TITLE" & Chr(129) & document.all.CNT_TITLE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION" & Chr(129) & document.all.CNT_DESC.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "FAX" & Chr(129) & document.all.CNT_FAX.value & Chr(129) & "1" & Chr(128)		
		sResult = sResult & "EMAIL" & Chr(129) & document.all.CNT_EMAIL.value & Chr(129) & "1" & Chr(128)		
		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmDetails.Submit()
		bRet = true
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If
	
	ExeSave = bRet
	window.returnValue = true
End Function

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
</script>

</head>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=12 topmargin=0>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Contact </td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="ContactSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" NAME="AHSID" value="<%=Request.QueryString("AHSID")%>">
<input type="hidden" NAME="CID" value="<%=Request.QueryString("CID")%>">
<%	

If cCID <> "" Then
	If cCID <> "NEW" Then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		cSQL = "SELECT * FROM CONTACT WHERE CONTACT_ID=" & Request.QueryString("CID")
		Set oRS = oConn.Execute(cSQL)
		cName = trim(oRS("NAME"))
		cType = trim(oRS("TYPE"))
		cTitle = trim(oRS("TITLE"))
		cPhone = trim(oRS("PHONE"))
		cFax = trim(oRS("FAX"))
		cEMail = trim(oRS("EMAIL"))
		cDesc = trim(oRS("DESCRIPTION"))
		oRS.Close
		set oRS = nothing
		oConn.Close
		Set oConn = Nothing
	End If
End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=s_DisplayMsg%></span>
</td>
</tr>
</table>

<table CLASS="LABEL">
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr>
	<td COLSPAN="3">Contact ID:&nbsp;<span id="spanCID"><%=Request.QueryString("CID")%></span></td></tr>
</table>
<TABLE ID="Table3">
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT CLASS=LABEL NAME=CNT_NAME VALUE="<%=cName%>" TYPE=TEXT SIZE=40 MAXLENGTH=60 ID="Text1"></TD>
</TR>
<TR>
<TD CLASS=LABEL>Type:<BR><INPUT CLASS=LABEL NAME=CNT_TYPE VALUE="<%=cType%>" TYPE=TEXT SIZE=40 MAXLENGTH=40 ID="Text2"></TD>
<TD CLASS=LABEL>Title:<BR><INPUT CLASS=LABEL NAME=CNT_TITLE VALUE="<%=cTitle%>" TYPE=TEXT SIZE=40 MAXLENGTH=80 ID="Text3"></TD>
</TR>
<TR>
<TD CLASS=LABEL>Phone:<BR><INPUT CLASS=LABEL NAME=CNT_PHONE VALUE="<%=cPhone%>" TYPE=TEXT SIZE=14 MAXLENGTH=14 ID="Text4"></TD>
<TD CLASS=LABEL>Fax:<BR><INPUT CLASS=LABEL NAME=CNT_FAX VALUE="<%=cFax%>" TYPE=TEXT SIZE=10 MAXLENGTH=10 ID="Text5"></TD>
</TR>
<TR>
<TD CLASS=LABEL>E-Mail:<BR><INPUT CLASS=LABEL NAME=CNT_EMAIL VALUE="<%=cEMail%>" TYPE=TEXT SIZE=40 MAXLENGTH=255 ID="Text6"></TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2>Description:<BR><INPUT CLASS=LABEL NAME=CNT_DESC VALUE="<%=cDesc%>" TYPE=TEXT SIZE=85 MAXLENGTH=2000 ID="Text7"></TD>
</TR>
</TABLE>

</form>
</body>
</html>


