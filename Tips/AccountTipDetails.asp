<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	
	Dim ATID, ATLID, isRequired	
	ATID =  CStr(Request.QueryString("ATID"))
	ATLID =  CStr(Request.QueryString("ATLID"))	
	s_DisplayMsg = "Ready"
	BranchTextLen = 30
	RuleTextLen = 30
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Account Tip Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">
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

Sub window_onload
<%if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>
End Sub

Sub UpdateATLID(inATLID)
	document.all.ATLID.value = inATLID
	document.all.spanATLID.innerText = inATLID
	if document.all.spanATLID.innerText <> "NEW" then
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

Function GetATLID
	if document.all.ATLID.value <> "NEW" then
	    GetATLID = document.all.ATLID.value
	else
		GetATLID = ""
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
		
	If document.all.TxtSequence.value <> "" Then
		If not IsNumeric(document.all.TxtSequence.value) then
			MsgBox "Please enter a number(1-12) in the Sequence field.",0,"FNSNetDesigner"
			ValidateScreenData = false
			exit Function
		end if
	Else
		MsgBox "Sequence is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if	
	
	If document.all.TxtTip.value = "" Then
		MsgBox "Tip is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	

	ValidateScreenData = true
End Function


Function ExeSave
	sResult = ""
	bRet = false
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	End If
	
	If document.all.ATID.value = "" Then
		ExeSave = false
		exit function
	End If
		
	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.ATLID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "Account_Tip_List_ID " & Chr(129) & document.all.ATLID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "Account_Tip_ID " & Chr(129) & document.all.ATID.value & Chr(129) & "1" & Chr(128)
		
		sResult = sResult & "TIP_SEQUENCE" & Chr(129) & document.all.TxtSequence.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "TIP_DESCRIPTION" & Chr(129) & document.all.TxtTip.value & Chr(129) & "1" & Chr(128)
		if(document.all.chkEnabled.Checked) then
			sResult = sResult & "ENABLED_FLG"& Chr(129) & "Y" & Chr(129) & "1" & Chr(128)
		else
			sResult = sResult & "ENABLED_FLG"& Chr(129) & "N" & Chr(129) & "1" & Chr(128)
		end if
		
        
		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmDetails.Submit()
		bRet = true
	Else
		SpanStatus.innerHTML = "Nothing to Save"
	End If
	
	ExeSave = bRet
	window.returnValue = true
End Function

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
	end if
end sub

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
</script>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>" IsThisRequired="<%=isRequired%>">
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="2" height="4">
            </td>
        </tr>
        <tr>
            <td class="GrpLabel" width="134" height="10">
                <nobr>&nbsp;» Account Tip Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8">
            </td>
            <td height="5" align="LEFT">
                <table cellpadding="0" cellspacing="0" height="100%">
                    <tr>
                        <td width="3" height="4">
                        </td>
                        <td width="300" height="4">
                        </td>
                    </tr>
                    <tr>
                        <td class="GrpLabelDrk" width="3" height="8" valign="BOTTOM" align="LEFT">
                        </td>
                        <td width="300" height="8">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td class="GrpLabelLine" colspan="2" height="1">
            </td>
        </tr>
        <tr>
            <td colspan="2" height="1">
            </td>
        </tr>
    </table>

<form Name="FrmDetails" METHOD="POST" ACTION="AccountTipSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="ATID" value="<%=Request.QueryString("ATID")%>">
<input type="hidden" NAME="ATLID" value="<%=Request.QueryString("ATLID")%>">

<%	

If ATLID <> "" Then
	If ATLID <> "NEW" Then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		SQLST = "SELECT * FROM ACCOUNT_TIP_LIST WHERE ACCOUNT_TIP_LIST_ID = "& ATLID 				
		Set oRS = oConn.Execute(SQLST)
	    If Not oRS.EOF Then
			RS_TIP = oRS("TIP_DESCRIPTION")
			RS_SEQUENCE = oRS("TIP_SEQUENCE")			
		End If
		oRS.Close
		Set oRS = Nothing
		oConn.Close
		Set oConn = Nothing
	End If
End If
%>
<table language="JScript" ondragstart="return false;" style="{position: absolute;
    top: 20; }" class="Label">
    <tr>
        <td valign="CENTER" width="5">
            <img id="StatusRpt" src="..\images\StatusRpt.gif" width="16" height="16" valign="CENTER"
                alt="View Status Report">
        </td>
        <td width="485">
            :<span valign="CENTER" id="SpanStatus" style="color: #006699" class="LABEL"><%=s_DisplayMsg%></span>
        </td>
    </tr>
</table>

<table class="LABEL">    
    <tr>
        <td colspan="3">
            Account Tip List ID:&nbsp;<span id="spanATLID"><%=Request.QueryString("ATLID")%></span>
            <br />
        </td>
    </tr>
    <tr>
        <td colspan="1">
            Sequence: <input class="Label" maxlength="2" size="5" type="text" value="<%=RS_SEQUENCE %>" 
            name="TxtSequence"  onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange"/>
        </td>
        <td>
            Enabled: <input type="checkbox" name="chkEnabled" checked="<%=RS_Enabled %>" class="Label"  onchange="VBScript::Control_OnChange" />
        </td><td></td>
    </tr>
    </tr>
    <tr>
        <td width="85%" colspan="3">
            Tip :<br>
            <input scrninput="TRUE" class="LABEL" maxlength="255" size="80" type="TEXT" name="TxtTip"
                value="<%=RS_TIP%>" onkeypress="VBScript::Control_OnChange" onchange="VBScript::Control_OnChange">
        </td>        
    </tr>    
</table>
</form>
</body>
</html>


