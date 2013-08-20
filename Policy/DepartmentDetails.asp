<%
'***************************************************************
'form for Department data entry.
'
'$History: DepartmentDetails.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 1:39p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Policy
'* Added Department Interface due to ESIS Project.  It allows User to
'* create Department record attached to the AHSID in PROD Designer. The
'* permission used is the same as for Branch.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 12:10p
'* Created in $/FNS_DESIGNER/Source/Designer/Policy
'* Added Department Interface due to the ESIS Project.  It allows user to
'* attach AHSID to the department record.  Also, it allows user to delete,
'* create a new record and Edit an record in PROD Designer.  Permission
'* setup is the same as for Branch.  
'* 

%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	Response.Buffer = true
	RuleTextLen = 30
	DEPTID = CStr(Request.QueryString("DEPTID"))
	RSAHSID = CStr(Request.QueryString("AHSID"))
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Department Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
var g_StatusInfoAvailable = false;

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

function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}

var AHSSearchObj = new CAHSSearchObj();
</script>
<script language =vbscript >
Sub window_onload
dim cInnerHTML

<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
	
<%	end if %>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "DepartmentSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub UpdateDEPTID(inDEPTID)
	document.all.DEPTID.value = inDEPTID
	document.all.spanDEPTID.innerText = inDEPTID
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

Function GetDEPTID
	if document.all.DEPTID.value <> "NEW" then
		GetDEPTID = document.all.DEPTID.value
	else
		GetDEPTID = ""
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
	If  document.all.AHSID_ID.innertext = "" then
		MsgBox "AHS ID is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	if IsNumeric(document.all.AHSID_ID.innertext) = false then
		MsgBox "Please enter a number in the AHS ID field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.TxtDeptName.value = "" then
		MsgBox "Department Name is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	If  document.all.TxtDEPTCODE.value = "" then
		MsgBox "Department Code is a required field.",0,"FNSNetDesigner"
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
	
	if document.all.DEPTID.value = "" then
		ExeCopy = false
		exit function
	end if

	document.all.DEPTID.value = "NEW"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function AttachAccount (ID, SPANID)
	AHSID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.AHSIDName = SPANID
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No location currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_BRANCH_ASSIGNMENT&SELECTONLY=TRUE&AHSID=" &AHSID
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

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText(SPANID,inText)
	If Len(inText) < <%=RuleTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid (inText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Function ExeSave
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.DEPTID.value = "" then
		ExeSave = false
		exit function
	end if
	

		If document.all.DEPTID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		
		sResult = sResult & "DEPARTMENT_CODES_ID"& Chr(129) & document.all.spanDEPTID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DEPARTMENT_NAME"& Chr(129) & document.all.TxtDEPTNAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DEPARTMENT_CODE"& Chr(129) & document.all.TxtDEPTCODE.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.INNERTEXT & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
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

sub SetBranchTypeFieldReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("SpecialFilterBtn") = "TRUE" then
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
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Department Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmDetails" METHOD="POST" ACTION="DepartmentSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="SearchDEPTCode" value="<%=Request.QueryString("SearchDEPTCode")%>">
<input type="hidden" name="SearchDeptName" value="<%=Request.QueryString("SearchDeptName")%>">
<input type="hidden" name="SearchAHSID" value="<%=Request.QueryString("SearchAHSID")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="DEPTID" value="<%=Request.QueryString("DEPTID")%>">
<%	

If DEPTID <> "" Then
	If DEPTID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT DC.*, AHS.NAME FROM DEPARTMENT_CODES DC, ACCOUNT_HIERARCHY_STEP AHS"
		SQLST = SQLST & " WHERE DC.ACCNT_HRCY_STEP_ID = AHS.ACCNT_HRCY_STEP_ID AND DC.DEPARTMENT_CODES_ID = " & DEPTID
		Set RS = Conn.Execute(SQLST)
		If Not RS.EOF then 
			RSDEPTCODE = ReplaceQuotesInText(RS("DEPARTMENT_CODE"))
			RSDEPTNAME = ReplaceQuotesInText(RS("DEPARTMENT_NAME"))
			RSAHSID = RS("ACCNT_HRCY_STEP_ID")
			RSAHSID_TEXT = ReplaceQuotesInText(RS("NAME"))
		end if	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	end if	
	
	
%>		
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

<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">

	<tr><td COLSPAN="5">Department ID:&nbsp;<span id="spanDEPTID"><%=Request.QueryString("DEPTID")%></span></td></tr>
	<tr></tr>
	<tr><td>
	<IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Location" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	<IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Location" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
	
	</td>
		
	<td width=305 nowrap>Location:&nbsp;<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE ="<%=RSAHSID_TEXT%>"><%=TruncateText(RSAHSID_TEXT,RuleTextLen)%></SPAN></td>
	<td>A.H.Step ID:&nbsp;<SPAN ID=AHSID_ID CLASS=LABEL><%=RSAHSID%></SPAN><input name=TxtAHSID type=hidden value="<%=RSAHSID%>" ID="Hidden1"></input></td>


	<tr>
	
	</tr> 
	<tr>
	<td COLSPAN="2">Department Name:<br><input size="25" ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="30" TYPE="TEXT" NAME="TxtDeptName" VALUE="<%=RSDEPTNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	
	<td>Department Code:<br><input ScrnInput="TRUE" size="10" CLASS="LABEL" MAXLENGTH="10" TYPE="TEXT" NAME="TxtDEPTCODE" VALUE="<%=RSDEPTCODE%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	
	</tr>
</table>


<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
No Department selected.
</div>

<% End If %>

</form>
</body>
</html>


