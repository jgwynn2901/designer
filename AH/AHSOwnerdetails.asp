<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->

<%	Response.Expires=0
	Response.AddHeader  "Pragma", "no-cache"
	
	Dim OID, AHSOID, isRequired
	
	OID =  CStr(Request.QueryString("OID"))
	AHSOID =  CStr(Request.QueryString("AHSOID"))
	s_DisplayMsg = "Ready"

	
	AccountTextLen = 20
	RuleTextLen = 30
%>
<html>
<head>
<!--#include file="..\lib\tablecommon.inc"-->
<meta name="VI60_defaultClientScript" content="VBScript">
<title>AHS Owner Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">

function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}

var AHSSearchObj = new CAHSSearchObj();

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
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>
End Sub
Function AttachAccount (ID, SPANID)
	
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.AHSIDName = SPANID.title
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No account currently attached.",0,"FNSNetDesigner"
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


Sub UpdateRuleText (SPANID)
	If Len(RuleSearchObj.RIDText) < <%=RuleTextLen%> Then
		SPANID.innertext = RuleSearchObj.RIDText
	Else
		SPANID.innertext = Mid ( RuleSearchObj.RIDText, 1, <%=RuleTextLen%>) & " ..."
	End If
	SPANID.title = RuleSearchObj.RIDText
End Sub


Sub UpdateAHSOID(inAHSOID)
	document.all.AHSOID.value = inAHSOID
	document.all.spanAHSOID.innerText = inAHSOID
	if document.all.spanAHSOID.innerText <> "NEW" then
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

Function GetAHSOID
	if document.all.AHSOID.value <> "NEW" then
		GetAHSOID = document.all.AHSOID.value
	else
		GetAHSOID = ""
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
    
       
	If  document.all.AHS_ID.innerText = "" then
	    msgbox("Please attach the AHS Owner to an Account")
		ValidateScreenData = false
		exit Function
	End If

	if document.all.TxtACTSTDT.value <> "" then 
    
       if NOT checkDate(document.all.TxtACTSTDT.value) then
	      msgbox("Active Start Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF)
	      f_ValidateScreenData = False
	   end if
	end if
    if document.all.TxtACTEDT.value <> "" then 
       if NOT checkDate(document.all.TxtACTEDT.value) then
	      msgbox("Active End Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF)
	      f_ValidateScreenData = False
	   end if
	end if

	ValidateScreenData = true
End Function


Function ExeSave
	sResult = ""
	bRet = false
	if NOT ValidateScreenData  then 
			ExeSave = false
			exit function
	end if
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	End If
	
	If document.all.OID.value = "" Then
		ExeSave = false
		exit function
	End If
	
	
'	If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.AHSOID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		Else
			document.all.TxtAction.value = "UPDATE"
		End If
		
		if document.all.TxtAction.value = "UPDATE" then
		   sResult = sResult & "AHS_OWNER_ID"& Chr(129) & document.all.AHSOID.value & Chr(129) & "1" & Chr(128)
           sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHS_ID.innerText & Chr(129) & "1" & Chr(128)
		   sResult = sResult & "ACTIVE_START_DT"& Chr(129) & "to_date('" & document.all.TxtACTstDT.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		   sResult = sResult & "ACTIVE_END_DT"& Chr(129) & "to_date('" & document.all.TxtACTEDT.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
        end if
        if document.all.TxtAction.value = "INSERT" then
		   sResult = sResult & "AHS_OWNER_ID"& Chr(129) & document.all.AHSOID.value & Chr(129) & "1" & Chr(128)
		   sResult = sResult & "OWNER_ID"& Chr(129) & document.all.OID.value & Chr(129) & "1" & Chr(128)
           sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHS_ID.innerText & Chr(129) & "1" & Chr(128)
           sResult = sResult & "ACTIVE_START_DT"& Chr(129) & "to_date('" & document.all.TxtACTstDT.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
		   sResult = sResult & "ACTIVE_END_DT"& Chr(129) & "to_date('" & document.all.TxtACTEDT.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)		   
        end if
        
		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		'msgbox( "sResult = " & sResult)
		document.all.FrmDetails.Submit()
		bRet = true
'	Else
'		SpanStatus.innerHTML = "Nothing to Save"
'	End If
	
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

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If
End Sub
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>" IsThisRequired="<%=isRequired%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 AHS Owner Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<form Name="FrmDetails" METHOD="POST" ACTION="AHSOwnerSave.asp" TARGET="hiddenPage">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="OID" value="<%=Request.QueryString("OID")%>" >
<input type="hidden" NAME="AHSOID" value="<%=Request.QueryString("AHSOID")%>" >

<%	

If AHSOID <> "" Then
	If AHSOID <> "NEW" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		
		SQLST = "SELECT AHSO.*, AHS.NAME  " &_
				"  FROM  AHS_OWNER AHSO, ACCOUNT_HIERARCHY_STEP  AHS " &_
				" WHERE AHSO.ACCNT_HRCY_STEP_ID = AHS.ACCNT_HRCY_STEP_ID AND AHS_OWNER_ID = " & AHSOID
		'RESPONSE.WRITE(SQLST)
		Set RS = Conn.Execute(SQLST)
	
		If Not RS.EOF Then
			RSAHSOWNERID    = RS("AHS_OWNER_ID")
			RSAHSID         = RS("ACCNT_HRCY_STEP_ID")
			RSAHSNAME       = RS("NAME")
			RSACTIVESTARTDT = RS("ACTIVE_START_DT")
			RSACTIVEENDDT   = RS("ACTIVE_END_DT")
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If
End If
%>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL><%=s_DisplayMsg%></SPAN>
</td>
</tr>
</table>

<table CLASS="LABEL" >
   
   <tr>
	<td COLSPAN=3>Owner ID:&nbsp<span id="spanOID"><%=Request.QueryString("OID")%></span></td>
   </tr>
   <tr>
	<td COLSPAN=3>AHS Owner ID:&nbsp<span id="spanAHSOID"><%=Request.QueryString("AHSOID")%></span></td>
   </tr>
   <TR>
     <TD><IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHS_ID, AHS_NAME">
	 <IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHS_ID, AHS_NAME"></TD>
     <td nowrap  WIDTH=200>Account:&nbsp<SPAN ID=AHS_NAME CLASS=LABEL TITLE="<%=RSAHSNAME%>" ><%=TruncateText(RSAHSNAME,AccountTextLen)%></SPAN></td>
	 <td Nowrap>A.H.Step ID:<span ID=AHS_ID CLASS="LABEL"><%=RSAHSID%></span></td>	  
   </TR>
   <tr>
	<td nowrap >Active Start Dt:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=15 size=11 TYPE="TEXT" NAME="TxtACTSTDT" VALUE="<%=RSACTIVESTARTDT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	<td nowrap >Active End Dt:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH=15 size=11 TYPE="TEXT" NAME="TxtACTEDT" VALUE="<%=RSACTIVEENDDT%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
   </tr>
</table>


</form>
</body>
</html>


