<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<%	Response.Expires=0 
	Dim PID	
	PID =  CStr(Request.QueryString("PID"))
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Jurisdiction State Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JScript">
var allStates = "*";

function loadStates(){
	var val = document.getElementsByName('selectedJurisStates');	
	var states = new Array();
	if (val[0].value != ""){		
		states = val[0].value.split(",");
		document.frames("JurisStateFrame").document.frames.setSelectedStates(states);
	}
}

function BuildInsertScript(strPIDValue){
	var strScript = "";	
	var newSelectedStates = new Array();	
	newSelectedStates = document.frames("JurisStateFrame").document.frames.getSelectedStates();	
	if (newSelectedStates == ""){
		return;
	}	
	for(var i=0; i < newSelectedStates.length; i++){			
			strScript = strScript + "INTO JURISDICTION_STATE (POLICY_ID, STATE) VALUES (" + strPIDValue + ",'" + newSelectedStates[i] + "') ";
	}		
	strScript = strScript.substring(0,strScript.length - 1);
	return strScript;
}
</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload		
	loadStates
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

Function ExeSave
	sResult = ""
	bRet = false	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	End If	
	If document.all.PID.value = "" Then
		ExeSave = false
		exit function
	End If								
	sResult = BuildInsertScript(document.all.PID.value)
	If sResult <> "" Then
		document.all.TxtAction.value = "INSERT_ALL"				
		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmDetails.Submit()			
	Else			
		sResult = "POLICY_ID=" & document.all.PID.value
		document.all.TxtAction.value = "DELETE"							
		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmDetails.Submit()					
	End If 						
	ExeSave = True
End Function

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
	end if
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
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR>
<TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Jurisdiction States for Policy ID "<%=Request.QueryString("PID")%>"</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<form Name="FrmDetails" METHOD="POST" ACTION="PolicyJurisStateSave.asp" TARGET="hiddenPage" ID="Form1">
<INPUT TYPE="HIDDEN" NAME="TxtSaveData">
<INPUT TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="PID" value="<%=Request.QueryString("PID")%>">
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr><td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report"></td>
<td width="485">:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td></tr>
</table>
<iframe FRAMEBORDER="0" ID="JurisStateFrame" SRC="JurisStates.asp?<%=Request.QueryString%>" WIDTH="100%" HEIGHT="90%"></iframe>

<%
	If PID <> "NEW" And PID <> "" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
				
		SQLST = "SELECT * FROM JURISDICTION_STATE WHERE POLICY_ID = " & PID 
		Set RS = Conn.Execute(SQLST)		
		Dim selectedJurisStates
		selectedJurisStates = ""		
		
		Do While Not RS.EOF And Not RS.BOF
			selectedJurisStates = selectedJurisStates & trim(RS("STATE")) & ","		
			RS.MoveNext		
		Loop						
		
		if (selectedJurisStates <> "") Then
			selectedJurisStates = Left(selectedJurisStates,  Len(selectedJurisStates) - 1)
		end if		
%>
	<Input type="hidden" name="selectedJurisStates" value="<%=selectedJurisStates%>"/>
<%	
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End If		
%>

</form>

</body>
</html>


