<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	

Response.Expires = 0 
Response.AddHeader  "Pragma", "no-cache"
Response.Buffer = true
RuleTextLen = 30
Dim RSAHSID 
RSAHSID = Request.QueryString("AHSID")
	
	
Dim ATID,oRS,SQLST,RSDESCRIPTION,RSLOB
ATID = Request.QueryString("ATID")


Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
If ATID <> "" Then
	If ATID <> "NEW" Then
	    SQLST = "SELECT AT.*,AHS.NAME,R.RULE_TEXT FROM " 
		SQLST =SQLST & "ACCOUNT_TIP AT INNER JOIN ACCOUNT_HIERARCHY_STEP AHS on AT.ACCNT_HRCY_STEP_ID = AHS.ACCNT_HRCY_STEP_ID LEFT OUTER JOIN RULES R on R.RULE_ID = AT.RULE_ID" 
		SQLST =SQLST & " WHERE AT.ACCNT_HRCY_STEP_ID = "& RSAHSID 
		SQLST =SQLST & " AND ACCOUNT_TIP_ID =" & ATID 
				
		Set oRS = oConn.Execute(SQLST)
		If Not oRS.EOF Then
		    RSDESCRIPTION = ReplaceQuotesInText(oRS("DESCRIPTION"))
			RSLOB = oRS("LOB_CD")
			RSAHSID_TEXT = oRS("NAME")
			RSRULE_ID = oRs("RULE_ID")
			RSRULE_TEXT = oRS("RULE_TEXT")
			'RSTHRESHOLD = oRS("THRESHOLD")
		End If
		oRS.Close
		'Set oRS = Nothing
	End If
	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript"/>
<title>Account Tips</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css"/>
<script type="text/javascript" language="javascript">
function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
function RuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}


function TipSearchObj()
{
	this.Selected = false;
}

var TipSearchObj = new TipSearchObj();

var AHSSearchObj = new CAHSSearchObj();
var RuleSearchObj = new RuleSearchObj();
</script>
<script type="text/jscript" language="JScript" for="window" event="onload">
<%	If CStr(Request.QueryString("MODE")) = "RO" Then %>	
	SetScreenFieldsReadOnly(true,"DISABLED");
<%	End If %>
	if (document.all.DataFrame != null)
		document.all.DataFrame.style.height = document.body.clientHeight - 200;
	if (document.all.fldSet != null)
		document.all.fldSet.style.height = document.body.clientHeight - 180;
	if (document.all.SPANDATA != null)
		document.all.SPANDATA.innerText = "";
<%
If ATID <> "" Then
%>		
	document.all.LOB_CD.value = "<%= RSLOB %>"
<%
end if
%>	
</script>

<script id="clientEventHandlersVBS" language="vbscript">
Sub ExeNewBranchRule

	dim ATID, ATLID, MODE
	
	If Not InEditMode Then
		Exit Sub
	End If
	If document.all.ATID.value = "" Or document.all.ATID.value = "NEW" Then
		Exit Sub
	End If

	ATLID = "NEW"
	ATID = document.all.ATID.value
	MODE = document.body.getAttribute("ScreenMode")

		
	If document.all.TIPSCOUNT.value = 12 then
		MsgBox "12 Tips already entered for selected LOB.", 0, "FNSNet Designer"
		Exit Sub
	End if
			
	
	TipSearchObj.Selected = false
	strURL = "AccountTipModal.asp?ATID=" & ATID & "&ATLID=" & ATLID & "&MODE=" & MODE 	
	showModalDialog strURL, TipSearchObj, "center:yes;" 
	If TipSearchObj.Selected Then 
		Refresh
	End if
End Sub

Sub Refresh
	ATID = document.all.ATID.value
	document.all.tags("IFRAME").item("DataFrame").src = "TipsDetailsData.asp?ATID=" & ATID
End Sub

Function GetSelectedATLID
	GetSelectedATLID = document.frames("DataFrame").GetSelectedATLID
End Function

Sub ExeRemoveBranchRule
	dim ATLID, sResult

	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.ATID.value = "" Or document.all.ATID.value = "NEW" Then
		Exit Sub
	End If

	ATLID = GetSelectedATLID
	ATID = document.all.ATID.value
	
	If ATLID <> "" Then
		sResult = sResult & ATLID
		document.all.TxtSaveData.Value = sResult
		document.all.TxtAction.Value = "DELETE"
		FrmDetails.action = "AccountTipSave.asp"
		FrmDetails.method = "POST"
		FrmDetails.target = "hiddenPage"	
		FrmDetails.submit
		Refresh
	Else
		MsgBox "Please select a Tip  to Remove.", 0, "FNSNet Designer"		
	End If

	Exit Sub
End Sub


Sub ExeEditBranchRule
	dim ATLID, ATID

	If Not InEditMode Then
		Exit Sub
	End If

	If document.all.ATID.value = "" Or document.all.ATID.value = "NEW" Then
		Exit Sub
	End If

	ATLID = GetSelectedATLID
	ATID = document.all.ATID.value
	MODE = document.body.getAttribute("ScreenMode")
	
	If ATLID <> "" Then
		TipSearchObj.Selected = false
		strURL = "AccountTipModal.asp?ATID=" & ATID & "&ATLID=" & ATLID & "&MODE=" & MODE 	
		showModalDialog  strURL, TipSearchObj, "center:yes"
		If TipSearchObj.Selected Then
			Refresh
		End if
	Else
		MsgBox "Please select an  Item from Account Tip List to Edit.", 0, "FNSNet Designer"		
	End If
	
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then 
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Sub SetSuccessfulQueryCount(intValue)    
	document.all.SuccessfulQueryCount.value = CInt(document.all.SuccessfulQueryCount.value)+intValue
End Sub

Sub UpdateATID(inATID)
	document.all.ATID.value = inATID
	document.all.spanATID.innerText = inATID
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
	end if
end sub

Function ValidateScreenData
    if  document.all.AHSID_ID.innerText = "" then
        MsgBox "A.H.Step ID is required.",0,"FNSNetDesigner"
        ValidateScreenData = false
		exit Function
	end if

	If  document.all.TxtDescription.value = "" then
		MsgBox "Description is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if
	
	If  document.all.LOB_CD.value = "" then
		MsgBox "LOB is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	end if	
		
		
	ValidateScreenData = true
End Function

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

	If document.all.ATID.value = "" Then
		ExeSave = false
		Exit Function
	End If

	bRet = false
	
	if ValidateScreenData = false then 
		ExeSave = false
		exit function
	end if

	If document.all.ATID.value = "NEW" then
		document.all.TxtAction.value = "INSERT"
	else
		document.all.TxtAction.value = "UPDATE"
	end if
		
	sResult=""
	sResult = sResult & "ACCOUNT_TIP_ID"& Chr(129) & document.all.ATID.value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & "<%=RSAHSID%>" & Chr(129) & "1" & Chr(128)
	sResult = sResult & "RULE_ID"& Chr(129) & document.all.RULE_ID.InnerText & Chr(129) & "1" & Chr(128)
	sResult = sResult & "LOB_CD" & Chr(129) & document.all.LOB_CD.value & Chr(129) & "1" & Chr(128)
	'sResult = sResult & "SEQUENCE" & Chr(129) & document.all.SEQUENCE(i).value & Chr(129) & "1" & Chr(128)
	sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDescription.value & Chr(129) & "1" & Chr(128)
	'sResult = sResult & "TIP" & Chr(129) & document.all.TIP(i).value & Chr(129) & "1" & Chr(128)			
	
	document.all.TxtSaveData.Value = sResult
	FrmDetails.action = "TipsSave.asp"
	FrmDetails.method = "POST"
	FrmDetails.target = "hiddenPage"	
	FrmDetails.submit	
		
	bRet = true
	ExeSave = bRet
	
End Function

Function AttachRule (ID, SPANID, strTITLE)
	RID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	RuleSearchObj.RID = RID
	RuleSearchObj.RIDText = SPANID.title
	RuleSearchObj.Selected = false

	If RID = "" Then RID = "NEW"
	
	If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\Rules\RuleMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_MC_BRANCH_ASSIGNMENT&RID=" & RID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,RuleSearchObj ,"center"

	'if Selected=true update everything, otherwise if RuleID is the same, update text in case of save
	If RuleSearchObj.Selected = true Then
		If RuleSearchObj.RID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = RuleSearchObj.RID
		end if
		UpdateSpanText SPANID,RuleSearchObj.RIDText
	ElseIf ID.innerText = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
		UpdateSpanText SPANID,RuleSearchObj.RIDText
	End If

End Function

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
	
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_TIP&SELECTONLY=TRUE&AHSID=" &AHSID
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

Function GetATID
	if document.all.ATID.value <> "NEW" then
		GetATID = document.all.ATID.value
	else
		GetATID = ""
	end if 
End Function
</script>
<!--#include file="..\lib\BABtnControl.inc"-->
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" bgcolor="<%=BODYBGCOLOR%>" screendirty="NO"
    screenmode="<%=Request.QueryString("MODE")%>">
    <table width="100%" cellpadding="0" cellspacing="0">
         <tr>
            <td class="GrpLabel" width="134" height="10">
                <nobr>&nbsp;» Account Tip &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8">
            </td>
            <td height="5" align="LEFT">
                <table cellpadding="0" cellspacing="0" height="100%">
                    <tr>
                        <td width="3" height="4"></td>
                        <td width="300" height="4"></td>
                    </tr>
                    <tr>
                        <td class="GrpLabelDrk" width="3" height="8" valign="BOTTOM" align="LEFT"></td>
                        <td width="300" height="8"></td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td class="GrpLabelLine" colspan="2" height="1"></td>
        </tr>
        <tr>
            <td colspan="2" height="1"></td>
        </tr>
    </table>
    <form name="FrmDetails" method="POST" action="TipsSave.asp" target="hiddenPage">
    <input type="HIDDEN" name="TxtSaveData">
    <input type="HIDDEN" name="TxtAction">
    
<% 'need to maintain these values in order to post back to the search tab %>
<input type="hidden" name="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" name="ATID" value="<%=Request.QueryString("ATID")%>">
<input type="hidden" name="AHSID" value="<%=RSAHSID%>" ID="Hidden1">
<input type="hidden" name="TIPSCOUNT" value="0" ID="Hidden1">

    <table language="JScript" ondragstart="return false;" style="{position: absolute;
        top: 20; }" class="Label">
        <tr>
            <td valign="CENTER" width="5">
                <img id="StatusRpt" src="..\images\StatusRpt.gif" width="16" height="16" valign="CENTER"
                    alt="View Status Report">
            </td>
            <td width="485">:<span valign="CENTER" id="SpanStatus" 
                style="color: #006699" class="LABEL">Ready</span>                
                <input type="hidden" name="SuccessfulQueryCount" value="0" />
            </td>
        </tr>
    </table>
    </br>
    <table class="Label">		
		<tr>
			<td colspan="2">Account Tip ID:&nbsp;<span id="spanATID"><%=Request.QueryString("ATID")%></span></td>
		</tr>
		<tr><td></td></tr>
		<tr>
        <td>
            <img name="BtnAttachAHSID" style="cursor: hand" src="..\IMAGES\Attach.gif" title="Attach Account"
                onclick="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
            <img name="BtnDetachAHSID" style="cursor: hand" src="..\IMAGES\Detach.gif" title="Detach Account"
                onclick="VBScript::Detach AHSID_ID, AHSID_TEXT">
        </td>
        <td width="305" nowrap>
            Account:&nbsp;<span id="AHSID_TEXT" class="LABEL" title=""><%=TruncateText(RSAHSID_TEXT,RuleTextLen)%></span>
        </td>
        <td>
            A.H.Step ID:&nbsp;<span id="AHSID_ID" class="LABEL"><%=RSAHSID%></span><input name="TxtAHSID"
                type="hidden" value="<%=RSAHSID%>"></input>
        </td>
	</tr>
	<tr>
		<td>
			<IMG NAME=BtnAttachRule STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Rule" ONCLICK="VBScript::AttachRule RULE_ID, RULE_TEXT,''">
			<IMG NAME=BtnDetachRule STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Rule" OnClick="VBScript::Detach RULE_ID, RULE_TEXT">
		</td>
		<td width=305 nowrap>Rule Text:&nbsp;<SPAN ID=RULE_TEXT CLASS=LABEL TITLE=""><%=TruncateText(RSRULE_TEXT,RuleTextLen)%></SPAN></td>
		<td>Rule ID:&nbsp;<SPAN ID=RULE_ID CLASS=LABEL><%=RSRULE_ID%></SPAN></td>
	</tr>	
    </table>
    <table class="LABEL">       
        <tr>
			<td colspan="3"></td>
		</tr>				
       <!-- <tr>
            <td colspan="2">Tip ID:&nbsp;
                <span id="spanATID"><%=Request.QueryString("ATID")%></span>
            </td>
			<td></td>
        </tr> -->
        <tr>
		<td>LOB:<br><select style="width:160" name="LOB_CD" class="LABEL" 
                    onkeypress="VBScript::Control_OnChange"
                    onchange="VBScript::Control_OnChange">
                    <%
	cSQL = "SELECT * FROM LOB"
	Set oRS2 = oConn.Execute(cSQL)
	Do WHile Not oRS2.EOF
                    %>
                    <option value="<%= oRS2("LOB_CD") %>">
                        <%= oRS2("LOB_NAME") %>
                        <%
		oRS2.MoveNext
	Loop
	oRS2.Close
                        %>
                </select>
            </td>
            <td>Description:<br>
                <input scrninput="TRUE" maxlength="20" class="LABEL" size="65" 
                    type="TEXT" name="TxtDescription" value="<%=RSDESCRIPTION%>" 
                    onkeypress="VBScript::Control_OnChange" 
                    onchange="VBScript::Control_OnChange">
            </td>
            
        </tr>
        </table>
		</br>
	<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
	<tr>
		<td colspan="2" HEIGHT="4"></td>
	</tr>
	<tr>
		<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Account Tip List</td>
		<td	HEIGHT="5" ALIGN="LEFT">
	<table  CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
		<tr>
			<td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td>
		</tr>	
		<tr>
			<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
			<td	WIDTH="300" HEIGHT="8"></td>
		</tr>
	</table>
	</td>
</tr>
<tr>
	<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td>
</tr>
<tr>
	<td colspan="2" HEIGHT="1"></td>
</tr>
</table>


<span class="Label" ID="SPANDATA">Retrieving...</span>
<fieldset id="fldSet" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;width:'100%'">
<object data="../Scriptlets/ObjButtons.asp?HIDEREFRESH=TRUE&amp;HIDEATTACH=TRUE&amp;HIDESEARCH=TRUE&amp;HIDECOPY=TRUE&amp;HIDEPASTE=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="BABtnControl" type="text/x-scriptlet"></object>
<iframe width="100%" height="0" name="DataFrame" src="TipsDetailsData.asp?<%=Request.QueryString%>">
</fieldset>
	
   
	<% End If 
'Set oRS = Nothing
'oConn.Close
'Set oConn = Nothing
    %> 

    </form>
</body>
</html>
