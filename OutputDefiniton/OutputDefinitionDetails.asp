<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\CheckSharedOD.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<%Response.Expires=0 
	Dim SharedCount, SharedCountText, ODID
	SharedCountText = "Ready"
	
	ODID	= CStr(Request.QueryString("ODID"))
	If ODID <> "" Then
		If ODID = "NEW" Then 
			SharedCount = 0
		else
			SharedCount = CheckSharedOD(Request.QueryString("ODID"), True, True, 1, False, False, 0)	
		End If
	End If	

If ODID <> "" Then
	If ODID <> "NEW" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = ""
		SQLST = SQLST  & "SELECT * FROM OUTPUT_DEFINITION WHERE OUTPUTDEF_ID = " & ODID
		SQLST2 = ""
		SQLST2 = SQLST2 & "SELECT * FROM OUTPUT_PAGE WHERE OUTPUTDEF_ID=" & ODID & " ORDER BY PAGE_NUMBER "
		Set RS = Conn.Execute(SQLST)
		Set RS2 = Conn.Execute(SQLST2)
		If Not RS.EOF then
			RSOUTPUTDEF_ID = RS("OUTPUTDEF_ID")
			RSNAME = ReplaceQuotesInText(RS("NAME"))
			RSDESCRIPTION = ReplaceQuotesInText(RS("DESCRIPTION"))
			RSDUPLEX_PRINT_FLG = RS("DUPLEX_PRINT_FLG")
		end if	
	end if	
End If
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Output Definition Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="JavaScript">
<!--
function CVehicleSearchObj()
{
	this.VehicleID = "";
}
var VehicleObj = new CVehicleSearchObj();

function COutputPageSearchObj()
{
this.OPID = "";
this.OPIDName = "";
}
var OutputPageSearchObj = new COutputPageSearchObj();
-->
</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
		SetScreenFieldsReadOnly true,"DISABLED"
<%	else 
		if ODID <> "" then %>
		<% If RSDUPLEX_PRINT_FLG = "Y" Then %>
			document.all.TxtDUPLEX_PRINT_FLG.checked = True
		<% End If %>
			<% if SharedCount <= 1 then %>
					document.all.ChkEdit.checked = true
					ChkEdit_OnClick
			<%	else %>
					document.all.ChkEdit.checked = false
					ChkEdit_OnClick
					SetStatusInfoAvailableFlag(true)
				<%	SharedCountText = "<SPAN STYLE='COLOR:#FF0000'>Warning!</SPAN> Shared Count is greater than 1." 
					If CInt(SharedCount) = CInt(Application("MaximumSharedCount")) Then %>
						document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>" & "<Font size=1 Color='Maroon'>+</Font>"
					<%Else %>
						document.all.SpanSharedCount.innerHTML = "<%=SharedCount%>"
					<%End If
				end if
		end if	
	end if 
If ODID <> "" AND ODID <> "NEW" Then
If Session("CONTAINERTYPE") = "FRAMEWORK" Then %>
ExpandTo = "ODID=<%= RS("OUTPUTDEF_ID") %>"
NodeX = TreeView1.AddNode ("",1 , "ODID=<%= RS("OUTPUTDEF_ID") %>", "PAGE", "Output Definition: <%= RS("NAME") %>", "OUTPUTDEFINITION", "OUTPUTDEFINITIONSEL")
<% Do While Not RS2.EOF %>
NodeX = TreeView1.AddNode ("ODID=<%= RS("OUTPUTDEF_ID") %>", 4 , "ODID=<%= RS("OUTPUTDEF_ID") %>&OPID=<%= RS2("OUTPUT_PAGE_ID") %>", "PAGE", "Output Page: <%= RS2("NAME") %> , <%= RS2("OUTPUT_PAGE_ID") %>" ,"PAGE", "PAGESEL" )
<% 
RS2.MoveNext
Loop
RS2.Close
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
%>
If TreeView1.Nodes(1).children >= 1 Then 
	lret = TreeView1.AddMenuItem("PAGE", "&Visual Editor", ErrStr)
	TreeView1.ExpandNode (ExpandTo) 
end If
	lret = TreeView1.AddMenuItem("PAGE", "&Add New Ouput Page", ErrStr)
<% End If %>
<% End If %>
End Sub

Sub PostTo(strURL)
	FrmDetails.action = "OutputDefinitionSearch-f.asp"
	FrmDetails.method = "GET"
	FrmDetails.target = "_parent"	
	FrmDetails.submit
End Sub

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
	Select Case MenuItem
		Case "&Visual Editor"
			Call LaunchODEditor(NodeKey)
		Case "&Add New Ouput Page"
			If document.all.ChkEdit.checked = true Then
				strurl = ""
				strurl = "OutputPageMaintenance.asp?CONTAINERTYPE=MODAL&DETAILONLY=TRUE&OPID=NEW&ODID=<%= ODID%>"
				lret = window.showModalDialog( strurl  ,OutputPageSearchObj ,"dialogWidth:450px;dialogHeight:450px;center")
				self.location.reload()
			Else
				msgbox "Edit mode not selected",0,"FNSDesigner"
			End If
			'Parent.Parent.location.href = "OutputPageMaintenance.asp?CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&OPID=NEW&ODID=<%= ODID%>"
	End Select 
End Sub

Function Handles(Obj, Title)
If InStr(1, top.frames("TOP").location.href, "Toppane.asp") <> 0 Then
	lret = top.frames("TOP").SetHandle(Obj, Title)
End If
End Function

Sub UpdateODID(inODID)
	document.all.ODID.value = inODID
	document.all.spanODID.innerText = inODID
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

Function GetODID
	if document.all.ODID.value <> "NEW" then
		GetODID = document.all.ODID.value
	else
		GetODID = ""
	end if 
End Function

Function GetODIDName
	GetODIDName = document.all.TxtName.value
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
	If  document.all.TxtNAME.value = "" then
		MsgBox "Name is a required field.",0,"FNSNetDesigner"
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
	
	if document.all.ODID.value = "" then
		ExeCopy = false
		exit function
	end if
	document.all.COPYODID.value = document.all.ODID.value
	document.all.ODID.value = "COPY"
	document.body.setAttribute "ScreenDirty", "YES"
	ExeCopy = ExeSave
End Function

Function ExeSave
	sResult = ""
	bRet = false
	
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	end if
	
	if document.all.ODID.value = "" then
		ExeSave = false
		exit function
	end if
	
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.ODID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		elseif document.all.ODID.value = "COPY" then
			document.all.TxtAction.value = "COPY"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
		sResult = sResult & "OUTPUTDEF_ID"& Chr(129) & document.all.ODID.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME"& Chr(129) & document.all.TxtNAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "DESCRIPTION"& Chr(129) & document.all.TxtDESCRIPTION.value & Chr(129) & "1" & Chr(128)
		If document.all.TxtDUPLEX_PRINT_FLG.checked = True Then
			DUP_FLG = "Y"
		Else
			DUP_FLG = "N"
		End If
		sResult = sResult & "DUPLEX_PRINT_FLG"& Chr(129) & DUP_FLG & Chr(129) & "1" & Chr(128)
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		bRet = true
	'Else
	'	SpanStatus.innerHTML = "Nothing to Save"
	'End If
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

sub ChkEdit_OnClick
	document.all.ChkEdit.setAttribute "ScrnBtn","FALSE"
	if document.all.ChkEdit.checked = true then
		SetScreenFieldsReadOnly false,"LABEL"
		document.body.setAttribute "ScreenMode", "RW"		
	else
		SetScreenFieldsReadOnly true,"DISABLED"
		document.body.setAttribute "ScreenMode", "RO"
	end if
	document.all.ChkEdit.setAttribute "ScrnBtn","TRUE"
end sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

Sub BtnStatus_onclick
	If CLng(<%=SharedCount%>) > 1 Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other details reported", 0 , "FNSNet"
	End If
End Sub

Sub BtnRefCount_onclick
	lret = window.showModalDialog ("..\StatusRpt\RefCountRpt.asp?CheckSharedOD=True&ID=<%=Request.QueryString("ODID")%>", Null,  "dialogWidth=580px; dialogHeight=400px; center=yes")
End Sub

<!--#include file="..\lib\Help.asp"-->

</script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
<% If ODID <> "NEW" Then %>
function LaunchODEditor(key) {
Url = "../RoutingPlan/OutputDefinitionEditor-f.asp?AHSID=<%= Request.QueryString("AHSID") %>&" + key
var VisEditorObj = window.open(Url, null, "height=500,width=750,status=no,toolbar=no,menubar=no,location=no,resizable=yes,top=0,left=0");
lret = Handles( VisEditorObj, "OUTPUT");
VisEditorObj.focus()
}
<% End If %>
//-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Output Defintion Details&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<form Name="FrmDetails" METHOD="POST" ACTION="OutputDefinitionSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">
<input type="hidden" name="SearchODID" value="<%=Request.QueryString("SearchODID")%>">
<input type="hidden" name="SearchNAME" value="<%=Request.QueryString("SearchNAME")%>">
<input type="hidden" name="SearchDESCRIPTION" value="<%=Request.QueryString("SearchDESCRIPTION")%>">
<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="ODID" value="<%=Request.QueryString("ODID")%>">
<input type="hidden" NAME="COPYODID" VALUE>
<%	

Function TruncateRuleText(inText)
	if not IsNull(inText) then
		If Len(inText) < 40 Then
			TruncateRuleText = inText
		Else
			TruncateRuleText = Mid ( inText, 1, 40) & " ..."
		End If
	end if
End Function

Function TruncateLookupText(inText)
	if not IsNull(inText) then
		If Len(inText) < 22 Then
			TruncateLookupText = inText
		Else
			TruncateLookupText = Mid ( inText, 1, 22) & " ..."
		End If
	end if
End Function

Function ReplaceRuleText(inText)
	if not IsNull(inText) then
		ReplaceRuleText = Replace(inText,"""","&quot;")
	end if
End Function
If ODID <> "" Then

%>
<table style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td CLASS="LABEL" width="16" height="16" valign="MIDDLE"><nobr><img SRC="..\images\RefCount.gif" STYLE="CURSOR:HAND" ID="BtnRefCount" align="absmiddle" TITLE="Shared Count">
<td CLASS="LABEL" width="16" height="16" valign="MIDDLE">
:<span id="SpanSharedCount"><%=SharedCount%></span> </td>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><%=SharedCountText%></span>
</td>
<td>
<input ScrnBtn="TRUE" TYPE="CHECKBOX" VALIGN="RIGHT" Name="ChkEdit">Edit
</td>
</tr>
</table>
<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
<table class="LABEL">
	<tr>
	<td COLSPAN="5" CLASS="LABEL">Output Definition ID:&nbsp;<span id="spanODID"><%=Request.QueryString("ODID")%></span></td>
	</tr>
	</table>
	<table>
	<tr>	
	<td CLASS="LABEL">Name:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="80" size="80" TYPE="TEXT" NAME="TxtNAME" VALUE="<%=RSNAME%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Description:<br><input ScrnInput="TRUE" CLASS="LABEL" MAXLENGTH="255" size="80" TYPE="TEXT" NAME="TxtDESCRIPTION" VALUE="<%=RSDESCRIPTION%>" ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange"></td>
	</tr>
	<tr>
	<td CLASS="LABEL"><input ScrnBtn="TRUE" CLASS="LABEL" TYPE="CHECKBOX" NAME="TxtDUPLEX_PRINT_FLG" ONCHANGE="VBScript::Control_OnChange" ONCLICK="VBScript::Control_OnChange">Duplex Print Flag:</td>
	</tr>
	</table>
	</table>
	</form>
<% If Session("CONTAINERTYPE") = "FRAMEWORK" Then %>	
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Output Definition Tree View</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<object CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0">
<param NAME="LPKPath" VALUE="LPKfilename.LPK"></object>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="50%">
<param NAME="ShowTips" VALUE="False">
</object>
<% End If %>
<% Else %>
<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
	No output defintion selected.
</div>
<% End If %>

</body>
</html>