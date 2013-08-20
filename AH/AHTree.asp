<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\AHSTree.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->

<%
'Response.Write("curAHSTreeTopID: " & session.Contents("curAHSTreeTopID") & "<HR>")
Response.Expires = 0
Response.Buffer = true
Session("AHSTreeProperties").RemoveAll

Dim bHighestLevelReached
dim cCmd
dim oRS
dim cConnStr

bHighestLevelReached = False
filterApplied = ""

if Request.QueryString("SHOWNODES") <> "" then
	addShowNodes(Request.QueryString("SHOWNODES"))
	AddExpandedNode(Request.QueryString("SHOWNODES"))
end if	

if Request.QueryString("SHOWALL") <> "" then
	addShowAllNode(Request.QueryString("SHOWALL"))
	parts = Split(Request.QueryString("SHOWALL"), "=")
	if Session("AHSTreeCurExpandedNodes").Item("AHSID=" & Request.QueryString("SHOWALL")) = "" Then
		AddExpandedNode(Request.QueryString("SHOWALL"))
	end if
end if	
If Request.QueryString("REFRESH") = "TRUE" Then
	cCmd = "AHTree.asp?MAXRS=" & Session("USERTREECOUNT") & "&MAXLEVEL=" &  Session("USERTREELEVELS") & "&AHSID=" & Request.QueryString("AHSID")
	Response.Redirect cCmd
End If

cConnStr = Session("ConnectionString")
Set oRS = Server.CreateObject("ADODB.Recordset")

Session("CurrentNode") = ""
If Request.QueryString("UPONELEVEL") <> "" Then
	parts = Split(Request.QueryString("UPONELEVEL"), "=")
	PARENTID = GetAHSParent(parts(1))

	If PARENTID <> "" Then
		Session("CurrentNode") = parts(1)
		Session("CurAHSTreeTopID") = PARENTID
		Session("CurAHSTreeTopDesc") = GetAHSDesc(PARENTID)
		If PARENTID <> "1" Then 
			Session("CurAHSTreeTopDesc") = Session("CurAHSTreeTopDesc") & "    [ " &PARENTID & " ]"
		end if

		'If PARENTID = "1" Then
		'	Session("CurrentMaxRecords") = 30
		'	Session("CurrentMaxLevel") = 1
		'End If
	Else
		bHighestLevelReached = True
	End If
End If

If Request.QueryString("EXPANDNODE") <> "" Then
	AddExpandedNode(Request.QueryString("EXPANDNODE"))
End If

If Request.QueryString("AHSID") <> "" Then
	Session("CurAHSTreeTopID") = Request.QueryString("AHSID")
End If

If Request.QueryString("AHSDesc") <> "" Then
	Session("CurAHSTreeTopDesc") = Request.QueryString("AHSDesc")
End If

If Request.QueryString("AHSDesc") = "UNKNOWN" Then
	Session("CurAHSTreeTopDesc") = GetAHSDesc(Request.QueryString("AHSID"))
	If Request.QueryString("AHSID") <> "1" Then 
		Session("CurAHSTreeTopDesc") = Session("CurAHSTreeTopDesc") &  "    [ " &Request.QueryString("AHSID") & " ]"
	end if
End If

If Request.QueryString("MAXRS") <> "" Then
	Session("CurrentMaxRecords") = CInt(Request.QueryString("MAXRS"))
End If

If Request.QueryString("MAXLEVEL") <> "" Then
	Session("CurrentMaxLevel") = CInt(Request.QueryString("MAXLEVEL"))
End If

If Session("CurrentMaxRecords") = "" Then
	oRS.MaxRecords = 30
	MaxLevel = 1
End If
oRS.MaxRecords = Session("CurrentMaxRecords")
MaxLevel = Session("CurrentMaxLevel")

Function GetAHSChildren(nParentID, NodeType, nLevel, lFirstCall)
Dim nCurLevel
dim lGetAll
dim cSQL
dim cNodes
dim lHasRootAccess
dim aNodes, x, cCurName
dim cInnerSQL
dim cWhereSQL
dim cIncludeSQL
dim cExcludeSQL
dim cMustInclude, cMustExclude
dim lIDOk

lGetAll = false
nCurLevel = nLevel + 1
If nCurLevel > MaxLevel And NodeType <> "RISK LOCATION" Then
%>
	NodeX = TreeView1.AddNode ("AHSID=<%=nParentID%>", 4 , "GRP=<%=nParentID%>", "EXPAND",  "..." ,"FOLDERGRP", "FOLDERGRPSEL")
<%		
	Exit Function
End If

aNodes = split( Session("ACCOUNT_SECURITY"), "," )
lHasRootAccess = false
aNodes = split( Session("ACCOUNT_SECURITY"), "," )
for x=lbound(aNodes) to ubound(aNodes)
	if aNodes(x) = "1" then
		lHasRootAccess = true
		exit for
	end if
next

if lFirstCall then
	cSQL = "SELECT AHS.Name AHSName, AHS.ACCNT_HRCY_STEP_ID AHSID, AHS.Type AHSType " & _
				"FROM ACCOUNT_HIERARCHY_STEP AHS " & _
				"WHERE (AHS.PARENT_NODE_ID = 1 AND AHS.ACTIVE_STATUS='ACTIVE') " & _
				"ORDER BY NAME"
	if not lHasRootAccess then
		cNodes = trim(Session("ACCOUNT_SECURITY"))
		if cNodes = "" then
			cNodes = "0"
		end if
		cSQL = "SELECT AHS.Name AHSName, AHS.ACCNT_HRCY_STEP_ID AHSID, AHS.Type AHSType " & _
					 "FROM ACCOUNT_HIERARCHY_STEP AHS " & _
					 "WHERE AHS.ACTIVE_STATUS='ACTIVE' AND AHS.ACCNT_HRCY_STEP_ID IN(" & cNodes & ") " & _
					 "ORDER BY NAME"
	end if 		
	oRS.MaxRecords = 999
	loadTree cSQL, nParentID, 0, true
	Exit Function
end if	
' has the User permit for these ID ?
'if not hasNodeAccess(nParentID, aNodes) then
'	Exit Function
'end if

cSQL = "SELECT AHS.Name AHSName, AHS.ACCNT_HRCY_STEP_ID AHSID, AHS.Type AHSType FROM ACCOUNT_HIERARCHY_STEP AHS WHERE (AHS.PARENT_NODE_ID =" & CStr(nParentID) & " AND AHS.ACTIVE_STATUS='ACTIVE'"
if Session("AHSTreeShowAllNodes").Item("AHSID=" & nParentID) <> "" then
	lGetAll = true
end if
If HasSpecificFilter("AHSID=" & nParentID, "DESIGNER_AHSFILTER") = true Then
	oRS.MaxRecords = 0
		
	cInnerSQL = ""
	cWhereSQL =""
	cIncludeSQL = ""
	cExcludeSQL = ""

	useWHERE = GetSpecificFilter("AHSID=" & nParentID, "DESIGNER_AHSFILTER", "USEWHERECLAUSE")	
	If useWhere = "TRUE" Then
		cWhereSQL = GetSpecificFilter("AHSID=" & nParentID, "DESIGNER_AHSFILTER", "WHERECLAUSE")
		cWhereSQL = replace(cWhereSQL, "#AND#", "AND")
		cInnerSQL = " AND " & cWhereSQL
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'Check if Special Character '*' is present in Filter Criteria ...
		'Replace it with '+'
		'Dated : 22/02/2007
		'ILOG Issue : JMCA-0128
		if (instr(cInnerSQL,"[[--]]")>0) then 
			   cInnerSQL = Replace(cInnerSQL,"[[--]]","+")
		End if 
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	End If			

	cMustExclude = GetSpecificFilter("AHSID=" & nParentID, "DESIGNER_AHSFILTER","MUSTEXCLUDE")
	If cMustExclude <> "" Then 
		cInnerSQL = cInnerSQL & " AND AHS.ACCNT_HRCY_STEP_ID NOT IN(" &  cMustExclude & ")" 
	end if

	cMustInclude = GetSpecificFilter("AHSID=" & nParentID,"DESIGNER_AHSFILTER", "MUSTINCLUDE")
	If cMustInclude <> "" Then 
		cIncludeSQL = " AND AHS.ACCNT_HRCY_STEP_ID IN(" &  cMustInclude & ")"  
		cInnerSQL = cInnerSQL & cIncludeSQL
	End If	
	cSQL = cSQL & cInnerSQL & ") ORDER BY NAME"
else
	cSQL = cSQL & ") ORDER BY NAME"
end if	
If Session("AHSTreeShowAllNodes").Item("AHSID=" & nParentID) <> "" Then
	oRS.MaxRecords = 999
elseif nParentID = "1" or lGetAll then	'	or NodeType = "RISK LOCATION" then
	oRS.MaxRecords = 999
else
	oRS.MaxRecords = Session("CurrentMaxRecords")
end if	
loadTree cSQL, nParentID, nCurLevel, lGetAll
End Function

Sub loadTree(cSQL, nParentID, nCurLevel, lGetAllNodes)
Dim aID
Dim cIDList, cCurName
Dim i, nMax, aSubparts
Dim cType, cFilter

' kmp hack
if nParentID = "31" then exit sub

i = 0
oRS.Open cSQL, cConnStr, adOpenStatic, adLockReadOnly, adCmdText		
cIDList = ""	
Set AHSTree = Session("AHSTreeProperties")
do	While Not oRS.EOF
	cIDList = cIDList & "|" & oRS("AHSID") & ":" & oRS("AHSType")
	cCurName = oRS("AHSName")
	If IsNull(cCurName) Then 
		cCurName = ""
	end if
	AHSTree.Add CStr(oRS("AHSID")), Replace(CStr(cCurName), """", """""") & "    [ " & CStr(oRS("AHSID")) & " ]"
	i = i + 1
	oRS.MoveNext
loop
oRS.close
nMax = i
i = 0
aID = Split(cIDList, "|")
For i = 1 To nMax
	aSubparts = Split(aID(i), ":")
	cType = "CLIENT"
	If aSubparts(1) = "FNS" Then
		cType = "FOLDER"
	ElseIf aSubparts(1) = "ACCOUNT" Then
		cTyper = "ACCOUNT"
	ElseIf aSubparts(1) = "RISK LOCATION" Then
		cType = "RISKLOCATION"
	End If

	If HasSpecificFilter("AHSID=" & aSubparts(0), "DESIGNER_AHSFILTER") Then
		cFilter = "FIL"
	else 
		cFilter = ""
	end if
	if clng(nParentID) <> 0 then
	%>
	NodeX = TreeView1.AddNode ("AHSID=<%=nParentID%>", 4 , "AHSID=<%=aSubparts(0)%>", "<%=cType%>",  "<%=AHSTree.Item(aSubparts(0))%>" ,"<%=cType & filterApplied%>", "<%=cType & cFilter%>SEL")
	<%
	else
	%>
	NodeX = TreeView1.AddNode ("AHSID=1", 4 , "AHSID=<%=aSubparts(0)%>", "<%=cType%>",  "<%=AHSTree.Item(aSubparts(0))%>" ,"<%=cType & filterApplied%>", "<%=cType & cFilter%>SEL")
	<%
	end if
	if Session("AHSTreeShowNodes").Item("AHSID=" & aSubparts(0)) <> "" then
		GetAHSChildren aSubparts(0), aSubparts(1), nCurLevel, false
	End If
Next
If not HasSpecificFilter("AHSID=" & nParentID, "DESIGNER_AHSFILTER") And _
	 nMax > 0 And i > oRS.MaxRecords - 1 and not lGetAllNodes and nParentID<>"1" and nParentID<>"0" Then
%>
	NodeX = TreeView1.AddNode ("AHSID=<%=nParentID%>", 4 , "GRP=<%=nParentID%>", "FILTER",  "More..." ,"FOLDERGRP", "FOLDERGRPSEL")
<%
End If
end sub

Function hasNodeAccess(cAHS_ID, aNodesID)
dim cSQL
dim cID
dim x
dim lIDOk
dim cIDToSearch

' check the given node ID 
lIDOk = false
for x=lbound(aNodesID) to ubound(aNodesID)
	if aNodesID(x) = cAHS_ID then
		lIDOk = true
		exit for
	end if
next
if not lIDOk then
	' check the parent and beyond
	cIDToSearch = cAHS_ID
	do while cID <> "1"
		cSQL = "Select parent_node_id " & _
					"From account_hierarchy_step " & _
					"Where accnt_hrcy_step_id = " & cIDToSearch
		oRS.Open cSQL, cConnStr, adOpenStatic, adLockReadOnly, adCmdText
		cID = CStr(oRS.Fields("parent_node_id").Value)
		oRS.Close
		lIDOk = false
		for x=lbound(aNodesID) to ubound(aNodesID)
			if aNodesID(x) = cID then
				lIDOk = true
				exit for
			end if
		next
		if not lIDOk then
			cIDToSearch = cID
		else
			exit do
		end if
	loop			
end if
hasNodeAccess = lIDOk
End function

Function GetAHSParent(nNodeID)
dim cSQL

cSQL = "SELECT AHS.PARENT_NODE_ID FROM ACCOUNT_HIERARCHY_STEP AHS WHERE AHS.ACCNT_HRCY_STEP_ID =" & CStr(nNodeID) 
cSQL =  cSQL & " ORDER BY NAME"
oRS.Open cSQL, cConnStr, adOpenStatic, adLockReadOnly, adCmdText
GetAHSParent = ""
If Not oRS.EOF Then
	If Not IsEmpty(oRS("PARENT_NODE_ID")) And Not IsNull(oRS("PARENT_NODE_ID")) Then
		GetAHSParent = oRS("PARENT_NODE_ID")
	End If
End If
oRS.Close 
End Function

Function GetAHSDesc(nNodeID)
dim cSQL

cSQL = "SELECT AHS.NAME FROM ACCOUNT_HIERARCHY_STEP AHS WHERE AHS.ACCNT_HRCY_STEP_ID =" & CStr(nNodeID) & " ORDER BY NAME"
oRS.Open cSQL, cConnStr, adOpenStatic, adLockReadOnly, adCmdText

GetAHSDesc = ""
If Not oRS.EOF Then
	If Not IsEmpty(oRS("NAME")) And Not IsNull(oRS("NAME")) Then
		GetAHSDesc = oRS("NAME")
	End If
End If
oRS.Close 
End Function
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>FNS Account Lookup Tree</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="vbscript">

Sub Window_Onload
<%
Dim cType, cNodeType
dim lHasRootAccess, aNodes, x

If bHighestLevelReached = True Then
%>
	MsgBox "The highest accessible level has been reached."
<%
End If
		
cType = "CLIENT"
cNodeType = Request.QueryString("NOdeType")

If cNodeType = "FNS" Then
	cType = "FOLDER"
ElseIf cNodeType = "ACCOUNT" Then
	cType = "ACCOUNT"
ElseIf cNodeType = "RISKLOCATION" Then
	cType = "RISKLOCATION"
End If

lHasRootAccess = false
aNodes = split( Session("ACCOUNT_SECURITY"), "," )
for x=lbound(aNodes) to ubound(aNodes)
	if aNodes(x) = "1" then
		lHasRootAccess = true
		exit for
	end if
next

If HasSpecificFilter("AHSID=" & Session("CurAHSTreeTopID"), "DESIGNER_AHSFILTER") Then
	filterApplied = "FIL"
end if
%>
NodeX = TreeView1.AddNode("", 1, "AHSID=<%=Session("CurAHSTreeTopID")%>", "TOPNODE", "<%=Session("CurAHSTreeTopDesc")%>", "<%=cType & filterApplied%>", "<%=cType & filterApplied%>SEL")
<%
if instr(1, Request.QueryString, "Build_Tree" ) <> 0 OR (Request.QueryString("UPONELEVEL") <> "" AND Session("CurAHSTreeTopID") = "1") then
	GetAHSChildren Session("CurAHSTreeTopID"), "FNS", 0, true
else
	GetAHSChildren Session("CurAHSTreeTopID"), "FNS", 0, false
end if		
%>
    NodeX = TreeView1.AddMenuItem("TOPNODE", "Up One Level", ErrStr)
    NodeX = TreeView1.AddMenuItem("TOPNODE", "Add to Favorites", ErrStr)
    NodeX = TreeView1.AddMenuItem("TOPNODE", "-", ErrStr)
<%if lHasRootAccess then%>
    NodeX = TreeView1.AddMenuItem("TOPNODE", "Add Child Node", ErrStr)
    NodeX = TreeView1.AddMenuItem("TOPNODE", "Node Search", ErrStr)
    NodeX = TreeView1.AddMenuItem("TOPNODE", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("TOPNODE", "Filter Properties", ErrStr)
    NodeX = TreeView1.AddMenuItem("TOPNODE", "-", ErrStr)
<%end if%>    
    NodeX = TreeView1.AddMenuItem("TOPNODE", "Refresh", ErrStr)
    
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "Show Nodes", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "Show as Top Node", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "Add to Favorites", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "Add Child Node", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "Node Search", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "Filter Properties", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("ACCOUNT", "Refresh", ErrStr)
    
    NodeX = TreeView1.AddMenuItem("CLIENT", "Show Nodes", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "Show as Top Node", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "Add to Favorites", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "Add Child Node", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "Node Search", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "Filter Properties", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("CLIENT", "Refresh", ErrStr)
    
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "Show Nodes", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "Show as Top Node", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "Add to Favorites", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "Add Child Node", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "Node Search", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "Filter Properties", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("RISKLOCATION", "Refresh", ErrStr)
            
    NodeX = TreeView1.AddMenuItem("FILTER", "Filter Properties", ErrStr)
    NodeX = TreeView1.AddMenuItem("FILTER", "Show All", ErrStr)

    NodeX = TreeView1.AddMenuItem("EXPAND", "Retrieve Nodes", ErrStr)
    
    TreeView1.ExpandNode("AHSID=<%=Session("CurAHSTreeTopID")%>")

<%
	If Request.QueryString("EXPANDNODE") <> "" Then
		part = Split(Request.QueryString("EXPANDNODE"),"=")
		key = "AHSID=" & part(1)
%>
	TreeView1.ExpandNode("<%=Key%>")
<%
	End If
%>
	call window_onresize()
<%
if isObject(oRS) then
	set oRS = nothing
end if

aKeys = Session("AHSTreeCurExpandedNodes").Keys   ' Get the keys.
for x=0 to Session("AHSTreeCurExpandedNodes").count - 1
%>	
	TreeView1.ExpandNode("<%=aKeys(x)%>")
<%
next	
%>	
End Sub

Sub window_onresize()
	TreeView1.style.posTop = 0
	TreeView1.style.posLeft = 0
	TreeView1.style.Width = document.body.clientWidth
	TreeView1.style.height = document.body.clientHeight - 12
End Sub


Sub TreeView1_NodeClicked( NodeType, NodeKey, NodeText , IsLoaded , Shift )
	If NodeType <> "FILTER" And  NodeType <> "EXPAND" Then
		Parent.frames("WORK").location.href = "NodeSummary.asp?" & NodeKey
	End If
End Sub

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )

	Select Case MenuItem

		Case "Node Search"

			If NodeKey = "AHSID=1" Then
				MsgBox "Node search is not available at this level.",0,"FNSNetDesigner"
			Else
				SearchObj.goto = false
				SearchObj.ahsid = false
				SearchObj.NodeText = ""
				lret = showModalDialog  ("TreeSearchModal.asp?" & NodeKey , SearchObj, "dialogWidth:700px;dialogHeight:500px")
				if SearchObj.goto = true then
					window.location = "AHTree.asp?MAXRS=<%= Session("USERTREECOUNT") %>&MAXLEVEL=<%= Session("USERTREELEVELS") %>&" & SearchObj.ahsid & "&NodeType=CLIENT&AHSDesc=UNKNOWN" 
					Parent.frames("WORK").location.href  = "NodeSummary.asp?" & SearchObj.ahsid
				End If
			End If
		Case "Filter Properties"
			part = Split(NodeKey,"=")
			key = CStr(part(1))
			lret = showModalDialog  ("TreeFilterModal.asp?KEY=" & key, null, "dialogWidth:700px;dialogHeight:500px")
			window.location = "AHTree.asp?Build_Tree&AHSID=<%=Request.QueryString("AHSID")%>&EXPANDNODE=" & "AHSID=" & key & "&NodeType=<%=Request.QueryString("NodeType")%>" 
		Case "Show All"
			window.location = "AHTree.asp?SHOWALL=" & NodeKey & "&NodeType=" & NodeType & "&Build_Tree"
		Case "Show as Top Node"
			window.location = "AHTree.asp?MAXRS=<%= Session("USERTREECOUNT") %>&MAXLEVEL=<%= Session("USERTREELEVELS") %>&" & NodeKey & "&NodeType=" & NodeType & "&AHSDesc=" & NodeText
		Case "Retrieve Nodes"
			window.location = "AHTree.asp?EXPANDNODE=" & NodeKey & "&NodeType=" & NodeType & "&Build_Tree"

		Case "Up One Level"
			window.location = "AHTree.asp?UPONELEVEL=" & NodeKey & "&NodeType=" & NodeType
		Case "Add to Favorites"
			Parent.frames("FAVORITES").location.href = "Favorites.asp?ADD=TRUE&" & NodeKey
		Case "Add Child Node"
			Parent.frames("WORK").location.href = "AHSMaintenance.asp?DETAILONLY=TRUE&AHSID=NEW&PARENT_" & NodeKey
		Case "Refresh"
			window.location = "AHTree.asp?REFRESH=TRUE&AHSID=<%= Request.QueryString("AHSID") %>" & "&Build_Tree"
		Case "Show Nodes"
			window.location = "AHTree.asp?SHOWNODES=" & NodeKey & "&NodeType=" & NodeType & "&Build_Tree"
		Case Else
			MsgBox "Not Yet Implemented"
	End Select
End Sub

Sub TreeView1_NodeDblClicked(NodeName, NodeKey, NodeText, IsLoaded, Shift)
	Select Case NodeName
		Case "EXPAND"
			window.location = "AHTree.asp?EXPANDNODE=" & NodeKey & "&NodeType=" & NodeType
		Case "FILTER"
			TreeView1_NodeMenuClicked "", NodeKey, NodeText, "Filter Properties"
	End Select
End Sub

<!--#include file="..\lib\Help.asp"-->

</script>
<script LANGUAGE="JavaScript">
<!--
function CRPSearchObj()
{
	this.goto = false;
	this.ahsid = "";
	this.copy = false;
	this.copyproperty = "";
	this.copyvalue = "";
	this.NodeText = "";
}
var SearchObj = new CRPSearchObj();
//-->
</script>
</head>
<body bgcolor="white" topmargin="0" leftmargin="0" RightMargin="0" bottommargin="0">
<table WIDTH="100%" Height="1" BGCOLOR="#006699" CELLPADDING="0" CELLSPACING="0">
<tr>
<td CLASS="LABEL" ALIGN="LEFT">
<font COLOR="WHITE">» Account Hierarchy</td>
<td ALIGN="RIGHT" CLASS="LABEL"><img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp('Hierarchy Tree.htm')" WIDTH="7" HEIGHT="8">&nbsp;</td>
</tr>
</table>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="100%" BORDER=0>
<PARAM NAME="ShowTips" VALUE="False"></OBJECT></BODY>
</html>