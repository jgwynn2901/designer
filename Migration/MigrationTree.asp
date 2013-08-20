<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->

<%	Response.Expires = 0 %>

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>FNS Account Lookup Tree</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

<!--#include file="..\lib\Help.asp"-->

Sub Window_Onload
	NodeX = TreeView1.AddNode ("",1  , "STEP=1060",  "MIGRATION", "Migration", "FOLDER", "FOLDERSEL" )
	NodeX = TreeView1.AddNode ("STEP=1060", 4 , "STEP=100", "MIGRATION_JOB","Migration Job", "PAGE", "PAGESEL" )
	NodeX = TreeView1.AddNode ("STEP=1060", 4 , "STEP=1020", "MIGRATION_VIEW","Migration Status", "PAGE", "PAGESEL" )
	Treeview1.ExpandNode("STEP=1060")
	window_onresize()
End Sub
	
Sub window_onresize()
	TreeView1.style.posTop = 0
	TreeView1.style.posLeft = 0
	TreeView1.style.pixelWidth = document.body.clientWidth
	TreeView1.style.height = document.body.clientHeight - 12
End Sub

Sub TreeView1_NodeClicked( NodeType, NodeKey, NodeText , IsLoaded , Shift )
	Select Case NodeType
		Case "MIGRATION_JOB"
			Parent.frames("WORK").location.href = "MigrationJob-f.asp"
		Case "MIGRATION_VIEW"
			Parent.frames("WORK").location.href = "MigrationStatus.asp"	
	End Select
End Sub

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
	Select Case MenuItem
		Case "Node Search"
			showModalDialog  "SearchModal.asp?CFID=1"  , "PropertiesModal", "dialogWidth:700px;dialogHeight:500px"
		Case Else
	End Select
End Sub
-->
</script>
</head>
<body bgcolor="white" topmargin="0" leftmargin="0" RightMargin="0" bottommargin="0">
<table WIDTH="100%" Height="1" BGCOLOR="#006699" CELLPADDING="0" CELLSPACING="0">
<tr>
<td CLASS="LABEL" ALIGN="LEFT">
<font COLOR="WHITE">» Migration</td>
<td ALIGN="RIGHT" CLASS="LABEL"><img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8">&nbsp;</td>
</tr>
</table>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="100%">
</object>
</body>
</html>