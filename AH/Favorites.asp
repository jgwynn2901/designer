<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\AHSTree.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<% Response.Expires=0 %>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

If Request.QueryString("PATH") = "TRUE" Then
	RemoveNode(Request.QueryString("AHSID"))
	If Session("EXPANDLIST") <> "" AND InStr(1, Session("EXPANDLIST"), Request.QueryString("AHSID")) = 0 Then
		Session("EXPANDLIST") = Trim(Session("EXPANDLIST")) & ","
	End If
	If InStr(1, Session("EXPANDLIST"), Request.QueryString("AHSID")) = 0 Then
		Session("EXPANDLIST") = Session("EXPANDLIST") & Trim(Request.QueryString("AHSID"))
	End If
	SetFilterByName "KEY=null" , "EXP_DESIGNER_FAVORITES", "EXP_FAVORITES_AHSID", Session("EXPANDLIST") 
End If

If Request.QueryString("REMOVENODE") = "TRUE" Then
	RemoveNode(Request.QueryString("AHSID"))
	RemoveExpandListNode(Request.QueryString("AHSID"))
End If
	
If Request.QueryString("ADD") <> "" Then
	If Session("AHLIST") <> "" Then
		Session("AHLIST") = Trim(Session("AHLIST")) & ","
	End If
	Session("AHLIST") = Session("AHLIST") & Trim(Request.QueryString("AHSID"))
	SetFilterByName "KEY=1" , "DESIGNER_FAVORITES", "EXP_FAVORITES_AHSID", Session("AHLIST") 
End If

If Session("AHLIST") <> "" Then
	SQL = ""
	SQL = SQL & "SELECT ACCNT_HRCY_STEP_ID, NAME FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID IN (" & Session("AHLIST") & ")"
	Set RS = Conn.Execute(SQL)
End If


sub removeNode(cID)
dim aFav, x
dim lFirst
dim cResult

cResult = ""
lFirst = true
aFav = split(Session("AHLIST"), ",")
for x=0 to ubound(aFav)
	if aFav(x) <> cID then
		if lFirst then
			cResult = aFav(x)
			lFirst = false		
		else
			cResult = cResult & "," & aFav(x)
		end if
	end if
next	
Session("AHLIST") = cResult
End sub

Function RemoveExpandListNode(ID)
If Len(Session("AHLIST")) = 0 Then
	MyLen = 1
Else
	MyLen = Len(Session("AHLIST"))
End If

	Session("EXPANDLIST") = Replace(Session("EXPANDLIST"), ID, "")
	Session("EXPANDLIST") = Replace(Session("EXPANDLIST"), ",,", ",")
	If Len(Session("EXPANDLIST")) <> 0 Then
	If Mid(Session("EXPANDLIST"), 1,1) = "," Then
		Session("EXPANDLIST") = Mid(Session("EXPANDLIST"), 2,Len(Session("EXPANDLIST")))
	End If
	If Mid(Session("AHLIST"), MyLen, 1) = "," Then
		Session("EXPANDLIST") = Mid(Session("EXPANDLIST"), 1,MyLen-1)
	End If
	End If
End Function

Function ExpandNode(ID, Ubound)
	Set RSPath = Server.CreateObject("ADODB.RecordSet")
	RSPath.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	QSQL = "{call Designer_2.GetUpTreePathNodes(" & ID & ", 1, {resultset 200, outAHSID, outAcctName, outLevel})}"
	RSPath.Open QSQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
    ParentNode = RSPath("outAHSID")
    If Ubound = 0 Then
%>    
	NodeX = TreeView1.AddNode (&quot;&quot;,1 , &quot;PATH=TRUE&amp;AHSID=<%= RSPath("outAHSID") %>&quot;, &quot;SHOWPATH&quot;, &quot;<%= RSPath("outAcctName") %> [ <%= RSPath("outAHSID")%> ]  &quot;, &quot;FOLDER&quot;, &quot;FOLDERSEL&quot;)
<%
	End If
	RSPath.MoveNext
	Do While Not RSPath.EOF 
	%>
		NodeX = TreeView1.AddNode (&quot;PATH=TRUE&amp;AHSID=<%= ParentNode %>&quot;,4 , &quot;PATH=TRUE&amp;AHSID=<%= RSPath("outAHSID") %>&quot;, &quot;TOPNODE&quot;, &quot;<%= RSPath("outAcctName") %> [ <%= RSPath("outAHSID")%> ]  &quot;, &quot;FOLDER&quot;, &quot;FOLDERSEL&quot;)	
	<%
		parentNode = RSPath("outAHSID")
		RSPath.MoveNext
	Loop
	RSPath.Close


End Function

%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub window_onload
<%
If Session("AHLIST") <> "" Then 
	Do While Not RS.EOF %>
	NodeX = TreeView1.AddNode ("",1 , "AHSID=<%=RS("ACCNT_HRCY_STEP_ID")%>", "TOPNODE", "<%= RS("NAME") %> [ <%= RS("ACCNT_HRCY_STEP_ID")%> ]  ", "FOLDER", "FOLDERSEL")	
	
<% RS.MoveNext
	 Loop
	 RS.CLose
 End If 
%>
	
<% 
If Session("EXPANDLIST") <> "" Then 
		ListArray = Split(Session("EXPANDLIST"), ",")
		For i = 0 to Ubound(ListArray) step 1
			ExpandNode ListArray(i), i
		Next
 End If
%>

    'NodeX = TreeView1.AddMenuItem("TOPNODE", "Show Path", ErrStr)
    'NodeX = TreeView1.AddMenuItem("TOPNODE", "-", ErrStr)
    NodeX = TreeView1.AddMenuItem("TOPNODE", "Remove from Favorites", ErrStr)
    
    TreeView1.style.pixelWidth = document.body.clientWidth
    TreeView1.style.height  = document.body.clientHeight - 17
    TreeView1.ExpandNode("PATH=TRUE&AHSID=<%= parentNode %>")
    
    
End Sub

Sub window_onresize()
	TreeView1.style.posTop = 0
	TreeView1.style.posLeft = 0
	TreeView1.style.pixelWidth = document.body.clientWidth
	If document.body.clientHeight > 17 Then
		TreeView1.style.height = document.body.clientHeight -17
	End If
End Sub

Sub TreeView1_NodeClicked( NodeType, NodeKey, NodeText , IsLoaded , Shift )
	Parent.frames("WORK").location.href = "NodeSummary.asp?" & NodeKey
End Sub

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
	Select Case MenuItem
		Case "Show Path"
			'self.location.href = "Favorites.asp?PATH=TRUE&" & NodeKey
		Case "Remove from Favorites"
			self.location.href = "Favorites.asp?REMOVENODE=TRUE&" & NodeKey	
    	Case Else
			MsgBox "Not Yet Implemented"
	End Select
End Sub

<!--#include file="..\lib\Help.asp"-->

</script>
</head>
<body bgcolor="white" topmargin="0" leftmargin="0" RightMargin="0" bottommargin="0">
<table WIDTH="100%" Height="1" BGCOLOR="#006699" CELLPADDING="0" CELLSPACING="0">
<tr>
<td CLASS="LABEL" ALIGN="LEFT">
<font COLOR="WHITE">» Favorites</td>
<td ALIGN="RIGHT" CLASS="LABEL">
<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp('Welcome.htm')" WIDTH="7" HEIGHT="8">&nbsp;</td>
</tr>
</table>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="100%" BORDER=0>
<PARAM NAME="ShowTips" VALUE="False">
</object>
</body>
</html>

