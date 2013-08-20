<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<html>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<title>Attribute Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub BtnFind_onclick
	If TxtName.Value <> "" OR TxtCaption.Value <> "" OR TxtDescription.value <> "" OR TxtHelpString.value <> "" Then
		if SearchType(0).checked = True Then
			Clause = Clause & "SEARCHTYPE=" & SearchType(0).Value
		End If
		If SearchType(1).checked = True Then
			Clause = Clause & "SEARCHTYPE=" & SearchType(1).Value
		End If
		If SearchType(2).checked = True Then
			Clause = Clause & "SEARCHTYPE=" & SearchType(2).Value
		End if
		Clause = Clause & "&TxtName=" & TxtName.Value 
		Clause = Clause & "&TxtCaption=" & TxtCaption.Value
		Clause = Clause & "&TxtDescription=" & TxtDescription.value
		Clause = Clause & "&TxtHelpString=" & TxtHelpString.value
		location.href = "AttribSearch.asp?PAGE=<%= Request.QueryString("PAGE") %>&" & Clause
	Else
		MsgBox "Please Enter Search Criteria", 0 , "FNSDesigner"
	End if
End Sub

Sub BtnClear_OnClick
	TxtName.Value = ""
	TxtCaption.Value = ""
	TxtDescription.Value = ""
	TxtHelpString.Value = ""
End Sub

Sub window_onload
TxtName.focus
SearchType(0).checked = True
<%
	If Request.QueryString("SEARCHTYPE") <> "" Then
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				NAME = Request.QueryString("TxtName") & "%"
				CAPTION = Request.QueryString("TxtCaption") & "%"
				DESCRIPTION = Request.QueryString("TxtDescription") & "%"
				HELP = Request.QueryString("TxtHelpString") & "%"
			Case "C"
				NAME = "%" & Request.QueryString("TxtName") & "%"
				CAPTION = "%" & Request.QueryString("TxtCaption") & "%"
				DESCRIPTION = "%" & Request.QueryString("TxtDescription") & "%"
				HELP = "%" & Request.QueryString("TxtHelpString") & "%"
			Case "E"
				NAME = Request.QueryString("TxtName")
				CAPTION = Request.QueryString("TxtCaption")
				DESCRIPTION = Request.QueryString("TxtDescription")
				HELP = Request.QueryString("TxtHelpString")
		End Select
	
		If Request.QueryString("TxtName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("TxtCaption") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(CAPTION) LIKE '" & UCASE(CAPTION) & "'"
		End If
		If Request.QueryString("TxtDescription") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		If Request.QueryString("TxtHelpString") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(HELPSTRING) LIKE '" & UCASE(HELP) & "'"
		End If
		Set Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = CONNECT_STRING
		Conn.Open ConnectionString
		
		SQLST = "SELECT NAME, ATTRIBUTE_ID FROM ATTRIBUTE WHERE " & WHERECLS & " ORDER BY NAME" 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF
%>
NodeX = TreeView1.AddNode ("", 4 , "ID=<%= RS("ATTRIBUTE_ID") %>", "ATTRIB", "<%= RS("NAME") %>","FOLDER", "FOLDERSEL" )

<%
	RS.MoveNext
	Loop
%>
Select Case "<%= Request.QueryString("SEARCHTYPE") %>"
	Case "C"
		SearchType(1).checked = True
	Case "B"
		SearchType(0).checked = True
	Case "E"
		SearchType(2).checked = True
	Case Else
		SearchType(0).checked = True
End Select
<%
	End If 
%>
End Sub

Sub TreeView1_NodeDblClicked( NodeType, NodeKey, NodeText , IsLoaded , Shift )
<% If Request.QueryString("PAGE") = "LAYOUT" Then %>
	window.opener.SearchResults NodeKey, NodeText
<% Else %>
	window.opener.document.all.ATTRIBUTENAME.Value = Trim(NodeText)
<% End If %>
End Sub

Sub document_onkeydown
		select case window.event.keyCode
			case 13
				call btnFind_onclick
			case else:

		end select
End Sub

Sub BtnOK_onclick
If G_CurNodeText <> "" Then
	window.opener.document.all.NEWATTR.value = G_CurNodeText
End If
window.close
End Sub

Dim G_CurNodeKey
Dim G_CurNodeText
Sub TreeView1_NodeClicked( NodeType, NodeKey, NodeText , IsLoaded , Shift )
G_CurNodeKey = NodeKey
G_CurNodeText	= NodeText
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR='<%=BODYBGCOLOR%>' >
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Attribute Search</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT NAME="TxtName" VALUE="<%= Request.QueryString("TxtName") %>" SIZE=25 CLASS=LABEL></TD>
<TD CLASS=LABEL>Caption:<BR><INPUT TYPE=TEXT NAME="TxtCaption" VALUE="<%= Request.QueryString("TxtCaption") %>" SIZE=25 CLASS=LABEL></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT NAME="TxtDescription" VALUE="<%= Request.QueryString("TxtDescription") %>" SIZE=25 CLASS=LABEL></TD>
<TD COLSPAN=4 CLASS=LABEL>Help String:<BR><INPUT TYPE=TEXT NAME="TxtHelpString" SIZE=25 CLASS=LABEL VALUE="<%= Request.QueryString("TxtHelpString") %>"></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL><INPUT TYPE=RADIO STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS=LABEL>Begins With</TD>
<TD CLASS=LABEL><INPUT TYPE=RADIO STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS=LABEL>Contains</TD>
<TD CLASS=LABEL><INPUT TYPE=RADIO STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS=LABEL>Exact</TD>
</TR>
</TABLE>

<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnFind">Find</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnClear">Clear</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnOK">Ok</BUTTON></TD>
</TR>
</TABLE>
<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0">
	<PARAM NAME="LPKPath" VALUE="LPKfilename.LPK">
</OBJECT>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="65%">
</OBJECT>
</BODY>
</HTML>

