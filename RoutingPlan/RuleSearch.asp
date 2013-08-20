<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<html>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<title>Rule Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub BtnFind_onclick
	If TxtName.Value <> "" OR TxtCaption.Value <> "" OR TxtDescription.value <> "" Then
		
		Clause = Clause & "TxtName=%" & TxtName.Value  & "%"
		Clause = Clause & "&TxtCaption=%" & TxtCaption.Value & "%"
		Clause = Clause & "&TxtDescription=" & TxtDescription.value
		location.href = "RuleSearch.asp?SEARCHTYPE=SEARCH" & Clause
	Else
		MsgBox "Please Enter Search Criteria", 0 , "FNSDesigner"
	End if
End Sub

Sub BtnClear_OnClick
	TxtName.Value = ""
	TxtCaption.Value = ""
	TxtDescription.Value = ""
End Sub

Sub window_onload
TxtName.focus
<%
	If Request.QueryString("SEARCHTYPE") <> "" Then
	
		If Request.QueryString("TxtName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(RULE_TEXT) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("TxtCaption") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(TYPE) LIKE '" & UCASE(CAPTION) & "'"
		End If
		If Request.QueryString("TxtDescription") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(RULE_ID) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = "DRIVER={Microsoft ODBC for Oracle};SERVER=190.15.5.4;ConnectString=FNS;UID=FNSOWNER;PWD=CTOWN"
		Conn.Open ConnectionString
		SQLST = "SELECT RULE_TEXT, RULE_ID FROM RULES WHERE " & WHERECLS & " ORDER BY RULE_TEXT" 
		Set RS = Conn.Execute(SQLST)
		Do While Not RS.EOF
%>
NodeX = TreeView1.AddNode ("", 4 , "ID=<%= RS("ATTRIBUTE_ID") %>", "ATTRIB", "<%= RS("NAME") %>","FOLDER", "FOLDERSEL" )

<%
	RS.MoveNext
	Loop
	End If 
%>
End Sub

Sub TreeView1_NodeDblClicked( NodeType, NodeKey, NodeText , IsLoaded , Shift )
	window.opener.document.all.EnblRule.Value = Trim(NodeText)
End Sub

Sub document_onkeydown
		select case window.event.keyCode
			case 13
				call btnFind_onclick
			case else:

		end select
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0  rightmargin=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Rule Search</FONT>
</TD>
</TABLE>
</FIELDSET>
<TABLE>
<TR>
<TD CLASS=LABEL>Rule Text:<BR><INPUT TYPE=TEXT NAME="TxtName" VALUE="<%= Request.QueryString("TxtName") %>" SIZE=25 CLASS=LABEL></TD>
<TD CLASS=LABEL>Rule Type:<BR><INPUT TYPE=TEXT NAME="TxtCaption" VALUE="<%= Request.QueryString("TxtCaption") %>" SIZE=25 CLASS=LABEL></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Rule ID:<BR><INPUT TYPE=TEXT NAME="TxtDescription" VALUE="<%= Request.QueryString("TxtDescription") %>" SIZE=25 CLASS=LABEL></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnFind">Find</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnClear">Clear</BUTTON></TD>
</TR>
</TABLE>
<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0">
	<PARAM NAME="LPKPath" VALUE="LPKfilename.LPK">
</OBJECT>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="65%">
</OBJECT>
</BODY>
</HTML>

