<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<title>Attribute Maintenance</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<% If Request.Form.Count = 0 Then %>
<% 
	WHERECLS = Request.QueryString("ID")
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = "DRIVER={Microsoft ODBC for Oracle};SERVER=190.15.5.4;ConnectString=FNS;UID=FNSOWNER;PWD=CTOWN"
	Conn.Open ConnectionString
	SQLST = "SELECT * FROM ATTRIBUTE WHERE ATTRIBUTE_ID=" & WHERECLS 
	Set RS = Conn.Execute(SQLST)

%>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub BtnSave_onclick
	AttribForm.submit
End Sub

-->
</SCRIPT>
</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Attribute Maintenance</FONT>
</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back (1)" CLASS=LABEL>
U</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
S</TD></TR>
</TABLE>
</FIELDSET>

<FORM NAME="AttribForm" Action="AttributeMaint2.asp?TX=True" Method="POST">
<TABLE>
<TR>
<TD><INPUT TYPE=HIDDEN NAME="ATTRIBUTE_ID" VALUE="<%= RS("ATTRIBUTE_ID") %>"></TD>
</TR>
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=35 NAME="NAME" VALUE="<%= RS("NAME") %>"></TD>
<TD CLASS=LABEL>Entry Mask:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=35 NAME="ENTRYMASK" VALUE="<%= RS("ENTRYMASK") %>"></TD>
</TR>
<TR>
<TD CLASS=LABEL>Input Type:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=35 NAME="INPUTTYPE" VALUE="<%= RS("INPUTTYPE") %>"></TD>
<TD CLASS=LABEL>Caption:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=35 NAME="CAPTION" VALUE="<%= RS("CAPTION") %>"></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=35 NAME="DESCRIPTION" VALUE="<%= RS("DESCRIPTION") %>"></TD>
<TD CLASS=LABEL>Help String:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=35 NAME="HELPSTRING" VALUE="<%= RS("HELPSTRING") %>"></TD>
</TR>
</TABLE>
<TABLE>
<TR>
<TD CLASS=LABEL>Text Length:<BR><INPUT TYPE=TEXT CLASS=LABEL SIZE=5 NAME="TEXTLENGTH" VALUE="<%= RS("TEXTLENGTH") %>"></TD>
<TD CLASS=LABEL VALIGN=BOTTOM><INPUT TYPE=CHECKBOX CLASS=LABEL NAME="SPELLCHECK_FLG">Spell Check Flag</TD>
</TR>
</TABLE>
</FORM>
<BR>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnSave">Save</BUTTON></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON NAME="BtnCancel">Cancel</BUTTON></TD>
</TR>
</TABLE>
</BODY>
<% Else 

If Request.Form("SPELLCHECK_FLG") <> "" Then
	SPELLCHECK_FLG = "Y"
Else
	SPELLCHECK_FLG = "N"
End If


%>



<% End If %>
</BODY>
</HTML>
