<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<%

Response.Expires = 0
If HasViewPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then  
	Session("NAME") = ""
	Response.Redirect "Override_Layout_Bottom.asp"
End If
If HasModifyPrivilege("FNSD_OUTPUT_DEFINITION",SECURITYPRIV) <> True Then MODE = "RO"

If Len(Request.QueryString("OPID")) < 1 OR IsNumeric(Request.QueryString("OPID")) = False Then
	Session("ErrorMessage") = "On page " &  Request.ServerVariables("SCRIPT_NAME") & " QueryString OPID was Null or Not Numeric"
	Response.Redirect "..\directerror.asp"
End If
If Request.QueryString("ACTION") <> "SAVE" Then
	If Request.QueryString("OPID") <> "" Then 
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "SELECT * FROM OUTPUT_PAGE WHERE OUTPUT_PAGE_ID = " & Request.QueryString("OPID")
	Set RS = Conn.Execute(SQLST)
		
If RS.EOF or isnull(RS) Then
	Session("ErrorMessage") = "Statement = " & SQLST & " ----- returned no records" & vbCrlf
	Response.redirect	 "..\directerror.asp"
End If
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

<!--#include file="..\lib\Help.asp"-->

Sub BtnSave_onclick
    if document.all.Orientation.checked  then
       document.all.inpORIENTATION.value = "L"
    else
       document.all.inpORIENTATION.value = "P"
    end if
    
	If document.all.NAME.value = "" OR document.all.PAGE_NUMBER.value = "" Then
		Msgbox "Name or Page Number cannot be null", 0, "FNSDesigner"
	Else
		If Isnumeric(document.all.PAGE_NUMBER.value) Then
			FrmOutputpage.Submit()
		Else
			MsgBox "Page Number must be numeric", 0, "FNSNetDesigner"
		End If
	End If
End Sub

Sub window_onload
	document.all.NAME.focus()
	'If IsObject(Top.Frames("TOP").Document.all.PAGENAMEID) Then
	'	Top.Frames("TOP").Document.all.PAGENAMEID.InnerHtml = "<%= RS("NAME") %>"
	'End If
End Sub

Sub SetDirty
	document.body.SetAttribute "CanDocUnloadNowInf" , "YES"
End Sub
-->
</SCRIPT>

<script LANGUAGE="JavaScript">
function CanDocUnloadNow()
{
	if (false == confirm("Data has changed. Leave page without saving?"))
		return false;
	else
		return true;
}
</script>
</HEAD>
<BODY BGCOLOR='<%=BODYBGCOLOR%>'  topmargin=5 rightmargin=0 leftmargin=0 CanDocUnloadNowInf=NO>
<% If Not RS.EOF Then %>
<FORM NAME=FrmOutputpage ACTION="OP_GENERAL.ASP?ACTION=SAVE&OPID=<%= Request.QueryString("OPID") %>" METHOD=POST>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10>&nbsp&#187 Output Page
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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
<TR><TD CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> MAXLENGTH=80 SIZE=55 OnChange='SetDirty' OnKeyPress="SetDirty" NAME=NAME VALUE="<%= RS("NAME") %>"></TD>
<TD CLASS=LABEL>Page Number:<BR><INPUT TYPE=TEXT OnChange="SetDirty" <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>MAXLENGTH=10 OnKeyPress="SetDirty" NAME=PAGE_NUMBER CLASS=LABEL SIZE=5 VALUE="<%= RS("PAGE_NUMBER") %>"></TD></TR>
<TR><TD CLASS=LABEL>Output Tray:<BR><INPUT TYPE=TEXT CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %>MAXLENGTH=30 OnChange="SetDirty" OnKeyPress="SetDirty" NAME=OUTPUT_TRAY SIZE=55 VALUE="<%= RS("OUTPUT_TRAY") %>"></TD></TR>
<TR><TD CLASS=LABEL>Background BMP:<BR><INPUT TYPE=TEXT CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' " ) %> MAXLENGTH=255 OnChange="SetDirty" OnKeyPress="SetDirty" SIZE=55 name=BACKGROUND_BMP VALUE="<%= RS("BACKGROUND_BMP") %>"></TD></TR>
<%if RS("ORIENTATION") = "L" then %>
	   <td CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' " ) %>><input ScrnBtn="False" ScrnInput="TRUE" TYPE="CHECKBOX"  VALIGN="RIGHT" Name="Orientation" ID="Checkbox1"  checked  value="L" >Landscape</td>
	<%else%>
	    <td CLASS=LABEL <% If MODE="RO" Then Response.write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' " ) %>><input ScrnBtn="False" ScrnInput="TRUE" TYPE="CHECKBOX"  VALIGN="RIGHT" Name="Orientation" ID="Checkbox2"   value="P" >Landscape</td>
<%end if%>
</tr>
</TABLE>
<input type=hidden name=inpOrientation >
<BR>
<TABLE>
<TR><TD CLASS=LABEL><BUTTON CLASS=STDBUTTON 
<% If MODE="RO" Then Response.write(" DISABLED " ) %>
NAME=BtnSave ACCESSKEY="S"><U>S</U>ave</BUTTON></TD></TR>
</TABLE>

<% If Request.QueryString("STATUS") ="TRUE" Then %>
<LABEL CLASS=LABEL><FONT COLOR=MAROON>&nbsp;Page Saved</FONT></LABEL>
<% End If %>

<% End If %>
<% End If %>
</BODY>
</HTML>
<% Else
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "UPDATE OUTPUT_PAGE SET NAME='" & Request.Form("NAME") & "', "
	SQLST = SQLST & "PAGE_NUMBER=" & Request.Form("PAGE_NUMBER") & ", "
	SQLST = SQLST & "OUTPUT_TRAY='" & Request.Form("OUTPUT_TRAY") & "', "
	SQLST = SQLST & "BACKGROUND_BMP='" & Request.Form("BACKGROUND_BMP") & "',"
	SQLST = SQLST & "ORIENTATION='" & Request.Form("INPORIENTATION") & "'"
	SQLST = SQLST & " WHERE OUTPUT_PAGE_ID=" & Request.QueryString("OPID")
	Set RS = Conn.Execute(SQLST)
	Response.Redirect "OP_General.asp?STATUS=TRUE&OPID=" & Request.QueryString("OPID")
End If
%>