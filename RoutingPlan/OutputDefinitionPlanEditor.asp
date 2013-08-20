<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<% Response.Expires=0 %>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE>Routing Plan Summary</TITLE>
<%= Request.QueryString("RPID") %>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	NodeX = TreeView1.AddNode ("",1 , "STEP=1", "ROOT", "XYZ Routing Plan", "FOLDER", "FOLDERSEL")
	NodeX = TreeView1.AddNode ("STEP=1", 4 , "STEP=32", "TRANSMISSION", "(1) Transmission Step: Fax, Dest. \\Server1\PC100" ,"FIELD", "FIELDSEL" )
	NodeX = TreeView1.AddNode ("STEP=1", 4 , "STEP=33", "TRANSMISSION", "(2) Transmission Step: Print, Dest. \\Server1\PC100" ,"FIELD", "FIELDSEL" )
	NodeX = TreeView1.AddNode ("STEP=32", 4 , "STEP=35", "PAGE", "(1) Page: XYZ CAU Cover Letter" ,"FRAME", "FRAMESEL" )
	NodeX = TreeView1.AddNode ("STEP=32", 4 , "STEP=36", "PAGE", "(2) Page: Form XYZ-1"  ,"FRAME", "FRAMESEL")
	NodeX = TreeView1.AddNode ("STEP=32", 4 , "STEP=37", "PAGE",  "(3) Page: Overflow XYZ-1 page"  ,"FRAME", "FRAMESEL")
	
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "Properties", ErrStr)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "New Assignment", ErrStr)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "-", ErrStr)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "Move Up", ErrStr)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "Move Down", ErrStr)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "-", ErrStr2)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "Remove", ErrStr)
	
	NodeX = TreeView1.AddMenuItem("ROOT", "New", ErrStr)
	
	NodeX = TreeView1.AddMenuItem("PAGE", "Override Fields", ErrStr)
	
	
End Sub

Sub TreeView1_DblClick
	If Not TreeView1.SelectedItem.Index = 1 Then
		showModalDialog  "RPPropertiesModal.asp?CFID=141"  , " PropertiesModal", "dialogWidth:300px;dialogHeight:300px"
	End If
End Sub

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
	Select Case MenuItem
		Case "Properties"
			showModalDialog  "RPPropertiesModal.asp?CFID=141"  , "PropertiesModal", "dialogWidth:300px;dialogHeight:300px"
		Case "Root"
			
		Case Else
	End Select
End Sub



Sub BtnVisual_OnClick
	Window.open  "OutputDefinitionEditor-f.asp?OPID=466", Null, "height=500,width=750,status=no,toolbar=no,menubar=no,location=no"
End Sub
-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=0 topmargin=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Routing Plan Summary</FONT>
</TD>
<TD STYLE="WIDTH:10" OnClick="Window.History.Back (1)" CLASS=UPMENU>U</TD>
<TD STYLE="WIDTH:10" OnClick="Window.History.Back(1)" CLASS=UPMENU>S</TD>
</TR>
</TABLE>
</FIELDSET>
<FIELDSET STYLE="BACKGROUND:SILVER">
<TABLE WIDTH="100%" BORDER=0>
<TR>
<TD CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPDescription" SIZE=30></TD>
<TD CLASS=LABEL>Dest. Type:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPDTYPE" SIZE=30></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnRPSave">Save</BUTTON></TD>
</TR>
<TR>
<TD CLASS=LABEL>LOB:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPLOB" SIZE=30></TD>
<TD CLASS=LABEL>Input System Name: <BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPSTATE" SIZE=30></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnRPCancel">Cancel</BUTTON></TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2 ALIGN=LEFT>Enable Rule:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPSTATE" SIZE=75></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnRPCompose">Compose</BUTTON></TD>
</TR>
<TR>
<TD CLASS=LABEL ALIGN=LEFT>State: <BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPSTATE" MAXLENGTH=2 SIZE=2></TD>
<TD></TD>
<TD CLASS=LABEL ALIGN=LEFT><INPUT CLASS=LABEL TYPE=CHECKBOX NAME="CHECKENABLED"> Enabled</TD>
</TR>
</TABLE>
</FIELDSET>
<BR>
<TABLE WIDTH="100%">
<TR>
<TD><BUTTON CLASS=STDBUTTON STYLE="WIDTH:25;CURSOR:N-RESIZE;">U</BUTTON><BUTTON CLASS=STDBUTTON STYLE="WIDTH:25;CURSOR:S-RESIZE;">D</BUTTON></TD>
<TD ALIGN=RIGHT><BUTTON STYLE="WIDTH:80" NAME="BtnVisual" CLASS=STDBUTTON>Visual Editor</BUTTON></TD>
</TR>
</TABLE>
<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0"1>
	<PARAM NAME="LPKPath" VALUE="LPKfilename.LPK">
</OBJECT>
<OBJECT ID="TreeView1" <%GetTreeCLSID()%>  Width="100%" Height="55%">
</OBJECT>
</BODY>
</HTML>
