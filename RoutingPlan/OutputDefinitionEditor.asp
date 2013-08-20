<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\TreeCLSID.inc"-->
<% Response.Expires=0 %>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE>Routing Plan Summary</TITLE>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload

	NodeX = TreeView1.AddNode ("", 4 , "STEP=33", "TRANSMISSION", "Output Definition" ,"FIELD", "FIELDSEL" )
	NodeX = TreeView1.AddNode ("STEP=33", 4 , "STEP=35", "PAGE", "(1) Page: XYZ CAU Cover Letter" ,"FRAME", "FRAMESEL" )
	NodeX = TreeView1.AddNode ("STEP=35", 4 , "STEP=55", "FIELD", "(1) Field1" ,"FIELD", "FIELDSEL" )
	NodeX = TreeView1.AddNode ("STEP=33", 4 , "STEP=36", "PAGE", "(2) Page: Form XYZ-1"  ,"FRAME", "FRAMESEL")
	
	NodeX = TreeView1.AddNode ("STEP=33", 4 , "STEP=37", "PAGE",  "(3) Page: Overflow XYZ-1 page"  ,"FRAME", "FRAMESEL")
	
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "Properties", ErrStr)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "New Assignment", ErrStr)
	NodeX = TreeView1.AddMenuItem("TRANSMISSION", "Remove", ErrStr)
	
		
	NodeX = TreeView1.AddMenuItem("PAGE", "Properties", ErrStr)
	NodeX = TreeView1.AddMenuItem("FIELD", "Properties", ErrStr)
	
End Sub

Sub TreeView1_DblClick
	If Not TreeView1.SelectedItem.Index = 1 Then
		showModalDialog  "RPPropertiesModal.asp?CFID=1"  , " PropertiesModal", "dialogWidth:300px;dialogHeight:300px"
	End If
End Sub

Sub TreeView1_NodeMenuClicked( NodeType,  NodeKey ,  NodeText ,  MenuItem )
	Select Case NodeType
		Case "FIELD"
			showModalDialog  "ODFieldproperties.asp?CFID=1"  , "PropertiesModal", "dialogWidth:300px;dialogHeight:300px"
		Case "PAGE"
			showModalDialog  "ODpageproperties.asp?CFID=1"  , "PropertiesModal", "dialogWidth:300px;dialogHeight:300px"
		Case Else
	End Select
End Sub

Sub BtnVisual_OnClick
	Window.open  "OutputDefinitionEditor-f.asp?OPID=25&ODID=25", Null, "height=500,width=750,status=no,toolbar=no,menubar=no,location=no"
End Sub
-->
</SCRIPT>
</HEAD>
<BODY  leftmargin=0 topmargin=0>
<FIELDSET STYLE="BACKGROUND:SILVER;WIDTH='100%'">
<TABLE WIDTH="100%" >
<TR BGCOLOR=SILVER>
<TD CLASS=LABEL>
<FONT SIZE=2>Output Definition Summary</FONT>
</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back (1)" CLASS=LABEL>
U</TD>
<TD STYLE="BORDER-STYLE:GROOVE;BORDER-WIDTH:1;WIDTH:10;CURSOR:HAND" OnCLick="Window.History.Back(1)" CLASS=LABEL>
S</TD>
</TR>
</TABLE>
</FIELDSET>
<FIELDSET STYLE="BACKGROUND:SILVER">
<TABLE WIDTH="100%" BORDER=0>
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPDescription" SIZE=30></TD>
<TD CLASS=LABEL>Description:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME="RPDTYPE" SIZE=30></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnRPSave">Save</BUTTON></TD>
</TR>
<TR>
<TD CLASS=LABEL>Duplex Print:<INPUT TYPE=CHECKBOX CLASS=LABEL NAME="RPLOB"></TD>
<TD CLASS=LABEL></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnRPCancel">Cancel</BUTTON></TD>
</TR>
</TABLE>
</FIELDSET>
<BR>
<TABLE WIDTH="100%">
<TR>
<TD><BUTTON CLASS=STDBUTTON STYLE="WIDTH:25;CURSOR:N-RESIZE;" id=button1 name=button1>U</BUTTON><BUTTON CLASS=STDBUTTON STYLE="WIDTH:25;CURSOR:S-RESIZE;" id=button2 name=button2>D</BUTTON></TD>
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

