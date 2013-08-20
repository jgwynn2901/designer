<%	
	 
	Response.Buffer = true
	dim NewCaption, EditCaption, SearchCaption, PasteCaption, CopyCaption, RemoveCaption		
	NewCaption = "New"
	EditCaption = "Edit"
	SearchCaption = "Search"
	PasteCaption = "Paste" 
	CopyCaption = "Copy" 
	RemoveCaption = "Delete"
	AttachCaption = "Attach"	
	RefreshCaption = "Refresh"
	
	If Request.QueryString <> "" Then
	
		dim bHideNew, bHideEdit, bHideSearch, bHidePaste, bHideCopy, bHideRemove, bHideAttach, bHideRefresh
		If CStr(Request.QueryString("HIDENEW")) = "TRUE" Then bHideNew = true
		If CStr(Request.QueryString("HIDEEDIT")) = "TRUE" Then bHideEdit = true
		If CStr(Request.QueryString("HIDESEARCH")) = "TRUE" Then bHideSearch = true
		If CStr(Request.QueryString("HIDEPASTE")) = "TRUE" Then bHidePaste = true
		If CStr(Request.QueryString("HIDECOPY")) = "TRUE" Then bHideCopy = true
		If CStr(Request.QueryString("HIDEREMOVE")) = "TRUE" Then bHideRemove = true
If CStr(Request.QueryString("HIDEATTACH")) = "TRUE" Then bHideAttach = true
If CStr(Request.QueryString("HIDEREFRESH")) = "TRUE" Then bHideRefresh = true

		If CStr(Request.QueryString("NEWCAPTION")) <> "" Then NewCaption = CStr(Request.QueryString("NEWCAPTION"))
		If CStr(Request.QueryString("EDITCAPTION"))<> "" Then EditCaption = CStr(Request.QueryString("EDITCAPTION"))
		If CStr(Request.QueryString("SEARCHCAPTION")) <> "" Then SearchCaption = CStr(Request.QueryString("SEARCHCAPTION"))
		If CStr(Request.QueryString("PASTECAPTION")) <> "" Then PasteCaption = CStr(Request.QueryString("PASTECAPTION"))
		If CStr(Request.QueryString("COPYCAPTION")) <> "" Then CopyCaption = CStr(Request.QueryString("COPYCAPTION"))
		If CStr(Request.QueryString("REMOVECAPTION")) <> "" Then RemoveCaption = CStr(Request.QueryString("REMOVECAPTION"))
If CStr(Request.QueryString("ATTACHCAPTION")) <> "" Then AttachCaption = CStr(Request.QueryString("ATTACHCAPTION"))
If CStr(Request.QueryString("REFRESHCAPTION")) <> "" Then RefreshCaption = CStr(Request.QueryString("REFRESHCAPTION"))		
	End If
%>
<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<SCRIPT>
var InScriptlet = (typeof(window.external.version)=="string");
</SCRIPT>

<SCRIPT LANGUAGE=vbscript>
Sub SwitchMenu(ID)
	ID.classname = "UPMENU"
End Sub

Sub SwitchBack(ID)
	ID.classname = "DOWNMENU"
End Sub

Function SwitchClick(ID)
	ID.classname = "CLICKMENU"
End Function

<%	If Not bHideEdit Then %>
Sub TDEDIT_OnMouseDown
	SwitchClick(TDEDIT)
End Sub

Sub TDEDIT_OnMouseUp
	SwitchMenu(TDEDIT)
	window.external.raiseEvent "EDITBUTTONCLICK", TDEDIT
End Sub
<%	End If %>

<%	If Not bHideNew Then %>
Sub TDNEW_OnMouseDown
	SwitchClick(TDNEW)
End Sub

Sub TDNEW_OnMouseUp
	SwitchMenu(TDNEW)
	window.external.raiseEvent "NEWBUTTONCLICK", TDNEW
End Sub
<%	End If %>

<%	If Not bHideCopy Then %>
Sub TDCOPY_OnMouseDown
	SwitchClick(TDCOPY)
End Sub

Sub TDCOPY_OnMouseUp
	SwitchMenu(TDCOPY)
	window.external.raiseEvent "COPYBUTTONCLICK", TDCOPY
End Sub
<%	End If %>

<%	If Not bHidePaste Then %>
Sub TDPASTE_OnMouseDown
	SwitchClick(TDPASTE)
End Sub

Sub TDPASTE_OnMouseUp
	SwitchMenu(TDPASTE)
	window.external.raiseEvent "PASTEBUTTONCLICK", TDPASTE
End Sub
<%	End If %>

<%	If Not bHideSearch Then %>
Sub TDSEARCH_OnMouseDown
	SwitchClick(TDSEARCH)
End Sub

Sub TDSEARCH_OnMouseUp
	SwitchMenu(TDSEARCH)
	window.external.raiseEvent "SEARCHBUTTONCLICK", TDSEARCH
End Sub
<%	End If %>

<%	If Not bHideRemove Then %>
Sub TDRemove_OnMouseDown
	SwitchClick(TDREMOVE)
End Sub

Sub TDREMOVE_OnMouseUp
	SwitchMenu(TDREMOVE)
	window.external.raiseEvent "REMOVEBUTTONCLICK", TDREMOVE
End Sub
<%	End If %>

<%	If Not bHideAttach Then %>
Sub TDATTACH_OnMouseDown
	SwitchClick(TDATTACH)
End Sub

Sub TDATTACH_OnMouseUp
	SwitchMenu(TDATTACH)
	window.external.raiseEvent "ATTACHBUTTONCLICK", TDATTACH
End Sub
<%	End If %>

<%	If Not bHideRefresh Then %>
Sub TDREFRESH_OnMouseDown
	SwitchClick(TDREFRESH)
End Sub

Sub TDREFRESH_OnMouseUp
	SwitchMenu(TDREFRESH)
	window.external.raiseEvent "REFRESHBUTTONCLICK", TDREFRESH
End Sub
<%	End If %>

</SCRIPT>
</HEAD>
<BODY  bottommargin=0 leftmargin=0 rightmargin=0 topmargin=0 >
<FIELDSET STYLE="BACKGROUND-COLOR:#d6cfbd">
<TABLE  BGCOLOR=#d6cfbd STYLE="HEIGHT:10" ID="TBLMENU">
<TR ALIGN=CENTER>
<%	If Not bHideEdit Then %>
<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDEDIT" NAME="TDEDIT" OnMouseOver="SwitchMenu(TDEDIT)" OnMouseOut="SwitchBack(TDEDIT)"><%=EditCaption%></TD>
<%	End If
	If Not bHideNew Then %>
<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDNEW" NAME="TDNEW" OnMouseOver="SwitchMenu(TDNEW)" OnMouseOut="SwitchBack(TDNEW)"><%=NewCaption%></TD>
<%	End If
	If Not bHideCopy Then %>
<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDCOPY" NAME="TDCOPY" OnMouseOver="SwitchMenu(TDCOPY)" OnMouseOut="SwitchBack(TDCOPY)"><%=CopyCaption%></TD>
<%	End If
	If Not bHidePaste Then %>
<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDPASTE" NAME="TDPASTE" OnMouseOver="SwitchMenu(TDPASTE)" OnMouseOut="SwitchBack(TDPASTE)"><%=PasteCaption%></TD>
<%	End If
	If Not bHideSearch Then %>
<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDSEARCH" NAME="TDSEARCH" OnMouseOver="SwitchMenu(TDSEARCH)" OnMouseOut="SwitchBack(TDSEARCH)"><%=SearchCaption%></TD>
<%	End If
	If Not bHideRemove Then %>
<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDREMOVE" NAME="TDREMOVE" OnMouseOver="SwitchMenu(TDREMOVE)" OnMouseOut="SwitchBack(TDREMOVE)"><%=RemoveCaption%></TD>
<%	End If
	If Not bHideAttach Then %>
	<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDATTACH" NAME="TDATTACH" OnMouseOver="SwitchMenu(TDATTACH)" OnMouseOut="SwitchBack(TDATTACH)"><%=AttachCaption%></TD>
<%	End If 
	If Not bHideRefresh Then %>
	<TD style="width:55" CLASS=DOWNMENU ALIGN=CENTER ID="TDREFRESH" NAME="TDREFRESH" OnMouseOver="SwitchMenu(TDREFRESH)" OnMouseOut="SwitchBack(TDREFRESH)"><%=RefreshCaption%></TD>
<%	End If  %>
</TR>
</TABLE>
</FIELDSET>
</BODY>
</HTML>
