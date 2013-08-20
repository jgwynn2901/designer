<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\AHSTree.inc"-->
<%
	ACTION = CStr(Request.QueryString("ACTION"))
	AHSID = CStr(Request.QueryString("AHSID"))
	USEWHERECLAUSE = CStr(Request.QueryString("USEWHERECLAUSE"))
	WHERECLAUSE = CStr(Request.QueryString("WHERECLAUSE"))
	WHERECLAUSE = Replace(WHERECLAUSE,"|","%")
	MUSTINCLUDE = CStr(Request.QueryString("MUSTINCLUDE"))
	MUSTEXCLUDE = CStr(Request.QueryString("MUSTEXCLUDE"))	
	NODEDELIM = CStr(Request.QueryString("NODEDELIM"))	
	If ACTION = "REMOVE" Then
		RemoveFilter "AHSID=" & AHSID, "DESIGNER_AHSFILTER" 
'		if Session("AHSTreeShowAllNodes").exists("AHSID=" & AHSID) then
'			Session("AHSTreeShowAllNodes").Remove("AHSID=" & AHSID)
'		end if
	ElseIf ACTION = "ADD" then	'	and Request.QueryString("SHOWALL") = "" Then
		if WHERECLAUSE <> "" then
			RemoveFilter "AHSID=" & AHSID, "DESIGNER_AHSFILTER"
			SetFilter "AHSID=" & AHSID, "DESIGNER_AHSFILTER", USEWHERECLAUSE,WHERECLAUSE, MUSTINCLUDE, MUSTEXCLUDE, NODEDELIM
		end if
 	End If
'	if Request.QueryString("SHOWALL") = "Y" then
'		addShowAllNode("GRP=" & AHSID)
'	else
'		if Request.QueryString("SHOWALL") = "N" then
'			Session("AHSTreeShowAllNodes").Remove("AHSID=" & AHSID)
'			RemoveFilter "AHSID=" & AHSID, "DESIGNER_AHSFILTER" 
'		end if
'	end if
	
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE></TITLE>
</HEAD>
<BODY>
<%=DisplayFilter%>
</BODY>
</HTML>
