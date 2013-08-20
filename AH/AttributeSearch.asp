<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</HEAD>
<BODY>

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
<%

		Set Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = "DRIVER={Microsoft ODBC for Oracle};SERVER=190.15.5.4;ConnectString=FNS;UID=FNSOWNER;PWD=CTOWN"
		Conn.Open ConnectionString
		SQLST = "SELECT NAME, ATTRIBUTE_ID FROM ATTRIBUTE WHERE UPPER(NAME) LIKE '" & WHERECLS & "%' ORDER BY NAME" 
		Set RS = Conn.Execute(SQLST)%>


Sub RenderRows()
	Dim objRow 

	if g_rsUsers.RecordCount <> 0 Then
		g_rsUsers.MoveFirst
		do while Not g_rsUsers.eof
			set objRow = g_rsUsers.Fields
			Response.Write("<tr DYNKEY=""" & objRow("CLAIMANTID") & """ ")
			Response.Write("CLASS=""ResultRow"" OnClick=""Javascript:multiselect(this);"" OnDblClick=""Javascript:dblhighlight(this);"">")
			Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIMANT_LAST_NAME")) & "</td>")
			Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIMANT_FIRST_NAME")) & "</td>")
			Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIMANT_DOB")) & "</td>")
			formatssno = RenderCellContents(objRow("CLAIMANT_SSNO"))
			if trim(formatssno) <> "&nbsp;" then
				formatssno = mid(formatssno ,1,3) & "-" &  mid(formatssno ,4,2) & "-" & mid(formatssno ,6,4)
			end if 
			Response.Write("<td><NOBR>" & formatssno & "</td>")
			If Enhanced = "on" Then
				Response.Write("<td><NOBR>" & RenderCellContents(objRow("CRA_NUMBER")) & "</td>")
				Response.Write("<td><NOBR>" & RenderCellContents(objRow("ACCOUNT_ACRONYM")) & "</td>")
				Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIM_NUM")) & "</td>")
				Response.Write("<td><NOBR>" & RenderCellContents(objRow("EMPLOYER_NAME")) & "</td>")
			End If
			formatphone = RenderCellContents(objRow("CLAIMANT_PHONE"))
			if trim(formatphone) <> "&nbsp;" then
				formatphone = "(" & mid(formatphone,1,3) & ") " & mid(formatphone,4,3) & "-" & mid(formatphone,7,4)
			End If
			Response.Write("<td><NOBR>" & formatphone & "</td>")
			Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIMANT_ADDRESS")) & "</td>")
			Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIMANT_CITY")) & "</td>")
			Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIMANT_STATE")) & "</td>")
			If Len(objRow("CLAIMANT_ZIP")) > 5 Then
				Response.Write("<td><NOBR>" & RenderCellContents(Mid(objRow("CLAIMANT_ZIP"),1,5)) & "-" & RenderCellContents(Mid(objRow("CLAIMANT_ZIP"),6,4)) & "</td>")
			Else
				Response.Write("<td><NOBR>" & RenderCellContents(objRow("CLAIMANT_ZIP")) & "</td>")
			End If
			IF CurrentUser("AdminLevel") > "0" Then 
				Response.Write("<td><NOBR>" & RenderCellContents(objRow("SITEID")) & "</td>")
				Response.Write("<td><NOBR>" & RenderCellContents(objRow("SITEDESC")) & "</td>")
			End If

			Response.Write("</tr>" & chr(13) )

			g_rsUsers.MoveNext
		loop
	else 
		Response.Write("<tr CLASS=""ResultRow"" >")
		Response.Write("<td align=""center"" colspan=""10"">" & "No matching Claimants found. Check your selection criteria." & "</td>" )
		Response.Write("</tr>" & chr(13) )
	end if
End Sub
%>


<html>
<!--#include file="../lib/vbs_server_common.inc"-->
<!--#include file="../lib/msg_processing.inc"-->
<!--#include file="../lib/msg.inc"-->
<!--#include file="../lib/DynTreePOS.inc"-->

<head>
<title>Claimant Search Results</title>
<!--#include file="../lib/tablecommon.inc"-->
<!--#include file="../lib/NavKeys.inc"-->
<LINK REL="Stylesheet" TYPE="text/css" HREF="../ICMS.css"> 

<SCRIPT LANGUAGE="VBScript">
<!--
<%

<%
'---------------------------------------------------------
' CheckKey
' Purpose:  handler for the OnKeyDown event of the 
'			Document
' Inputs:   n/a
' Returns:  n/a
'---------------------------------------------------------
%>
Sub ProcessKey(keycode,altkey)
	dim i
	dim stuff

	select case keycode
		case 8:
			window.event.keyCode = 0
			window.event.returnValue = 0
		case 38:
			call relativemultiselect( Document.all.tblResult, -1 )
		case 40:
			call relativemultiselect( Document.all.tblResult, 1 )
		case 13:
			i = getselectedindex( Document.all.tblResult )
			if 0 < i then
				dblhighlight(Document.all.tblResult.rows(i))
			end if
		case else:
	end select

Sub Document_OnKeyDown()

	call ProcessKey(window.event.keyCode, window.event.altKey)

End Sub

Sub Window_OnLoad()
	if 0 < Document.all.tblResult.rows.length then
		call multiselect( Document.all.tblResult.rows(1))
		call Document.all.tblResult.focus()
	end if
End Sub


-->
</SCRIPT>

<SCRIPT LANGUAGE="Javascript">
<!--
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	GetUser( objRow.getAttribute("DYNKEY") )
}
-->
</SCRIPT>
</head>

<body topmargin="0" leftmargin="0">
<DIV NAME=TEST ID=TEST></DIV>
<% DisplayProcessingMsg "Searching for matching records..." %>
<div align="left" id="SCREEN_DATA" style="display:block;">
<LABEL CLASS="LABEL">Search Results:<BR></LABEL>
<table NAME=tblResult ID=tblResult border="1" width="100%" cellspacing="0" cellpadding="3">
<TR CLASS="ResultHeader">
<TD CLASS="LABEL"><NOBR>Last Name</TD>
<TD CLASS="LABEL"><NOBR>First Name</TD>
<TD CLASS="LABEL"><NOBR>D.O.B.</TD>
<TD CLASS="LABEL"><NOBR>Soc.Sec.#</TD>
<% If Enhanced = "on" Then %>
	<TD CLASS="LABEL"><NOBR>Referral #</TD>
	<TD CLASS="LABEL"><NOBR>Acct Acronym</TD>
	<TD CLASS="LABEL"><NOBR>Claim #</TD>
	<TD CLASS="LABEL"><NOBR>Employer Name</TD>
<% End If %>
<TD CLASS="LABEL"><NOBR>Phone</TD>
<TD CLASS="LABEL"><NOBR>Address</TD>
<TD CLASS="LABEL"><NOBR>City</TD>
<TD CLASS="LABEL"><NOBR>State</TD>
<TD CLASS="LABEL"><NOBR>Zip</TD>
<%IF CurrentUser("AdminLevel") > "0" Then  %>
	<TD CLASS="LABEL"><NOBR>Site ID</TD>
	<TD CLASS="LABEL"><NOBR>Site Name</TD>
<% End If %>
<%
RenderRows
g_rsUsers.close
%>
</table>
</DIV> 

</BODY>
</HTML>
