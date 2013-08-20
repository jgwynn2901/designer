<%
	Response.Expires = 0
	Response.Buffer = true
%>

<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Coverage Code Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		'Parent.frames("TOP").document.all.XREFID.value = ""
	end if
End Sub

Function GetXREFID
   GetXREFID = getmultipleindex(document.all.tblFields, "XREFID")
End Function

Function GetXREFIDName
	GetXREFIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function

</script>
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "XREFID")
		return objRow.getAttribute("XREFID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
		
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>XREF ID</div></td>
			<td class="thd"><div id><nobr>A.H. Step ID</div></td>
			<td class="thd"><div id><nobr>Coverage Code</div></td>
			<td class="thd"><div id><nobr>Vendor Designator</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%

	dim RecCount
	RecCount = -1
	If Request.QueryString("SEARCHTYPE") <> "" Then
	RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
			    XREFID            = Request.QueryString("SearchXREFID") & "%"
				AHSID             = Request.QueryString("SearchAHSID") 
				COVERAGECODE      = Request.QueryString("SearchCoverageCode") & "%"
				VENDORDESIGNATOR = Request.QueryString("SearchVendorDesignator") & "%"
			Case "C"
				XREFID            = "%" & Request.QueryString("SearchXREFID") & "%"
				AHSID             =  Request.QueryString("SearchAHSID") 
				COVERAGECODE      = "%" & Request.QueryString("SearchCoverageCode") & "%"
				VENDORDESIGNATOR  = "%" & Request.QueryString("SearchVendorDesignator") & "%"
				
			Case "E"
				XREFID           = Request.QueryString("SearchXREFID")
				AHSID            = Request.QueryString("SearchAHSID") 
				COVERAGECODE     = Request.QueryString("SearchCoverageCode")
				VENDORDESIGNATOR = Request.QueryString("SearchVendorDesignator") 
		End Select
		XREFID = Replace(XREFID, "'", "''")
		'DESCRIPTION = Replace(DESCRIPTION, "'", "''")
		AHSID = Replace(AHSID, "'", "''")
		COVERAGECODE = Replace(COVERAGECODE, "'", "''")
		VENDORDESIGNATOR = Replace(VENDORDESIGNATOR, "'", "''")
		
		If Request.QueryString("SearchAHSID") <> "" Then
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID= " & AHSID
		End If
		If Request.QueryString("SearchXREFID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "COVERAGECODE_CONVERSION_ID LIKE '" & XREFID & "'"
		End If
		If Request.QueryString("SearchCoverageCode") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "COVERAGE_CODE LIKE '" & COVERAGECODE  & "'"
		End If
		
		If Request.QueryString("SearchVendorDesignator") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "VENDOR_DESIGNATOR LIKE '" & VENDORDESIGNATOR & "'"
		End If
		
          Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = "SELECT COVERAGECODE_CONVERSION_ID,ACCNT_HRCY_STEP_ID,COVERAGE_CODE,VENDOR_DESIGNATOR,DESCRIPTION FROM COVERAGECODE_CONVERSION"

		 
			if WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS
			End if
			
			SQLST = SQLST & " ORDER BY COVERAGECODE_CONVERSION_ID "
			
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	

			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" XREFID=''>
	<td COLSPAN="5" NOWRAP CLASS="ResultCell">No managed care branch assignment types found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
				
%>

<tr ID="FieldRow" CLASS="ResultRow"  OnClick="Javascript:multiselect(this);" XREFID='<%=RS("COVERAGECODE_CONVERSION_ID")%>'>
<td NOWRAP CLASS="ResultCell"ID="NAME" ><%=renderCell(RS("COVERAGECODE_CONVERSION_ID"))%></td>
<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("COVERAGE_CODE"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("VENDOR_DESIGNATOR"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DESCRIPTION"))%></td>
</tr>   

<%
				RS.MoveNext
				Loop
			End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing

		End if

%>

</tbody>
</table>
</div>
</fieldset>
<SCRIPT LANGUAGE="VBScript">
<%	If RecCount >= 0 Then %>
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=RecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=RecCount%>")
	end if		
end if
<%	End If %>
</SCRIPT>
</body>
</html>
