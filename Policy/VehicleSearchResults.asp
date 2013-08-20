<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.Buffer = true
Response.AddHeader  "Pragma", "no-cache"
%>
<!--#include file="..\lib\tablecommon.inc"-->
<% Response.Expires = 0 %>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vehicle Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.VID.value = ""
	end if
End Sub

Function GetVID
	GetVID = getmultipleindex(document.all.tblFields, "VID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "VID")
		return objRow.getAttribute("VID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
		
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Vehicle Id</div></td>
			<td class="thd"><div id><nobr>Make</div></td>
			<td class="thd"><div id><nobr>Model</div></td>
			<td class="thd"><div id><nobr>Year</div></td>
			<td class="thd"><div id><nobr>Plate</div></td>
			<td class="thd"><div id><nobr>State</div></td>
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
				VID = Request.QueryString("SearchVID") & "%"
				POLICY_ID = Request.QueryString("SearchPOLICY_ID") & "%"
				VIN = Request.QueryString("SearchVIN") & "%"
				CARYEAR = Request.QueryString("SearchYEAR") & "%"
				MAKE = Request.QueryString("SearchMAKE") & "%"
				MODEL = Request.QueryString("SearchMODEL") & "%"
				LICENSE_PLATE = Request.QueryString("SearchLICENSE_PLATE") & "%"
				LICENSE_PLATE_STATE = Request.QueryString("SearchLICENSE_PLATE_STATE") & "%"
				REGISTRATION_STATE = Request.QueryString("SearchREGISTRATION_STATE") & "%"
				COLOR = Request.QueryString("SearchCOLOR") & "%"
			Case "C"
				VID = "%" & Request.QueryString("SearchVID") & "%"
				POLICY_ID = "%" & Request.QueryString("SearchPOLICY_ID") & "%"
				VIN = "%" & Request.QueryString("SearchVIN") & "%"
				CARYEAR = "%" & Request.QueryString("SearchYEAR") & "%"
				MAKE = "%" & Request.QueryString("SearchMAKE") & "%"
				MODEL = "%" & Request.QueryString("SearchMODEL") & "%"
				LICENSE_PLATE = "%" & Request.QueryString("SearchLICENSE_PLATE") & "%"
				LICENSE_PLATE_STATE = "%" & Request.QueryString("SearchLICENSE_PLATE_STATE") & "%"
				REGISTRATION_STATE = "%" & Request.QueryString("SearchREGISTRATION_STATE") & "%"
				COLOR = "%" & Request.QueryString("SearchCOLOR") & "%"
			Case "E"
				VID = Request.QueryString("SearchVID")
				POLICY_ID = Request.QueryString("SearchPOLICY_ID")
				VIN = Request.QueryString("SearchVIN")
				CARYEAR = Request.QueryString("SearchYEAR")
				MAKE = Request.QueryString("SearchMAKE")
				MODEL = Request.QueryString("SearchMODEL")
				LICENSE_PLATE = Request.QueryString("SearchLICENSE_PLATE")
				LICENSE_PLATE_STATE = Request.QueryString("SearchLICENSE_PLATE_STATE")
				REGISTRATION_STATE = Request.QueryString("SearchREGISTRATION_STATE")
				COLOR = Request.QueryString("SearchCOLOR")
		End Select
		VID = Replace(VID,"'","''")
		POLICY_ID = Replace(POLICY_ID,"'","''")
		VIN = Replace(VIN,"'","''")
		CARYEAR = Replace(CARYEAR,"'","''")
		MAKE = Replace(MAKE,"'","''")
		MODEL = Replace(MODEL,"'","''")
		LICENSE_PLATE = Replace(LICENSE_PLATE,"'","''")
		LICENSE_PLATE_STATE = Replace(LICENSE_PLATE_STATE,"'","''")
		REGISTRATION_STATE = Replace(REGISTRATION_STATE,"'","''")
		COLOR = Replace(COLOR,"'","''")
		
		If Request.QueryString("SearchPOLICY_ID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(POLICY_ID) LIKE '" & UCASE(POLICY_ID)  & "'"
		End If
		If Request.QueryString("SearchVID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "VEHICLE_ID LIKE '" & VID & "'"
		End If
		If Request.QueryString("SearchVIN") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(VIN) LIKE '" & UCASE(VIN) & "'"
		End If
		If Request.QueryString("SearchYEAR") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(YEAR) LIKE '" & UCASE(CARYEAR) & "'"
		End If
		If Request.QueryString("SearchMAKE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(MAKE) LIKE '" & UCASE(MAKE) & "'"
		End If
		If Request.QueryString("SearchMODEL") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(MODEL) LIKE '" & UCASE(MODEL) & "'"
		End If
		If Request.QueryString("SearchLICENSE_PLATE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(LICENSE_PLATE) LIKE '" & UCASE(LICENSE_PLATE) & "'"
		End If
		If Request.QueryString("SearchLICENSE_PLATE_STATE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(LICENSE_PLATE_STATE) LIKE '" & UCASE(LICENSE_PLATE_STATE) & "'"
		End If
		If Request.QueryString("SearchREGISTRATION_STATE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(REGISTRATION_STATE) LIKE '" & UCASE(REGISTRATION_STATE) & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT VEHICLE_ID,POLICY_ID,MAKE, MODEL, YEAR, LICENSE_PLATE, LICENSE_PLATE_STATE FROM VEHICLE "
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS
			End If
			SQLST = SQLST &  " ORDER BY MAKE" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" VID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No Vehicles found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  VID='<%=RS("VEHICLE_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("VEHICLE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MAKE"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MODEL"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RS("YEAR"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LICENSE_PLATE"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LICENSE_PLATE_STATE"))%></td>
	</tr>

<%
				RS.MoveNext
				Loop
				
			End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing
	End If
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

