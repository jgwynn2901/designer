<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	On Error Resume Next
	ACTION = Request.Form("TxtAction")
	VID = Request.Form("VID")
	SQL_STRING = Request.Form("TxtSaveData")	
	cServiceTypes = Request.Form("ServiceTypes")	
	cServiceDays = Request.Form("VendorServiceDays")
	cZIPCode = Request.Form("Txt0ZIP")

	    '   get FIPS
    cSQL = "Select FIPS From LOCATION_CODE Where (ZIP='" & cZIPCode & "')"
    Set oRS0 = Conn.Execute(cSQL)
	cFIPS = oRS0.Fields("FIPS").Value
    oRS0.Close
	
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "VENDOR", "VENDOR_ID", "")
		UpdateSQL = replace(UpdateSQL, "ZZ7ZZ", cFIPS)
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Vendor " & ACTION)
		if len(cServiceTypes) <> 0 then
			updServTypes VID, cServiceTypes, False
		end if
		if len(cServiceDays) <> 0 then
			updServDays VID, cServiceDays, false
		end if
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewVID = CLng(NextPkey("VENDOR","VENDOR_ID"))
		If NewVID > 0 Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "VENDOR", "VENDOR_ID", NewVID)
			InsertSQL = replace(InsertSQL, "ZZ7ZZ", cFIPS)
			Set RSUpdate = Conn.Execute(InsertSQL)
			strError = CheckADOErrors(Conn,"Vendor " & ACTION)
			If strError = "" Then 
				Response.write("parent.frames('WORKAREA').UpdateVID('" & NewVID &  "');")
				updServTypes NewVID, cServiceTypes, True
				updServDays NewVID, cServiceDays, true
			end if
		Else
			strError = "Unable to obtain next primary key for VENDOR table."
		End If			
	End If
	Conn.Close
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "VENDOR", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames('WORKAREA').SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').UpdateStatus('Update successful.');
		parent.frames('WORKAREA').ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		
<%	End If

sub updServTypes(nVID, cNewTypes, lNewVID)
'				format is 2;4;5
dim cSQL, aTypes, x

if not lNewVID then
	cSQL = "Delete from VENDOR_SERVICE Where Vendor_ID = " & nVID
	Conn.execute cSQL
end if	
aTypes = split(cNewTypes, ";")
for x=lbound(aTypes) to ubound(aTypes)
	cSQL = "Insert into VENDOR_SERVICE values (" & nVID & "," & aTypes(x) & ")"
	Conn.execute cSQL
next	
end sub

sub updServDays(nVID, cNewDays, lNewVID)
'		format is MON$8:00$18:00~TUE$8:00$18:00
dim cSQL, aDays, aDayHours, x

if not lNewVID then
	cSQL = "Delete from VENDOR_DAY Where Vendor_ID = " & nVID
	Conn.execute cSQL
end if	
aDays = split(cNewDays, "~")
for x=lbound(aDays) to ubound(aDays)
	aDayHours = split(aDays(x), "$")
	cSQL = "Insert into VENDOR_DAY values (" & nVID & ",'" & aDayHours(0) & "','Y','" & aDayHours(1) & "','" & aDayHours(2) & "')"
	Conn.execute cSQL
next	
end sub
 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
