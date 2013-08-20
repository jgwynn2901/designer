<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	On Error Resume Next
	
	ACTION = CStr(Request.Form("SAVEACTION"))
	SQL_STRING = Request.Form("SAVEDATA")		
	If ACTION = "UPDATE" Then
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "ATTRIBUTE_OVERRIDE", "ATTRIBUTEOVERRIDE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"Carrier " & ACTION)
	Elseif ACTION = "INSERT" Then
		InsertSQL = ""
		NewOVID = NextPkey("ATTRIBUTE_OVERRIDE","ATTRIBUTEOVERRIDE_ID")
		InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "ATTRIBUTE_OVERRIDE", "ATTRIBUTEOVERRIDE_ID", NewOVID)		 
		Set RSUpdate = Conn.Execute(InsertSQL)
		strError = CheckADOErrors(Conn,"Carrier " & ACTION)
		If strError = "" Then
			Response.write("parent.frames('WORKAREA').UpdateStatus('" & NewOVID &  "');")	
		End If
	End If
	
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "AttributeOverride", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames('WORKAREA').SetSpan ("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.")
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').SetSpan("Update successful.")
<%	End If  %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
