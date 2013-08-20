<!--#include file="..\lib\common.inc"-->
<%
Response.Expires=0
dim oConn, cSQL, cFrameID, oRS, cMsg, lError
cFrameID = Request.QueryString("FRAMEID")
cMsg = """" & """"
If cFrameID <> "" Then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL = "{call Designer.sp_deleteFrame(" & cFrameID & ", {resultset 1, StatusMsg, StatusNum, cCFs})}"
	Set oRS = oConn.Execute(cSQL)
	lError = oRS.fields("StatusNum") <> "0"
	select case oRS.fields("StatusNum")
		case  "-1"
			' shared count
			cMsg = oRS.fields("StatusMsg")
			cMsg ="""" & cMsg & """" & " & vbCRLF & " & """" & "This frame is in use in Callflows" & """" & " & vbCRLF & vbCRLF & " & """" & oRS.fields("cCFs") & """"
		case "0"
			cMsg = """" & "Frame " & cFrameID & " deleted!" & """"
		case "-2"
			' overflow, meaning there are many callflows in the 'used in' list
			cMsg ="""" & "This frame is in use in more than 100 Callflows." & """"
		case else
			cMsg = """" & "Database Error " & oRS.fields("StatusNum") & """" & " & vbCRLF & " & """" & oRS.fields("StatusMsg") & """"
	end select
	oRS.close
	set oRS = nothing
	oConn.Close
	set Oconn = nothing
end if 	
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<script Language=VBScript>
msgbox <%=cMsg%>, vbExclamation, "FNSDesigner"
<%
if not lError then
%>
	a=parent.parent
	a.frames("WORKAREA").location.reload
<%
end if
%>	
</script>	
</BODY>
</HTML>
