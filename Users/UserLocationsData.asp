<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next
	
%>
<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Permissions Data</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedACCID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedACCID = document.all.tblFields.rows(idx).getAttribute("ACCID")
	Else
		GetSelectedACCID = ""
	End If
End Function

Function IsSelectedUserLevel
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		 curText = document.all.tblFields.rows(idx).cells("ACCESSTHROUGH").innerText
		 If Right(curText,3) = "(G)" Then
			IsSelectedUserLevel = false
		Else
			IsSelectedUserLevel = true
		End If
	Else
		IsSelectedUserLevel = false
	End If
End Function

</script>


<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Location UserID</div></td>
			<td class="thd"><div id"Div1"><nobr>Account UserID</div></td>
			<td class="thd"><div id"Div2"><nobr>Location AHSID</div></td>
			<td class="thd"><div id"Div3"><nobr>AHSID</div></td>
			<td class="thd"><div id"Div4"><nobr>Client CD</div></td>
			<td class="thd"><div id"Div4"><nobr>Name</div></td>
			<td class="thd"><div id"Div5"><nobr>Phone Number</div></td>
			<td class="thd"><div id"Div6"><nobr>Greeting</div></td>
			<td class="thd"><div id"Div7"><nobr>LOB</div></td>
			
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
dim UID, Conn, RS,  SQLST
UID = CStr(Request.QueryString("UID"))
If UID <> "NEW" And UID <> "" Then  
 
    Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING

	if instr(CONNECT_STRING,"SED")> 0 THEN
	 SQLST = "SELECT LU.LOCATION_USER_ID, LU.ACCOUNT_USER_ID,LU.ACCNT_HRCY_STEP_ID,"
     SQLST = SQLST & " LU.FNS_CLIENT_CD, LU.NAME ,LU.PHONENUMBER ,"
     SQLST = SQLST & " G.TEXT AS GREETING, LU.LOB_CD , LU.LOCATION_AHSID "
     SQLST = SQLST & " FROM LOCATIONS_USER LU, GREETINGS G "
     SQLST = SQLST & " WHERE LU.ACCOUNT_USER_ID = " & UID 
      SQLST = SQLST & " AND G.GREETINGS_ID = LU.GREETINGS_ID "
     
   else
    SQLST = "SELECT LOCATION_USER_ID, ACCOUNT_USER_ID,ACCNT_HRCY_STEP_ID,"
     SQLST = SQLST & " FNS_CLIENT_CD, NAME ,PHONENUMBER ,"
     SQLST = SQLST & " GREETING, LOB_CD , LOCATION_AHSID "
     SQLST = SQLST & " FROM LOCATIONS_USER "
     SQLST = SQLST & " WHERE ACCOUNT_USER_ID = " & UID 
     
    
	end if
	  Set RS = Conn.Execute(SQLST)  

Do While Not RS.EOF And Not RS.BOF

             RsLOCATION_USER_ID       =  "" & RS("LOCATION_USER_ID")
			 RsACCOUNT_USER_ID        =  "" & RS("ACCOUNT_USER_ID")
			 RsACCNT_HRCY_STEP_ID     =  "" & RS("ACCNT_HRCY_STEP_ID")
			 RsFNS_CLIENT_CD          =  "" & RS("FNS_CLIENT_CD")
			 RsPHONE_NUMBER           =  "" & RS("PHONENUMBER")
			 RsGREETING               =  "" & RS("GREETING")                                       
             RsLOB                    =  "" & RS("LOB_CD")                                        
             RsLOCATION_AHSID         =  "" & RS("LOCATION_AHSID")
             RsNAME                   =  "" & RS("NAME")              
%>
	<!--<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" AHS_POL_ID="<%'=oRS("AHS_POLICY_ID")%>" AHS_ID="<%'=oRS("ACCNT_HRCY_STEP_ID")%>">-->
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  ACCID="<%=RsLOCATION_USER_ID%>" AHSID = "<%=RsACCNT_HRCY_STEP_ID%> ">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsLOCATION_USER_ID)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsACCOUNT_USER_ID)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsLOCATION_AHSID)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsACCNT_HRCY_STEP_ID)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsFNS_CLIENT_CD)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsNAME)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsPHONE_NUMBER)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsGREETING)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(RsLOB)%></td>
	
	</tr>

<%
	    RS.MoveNext
  Loop
	  RS.Close 
	   Set RS  = Nothing
	  Conn.Close
	 Set Conn = Nothing
 End If
%>
</tbody>
</table>
</div>
</BODY>
</HTML>