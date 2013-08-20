<%@ LANGUAGE="VBSCRIPT" %>
<%  Option Explicit
	on error resume next
	Response.expires=0
	
	dim strList ' Global
	If Session("CurrentUser").CanUserDo("Account Maintenance","MODIFY") <> true  Then 
		Response.redirect "../AccessDenied.asp"
	End If
	

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../fnsnet.css">
<TITLE>Data Cache list</TITLE>
</head>
<body CLASS="DataFrame" align=center leftmargin=0>
<form name="frmCache" method="post" action="AM_DataCacheMaintenance.asp" >
<p></p>

<table width="100%">
   <tr><td width="100%" class=label bgcolor=teal><font color=white> Data Cache List Maintenance</td></tr>
</table>
<table cellspacing=20> 
   <tr>
     <td class="label"> Data Cache entries</td>
     <td>
         <SELECT NAME="Cache Entries">
          <OPTION VALUE="1">VALID_VALUE       </option>
          <OPTION VALUE="2">TRANSMISSION_TYPE </option>
	      <OPTION VALUE="3">RESUBMIT_REASON   </option>
	      <OPTION VALUE="3">LOB               </option>
	      <OPTION VALUE="3">SITE              </option>
         </SELECT>   
     </td>
   </tr>
   <tr>
      <td><input type=submit name="listaction" value="Release All Entries " onClick='onClickSubmit()'></td>
      <td><input type=hidden name="inpFlag" value="dee" ></td>
   </tr>
</table>
</form>


</body>
<script language="VBScript">

sub onClickSubmit()
      
<% 

if request("inpFlag") = "dee" then
	  strlist = Application("StaticDataObj").getvaluenamesforgroupname("PRINTER_NAMES")
	  Application.Lock
	  Application("StaticDataObj").ReleaseGenericObjArray()
	  Application.Unlock
'	  strlist = Application("StaticDataObj").getvaluenamesforgroupname("PRINTER_NAMES")
end if
   %>
   
  
end sub

   
</script>

</html>
