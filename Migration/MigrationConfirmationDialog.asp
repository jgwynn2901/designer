<!--#include file="..\lib\common.inc"-->
<html><head>
<title>Migration Confirmation Dialog</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script LANGUAGE="vbscript">
   
	Sub btnYes_onclick
		window.returnValue="true"
		self.close()
	End Sub
	
	Sub btnNo_onclick
		window.returnValue="false"
		self.close()
	End Sub
	
</script></head>
<body BGCOLOR="<%=BODYBGCOLOR%>"><p><br /></p>
<table>
<tr><td><font size="4" color="red"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;WARNING YOU ARE GOING TO PRODUCTION...!!!</b></font> </td></tr>
<tr><td CLASS="LABEL" style="font-size:13px;">&nbsp;Click <b>YES</b> to proceed with migration, <b>NO</b> to NOT proceed with migration!</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td align="center"><button NAME="btnYes" STYLE CLASS="STDBUTTON">Yes</button>&nbsp;&nbsp;
<button NAME="btnNo" STYLE CLASS="STDBUTTON">No</button></td></tr>
</table>
</body></html>
