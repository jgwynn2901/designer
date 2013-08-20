<html>
<head>
<SCRIPT LANGUAGE="Javascript">
function VendorSearch()
{
	this.VendorID;
	this.Selected = false;
}

function NetworkSearch()
{
	this.NetID;
	this.Selected = false;
}

var AccVendObj = new AVSearchObj();
var VendObj = new VendorSearch();
var NetObj = new NetworkSearch();
</SCRIPT>

<meta name="VI60_defaultClientScript" content="VBScript">
<title>Carrier Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="#d6cfbd">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd" width="4"><div id><nobr>Vendor ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd" width="2"><div id><nobr>City</div></td>			
			<td class="thd" width="2"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>ZIP Code</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">



</tbody>
</table>
</div>
</fieldset>
<SCRIPT LANGUAGE="VBScript">

</SCRIPT>
</body>
</html>
