<!--#include file="..\lib\common.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">


<script LANGUAGE="JScript">
var BackObjArray = new Array();

function BackObj(inURL)
{
	this.URL = inURL;
}

function AddBackObj(inURL)
{
	var curObj = new BackObj(inURL);
	var idx;
	idx = BackObjArray.length;
	//alert("idx is " + idx);
	BackObjArray.length = BackObjArray.length + 1;
	//alert("length is " + BackObjArray.length);
	BackObjArray[idx] = curObj;
	//alert("BackObjArray[idx] is " + BackObjArray[idx]);
	//document.all.BtnBack.disabled = false;
}

function GetLastBackObj()
{
	var idx, retval;
	retval = null;
	idx = BackObjArray.length -1;
	//alert("idx is " + idx);
	
	if (idx >= 0 )
	{	retval = BackObjArray[idx];
		//alert("BackObjArray[idx] is " + BackObjArray[idx]);	
		BackObjArray.length = idx;
		//alert("BackObjArray.length is " + BackObjArray.length);	
	}
		
//	if (BackObjArray.length == 0)
//		document.all.BtnBack.disabled = true;
		
	return retval;
}
</script>

<script LANGUAGE="JavaScript" FOR="BtnBack" EVENT="onclick">

	var curObj, i;
	curObj = GetLastBackObj();
	if (curObj == null) 
	{
		//alert("At top level.");
		return;
	}	

	// alert("Last query string " + curObj.URL);
	// get the current form name

	/*if (curObj.method == "POST")
	{
		for (i=0; i <  parent.frames("TOP").document.all.length; i++)
		{

			if (parent.frames("TOP").document.all(i).tagName  == "FORM")
			{
				parent.frames("TOP").document.all(i).action = curObj.URL;
				parent.frames("TOP").document.all(i).method = "POST";
				parent.frames("TOP").document.all(i).submit();
				break;
			}
		}
	}
	else*/
		parent.frames("TOP").document.location.href = curObj.URL
	
</script>

</head>
<body TOPMARGIN="0" LEFTMARGIN="0" BOTTOMMARGIN="0" BGCOLOR="#d6cfbd">
<table ALIGN="RIGHT">
<img id="BtnBack" name="BtnBack" SRC="../ah/..\images\backup.gif" ALIGN="LEFT" VALIGN="BOTTOM" ALT="Back">
</table>
</body>

</html>