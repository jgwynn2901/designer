<HTML>
<HEAD>
</HEAD>
<BODY>
<script language="JScript" src="../_ScriptLibrary/rs.htm"></script>
<script src="../jQuery/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script language="JScript">
function doesUserExist(cName)
{
	// Ajax call start
	var check;
	var request = $.ajax({
	url: '../_ScriptLibrary/remoteScriptMethods.asp?method=UserExists&parms=' + cName,
	type: "POST",
	cache: false,
	async: false,
	success :function(results){
		if (results === "0")
			check = false;
		else
			check = true;	
	},
	error :function(request){
		check = true;
		parent.frames("WORKAREA").isJqueryError = true;
		writeError (request.statusText);
	}
	});	//Ajax call end
	return check;	
}

function createPwd(cPrefix, nLen)
{
    var nRnd;
    var cPWD, nSet, lGotNumber, lGotUpper, lGotLower;
	
    cPWD = cPrefix + "";
    lGotNumber = false;
    lGotUpper = false;
    lGotLower = false;
    while (cPWD.length < nLen)
		{
		nSet = Math.round(Math.random() * 2 + 1);
		switch (nSet)
			{
			case 1:
				nRnd = getUpper();
				lGotUpper = true;
				break;
			case 2:
				nRnd = getLower();
				lGotLower = true;
				break;
			case 3:
				nRnd = getNumber();
				lGotNumber = true;
			}
		cPWD += String.fromCharCode(nRnd);
		}
	if (!lGotUpper)
		cPWD += String.fromCharCode(getUpper());
	if (!lGotNumber)
		cPWD += String.fromCharCode(getNumber());
	if (!lGotLower)
		cPWD += String.fromCharCode(getLower());
	return(cPWD);
}

function getDefaultPassword()
{
    var d = new Date();
    return "Password"+(Number(d.getMonth())+1)+d.getFullYear();
}

function getNumber()
{
return (Math.round(("9".charCodeAt(0) - "0".charCodeAt(0) + 1) * Math.random() + "0".charCodeAt(0)));
}

function getUpper()
{
return (Math.round(("Z".charCodeAt(0) - "A".charCodeAt(0) + 1) * Math.random() + "A".charCodeAt(0)));
}

function getLower()
{
return (Math.round(("z".charCodeAt(0) - "a".charCodeAt(0) + 1) * Math.random() + "a".charCodeAt(0)));
}			

function writeError(cMsg)
{
	  var w = window.open("","error_window","width=500,height=300,toolbar=no,location=no,directories=no,status=no,menubar=no")	  	  
	  w.document.write("<HTML>");
	  w.document.write("<BODY>");
	  w.document.write("<CENTER>");
	  w.document.write("<H2>Remote Scripting Call Returned the following:</H2>");
	  w.document.write("<TABLE border=1 cellpadding=10 bgcolor=#dddddd><TR><TD>");
	  w.document.write(cMsg)
	  w.document.write("</TD></TR></TABLE>");
	  w.document.write("<FORM id=form1 name=form1><INPUT type=button value=\" OK \" onclick=self.close()></FORM>");
	  w.document.write("</CENTER>");
	  w.document.write("</BODY>");
	  w.document.write("</HTML>");
}
</script>
</BODY>
</HTML>
