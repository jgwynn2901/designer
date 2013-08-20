<%@ LANGUAGE=VBSCRIPT %>
<!--#INCLUDE file="../_ScriptLibrary/rs.asp"-->

<SCRIPT RUNAT=SERVER LANGUAGE=JScript>

/* remote scripting functions	-	ra	- nov 2003	*/

function serverMethods()
{
  this.doesUserExist = doesUserExist;
  this.createPassword = createPassword;
}
var public_description = new serverMethods();
var CONNECT_STRING = Session("ConnectionString");

RSDispatch();
</script>

<SCRIPT RUNAT=SERVER LANGUAGE=JScript>
function doesUserExist(cName)
{
var oRS, oConn;
var cSQL, lResult;

oConn = new ActiveXObject("ADODB.Connection");
oConn.Open(CONNECT_STRING);
cSQL = "Select USER_ID From USERS Where Name = '" + cName + "'";
oRS = oConn.execute(cSQL);
lResult = !oRS.eof;
oRS.close();
oConn.close();
return lResult;
}

function createPassword(cPrefix, nLen)
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
</script>
