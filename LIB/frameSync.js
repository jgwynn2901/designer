/***********************************************
* Frame syncronization functions - RA - 12/00
************************************************/
var cFunctionToExecute = "";
var oWindow = null;

function executePendingFunction()
{
	eval(cFunctionToExecute);
	cFunctionToExecute = "";
	oWindow = null;
}

function asyncWait4Frame()
{
	if (oWindow.document.readyState == "complete")
		{
		oWindow.document.onreadystatechange = null;
		executePendingFunction();
		}
}

function wait4Frame(oWin, cFunction)
{
if (cFunctionToExecute != "" || oWindow != null)
	alert("Error! Function '" + 	cFunctionToExecute + "' is still pending execution.")
else
	{
	oWindow = oWin;
	cFunctionToExecute = cFunction;
	if (oWindow.document.readyState != "complete")
		oWindow.document.onreadystatechange = asyncWait4Frame;
	else	
		executePendingFunction();
	}
}
