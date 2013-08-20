///////////////////////////////////////////////////////////
// Expression Tree class.  evaluates an expression of the
// format (lVal Op rVal) where lVal and rVal are the 
// leftvalue and rightvalue for the expression, and Op
// is the operation used to combine the values.
// Either the lVal or rVal (or both) may be referneces to
// other CExprTree objects.
//
function CExprTree(lVal, Op, rVal, strName, referenceArray)
{

	var obj = CExprTree.compileCache.itemValue (strName);

	if (obj != (void 0))
	{
		//*DebugWriteLnBr ("I see a precompiled '" + strName + "' in the cache, you'll get that");
		return (obj);
		//*DebugWriteLnBr ("How did I get here?");  // I didn't.
	}

	this.lVal = lVal;
	this.Op = Op;
	this.rVal = rVal;
	this.name = strName;
	this.referenceArray = referenceArray;

	//public methods
	this.lastError = "";

	//if (this.lastError == "")
	//{
		//*DebugWriteLnBr ("Adding '" + strName + "' to cache.");
		CExprTree.compileCache.addItem (strName, this);
	//}

}

Math.MSPerYear = 31536000000;

CExprTree.compileCache = new CReferenceArray();

CExprTree.resultCacheDirty = false;
CExprTree.reVariable = /~(.+[^~])+~/;
CExprTree.reVariableG = /~[^~](\w|\[|\]|:|\.)+~/g;
CExprTree.reString = /\s*(\x22|\x27).*\1\s*/;
CExprTree.aryFunctions = new Array();

///////////////////////////////////////////////////////////
// Adds a new function to the list of available functions
// strName is the name of the function to be used in expressions
// (including the leading $) and objFunction is the function
// to call back when strName is encountered in an expression.
//
CExprTree.addFunction = function (strName, objFunction, cParams)
{

	strName = strName.toUpperCase()
	if (cParams == (void 0)) cParams = -1;
	//alert ("adding function: " + TRACE_getFuncName(objFunction) + " as '" + strName + "' with " + cParams + " params.");
	CExprTree.aryFunctions[strName] = new Array();
	CExprTree.aryFunctions[strName][0] = objFunction;
	CExprTree.aryFunctions[strName][1] = cParams;

}

///////////////////////////////////////////////////////////
// Returns val in a format which can be used directly in
// javascript comparison statements.  If val is an object,
// this function will return the result of a call to
// .getEvaluate, else, if val matches the regexp this.reVariable
// it's value will be determined via a call to 
// ReferenceArray.getValue.  if val matches the regexp defined
// by CExprTree.reString, the function will return it without the
// surrounding delimiters (eg "" or '')  If all else fails,
// the function will return val converted to a number.
//
// If this function encounters an object in the tree with no
// getEvaluate method, it will return (void 0) and
// this.lastError will contain a text descript of what went wrong.
//
CExprTree.prototype.getProcessedVal = function (val)
{
	var result = "";

	if (typeof val == "object")						// if val is an object, then it needs to
	{												// evaluate and format itself

		if (val.getEvaluate != (void 0))			// make sure this method exists
		{
			result = val.getEvaluate();				// Get the answer to this object
			this.lastError = val.lastError;			// Pass any error back up the recurse stack
			return (result);						// all done
		}
		else										// what the hell is this object?
		{
			this.lastError = "getProcessedVal: Unknown object discovered in graph";
			return (void 0);
		}
	}
	else											// otherwise we'll massage it here...
	{

		if (val.charAt(0) == "$") return (val);											// it's a function name, leave it alone

		val = this.replaceVars(val);													// search and replace variable names

		if (val.search(CExprTree.reString) == 0)										// if it looks like a string
		{
			return (val.substring (1, val.length - 1));									// chop the delimeters and return it
		}

		else
		{
			var result = (Number(val));													// if it can be converted to a number
			if (!isNaN(result))
				return (result);														// it must be a number
			else
			{
				switch (val)															// or maybe it's an internal identifier
				{

					case "TRUE":														// like the boolean 'true'
					case "true":
					case "True":
						return (true);
						break;

					case "FALSE":														// or 'false'
					case "false":
					case "False":
						return (false);
						break;

					case "[blank]":														// or maybe "" ?
					case "[BLANK]":
					case "[Blank]":
						return ("");
						break;

					case "[today]":
					case "[TODAY]":
					case "[Today]":
						var d = new Date();
						var dd = String(d.getDate());
						if (dd.length == 1) dd = "0" + dd;
						var mm = String(d.getMonth())
						if (mm.length == 1) mm = "0" + mm;
						var yy = String(d.getFullYear());
						return (dd + mm + yy);


					default:
						this.lastError = "Unknown identifier '" + val + "'";			// then again, mabye it's trash
						return (void 0);
						break;

				}

			}

		}

	}

}

///////////////////////////////////////////////////////////
// Replaces any tokens matching reVariable with it's
// value in the reference array
//
CExprTree.prototype.replaceVars = function(str)
{
	var i;
	if (typeof(str) != "string")
	{
		debugWrite ("replaceVars: " + str + " is not a string, exiting.");
		return (str);
	}

	debugWrite ("replaceVars: searching " + str + " for the pattern " + CExprTree.reVariableG);

	var aryVarTokens = str.match (CExprTree.reVariableG);

	if (this.referenceArray == (void 0))
	{
		debugWrite ("no referenceArray, cant continue.");
		return(str);
	}

	if (aryVarTokens != (void 0))
	{

		for (i=0; i < aryVarTokens.length; i++)
		{
		
			 str = str.replace (aryVarTokens[i],this.referenceArray.itemValue(aryVarTokens[i].match(CExprTree.reVariable)[1]));
			 debugWrite ("replacing " + aryVarTokens[i] + " with " + this.referenceArray.itemValue(aryVarTokens[i].match(CExprTree.reVariable)[1]));
			 //*alert ("replacing " + aryVarTokens[i] + " with " + this.referenceArray.itemValue(aryVarTokens[i].match(CExprTree.reVariable)[1]));

		}

	}

	debugWrite ("returning: " + str);
	return (str);

}

CExprTree.prototype.doFunction = function (strName, strParamList, objThat)
{
	//*alert ("looking for function: " + strName.toUpperCase());

	var func = CExprTree.aryFunctions[strName.toUpperCase()];
	var aryParamList = new Array();
	var objEv;
	var cParams;
	var i;
	
	if (func == (void 0))
	{
		this.lastError = "Unknown function: " + strName;
		return (false);
	}

	aryParamList = splitExpr(strParamList);

	cParams = func[1];
	if (cParams != -1) 
		if (aryParamList.length != cParams)
		{
			objThat.lastError = strName + ": incorrect number of parameters, " + cParams + " expected, " + aryParamList.length + " found.";
			return (false);
		}

	func = func[0];

	for (i=0; i< aryParamList.length; i++)
	{
		objEv = ETreeFrom (aryParamList[i], this.referenceArray);

		if (typeof(objEv) != "object")
		{
			objThat.lastError = strName + " (param " + i + "): " + objEv
			return (false);
		}

		aryParamList[i] = objEv.getEvaluate();
		if (objEv.lastError != "")
		{
			objThat.lastError = objEv.lastError;
			return (false);
		}
	}

	return (func(aryParamList, objThat));

}

///////////////////////////////////////////////////////////
// Returns the evaluated object.
// If the expression is non-trivial (lval or rval are objects)
// This function will recurse via the getProcessedVal function,
// evaluating left-to-right, depth first.
// On error this function will return (void 0), and 
// this.lastError will be a text description of what went wrong.
//
CExprTree.prototype.getEvaluate = function (fIgnoreCache)
{
	var lVal;
	var rVal;
	var result;
	if (fIgnoreCache == (void 0)) fIgnoreCache = false;

	if (!fIgnoreCache)
	{
		if ((CExprTree.resultCache != (void 0)) && (!CExprTree.resultCacheDirty))
		{

			//*DebugWriteLnBr ("Checking for '" + this.name + "' in the cache.");
			var obj = CExprTree.resultCache.itemValue(this.name);
			if (obj != (void 0))
			{
				//*DebugWriteLnBr ("I see the result for '" + this.name + "' (" + obj + ") in the cache, you'll get that.");
				return (obj);
			}
		}
		else
		{
			//*DebugWriteLnBr ("Cache is dirty or not created yet, I'll make a new one.");
			CExprTree.resultCache = new CReferenceArray();
			CExprTree.resultCacheDirty = false;
		}
	}

	lVal = this.getProcessedVal (this.lVal);											// only process the left value for now,
	if (lVal == (void 0))																// in case we can short circuit.
		return (void 0);																// if it won't process, then bail out.																						

	if (this.rVal == '') return (lVal);													// if the right value is empty, then the
																						// expr. must have been exreaneous parens 
																						// like '((a) and b)'
																						//        ^^^

	switch (this.Op.toUpperCase())														// what to do?
	{

		case ".OR.":
		case ".||.":
		case "OR":
		case "||":
			if (lVal) return (true);													// short circuit OR statement if lVal is true
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (lVal || rVal);
			break;

		case ".AND.":
		case ".&&.":
		case "AND":
		case "&&":
			if (!lVal) return (false);													// short circuit AND statement if lVal is false
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (lVal && rVal);
			break;

		case ";":																		// AND w/o short circuit
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (lVal && rVal);
			break;


		case ".EQ.":
		case "=":
		case "==":
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (lVal == rVal);
			break;

		case ".NE.":
		case "<>":
		case "!=":
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (lVal != rVal);
			break;

		case "-":
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (lVal - this.rVal);
			break;

		case "+":
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (lVal + rVal);
			break;

		case ">":
		case ".GT.":
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (Number(lVal) > Number(rVal));
			break;

		case ">=":
		case ".GE.":
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (Number(lVal) >= Number(rVal));
			break;

		case "<=":
		case ".LE.":
			if ((rVal = this.getProcessedVal(this.rVal)) == (void 0)) return (void 0);
			result = (Number(lVal) <= Number(rVal));
			break;

 		case ".IN.":
 		case "IN":
			result = (rVal.search(RegExp("[\s+;]" + lVal + "[\s+;]") != -1));
			break;
 
 		case ".NI.":
 		case "NI":
			result = (rVal.search(RegExp("[\s+;]" + lVal + "[\s+;]") == -1));
			break;

		case "$":
			result = this.doFunction(lVal, this.rVal, this);
			break;

		default:
			result = (void 0)
			this.lastError = "Unknown operator: " + this.Op;

	}

	if (!fIgnoreCache && (this.lastError == ""))
	{
		//*DebugWriteLnBr ("Adding '" + this.name + "' (" + result + ") to the cache.")
		CExprTree.resultCache.addItem (this.name, result);
	}
	else
	{
		//*DebugWriteLnBr ("Not adding '" + this.name + "' to the cache due to an evaluation error.");
	}

	//*alert (this.name + " evaluated to " + result);
	return (result);

}

//
//
// end of CExprTree class. (pretty short, no?)
///////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////
// Helper functions to build a CExprTree from a text
// expression.  No need to include these in the class.
// 

///////////////////////////////////////////////////////////
// Splits strings into array elements using common sense
// rules for expressions.  Anything included in parens
// will be returned as a single token (regardless of it's 
// contents.  Parens in strings (delimited by "" or '') are
// ignored.  Strings will be returned as a single token
// regardless of embedded whitespace.  Whitespace between
// tokens is ignored.
// 
// Don't put anything but normal text ascii chars in the
// string.  Control chars will be returned as single tokens
// or embedded in other tokens, depending on the context.
//
// On error the return type of this function will be 
// "string" (as opposed to "object" with an Array() 
// constructor.)  The value of the string will be (you
// guessed it) the text description of the error.
//
// On success, the function will return an array with
// one element for each token it found.
//
function splitExpr(str)
{
	var i = 0, iToken = 0;					// character position in string and current token number
	var iParenLevel = 0;					// how many levels of parens deep we are currently
	var c = str.length;						// uh... the length of the string
	var fInToken = false;					// currently looking at a token?
	var fInFunctionName = false				// currently looking at a function name token?  (fInToken will be true also)
	var fInFunctionParams = false			// currently looking at a function parameters token?  (fInToken will be true also)
	var fInString = false;					// currently looking at a string?
	var strStringDelimeter = ""				//  if so, which delimeter started it?
	var aryResult = new Array();			// space for return values
	var ch;									// current character


	//alert ("splitExpr: " + str);

	for (i=0; i< c; i++)
	{										// pretty basic stuff, right?
											// it's grown a bit since that comment.
		ch = str.charAt(i);					// will add more soon

		switch (ch)
		{

			//case "=":
			//case "-":
			//case "+":
			//case ";":
			case ',':
			case ' ':
				if (fInToken && ((iParenLevel > 0) || (fInString)))
					if (aryResult[iToken] != (void 0))
						aryResult[iToken] += ch;
					else
						aryResult[iToken] = ch;

				else
					
					if (!fInFunctionName && ((fInToken) && (iParenLevel == 0)))
					{
						fInToken = false;
						iToken++;
						if ((ch != ' ') && (ch != ','))
						{
							aryResult[iToken] = ch;
							iToken++;
						}
					}

				break;

			case '"':
			case "'":
				if (aryResult[iToken] != (void 0))
					aryResult[iToken] += ch;
				else
					aryResult[iToken] = ch;
				if (fInString)
				{
					if (ch == strStringDelimeter)
					{
						fInString = false;
						strStringDelimeter = "";
						if (iParenLevel == 0)
						{
							fInToken = false;
							iToken++;
						}
					}
				}
				else
				{
					fInString = true;
					fInToken = true;
					strStringDelimeter = ch;
				}
				break;

			case '(':		
				if (!fInString)
				{
					if (fInToken && (iParenLevel == 0))
					{
						iToken++;
						if (fInFunctionName)
						{
							aryResult[iToken] = "$"
							iToken++;
							fInFunctionName = false
							fInFunctionParams = true
						}
							
					}

					fInToken = true;
					iParenLevel++;

				}

				if (!fInFunctionParams || (iParenLevel != 1))
					if (aryResult[iToken] != (void 0))
						aryResult[iToken] += ch;
					else
						aryResult[iToken] = ch;

				break;

			case ')':
				if (fInToken)
				{
					if (fInFunctionParams && iParenLevel == 1)
					{
						fInFunctionParams = false;
						fInToken = false;
					}
					else
						if (aryResult[iToken] != (void 0))
							aryResult[iToken] += ch;
						else
							aryResult[iToken] = ch;

						

					if (!fInString)
						if (--iParenLevel < 0)
							return ("Syntax error in '" + str + "' at " + i + ":'" + ch + "', expected seperator.");

				}
				else
					return ("Syntax error in '" + str + "' at " + i + ":'" + ch + "', expected token.");

				break;

			default:
				if (!fInToken && (ch == "$"))
					fInFunctionName = true;

				fInToken = true;

				if (aryResult[iToken] != (void 0))
					aryResult[iToken] = aryResult[iToken] + ch;
				else
					aryResult[iToken] = ch;
				break;
		}

	}

	if (iParenLevel != 0)
		return ("Syntax Error, " + iParenLevel + " Mismatched open parens.");

	return (aryResult);

}

///////////////////////////////////////////////////////////
// Compiles a string representation of an expression into a 
// expression tree structure.
// This function returns a reference to a CExprTree instance which
// represents the root node of the copmiled expression 'strExpression'.
// Call this function to turn "(true or false)" (or any other
// expression) into an compiled expression object.  Then call
// et.getEvaluate to find out the result of the expression.
//
// objReferenceArray is the reference to the reference array 
// that holds values for the varaibles used in this expression.
// If the function contains no variables then this parameter
// can be safely ignored.
//
// On error this function will return a typeof "string", which
// contains a text description of what went wrong.
//
// On success, the function returns the root node of an 
// expression tree.
//

function ETreeFrom (strExpression, objReferenceArray)
{
	var aryValues;
	var cArgs;

	if (typeof (aryValues = splitExpr (strExpression)) != "object")						// split the expression up
		return (aryValues);																// if we get a string as the 
																						// result, something bad happened.

	cArgs = aryValues.length;															// how many tokens did we find

	switch (cArgs)
	{

		case 3:																			// got 3, it's a full expression
		{																				// lval=[0]; Op=[1]; rval=[2]

			if (aryValues[0].search(/\s*\(/) == 0)										// if the value starts w/ a '(', its
			{																			// another expression, recurse
				aryValues[0] = ETreeFrom (aryValues[0], objReferenceArray);
				if (typeof(aryValues[0]) != "object") return (aryValues[0])				// if recursing didnt return an object
			}																			// something bad happened.  Pass the buck.

			if (aryValues[2].search(/\s*\(/) == 0)										// now do the same for the right value.
			{
				aryValues[2] = ETreeFrom (aryValues[2], objReferenceArray);
				if (typeof(aryValues[2]) != "object") return (aryValues[2]);
			}

			break;

		}


		case 1:																			// only got one, it's not really an
		{																				// expression, just extreanous parens
			if (aryValues[0].charAt(0) == '(')											// eg '((a) == b)'	
			{																			//      ^^^
				return (ETreeFrom(aryValues[0].substr(1,aryValues[0].length - 2), objReferenceArray));
			}

			aryValues[1] = '';
			aryValues[2] = '';
			break;

		}

		default:																		// sorry charlie, wrong number of tokens
		{																				// in that expression.
			return ("Parse Error in '" + strExpression + "' found " + aryValues.length + " tokens, expected 1 or 3.");
			break;
		}


	}
																						// all done, create and return 
																						// the new node
	return (new CExprTree(aryValues[0], aryValues[1], aryValues[2], strExpression, objReferenceArray));	


}

//////////////////////////////////////////////
//  Functions added to the ETree
//

////////////////////////////////////////////////////
// formats a string into a currency value.
// returns "" on failure.
//
function formatMoney (theStr)
{
	var rtnStr;
	var theNewStr = theStr;
	var strPrefix = "";		//"$";

	if (theNewStr.match(/\w*[.][0-9]$/))
		theNewStr = theNewStr + "0";

	if (theNewStr.match(/\w*[.]$/))
		theNewStr = theNewStr + "00";
		
	// Remove all non-numeric characters
	while (theNewStr.match(/\D+/))
	{
		theNewStr = theNewStr.replace(/\D+/, "");
	}
	
	// Remove leading zeros
	if ((theNewStr.length > 1) && (theNewStr.charAt(0) == "0"))
	{
		theNewStr = theNewStr.replace(/[0]+/, "");
		
		if (theNewStr == "") //if all gone, put one back
			theNewStr = "0";
	}
	
	if (theNewStr.length == 0)
	{
		rtnStr = "";
	}
	else if (theNewStr.length == 1)
	{
		rtnStr = strPrefix + "0.0" + theNewStr;
	}
	else if (theNewStr.length == 2)
	{
		rtnStr = strPrefix + "0." + theNewStr;
	}
	else
	{
		rtnStr = strPrefix + 
				 theNewStr.substr(0, theNewStr.length - 2) + "." + 
		         theNewStr.substr(theNewStr.length - 2, 2);
	}

	return rtnStr;

}

////////////////////////////////////////////////////
//
function EVAL_formatMoney(astrParams)
{

	return(formatMoney(astrParams[0]));

}

CExprTree.addFunction ("$formatMoney", EVAL_formatMoney, 1);

////////////////////////////////////////////////////
//
function EVAL_parseFloat (astrParams)
{
	var rval = "", input = astrParams[0];
	var hasDecimal = false;

	for (i = 0; i < input.length; i++)
	{
		c = input.charAt(i);
		if (!isNaN(c) || ((c == '.') && !hasDecimal))
		{
			if (c == '.')
				hasDecimal = true;
			rval += c;
		}
		else
			if (rval != "")
				return (rval);
	}

	if (rval.charAt(rval.length-1) == '.')
		rval += "00";

	return(rval);

}

CExprTree.addFunction ("$formatFloat", EVAL_parseFloat, 1);

////////////////////////////////////////////////////
//
function EVAL_parseInt (astrParams)
{
	var rval = "", input = astrParams[0];

	for (i = 0; i < input.length; i++)
	{
		c = input.charAt(i);
		if (!isNaN(c))
		{
			rval += c;
		}
		else
			if (rval != "")
				return (rval);
	}

	return(rval);

}

CExprTree.addFunction ("$formatInt", EVAL_parseInt, 1);

////////////////////////////////////////////////////
//
function EVAL_len (astrParams)
{

	alert (astrParams[0].toString());

	return(astrParams[0].toString().length);

}

CExprTree.addFunction ("$len", EVAL_len, 1);


////////////////////////////////////////////////////
// Input Format: HHMMAP where A = {a, A, p, P} and 
// P = {null, m, M} Ex: 11:32a or 0532PM or 1345
// Return 12hour format with AM or PM suffix
function formatTime(theStr)
{
	var rtnStr = "";
	var theNewStr = theStr;
		
	if (theNewStr.length > 4)
	{
		if (theNewStr.match(/[01][0-9][0-5][0-9][aA]|[01][0-9][0-5][0-9][pP]|[01][0-9][0-5][0-9][aA][mM]|[01][0-9][0-5][0-9][pP][mM]]/))
		{
			if ((parseInt(theNewStr.substr(0,2)) <= 12) && (theNewStr.substr(0,2) != "00"))
			{
				rtnStr = theNewStr.toUpperCase();
				var lastChar = rtnStr.substr(theNewStr.length-1,1);
				if (lastChar.match(/[AP]/))
					rtnStr = rtnStr + "M";
			}
		}
	}
	else if (theNewStr.length == 4)
	{
		if (theNewStr.match(/[0-2][0-9][0-5][0-9]/))
		{
			if (theNewStr.substr(0,2) != "00")
			{
				if (parseInt(theNewStr.substr(0,2)) < 12)
					rtnStr = theNewStr + "AM";
				else if (parseInt(theNewStr.substr(0,2)) == 12)
					rtnStr =  theNewStr + "PM";
				else if (parseInt(theNewStr.substr(0,2)) < 24)
				{
					var hr = (parseInt(theNewStr.substr(0,2)) - 12).toString();
				
					if (hr.length == 1)
						hr = "0" + hr;
					
					rtnStr = hr + theNewStr.substr(2,2) + "PM";
				}
			}
		}
	}
	return rtnStr;
}

////////////////////////////////////////////////////
//
function EVAL_formatTime(aryParams)
{

	return (formatTime(aryParams[0]));
}

CExprTree.addFunction ("$formatTime", EVAL_formatTime, 1);

////////////////////////////////////////////////////
//
function formatDate(strInDate, strSep)
{
	var strDate = strInDate;

	if (strSep == (void 0)) strSep = "";

	// Remove all non-numeric characters
	strDate.replace(/\D+/g, "")

	if (strDate.length != 8)
	{
		return "";
	}

	strDate = strDate.substr(0,2) + strSep + strDate.substr(2,2) + strSep + strDate.substr(4,4);

	if (!isDate(strDate))
		return "";

	return strDate;

}

////////////////////////////////////////////////////
//
function EVAL_formatDate(aryParams)
{

	return (formatDate(aryParams[0]));

}

CExprTree.addFunction ("$formatDate", EVAL_formatDate, 1);

////////////////////////////////////////////////////
// converts a date into its number of days past
// the epoch.
//
function toDays(strDate)
{

	var d = formatDate(strDate,"/");

	if (d == "") return (0);

	alert (d);
	
	var secd = Date.parse(d); return (parseInt(parseInt(secd) / 1000 / 60 / 60 / 24));

}

////////////////////////////////////////////////////
//
function EVAL_toDays(aryParams)
{

	return (toDays(aryParams[0]));

}

CExprTree.addFunction ("$toDays", EVAL_toDays, 1);

////////////////////////////////////////////////////
//
function getAge(strDate)
{
	if (strDate == "") return ("");

	var today = new Date;
	var start = new Date(formatDate(strDate, "/"));

	var diff = today.valueOf() - start.valueOf();
	var diff = parseInt (diff / Math.MSPerYear)

	return (diff);

}

////////////////////////////////////////////////////
//
function EVAL_getAge(aryParams)
{

	return (getAge(aryParams[0]));

}

CExprTree.addFunction ("$getAge", EVAL_getAge, 1);

////////////////////////////////////////////////////
// (Note, this function uses the isDate function)
// dateAdd usage: strDate ="01012000" or ="01/01/2000" 
// and strDays ="-10" or =-10 (negative or positive);
// returns "12221999"
//
function dateAdd(strDate, strDays)
{
	var strDateInput = strDate;
	var intDays = parseInt(strDays);1

	// Remove all non-numeric characters
	while (strDateInput.match(/\D+/))
	{
		strDateInput = strDateInput.replace(/\D+/, "");
	}

	if ((strDateInput.length != 8) || (isDate(strDateInput) == false) || ("NaN" == intDays))
	{
		return "";
	}

	var dateFmt = strDateInput.substr(0,2) + "/" + strDateInput.substr(2,2) + "/" + strDateInput.substr(4,4);
	
	var dateObj  = new Date(dateFmt);

	var datePart = dateObj.getDate();
	datePart = datePart + intDays;
	dateObj.setDate(datePart);

	var strNewMonth = (dateObj.getMonth() + 1).toString();
	if (strNewMonth.length == 1)
		strNewMonth = "0" + strNewMonth;

	var strNewDate = dateObj.getDate().toString();
	if (strNewDate.length == 1)
		strNewDate = "0" + strNewDate;

	var strResult = strNewMonth + strNewDate + (dateObj.getFullYear()).toString();

	return strResult;
}



////////////////////////////////////////////////////
//
function EVAL_dateAdd(aryParams)
{

	return (dateAdd(aryParams[0],aryParams[1]));

}

CExprTree.addFunction ("$dateAdd", EVAL_dateAdd, 2);


////////////////////////////////////////////////////
//
function isDate(strDate)
{

	/* rgg: VB functions cant be declared in a .js file
		so I wanted to be safe and make sure it was there
		before it was used but, any check I can think of to
		verify the existance of VBIsDate causes a
		'object doesnt support this p/m' error.  This
		sucks.
	*/

	strDate = strDate.replace (/\D+/g, "");

	if ((strDate.length != 8) || (Number(strDate.substr(0,2)) > 12))
		return false;

	strDate = strDate.substr(0,2) + "/" + strDate.substr(2,2) + "/" + strDate.substr(4,4);

//-	if (typeof VBIsDate != "undefined")

		return VBIsDate (strDate);

//-	else
//-		return JSIsDate (strDate);

}

function JSIsDate (theStr)
{
	var trnVal = false;
			
	if (theStr.length == 8)
	{
		if (theStr.match(/[01][0-9][0-3][0-9][0-9]{4}/))
		{				
			if ((theStr.substr(0,2) != "00") && 
				(theStr.substr(2,2) != "00") &&
				(theStr.substr(4,4) != "0000"))
			{
				var finalTest = new Date(theStr.substr(0,2) + "/" + theStr.substr(2,2) + "/" + theStr.substr(4,4));
				if (("NaN" == finalTest) || (null == finalTest) || (finalTest.getDate() < 1))
					trnVal = false;
				else
					trnVal = true;
			}
		}
	}

	return trnVal;

}
////////////////////////////////////////////////////
//
function EVAL_isDate(aryParams)
{

	return (isDate(aryParams[0]));

}

CExprTree.addFunction ("$isDate", EVAL_isDate, 1);

//
//
// End of helper functions
///////////////////////////////////////////////////////////
