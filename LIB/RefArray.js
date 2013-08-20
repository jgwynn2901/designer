///////////////////////////////////////////////////////////
// Associative array class.
// rgg 9/28/98
//

function CReferenceArray()
{
	this.isDirty = true;
}


///////////////////////////////////////////////////////////
// Adds an item named 'strItemName' to the array and 
// assigns it the value /strItemValue'
//
CReferenceArray.prototype.addItem = function (strItemName, strItemValue)
{

	//*alert ("RA Adding '" + strItemName + "' as " + strItemValue + "");
	this["ITEM_" + strItemName] = strItemValue;
	this.isDirty = true;

}

///////////////////////////////////////////////////////////
// Returns the value stored in 'strItemName'
// if the item isn't found, RE_getUnknownValue is called.
// the client can override this function to provide a
// function, such as, when an item named "[currentindex]"
// is referenced, a function is run, or something like that.
//
CReferenceArray.prototype.itemValue = function (strItemName)
{
	var value = this["ITEM_" + strItemName];

	if (value != (void 0))
		return (this["ITEM_" + strItemName]);
	else
		return (this.getUnknownValue(strItemName));

}

///////////////////////////////////////////////////////////
// If the item name starts with a "[" then it's value is 
// retrieved from the user, UNLESS it's '[blank]', in which
// case an empty string is returned.
//
CReferenceArray.prototype.getUnknownValue = function (strItemName)
{

	//*alert ("Cant find " + strItemName);
/*
	if (strItemName.search(/^\[\S+\]$/) != -1)
	{

		switch (strItemName)
		{
			
			case "[blank]":
			case "[BLANK]":
			case "[Blank]":
				return ("");

			default:
				return (prompt ("Enter a value for " + strItemName, ""));
		}

	}
*/
	return (void 0);

}

//
//
// End of assocatiave array class
///////////////////////////////////////////////////////////
