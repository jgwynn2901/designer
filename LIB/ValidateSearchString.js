function f_EncodeURLString(in_SearchString){
	var s_OutPutString = new String();
	s_OutPutString = escape(in_SearchString);
	return s_OutPutString;
}

function f_NumericCheck(in_NumberString){
	var s_Number = "0123456789";
	if(in_NumberString.length == 0) {
		return "0";
	}
	else{
		for(var i_Loop = 0; i_Loop < in_NumberString.length; i_Loop++){
			if (-1 == s_Number.indexOf(in_NumberString.charAt(i_Loop))){
				return "1";}
		}
	}
	return "0";
}