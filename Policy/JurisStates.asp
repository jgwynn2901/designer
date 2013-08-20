<!--#include file="..\lib\common.inc"-->
<html>
<head>
<script type="text/javascript">

var allStates = "*";
function setAllStatesValue(allStatesValue){
	allStates = allStatesValue;
}

function selectAll () {
      var state = document.getElementsByName('USstate'); 
      for (var i =0; i < state.length; i++) 
		{
		 state[i].checked = true;
		}	
	  state = document.getElementsByName('CANstate'); 
      for (var i =0; i < state.length; i++) 
		{
		 state[i].checked = true;
		}		
		document.getElementsByName('BtnSelectOrDeselectAll')[0].value = "Deselect All";
}     

function selectOrdeselectAll(){
	var val = document.getElementsByName('BtnSelectOrDeselectAll');
	if (val[0].value == "Select All"){
		selectAll();
	}
	else{
		deselectAll();
	}	
}

function deselectAll(){
      var state = document.getElementsByName('USstate'); 
      for (var i =0; i < state.length; i++) 
		{
		 state[i].checked = false;
		}		
	  state = document.getElementsByName('CANstate'); 
      for (var i =0; i < state.length; i++) 
		{
		 state[i].checked = false;
		}
		document.getElementsByName('BtnSelectOrDeselectAll')[0].value = "Select All";
} 

function getSelectedStates(){
		var selectedStates = new Array();	  			
		var USstates = document.getElementsByName('USstate'); 		  
		var count = 0;
		for (var i =0; i < USstates.length; i++) 
		{
			if(USstates[i].checked == true){
				selectedStates[count] = USstates[i].value;
				count++;			
			}		
		}		
		var CANstate = document.getElementsByName('CANstate'); 
		for (var i =0; i < CANstate.length; i++) 
		{
		if(CANstate[i].checked == true){
			selectedStates[count] = CANstate[i].value;
			count++;
			}
		}
		if(count == (USstates.length + CANstate.length)){
			selectedStates = new Array();
			selectedStates[0] = allStates; 
		}	
	return selectedStates;
}

function getAllStates(){	
		var States = new Array();	  			
		var USstates = document.getElementsByName('USstate'); 		  
		var j = 0;
		for (var i =0; i < USstates.length; i++){
			if(USstates[i].checked == true){
				States[j] = USstates[i].value;
				j++;						
			}		
		}		
		var CANstate = document.getElementsByName('CANstate'); 
		for (var i =0; i < CANstate.length; i++){
			if(CANstate[i].checked == true){
				States[j] = CANstate[i].value;
				j++;
			}
		}			
	return States;
}

function setSelectedStates(SelectedStates){
	if(SelectedStates[0] == allStates){	
		selectAll();
	}else{
		var US_States = document.getElementsByName('USstate'); 
		var CAN_States = document.getElementsByName('CANstate');
		for(var j = 0; j < SelectedStates.length; j++){					
			if(SelectedStates[j] == allStates){
				selectAll();
				return;
			}				
			for (var i =0; i < US_States.length; i++){				
				if(US_States[i].value == SelectedStates[j]){
					US_States[i].checked = true;					
					break;		 
				}
			}			   
			for (var i =0; i < CAN_States.length; i++){
				if(CAN_States[i].value == SelectedStates[j]){
					CAN_States[i].checked = true;
					break;
				}
			}
		}	
	}		 	 	
}
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true
<%	end if %>			
End Sub
sub SetScreenFieldsReadOnly(bDisabled)
	for iCount = 0 to document.all.length-1		
		if document.all(iCount).getAttribute("ScrnChkBox") = "TRUE" OR document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then	
			document.all(iCount).disabled = bDisabled
		end if
	next
end sub
</script>

<meta name="VI60_defaultClientScript" content="VBScript">
<title>Policy Jurisdiction State Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<BODY  topmargin=0 leftmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<form>
<table CLASS="LABEL" align="left" cellspacing="5">
<tr><TD colspan="8"><u>State/US Territory:</u></TD></tr>
<tr>
	<td>AK</td><td><INPUT type="checkbox" name="USstate" value="AK" ScrnChkBox="TRUE"></td>
	<td>AL</td><td><INPUT type="checkbox" name="USstate" value="AL" ScrnChkBox="TRUE"></td>
	<td>AR</td><td><INPUT type="checkbox" name="USstate" value="AR" ScrnChkBox="TRUE"></td>
	<td>AZ</td><td><INPUT type="checkbox" name="USstate" value="AZ" ScrnChkBox="TRUE"></td>
	<td>CA</td><td><INPUT type="checkbox" name="USstate" value="CA" ScrnChkBox="TRUE"></td>
	<td>CO</td><td><INPUT type="checkbox" name="USstate" value="CO" ScrnChkBox="TRUE"></td>
	<td>CT</td><td><INPUT type="checkbox" name="USstate" value="CT" ScrnChkBox="TRUE"></td>
	<td>DC</td><td><INPUT type="checkbox" name="USstate" value="DC" ScrnChkBox="TRUE"></td>	
</tr>
<tr>
	<td>DE</td><td><INPUT type="checkbox" name="USstate" value="DE" ScrnChkBox="TRUE"></td>
	<td>FL</td><td><INPUT type="checkbox" name="USstate" value="FL" ScrnChkBox="TRUE"></td>
	<td>GA</td><td><INPUT type="checkbox" name="USstate" value="GA" ScrnChkBox="TRUE"></td>
	<td>GU</td><td><INPUT type="checkbox" name="USstate" value="GU" ScrnChkBox="TRUE"></td>
	<td>HI</td><td><INPUT type="checkbox" name="USstate" value="HI" ScrnChkBox="TRUE"></td>
	<td>IA</td><td><INPUT type="checkbox" name="USstate" value="IA" ScrnChkBox="TRUE"></td>
	<td>ID</td><td><INPUT type="checkbox" name="USstate" value="ID" ScrnChkBox="TRUE"></td>
	<td>IL</td><td><INPUT type="checkbox" name="USstate" value="IL" ScrnChkBox="TRUE"></td>				
</tr>
<tr>
	<td>IN</td><td><INPUT type="checkbox" name="USstate" value="IN" ScrnChkBox="TRUE"></td>
	<td>KS</td><td><INPUT type="checkbox" name="USstate" value="KS" ScrnChkBox="TRUE"></td>
	<td>KY</td><td><INPUT type="checkbox" name="USstate" value="KY" ScrnChkBox="TRUE"></td> 
	<td>LA</td><td><INPUT type="checkbox" name="USstate" value="LA" ScrnChkBox="TRUE"></td>
	<td>MA</td><td><INPUT type="checkbox" name="USstate" value="MA" ScrnChkBox="TRUE"></td>
	<td>MD</td><td><INPUT type="checkbox" name="USstate" value="MD" ScrnChkBox="TRUE"></td>
	<td>ME</td><td><INPUT type="checkbox" name="USstate" value="ME" ScrnChkBox="TRUE"></td>
	<td>MI</td><td><INPUT type="checkbox" name="USstate" value="MI" ScrnChkBox="TRUE"></td>	
</tr>
<tr>
	<td>MN</td><td><INPUT type="checkbox" name="USstate" value="MN" ScrnChkBox="TRUE"></td>
	<td>MO</td><td><INPUT type="checkbox" name="USstate" value="MO" ScrnChkBox="TRUE"></td>
	<td>MS</td><td><INPUT type="checkbox" name="USstate" value="MS" ScrnChkBox="TRUE"></td>
	<td>MT</td><td><INPUT type="checkbox" name="USstate" value="MT" ScrnChkBox="TRUE"></td>
	<td>NC</td><td><INPUT type="checkbox" name="USstate" value="NC" ScrnChkBox="TRUE"></td>
	<td>ND</td><td><INPUT type="checkbox" name="USstate" value="ND" ScrnChkBox="TRUE"></td>
	<td>NE</td><td><INPUT type="checkbox" name="USstate" value="NE" ScrnChkBox="TRUE"></td>
	<td>NH</td><td><INPUT type="checkbox" name="USstate" value="NH" ScrnChkBox="TRUE"></td>	
</tr>
<tr>
	<td>NJ</td><td><INPUT type="checkbox" name="USstate" value="NJ" ScrnChkBox="TRUE"></td>
	<td>NM</td><td><INPUT type="checkbox" name="USstate" value="NM" ScrnChkBox="TRUE"></td>
	<td>NV</td><td><INPUT type="checkbox" name="USstate" value="NV" ScrnChkBox="TRUE"></td>
	<td>NY</td><td><INPUT type="checkbox" name="USstate" value="NY" ScrnChkBox="TRUE"></td>
	<td>OH</td><td><INPUT type="checkbox" name="USstate" value="OH" ScrnChkBox="TRUE"></td>
	<td>OK</td><td><INPUT type="checkbox" name="USstate" value="OK" ScrnChkBox="TRUE"></td>
	<td>OR</td><td><INPUT type="checkbox" name="USstate" value="OR" ScrnChkBox="TRUE"></td>
	<td>PA</td><td><INPUT type="checkbox" name="USstate" value="PA" ScrnChkBox="TRUE"></td>	
</tr>
<tr>
	<td>PR</td><td><INPUT type="checkbox" name="USstate" value="PR" ScrnChkBox="TRUE"></td>
	<td>RI</td><td><INPUT type="checkbox" name="USstate" value="RI" ScrnChkBox="TRUE"></td>
	<td>SC</td><td><INPUT type="checkbox" name="USstate" value="SC" ScrnChkBox="TRUE"></td>
	<td>SD</td><td><INPUT type="checkbox" name="USstate" value="SD" ScrnChkBox="TRUE"></td>
	<td>TN</td><td><INPUT type="checkbox" name="USstate" value="TN" ScrnChkBox="TRUE"></td>
	<td>TX</td><td><INPUT type="checkbox" name="USstate" value="TX" ScrnChkBox="TRUE"></td>
	<td>UT</td><td><INPUT type="checkbox" name="USstate" value="UT" ScrnChkBox="TRUE"></td>
	<td>VA</td><td><INPUT type="checkbox" name="USstate" value="VA" ScrnChkBox="TRUE"></td>	
</tr>
<tr>
	<td>VI</td><td><INPUT type="checkbox" name="USstate" value="VI" ScrnChkBox="TRUE"></td>
	<td>VT</td><td><INPUT type="checkbox" name="USstate" value="VT" ScrnChkBox="TRUE"></td>
	<td>WA</td><td><INPUT type="checkbox" name="USstate" value="WA" ScrnChkBox="TRUE"></td>
	<td>WI</td><td><INPUT type="checkbox" name="USstate" value="WI" ScrnChkBox="TRUE"></td>
	<td>WV</td><td><INPUT type="checkbox" name="USstate" value="WV" ScrnChkBox="TRUE"></td>
	<td>WY</td><td><INPUT type="checkbox" name="USstate" value="WY" ScrnChkBox="TRUE"></td>	
</tr>
<tr>
	<td>AA</td><td><INPUT type="checkbox" name="USstate" value="AA" ScrnChkBox="TRUE"></td>
	<td>AE</td><td><INPUT type="checkbox" name="USstate" value="AE" ScrnChkBox="TRUE"></td>
	<td>AP</td><td><INPUT type="checkbox" name="USstate" value="AP" ScrnChkBox="TRUE"></td>
	<td>AS</td><td><INPUT type="checkbox" name="USstate" value="AS" ScrnChkBox="TRUE"></td>
	<td>FM</td><td><INPUT type="checkbox" name="USstate" value="FM" ScrnChkBox="TRUE"></td>
	<td>MH</td><td><INPUT type="checkbox" name="USstate" value="MH" ScrnChkBox="TRUE"></td>	
	<td>MP</td><td><INPUT type="checkbox" name="USstate" value="MP" ScrnChkBox="TRUE"></td>	
	<td>PW</td><td><INPUT type="checkbox" name="USstate" value="PW" ScrnChkBox="TRUE"></td>	
</tr><tr/>
<tr><TD colspan="8"><u>Canadian Provinces:</u></TD></tr>
<tr>
	<td>AB</td><td><INPUT type="checkbox" name="CANstate" value="AB" ScrnChkBox="TRUE"></td>
	<td>BC</td><td><INPUT type="checkbox" name="CANstate" value="BC" ScrnChkBox="TRUE"></td>
	<td>MB</td><td><INPUT type="checkbox" name="CANstate" value="MB" ScrnChkBox="TRUE"></td>
	<td>NB</td><td><INPUT type="checkbox" name="CANstate" value="NB" ScrnChkBox="TRUE"></td>
	<td>NL</td><td><INPUT type="checkbox" name="CANstate" value="NL" ScrnChkBox="TRUE"></td>
	<td>NS</td><td><INPUT type="checkbox" name="CANstate" value="NS" ScrnChkBox="TRUE"></td>	
	<td>NT</td><td><INPUT type="checkbox" name="CANstate" value="NT" ScrnChkBox="TRUE"></td>	
	<td>NU</td><td><INPUT type="checkbox" name="CANstate" value="NU" ScrnChkBox="TRUE"></td>	
</tr>
<tr>
	<td>ON</td><td><INPUT type="checkbox" name="CANstate" value="ON" ScrnChkBox="TRUE"></td>
	<td>PE</td><td><INPUT type="checkbox" name="CANstate" value="PE" ScrnChkBox="TRUE"></td>
	<td>QC</td><td><INPUT type="checkbox" name="CANstate" value="QC" ScrnChkBox="TRUE"></td>
	<td>SK</td><td><INPUT type="checkbox" name="CANstate" value="SK" ScrnChkBox="TRUE"></td>
	<td>YT</td><td><INPUT type="checkbox" name="CANstate" value="YT" ScrnChkBox="TRUE"></td>
	<td>CN</td><td><INPUT type="checkbox" name="CANstate" value="CN" ScrnChkBox="TRUE"></td>	
</tr><tr/><tr/>
<tr><td colspan="14">
<table CLASS="LABEL" align="left">
	<tr><TD>Select/Deselect individual States/Provinces by clicking the appropriate box.</TD>
	<td><input CLASS=StdButton type="button" value="Select All" NAME="BtnSelectOrDeselectAll" ScrnBtn="TRUE" onclick="selectOrdeselectAll()" ID="Button1"/></td></tr>
	<tr></tr><tr/><tr/>
</table></td></tr>
</table>
</form>
</body>
</html>