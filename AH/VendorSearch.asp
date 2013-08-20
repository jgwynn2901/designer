
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vendor Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="#d6cfbd">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Vendor Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<form Name="FrmSearch" METHOD="GET" ACTION="CarrierSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="RW">
<input type="hidden" NAME="CID" value="">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align="left">
	<tr>
	<tr>
	<tr>
		<td CLASS="LABEL">Vendor ID:<br><input CLASS="LABEL" tabindex="1" TYPE="TEXT" size="10" MAXLENGTH="10" NAME="SearchCarrierNumber" VALUE=""></td>
		<td CLASS="LABEL" COLSPAN="2">Name:<br><input CLASS="LABEL" tabindex="2" TYPE="TEXT" NAME="SearchName" size="30" VALUE=""></td>
	</tr>
	<tr>
		<td CLASS="LABEL">Service Type:<br>
			<select NAME="SearchState" CLASS="LABEL" tabindex="8">
				<option VALUE>
				<OPTION VALUE='Body'>Body
			</select>
		</td>
	</tr>
	<tr>
		<td CLASS="LABEL" COLSPAN="2">City:<br><input tabindex="7" CLASS="LABEL" TYPE="TEXT" NAME="SearchCity" VALUE=""></td>
		<td CLASS="LABEL">State:<br>
			<select NAME="SearchState" CLASS="LABEL" tabindex="8">
				<option VALUE>
				<OPTION VALUE='AA'>AA<OPTION VALUE='AE'>AE<OPTION VALUE='AK'>AK<OPTION VALUE='AL'>AL<OPTION VALUE='AP'>AP<OPTION VALUE='AR'>AR<OPTION VALUE='AS'>AS<OPTION VALUE='AZ'>AZ<OPTION VALUE='CA'>CA<OPTION VALUE='CO'>CO<OPTION VALUE='CT'>CT<OPTION VALUE='DC'>DC<OPTION VALUE='DE'>DE<OPTION VALUE='FL'>FL<OPTION VALUE='FM'>FM<OPTION VALUE='GA'>GA<OPTION VALUE='GU'>GU<OPTION VALUE='HI'>HI<OPTION VALUE='IA'>IA<OPTION VALUE='ID'>ID<OPTION VALUE='IL'>IL<OPTION VALUE='IN'>IN<OPTION VALUE='KS'>KS<OPTION VALUE='KY'>KY<OPTION VALUE='LA'>LA<OPTION VALUE='MA'>MA<OPTION VALUE='MD'>MD<OPTION VALUE='ME'>ME<OPTION VALUE='MH'>MH<OPTION VALUE='MI'>MI<OPTION VALUE='MN'>MN<OPTION VALUE='MO'>MO<OPTION VALUE='MP'>MP<OPTION VALUE='MS'>MS<OPTION VALUE='MT'>MT<OPTION VALUE='NC'>NC<OPTION VALUE='ND'>ND<OPTION VALUE='NE'>NE<OPTION VALUE='NH'>NH<OPTION VALUE='NJ'>NJ<OPTION VALUE='NM'>NM<OPTION VALUE='NV'>NV<OPTION VALUE='NY'>NY<OPTION VALUE='OH'>OH<OPTION VALUE='OK'>OK<OPTION VALUE='OR'>OR<OPTION VALUE='PA'>PA<OPTION VALUE='PR'>PR<OPTION VALUE='PW'>PW<OPTION VALUE='RI'>RI<OPTION VALUE='SC'>SC<OPTION VALUE='SD'>SD<OPTION VALUE='TN'>TN<OPTION VALUE='TX'>TX<OPTION VALUE='UT'>UT<OPTION VALUE='VA'>VA<OPTION VALUE='VI'>VI<OPTION VALUE='VT'>VT<OPTION VALUE='WA'>WA<OPTION VALUE='WI'>WI<OPTION VALUE='WV'>WV<OPTION VALUE='WY'>WY 
			</select>
		</td>
		<td CLASS="LABEL">Zip:<br><input tabindex="9" SIZE="7" CLASS="LABEL" TYPE="TEXT" NAME="SearchZip" VALUE=""></td>
	</tr>
	</table>
</td>
<td VALIGN="TOP" rowspan="3">
	<table>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="12" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="13" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
	</table>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="9" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="10" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="11" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>


</form>
</body>
</html>
