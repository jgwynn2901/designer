<HTML>
	<HEAD>
		<title>Finance Report</title>
		<%
const cServerName = "//cha0s50t"

dim cRepDate, cErrMsg, cEmailAddress, cScheduleDate

if len(Request.form("emailAddress")) <> 0 then
	cRepDate = Request.form("selMonth") & Request.form("selYear")
	cEmailAddress = server.URLEncode(Request.form("emailAddress"))
	cScheduleDate = server.URLEncode(Request.form("scheduleData"))
	response.Redirect cServerName & "/scheduleRep/default.aspx?rd=" & cRepDate & "&ea=" & cEmailAddress & "&sd=" & cScheduleDate
end if
%>
		<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
		<meta name="ProgId" content="FrontPage.Editor.Document">
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<script language="vbscript">
Sub window_onload
<%
dim dDate, cDay, cMonth, cYear

dDate = date
cDay = cstr(Day(dDate))
if len(cDay ) < 2 then
	cDay = "0" & cDay 
end if	
cMonth = cstr(Month(dDate))
if len(cMonth) < 2 then
	cMonth = "0" & cMonth
end if	
cYear = cstr(Year(dDate))
%>
document.all.selMonth.selectedIndex = Month(Date) - 1
document.all.selYear.selectedIndex = 1
document.all.scheduleDate.value = ""
document.all.scheduleDate.DateMin = "<%=cYear%>" & "<%=cMonth%>" & "<%=cDay%>"
document.all.scheduleDate.TimeValue = "22:00:00"
End Sub

Sub cmdRun_onclick
dim cErrMsg

cErrMsg = ""
if document.all.scheduleDate.value = "0" then
	cErrMsg = "Please enter a date to schedule the report to run." & vbcrlf
end if
if len(document.all.emailAddress.value) = 0 then
	cErrMsg = cErrMsg  & "Please enter the Email address where the report will be send to."
end if
if len(cErrMsg) <> 0 then
	msgbox cErrMsg, 48, "Finance Billing Report"
else
	if msgbox("Please confirm:" & vbcrlf & vbcrlf & "Report will run on: " & formatdatetime(document.all.scheduleDate.value) & vbcrlf & _
		"Email address: " & document.all.emailAddress.value,1,"Finance Billing Report") = 1 then
		document.all.scheduleData.value = document.all.scheduleDate.value
		frmSchedule.submit
	end if
end if
End Sub

		</script>
	</HEAD>
	<body bgcolor="seashell">
		<form Name="frmSchedule" METHOD="post" ACTION="finRep.asp" ID="Form1">
			<input type="hidden" NAME="scheduleData" value='ID="Hidden1"'>
			<div align="left">
				<table border="0" width="66%">
					<tr>
						<td CLASS="GrpLabel" WIDTH="70" HEIGHT="12"><font face="Verdana, Helvetica, Arial"><nobr>&nbsp;» 
								Finance Billing Report</font></NOBR></td>
					</tr>
				</table>
			</div>
			<div align="left">
				<br>
				<table border="0" width="80%">
					<tr>
						<td class="Label" style="font-family: Verdana; font-size: 11pt">* Select the 
							period, month and year,&nbsp; for the report:<br>
						</td>
					</tr>
				</table>
				<br>
				<div STYLE="LEFT: 30px; POSITION: relative">
					<table border="0" width="80%" cellspacing="1" style="BORDER-COLLAPSE: collapse" bordercolor="#111111" cellpadding="2">
						<tr>
							<td class="Label" width="6%" style="font-family: Verdana; font-size: 11pt"><b>Period</b>:<br>
							</td>
							<td width="8%">
								<select name="selMonth" size="1" class="label">
									<option value="Jan" selected>Jan</option>
									<option value="Feb">Feb</option>
									<option value="Mar">Mar</option>
									<option value="Apr">Apr</option>
									<option value="May">May</option>
									<option value="Jun">Jun</option>
									<option value="Jul">Jul</option>
									<option value="Aug">Aug</option>
									<option value="Sep">Sep</option>
									<option value="Oct">Oct</option>
									<option value="Nov">Nov</option>
									<option value="Dec">Dec</option>
								</select>
							</td>
							<td width="11%">
								<select name="selYear" size="1" class="label">
									<option value="<%=Year(Date) - 1%>" selected><%=Year(Date) - 1%></option>
									<option value="<%=Year(Date)%>"><%=Year(Date)%></option>
									<option value="<%=Year(Date) + 1%>"><%=Year(Date) + 1%></option>
								</select>
							</td>
							<td width="29%" align="left">&nbsp;&nbsp;
							</td>
						</tr>
					</table>
				</div>
				<br>
				<table border="0" width="80%">
					<tr>
						<td class="Label" style="font-family: Verdana; font-size: 11pt">* Enter the date 
							when the report will execute (It will be scheduled to run overnight):</td>
					</tr>
				</table>
				<br>
				<div STYLE="LEFT: 30px; POSITION: relative">
					<table border="0" height="36" ID="Table1">
						<tr>
							<td class="Label" style="font-size: 11pt; font-family: Verdana" height="32"><b>This 
									report will run on</b>:
							</td>
							<td style="font-family: Verdana; font-size: 10pt" height="32">
								<OBJECT CLASSID="clsid:5220CB21-C88D-11CF-B347-00AA00A28331" VIEWASTEXT ID="Object1">
									<PARAM NAME="LPKPath" VALUE="../../bin/controls.lpk">
								</OBJECT>
								<object CLASSID="clsid:B8958DE0-BAC9-101C-933E-0000C005958C" CODEBASE="../../bin/Edt32x20.cab#version=2,1,0,10" id="scheduleDate" width="114" height="23" VIEWASTEXT>
									<param name="_Version" value="131073">
									<param name="_ExtentX" value="2646">
									<param name="_ExtentY" value="741">
									<param name="_StockProps" value="68">
									<param name="Enabled" value="1">
									<param name="BackColor" value="-2147483643">
									<param name="ForeColor" value="-2147483640">
									<param name="ThreeDInsideStyle" value="0">
									<param name="ThreeDInsideHighlightColor" value="-2147483633">
									<param name="ThreeDInsideShadowColor" value="-2147483642">
									<param name="ThreeDInsideWidth" value="1">
									<param name="ThreeDOutsideStyle" value="0">
									<param name="ThreeDOutsideHighlightColor" value="16777215">
									<param name="ThreeDOutsideShadowColor" value="-2147483632">
									<param name="ThreeDOutsideWidth" value="1">
									<param name="ThreeDFrameWidth" value="0">
									<param name="BorderStyle" value="1">
									<param name="BorderColor" value="-2147483642">
									<param name="BorderWidth" value="1">
									<param name="ButtonDefaultAction" value="-1">
									<param name="ButtonDisable" value="0">
									<param name="ButtonHide" value="0">
									<param name="ButtonIncrement" value="1">
									<param name="ButtonMin" value="0">
									<param name="ButtonMax" value="100">
									<param name="ButtonStyle" value="3">
									<param name="ButtonWidth" value="0">
									<param name="ButtonWrap" value="-1">
									<param name="ThreeDText" value="0">
									<param name="ThreeDTextHighlightColor" value="-2147483633">
									<param name="ThreeDTextShadowColor" value="-2147483632">
									<param name="ThreeDTextOffset" value="1">
									<param name="AlignTextH" value="0">
									<param name="AlignTextV" value="0">
									<param name="AllowNull" value="0">
									<param name="NoSpecialKeys" value="0">
									<param name="AutoAdvance" value="0">
									<param name="AutoBeep" value="0">
									<param name="CaretInsert" value="0">
									<param name="CaretOverWrite" value="3">
									<param name="UserEntry" value="0">
									<param name="HideSelection" value="-1">
									<param name="InvalidColor" value="-2147483637">
									<param name="InvalidOption" value="0">
									<param name="MarginLeft" value="3">
									<param name="MarginTop" value="3">
									<param name="MarginRight" value="3">
									<param name="MarginBottom" value="3">
									<param name="NullColor" value="-2147483637">
									<param name="OnFocusAlignH" value="0">
									<param name="OnFocusAlignV" value="0">
									<param name="OnFocusNoSelect" value="0">
									<param name="OnFocusPosition" value="0">
									<param name="ControlType" value="0">
									<param name="Text" value="fpDateTime1">
									<param name="DateCalcMethod" value="0">
									<param name="DateTimeFormat" value="0">
									<param name="UserDefinedFormat" value>
									<param name="DateMax" value="00000000">
									<param name="DateMin" value="00000000">
									<param name="TimeMax" value="000000">
									<param name="TimeMin" value="000000">
									<param name="TimeString1159" value>
									<param name="TimeString2359" value>
									<param name="DateDefault" value="00000000">
									<param name="TimeDefault" value="000000">
									<param name="TimeStyle" value="0">
									<param name="BorderGrayAreaColor" value="-2147483637">
									<param name="ThreeDOnFocusInvert" value="0">
									<param name="ThreeDFrameColor" value="-2147483633">
									<param name="Appearance" value="0">
									<param name="BorderDropShadow" value="0">
									<param name="BorderDropShadowColor" value="-2147483632">
									<param name="BorderDropShadowWidth" value="3">
									<param name="PopUpType" value="1">
									<param name="DateCalcY2KSplit" value="60">
									<param name="MousePointer" value="0">
								</object>
							</td>
						</tr>
					</table>
				</div>
				<br>
				<br>
				<table border="0" width="80%">
					<tr>
						<td class="Label" style="font-family: Verdana; font-size: 11pt">* Enter the Email 
							address, where the report will be emailed to:<br>
						</td>
					</tr>
				</table>
				<br>
				<div STYLE="LEFT: 30px; POSITION: relative">
					<table border="0" width="718">
						<tr>
							<td class="Label" style="font-family: Verdana; font-size: 11pt" width="521">
								<input type="text" name="emailAddress" size="30"></td>
							<td class="Label" width="362">
								<input id="cmdRun" name="cmdRun" CLASS="StdButton" type="button" value="Schedule" width="100"></td>
						</tr>
					</table>
				</div>
		</form>
		</DIV>
	</body>
</HTML>
