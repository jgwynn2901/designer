<!--#include file="..\lib\common.inc"-->
<% If Request.QueryString("ACTION") = "SAVE" Then 
		
		Set X = Server.CreateObject("IDGen.IDGen.1")
		X.TableName = "RULES"
		Set Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = "DRIVER={Microsoft ODBC for Oracle};SERVER=190.15.5.4;ConnectString=FNS;UID=FNSOWNER;PWD=CTOWN"
		Conn.Open ConnectionString
		RULEID = X.Next
		RULETEXT = "'" & Request.Form("EnblRule") & "'"
		RULETYPE = "'ROUTING'"
		SQLST = "INSERT INTO RULES (RULE_ID, RULE_TEXT, TYPE) VALUES( " & RULEID & " ,  " & RULETEXT & " , " & RULETYPE & ")"
		Set RS = Conn.Execute(SQLST)
		SUCCESS = True
End If
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub BtnFullSearch_onclick
	Window.open  "AttribSearch.asp?PAGE=RULE", "ATTRIBSEARCH", "height=400,width=350,status=no,toolbar=no,menubar=no,location=no"
	'call Window.showHelp("CFAttribSearch.asp","AttributeSearch","POPUP") 
End Sub

Sub BtnEnterAttrib_onclick
	If ATTRIBUTENAME.Value <> "" Then
		document.all.EnblRule.value =document.all.EnblRule.value & """~" & ATTRIBUTENAME.Value & "~"""
	End If
End Sub

Sub BtnEnterOperator_onclick
	If OPERATOR.value <> "" Then
		document.all.EnblRule.Value = document.all.EnblRule.value & " " & OPERATOR.value & " " 
	End If
End Sub

Sub BtnClear_onclick
	TxtBuffer.value = ""
	document.all.EnblRule.Value = ""
	RULELABEL.style.display = "None"
	
End Sub

Sub BtnEnterValue_onclick
	If TxtValue.Value <> "" Then
		document.all.EnblRule.value = document.all.EnblRule.value & """" & TxtValue.Value & """" 
	End If
End Sub

Sub BtnToBuffer_onclick
	TxtBuffer.Value = ""
	TxtBuffer.Value = document.all.EnblRule.value
	document.all.EnblRule.value = ""
	
End Sub

Sub BtnEnterBuffer_onclick
	document.all.EnblRule.value = document.all.EnblRule.value & TxtBuffer.Value
End Sub

Sub BtnSave_onclick
	If document.all.EnblRule.value <> "" Then
		document.all.RULEDATA.Submit()
	Else
		MsgBox "You must enter a rule before saving",  0, "FNSNet Designer"
	End If
End Sub

Sub BtnSelect_onclick
	Window.open  "RuleSearch.asp", "ATTRIBSEARCH", "height=400,width=350,status=no,toolbar=no,menubar=no,location=no"
	'call Window.showHelp("CFAttribSearch.asp","AttributeSearch","POPUP") 
End Sub

Sub BtnVerify_onclick
	lret = window.showModalDialog ("SyntaxCheck.asp?RULE=" & document.all.EnblRule.Value , Null,  "dialogWidth=400px; dialogHeight=400px; center=yes")
End Sub

-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%=BODYBGCOLOR%>'  topmargin=0 leftmargin=0>

<FIELDSET STYLE="WIDTH:300">
<LEGEND CLASS=LABEL>Attribute</LEGEND>
<TABLE>
<TR>
<TD CLASS=LABEL>
<INPUT TYPE=TEXT SIZE=40 CLASS=LABELREADONLY READONLY NAME="ATTRIBUTENAME">
</TD>
<TD CLASS=LABEL>
</TD>
<TD CLASS=LABEL>
<BUTTON CLASS=STDBUTTON NAME="BtnFullSearch">Full Search</BUTTON>
</TD>
<TD CLASS=LABEL>
<BUTTON CLASS=STDBUTTON STYLE="WIDTH:50" NAME="BtnEnterAttrib">Enter</BUTTON>
</TD>
</TR>
</TABLE>
</FIELDSET>
<TABLE>
<TR>
<TD>
<FIELDSET WIDTH=200>
<LEGEND CLASS=LABEL>Operator</LEGEND>
<TABLE>
<TR>
<TD CLASS=LABEL>
<SELECT NAME="OPERATOR" CLASS=LABEL>
<OPTION VALUE="AND">AND
<OPTION VALUE="OR">OR
<OPTION VALUE="=">EQUAL
<OPTION VALUE="<">LESS THAN
<OPTION VALUE="<=">LESS THAN OR EQUAL TO
<OPTION VALUE=">">GREATER THAN
<OPTION VALUE=">=">GREATER THAN OR EQUAL TO
</SELECT>
</TD>
<TD CLASS=LABEL>
<BUTTON CLASS=STDBUTTON STYLE="WIDTH:50" NAME="BtnEnterOperator">Enter</BUTTON>
</TD></TABLE>
</FIELDSET>
</TD><TD>
<FIELDSET>
<LEGEND CLASS=LABEL>Value</LEGEND>
<TABLE><TR>
<TD CLASS=LABEL><INPUT TYPE=TEXT SIZE=53 NAME="TxtValue" CLASS=LABEL></TD>
<TD CLASS=LABEL>
<BUTTON CLASS=STDBUTTON STYLE="WIDTH:50" NAME="BtnEnterValue">Enter</BUTTON>
</TD></TR>
</TABLE>
</FIELDSET>
</TD></TR>
</TABLE>
<FIELDSET STYLE="WIDTH:300">
<LEGEND CLASS=LABEL>Expression Buffer</LEGEND>
<TABLE>
<TR>
<TD CLASS=LABEL><INPUT TYPE=TEXT SIZE=105 NAME="TxtBuffer" READONLY CLASS=LABEL></TD>
<TD CLASS=LABEL><BUTTON CLASS=STDBUTTON STYLE="WIDTH:50" Name="BtnEnterBuffer">Enter</BUTTON>
</TR>
</TABLE>
</FIELDSET>
<BR>


<FORM NAME="RULEDATA" METHOD="POST" ACTION="OP_RULES.ASP?ACTION=SAVE">
<LABEL CLASS=LABEL>Routing Rule:</LABEL><BR>
<TEXTAREA NAME="EnblRule" CLASS=LABEL STYLE="WIDTH:600;HEIGHT:150">
<%= Request.Form("EnblRule") %></TEXTAREA><BR><BR>
</FORM>
<DIV ID="RULELABEL">
<% If SUCCESS=True Then %>
<LABEL CLASS=LABEL NAME="RULELABEL">&nbsp;&nbsp;&nbsp;Rule Saved Successfuly</LABEL><BR>
<% End If %>
</DIV>
<TABLE>
<TR>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnToBuffer">To Buffer</BUTTON></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnClear">Clear</BUTTON></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnSave">Save</BUTTON></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnSelect">Select</BUTTON></TD>
<TD><BUTTON CLASS=STDBUTTON NAME="BtnVerify">Verify</BUTTON></TD>
</TR>
</TABLE>
</BODY>
</HTML>
