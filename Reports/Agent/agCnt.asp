<%
nCurrAgent = request.querystring("x")
nTotalAgents = request.querystring("top")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Thank you...</title>
</head>

<body bgColor="rosybrown">
<div STYLE="position: absolute; left: 250; top: 200; width: 562; height: 50">
<table border="0" width="598" >
  <tr>
    <td align="left" width="276"><font face="Comic Sans MS" size="5">Processing Agent</font></td>
    <td align="left" width="98"><font face="Comic Sans MS" size="5"><%=nCurrAgent%></font></td>
    <td align="left" width="72"><font face="Comic Sans MS" size="5">of</font></td>
    <td align="left" width="126"><font face="Comic Sans MS" size="5"><%=nTotalAgents%></font></td>
  </tr>
 </table>

</div>
</body>

</html>
