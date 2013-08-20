<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=".\lib\CheckLogicExpression.inc"-->
<%
expr = """~x~"" = ""z1"" or (3= 6 and 4 = 6)"
bStatus = CheckLogicExpression(expr, "VBScript")

bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")
bStatus = CheckLogicExpression(expr, "VBScript")


Response.Expires = 0
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<body BGCOLOR="#d6cfbd">
Expression:  <%=expr%><BR>
<%If bStatus = True Then%>
Valid
<%Else%>
Not Valid
<%End If%>
</body>
</html>
