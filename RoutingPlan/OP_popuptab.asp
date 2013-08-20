<!--#include file="..\lib\common.inc"-->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

<title>Document Title</title>
</head>
<body>

<button ID="OpenWin">Open Window</button>

</body>
</html>
<script LANGUAGE="JavaScript">

function OpenWin.onclick()
{
	var strFeatures = "top=100, left=100, width=400, height=450, toolbar=no, menubar=no, location=no, directories=no";
	window.open("MyTabWindow.asp", "Properties", strFeatures);
}
</script>