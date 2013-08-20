<script language=jscript>

function CityZip()
{
	this.City = "";
	this.State = "";
	this.Zip = "";
	this.County = "";
	this.FIPS = "";
	this.Country = "";
}

var oZip = new CityZip();

</script>

<script language=vbscript>
sub fillZip
dim cURL

oZip.Zip = ""
oZip.State = ""
oZip.City = ""
oZip.County = ""
oZip.FIPS = ""
oZip.Country = ""
cURL = "../LIB/ZIPLookupIFrame.asp?ZIP=" & document.all.ZIP.value
showModalDialog  cURL, oZip ,"center"

on error resume next
document.all.ZIP.value = oZip.Zip
document.all.STATE.value = oZip.State
document.all.CITY.value = oZip.City
document.all.COUNTY.value = oZip.County
document.all.FIPS.value = oZip.FIPS
document.all.COUNTRY.value = oZip.Country
on error goto 0
end sub

sub ZIP_onblur
dim x, oZip, z, lFocusSet

fillZIP
on error resume next
for x=0 to document.all.length-1
	if uCase(Trim(document.all(x).name)) = "ZIP" then
		if err.number = 0 then
			set oZip = document.all(x)
			exit for
		elseif err.number <> 438 then	'	no such property
			msgbox "Internal Error: " & err.number & " - " & err.description
			exit for
		else
			err.clear
		end if
	end if
next
lFocusSet = false
for z=x+1 to document.all.length-1
	document.all(z).focus
	if err.number = 0 then
		lFocusSet = true
		exit for
	else
		err.clear
	end if
next
on error goto 0
if lFocusSet then
	'msgbox "Focus switched to " & document.all(z).tagname
else
	' start from the beginning
	on error resume next
	for z=0 to document.all.length-1
		document.all(z).focus
		if uCase(Trim(document.all(z).name)) <> "ZIP" then
			if err.number = 0 then
				lFocusSet = true
				exit for
			else
				err.clear
			end if
		end if
	next
	on error goto 0
	if lFocusSet then
		'msgbox "Focus switched to " & document.all(z).tagname
	else
		msgbox "Couldn't find an element to switch focus to."
	end if	
end if
end sub

</script>