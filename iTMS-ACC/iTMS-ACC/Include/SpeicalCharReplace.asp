<%
function SpeicalCharReplace(val)
	Dim sTemp
	'sTemp = UCase( val )
	sTemp = trim(val)
	sTemp = Replace(sTemp,"&","&amp;")
	sTemp = Replace(sTemp,">","&gt;")
	sTemp = Replace(sTemp,"<","&lt;")
	sTemp = Replace(sTemp,"'","&apos;")
	sTemp = Replace(sTemp,"""","&quot;")
	SpeicalCharReplace = sTemp
end function
%>