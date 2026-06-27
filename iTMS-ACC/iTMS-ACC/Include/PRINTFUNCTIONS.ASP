<%
function formattprint(command,str)

select case command
	case "WIDE"			str=Chr(27) & "W1" & str & Chr(27) + "W0"
	case "ENHANCE"		str=Chr(27) & "E" & str & Chr(27) + "F"

	case "ITALIC"		str=Chr(27) & "4" & str & Chr(27) + "5"
	case "UNDERLINE"	str=Chr(27) & Chr(45) & "1" & str & Chr(27) + Chr(45) & "0"
	case "BOLD"			str=Chr(27) & Chr(71) & str & Chr(27) & Chr(72)

	case "RESET"		str=Chr(27) & Chr(64)

	case "PAGESKIP"		str=str & Chr(12)

	' Page Length in Inches
	case "PAGELEN4IN"	str=Chr(27) & Chr(67) & Chr(0) & Chr(4)
	case "PAGELEN6IN"	str=Chr(27) & Chr(67) & Chr(0) & Chr(6)
	case "PAGELEN12IN"	str=Chr(27) & Chr(67) & Chr(0) & Chr(12)

	' Page Length in Lines
	case "PAGELEN4LN"	str=Chr(27) & Chr(67) & Chr(24)
	case "PAGELEN6LN"	str=Chr(27) & Chr(67) & Chr(36)
	case "PAGELEN12LN"	str=Chr(27) & Chr(67) & Chr(72)

	case "BLANKLINE"	str= " " & vbCrLf

	case "12PITCH"			str=Chr(27) & Chr(77) & str
	case "10PITCH"			str=Chr(27) & Chr(80) & str
	case "L12PITCH"			str=Chr(27) & Chr(77) & str & Chr(27) & Chr(80)
	case "L10PITCH"			str=Chr(27) & Chr(80) & str & Chr(27) & Chr(77)
	case "CONDENSE"			str=Chr(27) & Chr(15) & str & Chr(27) & Chr(18)
	Case "CONDENSESTART"	str=Chr(27) & Chr(15)
	Case "CONDENSEEND"  	str=Chr(27) & Chr(18)

	case else			str=str & vbCrLf
end select

formattprint = str

end function
%>
<%
'------------------------Declaration Of User Define Function----------------
function myalign(val1,alen,str1)
	dim vlen,k,str2,val,nkk,njj
	val = val1
	IF IsNull(val) Then
		Exit Function
	End IF
	vlen = CInt(len(val))
	if (vlen > alen) then
		val = Mid(val,1,alen)
		vlen = CInt(len(val))
	end if

	k = (alen - vlen)

	if alen = vlen then
		str2 = val
		myalign = str2
	' Left Align
	elseif (str1="L") then
		str2 = val & String(k," ")
		myalign = str2
	' Right Align
	elseif (str1 = "R") then
	    str2 = String(k," ") & val
	    myalign = str2
	' Center Align
	elseif (str1 = "C") then
		nkk = int((alen - len(val)) / 2)
	    str2 = String(nkk," ") & val 
	    njj = alen - len(str2)
	    str2 = str2 & String(njj," ")
	    myalign = str2
	end if
end function

'------------------------End OF MyAlign Function----------------------------
%>