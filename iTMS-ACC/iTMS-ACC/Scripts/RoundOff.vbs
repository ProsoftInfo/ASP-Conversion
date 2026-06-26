'Function Used to Round Off the given value
'Rounds to the next greatest integer if the decimal part of the given value is >=5
'Rounds to the next lowest integer if the decimal part of the given value is <5

Function RndOff(iValue)
	Dim n, iIntValue, iDecValue, iCheck

	n = InStr(1,iValue,".")
	if n > 0 then
		iIntValue = mid(iValue,1,n-1)
		iDecValue = mid(iValue,n+1,1)

		iCheck = cdbl("5")

		if cdbl(iDecValue) >= iCheck then 
			RndOff = cdbl(iIntValue) + 1
		else
			RndOff = cdbl(iIntValue)
		end if

	else
		RndOff = cdbl(iValue)
	end if

End Function
