Function Cancel(sLoc)
	if confirm("Do you want to Cancel, If so the data entered will be lost.") then
		window.location.href = sLoc
	else
		exit function
	end if
end Function
