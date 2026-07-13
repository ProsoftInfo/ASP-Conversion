dim objTemp,Root,newElem
dim iClass,iItem,iQty,i,Q,j

Function fnInit(sItem,sClass)
	iClass = sClass
	iItem = sItem
	
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iClass and HeaderNode.Attributes.Item(1).nodeValue = iItem then
				'idItemName.innerHTML = HeaderNode.Attributes.Item(2).nodeValue & "&nbsp;"
				For Each childNod In HeaderNode.childNodes
					if StrComp(Trim(childNod.NodeName),"QtyPara") = 0 then
						For Each PageNod In childNod.childNodes
							i = i + 1
							if (document.formname.elements(i).type = "text") then
								if StrComp(Trim(document.formname.elements(i).Name),"txt") > 0 then
									document.formname.elements(i).value = PageNod.Attributes.Item(2).nodeValue
								end if
							end if
						next
					end if
				next
			end if
		next
	end if
end Function
