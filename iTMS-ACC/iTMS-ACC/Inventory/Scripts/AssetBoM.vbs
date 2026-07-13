dim objTemp,Root,j,RootO
j = 0

Set Root = Data.documentElement

Function fnInit()
	set objTemp = window.dialogArguments
	Set RootO = objTemp.documentElement

	sExp4 ="//ASSET/ITMDET"
	Set DisplayNode = RootO.Selectnodes(sExp4)

	for z = 0 to DisplayNode.Length - 1
		j = j + 1
		set oRow = document.all.tblData.insertRow(j)

		set headerCell=oRow.insertCell()									
		headerCell.innerHTML=j
		headerCell.className="ExcelSerial"
		headerCell.align="center"

		set headerCell=oRow.insertCell()									
		headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("DESC").Value)
		headerCell.className="ExcelDisplayCell"
		headerCell.align="left"

		set headerCell=oRow.insertCell()									
		headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("QTY").Value)
		headerCell.className="ExcelDisplayCell"
		headerCell.width="10"

		set headerCell=oRow.insertCell()									
		headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("SUOM").Value)
		headerCell.className="ExcelDisplayCell"

		set headerCell=oRow.insertCell()									
		headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("ITYPENAME").Value)
		headerCell.className="ExcelDisplayCell"

	next		

end Function

Function DisplayUoM(obj)
	dim RootNode
	
	if document.formname.selItem.length = 0 then exit function
	idUoM.innerHTML = "&nbsp;"
	iClass = document.formname.selClass.value
	
	Set RootNode = ItemData.documentElement
	
	sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&obj.value&"]"
	Set ItemNode = RootNode.Selectnodes(sExp)
	
	if ItemNode.Length > 0 then
		idUoM.innerHTML = ItemNode.Item(0).Attributes.getNamedItem("SUOM").Value & "&nbsp;"
	end if

end Function


Function GetItem(obj)
	dim RootNode
	
	document.formname.selItem.options.length = 0

	if obj.value ="select" then exit function

	sOrgID = document.formname.hOrgID.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","XMLAssetItem.asp?sOrgID="& sOrgID &"&sClass="&obj.value, false
	
	objhttp.send 

	if objhttp.responseXML.xml <> "" then
		ItemData.loadXML objhttp.responseXML.xml
		Set RootNode = ItemData.documentElement
		if RootNode.HaschildNodes() then
			For Each HeaderNode In RootNode.childNodes
				document.formname.selItem.length = document.formname.selItem.length+1
				document.formname.selItem.options(document.formname.selItem.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
				document.formname.selItem.options(document.formname.selItem.length-1).Value = HeaderNode.Attributes.Item(1).nodeValue
			next
		end if
	else
		alert("No Item defined for the Classification Selected")
		obj.focus()
		Exit Function
	end if

end Function

Function CheckEntry()

	if document.formname.selClass.selectedIndex = "0" then
		alert("Select Classification")
		document.formname.selClass.focus()
		exit Function
	elseif document.formname.selItem.value = "" then
		alert("Select Item")
		document.formname.selItem.focus()
		exit Function
	elseif trim(document.formname.txtQty.value) = "" then
		alert("Enter Quantity")
		document.formname.txtQty.select()
		exit Function
	elseif not isNumeric(document.formname.txtQty.value) then
		alert("Enter Quantity in numerals")
		document.formname.txtQty.select()
		exit Function
	elseif cdbl(document.formname.txtQty.value) <= 0 then
		alert("Quantity cannot be ZERO")
		document.formname.txtQty.select()
		exit Function
	elseif not (document.formname.radType(0).checked or document.formname.radType(1).checked) then
		alert("Select Type")
		document.formname.radType(0).select()
		exit Function
	else
		iItem = document.formname.selItem.value
		iClass = document.formname.selClass.value
		
		sExp ="//ASSET [ @CLACODE = "&iClass&"]"
		Set ItemNode = RootO.Selectnodes(sExp)
		
		if ItemNode.Length > 0 then
			sExp3 ="//ASSET [ @CLACODE = "&iClass&"]/ITMDET [ @ITMCODE = "&iItem&"]"
			Set AddNode = RootO.Selectnodes(sExp3)

			if AddNode.Length > 0 then
				ClearTable

				AddNode.Item(0).Attributes.getNamedItem("DESC").Value = document.formname.selItem(document.formname.selItem.selectedIndex).text & " / " & document.formname.selClass(document.formname.selClass.selectedIndex).text
				AddNode.Item(0).Attributes.getNamedItem("QTY").Value = trim(document.formname.txtQty.value)
				AddNode.Item(0).Attributes.getNamedItem("SUOM").Value = idUoM.innerText

				if document.formname.radType(0).checked then
					AddNode.Item(0).Attributes.getNamedItem("ITYPE").Value = document.formname.radType(0).value
					AddNode.Item(0).Attributes.getNamedItem("ITYPENAME").Value = "Final Component"
				else
					AddNode.Item(0).Attributes.getNamedItem("ITYPE").Value = document.formname.radType(1).value
					AddNode.Item(0).Attributes.getNamedItem("ITYPENAME").Value = "Assembly"
				end if

				sExp4 ="//ASSET/ITMDET"
				Set DisplayNode = RootO.Selectnodes(sExp4)

				for z = 0 to DisplayNode.Length - 1
					j = j + 1
					set oRow = document.all.tblData.insertRow(j)

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML=j
					headerCell.className="ExcelSerial"
					headerCell.align="center"

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("DESC").Value)
					headerCell.className="ExcelDisplayCell"
					headerCell.align="left"

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("QTY").Value)
					headerCell.className="ExcelDisplayCell"
					headerCell.width="10"

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("SUOM").Value)
					headerCell.className="ExcelDisplayCell"

					set headerCell=oRow.insertCell()									
					headerCell.innerHTML= trim(DisplayNode.Item(z).Attributes.getNamedItem("ITYPENAME").Value)
					headerCell.className="ExcelDisplayCell"

				next		

			else
				Set newElem3 = objTemp.createElement("ITMDET")
				
				newElem3.setAttribute "ITMCODE", document.formname.selItem.value
				newElem3.setAttribute "DESC", document.formname.selItem(document.formname.selItem.selectedIndex).text & " / " & document.formname.selClass(document.formname.selClass.selectedIndex).text
				newElem3.setAttribute "QTY", trim(document.formname.txtQty.value)
				newElem3.setAttribute "SUOM", idUoM.innerText
				
				if document.formname.radType(0).checked then
					newElem3.setAttribute "ITYPE", document.formname.radType(0).value
					newElem3.setAttribute "ITYPENAME", "Final Component"
				else
					newElem3.setAttribute "ITYPE", document.formname.radType(1).value
					newElem3.setAttribute "ITYPENAME", "Assembly"
				end if
				
				ItemNode.Item(0).appendChild newElem3

				j = j + 1
				set oRow = document.all.tblData.insertRow(j)

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML=j
				headerCell.className="ExcelSerial"
				headerCell.align="center"

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML= document.formname.selItem(document.formname.selItem.selectedIndex).text & " / " & document.formname.selClass(document.formname.selClass.selectedIndex).text
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML= trim(document.formname.txtQty.value)
				headerCell.className="ExcelDisplayCell"
				headerCell.width="10"

				set headerCell=oRow.insertCell()									
				headerCell.innerHTML= idUoM.innerText
				headerCell.className="ExcelDisplayCell"

				set headerCell=oRow.insertCell()									
				if document.formname.radType(0).checked then
					headerCell.innerHTML= "Final Component"
				else
					headerCell.innerHTML= "Assembly"
				end if
				headerCell.className="ExcelDisplayCell"

			end if

		else
			Set newElem = objTemp.createElement("ASSET")

			newElem.setAttribute "CLACODE", document.formname.selClass.value

			Set newElem3 = objTemp.createElement("ITMDET")

			newElem3.setAttribute "ITMCODE", document.formname.selItem.value
			newElem3.setAttribute "DESC", document.formname.selItem(document.formname.selItem.selectedIndex).text & " / " & document.formname.selClass(document.formname.selClass.selectedIndex).text
			newElem3.setAttribute "QTY", trim(document.formname.txtQty.value)
			newElem3.setAttribute "SUOM", idUoM.innerText
			
			if document.formname.radType(0).checked then
				newElem3.setAttribute "ITYPE", document.formname.radType(0).value
				newElem3.setAttribute "ITYPENAME", "Final Component"
			else
				newElem3.setAttribute "ITYPE", document.formname.radType(1).value
				newElem3.setAttribute "ITYPENAME", "Assembly"
			end if

			newElem.appendChild newElem3

			RootO.appendChild newElem

			j = j + 1
			set oRow = document.all.tblData.insertRow(j)

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML=j
			headerCell.className="ExcelSerial"
			headerCell.align="center"

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML= document.formname.selItem(document.formname.selItem.selectedIndex).text & " / " & document.formname.selClass(document.formname.selClass.selectedIndex).text
			headerCell.className="ExcelDisplayCell"
			headerCell.align="left"

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML= trim(document.formname.txtQty.value)
			headerCell.className="ExcelDisplayCell"
			headerCell.width="10"

			set headerCell=oRow.insertCell()									
			headerCell.innerHTML= idUoM.innerText
			headerCell.className="ExcelDisplayCell"

			set headerCell=oRow.insertCell()									
			if document.formname.radType(0).checked then
				headerCell.innerHTML= "Final Component"
			else
				headerCell.innerHTML= "Assembly"
			end if
			headerCell.className="ExcelDisplayCell"

		
		end if

		document.formname.selClass.selectedIndex = "0"
		document.formname.selItem.length = "0"
		document.formname.txtQty.value = ""
		idUoM.innerHTML = "&nbsp;"
		document.formname.radType(0).checked = false
		document.formname.radType(1).checked = false

	end if
end Function

Function CheckSubmit()
	if j = 0 then
		alert("Enter BoM details")
		exit function
	end if

	window.close
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

Function ClearTable()
	dim i
	for i=1 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(1) 
	next
	j = 0
end Function
