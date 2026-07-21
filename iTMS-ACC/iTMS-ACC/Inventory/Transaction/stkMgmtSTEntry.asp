<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtSTEntry.asp
	'Module Name				:	Inventory (Stock Management Stock Transfer)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 27, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 22,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock Management - Stock Transfer</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Output/></script>
<script type="application/xml" data-itms-xml-island="1" id="IssueData"><ISSTYPE></ISSTYPE></script>
<script type="application/xml" data-itms-xml-island="1" id="IntReceipt"><ROOT></ROOT></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,OutDataValue
dim sorgID,iClass,sStore,iInvRec,sLot,sBin

Function FnInit(sItemCode)
	'document.formname.selItem.value = sItemCode
	GetXML()
End Function

Function GetXML()
	ClearTable
	clearXML

'	set obj = document.formname.selItem

'	sorgID = trim(document.formname.hOrgID.value)
'	iClass = trim(document.formname.hClass.value)

'	if obj.value <> "select" then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'objhttp.Open "GET","itmStoreXMLSelect.asp", false
		objhttp.Open "GET","itmStoreXMLSelectNew.asp", false
		objhttp.send
		'alert(objhttp.responseText)
		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			DisplayDetails()
			popClaDisplay
		else
			clearXML
		end if
'	end if
End Function

Function CheckLot(obj)
	'sTemp = document.formname.selItem(document.formname.selItem.selectedIndex).text & " -- " & trim(idClass.innerText)

	set OutDataValue = showModalDialog("stkMgmtSTPoP.asp?sTemp="&obj.name&"&sValue="&sTemp,OutData,"dialogHeight:310px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	
	arrTemp = split(obj.name,"AAAA")
	sOrgID = arrTemp(1)
	iItem = arrTemp(2)
	iClass = arrTemp(3)
	sLot = arrTemp(4)
	if sLot = "0" then sLot = "NULL"
	sStore = arrTemp(6)
	sBin = arrTemp(7)
	iInvRec = arrTemp(8)

	Set RootO = OutData.documentElement
	'alert(RootO.xml)
	For Each ItemNode in RootO.childNodes
	    if StrComp(trim(ItemNode.nodeName),"Item")=0 then
	        For Each HeaderONode In ItemNode.childNodes
	            if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			        if (HeaderONode.Attributes.Item(0).nodeValue = sStore and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
				        For Each PickNode In HeaderONode.childNodes
					        i = i + 1
					        set objQ = eval("document.formname.txtQtyA"&sStore&"A"&sBin)
					        'objQ.value = "0"

					        if PickNode.Attributes.Item(0).nodeValue = sLot and PickNode.Attributes.Item(1).nodeValue = iInvRec then
						        objQ.value = PickNode.Attributes.Item(3).nodeValue
					        end if
					        'if objQ.value = "" then objQ.value = "0"
				        next
				        exit for
			        end if
		        end if
	        next
	    end if
	Next
	
End Function

Function DisplayDetails()
	Dim sRecNumStatus
	ClearTable
	j = 0

'	iItem = trim(document.formname.selItem.value)

'	arrTemp1 = split(iSTNo,":")
'	if ubound(arrTemp1) > 0 then
'		sLoc = arrTemp1(0)
'		sBin = arrTemp1(1)
''	else
'		sLoc = arrTemp1(0)
'		sBin = "0"
'	end if

	Set Root = OutData.documentElement
	Set RootO = OutData.documentElement

	For Each ndItem In RootO.childNodes
		For each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
				For Each PickNode In HeaderONode.childNodes
					PickNode.setAttribute "QTYISS", ""
					PickNode.setAttribute "STORE", ""
				next
			end if
			if StrComp(Trim(HeaderONode.NodeName),"UOM") = 0 then
				sCheck = HeaderONode.Attributes.Item(2).nodeValue
			end if
		Next
	Next
	i = 1

	For Each ndItem In Root.childNodes
		sItemName = ndItem.getAttribute("IName")
		iItem = ndItem.getAttribute("ICode")
		iClass = ndItem.getAttribute("CCode")
		sorgID = ndItem.getAttribute("Unit")
		sRecNumStatus = ndItem.getAttribute("RecNumStatus")

		For Each HeaderNode in ndItem.childNodes
			sStoreName = HeaderNode.getAttribute("LOCNAME")
			if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
				sLoc = HeaderNode.getAttribute("LOC")
				sBin = HeaderNode.getAttribute("BIN")
					For Each ItemNode In HeaderNode.childNodes
						
						j = j + 1
						set oRow = document.all.tblData.insertRow(j+1)

						set headerCell=oRow.insertCell()
						headerCell.innerHTML=j
						headerCell.className="ExcelSerial"
						headerCell.align="center"

						set headerCell=oRow.insertCell()
						headerCell.innerHTML=sItemName
						headerCell.className="ExcelDisplayCell"
						headerCell.align="center"

						set headerCell=oRow.insertCell()
						If sBin <> "0" and sBin <> "" Then
							headerCell.innerHTML=sStoreName & "-" & sBin
						Else
							headerCell.innerHTML=sStoreName
						End IF
						headerCell.className="ExcelDisplayCell"
						headerCell.align="center"

						set headerCell=oRow.insertCell()
						if trim(ItemNode.Attributes.Item(0).nodeValue) = "NULL" then
							set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value=""-"" READONLY class=""FormelemRead"">" )
							headerCell.appendChild(oText)
						else
							set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value="""&ItemNode.Attributes.Item(0).nodeValue&""" READONLY class=""FormelemRead"">" )
							headerCell.appendChild(oText)
						end if
						headerCell.className="ExcelDisplayCell"
						headerCell.align="left"
						headerCell.width = "150"

						set headerCell=oRow.insertCell()
						set oText = document.createElement("<input type=""text"" name=""txtStk"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY class=""FormelemRead"" style=""text-align=right"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						headerCell.width = "10"
						'alert(ItemNode.Attributes.Item(5).nodeValue)
						'if trim(ItemNode.Attributes.Item(5).nodeValue) = "N" then
						If sRecNumStatus = "N" Then
							set headerCell=oRow.insertCell()
							set oText = document.createElement("<input type=""text"" name=""txtQtyA"&sLoc&"A"&sBin&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value=""0"" class=""Formelem"" style=""text-align=right"" onchange=""CheckQty('"&ItemNode.Attributes.Item(2).nodeValue&"','"&sLoc&"','"&sBin&"')"">" )
							headerCell.appendChild(oText)
							headerCell.width = "10"
							headerCell.className="ExcelInputCell"

							set headerCell=oRow.insertCell()
							set oText = document.createElement("<a href=""#"">" )

							set oText1 = document.createElement("<img name=""btnAAAA"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"">")
							oText.appendChild(oText1)

							headerCell.appendChild(oText)
							headerCell.className="ExcelDisplayCell"
							headerCell.align="center"
							headerCell.width = "10"
						else
							'alert(Root.xml)
							set headerCell=oRow.insertCell()
							set oText = document.createElement("<input type=""text"" name=""txtQtyA"&sLoc&"A"&sBin&""" size=""12"" maxlength=""10"" value=""0"" READONLY class=""Formelem"" style=""text-align=right"">" )
							headerCell.appendChild(oText)
							headerCell.width = "10"
							headerCell.className="ExcelInputCell"

							set headerCell=oRow.insertCell()
							set oText = document.createElement("<a href=""#"">" )
							set oText1 = document.createElement("<img name=""btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&SCheck&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Serial Details"" onClick=""CheckLot(this)"">")
							oText.appendChild(oText1)

							headerCell.appendChild(oText)
							headerCell.className="ExcelDisplayCell"
							headerCell.align="center"
							headerCell.width = "10"
						end if


						set headerCell=oRow.insertCell()
						set oText = document.createElement("<SELECT name=""selSTA"&CStr(j)&""" class=""FormElem"" >" )

						For Each ndItemNew In RootO.childNodes
							For Each HeaderONode in ndItemNew.childNodes
								if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
									if not (HeaderONode.Attributes.Item(0).nodeValue = sLoc and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
										if not HeaderONode.Attributes.Item(1).nodeValue = "0" then
											set oText1 = document.createElement("<Option>" )
											'oText1.Text = trim(HeaderONode.Attributes.Item(2).nodeValue)&" -- "&trim(HeaderONode.Attributes.Item(3).nodeValue)&" ["&trim(HeaderONode.Attributes.Item(4).nodeValue)&"]"
											oText1.Text = trim(HeaderONode.Attributes.Item(2).nodeValue)&" -- "&trim(HeaderONode.Attributes.Item(1).nodeValue)&" ["&trim(HeaderONode.Attributes.Item(4).nodeValue)&"]"
											oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)&":"&trim(HeaderONode.Attributes.Item(1).nodeValue)
											oText.Options.Add(oText1)
										else
											set oText1 = document.createElement("<Option>" )
											oText1.Text = trim(HeaderONode.Attributes.Item(2).nodeValue)&" ["&trim(HeaderONode.Attributes.Item(4).nodeValue)&"]"
											oText1.Value = trim(HeaderONode.Attributes.Item(0).nodeValue)& ":" & "0"
											oText.Options.Add(oText1)
										end if
									end if
								end if
							Next
						Next

						headerCell.appendChild(oText)
						headerCell.className="ExcelFieldCell"
						headerCell.width = "10"
						headerCell.align="center"
						

					next
			end if
		Next
	next
end Function

Function CheckQty(nStkQty,sLoc,sBin)
	nIssueQty = cdbl(Eval("document.formname.txtQtyA"&sLoc&"A"&sBin).value)
	If nIssueQty > cdbl(nStkQty) Then
		alert("Enter Issue Qty which is less than Stock Qty")
		Eval("document.formname.txtQtyA"&sLoc&"A"&sBin).value = "0"
		Exit Function
	End IF
End Function

Function popClaDisplay()

	'document.formname.selStore.options.length = 1
	Set Root = OutData.documentElement

	For Each Node In Root.childNodes
		For Each HeaderNode in Node.childNodes
			if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then

				if not HeaderNode.Attributes.Item(1).nodeValue = "0" then
				'	document.formname.selStore.length = document.formname.selStore.length+1
				'	document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)&" -- "&trim(HeaderNode.Attributes.Item(3).nodeValue)
				'	document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)&":"&trim(HeaderNode.Attributes.Item(1).nodeValue)
				else
				'	document.formname.selStore.length = document.formname.selStore.length+1
				'	document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)
				'	document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)
				end if
			end if
			if StrComp(Trim(HeaderNode.NodeName),"UOM") = 0 then
				idUoM.innerHTML = trim(HeaderNode.Attributes.Item(1).nodeValue) & "&nbsp;"
			end if
		Next
	next

end Function

Function clearXML()
	Set Root = OutData.documentElement
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			set a=Root.removeChild(HeaderNode)
		next
	end if
	'document.formname.selStore.options.length = 1
end Function

Function ClearTable()
	dim i
	for i=2 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(2)
	next
end Function

Function CheckSubmit()
	dim objQ,iQtyTot
	Set Root = OutData.documentElement
	if j = 0 then exit function
	'alert(Root.xml)

	iLen =  document.formname.selSTA1.length

'	alert(iLen)
	if cint(iLen) < 1 then exit function
'
'	iSTNo = trim(document.formname.selStore.value)
'
'	arrTemp1 = split(iSTNo,":")
'	if ubound(arrTemp1) > 0 then
'		sLoc = arrTemp1(0)
'		sBin = arrTemp1(1)
'	else
'		sLoc = arrTemp1(0)
'		sBin = "0"
'	end if

	For Each ndItem In Root.childNodes
		For Each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			sLoc = HeaderONode.getAttribute("LOC")
			sBin = HeaderONode.getAttribute("BIN")
				if (HeaderONode.Attributes.Item(0).nodeValue = sLoc and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
					For Each PickNode In HeaderONode.childNodes
						i = i + 1
						set objQ = eval("document.formname.txtQtyA"&sLoc&"A"&sBin)
						if trim(objQ.value) = "" then
							alert("Enter Quantity")
							objQ.select()
							exit function
						else
							set obj = eval("document.formname.txtStk"&cstr(i))

							if (cdbl(objQ.value) > cdbl(obj.value)) then
								alert("Transfer Quantity should be equal to or less than Stock Quantity")
								objQ.select()
								exit function
							end if

							iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
						end if
					next
				end if
			end if
		Next
	Next

	i = 1

	Set RootO = OutData.documentElement
	For Each ndItem In RootO.childNodes
		For Each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			sLoc = HeaderONode.getAttribute("LOC")
			sBin = HeaderONode.getAttribute("BIN")
			'	if (HeaderONode.Attributes.Item(0).nodeValue = sLoc and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
					For Each PickNode In HeaderONode.childNodes
						set objQ = eval("document.formname.txtQtyA"&sLoc&"A"&sBin)
						PickNode.setAttribute "QTYISS", cdbl(objQ.value)

						set objQ = eval("document.formname.selSTA"&cstr(i))
						sTemp = split(objQ.value,":")
						PickNode.setAttribute "STORE", sTemp(0)'objQ.value
						PickNode.setAttribute "BIN", sTemp(1)
						i = i + 1
					next
			'	else
			'		set a = RootO.removeChild(HeaderONode)
			'	end if
			end if
		Next
	Next

'	sorgID = trim(document.formname.hOrgID.value)
'	iClass = trim(document.formname.hClass.value)

'	RootO.setAttribute "ITEM", trim(document.formname.selItem.value)
'	RootO.setAttribute "CLASS", iClass
'	RootO.setAttribute "ORG", sorgID
'	alert(OutData.xml)
	
	'ISSUES XML CONSTRUCTION
	set OutDataRoot = OutData.documentElement
	
	set ndIssueRoot = IssueData.documentElement
	ndIssueRoot.setAttribute "ISSTYPE","F"
	ndIssueRoot.setAttribute "ISSTOCODE","INV"
	ndIssueRoot.setAttribute "ISSTOTYPE","DEPT"
	ndIssueRoot.setAttribute "ISSTOSUBCODE",""
	ndIssueRoot.setAttribute "POConfirm","N"
	ndIssueRoot.setAttribute "SInvConfirm","N"
	ndIssueRoot.setAttribute "Invoice","A"
	ndIssueRoot.setAttribute "GPConfirm","N"
	ndIssueRoot.setAttribute "ProConfirm","N"
	ndIssueRoot.setAttribute "MCallFrom","N"
	ndIssueRoot.setAttribute "RedirectTo",""
	ndIssueRoot.setAttribute "AppRefType",""
	ndIssueRoot.setAttribute "AppRefNo",""
	ndIssueRoot.setAttribute "AppRefDate",""
	ndIssueRoot.setAttribute "ConsumptionAccHead",""
	ndIssueRoot.setAttribute "IssueToCode",""
	ndIssueRoot.setAttribute "PickPackFlag",""
	ndIssueRoot.setAttribute "IssFrom","IN"
	ndIssueRoot.setAttribute "Returnable","N"
	ndIssueRoot.setAttribute "ReturnItem","S"
	ndIssueRoot.setAttribute "TYPE","GEN"
	
	
	
	set ndRcptRoot = IntReceipt.documentElement
		ndRcptRoot.setAttribute "DEPT","OTH"
		ndRcptRoot.setAttribute "SOURCE","N"
		ndRcptRoot.setAttribute "ORGCODE",document.formname.hOrgID.value
		ndRcptRoot.setAttribute "STYPE","N"
		ndRcptRoot.setAttribute "ITEMTYPE",""
		ndRcptRoot.setAttribute "PACKNUM",""
		ndRcptRoot.setAttribute "SRCREFTYPE","N"
		ndRcptRoot.setAttribute "SRCREFNO",""
		ndRcptRoot.setAttribute "RCPTNUMBERINV",""
		ndRcptRoot.setAttribute "sTypeRcpt",""
		ndRcptRoot.setAttribute "APPREFTYPE",""
		ndRcptRoot.setAttribute "APPREFNO",""
		ndRcptRoot.setAttribute "APPREFDATE",""
		ndRcptRoot.setAttribute "RCVDON",date()
		ndRcptRoot.setAttribute "AUTOACCOUNT","Y"
	
	set ndRcptDetails = IntReceipt.createElement("Details")
	ndRcptRoot.appendChild ndRcptDetails
	
	
	sExp = "//Item"
	set TempNode = OutDataRoot.selectNodes(sExp)
	If TempNode.length <> "0" Then
		For i = 1 to TempNode.length
		sItemRcptNum = TempNode.item(i-1).Attributes.getNamedItem("RecNumStatus").value
			nTotQtyIssued = "0"
			sExp1 = "//Item[@ICode="&trim(TempNode.item(i-1).Attributes.getNamedItem("ICode").value)&" and @CCode="& trim(TempNode.item(i-1).Attributes.getNamedItem("CCode").value)&"]/LOCDET/PICK"
			set TempNode1 = OutDataRoot.selectNodes(sExp1)
			If TempNode1.length <> "0" Then
				For j = 1 to TempNode1.length
					nTotQtyIssued = cdbl(nTotQtyIssued) + cdbl(TempNode1.item(j-1).Attributes.getNamedItem("QTYISS").value)
				Next
			End IF
			
			set ndIssueItem = IssueData.createElement("ITEM")
	        ndIssueItem.setAttribute "ENTRYNO",i
	        ndIssueItem.setAttribute "ITMCODE",TempNode.item(i-1).Attributes.getNamedItem("ICode").value
	        ndIssueItem.setAttribute "CLACODE",TempNode.item(i-1).Attributes.getNamedItem("CCode").value
	        ndIssueItem.setAttribute "ITMNAME",TempNode.item(i-1).Attributes.getNamedItem("IName").value
	        ndIssueItem.setAttribute "SSTORE",""
	        ndIssueItem.setAttribute "REQQTY","0"
	        ndIssueItem.setAttribute "REQBY",""
	        ndIssueItem.setAttribute "REMARKS",""
	        ndIssueItem.setAttribute "ITEMTYPE",""'TempNode.item(i-1).Attributes.getNamedItem("ItemTypeID").value
	        ndIssueItem.setAttribute "ISSUEDATE",date()
	        ndIssueItem.setAttribute "ISSQTY",nTotQtyIssued
	        ndIssueItem.setAttribute "TRAQTY","0"
	        ndIssueItem.setAttribute "PRQTY","0"
	        ndIssueItem.setAttribute "IVALUE","0"
	        ndIssueItem.setAttribute "ORGCODE",document.formname.hOrgID.value
	        ndIssueItem.setAttribute "MRSNO",""
	        ndIssueItem.setAttribute "MRSDATE",""
	        ndIssueItem.setAttribute "ATTRIBUTELIST",""
	        ndIssueItem.setAttribute "CREATEDBY", document.formname.hUserID.value 
	        ndIssueItem.setAttribute "CREATEDON", date()
	        ndIssueItem.setAttribute "RETURNABLE","N"
	        ndIssueItem.setAttribute "RefNo",""
	        ndIssueItem.setAttribute "ONLYLOT",""
	        ndIssueItem.setAttribute "RETURNITEM","S"
	        ndissueItem.setAttribute "MatType",""
	        ndIssueRoot.appendChild ndIssueItem 
	        
	        set ndIssuePick = IssueData.createElement("Pick")
			ndIssuePick.setAttribute "TOT",nTotQtyIssued
			ndIssuePick.setAttribute "NoofPack",""
			ndIssueItem.appendChild ndIssuePick
			
			sExp1 = "//Item[@ICode="&trim(TempNode.item(i-1).Attributes.getNamedItem("ICode").value)&" and @CCode="& trim(TempNode.item(i-1).Attributes.getNamedItem("CCode").value)&"]/LOCDET"
			set TempNode1 = OutDataRoot.selectNodes(sExp1)
			If TempNode1.length <> "0" Then
				For k = 1 to TempNode1.length
					nLocNo = TempNode1.item(k-1).Attributes.getNamedItem("LOC").value
					nBinNo = TempNode1.item(k-1).Attributes.getNamedItem("BIN").value
					
					sExp1 = "//Item[@ICode="&trim(TempNode.item(i-1).Attributes.getNamedItem("ICode").value)&" and @CCode="& trim(TempNode.item(i-1).Attributes.getNamedItem("CCode").value)&"]/LOCDET[@LOC="&trim(nLocNo)&" and @BIN="&trim(nBinNo)&"]/PICK"
					'alert(sExp1)
					set TempNode2 = OutDataRoot.selectNodes(sExp1)
					If TempNode2.length <> "0" Then
					
						If TempNode2.item(0).Attributes.getNamedItem("QTYISS").value <> "0" Then
							If TempNode2.item(0).Attributes.getNamedItem("LOTNO").value = "" Then
								nLotNo = "N/A"
							End IF
							if trim(sItemRcptNum)="N" then
							    set ndIssueStore = IssueData.createElement("STORE")
							    ndIssueStore.setAttribute "LOC",nLocNo
							    ndIssueStore.setAttribute "BIN",nBinNo 
							    ndIssueStore.setAttribute "LOTNO",nLotNo
							    ndIssueStore.setAttribute "INVRECNO",TempNode2.item(0).Attributes.getNamedItem("INVRECNO").value
							    ndIssueStore.setAttribute "QTYISS",TempNode2.item(0).Attributes.getNamedItem("QTYISS").value
							    ndIssueStore.setAttribute "NoofPack",""
							    ndIssuePick.appendChild ndIssueStore
							else
							
							    set ndIssueStore = IssueData.createElement("PICK")
							    ndIssueStore.setAttribute "LOC",nLocNo
							    ndIssueStore.setAttribute "BIN",nBinNo 
							    ndIssueStore.setAttribute "LOTNO",nLotNo
							    ndIssueStore.setAttribute "INVRECNO",TempNode2.item(0).Attributes.getNamedItem("INVRECNO").value
							    ndIssueStore.setAttribute "QTYISS",TempNode2.item(0).Attributes.getNamedItem("QTYISS").value
							    ndIssueStore.setAttribute "NoofPack",""
							    ndIssuePick.appendChild ndIssueStore
							    
							sExp3 = "//Item[@ICode="&trim(TempNode.item(i-1).Attributes.getNamedItem("ICode").value)&" and @CCode="& trim(TempNode.item(i-1).Attributes.getNamedItem("CCode").value)&"]/LOCDET[@LOC="&trim(nLocNo)&" and @BIN="&trim(nBinNo)&"]/PICK/SERIALDETAILS"
							'alert(sExp3)
					            set TempNode3 = OutDataRoot.selectNodes(sExp3)
					            If TempNode3.length <> "0" Then
					                set ndIssSerHdr = IssueData.createElement("SERIALHEADER")
					                ndIssueStore.appendChild ndIssSerHdr
					                For K3 = 1 to TempNode3.length
					                    set ndIssSerDet = IssueData.createElement("SERIALDETAILS")
					                        ndIssSerDet.setAttribute "SERIALNO",TempNode3.Item(K3-1).Attributes.getNamedItem("SERIALNO").value
					                        ndIssSerDet.setAttribute "QTY",TempNode3.Item(K3-1).Attributes.getNamedItem("QTY").value
					                    ndIssSerHdr.appendChild ndIssSerDet
					                Next
					            End If
					            
					        End If 'if trim(sItemRcptNum)="N" then
							
						End IF
					End IF
					
				Next
			End IF
			
			
			
			set ndRcptItem = IntReceipt.createElement("ItemDetail")
				ndRcptItem.setAttribute "ItemCode",TempNode.item(i-1).Attributes.getNamedItem("ICode").value
				ndRcptItem.setAttribute "CLACODE",TempNode.item(i-1).Attributes.getNamedItem("CCode").value
				ndRcptItem.setAttribute "QTY",nTotQtyIssued
				ndRcptItem.setAttribute "MRSNO","N"
				ndRcptItem.setAttribute "ISSNO","N"
				ndRcptItem.setAttribute "ENTRYNO",i
				ndRcptItem.setAttribute "UNIT",document.formname.hOrgID.value
				ndRcptItem.setAttribute "ITEMNAME",TempNode.item(i-1).Attributes.getNamedItem("IName").value
				ndRcptItem.setAttribute "UOM",""
				ndRcptItem.setAttribute "ATTRIBUTELIST",""
				ndRcptItem.setAttribute "RefNo",""
				ndRcptItem.setAttribute "RefQty",""
				ndRcptItem.setAttribute "RECEIPTNUM",TempNode.item(i-1).Attributes.getNamedItem("RecNumStatus").value
				ndRcptItem.setAttribute "BYPRODUCT","P"
				ndRcptDetails.appendChild ndRcptItem
				
				sExp1 = "//Item[@ICode="&trim(TempNode.item(i-1).Attributes.getNamedItem("ICode").value)&" and @CCode="& trim(TempNode.item(i-1).Attributes.getNamedItem("CCode").value)&"]/LOCDET"
				set TempNode1 = OutDataRoot.selectNodes(sExp1)
				If TempNode1.length <> "0" Then
					For k = 1 to TempNode1.length
						nLocNo = TempNode1.item(k-1).Attributes.getNamedItem("LOC").value
						nBinNo = TempNode1.item(k-1).Attributes.getNamedItem("BIN").value
						
						iTotRcptVal  = 0
						sExp1 = "//Item[@ICode="&trim(TempNode.item(i-1).Attributes.getNamedItem("ICode").value)&" and @CCode="& trim(TempNode.item(i-1).Attributes.getNamedItem("CCode").value)&"]/LOCDET[@LOC="&trim(nLocNo)&" and @BIN="& nBinNo &"]/PICK"
						
						set TempNode2 = OutDataRoot.selectNodes(sExp1)
						If TempNode2.length <> "0" Then
						    If TempNode2.item(0).Attributes.getNamedItem("QTYISS").value <> "0" Then
								set ndRcptStorage = IntReceipt.createElement("STORAGE")
								ndRcptStorage.setAttribute "STORE",TempNode2.item(0).Attributes.getNamedItem("STORE").value
								if trim(TempNode2.item(0).Attributes.getNamedItem("BIN").value)="0" then
								    ndRcptStorage.setAttribute "BIN","NULL"
								else
								    ndRcptStorage.setAttribute "BIN",TempNode2.item(0).Attributes.getNamedItem("BIN").value
								end if
								ndRcptStorage.setAttribute "APPLICABLE","IN"
								ndRcptStorage.setAttribute "MONTHYEAR",Date()
								ndRcptStorage.setAttribute "QTY",TempNode2.item(0).Attributes.getNamedItem("QTYISS").value
								ndRcptStorage.setAttribute "STORAGEVALUE","0"
								ndRcptItem.appendChild ndRcptStorage
								
								
								sExp3 = "//Item[@ICode="&trim(TempNode.item(i-1).Attributes.getNamedItem("ICode").value)&" and @CCode="& trim(TempNode.item(i-1).Attributes.getNamedItem("CCode").value)&"]/LOCDET[@LOC="&trim(nLocNo)&" and @BIN="&trim(nBinNo)&"]/PICK/SERIALDETAILS"
					            set TempNode3 = OutDataRoot.selectNodes(sExp3)
					            If TempNode3.length <> "0" Then
					                set ndRcptLotSer = IssueData.createElement("LotSerial")
					                ndRcptLotSer.setAttribute "QTYIN","N"
					                ndRcptLotSer.setAttribute "TARE","0" 
					                if trim(TempNode2.item(0).Attributes.getNamedItem("LOTNO").value)<>"NULL" then
					                    ndRcptLotSer.setAttribute "LOT",TempNode2.item(0).Attributes.getNamedItem("LOTNO").value
					                else
					                    ndRcptLotSer.setAttribute "LOT",""
					                end if
					                ndRcptLotSer.setAttribute "SERIALFROM","" 
					                ndRcptLotSer.setAttribute "SERIALTO","" 
					                ndRcptLotSer.setAttribute "TAREWEIGHT","U" 
					                ndRcptLotSer.setAttribute "QTY",TempNode2.item(0).Attributes.getNamedItem("QTYISS").value
					                ndRcptLotSer.setAttribute "COUNTER","1" 
					                ndRcptLotSer.setAttribute "STAGE","" 
					                ndRcptLotSer.setAttribute "ALTGROSS","0" 
					                ndRcptLotSer.setAttribute "ALTNETT","0" 
					                ndRcptLotSer.setAttribute "ALTUOM","select" 
					                ndRcptLotSer.setAttribute "IVALUE","0"
					                ndRcptLotSer.setAttribute "AUTOGEN","AUTO"
					                ndRcptStorage.appendChild ndRcptLotSer
					                
					                For K3 = 1 to TempNode3.length
					                    set ndRcptLotSerDet = IssueData.createElement("LotSerialDetails")
					                        sArrTemp = split(TempNode3.Item(K3-1).Attributes.getNamedItem("OTHDET").value,":")
					                        
					                        iTotRcptVal = iTotRcptVal + ((TempNode3.Item(K3-1).Attributes.getNamedItem("QTY").value)*cdbl(sArrTemp(2)))
					                        
					                        ndRcptLotSerDet.setAttribute "LOTSERIAL",K3
					                        ndRcptLotSerDet.setAttribute "QTYREC",TempNode3.Item(K3-1).Attributes.getNamedItem("QTY").value
					                        ndRcptLotSerDet.setAttribute "TAREREC","0" 
					                        ndRcptLotSerDet.setAttribute "SELLINGTYPE","" 
					                        ndRcptLotSerDet.setAttribute "WEIGHTSTYPE","0" 
					                        ndRcptLotSerDet.setAttribute "PACKINGTYPE",sArrTemp(1)
					                        ndRcptLotSerDet.setAttribute "LOT",sArrTemp(4)
					                        ndRcptLotSerDet.setAttribute "SELLINGFORM","" 
					                        ndRcptLotSerDet.setAttribute "PACKNUMBER",sArrTemp(0)
					                        ndRcptLotSerDet.setAttribute "IVALUE",((TempNode3.Item(K3-1).Attributes.getNamedItem("QTY").value)*cdbl(sArrTemp(2)))
					                        ndRcptLotSerDet.setAttribute "ATTRIBUTELIST",sArrTemp(3)
					                        ndRcptLotSerDet.setAttribute "NOOFCONE","0" 
					                        ndRcptLotSerDet.setAttribute "SUBLEVELID",""
					                        
					                    ndRcptLotSer.appendChild ndRcptLotSerDet
					                Next
					                
					                ndRcptLotSer.setAttribute "IVALUE",iTotRcptVal
					                ndRcptStorage.setAttribute "STORAGEVALUE",iTotRcptVal
					                
					            End If
								
								
							End IF
						End IF
					Next
				End IF
				
		Next
	End IF
'	alert(IntReceipt.xml)
'   exit function
	
	set objhttp2 = CreateObject("Microsoft.XMLHTTP")
	objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=StockTransferData",false
	objhttp2.send OutData.xmlDocument
	
	set objhttp2 = CreateObject("Microsoft.XMLHTTP")
	objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=mrsIssueData",false
	objhttp2.send IssueData.xmlDocument
	
	set objhttp2 = CreateObject("Microsoft.XMLHTTP")
	objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=ReceiptLotData",false
	objhttp2.send IntReceipt.xmlDocument
	'Exit Function
	document.formname.B7.disabled = True
	
	
	document.formname.action = "StkTransferInsert.asp"
	document.formname.submit 
end Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp


	set rsTemp = server.CreateObject("ADODB.Recordset")

	sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	'sOrgID = trim(Request.Form("selUnit"))
	sOrgID =session("organizationcode")
	iClass = trim(Request.Form("selClass"))
	sTemp = trim(Request.Form("hSelectedValue"))
	sTempName = trim(Request.Form("hItemNames"))
	
	
	
	if sTempName  = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select ItemDescription from Inv_M_ItemMaster where ItemCode = " & mid(sTemp,1,len(sTemp)-1)
			.Open
		end with

		if not rsTemp.EOF then
			sTempName = trim(rsTemp(0)) & "|"
		end if
		rsTemp.Close
	end if 'if sTempName  = "" then

	if trim(sClassName) = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select GroupName from Inv_M_Classification where GroupCode = " & iClass
			.Open
		end with

		if not rsTemp.EOF then
			sClassName = trim(rsTemp(0))
		end if
		rsTemp.Close
	end if 'if trim(sClassName) = "" then
	
	'Response.Write "<p><font color=red>sTempName="&sTempName & "-"& sTemp 
	
	arrTempName = split(mid(sTempName,1,len(sTempName)-1),"|")
	arrTemp = split(mid(sTemp,1,len(sTemp)-1),"|")

%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit('<%= mid(sTemp,1,len(sTemp)-1)%>')">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hClass" value="<%=iClass%>">
<input type=hidden name="hUserID" value="<%=Session("userID")%>">

<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Stock Transfer
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<!--<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td width="100%" align="center">Header</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Control</td>
										</tr>
									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    <font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>
					</td>
				</tr>-->
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <!--<tr>
                                            <td class="FieldCell">Organization</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                        </tr>-->
                                        <tr>
                                            <td class="FieldCell">Classification</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idClass"><%=sClassName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <!--<tr>
                                            <td class="FieldCell">Item </td>
                                            <td class="FieldCellSub" colspan=4>
												<select size="1" name="selItem" class="FormElem" onChange="GetXML()">
													<option value="select">Select</option>
												<%
													for iCtr = 0 to UBound(arrTempName)
														If iCtr = 0 Then
													%>
														<option value="<%=arrTemp(iCtr)%>" selected><%=arrTempName(iCtr)%></option>
													<%
														Else
													%>
														<option value="<%=arrTemp(iCtr)%>"><%=arrTempName(iCtr)%></option>
													<%	End IF
													next
												%>
												</select>
                                            </td>
                                        </tr>-->
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly" id="idUoM"></span>
                                            </td>
                                        </tr>
                                        <!--<tr>
                                            <td class="FieldCell">Store -- Bin</td>
                                            <!--<td class="FieldCellSub">-->
												<!--<select size="1" name="selStore" class="FormElem" onChange="DisplayDetails(this.value)">-->
												<!--<select size="1" name="selStore" class="FormElem">
													<option value="select">Select</option>
												</select>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell"></td>
                                            <td class="FieldCellSub"></td>
                                        </tr>-->
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="frm2" style="width: 750; height:300;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Item Name</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Store - Bin </td>
                                                <td class="ExcelHeaderCell" align="center" colspan="2">Existing Stock Information</td>
                                                <td class="ExcelHeaderCell" align="center" colspan="4">Stock Transfer</td>
                                            </tr>
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center">Lot Number</td>
                                                <td class="ExcelHeaderCell" align="center">Stock</td>
                                                <td class="ExcelHeaderCell" align="center">Quantity</td>
                                                <td class="ExcelHeaderCell" align="center">Serial</td>
                                                <td class="ExcelHeaderCell" align="center">Store -- Bin [Stock]</td>
                                            </tr>
                                        </table>
                                    </div>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="left" colspan="3" class="FieldCell">&nbsp; <B>Note: The above stock shown excludes the quantity already reserved for Issue</B></td>
							</tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Transfer" name="B7" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="reset" value="Reset" name="B1" class="ActionButton">
                                                    <input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('stkMgmtEntry.asp')">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
