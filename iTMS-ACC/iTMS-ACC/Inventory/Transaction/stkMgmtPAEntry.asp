<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtPAEntry.asp
	'Module Name				:	Inventory (Stock Management Physical Adjustment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 30, 2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	stkMgmtPAInsert.asp
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
<html><head><title>Stock Management - Physical Adjustment</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1"/>
<meta content="Microsoft FrontPage 4.0" name="GENERATOR"/>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css"/>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="IssueData"><ISSTYPE></ISSTYPE></script>
<script type="application/xml" data-itms-xml-island="1" id="IntReceipt"><ROOT></ROOT></script>
<script language="javascript" type="text/javascript" src="../../scripts/rolloverout.js"></script>
<script language="javascript" type="text/vbscript" src="../../scripts/Cancel.js"></script>
<script language="javascript" type="text/javascript" src="../../scripts/ValidateFormat.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,OutDataValue
dim sorgID,iClass,sStore,iInvRec,sLot,sBin

Function FnInit(sItemCode)
	'document.formname.selItem.value = sItemCode
	GetXML()
	DisplayDetails
End Function

Function GetXML()
		Dim iItem,iClass,sOrgID
		
		clearXML
		
		iItem  =  Trim(document.formname.hOrgID.value)
		iClass =  Trim(document.formname.hItem.value)
		sorgID =  Trim(document.formname.hClass.value)
		sTemp = sorgID &":"& iClass & ":" & iItem  
		
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","itmStoreXMLSelect.asp?Data="&sTemp, false
		objhttp.send
		
		'alert(objhttp.responseText)
		
		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
		else
			clearXML
		end if
	
End Function

Function CheckLot(obj,iValue)
    sReceiptType = document.formname.hRcptNum.value
    if trim(sReceiptType)="N" then exit function
    
    arrTemp = split(obj.value,"AAAA")
	    sOrgID = arrTemp(1)
	    iItem = arrTemp(2)
	    iClass = arrTemp(3)
	    sLot = arrTemp(4)
    	
	    if sLot = "0" then sLot = "NULL"
    	
	    sStore = arrTemp(6)
	    sBin = arrTemp(7)
	    iInvRec = arrTemp(8)
	    sAttList = arrTemp(10)
	    'alert(sAttList)
	    Set RootO = OutData.documentElement
	    if RootO.HasChildNodes() then
	        for each ndItem in RootO.childNodes
	            if ndItem.nodeName="Item" then
	                For Each HeaderONode In ndItem.childNodes
            	        if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			                if (HeaderONode.Attributes.Item(0).nodeValue = sStore and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
			                    sStoreName = trim(HeaderONode.Attributes.Item(2).nodeValue)
			                    exit for
			                end if
			            end if
		            Next
		        end if
	        next
		end if
		
    
	if cdbl(iValue)<0 then
	    sTemp = "" & " -- " & replace(sStoreName,"&","and") &"`"&iValue
	
	    set OutDataValue = showModalDialog("stkMgmtPAPoP.asp?sTemp="&obj.value&"&sValue="&sTemp,OutData,"dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")

	    Set RootO = OutData.documentElement
	    if RootO.HasChildNodes() then
	        for each ndItem in RootO.childNodes
	            if ndItem.nodeName="Item" then
	                For Each HeaderONode In ndItem.childNodes
		                if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			                if (HeaderONode.Attributes.Item(0).nodeValue = sStore and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
				                For Each PickNode In HeaderONode.childNodes
					                i = i + 1
					                if PickNode.Attributes.Item(0).nodeValue = sLot and PickNode.Attributes.Item(1).nodeValue = iInvRec then
						                if PickNode.Attributes.Item(3).nodeValue = "" then
							                set objS = eval("document.formname.txtStk"&cstr(i))

							                set objQ = eval("document.formname.txtQtyA"&i)
							                objQ.value = objS.value

							                set obj = eval("document.formname.txtJusA"&i)
							                obj.value = cdbl(objS.value) - cdbl(objQ.value)
						                end if
					                end if
				                next
				                exit for
			                end if
		                end if
	                Next
		        end if
	        next
		end if
	else
	'alert(OutData.xml)
	        
		sTempValues = sReceiptType&"``"&iItem&"``"&iClass&"``"&sOrgID&"``"&sStore&"``"&sBin&"``"&iValue&"``"&replace(sStoreName,"&","and")&"``"&sUOM&"``0``"&date&"````"&sAttID&"``"&sLot&"``"&sAttList
		Set OutDataValue = showModalDialog("../../Common/PackingLotSerForPA.asp?sTemp="&sTempValues,OutData,"dialogHeight:580px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
		
		set RootO = OutData.documentElement
		if RootO.haschildNodes() then
		    for each ndItem in RootO.childNodes
		        if ndItem.nodeName="Item" then
		            for each ndLocDet in ndItem.childNodes
		                if ndLocDet.nodeName="LOCDET" then
		                    
		                end if' if ndLocDet.nodeName="LOCDET" then
		            next
		        end if
		    next
		end if
	end if 'if cdbl(iValue)<0 then
End Function

Function CheckSerial(obj,iValue)

    set obj = eval("document.formname.txtQtyA"&iValue)
    Check(obj)

End Function

Function CheckStock(obj)
	if trim(obj.value) = "" then obj.value = obj.defaultValue
	obj.defaultValue = obj.value
	sStr = obj.name
	iValue = obj.value

	arrTemp = split(sStr,"A")
	iCount = arrTemp(1)
	
	set objS = eval("document.formname.txtStk"&cstr(iCount))
	set obj = eval("document.formname."&Replace(obj.name,"Qty","Jus"))
	obj.value = cdbl(iValue) - cdbl(objS.value)

End Function

Function Check(obj)

	if trim(obj.value) = "" then obj.value = obj.defaultValue
	obj.defaultValue = obj.value
	sStr = obj.name
	iValue = obj.value
	
	arrTemp = split(sStr,"A")
	iCount = arrTemp(1)
	
	set objS = eval("document.formname.txtStk"&cstr(iCount))
	set objJ = eval("document.formname."&Replace(obj.name,"Qty","Jus"))
	
	objJ.value = cdbl(iValue) - cdbl(objS.value)
	
	set objH = eval("document.formname."&Replace(obj.name,"txtQty","h"))
	if cdbl(objJ.value) <> 0 then CheckLot objH,objJ.value

End Function

Function DisplayDetails()
	Dim sPrintUOM,sorgID,iItem,iClass,sLoc,sBin,sCheck
	
	ClearTable
	j = 0
	sPrintUOM = "N"
	
	Set Root = OutData.documentElement
	Set RootO = OutData.documentElement

	For Each ndItem In RootO.childNodes
		For each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
				For Each PickNode In HeaderONode.childNodes
					PickNode.setAttribute "QTYISS", ""
					PickNode.setAttribute "ADJUSTED", ""
				next
			end if
			if StrComp(Trim(HeaderONode.NodeName),"UOM") = 0 then
				sCheck = HeaderONode.Attributes.Item(2).nodeValue
			end if
		Next 'For each HeaderONode in ndItem.childNodes
	next
	
	i = 1
	'alert(Root.xml)
	For Each ndItem In Root.childNodes
		sItemName = ndItem.getAttribute("IName")
		sorgID = ndItem.getAttribute("Unit")
		iItem = ndItem.getAttribute("ICode")
		iClass = ndItem.getAttribute("CCode")
		For each HeaderNode in ndItem.childNodes
			if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
				sStoreName= HeaderNode.getAttribute("LOCNAME")
				sLoc = HeaderNode.getAttribute("LOC")
				sBin = HeaderNode.getAttribute("BIN")
					For Each ItemNode In HeaderNode.childNodes
					sAttName = ""
					sAttList =""
					
					sAttList = ItemNode.Attributes.Item(6).nodeValue
					
					if trim(sAttList)<>"" and (not IsNull(sAttList)) then
					    set objhttp = createObject("Microsoft.XMLHTTP")
					    objhttp.open "GET","../../Include/GetAttrName.asp?AttID="&replace(sAttList,",",":"),false
					    objhttp.send
					    if trim(objhttp.responseText)<>"" then
					        sAttName = objhttp.responseText
					    end if
					end if
					'alert(sAttName)
					
					'alert(ItemNode.xml)
						j = j + 1
						set oRow = document.all.tblData.insertRow(j+1)

						set headerCell=oRow.insertCell()									
						headerCell.innerHTML=j
						headerCell.className="ExcelSerial"
						headerCell.align="center"
						
						set headerCell=oRow.insertCell()									
						headerCell.innerHTML=sItemName
						set oText = document.createElement("<input type=""hidden"" name=""txtItemName"&CStr(j)&""" value="""&sItemName&""" class=""FormelemRead"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						
						set headerCell=oRow.insertCell()									
						headerCell.innerHTML=sStoreName
						set oText = document.createElement("<input type=""hidden"" name=""txtClassName"&CStr(j)&""" value="""&sStoreName&""" class=""FormelemRead"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"

						set headerCell=oRow.insertCell()									
						if trim(ItemNode.Attributes.Item(0).nodeValue) = "NULL" then
							set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value=""-"" READONLY class=""FormelemRead"">" )
							headerCell.appendChild(oText)
						else
							set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value="""&ItemNode.Attributes.Item(0).nodeValue&"["&sAttName&"]"&""" READONLY class=""FormelemRead"">" )
							headerCell.appendChild(oText)
						end if
						headerCell.className="ExcelDisplayCell"
						headerCell.align="left"
						headerCell.width = "50"

						set headerCell=oRow.insertCell()									
						set oText = document.createElement("<input type=""text"" name=""txtStk"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY class=""FormelemRead"" style=""text-align=right"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						headerCell.width = "10"
						
						if trim(ItemNode.Attributes.Item(5).nodeValue) = "N" then
							sPrintUOM = "Y"
							set headerCell=oRow.insertCell()									
							set oText = document.createElement("<input type=""text"" name=""txtQtyA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" class=""Formelem"" onBlur=""CheckStock(this)"" style=""text-align=right"">" )
							headerCell.appendChild(oText)
							headerCell.width = "10"
							headerCell.className="ExcelInputCell"

							set oText = document.createElement("<input type=""hidden"" name=""hA"&j&""">" )
							document.formname.appendChild(oText)

							set headerCell=oRow.insertCell()									
							set oText = document.createElement("<a href=""#"">" )

							set oText1 = document.createElement("<img name=""btnA"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"">")
							oText.appendChild(oText1)

							headerCell.appendChild(oText)
							headerCell.className="ExcelDisplayCell"
							headerCell.align="center"
							headerCell.width = "10"
						else
							
							
							'alert(sAttList)
							
							sPrintUOM = "Y"
							set headerCell=oRow.insertCell()									
							set oText = document.createElement("<input type=""text"" name=""txtQtyA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" class=""Formelem"" onBlur=""Check(this)"" style=""text-align=right"">" )
							headerCell.appendChild(oText)
							headerCell.width = "10"
							headerCell.className="ExcelInputCell"

							set oText = document.createElement("<input type=""hidden"" name=""hA"&j&""" value=""btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&SCheck&"AAAA"&sAttList&""">" )
							document.formname.appendChild(oText)

							set headerCell=oRow.insertCell()									
							set oText = document.createElement("<a href=""#"">" )
							'set oText1 = document.createElement("<img name=""btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&sCheck&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Serial Details"" DISABLED onClick=""CheckSerial(this,"&j&")"">")
							set oText1 = document.createElement("<img name=""btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&sCheck&"AAAA"&sAttList&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Serial Details"" onClick=""CheckSerial(this,"&j&")"">")
							oText.appendChild(oText1)

							headerCell.appendChild(oText)
							headerCell.className="ExcelDisplayCell"
							headerCell.align="center"
							headerCell.width = "10"
						end if


						set headerCell=oRow.insertCell()
						set oText = document.createElement("<input type=""text"" name=""txtJusA"&j&""" size=""12"" maxlength=""10"" value=""0"" READONLY class=""Formelem"" style=""text-align=right"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelInputCell"
						headerCell.width = "10"
						headerCell.align="center"
						
						set headerCell=oRow.insertCell()
						set oText = document.createElement("<input type=""text"" name=""txtReasonA"&j&""" size=""12"" maxlength=""100"" value="""" class=""Formelem"" style=""text-align=right"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelInputCell"
						headerCell.width = "10"
						headerCell.align="center"
					next
			elseif strcomp(HeaderNode.nodeName,"UOM")=0 and sPrintUOM = "Y" then
				set headerCell=oRow.insertCell()									
				headerCell.innerHTML=HeaderNode.getAttribute("UoMName")
				headerCell.className="ExcelDisplayCell"
			end if
		Next 'For each HeaderNode in ndItem.childNodes
	next
end Function

Function popClaDisplay()
	document.formname.selStore.options.length = 1
	Set Root = OutData.documentElement
	For Each HeaderNode In Root.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.hasChildNodes() then
				if not HeaderNode.Attributes.Item(1).nodeValue = "0" then
					document.formname.selStore.length = document.formname.selStore.length+1
					document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)&" -- "&trim(HeaderNode.Attributes.Item(3).nodeValue)
					document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)&":"&trim(HeaderNode.Attributes.Item(1).nodeValue)
				else
					document.formname.selStore.length = document.formname.selStore.length+1
					document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)
					document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)
				end if
			end if
		end if
		if StrComp(Trim(HeaderNode.NodeName),"UOM") = 0 then
			idUoM.innerHTML = trim(HeaderNode.Attributes.Item(1).nodeValue) & "&nbsp;"
		end if
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

	For Each ndItem In Root.childNodes
		For Each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
				sLoc = HeaderONode.getAttribute("LOC")
				sBin = HeaderONode.getAttribute("BIN")
				For Each PickNode In HeaderONode.childNodes
				    if PickNode.nodeName="PICK" then
					    i = i + 1
					    set objQ = eval("document.formname.txtQtyA"&i)
					    if trim(objQ.value) = "" then
						    msgbox "Enter Quantity",0,"Quantity"
						    objQ.select()
						    exit function
					    else
						    iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
					    end if
					    
					    set obj = eval("document.formname.txtJusA"&i)
					    
					    'alert(obj.value)
					    if cdbl(obj.value)>0 then
					        set objReason = eval("document.formname.txtReasonA"&i)
					        if trim(objReason.value)="" then
						        MsgBox "Enter the Reason",0,"Reason"
						        objReason.focus()
						        exit function
					        end if
					    end if
					end if 'if PickNode.nodeName="PICK" then
				next
			end if
		Next 'For Each HeaderONode in ndItem.childNodes
	Next

	i = 1	
	Set RootO = OutData.documentElement
	For Each ndItem In RootO.childNodes
		For each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
				For Each PickNode In HeaderONode.childNodes
				    if PickNode.nodeName="PICK" then
					    ndItem.setAttribute "Reason",eval("document.formname.txtReasonA"&i).value
    					
					    set objQ = eval("document.formname.txtQtyA"&i)
					    PickNode.setAttribute "QTYISS", cdbl(objQ.value)
    						
					    set obj = eval("document.formname.txtJusA"&i)
    						
					    set objS = eval("document.formname.txtStk"&cstr(i))
					    obj.value = cdbl(objQ.value) - cdbl(objS.value)

					    PickNode.setAttribute "ADJUSTED", cdbl(objS.value) - cdbl(objQ.value)
					    
					    nTotQtyIssued = nTotQtyIssued +(cdbl(objS.value) - cdbl(objQ.value))
		                nTotQtyRcpt = nTotQtyRcpt +(cdbl(objS.value) - cdbl(objQ.value))
					    i = i + 1
					end if
				next
			end if
		Next
	next
'alert(nTotQtyIssued)
'	sorgID = trim(document.formname.hOrgID.value)
'	iClass = trim(document.formname.hClass.value)
	
'	RootO.setAttribute "ITEM", trim(document.formname.selItem.value)
'	RootO.setAttribute "CLASS", iClass
'	RootO.setAttribute "ORG", sorgID
'	RootO.setAttribute "REASON", trim(document.formname.txtReason.value)
'	
'	alert(OutData.xml)
'	exit function
	document.formname.B7.disabled = True
	
	'alert(OutData.xml)
	
	
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
	
	set ndRcptDetails = IntReceipt.createElement("Details")
	ndRcptRoot.appendChild ndRcptDetails
	
	IssueFlag = false
	ReceiptFlag  = false
	sExp = "//Item"
	set TempNode = OutDataRoot.selectNodes(sExp)
	If TempNode.length <> "0" Then
		For i = 1 to TempNode.length
		sItemRcptNum = TempNode.item(i-1).Attributes.getNamedItem("RecNumStatus").value
		
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
	            ndIssueItem.setAttribute "RETURNABLE","1"
	            ndIssueItem.setAttribute "RefNo",""
	            ndIssueItem.setAttribute "ONLYLOT",""
	            ndIssueItem.setAttribute "RETURNITEM","S"
	            ndIssueItem.setAttribute "MatType",""
	            ndIssueRoot.appendChild ndIssueItem 
	            
	            
            set ndRcptItem = IntReceipt.createElement("ItemDetail")
			    ndRcptItem.setAttribute "ItemCode",TempNode.item(i-1).Attributes.getNamedItem("ICode").value
			    ndRcptItem.setAttribute "CLACODE",TempNode.item(i-1).Attributes.getNamedItem("CCode").value
			    ndRcptItem.setAttribute "QTY",nTotQtyRcpt
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
		
			nTotQtyIssued = "0"
			nTotQtyRcpt = "0"
			
			For each ndItem in OutDataRoot.childNodes
			    if ndItem.nodeName="Item" then
			        For each ndLocDet in nditem.childNodes
			            if ndLocDet.nodeName="LOCDET" then
			                nLocNo = ndLocDet.getAttribute("LOC")
			                nBinNo = ndLocDet.getAttribute("BIN")
			                iRate = ndLocDet.getAttribute("AVGRATE")
			                
			                
			                
			                For each ndPick in ndLocDet.childNodes
			                    if ndPick.nodeName="PICK" then
			                        nQtyStk = ndPick.getAttribute("QTYSTK")
			                        nQtyIss = ndPick.getAttribute("QTYISS")
			                        nQtyAdj = ndPick.getAttribute("ADJUSTED")
			                        nLotNo = ndPick.getAttribute("LOTNO")
			                        nInvNo = ndPick.getAttribute("INVRECNO")
			                        
			                        if cdbl(nQtyStk)-cdbl(nQtyIss)>0 then
			                            nTotQtyIssued = cdbl(nTotQtyIssued)+(cdbl(nQtyStk)-cdbl(nQtyIss))
			                            
			                            ndIssueItem.setAttribute "IVALUE",nTotQtyIssued*iRate
			                            
			                            set ndIssuePick = IssueData.createElement("Pick")
                                            ndIssuePick.setAttribute "TOT",(cdbl(nQtyStk)-cdbl(nQtyIss))
                                            ndIssuePick.setAttribute "NoofPack",""
                                            ndIssueItem.appendChild ndIssuePick
                                            
                                            If cdbl(nQtyAdj) <> 0 Then
                                                if trim(nLotNo)="" or trim(nLotNo)="NULL" then
                                                    nLotNo = "N/A"
                                                end if
                                                if trim(sItemRcptNum)="N" then
			                                        set ndIssueStore = IssueData.createElement("STORE")
			                                        ndIssueStore.setAttribute "LOC",nLocNo
			                                        ndIssueStore.setAttribute "BIN",nBinNo 
			                                        ndIssueStore.setAttribute "LOTNO",nLotNo
			                                        ndIssueStore.setAttribute "INVRECNO",nInvNo
			                                        ndIssueStore.setAttribute "QTYISS",cdbl(nQtyAdj)
			                                        ndIssueStore.setAttribute "NoofPack",""
			                                        ndIssuePick.appendChild ndIssueStore
			                                    else
                    							
			                                        set ndIssueStore = IssueData.createElement("PICK")
			                                        ndIssueStore.setAttribute "LOC",nLocNo
			                                        ndIssueStore.setAttribute "BIN",nBinNo 
			                                        ndIssueStore.setAttribute "LOTNO",nLotNo
			                                        ndIssueStore.setAttribute "INVRECNO",nInvNo
			                                        ndIssueStore.setAttribute "QTYISS",cdbl(nQtyAdj)
			                                        ndIssueStore.setAttribute "NoofPack",""
			                                        ndIssuePick.appendChild ndIssueStore
    			                                    
			                                        set ndIssSerHdr = IssueData.createElement("SERIALHEADER")
	                                                    ndIssueStore.appendChild ndIssSerHdr
    	                                                
			                                        For each ndSer in ndPick.childNodes
		                                                set ndIssSerDet = IssueData.createElement("SERIALDETAILS")
                                                            ndIssSerDet.setAttribute "SERIALNO",ndSer.getAttribute("SERIALNO")
                                                            ndIssSerDet.setAttribute "QTY",cdbl(ndSer.getAttribute("QTY"))*-1
                                                        ndIssSerHdr.appendChild ndIssSerDet
			                                        next
	                                            End If 'if trim(sItemRcptNum)="N" then
	                                        end if 'If cdbl(nQtyAdj) <> 0 Then
	                                end if ' if cdbl(nQtyStk)-cdbl(nQtyIss)>0 then
			                    elseif ndPick.nodeName="LotSerial" then
			                        iLotQtyStk = ndPick.getAttribute("QTY")
			                        iLotQtyVal = ndPick.getAttribute("IVALUE")
			                        iLotQtyTare = ndPick.getAttribute("TARE")
			                        iLotQty = cdbl(iLotQtyStk)-cdbl(iLotQtyTare)
			                        nTotQtyRcpt = cdbl(nTotQtyRcpt)+(cdbl(iLotQty))
                                    if cdbl(nQtyIss)<>0 then
                                        bFlag = false
                                        if IntReceipt.hasChildNodes() then
                                            for each ndIntDet in IntReceipt.childNodes
                                                for each ndIntItem in ndIntDet.childNodes
                                                    for each ndStorage in ndIntItem.childNodes
                                                        if ndStorage.nodeName = "STORAGE" and trim(ndStorage.getAttribute("STORE"))=trim(nLocNo) then
                                                            set ndRcptStorage = ndStorage
                                                            bFlag = true 
                                                        end if
                                                    next
                                                next
                                            next
                                        end if 
                                        
                                        if bFlag = false then
		                                    set ndRcptStorage = IntReceipt.createElement("STORAGE")
		                                    ndRcptItem.appendChild ndRcptStorage
		                                end if 
		                                
		                                ndRcptStorage.setAttribute "STORE",nLocNo
	                                    if trim(nBinNo)="0" then
	                                        ndRcptStorage.setAttribute "BIN","NULL"
	                                    else
	                                        ndRcptStorage.setAttribute "BIN",nBinNo
	                                    end if
	                                    ndRcptStorage.setAttribute "APPLICABLE","IN"
	                                    ndRcptStorage.setAttribute "MONTHYEAR",Date()
	                                    ndRcptStorage.setAttribute "QTY",iLotQty
	                                    ndRcptStorage.setAttribute "STORAGEVALUE",iLotQtyVal
                                        
		                                if ndPick.hasChildNodes() then
		                                    ndRcptStorage.appendChild(ndPick)
		                                end if
                                    end if
			                    end if
			                Next
			            end if
			        Next
			    end if
			Next
			if trim(sItemRcptNum)<>"N" then
			    ndIssueItem.setAttribute "ISSQTY",nTotQtyIssued
			    ndRcptItem.setAttribute "QTY",nTotQtyRcpt
			end if
		Next
	End IF
	
	set objhttp2 = CreateObject("Microsoft.XMLHTTP")
	objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=PhyAdjData",false
	objhttp2.send OutData.xmlDocument
	
    set objhttp2 = CreateObject("Microsoft.XMLHTTP")
    objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=mrsIssueData",false
    objhttp2.send IssueData.xmlDocument
    
    if cdbl(nTotQtyRcpt)<0 then nTotQtyRcpt = cdbl(nTotQtyRcpt)*-1
	if cdbl(nTotQtyRcpt)>0 then
        set objhttp2 = CreateObject("Microsoft.XMLHTTP")
        objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=ReceiptLotData",false
        objhttp2.send IntReceipt.xmlDocument
    end if
	
	'exit function
	
	document.formname.action = "stkMgmtPAInsert.asp"
	document.formname.submit
	
end Function

</script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp,sReceiptNum,sQuery
	
	set rsTemp = server.CreateObject("ADODB.Recordset")
	
	sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	sOrgID = trim(session("organizationcode"))
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
	
	sQuery = "Select ReceiptNumbering from Inv_M_ItemMaster where ItemCode = " & mid(sTemp,1,len(sTemp)-1) 
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        sReceiptNum = rsTemp(0)
    end if
    rsTemp.close
    	
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
	
	arrTempName = split(mid(sTempName,1,len(sTempName)-1),"|")
	arrTemp = split(mid(sTemp,1,len(sTemp)-1),"|")
	
	
	
%>
<body onload="FnInit('<%= mid(sTemp,1,len(sTemp)-1)%>')">
<form method="post" name="formname" action="">
<input type="hidden" name="hOrgID" value="<%=sOrgID%>"/>
<input type="hidden" name="hClass" value="<%=iClass%>"/>
<input type="hidden" name="hClassName" value="<%=sClassName%>"/>
<input type="hidden" name="hItem" value="<%= mid(sTemp,1,len(sTemp)-1)%>"/>
<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>"/>
<input type="hidden" name="hRcptNum" value="<%=sReceiptNum%>" />
<input type=hidden name="hUserID" value="<%=Session("userID")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="PageTitle" height="20"><p align="center">Physical Adjustment</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
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
				<tr>
					<td class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5" alt=""/>
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="frm2" style="width: 700; height:390;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Item Name</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Store</td>
                                                <td class="ExcelHeaderCell" align="center" colspan="2">Existing Stock Information</td>
                                                <td class="ExcelHeaderCell" align="center" colspan="3">Physical Adjustment</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Reason</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">UoM</td>
                                            </tr>
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center">Lot No.[Attributes]</td>
                                                <td class="ExcelHeaderCell" align="center">Quantity</td>
                                                <td class="ExcelHeaderCell" align="center">Stock</td>
                                                <td class="ExcelHeaderCell" align="center">Serial</td>
                                                <td class="ExcelHeaderCell" align="center">Adjusted</td>
                                            </tr>
                                        </table>
                                    </div>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5" alt=""/>
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Save" name="B7" class="ActionButton" onclick="CheckSubmit()"/>
                                                    <input type="reset" value="Reset" name="B1" class="ActionButton"/>
                                                    <input type="button" value="Cancel" name="B1" class="ActionButton" onclick="Cancel('stkMgmtEntry.asp')"/>
                                                </p>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5" alt=""/>
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5" alt=""/>
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
</body>
</html>
