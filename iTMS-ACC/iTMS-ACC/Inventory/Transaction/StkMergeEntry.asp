<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	StkMergeEntry.asp
	'Module Name				:	Inventory (Stock Management Status Management)
	'Author Name				:	Ragavendran R
	'Created On					:	May 25, 2011
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
<HTML><HEAD><TITLE>Stock Management - Stock Merge</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData1">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemSelectData">
    <Root />
</script>
<script type="application/xml" data-itms-xml-island="1" id="LotData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempItemData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTypeData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="IssueData"><ISSTYPE></ISSTYPE></script>
<script type="application/xml" data-itms-xml-island="1" id="IntReceipt"><ROOT></ROOT></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script language="javascript" src="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,OutDataValue
dim sorgID,iClass,sStore,iInvRec,sLot,sBin
j = 0
Dim ndIssueRoot,ndIssueItem,ndIssuePick,ndRcptRoot,ndRcptDet,ndRcptItem,ndRcptStore
Function FnInit(sItemCode)
	'document.formname.selItem.value = sItemCode
	GetXML()
	DisplayDetails()
End Function
'**************************************************
Function ViewLotDetails(iCnt)
   iItemCode = eval("document.formname.hFromItemCode"&cstr(iCnt)).value
   
   showModalDialog "ViewLotDetailsPop.asp?ItemCode="&iItemCode,"","dialogWidth:450px;dialogHeight:450px;Status:No;"
End Function
'****************************************************
Function GetLotDetails(iCnt,sItemCode,sClassCode,iLocNo,iBinNo)

    if eval("document.formname.hMergeItemCode").value = "" then
        alert("Select Item Merged With")
        exit function
    end if

    sParam1 = sItemCode&":"&sClassCode&":"&sItemName&":"& iLocNo &":"& iBinNo &"::::"&iCnt
    'alert(sParam1)
	set OutValue =  showModalDialog("ItemMergePickPop.asp?sTemp="&sParam1,TempItemData,"dialogHeight:390px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
	set ndRoot = TempItemData.documentElement
	if ndRoot.getAttribute("DONE")="YES" then
		sExp = "//Item[@ITEMCODE="& sItemCode &" and @CLASSCODE="& sClassCode &" and @LOCNO="& iLocNo &" and @BINNO = "& iBinNo &"]/Pick"
		set TempNode = ndRoot.selectNodes(sExp)
		if TempNode.length> 0 then
			eval("document.formname.hItemStock"&Cstr(iCnt)).value=TempNode.Item(0).Attributes.getNamedItem("TOT").Value
		end if
	end if
   CalculateVal(iCnt)
End Function
'******************************************************
Function SelectItem()

		sorgID = document.formname.hOrgID.value
		sIType = eval("document.formname.hFromItemType").value
		
		nFlag = 1
		iStock = "Y"

    iTotalQty = 0
     sTempValWindowSize = GetWindowSizeForPopup("1")
		sArrTempValWindowSize = split(sTempValWindowSize,":")
		sProgramName = sArrTempValWindowSize(0)
		sPopupHeight = sArrTempValWindowSize(1)
		sPopupWidth = sArrTempValWindowSize(2)

		set OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=S&Flag="+cstr(nFlag)&"&hDispButt="&bAddButton&"&hDispItem="&sDispItem&"&CallFrom="&sCallFrom,ItemSelectData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))

		set ndRoot = ItemSelectData.documentElement
		if ndRoot.hasChildNodes() then
		    for each ndChild in ndRoot.childNodes
		        'alert(ndChild.xml)
		        sItemName = ndChild.getAttribute("ItemName")
		        sItemCode = ndChild.getAttribute("ItemCode")
		        sClassCode = ndChild.getAttribute("ClassCode")
		        sReceiptNum = ndChild.getAttribute("ReceiptNum")
		        sAttributeList = ndChild.getAttribute("AttributeList")
		    next
		    
		    if trim(sAttributeList)<>"" then
		        sArrTemp = split(sAttributeList,"#")
		        if ubound(sArrTemp)=1 then
		            sArrTemp1 = split(sArrTemp(1),":")
		            sAttID = sArrTemp1(0)
		        end if
		    else
		        sAttID =""
		    end if
		    
		    document.formname.hReceiptNum.value = sReceiptNum
		    
		    set objhttp = CreateObject("MSXML2.XMLHTTP")
	        objhttp.Open "GET","ItmMergeStockDet.asp?ItemCode="&sItemCode&"&ClassCode="&sClassCode, false
	        objhttp.send
    		
		    if trim(objhttp.responseXML.xml)<>"" then
		        TempData.loadXml(objhttp.responseXML.xml)
		    else
		        alert(objhttp.responseText)
		    end if
    		
		    set ndRoot = TempData.documentElement
		    'alert(ndRoot.xml)
		    if ndRoot.hasChildNodes() then
		        for each ndChild in ndRoot.childNodes
		            if ndChild.nodeName="LOCDET" then
						sStorageCode = ndChild.getAttribute("LOC")
						sStorageBin = ndChild.getAttribute("BIN")
		                sMergeLocName = ndChild.getAttribute("LOCNAME")
		                for each ndPick in ndChild.childNodes
		                    if ndPick.nodeName="PICK" then
		                        iMergeStock = ndPick.getAttribute("QTYSTK")
		                    end if
    		            next 'for each ndPick in ndChild.childNodes
		            end if 
		        next
		        For iCnt = 1 to Cint(document.formname.hItemRow.value)
		            iPrevQty =cdbl(iPrevQty)+eval("document.formname.hItemStock"&cstr(iCnt)).value
		        Next
		        
		        
		        iTotalQty = cdbl(iPrevQty) + cdbl(iMergeStock)
		        eval("spaMergeItem").innertext = sItemName
                eval("document.formname.hMergeItemCode").value=sItemCode
                eval("document.formname.hMergeClassCode").value=sClassCode
                eval("document.formname.hLocCode").value = sStorageCode
                eval("document.formname.hBinNo").value = sStorageBin
                eval("document.formname.hAttID").value = sAttID
                eval("spaMergeStore").innerText = sMergeLocName
                eval("spaMergeStock").innerText = iMergeStock
                eval("spaMergedQty").innerText = iTotalQty
		    end if
		end if
		
		
		'alert(ndRcptRoot.xml)
End Function
'***********************************************************
Function GetXML()
	clearXML

	'set obj = document.formname.selItem
	Dim	sTemp
	
	sorgID = trim(document.formname.hOrgID.value)
	iClass = trim(document.formname.hClass.value)
	iItem = Trim(document.formname.hItemCode.value)
	
	sTemp = sorgID & ":" & iClass & ":" & iItem
	
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","itmStatusXMLSelect.asp", false
	objhttp.send
	
	'alert(objhttp.responseText)
	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
	else
		clearXML
	end if
	
End Function

Function DisplayDetails()
	Dim sRcptNumbering
	ClearTable
	j = 0
	
	blnPrintUOM = "N"
	iEntryNo = 0
	Set Root = OutData.documentElement
	Set ndTempItemRoot = TempItemData.documentElement
	'alert(Root.xml)
	if Root.hasChildNodes() then
	    for each ndChild in Root.childNodes
	        if ndChild.nodeName="Item" then
	            iEntryNo = iEntryNo + 1
	            sItemName = ndChild.getAttribute("IName")
	            iFromICode = ndChild.getAttribute("ICode")
	            iFromCCode = ndChild.getAttribute("CCode")
	            
	            
	            set objhttp = CreateObject("Microsoft.XMLHTTP")
	            objhttp.open "GET","../../include/GetItemTypeForItem.asp?ItemCode="&iFromICode,false
	            objhttp.send
	            if trim(objhttp.responseXML.xml)<>"" then
	                ItemTypeData.loadXML(objhttp.responseXML.xml)
	                set ndRoot = ItemTypeData.documentElement
	                sFromIType = ndRoot.getAttribute("ItemType")
	                document.formname.hFromItemType.value = sFromIType
	            end if
	            iRowspan = 0
	            
	                iLocCount = 0
	                For each ndLoc in ndChild.childNodes
                        if ndLoc.nodeName="LOCDET" then
                            iLocCount = iLocCount + 1
  				            iTotStock = 0
  				            iLocNo = ndLoc.getAttribute("LOC")
  				            iBinNo = ndLoc.getAttribute("BIN")
  				            sLocName = ndLoc.getAttribute("LOCNAME")
					        sRcptNumbering = trim(ndLoc.getAttribute("RECNUM"))
					        document.formname.hFromRcptNum.value = sRcptNumbering
					        iValue = ndLoc.getAttribute("VALUE")
					        
					        j=j+1
					        
					        Set ndTempItmDet = TempItemData.createElement("Item")
					            ndTempItmDet.setAttribute "ENTRYNO",j
	                            ndTempItmDet.setAttribute "ITEMCODE",iFromICode
	                            ndTempItmDet.setAttribute "CLASSCODE",iFromCCode
	                            ndTempItmDet.setAttribute "LOCNO",iLocNo
	                            ndTempItmDet.setAttribute "BINNO",iBinNo
	                            ndTempItmDet.setAttribute "ITEMNAME",sItemName
					            ndTempItemRoot.appendChild ndTempItmDet
	                
	                	    set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length)

						    set headerCell=oRow.insertCell()
						    headerCell.innerHTML=j
						    headerCell.className="ExcelSerial"
						    headerCell.align="center"
						    
						    set headerCell=oRow.insertCell()
							    headerCell.innerText = sItemName
							    headerCell.className="ExcelDisplayCell"
							    					           
					                
					                
					                For each ndPick in ndLoc.childNodes
					                    if ndPick.nodeName="PICK" then
					                        iStock = ndPick.getAttribute("QTYSTK")
					                        iTotStock = cdbl(iTotStock)+ cdbl(iStock)
					                    end if
					                Next
					                
					                set headerCell=oRow.insertCell()
							            headerCell.innerText = sLocName
							            set oText = document.createElement("<input type=""hidden"" name=""hItemEntryNo"&CStr(j)&""" value="""& iEntryNo &""" size=""30"">")
                   						headerCell.appendChild(oText)
							            set oText = document.createElement("<input type=""hidden"" name=""hFromItemCode"&CStr(j)&""" value="""& iFromICode &""" size=""30"">")
                   						headerCell.appendChild(oText)
                   						set oText = document.createElement("<input type=""hidden"" name=""hFromClassCode"&CStr(j)&""" value="""& iFromCCode &""" size=""30"">")
                   						headerCell.appendChild(oText)
                   						set oText = document.createElement("<input type=""hidden"" name=""hFromItemName"&CStr(j)&""" value="""& sItemName &""" size=""30"">")
                   						headerCell.appendChild(oText)
                   						set oText = document.createElement("<input type=""hidden"" name=""hFromLocNo"&CStr(j)&""" value="""& iLocNo &""" size=""30"">")
                   						headerCell.appendChild(oText)
                   						set oText = document.createElement("<input type=""hidden"" name=""hFromBinNo"&CStr(j)&""" value="""& iBinNo &""" size=""30"">")
                   						headerCell.appendChild(oText)
                   						set oText = document.createElement("<input type=""hidden"" name=""hFromItemValue"&CStr(j)&""" value="""& iValue &""" size=""30"">")
                   						headerCell.appendChild(oText)
							            headerCell.className="ExcelDisplayCell"
            					    
					                set headerCell=oRow.insertCell()
							            headerCell.innerHtml = "<a href='#' class='ExcelDisplayLink' onClick='ViewLotDetails("&CStr(j)&")' >"&iTotStock&"</a>"
							            headerCell.className="ExcelDisplayCell"
							            headercell.align="right"
							            
            						set headerCell=oRow.insertCell()
						            if Trim(sRcptNumbering)="LS" or trim(sRcptNumbering)="S" then
						                headerCell.innerHtml = "<a href='#' class='ExcelDisplayLink' onClick=GetLotDetails("&CStr(j)&","& iFromICode &","& iFromCCode &","& iLocNo &","& iBinNo &")><img src='../../assets/images/iTMS%20Icons/Entryicon.gif' border=0 style='cursor:hand'></a>"
							            set oText = document.createElement("<input type=""text"" size=8 style='text-align=right' class=FormElemRead name=""hItemStock"&CStr(j)&""" value=""0"" readonly>")
							        else
							            set oText = document.createElement("<input type=""text"" size=8 style='text-align=right' class=FormElem name=""hItemStock"&CStr(j)&""" value="""& iTotStock &""" onChange=CalculateVal("""&Cstr(j)&""")>")
							        end if
                   						headerCell.appendChild(oText)
							            headerCell.className="ExcelDisplayCell"
							            headercell.align="right"
							            
					        end if 'if ndLoc.nodeName="LOCDET" then
					Next 'For each ndLoc in ndChild.childNodes
		    end if 'if ndChild.nodeName="Item" then
		next ' for each ndChild in Root.childNodes				


		            set oRow = document.all.tblToItemData.insertRow(document.all.tblToItemData.rows.length)
			                
	                set headerCell=oRow.insertCell()
		                headerCell.innerHTML="1"
		                headerCell.className="ExcelSerial"
		                headerCell.align="center"
		                
						
		            set headerCell=oRow.insertCell()
		                headerCell.innerHtml = "<span id=spaMergeItem></span>"
		                headerCell.innerHtml = headerCell.innerHtml & "<img border='0' src='../../assets/images/iTMS%20icons/Entryicon.gif' width='12' height='12' onClick='SelectItem()'  >"
   						set oText = document.createElement("<input type=""hidden"" name=""hMergeItemCode"" size=""30"">")
   						headerCell.appendChild(oText)
   						set oText = document.createElement("<input type=""hidden"" name=""hMergeClassCode"" size=""30"">")
   						headerCell.appendChild(oText)
   						set oText = document.createElement("<input type=""hidden"" name=""hLocCode"" size=""30"">")
   						headerCell.appendChild(oText)
   						set oText = document.createElement("<input type=""hidden"" name=""hBinNo"" size=""30"">")
   						headerCell.appendChild(oText)
   						set oText = document.createElement("<input type=""hidden"" name=""hAttID"" size=""30"">")
   						headerCell.appendChild(oText)
   						headerCell.className="ExcelDisplayCell"
						
		            set headerCell=oRow.insertCell()
			            headerCell.innerHtml = "<span id=spaMergeStore></span>"
			            headerCell.className="ExcelDisplayCell"
										
		            set headerCell=oRow.insertCell()
		                headerCell.innerHtml = "<span id=spaMergeStock style='text-align:right;'></span>"
			            headerCell.className="ExcelDisplayCell"
			            headerCell.align="right"
		
		            set headerCell=oRow.insertCell()
		                headerCell.innerHtml = "<span id=spaMergedQty style='text-align:right;'></span>"
			            headerCell.className="ExcelDisplayCell"
			            headerCell.align="right"
		
	end if 'if Root.hasChildNodes() then
	
	document.formname.hItemRow.value = j
end Function
'****************
Function CalculateVal(nTempCtr)

	For nKK = 1 to cint(document.formname.hItemRow.value)
		iFromIQty = cdbl(iFromIQty) +cdbl(eval("document.formname.hItemStock"&Cstr(nKK)).value)
	Next
	
	iToIQty = eval("spaMergeStock").innerText
	if trim(iToIQty)="" then iToIQty = "0"
	iTotQty = cdbl(iToIQty)+Cdbl(iFromIQty)
	eval("spaMergedQty").innerText = iTotQty
	
End Function
'*****************

Function popClaDisplay()
	'document.formname.selStore.options.length = 1
	Set Root = OutData.documentElement
	For Each ndChild In Root.childNodes
		For each HeaderNode in ndChild.childNodes
			'alert(HeaderNode.NodeName)
			if Trim(HeaderNode.NodeName) = "LOCDET" then
				'if HeaderNode.hasChildNodes() then
				'alert HeaderNode.Attributes.Item(4).nodeValue
				document.formname.hRcptNumbering.value = HeaderNode.Attributes.Item(4).nodeValue  
					if not HeaderNode.Attributes.Item(1).nodeValue = "0" then
						document.formname.selStore.length = document.formname.selStore.length+1
						document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)&" -- "&trim(HeaderNode.Attributes.Item(3).nodeValue)
						document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)&":"&trim(HeaderNode.Attributes.Item(1).nodeValue)
					else
						document.formname.selStore.length = document.formname.selStore.length+1
						document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)
						document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)
					end if
				'end if
			end if
			if StrComp(Trim(HeaderNode.NodeName),"UOM") = 0 then
				idUoM.innerHTML = trim(HeaderNode.Attributes.Item(1).nodeValue) & "&nbsp;"
			end if
		next
	next
end Function

Function clearXML()
	Set Root = OutData.documentElement
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			set a=Root.removeChild(HeaderNode)
		next
	end if
'	document.formname.selStore.options.length = 1
end Function

Function ClearTable()
	dim i
	for i=2 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(2)
	next
end Function

Function CheckSubmit()
	dim objQ,iQtyTot,iTempQty
	
	GenIssueXML()
	GenReceiptXML()
	
	Set ndItemRoot = TempItemData.documentElement
	
	nItemCnt = document.formname.hItemRow.value
	
	For iSCnt = 1 to cint(nItemCnt)
		iTotStoreValue = cdbl(iTotStoreValue) + cdbl(eval("document.formname.hFromItemValue"&Cstr(iSCnt)).value)
	Next
	
	For iCnt = 1 to cint(nItemCnt)
		iMergeQty = cdbl(iMergeQty) + cdbl(eval("document.formname.hItemStock"&Cstr(iCnt)).value)
		iItemCode  = eval("document.formname.hFromItemCode"&Cstr(iCnt)).value
		iClassCode = eval("document.formname.hFromClassCode"&Cstr(iCnt)).value
		sLocNo = eval("document.formname.hFromLocNo"&Cstr(iCnt)).value
		sBinNo = eval("document.formname.hFromBinNo"&Cstr(iCnt)).value
		
	    IF iCnt = 1 then	
	    
	        set objhttp = CreateObject("Microsoft.XMLHTTP")
	        objhttp.open "GET","../../Common/GetItemRcptNumbering.asp?ItemCode="&eval("document.formname.hMergeItemCode").value,false
	        objhttp.send
	        if trim(objhttp.responseText)<>"" then
	            sMerRcptNum = objhttp.responseText
	        end if
		    set ndRcptItem = IntReceipt.createElement("ItemDetail")
		        ndRcptItem.setAttribute "ItemCode",eval("document.formname.hMergeItemCode").value
		        ndRcptItem.setAttribute "CLACODE",eval("document.formname.hMergeClassCode").value
		        ndRcptItem.setAttribute "QTY",eval("spaMergedQty").innerText - eval("spaMergeStock").innerText
		        ndRcptItem.setAttribute "MRSNO","N"
		        ndRcptItem.setAttribute "ISSNO","N"
		        ndRcptItem.setAttribute "ENTRYNO",iCnt
		        ndRcptItem.setAttribute "UNIT",document.formname.hOrgID.value
		        ndRcptItem.setAttribute "ITEMNAME",eval("spaMergeItem").innerText
		        ndRcptItem.setAttribute "UOM",""
		        ndRcptItem.setAttribute "ATTRIBUTELIST",eval("document.formname.hAttID").value
		        ndRcptItem.setAttribute "RefNo",""
		        ndRcptItem.setAttribute "RefQty",""
		        ndRcptItem.setAttribute "RECEIPTNUM",sMerRcptNum
		        ndRcptItem.setAttribute "BYPRODUCT","P"
	        ndRcptDet.appendChild ndRcptItem

	        set ndRcptStore = IntReceipt.createElement("STORAGE")
		        ndRcptStore.setAttribute "STORE",eval("document.formname.hLocCode").value
		        if trim(eval("document.formname.hBinNo").value)="0" then
				    ndRcptStore.setAttribute "BIN","NULL"
				else
				    ndRcptStore.setAttribute "BIN",eval("document.formname.hBinNo").value
				end if
		        ndRcptStore.setAttribute "APPLICABLE","IN"
		        ndRcptStore.setAttribute "MONTHYEAR",date()
		        ndRcptStore.setAttribute "QTY",eval("spaMergedQty").innerText - eval("spaMergeStock").innerText
		        ndRcptStore.setAttribute "STORAGEVALUE",iTotStoreValue
	        ndRcptItem.appendChild ndRcptStore	        
		end if 'IF iCnt = 1 then
		
		'alert(ndItemRoot.xml)
		if ndItemRoot.hasChildNodes() then
			
			sExp = "//Item[@CLASSCODE="& trim(iClassCode) &" and @ITEMCODE="& trim(iItemCode) &" and @LOCNO = "& trim(sLocNo) &" and @BINNO = "& trim(sBinNo)&"]/Pick/PICK"
		    set ndTemp = ndItemRoot.selectnodes(sExp)
		    If ndTemp.length > 0 Then
				For iCtr = 1 to ndTemp.length
					iTotPickQty = 0
					nLotNo = ndTemp.item(iCtr-1).Attributes.getNamedItem("LOTNO").value
					'alert(document.formname.hReceiptNum.value)
					If trim(document.formname.hReceiptNum.value) = "LS" or trim(document.formname.hReceiptNum.value) = "L" Then
						sExp1 = "//Item[@CLASSCODE="& trim(iClassCode) &" and @ITEMCODE="& trim(iItemCode) &" and @LOCNO = "& trim(sLocNo) &" and @BINNO = "& trim(sBinNo)&"]/Pick/PICK[@LOTNO='"&nLotNo&"']/Selection[@YesNo='Y']"
					Elseif trim(document.formname.hReceiptNum.value) = "S" Then
						sExp1 = "//Item[@CLASSCODE="& trim(iClassCode) &" and @ITEMCODE="& trim(iItemCode) &" and @LOCNO = "& trim(sLocNo) &" and @BINNO = "& trim(sBinNo)&"]/Pick/PICK/Selection[@YesNo='Y']"
					End IF
					
					'alert(sexp1)
					set nCheckNode = ndItemRoot.selectNodes(sExp1)
					If nCheckNode.length > 0 Then
						iCounter = iCounter + 1
						For nKK = 1 to nCheckNode.length
							iTotPickQty = cdbl(iTotPickQty) + nCheckNode.item(nKK-1).Attributes.getNamedItem("Qty").value
						Next
						'alert("Welcome" & nLotNo)
						Set ndLotSerial =IntReceipt.createElement("LotSerial")
							ndLotSerial.setAttribute "QTYIN","N"
							ndLotSerial.setAttribute "TARE","0"
							ndLotSerial.setAttribute "LOT",nLotNo
							ndLotSerial.setAttribute "SERIALFROM",""
							ndLotSerial.setAttribute "SERIALTO",""
							ndLotSerial.setAttribute "TAREWEIGHT","U"
							ndLotSerial.setAttribute "IVALUE",eval("document.formname.hFromItemValue"&Cstr(iCnt)).value
							ndLotSerial.setAttribute "QTY",iTotPickQty
							ndLotSerial.setAttribute "COUNTER",iCounter
							ndLotSerial.setAttribute "STAGE","select"
							ndLotSerial.setAttribute "ALTGROSS","0" 
							ndLotSerial.setAttribute "ALTNETT","0"
							ndLotSerial.setAttribute "ALTUOM","select"
							ndLotSerial.setAttribute "AUTOGEN",""
							ndRcptStore.appendChild ndLotSerial
							
						    iLotSer = 0
						    If trim(document.formname.hReceiptNum.value) = "LS" or trim(document.formname.hReceiptNum.value) = "L" Then
								sExp2 = "//Item[@CLASSCODE="& trim(iClassCode) &" and @ITEMCODE="& trim(iItemCode) &" and @LOCNO = "& trim(sLocNo) &" and @BINNO = "& trim(sBinNo)&"]/Pick/PICK[@LOTNO='"&nLotNo&"']/Selection[@YesNo='Y']"
							Elseif trim(document.formname.hReceiptNum.value) = "S" Then
								sExp2 = "//Item[@CLASSCODE="& trim(iClassCode) &" and @ITEMCODE="& trim(iItemCode) &" and @LOCNO = "& trim(sLocNo) &" and @BINNO = "& trim(sBinNo)&"]/Pick/PICK/Selection[@YesNo='Y']" 
							End IF
							
	                        set ndTempSelection = ndItemRoot.selectnodes(sExp2)
	                        if ndTempSelection.length > 0 then
	                            For nLL = 1 to ndTempSelection.length
	                                iLotSer = iLotSer + 1
	                                iSerNo = ndTempSelection.Item(nLL-1).Attributes.getNamedItem("SerialNo").Value
	                                InvRecNo = ndTempSelection.Item(nLL-1).Attributes.getNamedItem("InvRecNo").Value
	                                set Objhttp = CreateObject("Microsoft.XMLHTTP")
	                                objhttp.open "GET","GetLotDetails.asp?SerialNo="&iSerNo&"&InvNo="&InvRecNo,false
	                                objhttp.send
	                                if trim(objhttp.responseXML.xml)<>"" then
	                                    LotData.loadXML(objhttp.responseXML.xml)
	                                else
	                                    alert(objhttp.responseText)
	                                end if
	                                
	                                set ndLotRoot = LotData.documentElement
	                                
	                                if ndLotRoot.hasChildNodes() then
	                                    For Each ndChildLot in ndLotRoot.childNodes
	                                        set ndLotSerDet =IntReceipt.createElement("LotSerialDetails") 
	                                            ndLotSerDet.setAttribute "LOTSERIAL",iSerNo
	                                            ndLotSerDet.setAttribute "QTYREC",ndTempSelection.Item(nLL-1).Attributes.getNamedItem("Qty").Value
	                                            ndLotSerDet.setAttribute "TAREREC","0"
	                                            ndLotSerDet.setAttribute "SELLINGTYPE",ndChildLot.getAttribute("SellNo")
	                                            ndLotSerDet.setAttribute "WEIGHTSTYPE",ndChildLot.getAttribute("WeightPerSellForm")
	                                            ndLotSerDet.setAttribute "PACKINGTYPE",ndChildLot.getAttribute("PackCode")
	                                            ndLotSerDet.setAttribute "LOT",nLotNo
	                                            ndLotSerDet.setAttribute "SELLINGFORM",ndChildLot.getAttribute("SellForm")
	                                            ndLotSerDet.setAttribute "PACKNUMBER",ndChildLot.getAttribute("PackNo")
	                                            ndLotSerDet.setAttribute "IVALUE","0"
	                                            ndLotSerDet.setAttribute "ATTRIBUTELIST",eval("document.formname.hAttID").value
	                                            ndLotSerDet.setAttribute "NOOFCONE","0" 
					                            ndLotSerDet.setAttribute "SUBLEVELID",""
					                            ndLotSerDet.setAttribute "SQ",""
	                                            ndLotSerial.appendChild ndLotSerDet
	                                    Next
	                                end if 'if ndLotRoot.hasChildNodes() then
	                            Next
	                        end if
					End IF
				Next
		    End	If 'If ndTemp.length > 0 Then
		End If
		
	Next
	
	For iCnt = 1 to cint(nItemCnt)
	
		iMergeQty = cdbl(iMergeQty) + cdbl(eval("document.formname.hItemStock"&Cstr(iCnt)).value)
		iItemCode  = eval("document.formname.hFromItemCode"&Cstr(iCnt)).value
		iClassCode = eval("document.formname.hFromClassCode"&Cstr(iCnt)).value
		sItemName = eval("document.formname.hFromItemName"&Cstr(iCnt)).value
		sLocNo = eval("document.formname.hFromLocNo"&Cstr(iCnt)).value
		sBinNo = eval("document.formname.hFromBinNo"&Cstr(iCnt)).value
		iStoreIssQty = eval("document.formname.hItemStock"&Cstr(iCnt)).value
		
		iIssueQty = cdbl(iIssueQty) + cdbl(eval("document.formname.hItemStock"&Cstr(iCnt)).value)
		if iTempItemCode <>iItemCode then
		    iItemEntryNo = iItemEntryNo + 1
		    set ndIssueItem = IssueData.createElement("ITEM")
	        ndIssueItem.setAttribute "ENTRYNO",iItemEntryNo
	        ndIssueItem.setAttribute "ITMCODE",iItemCode
	        ndIssueItem.setAttribute "CLACODE",iClassCode
	        ndIssueItem.setAttribute "ITMNAME",sItemName
	        ndIssueItem.setAttribute "SSTORE",""
	        ndIssueItem.setAttribute "REQQTY","0"
	        ndIssueItem.setAttribute "REQBY",""
	        ndIssueItem.setAttribute "REMARKS",""
	        ndIssueItem.setAttribute "ITEMTYPE",eval("document.formname.hFromItemType").value
	        ndIssueItem.setAttribute "ISSUEDATE",document.formname.hCreatedOn.value 
	        ndIssueItem.setAttribute "ISSQTY",iIssueQty
	        ndIssueItem.setAttribute "TRAQTY","0"
	        ndIssueItem.setAttribute "PRQTY","0"
	        ndIssueItem.setAttribute "IVALUE","0"
	        ndIssueItem.setAttribute "ORGCODE",document.formname.hOrgID.value
	        ndIssueItem.setAttribute "MRSNO",""
	        ndIssueItem.setAttribute "MRSDATE",""
	        ndIssueItem.setAttribute "ATTRIBUTELIST",""
	        ndIssueItem.setAttribute "CREATEDBY", document.formname.hUserID.value 
	        ndIssueItem.setAttribute "CREATEDON", document.formname.hCreatedOn.value 
	        ndIssueItem.setAttribute "RETURNABLE","1"
	        ndIssueItem.setAttribute "RefNo",""
	        ndIssueItem.setAttribute "ONLYLOT",""
	        ndIssueItem.setAttribute "RETURNABLE","N"
	        ndIssueItem.setAttribute "RETURNITEM","S"
	        ndIssueItem.setAttribute "MatType",""
	        ndIssueRoot.appendChild ndIssueItem 
	        iTempItemCode = iItemCode
	        
			if document.formname.hFromRcptNum.value = "N" then
			
			    set ndIssuePick = IssueData.createElement("Pick")
			        ndIssuePick.setAttribute "TOT",iIssueQty
			        ndIssuePick.setAttribute "NoofPack",""
			        ndIssueItem.appendChild ndIssuePick
			
			    set ndIssueStore = IssueData.createElement("STORE")
			    ndIssueStore.setAttribute "LOC",sLocNo
			    ndIssueStore.setAttribute "BIN",sBinNo
			    ndIssueStore.setAttribute "LOTNO","N/A"
			    ndIssueStore.setAttribute "INVRECNO",""
			    ndIssueStore.setAttribute "QTYISS",iStoreIssQty
			    ndIssueStore.setAttribute "NoofPack",""
			    ndIssuePick.appendChild ndIssueStore
			else
			    if ndItemRoot.hasChildNodes() then
	                sExp = "//Item[@CLASSCODE="& trim(iClassCode) &" and @ITEMCODE="& trim(iItemCode) &" and @LOCNO = "& trim(sLocNo) &" and @BINNO = "& trim(sBinNo)&"]/Pick"
	                set ndTemp = ndItemRoot.selectnodes(sExp)
	                if ndTemp.length > 0 then
	                    For iCtr = 1 to ndTemp.length
	                        ndIssueItem.appendChild ndTemp.Item(iCtr-1)
	                    Next
	                end if
	            end if
			end if 'if document.formname.hFromRcptNum.value = "N" then
			
	    else
	        ndIssueItem.setAttribute "ISSQTY",iIssueQty
	        ndIssuePick.setAttribute "TOT",iIssueQty
	        
	        if document.formname.hFromRcptNum.value = "N" then
	            set ndIssueStore = IssueData.createElement("STORE")
			    ndIssueStore.setAttribute "LOC",sLocNo
			    ndIssueStore.setAttribute "BIN",sBinNo
			    ndIssueStore.setAttribute "LOTNO","N/A"
			    ndIssueStore.setAttribute "INVRECNO",""
			    ndIssueStore.setAttribute "QTYISS",iStoreIssQty
			    ndIssueStore.setAttribute "NoofPack",""
			    ndIssuePick.appendChild ndIssueStore
			end if 'if document.formname.hFromRcptNum.value = "N" then
	    end if    
	Next
	
		set objhttp2 = CreateObject("Microsoft.XMLHTTP")
		objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=mrsIssueData",false
		objhttp2.send IssueData.xmlDocument
	
		set objhttp2 = CreateObject("Microsoft.XMLHTTP")
		objhttp2.open "POST","XMLSave.asp?SessionFlag=true&Name=ReceiptLotData",false
		objhttp2.send IntReceipt.xmlDocument
	'exit function
	document.formname.action = "StkMergeInsert.asp"
	document.formname.submit

end Function
'************************************************
Function GenReceiptXML()
	set ndRcptRoot = IntReceipt.documentElement
	ndRcptRoot.setAttribute "DEPT","OTH"
	ndRcptRoot.setAttribute "SOURCE","N"
	ndRcptRoot.setAttribute "ORGCODE",document.formname.hOrgID.value
	ndRcptRoot.setAttribute "STYPE","N"
	ndRcptRoot.setAttribute "ITEMTYPE",""
	ndRcptRoot.setAttribute "PACKNUM",""
	ndRcptRoot.setAttribute "SRCREFTYPE","N"
	ndRcptRoot.setAttribute "SRCREFNO",""
	ndRcptRoot.setAttribute "RCPTNUMBERINV",trim(document.formname.hReceiptNum.value)
	ndRcptRoot.setAttribute "sTypeRcpt",""
	ndRcptRoot.setAttribute "APPREFTYPE",""
	ndRcptRoot.setAttribute "APPREFNO",""
	ndRcptRoot.setAttribute "APPREFDATE",""
	ndRcptRoot.setAttribute "RCVDON",date()
	ndRcptRoot.setAttribute "AUTOACCOUNT","Y"
	
	set ndRcptDet = IntReceipt.createElement("Details")
	ndRcptRoot.appendChild ndRcptDet
End Function
'***********************************************
Function GenIssueXML()
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
End Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp,iCreatedBy,dCreatedOn
	
	set rsTemp = server.CreateObject("ADODB.Recordset")
	'sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	sOrgID = Session("organizationcode")
	iClass = trim(Request.Form("selClass"))
	sTemp  = trim(Request.Form("hSelectedValue"))
	sTempName = trim(Request.Form("hItemNames"))
	
	'Response.Write "<p>sTemp="&sTemp
	iCreatedBy = Session("userid")
	dCreatedOn = FormatDate(Date())
	
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
	
	arrTempName = split(mid(sTempName,1,len(sTempName)-1),"|")
	arrTemp = split(mid(sTemp,1,len(sTemp)-1),"|")
	
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit('<%= mid(sTemp,1,len(sTemp)-1)%>')">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hClass" value="<%=iClass%>">
<input type=hidden name="hItemCode" value="<%= mid(sTemp,1,len(sTemp)-1)%>">
<input type=hidden name="hItem" value="">
<input type=hidden name="hRcptNumbering" value="">
<input type=hidden name="hLoc" value="">
<input type=hidden name="hBin" value="">
<input type="hidden" name="hItemType" value="">
<input type=hidden name="hUserID" value="<%=iCreatedBy%>">
<input type=hidden name="hCreatedOn" value="<%=dCreatedOn%>">
<input type=hidden name="hItemRow" value="">
<input type=hidden name="hFromItemType" value="">
<input type=hidden name="hReceiptNum" value="">
<input type=hidden name="hFromRcptNum" value="">

<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Stock Merge
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<td height="20" valign="bottom">
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="frm2" style="width: 700; height:150;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" rowspan=2 >S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan=2>Item To Merge</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan=2>Store</td>
                                                <td class="ExcelHeaderCell" align="center" colspan=2>Quantity</td>
                                            </tr>
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" >Stock</td>
                                                <td class="ExcelHeaderCell" align="center" >Merge</td>
                                            </tr>
                                        </table>
                                    </div>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="Div1" style="width: 700; height:150;">
                                        <table border="0" cellspacing="1" id="tblToItemData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" >Item Merged With</td>
                                                <td class="ExcelHeaderCell" align="center" >Store</td>
                                                <td class="ExcelHeaderCell" align="center" >Stock</td>
                                                <td class="ExcelHeaderCell" align="center" >Merged Qty</td>
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
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                <input type="button" value="Save" name="B3" class="ActionButton" onClick="CheckSubmit()">
                                                <input type="reset" value="Reset" name="B1" class="ActionButton">
                                                <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="Cancel('stkMgmtEntry.asp')">
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
