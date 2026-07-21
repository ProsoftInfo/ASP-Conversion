<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	receiptInternalEntry.asp
	'Module Name				:	Inventory (Receipt Creation)
	'Author Name				:	KUMAR K A
	'Created On					:
	'Modified By				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
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
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!--#include file="../../include/CommonFunctions.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim iCtr,arrTemp,sTemp,sSource,sOrgID,sDept,iItem, iClass , sReceiptType, sReceiptName
	Dim dcrs, sSalesEli, sPurEli, sManEli, sEligible, arrUoM, sUoMCode, sUoMDesc, sCheck
	Dim sTempMonYr, sMonYr, arrFin, sFinFrom, sFinTo,sType,sItmType,sQuery,Arr1,sPassData
	Dim dtCurrDate,sAttributeList,sFinPeriod,Arr,sMinDate,sMaxDate,sAutoInternalRcptAccount
	Dim iRcptNo,sMode
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sOrgID = Session("organizationcode")

	iRcptNo = Request("RcptNo")
	if Trim(iRcptNo)<>"" then
	    sMode = "E"
	else
	    sMode = "N"
	end if

	sFinPeriod = session("FinPeriod")
	Arr = split(sFinPeriod,":")
	sMinDate = "01/04/"& Arr(0)
	sMaxDate = "31/03/"& Arr(1)


	iItem = trim(Request.Form("hItmCode"))

	sSource = "N"
	sTemp = trim(Request.Form("hSelectedValue"))


	if trim(sOrgID) = "" then
		If iSAApplicationPop <> "" then
			sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY OUDEFINITIONID"
		Else
			sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		End If

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			sOrgID = dcrs(0)
		end if
		dcrs.Close
	end if

	sTempMonYr = mid(FormatDate(date),4,2)
	sMonYr = sTempMonYr&Year(date())

	arrFin = split(Session("FinPeriod"),":")
	sFinFrom = "01/04/"& arrFin(0)
	sFinTo = "31/03/"& arrFin(1)

	if DateDiff("D",FormatDate(sFinTo),FormatDate(date)) > 0 then
		dtCurrDate = sFinTo
	else
		dtCurrDate = FormatDate(date)
	end if

	sQuery = "Select IsNull(AutoInternalRcptAccounting,'N') from INV_M_ApplicationSetup"
	dcrs.open sQuery,con
	if not dcrs.eof then
	    sAutoInternalRcptAccount =dcrs(0)
	else
	    sAutoInternalRcptAccount = "N"
	end if
	dcrs.close


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receipt Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!--<ITEM CLACODE="<%=trim(iClass)%>" ITMCODE="<%=trim(iItem)%>" QTY="" MRSNO="<%="N"%>" ISSNO="<%="N"%>" />-->
<script type="application/xml" data-itms-xml-island="1" id="OutData2">
	<ROOT DEPT="" SOURCE="<%=sSource%>" ORGCODE="<%=sOrgID%>" STYPE="" ITEMTYPE="" PACKNUM="" SRCREFTYPE="" SRCREFNO="" RCPTNUMBERINV="" APPREFTYPE="" APPREFNO="" APPREFDATE="">
	</ROOT>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData">
<root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="NewData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="StoreData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="RefXML"><Root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/newReceipt.js"></script>
<script LANGUAGE=javascript SRC="../scripts/receiptInternalEntry.js"></script>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function AddItem()
    nFlag=1
	iStock = "N"
	sorgID = document.formname.hOrgID.value

    sTempValWindowSize = GetWindowSizeForPopup("1")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)

	set OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt=Y&CallFrom=PUR",OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	sAct = UCase(trim(OutValue.getAttribute("Action")))
	sQuery = trim(OutValue.getAttribute("PassQuery"))

Set ndRoot = OutData.documentElement
Set Root = OutData2.documentElement

bFlag = false
if Root.hasChildNodes() then
    for each ndDet in Root.childNodes
        if ndDet.nodeName="Details" then
            bFlag = true
            set ndDetails = ndDet
        end if
    next
end if
if bFlag = flase then
    set ndDetails = OutData2.createElement("Details")
    Root.appendChild ndDetails
end if

    if ndRoot.hasChildNodes() then
        for each ndChild in ndRoot.childNodes
            if ndChild.nodeName="Item" then
                Set newElem = OutData2.createElement("ItemDetail")
	                newElem.setAttribute "ItemCode", ndChild.getAttribute("ItemCode")
                    newElem.setAttribute "CLACODE", ndChild.getAttribute("ClassCode")
                    newElem.setAttribute "QTY",""
                    newElem.setAttribute "MRSNO","N"
                    newElem.setAttribute "ISSNO","N"
                    newElem.setAttribute "ENTRYNO",ndChild.getAttribute("EntryNo")
                    newElem.setAttribute "UNIT", document.formname.hOrgID.value
                    newElem.setAttribute "ITEMNAME", ndChild.getAttribute("ItemName")
                    newElem.setAttribute "UOM", ndChild.getAttribute("StoresUoM")
                    if trim(ndChild.getAttribute("AttributeList"))="0" then
                        newElem.setAttribute "ATTRIBUTELIST",""
                    else
                        newElem.setAttribute "ATTRIBUTELIST",ndChild.getAttribute("AttributeList")
                    end if
                    newElem.setAttribute "RefNo",""
                    newElem.setAttribute "ReqQty","0"
                    newElem.setAttribute "RECEIPTNUM",ndChild.getAttribute("ReceiptNum")
                    newElem.setAttribute "BYPRODUCT","P"
                    if not CheckAvailability(ndChild.getAttribute("ItemCode")) then
                        ndDetails.appendChild newElem
                    end if
	        end if  'if ndChild.nodeName="Item" then
	    next 'for each ndChild in ndRoot.childNodes
    end if 'if ndRoot.hasChildNodes() then

    DisplayTable date
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.Open "POST","XMLSave.asp?Name=ReceiptLotData&SessionFlag=true", false	'
    objhttp.send OutData2.XMLDocument
End Function
'**********************************************
Function CheckAvailability(ItemCode)
    Set Root = OutData2.documentElement

    if Root.hasChildNodes() then
        for each ndItem in Root.childNodes
            if ndItem.nodeName="Details" then
                for each ndItemdet in ndItem.childNodes
                    if ndItemdet.getAttribute("ItemCode")=ItemCode then
                        CheckAvailability = true
                    end if
                next
            end if
        next
    end if
End Function
'**********************************************
'---------------------------------------------------------------------------------
Function Search()

	sUnit = document.formname.hOrgID.value



End Function
'*******************************************
Function DisplayItem(obj)
	sArrValues = Split(obj,"A")
	sTempValues = sArrValues(0)&"A"&sArrValues(2)&"A"&sArrValues(1)&"A"&sArrValues(3)
	showModalDialog "itmDetailsPop.asp?sTemp="&sTempValues,"","dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
end Function
'*********************************************
    Function CheckStorage(obj)
    Dim sArrTemp,sItemCode,sClassCode,iEntNo,sOrgCode,sAttList
    Dim ndRoot,OutVlaue,StoreRoot

        set ndRoot = OutData2.documentElement
        sArrTemp = split(obj.name,"Z")
        sClassCode = sArrTemp(1)
        sItemCode = sArrTemp(2)
        iEntNo = sArrTemp(3)
        sOrgCode = sArrTemp(4)
        sAttList = sArrTemp(5)

        OutValue = showModalDialog("StorageSelectForItemPop.asp?sUnit="&sOrgCode&"&ItemCode="&sItemCode,StoreData,"dialogHeight:200px;dialogWidth:300px;")
        set StoreRoot = StoreData.documentElement
        'alert(storeRoot.xml)
        if StoreRoot.hasChildNodes() then
            for each ndStoreChild in StoreRoot.childNodes
                if ndStoreChild.nodeName="STOREDET" then
                    if ndRoot.haschildNodes() then
                        for each ndDet in ndRoot.childNodes
                            if ndDet.nodeName="Details" then
                                for each ndItem in ndDet.childNodes
                                    if ndItem.nodeName="ItemDetail" then
                                        if ndItem.getAttribute("ItemCode")=sItemCode and ndItem.getAttribute("CLACODE")=sClassCode and ndItem.getAttribute("ENTRYNO") = iEntNo then
                                            sStoreName = ndStoreChild.getattribute("STORE")
                                            sArrUnitStore = split(ndStoreChild.getattribute("UNITSTORE"),"-")
                                            sStore = sArrUnitStore(0)
                                            sBin = sArrUnitStore(1)

                                            if trim(sBin)="" or isNull(sBin) or sBin="0" then sBin = "NULL"

                                            set ndStorage = OutData2.createElement("STORAGE")
                                            ndStorage.setAttribute "STORE",trim(sStore)
                                            ndStorage.setAttribute "BIN",trim(sBin)
                                            ndStorage.setAttribute "APPLICABLE","IN"
                                            ndStorage.setAttribute "MONTHYEAR",""
                                            ndStorage.setAttribute "QTY","0"
                                            ndStorage.setAttribute "STORAGEVALUE","0"
                                            ndStorage.setAttribute "SQ","0"
                                            ndItem.appendchild ndStorage

			                                eval("document.formname.txtStore"&cstr(sClassCode)&"A"&cstr(sItemCode)&"A"&cstr(iEntNo)).value = sStoreName
			                                StoreRoot.removeChild ndStoreChild
                                        end if
                                    end if
                                next
                            end if
                        next
                    end if
                end if
            next
        end if

        set objhttp = CreateObject("Microsoft.XMLHTTP")
	    objhttp.Open "POST","XMLSave.asp?Name=ReceiptLotData&SessionFlag=true", false	'
	    objhttp.send OutData2.XMLDocument

    End Function
'**************************************************
    Function GetDetails()
        Dim sRefType,sOrgID,sItemType,sPartyCode,iStock,nFlag,bAddButton
        sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
        sorgID = document.formname.hOrgID.value

        if document.formname.selDepart.selectedIndex = 0 then
            alert("Select Received From")
                document.formname.selDepart.focus()
            exit function
        end if

        sReceivedFrom = document.formname.selDepart(document.formname.selDepart.selectedIndex).value


    sDispItem = 0

    sCallFrom = "PUR"
    iStock = "N"

        RefTypeSelection sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem,sCallFrom
            Set ndRoot = OutData.documentElement
            Set Root = OutData2.documentElement

            bFlag = false
            if Root.hasChildNodes() then
                for each ndDet in Root.childNodes
                    if ndDet.nodeName="Details" then
                        bFlag = true
                        set ndDetails = ndDet
                    end if
                next
            end if
            if bFlag = flase then
                set ndDetails = OutData2.createElement("Details")
                Root.appendChild ndDetails
            end if
        if trim(sRefType)<>"N" then
            if ndRoot.hasChildNodes() then
                for each ndChild in ndRoot.childNodes
                    if ndChild.nodeName="Reference" then
                        sRefCode = trim(ndChild.getAttribute("ReferenceCode"))
                        sRefDate = sRefDate &","& ndChild.getAttribute("ReferenceDate")
                        sRefNo = sRefNo &","& ndChild.getAttribute("ReferenceNo")
                        sRefDt = ndChild.getAttribute("ReferenceDate")
                        sRefCodeDate =  sRefCodeDate &","& sRefCode &" - "& sRefDt
                    end if
                next

                sRefCodeDate = mid(sRefCodeDate,2)
                sRefNo = mid(sRefNo,2)
                sRefDate = mid(sRefDate,2)

                RefNoDate.innerHTML = sRefCodeDate
                document.formname.hRefNo.value = sRefNo
                document.formname.hRefDate.value = sRefDate

                    if sRefNo <>"" then

                       if trim(sRefType)="12" then
                         set objhttp = CreateObject("Microsoft.XMLHTTP")
                            objhttp.open "GET","GetIssItemReturnable.asp?RefCodes="&trim(sRefNo),false
                            objhttp.send
                            if trim(Objhttp.responseText)<>"" then
                                sArrReturnable = split(objhttp.responsetext,":")
                                'alert(ubound(sArrReturnable))
                                if Ubound(sArrReturnable)=1 then
                                    sRetunable = sArrReturnable(0)
                                    sRetunitem = sArrReturnable(1)
                                else
                                    alert(objhttp.responseText)
                                end if
                            end if
                        else
                            sRetunable = ""
                            sRetunitem = ""
                        end if'if trim(sRefType)="12" then

                       if sRetunable="Y" and sRetunitem="D" then
                            Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

                            sTempValWindowSize = GetWindowSizeForPopup("1")
                            sArrTempValWindowSize = split(sTempValWindowSize,":")
                            sProgramName = sArrTempValWindowSize(0)
                            sPopupHeight = sArrTempValWindowSize(1)
                            sPopupWidth = sArrTempValWindowSize(2)


	                        set ResData = showModalDialog("../../Common/"&sProgramName&"?orgID="& sOrgID &"&sIType=" & sIType & "&hSelectMode=M&Flag="+cstr(nFlag),ItemData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	                        sAct = UCase(trim(ResData.getAttribute("Action")))
	                        sQuery = trim(ResData.getAttribute("PassQuery"))
                       else
                            set objhttp = CreateObject("Microsoft.XMLHTTP")
                            'alert("RefType="&sRefType &"&RefCodes="&trim(sRefNo)&"&orgID="& document.formname.hOrgID.value)
                                objhttp.open "GET","InvGetItemDetForInternalReceipt.asp?RefType="&sRefType &"&RefCodes="&trim(sRefNo)&"&orgID="& document.formname.hOrgID.value,false
                                objhttp.send
                                 'alert(Objhttp.responseText)
                                if trim(objhttp.responseXML.xml)<>"" then
                                    ItemData.loadXML(objhttp.responseXML.xml)
                                else
                                    alert(Objhttp.responseText)
                                end if
                       end if ' end if'if sReturnable="Y" and sReturnItem="D" then

                           set ndIRoot = ItemData.documentElement
                           if ndIRoot.hasChildNodes() then
                                for each ndIChild in ndIRoot.childNodes
                                    if ndIChild.nodeName="Item" then
                                        if trim(document.formname.hItemType.value)="" then
                                            sItemCode = ndIChild.getAttribute("ItemCode")
                                        end if
                                        bItemFlag = false

                                        if not bItemFlag then
                                            Set newElem = OutData2.createElement("ItemDetail")
			                                    newElem.setAttribute "ItemCode", ndIChild.getAttribute("ItemCode")
			                                    newElem.setAttribute "CLACODE", ndIChild.getAttribute("ClassCode")
			                                    newElem.setAttribute "QTY",""
			                                    newElem.setAttribute "MRSNO","N"
			                                    newElem.setAttribute "ISSNO","N"
			                                    newElem.setAttribute "ENTRYNO",ndIChild.getAttribute("EntryNo")
			                                    newElem.setAttribute "UNIT", document.formname.hOrgID.value
			                                    newElem.setAttribute "ITEMNAME", ndIChild.getAttribute("ItemName")
			                                    newElem.setAttribute "UOM", ndIChild.getAttribute("StoresUoM")
			                                    if trim(ndIChild.getAttribute("AttributeList"))="0" then
			                                        newElem.setAttribute "ATTRIBUTELIST",""
			                                    else
			                                        newElem.setAttribute "ATTRIBUTELIST",ndIChild.getAttribute("AttributeList")
			                                    end if
			                                    newElem.setAttribute "RefNo",sRefNo
			                                    newElem.setAttribute "ReqQty",""'ndIChild.getAttribute("Qty")
			                                    newElem.setAttribute "RECEIPTNUM",ndIChild.getAttribute("ReceiptNum")
			                                    newElem.setAttribute "BYPRODUCT","P"
			                                    if not CheckAvailability(ndIChild.getAttribute("ItemCode")) then
                                                    ndDetails.appendChild newElem
                                                end if
			                            end if 'if not bItemFlag then
			                        end if  'if ndChild.nodeName="Item" then
                                next
                            end if

                    end if 'if sRefNo <>"" then
            end if
        else
            if ndRoot.hasChildNodes() then
                for each ndChild in ndRoot.childNodes

                    if ndChild.nodeName="Item" then
                        Set newElem = OutData2.createElement("ItemDetail")
			                newElem.setAttribute "ItemCode", ndChild.getAttribute("ItemCode")
                            newElem.setAttribute "CLACODE", ndChild.getAttribute("ClassCode")
                            newElem.setAttribute "QTY",""
                            newElem.setAttribute "MRSNO","N"
                            newElem.setAttribute "ISSNO","N"
                            newElem.setAttribute "ENTRYNO",ndChild.getAttribute("EntryNo")
                            newElem.setAttribute "UNIT", document.formname.hOrgID.value
                            newElem.setAttribute "ITEMNAME", ndChild.getAttribute("ItemName")
                            newElem.setAttribute "UOM", ndChild.getAttribute("StoresUoM")
                            if trim(ndChild.getAttribute("AttributeList"))="0" then
                                newElem.setAttribute "ATTRIBUTELIST",""
                            else
                                newElem.setAttribute "ATTRIBUTELIST",ndChild.getAttribute("AttributeList")
                            end if
                            newElem.setAttribute "RefNo",""
                            newElem.setAttribute "ReqQty","0"
                            newElem.setAttribute "RECEIPTNUM",ndChild.getAttribute("ReceiptNum")
                            newElem.setAttribute "BYPRODUCT","P"
			                if not CheckAvailability(ndChild.getAttribute("ItemCode")) then
                                ndDetails.appendChild newElem
                            end if
			        end if  'if ndChild.nodeName="Item" then
			    next 'for each ndChild in ndRoot.childNodes
            end if 'if ndRoot.hasChildNodes() then
        end if
        DisplayTable date
        set objhttp = CreateObject("Microsoft.XMLHTTP")
	    objhttp.Open "POST","XMLSave.asp?Name=ReceiptLotData&SessionFlag=true", false	'
	    objhttp.send OutData2.XMLDocument
    End Function
'***************************************
Function DisplayTable(todaysDate)
    Dim sRefType
	iCounter1 = 0

	set root = OutData2.DocumentElement

	IF document.formname.hCtr.value = "" then document.formname.hCtr.value = 0
	IF document.formname.hCtr.value <> 0 then 	ClearTable()

    sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value

	sUnit = document.formname.hOrgID.value
	sMode = document.formname.hMode.value


	if root.hasChildNodes then
		sExp ="//ItemDetail"
		Set CheckNode = Root.Selectnodes(sExp)
		if CheckNode.Length > 0 then
			iCounter = 0
			For iCounter = 0 to CheckNode.Length - 1
			'alert(CheckNode.Item(iCounter).xml)
				iEntNo = CheckNode.Item(iCounter).Attributes.getNamedItem("ENTRYNO").value
				iItem = CheckNode.Item(iCounter).Attributes.getNamedItem("ItemCode").value
				iClass = CheckNode.Item(iCounter).Attributes.getNamedItem("CLACODE").value
				sOrgId = CheckNode.Item(iCounter).Attributes.getNamedItem("UNIT").value
				sItemDesc = replace(CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").value,"~~",chr(34))
				sItemName = CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").value
				sAttributeList = CheckNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").value
			 	sUOM = CheckNode.Item(iCounter).Attributes.getNamedItem("UOM").value
			 	iRequested =CheckNode.Item(iCounter).Attributes.getNamedItem("ReqQty").value
			 	if Trim(sMode)="E" then
			 	    iQty = CheckNode.Item(iCounter).Attributes.getnamedItem("QTY").value
			 	    iItemRate = CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMRATE").value
			 	    if iItemRate<>"" then
			 	        sValue = cdbl(iQty)*cdbl(iItemRate)
			 	    else
			 	        sValue = cdbl(iQty)*1
			 	    end if
			 	end if
				sOptName=""
			 	sStoreCode = 0
				sBinCode = 0

				set objhttp = CreateObject("Microsoft.XMLHTTP")
				objhttp.open "GET","../../Common/GetItemRcptNumbering.asp?ItemCode="&iItem,false
				objhttp.send
				if trim(objhttp.responseText)<>"" then
				    sRcptNumbering = trim(objhttp.responseText)
				end if


	            AttTemp = split(sAttributeList,",")
                IF sAttributeList <> "" then
	                For k = 0 to UBOUND(AttTemp)
		                nTemp   = AttTemp(k)
		                nTempVal  = split(nTemp,":")
		                nVal  = nTempVal(0)
		                if uBound(nTempVal)=0 then
		                    set objhttp = CreateObject("Microsoft.XMLHTTP")
		                    objhttp.open "GET","../../include/GetAttrName.asp?AttID="&nVal,false
		                    objhttp.send
		                    if trim(objhttp.responseText)<>"" then
		                        sOptName = objhttp.responseText
		                    end if
		                else
		                    nValTemp = split(nVal,"#")
		                    IF trim(nValTemp(1)) = "0" then
		                        if Trim(sOptName)="" then
			                        sOptName = ""
			                    end if
		                    Else
		                        sOptName = sOptName &"["& nTempVal(1)&"]"
		                    End IF
		                end if
	                Next
	            End If

	            if trim(sOptName)<>"" then
	                sItemDesc = sItemDesc &" [ "& sOptName &" ]"
	            end if

	         	'To Get Stock Details

				sTemp = iItem &":"& iEntNo &":"& iClass &":"& document.formname.hOrgID.value&":"&document.formname.hMinDate.value&":"&document.formname.hMaxDate.value&":"&Replace(Replace(sAttributeList,":","@"),"#","$")

				objhttp.Open "GET","XMLGetStockDetails.asp?Para="&sTemp,false
				objhttp.send

				If objhttp.responseXML.xml <> "" then
					NewData.loadXML objhttp.responseXML.xml
					Set Root1 = NewData.documentElement
					if Root1.HaschildNodes() then
						For Each Node In Root1.childNodes
							IF Node.getAttribute("ITEMCODE") = iItem and  Node.getAttribute("CLASSCODE") = iClass then
								iStock = Node.getAttribute("STOCK")
							End IF
						Next
					end if
				end if

				sTempValues = "orgID="&document.formname.hOrgID.value&"&RefCodes="&document.formname.hRefNo.value&"&RefType="&sRefType&"&ItemCode="&iItem&"&ClassCode="&iClass
				objhttp.Open "GET","XMLGetIssueQtyForRef.asp?"&sTempValues,false
				objhttp.send
				If objhttp.responseXML.xml <> "" then
				    NewData.loadXML objhttp.responseXML.xml
					Set Root1 = NewData.documentElement
					if Root1.HaschildNodes() then
						For Each Node In Root1.childNodes
						    IF Node.getAttribute("ItemCode") = iItem and  Node.getAttribute("ClassCode") = iClass then
							    iToIssue = Node.getAttribute("ToIssueQty")
							    exit for
							End IF
						Next
				    else
				        iToIssue = 0
					end if
			 	end if

			 	sStore = ""
 	            sBin   = ""
 	            sStoreName = ""
			 	objhttp.open "GET","GetItemStoreInfo.asp?ItemCode="&iItem,false
			 	objhttp.send
			 	    'alert(objhttp.responsetext)
			 	if objhttp.responseXML.xml<>"" then
			 	    NewData.loadXML objhttp.responseXML.xml
			 	    set ndStorage = NewData.documentElement
			 	    iCountStore = ndStorage.getAttribute("StoreCount")
			 	    if ndStorage.hasChildNodes() then
			 	        for each ndStore in ndStorage.childNodes
			 	            sStore = ndStore.getAttribute("LocNo")
			 	            sBin   = ndStore.getAttribute("BinNo")
			 	            sStoreName = ndStore.getAttribute("StoreName")
			 	            if trim(sBin)="" or isNull(sBin) or sBin="0" then sBin = "NULL"
			 	            bFlag = false
			 	            for each ndSto in CheckNode.Item(iCounter).childNodes
			 	                if ndSto.nodeName="STORAGE" then
			 	                    bFlag = true
			 	                else
			 	                    bFlag = false
			 	                end if
			 	            next

			                if not bFlag then
			                set ndStorage = OutData2.createElement("STORAGE")
                                ndStorage.setAttribute "STORE",sStore
                                ndStorage.setAttribute "BIN",sBin
                                ndStorage.setAttribute "APPLICABLE","IN"
                                ndStorage.setAttribute "MONTHYEAR",""
                                ndStorage.setAttribute "QTY","0"
                                ndStorage.setAttribute "STORAGEVALUE","0"
                                ndStorage.setAttribute "SQ","0"
                                CheckNode.Item(iCounter).appendchild ndStorage
                            end if 'if not bFlag then
			 	        next
			 	    end if
			 	end if
			 	'alert(iCountStore)
				''''''''''''''''''''
				j = j + 1
				iser = iser + 1

				set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length - cint(iCtr))
				set headerCell=oRow.insertCell()
				headerCell.innerText = iSer
				headerCell.className="ExcelSerial"
				headerCell.align="center"


				set headerCell=oRow.insertCell()
				headerCell.innerHTML=sItemDesc
				headerCell.innerHTML="<a class='ExcelDisplayLink' href=#  name=""lnkA"&CStr(iItem)&"A"&CStr(iClass)&"A"&CStr(sOrgId)&""" onClick=""javascript:DisplayItem(this.name)"">" & sItemDesc & "</a>"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"

				set headerCell=oRow.insertCell()
				headerCell.innerHtml =  "<input type=""text"" name=""txtStore"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" class='FormElemRead' value="& sStoreName &" >"
				if iCountStore > 1 then
				    headerCell.innerHtml = headerCell.innerHtml + "<a><img name=""btnZ"&CStr(iClass)&"Z"&CStr(iItem)&"Z"&CStr(iEntNo)&"Z"& CStr(sOrgId)&"Z"&cstr(sAttributeList)&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Pick Details"" onClick=""CheckStorage(this)""></a>"
			    end if
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"


			'	set headerCell=oRow.insertCell()
			 '   set oText = document.createElement("<input type=""text"" name=""txtMY"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" size=""12"" maxlength=""10"" class=""FormElem"">" )
			  '  headerCell.width = "10"
			  '  headerCell.appendChild(oText)
			   ' headerCell.className="ExcelInputCell"

			    set headerCell=oRow.insertCell()
			    if trim(sMode)="E" then
    			    set oText = document.createElement("<input type=""text"" name=""txtQTY"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" class=""FormElem"" value="""&iQty&""" style=""text-align=right"">" )
			    else
			        set oText = document.createElement("<input type=""text"" name=""txtQTY"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" class=""FormElem"" value=""0"" style=""text-align=right"">" )
			    end if
			    headerCell.width = "10"
			    headerCell.appendChild(oText)
			    headerCell.className="ExcelInputCell"

			    set headerCell=oRow.insertCell()
			    if trim(sMode)="E" then
			        set oText = document.createElement("<input type=""text"" name=""txtVAL"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" size=""12"" onkeypress=""DoKeyPress('Y',10,2)"" class=""FormElem"" value="""& sValue &""" style=""text-align=right"">" )
			    else
			        set oText = document.createElement("<input type=""text"" name=""txtVAL"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" size=""12"" onkeypress=""DoKeyPress('Y',10,2)"" class=""FormElem"" value=""0"" style=""text-align=right"">" )
			    end if
			    headerCell.width = "10"
			    headerCell.appendChild(oText)
			    headerCell.className="ExcelInputCell"

			    set headerCell=oRow.insertCell()
			    set oText = document.createElement("<input type=""button"" value="" Yes "" name=""btn:"&cstr(iClass)&":"&cstr(iItem)&":"&cstr(iEntNo)&""" size=""12"" maxlength=""10"" class=""AddButtonX"" onClick=""GetLot(this)"">" )
			    headerCell.appendChild(oText)
			    headerCell.className="ExcelFieldCell"
			    headerCell.align="center"

			    set headerCell=oRow.insertCell()
			    set oText = document.createElement("<input type=""checkbox"" value="" W "" name=""ChkProduct"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" class=""FormElem"">" )
			    headerCell.appendChild(oText)
			    headerCell.className="ExcelFieldCell"
			    headerCell.align="center"

				document.formname.hCtr.value = document.all.tblData.rows.length + 1

			Next
		end if

	end if

End Function
'********************************************
	Function DoChanges(obj)
	    if obj.value = "PRD" then
	       set WorkCenter = showModalDialog("../../Common/WorkCenterPopup.asp","","dialogHeight:150px;dialogWidth:300px;")
	       if WorkCenter.hasChildNodes() then
	            for each ndChild in WorkCenter.childNodes
	                document.formname.hWCCode.value = ndChild.getAttribute("Code")
	                spanWCName.innerText = ndChild.getAttribute("Name")
	                exit for
	            next
	       else
	            if not confirm("Do you want to continue without selecting the workcenter?") then
	                DoChanges(obj)
	            end if
	       end if
	    end if

	end Function

	Function DoDisable(obj)
		exit function
	'	document.formname.selItmType.selectedIndex = 0
		document.formname.selClass.options.length = 0

	'	document.formname.selItmType.disabled = false
		document.formname.selClass.disabled = false
		if obj.value = "M" then
		'	document.formname.selItmType.disabled = true
			document.formname.selClass.disabled = true
		end if
	end Function

dim j

j = 0

Function checkNumbers(val)
	dim valid,temp,i
	valid = "0123456789."
	for i=1 to len(val)
		temp = mid(val,i,1)
		if Instr(1,valid,temp) > 0 then
			checkNumbers = true
		else
			checkNumbers = false
			exit for
		end if
	next
end Function


Function ClearTable()
	dim i
	for i=2 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(2)
	next
	j = 0
end Function

Function clearXML()
	Set Root = OutData.documentElement
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.nodeName = "STORAGE" then
				Root.removeChild HeaderNode
			end if
		next
	end if
end Function

Function GetLot(obj)
Dim dRcvdOn
	arrTemp = split(obj.name,":")
	iClass = arrTemp(1)
	iItem= arrTemp(2)
	iEntNo = arrTemp(3)
	'set iMonth = eval("document.formname.txtMY"&iClass&"A"&iItem&"A"&iEntNo)
	set iQty = eval("document.formname.txtQTY"&iClass&"A"&iItem&"A"&iEntNo)
	set iVal = eval("document.formname.txtVAL"&iClass&"A"&iItem&"A"&iEntNo)

	dRcvdOn = document.formname.ctlRcvdOn.getDate()

	sOrgID = document.formname.hOrgID.value

	sStoreName =  eval("document.formname.txtStore"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)).value

'	if trim(iMonth.value) = "" then
'		alert("Enter Date")
'		iMonth.select()
'		exit function
'	else
    if trim(iQty.value) = "" then
		alert("Enter Quantity")
		iQty.select()
		exit function
	elseif not checkNumbers(iQty.value) then
		alert("Enter Numerals Only")
		iQty.select()
		exit function
	elseif trim(iVal.value) = "" then
		alert("Enter Value")
		iVal.select()
		exit function
	elseif not checkNumbers(iVal.value) then
		alert("Enter Numerals Only")
		iVal.select()
		exit function
	else
		'	for i = 0 to document.formname.elements.length - 1
		'		if document.formname.elements(i).type = "text" then
'
'					if (Instr(document.formname.elements(i).name,"txtMY") > 0) then
'						set objD = document.formname.elements(i)
'
'							if trim(objD.value) = "" then
'								alert "Enter Date"
'								objD.select()
'								Exit Function
'							end if
'					end if
'				end if
'			next

			set ndRoot = OutData2.documentElement

			if ndRoot.haschildNodes() then
			    for each ndDet in ndRoot.childNodes
			        if ndDet.nodeName="Details" then
			            for each ndItem in ndDet.childNodes
			                if ndItem.nodeName="ItemDetail" then
			                    if ndItem.getAttribute("ItemCode")=iItem and ndItem.getAttribute("CLACODE") = iClass and ndItem.getAttribute("ENTRYNO") = iEntNo then
			                        sAttID = ndItem.getAttribute("ATTRIBUTELIST")
			                        sReceiptType = ndItem.getAttribute("RECEIPTNUM")
			                        for each ndStorage in ndItem.childNodes
			                            if ndStorage.nodeName="STORAGE" then
			                                sLoc = ndStorage.getAttribute("STORE")
			                                sBin = ndStorage.getAttribute("BIN")
			                            end if
			                        next
			                    end if
			                end if
			            next
			        end if
			    next
			end if
			if trim(sLoc)="" or isNull(sLoc) then
			    alert("Select Storage Location and continue")
			    exit function
			end if
		    if trim(sReceiptType)="N" then exit function

			sTempValues = sReceiptType&"``"&iItem&"``"&iClass&"``"&sOrgID&"``"&sLoc&"``"&sBin&"``"&iQty.value&"``"&replace(sStoreName,"&","and")&"``"&document.formname.hStoresUom.value&"``"&iVal.value&"``"&dRcvdOn&"``"&sItemType&"``"&sAttID


		sRefNo = document.formname.hRcptNo.value
		Set OutDataValue = showModalDialog("../../Common/PackingLotSerialDetails.asp?sTemp="&sTempValues&"&CallFrom=IRCPT&RefNo="&sRefNo,OutData2,"dialogHeight:580px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")

		sExp = "//ItemDetail[@ItemCode="""&iItem&"""]/STORAGE[@STORE="""&trim(sLoc)&""" and @BIN="""&trim(sBin)&"""]"
		set ndStorage = ndRoot.selectNodes(sExp)
		if ndStorage.length>0 then
		    ndStorage.Item(0).Attributes.getNamedItem("MONTHYEAR").value = dRcvdOn 'iMonth.value
		    iQty.value = ndStorage.Item(0).Attributes.getNamedItem("QTY").value
		    iVal.value = ndStorage.Item(0).Attributes.getNamedItem("STORAGEVALUE").value
		end if
	end if
	'alert(OutData2.xml)
end Function

Function CheckSubmit(sFinFrom,sFinTo,todaysdate)
	Dim Root, i, iTotalQuantity, objStorage,sItemType

	'Set Root = OutData.documentElement
	dMinDate = sFinFrom
	sFinFrom = right(sFinFrom,4)&left(sFinFrom,2)
	sFinTo = right(sFinTo,4)&left(sFinTo,2)
	sItemType = document.formname.hIType.value
	sDept = document.formname.hDept.value
	sType = document.formname.hType.value

'	if document.formname.selStorage.value = "" then
'		alert("Select Storage Location")
'		document.formname.selStorage.focus
'		exit function
'	elseif (document.formname.radSource(0).checked or document.formname.radSource(1).checked or document.formname.radSource(2).checked or document.formname.radSource(3).checked) and document.formname.txtSource.value = "" then
'		alert("Enter Source Reference")
'		document.formname.txtSource.focus
'		exit function
'	else

		if j > 0 then
			for i = 0 to document.formname.elements.length - 1
				if document.formname.elements(i).type = "text" then
					if (Instr(document.formname.elements(i).name,"txtQTY") > 0) then

						set objD = document.formname.elements(i)
						    if trim(objD.value) = "" or trim(objD.value) = "0" then
								alert "Enter Quantity"
								objD.select()
								exit function
							elseif not checkNumbers(objD.value) then
								alert "Enter Numerals Only"
								objD.select()
								exit function
							elseif trim(document.formname.elements(i+1).value) = "" or trim(document.formname.elements(i+1).value) = "0" then
								alert "Enter Value"
								document.formname.elements(i+1).select()
								exit function
							elseif not checkNumbers(document.formname.elements(i+1).value) then
								alert "Enter Numerals Only"
								document.formname.elements(i+1).select()
								exit function
							end if
					end if
				end if
			next
		end if

		set ndRoot = OutData2.documentElement
		    if ndRoot.hasChildNodes() then
		        for each ndDet in ndRoot.childNodes
		            if ndDet.nodeName="Details" then
		                for each ndItem in ndDet.childNodes
		                    if ndItem.nodeName="ItemDetail" then
		                    iTotalQuantity =0
		                        iItem = nditem.getAttribute("ItemCode")
		                        iClass = ndItem.getAttribute("CLACODE")
		                        iEntNo = ndItem.getAttribute("ENTRYNO")
		                        if ndItem.hasChildNodes() then
		                            for each ndStorage in ndItem.childNodes
		                                if ndStorage.nodeName = "STORAGE" then
	                                        'ndStorage.setAttribute "MONTHYEAR",trim(eval("document.formname.txtMY"&iClass&"A"&iItem&"A"&iEntNo).value)
	                                        ndStorage.setAttribute "MONTHYEAR",document.formname.ctlRcvdOn.getDate()
				                            ndStorage.setAttribute "QTY",trim(eval("document.formname.txtQTY"&iClass&"A"&iItem&"A"&iEntNo).value)
				                            ndStorage.setAttribute "STORAGEVALUE",trim(eval("document.formname.txtVAL"&iClass&"A"&iItem&"A"&iEntNo).value)
				                            ndStorage.setAttribute "SQ","0"
				                            iTotalQuantity = cdbl(iTotalQuantity) + cdbl(trim(eval("document.formname.txtQTY"&iClass&"A"&iItem&"A"&iEntNo).value))
		                                end if
		                            next
		                        end if
		                        ndItem.setAttribute "QTY",iTotalQuantity
		                        if eval("document.formname.chkProduct"&iClass&"A"&iItem&"A"&iEntNo).checked then
		                            ndItem.setAttribute "BYPRODUCT","W"
		                        else
		                            ndItem.setAttribute "BYPRODUCT","P"
		                        end if
		                    end if
		                next
		            end if
		        next
		    end if

		iCounter = 0
		    if ndRoot.hasChildNodes() then
		        for each ndDet in ndRoot.childNodes
		            if ndDet.nodeName="Details" then
		                for each ndItem in ndDet.childNodes
		                    if ndItem.nodeName="ItemDetail" then
		                        sReceiptNum = ndItem.getAttribute("RECEIPTNUM")
		                        iItem = nditem.getAttribute("ItemCode")
		                        iClass = ndItem.getAttribute("CLACODE")
		                        iEntNo = ndItem.getAttribute("ENTRYNO")
		                        if ndItem.haschildNodes() then
		                            for each ndStorage in ndItem.childNodes
		                                if ndStorage.nodeName = "STORAGE" then
		                                    if trim(sReceiptNum)<>"N" then
		                                        if not ndStorage.haschildNodes() then
    		                                        alert("Enter Lot / Serial Details for "&ndItem.getAttribute("ITEMNAME"))
					                                exit function
		                                        end if
		                                    end if 'if trim(sReceiptNum)<>"N" then
		                                end if
		                            next
		                        else
		                             alert("Select Storage Location for "&ndItem.getAttribute("ITEMNAME"))
					                 exit function
		                        end if 'if ndItem.haschildNodes() then
		                    end if
		                next
		            end if
		        next
		    end if

'		if document.formname.radSource(0).checked then
'			sRefType = "S"
'		elseif document.formname.radSource(1).checked then
'			sRefType = "P"
'		elseif document.formname.radSource(2).checked then
'			sRefType = "R"
'		elseif document.formname.radSource(3).checked then
'			sRefType = "M"
'		elseif document.formname.radSource(4).checked then
'			sRefType = "I"
'		else
'			sRefType = "N"
'		end if
'

        sReferenceName = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
        if trim(sReferenceName)="12" then
            sRefType = "I"
        elseif trim(sReferenceName)="17" then
            sRefType = "P"
        else
            sRefType = "N"
        end if

		if (sDept = "PRD" and sType = "T") then
			'if document.formname.radNoSeries(0).checked then
			'	sPack = "M"
			'elseif document.formname.radNoSeries(1).checked then
				sPack = "N"
			'end if
		end if

        ndRoot.setAttribute "DEPT",document.formname.selDepart(document.formname.selDepart.selectedIndex).value
		ndRoot.Attributes.getNamedItem("PACKNUM").Value = sPack
		ndRoot.Attributes.getNamedItem("SRCREFTYPE").Value = sRefType
		ndRoot.Attributes.getNamedItem("SRCREFNO").Value =document.formname.hRefNo.value
		ndRoot.Attributes.getNamedItem("STYPE").Value = document.formname.hWCCode.value
		ndRoot.setAttribute "APPREFTYPE",sReferenceName
		ndRoot.setAttribute "APPREFNO",document.formname.hRefNo.value
		ndRoot.setAttribute "APPREFDATE",document.formname.hRefDate.value
		ndRoot.setAttribute "RCVDON",document.formname.ctlRcvdOn.getDate
		ndRoot.setAttribute "AUTOACCOUNT",document.formname.hAutoAccount.value

		sMode = document.formname.hMode.value
		Set objhttp = CreateObject("Microsoft.XMLHTTP")
		if trim(sMode)="E" then
		    objhttp.Open "POST","../Master/XMLSave.asp?SessionFlag=true&Value=ReceiptLotDataEdit&Folder=Transaction", false
		else
		    objhttp.Open "POST","../Master/XMLSave.asp?SessionFlag=true&Value=ReceiptLotData&Folder=Transaction", false
		end if
		objhttp.send OutData2.XMLDocument


	   '' added by ragav on april 27,2011
	   'alert(OutData2.xml)

	   sRcptNo = document.formname.hRcptNo.value
	    if trim(sMode)="E" then
		    document.formname.action = "InternalReceiptUpdate.asp?CurrDate="&document.formname.hCurrDate.value&"&RcptNo="&sRcptNo
		else
		    document.formname.action = "receiptNewInsert.asp?CurrDate="&document.formname.hCurrDate.value
		end if
		document.formname.submit
'	end if

End Function

Function SelectReference(RefType)
Dim OrgID
Dim RootRef
OrgID =document.formname.hOrgID.value

   ' if document.formname.selStorage.selectedIndex < 0 then
   '     alert("Select the Storage Location")
   '     document.formname.radSource(5).checked = true
   '     exit function
   ' end if
    if RefType = "I" then
        document.formname.txtSource.disabled = true
        document.formname.txtSource.className="FormElemRead"

        set RootRef = showModalDialog("IntRcptRefNoSel.asp?RefType="&RefType&"&OrgID="&OrgID,"","dialogHeight=200px;dialogWidth:500px;")
        if RootRef.hasChildNodes() then
            for each ChildNode in RootRef.childNodes
                document.formname.txtSource.value=ChildNode.getAttribute("value")
                spanSource.innerText = ChildNode.getAttribute("name")
                exit for
            next
        else
            if not confirm("Do you want to continue without select the Issue Reference") then
                set RootRef = showModalDialog("IntRcptRefNoSel.asp?RefType="&RefType&"&OrgID="&OrgID,"","dialogHeight=200px;dialogWidth:500px;")
                if RootRef.hasChildNodes() then
                    for each ChildNode in RootRef.childNodes
                        document.formname.txtSource.value=ChildNode.getAttribute("value")
                        spanSource.innerText = ChildNode.getAttribute("name")
                        exit for
                    next
                end if
            end if
        end if
    else
        document.formname.txtSource.disabled = false
        document.formname.txtSource.className="FormElem"
        document.formname.txtSource.value=""
        spanSource.innerText = ""
    end if
End Function
'*****************************************
Function EditInit(RcptNo)
sMode = document.formname.hMode.value
    if trim(sMode)="E" then
        set objhttp=CreateObject("Microsoft.XMLHTTP")
        objhttp.open "GET","PopIntRcptData.asp?RcptNo="&RcptNo,false
        objhttp.send
        'alert(objhttp.responseText)
        if Trim(objhttp.responseXML.xml)<>"" then
            OutData2.loadXML(objhttp.responseXML.xml)
        end if
        set ndRoot = OutData2.documentElement
        if ndRoot.hasChildNodes() then
            document.formname.ctlRcvdOn.setDate = ndRoot.getAttribute("RCVDON")
            sOrgCode= ndRoot.getAttribute("ORGCODE")
            sAppRefType = ndRoot.getAttribute("APPREFTYPE")
            sAppRefNo = ndRoot.getAttribute("APPREFNO")
            sDept = ndRoot.getAttribute("DEPT")
            set objDept = eval("document.formname.selDepart")
            For iCnt = 0 to objDept.length-1
                if trim(sDept)=trim(objDept(iCnt).value) then
                    objDept.selectedIndex = iCnt
                    exit for
                end if
            Next

            set objRef = eval("document.formname.selRefName")
            For iCnt = 0 to objRef.length-1
                if trim(sAppRefType)=trim(objRef(iCnt).value) then
                    objRef.selectedIndex = iCnt
                    exit for
                end if
            Next

            set objhttp = CreateObject("Microsoft.XMLHTTP")
            objhttp.open "GET","../../Common/GetInfoForRefType.asp?orgID="&sOrgCode&"&RefType="&sAppRefType&"&RefNo="& sAppRefNo,false
            objhttp.send
            if trim(objhttp.responseXML.xml)<>"" then
                RefXML.loadXML(objhttp.responseXML.xml)
            end if
            set ndRefRoot = RefXML.documentElement

            if ndRefRoot.hasChildNodes() then
                For Each ndRef in ndRefRoot.childNodes
                    if ndRef.nodeName="Ref" then
                        sCode = ndRef.getAttribute("Code")
                        sDate = ndRef.getAttribute("Date")
                        RefNoDate.innerHTML = sCode &" - "& sDate
                        exit for
                    end if
                Next
            end if
        end if

        DisplayTable date
        set objhttp = CreateObject("Microsoft.XMLHTTP")
	    objhttp.Open "POST","XMLSave.asp?Name=ReceiptLotDataEdit&SessionFlag=true", false	'
	    objhttp.send OutData2.XMLDocument
    end if

End Function
'**********************************************
Function setdate()
sFromDate = document.formname.hMinDate.value
sToDate = document.formname.hMaxDate.value

if DateDiff("d",sToDate,date())>0 then
    document.formname.ctlRcvdOn.setmindate = sFromDate
    document.formname.ctlRcvdOn.setmaxdate = sToDate
    document.formname.ctlRcvdOn.setDate = sToDate
else
    document.formname.ctlRcvdOn.setmindate = sFromDate
    document.formname.ctlRcvdOn.setmaxdate = date()
end if

End Function
</SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<BODY leftMargin=0 topMargin=0 onLoad="EditInit('<%=iRcptNo%>');setdate()">
<form method="POST" name="formname">
<input type=hidden name="hSelectedValue" value="">
<input type=hidden name="hItemNames" value="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hIType" value="<%=sItmType%>">
<input type=hidden name="hOrgCode" value="<%=sOrgID%>">
<input type=hidden name="hItmCode" value="<%=iItem%>">
<input type=hidden name="hClassCode" value="<%=iClass%>">
<INPUT TYPE=HIDDEN NAME="hStoresUom" VALUE="<%=sUoMDesc%>">
<INPUT TYPE=HIDDEN NAME="hReceiptType" VALUE="<%=sReceiptType%>">
<INPUT TYPE=HIDDEN NAME="hDept" VALUE="<%=sDept%>">
<INPUT TYPE=HIDDEN NAME="hType" VALUE="<%=sType%>">
<input type=hidden name="hCurrDate" value="<%=dtCurrDate%>">
<input type=hidden name="hAttributeList" value="<%=sAttributeList%>">
<input type="hidden" name="hItemType" value="">
<input type="hidden" name="hCtr" value="0">
<input type="hidden" name="hMinDate" value="<%=sMinDate%>">
<input type="hidden" name="hMaxDate" value="<%=sMaxDate%>">
<input type="hidden" name="hAutoAccount" value="<%=sAutoInternalRcptAccount%>" />
<input type="hidden" name="hMode" value="<%=sMode%>" />
<input type="hidden" name="hRcptNo" value="<%=iRcptNo%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Internal Receipts
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%"  >
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell">
                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                       <!-- <tr>
                                            <td class="FieldCell">Select Unit</td>
                                            <td class="FieldCellSub" colspan=3>
												<select size="1" name="selUnit" class="FormElem" onChange="resetAll('U')">
													<!--option value="select">Select</option-->
													<%	'Calling the Function which populates Organization Unit list
													'	populateUnitSelected(sOrgID)
													%>
												<!--</select>
											</td>
                                        </tr>-->
                                        <!--<tr>
                                            <td class="FieldCell">Usage</td>
                                            <td class="FieldCellSub" colspan=3>
												<select size="1" name="selDepart" class="FormElem" onChange="resetAll('R');DoChanges(this)" <%if sPassData <> "" then Response.Write " Disabled " %> >
													<option value="select">Select</option>
													<%	'Calling the Function which populates Department list
														'populateDepartment(sDept)
													%>
												</select>
                                            </td>
                                        </tr>-->
                                       <tr>
                                            <td class="FieldCell">Received From</td>
                                            <td class="FieldCellSub">
												<select size="1" name="selDepart" class="FormElem" onChange="resetAll('R');DoChanges(this)" <%if sPassData <> "" then Response.Write " Disabled " %> >
													<option value="select">Select</option>
													<%	'Calling the Function which populates Department list
														ReceivedFrom(sOrgID)
													%>
												</select>&nbsp;
												<span id="spanWCName" class="dataonly"></span>
												<input type="hidden" name="hWCCode" value="">
                                            </td>
                                            <td class="FieldCellSub">
                                                Received On
                                            </td>
                                            <td class="FieldCellSub">
                                                <%
                                                    InsertDatePicker("ctlRcvdOn")
                                                %>
                                            </td>
										<!--<tr>
										    <td class="FieldCell">Select Type</td>
										    <td class="FieldCellSub" colspan=3>
												<select size="1" name="selAddType" class="FormElem" onChange="DoDisable(this)"  <%if sDept <> "PRD" then Response.Write " Disabled " %> >
													<option value="N" <% If sType = "N" Then Response.Write "Selected" %> >Select</option>
													<option value="W" <% If sType = "W" Then Response.Write "Selected" %>>Work Center</option>
													<option value="P" <% If sType = "P" Then Response.Write "Selected" %>>Packing</option>
													<option value="M" <% If sType = "M" Then Response.Write "Selected" %>>Mixing</option>
													<option value="T" <% If sType = "T" Then Response.Write "Selected" %>>Waste</option>
											    </select>
											</td>
										</tr>-->
										<!--<tr>
										    <td class="FieldCell">Item Type</td>
										    <td class="FieldCellSub">
											    <select size="1" name="selItmType" class="FormElem" onChange="resetAll('T')">
													<option value="select">Select</option>
													<%	'Calling the Function which populates the Item Type list
												'		populateItemTypeSelected sItmType
													%>
												</select>
											</td>
										</tr>-->

                                        <!--tr>
                                            <td class="FieldCell">Source Reference</td>
                                            <td class="FieldCellSub" colspan="3">
												<select size="1" name="selSrc" class="FormElem" onChange="CheckType(this)">
													<option value="select">Select</option>
													<option value="M">MR / Direct Issue</option>
													<option value="N">None</option>
												</select>
                                            </td>
                                        </tr-->

											<!--<tr>
											   <td class="FieldCell">Select Item</td>
											   <td class="FieldCellSub" colspan=4>
											   <% If iItem <> "" Then %>
											   <span class="DataOnly" ><%=ItemDisplay(iItem,iClass)%>&nbsp;</span>
											   <% End If %>
											   <a onClick="Search()" href="#">
													<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif"  align="top" width="11" height="11" alt="Select Item">
												</a>
												</td>
											 </tr>-->


											<!--<tr>
												<td class=FieldCell> Receipt Numbering</td>
												<td class='FieldCellSub'>
													<Span Class="DataOnly"><%=sReceiptName%>&nbsp;</Span>
                                                </td>
											</tr>-->

											<!--<tr>
												<td class=FieldCell valign="top"> Storage Location</td>
												<td class='FieldCellSub' colspan="4">
													<select size="5" name="selStorage" class="FormElem" multiple onBlur="GetStockDet(this,'<%=sCheck%>')">
													<%	'Calling the Function which populates the Store list
														'populateStores sOrgID, sEligible
													%>
													</select>
                                                </td>
											</tr>-->

											<!--<tr>
												<td class=FieldCell >Source Reference</td>
												<td class='FieldCellSub' colspan="4">
												<%	if sDept <> "PRD" then %>
													<input type="radio" value="S" name="radSource" class="FormElem" onclick="SelectReference('S')"> Sales Order&nbsp;
													<input type="radio" value="P" name="radSource" class="FormElem" onclick="SelectReference('P')"> Purchase Order&nbsp;
													<input type="radio" value="R" name="radSource" class="FormElem" onclick="SelectReference('R')"> Production Order&nbsp;
													<input type="radio" value="M" name="radSource" class="FormElem" onclick="SelectReference('M')"> Mixing&nbsp;
													<input type="radio" value="I" name="radSource" class="FormElem" onclick="SelectReference('I')"> Issue&nbsp;
													<input type="radio" value="N" name="radSource" class="FormElem" onclick="SelectReference('N')" CHECKED> None
                                                <%	elseif sDept = "PRD" then %>
													<input type="radio" value="S" name="radSource" class="FormElem" onclick="SelectReference('S')" DISABLED> Sales Order&nbsp;
													<input type="radio" value="P" name="radSource" class="FormElem" onclick="SelectReference('P')" DISABLED> Purchase Order&nbsp;
													<input type="radio" value="R" name="radSource" class="FormElem" onclick="SelectReference('R')" CHECKED> Production Order&nbsp;
													<input type="radio" value="M" name="radSource" class="FormElem" onclick="SelectReference('M')" > Mixing&nbsp;
													<input type="radio" value="I" name="radSource" class="FormElem" onclick="SelectReference('I')"> Issue&nbsp;
													<input type="radio" value="N" name="radSource" class="FormElem" onclick="SelectReference('N')" DISABLED> None
                                                <%	end if %>
                                                </td>
											</tr>-->
											<tr>
											    <td class="FieldCell">Reference Name</td>
											    <td class="FieldCellSub">
											       <select name="selRefName" class="FormElem" Onchange="GetDetails()">
													<%
													    RefTypePop 4,4
													%>
													</select>
													<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click Here to Edit Usage Information" width="11" height="11" onClick="GetDetails()"></a>
													&nbsp;<span id="RefNoDate" class="dataonly"></span>
													<input type="hidden" name="hRefNo" value="">
													<input type="hidden" name="hRefDate" value="">
											    </td>
											    <td class="FieldCellSub">
                                                    Created By
                                                </td>
                                                <td class="FieldCellSub">
                                                    <span id="spancreatedby" class="DataOnly"><%=session("username")%></span>
                                                </td>
											</tr>



											<!--<tr>
												<td class=FieldCell valign="top"></td>
												<td class='FieldCellSub' colspan="4">
													<input type="text" name="txtSource" size="20" maxlength=30 class="FormElem">
													&nbsp;<span id="spanSource" class="dataonly"></span>
                                                </td>
											</tr>-->


                                    </table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

							<tr>

								<td valign="top" colspan="3">
									<table border="0" cellspacing="1" width="100%" >

										<tr>
											<td align="center"></td>
											<td>
												<DIV class="frmBody" id="frm1" style="width: 750; height:340;">
													<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width=100%>
														<tr>
															<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2">Item Name</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2">Storage Location</td>
															<td class="ExcelHeaderCell" align="center" colspan="2">Stock Details</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2" width="75">Lot & Serial</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2" width="75">By Product?</td>
														</tr>
														<tr>
															<td class="ExcelHeaderCell" align="center">Net. Quantity</td>
															<td class="ExcelHeaderCell" align="center">Unit Rate</td>
														</tr>
													</table>
													<input type="button" name="btnAddItem" class="AddButtonX" value="Add Item" onclick="AddItem()">
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
														<td valign="middle" class="ActionCell" align="center">
														<% 'Response.Write sFinFrom & "," & sFinTo %>
															<input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit('<%=sFinFrom%>','<%=sFinTo%>','<%=dtCurrDate%>')">
															<input type="reset" value="Reset" name="B1" class="ActionButton">
															<input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('MATERIALRECEIPTS.ASP?RCPT=A')">
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
											</td>
										</tr>
                                    </table>
								</td>

							</tr>

							<tr>
								<td align="center" class="ClearPixel" colspan="3">
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
<%
	' Function to populate Department
	Function populateDepartment(Dept)
		' Declaration of variables
		Dim dcrs,sDepartCode,sDepartDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DEPTNO,DEPTNAME FROM INV_M_DEPARTMENT ORDER BY DEPTNO"
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR ORDER BY ISSUEDFORDESCRIPTION"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sDepartCode = dcrs(0)
		set sDepartDesc = dcrs(1)

		Do While Not dcrs.EOF
			If Dept = trim(sDepartCode) Then
				Response.Write("<OPTION VALUE="""&trim(sDepartCode)&""" Selected>"&trim(sDepartDesc)&" </OPTION>" &vbcrlf)
			Else
				Response.Write("<OPTION VALUE="""&trim(sDepartCode)&""">"&trim(sDepartDesc)&"</OPTION>" &vbcrlf)
			End If
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate Store
	Function populateStores(sOrgID,sEligible)
		' Declaration of variables
		Dim dcrs,dcrs1,sLoc,sBin,sBinName,sLocName,sLocCode,imaxLoc,sSql
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT LOCATIONNUMBER,LOCATIONNAME,APPLICABLEFOR FROM Inv_M_Storage WHERE OUDEFINITIONID = " & Pack(sOrgID) & " AND APPLICABLEFOR IN ('IN') ORDER BY 1"
			'Response.Write dcrs.source
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			Do While Not dcrs.EOF
				sLoc = trim(dcrs(0))
				sLocName = trim(dcrs(1))
				sLocCode = trim(dcrs(2))

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " ORDER BY BINNUMBER"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					do while not dcrs1.EOF
						Response.Write("<OPTION VALUE="""&sLoc&"-"&trim(dcrs1(0))&"-"&sLocCode&""">"&sLocName&" -- "&trim(dcrs1(1))&"</OPTION>" &vbcrlf)
					dcrs1.MoveNext
					loop
				else
					Response.Write("<OPTION VALUE="""&sLoc&"-NULL-"&sLocCode&""">"&sLocName&"</OPTION>" &vbcrlf)
				end if
				dcrs1.Close

			dcrs.MoveNext
			Loop
		else
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'PUR','PURCHASE', " &_
				" 'PU','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'IOO','INSPECTION-OUTORDER', " &_
				" 'OI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'POI','INSPECTION-PREORDER', " &_
				" 'POI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'SAL','SALES', " &_
				" 'SA','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'INV','INVENTORY', " &_
				" 'IN','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'PI','INSPECTION PROCESS', " &_
				" 'PI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_ORGSTORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'PSI','POST SALE', " &_
				" 'PSI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'MAN','MANUFACTURING', " &_
				" 'MA','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			Server.Execute ("XMLStorageDefault.asp")

			populateStores sOrgID,sEligible

		end if
		dcrs.Close
	End Function
%>


<%
	' Function to Display UoM
	Function DisplayUoM(iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>

<%
    ' Function to Display ReceivedFrom
	Function ReceivedFrom(sOrgCode)
		Dim sQuery,rsTemp,objrs
        set rsTemp = Server.CreateObject("ADODB.Recordset")
        set objrs = Server.CreateObject("ADODB.Recordset")

	    sQuery = "Select DeptShortName,DepartmentName from APP_M_Departments"
	    rsTemp.Open sQuery,con
	    if not rsTemp.Eof then
		    do while not rsTemp.EOF
			    Response.Write "<option value='"& trim(rsTemp(0)) &"'>"&rsTemp(1)&"</option>"
			    rsTemp.MoveNext
		    loop
		end if
	    rsTemp.Close
	    'Response.Write "<option value='Party'>Party</option>"
	    Response.Write "<option value='Unit'>Other Unit</option>"
		    sQuery = "Select OuDefinitionID,OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OuDefinitionID)>4 and OuDefinitionID not in('"& sOrgCode &"')"
		    objrs.Open sQuery,con
		    if not objrs.EOF then
		        do while not objrs.EOF
		            Response.Write "<option value='"& trim(objrs(0)) &"'>&nbsp;&nbsp;&nbsp;"&trim(objrs(1))&"</option>"
		            objrs.MoveNext
		        loop
		    end if
		    objrs.Close

	End Function
%>
