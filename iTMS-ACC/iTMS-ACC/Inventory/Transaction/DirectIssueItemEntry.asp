<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	DirectIssueItemEntry.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	MAHESWARI
	'Created On					:	March 21, 2008
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	Sep 13,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsIssueInsert.asp
	'Procedures/Functions Used	:	populateStore
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
<!-- #include File="../../include/ItemDisplay.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!--#include file="../../include/CommonFunctions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>MR Issue - Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%

	Dim oDom,objfs,Root,HeaderNode,iMRSNo,newElem,sUnit
	Dim arrSSelected,arrSSelectedName,iCtr,arrItemClass
	dim dcrs,dcrs1

	dim sOrgName,dMRSDate,sIssue,sItmTypeName,sUsage,sItmType,sOrgID
	dim sReqType,sUsageCode,sITypeName,arrSchTemp,sSchTemp,sSchTempValue
	dim arrLocation,sStoreName,sStoreCode,sBinCode,arrStore,sItemName
	dim arrQty,iQtyReq,iQtyIssued,iQtyPending,iQtyAppr,iQtyTrans,iQtyPur
	dim iUnitQty,iOthUnitQty,iMarkQty
	dim iQtyRes,iQtyOnHold,iQtyRej
	dim arrUoM,sUoMDesc,sUoMCode,sRecBy
	dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin
	dim iuserId,sCreatedBy,rsUser
	dim sFinPeriod,arr,sMaxDate,sMinDate,sAutoConsumption,sQuery

	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	sMinDate = "01/04/"& Arr(0)
	sMaxDate = "31/03/"& Arr(1)

	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")

	if len(Month(date())) = 1 then
		sTempMonYr = "0"&Month(date())
	else
		sTempMonYr = Month(date())
	end if
	sMonYr = sTempMonYr&Year(date())

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	sUnit = Request("hUnit")
'	Response.Write "sUnit="&sUnit
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID= '"& sUnit&"' "
		.ActiveConnection = con
		.Open
	end with
	if not dcrs.EOF then
		sOrgName = dcrs(0)
	end if
	dcrs.Close
	sUsage = Request.Form("selUsage")
	sIssue = Request.Form("selReqType")
	sUsageCode = Request.Form("selIssType")
	
	if sIssue = "0" then
		sIssue = "Returnable"
	else
		sIssue = "Non Returnable"
	end if
	Set Root = oDOM.createElement("MRSApproval")
	oDOM.appendChild Root
	Set newElem = oDOM.createElement("MRSHeader")
	newElem.setAttribute "MRSNO",""
	newElem.setAttribute "MRSDATE",""
	newElem.setAttribute "ORGID",sUnit
	newElem.setAttribute "ORGNAME", sOrgName
	newElem.setAttribute "REQTYPE", sIssue

	newElem.setAttribute "USAGE", sUsageCode
	newElem.setAttribute "USAGENAME", sUsage
	newElem.setAttribute "ITYPE", ""
	newElem.setAttribute "ITYPENAME",""

	Root.appendChild newElem
	oDOM.Save server.MapPath("../temp/transaction/MRSIssue"&Session.SessionID&".xml")
	
	sQuery = "Select isNull(AutomaticConsumptionEntry,'N') from APP_M_ApplicationSetup"
	dcrs.open sQuery,con
	if not dcrs.eof then
	    sAutoConsumption = trim(dcrs(0))
	end if
	dcrs.close

	iuserid = Session("userid")
	Set rsUser = Server.CreateObject("ADODB.RecordSet")

	'To get User name
	with rsUser
		.CursorLocation = 3
		.CursorType = 3
		.Source ="Select isNull(UserName,'') from DCS_User where EmployeeNumber = "&iuserid&" "
		.ActiveConnection = con
		.Open
	end with
	if not rsUser.EOF then
		sCreatedBy = trim(rsUser(0))
	end if
	rsUSer.Close
	if objfs.FileExists(Server.MapPath("../temp/transaction/MRS"&iMRSNo&".xml")) then

%>
<script type="application/xml" data-itms-xml-island="1" id="Data" data-src="<%="../temp/transaction/MRS"&iMRSNo&".xml"%>"></script>
<%	else %>
<script type="application/xml" data-itms-xml-island="1" id="Data"><root/></script>
<%	end if %>
<script type="application/xml" data-itms-xml-island="1" id="OutData1"><root/></script> <%'src="<%="../temp/transaction/MRISSUEDETAILS"&Session.SessionID&".xml"%>
<script type="application/xml" data-itms-xml-island="1" id="OutData2"><root></root></script>
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData"></script>
<script type="application/xml" data-itms-xml-island="1" id="UoMData" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutSelectData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData"></script>
<script type="application/xml" data-itms-xml-island="1" id="NewData"></script>
<script type="application/xml" data-itms-xml-island="1" id="RefData" data-src="<% Response.Write("../temp/transaction/UsageSelection"&Session.SessionID&".xml")%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="POrder"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="POConfirm"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="SalesInvoice"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="GatePass"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="RefBasedItem"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTypeData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ConfData"><Root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsIssueItemDetails.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>

<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
'==========================================================
Function GetDetails()
    Dim sRefType,sOrgID,sItemType,sPartyCode,iStock,nFlag,bAddButton
    sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
    sorgID = document.formname.hUnit.value
    sItemType = ""
    sPartyCode = ""
    iStock = "Y"
    nFlag=1
    bAddButton = "N"

    set Root = RefData.documentElement
    'alert(Root.xml)
    sUsage = Root.getAttribute("Usage")
    sIssFor = Root.getAttribute("IssueTo")

    if sUsage = "PRD" and sIssFor="M" and sRefType<>"14" then
        alert("Select Mix Code Reference")
        document.formname.selRefName.focus
        exit function
    end if

sDispItem = 0
    'RefTypeSelection sRefType,sOrgID,sItemType,sPartyCode,iStock,nFlag,bAddButton,sDispItem
    RefTypeSelection sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem
        Set ndRoot = OutData.documentElement
        Set Root = OutData2.documentElement
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
                   'alert(ndChild.xml)
                    set objhttp = CreateObject("Microsoft.XMLHTTP")
                    objhttp.open "GET","InvGetItemDetForRefType.asp?RefType="&sRefType &"&RefCodes="&sRefNo&"&orgID="& document.formname.hUnit.value,false
                    objhttp.send
                    'alert(Objhttp.responseText)
                    if trim(objhttp.responseXML.xml)<>"" then
                        ItemData.loadXML(objhttp.responseXML.xml)
                    else
                        alert(Objhttp.responseText)
                    end if
                    set ndIRoot = ItemData.documentElement
                   if ndIRoot.hasChildNodes() then
                        for each ndIChild in ndIRoot.childNodes
                            if ndIChild.nodeName="Item" then
                                if trim(document.formname.hItemType.value)="" then
                                    sItemCode = ndIChild.getAttribute("ItemCode")
                                    set objhttp=CreateObject("Microsoft.XMLHTTP")
                                    objhttp.open "GET","../../Include/GetItemTypeForItem.asp?ItemCode="&sItemCode ,false
                                    objhttp.send

                                    if trim(objhttp.responseXML.xml)<>"" then
                                        ItemTypeData.loadXML(objhttp.responseXML.xml)
                                        set iRoot = ItemTypeData.documentElement
                                        document.formname.hItemType.value = iRoot.getAttribute("ItemType")
                                    else
                                        alert(objhttp.responseText)
                                    end if
                                end if
                                bItemFlag = false
                                if Root.haschildNodes() then
                                    for each child in Root.childNodes
                                        if child.nodename = "ITEMDETAILS" then
                                            if Child.getAttribute("ITEMCODE")=ndIChild.getAttribute("ItemCode") then
                                                    Child.setAttribute "RefNo",Child.getAttribute("RefNo") &","& ndIChild.getAttribute("No")
                                                bItemFlag = true
                                            end if
                                        end if
                                    next
                                end if

                                if not bItemFlag then
                                    Set newElem = OutData2.createElement("ITEMDETAILS")
			                            newElem.setAttribute "ENTRYNO",ndIChild.getAttribute("EntryNo")
			                            newElem.setAttribute "ITEMCODE", ndIChild.getAttribute("ItemCode")
			                            newElem.setAttribute "CLASSCODE", ndIChild.getAttribute("ClassCode")
			                            newElem.setAttribute "UNIT", document.formname.hUnit.value
			                            newElem.setAttribute "ITEMNAME", ndIChild.getAttribute("ItemName")
			                            newElem.setAttribute "UOM", ndIChild.getAttribute("StoresUoM")
			                            newElem.setAttribute "DECIMAL", ndIChild.getAttribute("Decimal")
			                            newElem.setAttribute "DISPALYED", "N"
			                            newElem.setAttribute "QTY", "0"
			                            newElem.setAttribute "REQUIREDBY", ""
			                            newElem.setAttribute "REQUIREDVALUE", ""
			                            if trim(ndIChild.getAttribute("AttributeList"))="0" then
			                                newElem.setAttribute "ATTRIBUTELIST",""
			                            else
			                                newElem.setAttribute "ATTRIBUTELIST",ndIChild.getAttribute("AttributeList")
			                            end if
			                            newElem.setAttribute "REMARKS",""
			                            newElem.setAttribute "RefNo",ndIChild.getAttribute("No")
			                            Root.appendChild newElem
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
                    Set newElem = OutData2.createElement("ITEMDETAILS")
			            newElem.setAttribute "ENTRYNO",ndChild.getAttribute("EntryNo")
			            newElem.setAttribute "ITEMCODE", ndChild.getAttribute("ItemCode")
			            newElem.setAttribute "CLASSCODE", ndChild.getAttribute("ClassCode")
			            newElem.setAttribute "UNIT", document.formname.hUnit.value
			            newElem.setAttribute "ITEMNAME", ndChild.getAttribute("ItemName")
			            newElem.setAttribute "UOM", ndChild.getAttribute("StoresUoM")
			            newElem.setAttribute "DECIMAL", ndChild.getAttribute("Decimal")
			            newElem.setAttribute "DISPALYED", "N"
			            newElem.setAttribute "QTY", "0"
			            newElem.setAttribute "REQUIREDBY", ""
			            newElem.setAttribute "REQUIREDVALUE", ""
			            newElem.setAttribute "ATTRIBUTELIST",ndChild.getAttribute("AttributeList")
			            newElem.setAttribute "REMARKS",""
			            newElem.setAttribute "RefNo",""
			            Root.appendChild newElem
			    end if  'if ndChild.nodeName="Item" then
			next 'for each ndChild in ndRoot.childNodes
        end if 'if ndRoot.hasChildNodes() then
    end if
    'alert(Root.xml)
    DisplayTable date
    OutData1.loadxml OutData2.xml
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=MRISSUEDETAILS&SessionFlag=true", false	'
	objhttp.send OutData1.XMLDocument
End Function
'***********************************************************
Function GetAddDet(iClass,iItem,sOrgid,iEntNo)
	sUsage = document.formname.hUsage.value
	set Q = eval("document.formname.txtQtyPPX"&iClass&"X"&iItem&"X"&iEntNo)
		if trim(Q.value) = "" then
			alert("Enter Quantity")
			exit function
		end if
		if cdbl(Q.value) = 0 then exit function

		sTempValues = iClass&"|"&iItem&"|"&sOrgid&"|"&sUsage&"|"&trim(Q.value)

	set OutValue = showModalDialog("DirectIssueAddEntry.asp?sTemp="&sTempValues,OutData2,"dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
End Function
'===================================================================
Function PRDITEMPOPULATE(todaysDate)
	Dim sRefCodes
	sRefCodes = document.formname.hRefCodes.value

Set objhttp = CreateObject("Microsoft.XMLHTTP")
	if 1 = 1 then
		sorgID = document.formname.hUnit.value
		sIType = document.formname.hItemType.value
		set rootData = OutSelectData.DocumentElement
	nFlag = 1
		set Root = OutSelectData.DocumentElement
		iStock = "Y"
		'alert(document.formname.hRefType.value)
		if trim(document.formname.hRefType.value)<>"" and document.formname.hDocType.value="WI" then

			objhttp.open "POST","InvGetItemDetForRefType.asp?orgID=" & sorgID &"&RefCodes="&sRefCodes&"&RefType="& document.formname.hRefType.value,false
			objhttp.send
			if trim(objhttp.responseXML.xml)<>"" then
				OutSelectData.loadXML(objhttp.responseXML.xml)
			else
				alert(objhttp.responseText)
				exit function
			end if
			set rootData = OutSelectData.DocumentElement
			sIType = trim(rootData.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType

			'alert(rootData.xml)
			 'alert OutValue.xml
			If not rootData.hasChildNodes Then 	exit function
			sIType = trim(rootData.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
			sTemp1 = sIType
		else
			if trim(document.formname.hUsage.value)="PRD" then
				sIType = "FIB"
				sTemp = "orgID=" & sorgID &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="&cstr(nFlag)&"&RefCodes="& sRefCodes

				objhttp.open "POST","INVISSUEPRDITEMPOPULATE.asp?"&sTemp,false
				objhttp.send
				if trim(objhttp.responseXML.xml)<>"" then
					OutSelectData.loadXML(objhttp.responseXML.xml)
				else
					alert(objhttp.responseText)
					exit function
				end if
				set rootData = OutSelectData.DocumentElement
				sIType = trim(rootData.getAttribute("ItemType"))
				document.formname.hItemType.value = sIType

				'alert(rootData.xml)
				 'alert OutValue.xml
				If not rootData.hasChildNodes Then 	exit function
				sIType = trim(rootData.getAttribute("ItemType"))
				document.formname.hItemType.value = sIType
				sTemp1 = sIType
			end if ' if trim(document.formname.selUsage.value)="PRD" then
		end if ' if trim(document.formname.hRefType.value)<>"" then

		If not rootData.hasChildNodes Then 	exit function

		set root = OutData2.DocumentElement
		if root.haschildnodes then
			for each node in root.childnodes
				if trim(node.NodeName) = trim("ITEMDETAILS") then
					root.removechild node
				end if
			next
		end if
		if rootData.hasChildNodes then

			For each ndTemp in rootData.childNodes
				sAttList =""
			 	'iEntNo = iEntNo + 1
				if root.hasChildNodes then

					sExp ="//ITEMDETAILS [@ENTRYNO = "&ndTemp.attributes.getNamedItem("EntryNo").Value&" and @ITEMCODE = "&ndTemp.attributes.getNamedItem("ItemCode").value&" and @CLASSCODE = "&ndTemp.attributes.getNamedItem("ClassCode").value&"]"
					Set CheckNode = Root.Selectnodes(sExp)
					if CheckNode.Length = 0 then
						sAttrbTemp = ndTemp.getAttribute("AttributeList")
						'alert(sAttrbTemp)

						if sAttrbTemp <> "" then
							Temp = split(sAttrbTemp,",")
							sAttList = ""
							For i = 0 to UBOUND(Temp)
								sValTemp = Split(Temp(i),":")
								sAttList = sAttList &","&sValTemp(0)
							Next
						end if

						sAttList = Mid(sAttList,2)

						Set newElem = OutData2.createElement("ITEMDETAILS")
						newElem.setAttribute "ENTRYNO",ndTemp.attributes.getNamedItem("EntryNo").Value
						newElem.setAttribute "ITEMCODE", ndTemp.attributes.getNamedItem("ItemCode").value
						newElem.setAttribute "CLASSCODE", ndTemp.attributes.getNamedItem("ClassCode").value
						newElem.setAttribute "UNIT", sorgID
						newElem.setAttribute "ITEMNAME", ndTemp.attributes.getNamedItem("ItemName").value
						newElem.setAttribute "UOM", ndTemp.attributes.getNamedItem("StoresUoM").value
						newElem.setAttribute "DECIMAL", ndTemp.attributes.getNamedItem("Decimal").value
						newElem.setAttribute "DISPALYED", "N"
						newElem.setAttribute "QTY", ""
						newElem.setAttribute "REQUIREDBY", ""
						newElem.setAttribute "REQUIREDVALUE", ""
						newElem.setAttribute "REMARKS", ""
						newElem.setAttribute "ATTRIBUTELIST",sAttList

						Root.appendChild newElem
					end if
				else
					sAttrbTemp = ndTemp.getAttribute("AttributeList")

					if sAttrbTemp <> "" then
						Temp = split(sAttrbTemp,",")
						sAttList = ""
						For i = 0 to UBOUND(Temp)
							sValTemp = Split(Temp(i),":")
							sAttList = sAttList &","&sValTemp(0)
						Next
					end if
					sAttList = Mid(sAttList,2)
					Set newElem = OutData2.createElement("ITEMDETAILS")
					newElem.setAttribute "ENTRYNO",ndTemp.attributes.getNamedItem("EntryNo").Value
					newElem.setAttribute "ITEMCODE", ndTemp.attributes.getNamedItem("ItemCode").value
					newElem.setAttribute "CLASSCODE", ndTemp.attributes.getNamedItem("ClassCode").value
					newElem.setAttribute "UNIT", sorgID
					newElem.setAttribute "ITEMNAME", ndTemp.attributes.getNamedItem("ItemName").value
					newElem.setAttribute "UOM", ndTemp.attributes.getNamedItem("StoresUoM").value
					newElem.setAttribute "DECIMAL", ndTemp.attributes.getNamedItem("Decimal").value
					newElem.setAttribute "DISPALYED", "N"
					newElem.setAttribute "QTY", ""
					newElem.setAttribute "REQUIREDBY", ""
					newElem.setAttribute "REQUIREDVALUE", ""
					newElem.setAttribute "REMARKS", ""
					newElem.setAttribute "ATTRIBUTELIST",sAttList
					Root.appendChild newElem
				end if
			next
		end if

		DisplayTable todaysDate
	end if
	'''''''''
	OutData1.loadxml OutData2.xml

	objhttp.Open "POST","XMLSave.asp?Name=MRISSUEDETAILS&SessionFlag=true", false	'
	objhttp.send OutData1.XMLDocument
End Function
'*********************************************************************
Function GetItems(todaysDate)
	Dim sRefCodes
	sRefCodes = document.formname.hRefCodes.value

	'alert(sRefCodes)
Set objhttp = CreateObject("Microsoft.XMLHTTP")
	'if document.formname.selUsage.value = "JWK" or document.formname.selUsage.value = "select" then exit function
	if 1 = 1 then
		sorgID = document.formname.hUnit.value
		sIType = document.formname.hItemType.value
	nFlag = 1
		set Root = OutSelectData.DocumentElement
		iStock = "Y"
		if trim(document.formname.hUsage.value)<>"PRD" then
			'Set OutValue = showModalDialog("ItemSelect.asp?orgID=" & sorgID &"&sIType=" & sIType & "&hSelectMode=M&Flag="+cstr(nFlag),OutSelectData,"dialogHeight:600px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			set OutValue = showModalDialog("../../Common/ItemSelectCommon.asp?orgID=" & sorgID &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag),OutSelectData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			sQuery = trim(OutValue.getAttribute("PassQuery"))
			sIType = trim(OutValue.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
			if ucase(trim(sAct)) <> "CLOSE" then
				do while sAct <> "DONE"
					set OutValue = showModalDialog("../../Common/ItemSelectCommon.asp?"&sQuery,OutSelectData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
					sAct = UCase(trim(OutValue.getAttribute("Action")))
					if ucase(Trim(sAct)) = "CLOSE" then exit do
					sQuery = trim(OutValue.getAttribute("PassQuery"))
					sIType = trim(OutValue.getAttribute("ItemType"))
				loop
			end if
			 'alert OutValue.xml
			If not OutValue.hasChildNodes Then 	exit function
			'alert(OutSelectData.xml)
			'sIType = right(sQuery,3)
			sIType = trim(OutValue.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
			'alert(sTemp1 & " <> " & sIType)
			if sTemp1 <> sIType then
				'RemoveItems
			end if
			'alert(sIType)
			sTemp1 = sIType
		else
			sIType = "FIB"
			set OutValue = showModalDialog("../../Include/ItemSelectRefBased.asp?orgID=" & sorgID &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="&cstr(nFlag)&"&RefCodes="& sRefCodes ,OutSelectData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			sQuery = trim(OutValue.getAttribute("PassQuery"))
			sIType = trim(OutValue.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
			if ucase(trim(sAct)) <> "CLOSE" then
				do while sAct <> "DONE"
					set OutValue = showModalDialog("../../Include/ItemSelectRefBased.asp?"&sQuery,OutSelectData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
					sAct = UCase(trim(OutValue.getAttribute("Action")))
					if ucase(Trim(sAct)) = "CLOSE" then exit do
					sQuery = trim(OutValue.getAttribute("PassQuery"))
				loop
			end if
			 'alert OutValue.xml
			If not OutValue.hasChildNodes Then 	exit function
			sIType = trim(OutValue.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
			'alert(sTemp1 & " <> " & sIType)
			if sTemp1 <> sIType then
				'RemoveItems
			end if
			'alert(sIType)
			sTemp1 = sIType
		end if ' if trim(document.formname.selUsage.value)<>"PRD" then

		If not OutValue.hasChildNodes Then 	exit function

'		alert(OutValue.xml)
		set rootData = OutSelectData.DocumentElement
		set root = OutData2.DocumentElement
		'iEntNo = 0
		'alert(rootData.xml)
		if root.haschildnodes then
			for each node in root.childnodes
				if trim(node.NodeName) = trim("ITEMDETAILS") then
					root.removechild node
				end if
			next
		end if
		if rootData.hasChildNodes then

			For each ndTemp in rootData.childNodes
				sAttList =""
			 	'iEntNo = iEntNo + 1
				if root.hasChildNodes then

					sExp ="//ITEMDETAILS [@ENTRYNO = "&ndTemp.attributes.getNamedItem("EntryNo").Value&" and @ITEMCODE = "&ndTemp.attributes.getNamedItem("ItemCode").value&" and @CLASSCODE = "&ndTemp.attributes.getNamedItem("ClassCode").value&"]"
					Set CheckNode = Root.Selectnodes(sExp)
					if CheckNode.Length = 0 then
						sAttrbTemp = ndTemp.getAttribute("AttributeList")
						'alert(sAttrbTemp)

						if sAttrbTemp <> "" then
							Temp = split(sAttrbTemp,",")
							sAttList = ""
							For i = 0 to UBOUND(Temp)
								sValTemp = Split(Temp(i),":")
								sAttList = sAttList &","&sValTemp(0)
							Next
						end if

						sAttList = Mid(sAttList,2)

						Set newElem = OutData2.createElement("ITEMDETAILS")
						newElem.setAttribute "ENTRYNO",ndTemp.attributes.getNamedItem("EntryNo").Value
						newElem.setAttribute "ITEMCODE", ndTemp.attributes.getNamedItem("ItemCode").value
						newElem.setAttribute "CLASSCODE", ndTemp.attributes.getNamedItem("ClassCode").value
						newElem.setAttribute "UNIT", sorgID
						newElem.setAttribute "ITEMNAME", ndTemp.attributes.getNamedItem("ItemName").value
						newElem.setAttribute "UOM", ndTemp.attributes.getNamedItem("StoresUoM").value
						newElem.setAttribute "DECIMAL", ndTemp.attributes.getNamedItem("Decimal").value
						newElem.setAttribute "DISPALYED", "N"
						newElem.setAttribute "QTY", ""
						newElem.setAttribute "REQUIREDBY", ""
						newElem.setAttribute "REQUIREDVALUE", ""
						newElem.setAttribute "REMARKS", ""
						newElem.setAttribute "ATTRIBUTELIST",sAttList

						Root.appendChild newElem
					end if
				else
					sAttrbTemp = ndTemp.getAttribute("AttributeList")

					if sAttrbTemp <> "" then
						Temp = split(sAttrbTemp,",")
						sAttList = ""
						For i = 0 to UBOUND(Temp)
							sValTemp = Split(Temp(i),":")
							sAttList = sAttList &","&sValTemp(0)
						Next
					end if
					sAttList = Mid(sAttList,2)
					Set newElem = OutData2.createElement("ITEMDETAILS")
					newElem.setAttribute "ENTRYNO",ndTemp.attributes.getNamedItem("EntryNo").Value
					newElem.setAttribute "ITEMCODE", ndTemp.attributes.getNamedItem("ItemCode").value
					newElem.setAttribute "CLASSCODE", ndTemp.attributes.getNamedItem("ClassCode").value
					newElem.setAttribute "UNIT", sorgID
					newElem.setAttribute "ITEMNAME", ndTemp.attributes.getNamedItem("ItemName").value
					newElem.setAttribute "UOM", ndTemp.attributes.getNamedItem("StoresUoM").value
					newElem.setAttribute "DECIMAL", ndTemp.attributes.getNamedItem("Decimal").value
					newElem.setAttribute "DISPALYED", "N"
					newElem.setAttribute "QTY", ""
					newElem.setAttribute "REQUIREDBY", ""
					newElem.setAttribute "REQUIREDVALUE", ""
					newElem.setAttribute "REMARKS", ""
					newElem.setAttribute "ATTRIBUTELIST",sAttList
					Root.appendChild newElem
				end if
			next
		end if

		DisplayTable todaysDate
	end if
	'''''''''
	OutData1.loadxml OutData2.xml

	objhttp.Open "POST","XMLSave.asp?Name=MRISSUEDETAILS&SessionFlag=true", false	'
	objhttp.send OutData1.XMLDocument
	 'alert(OutData1.xml)
End Function
'***************************************************
Function DoChanges(obj)
	dim OutValue

	idUsage.innerHTML = obj(obj.selectedIndex).Text
	document.formname.hUsage.value = obj(obj.selectedIndex).value
	if obj.value = "JWK" then

		sOrgID = document.formname.hUnit.value
	'	sIType = document.formname.selItmType.value
		OutValue = showModalDialog("JobworkPop.asp?sUnit="& sOrgID,"","dialogHeight:200px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		document.formname.hJobWorkNo.value = OutValue

	elseif obj.value = "DIS" or obj.value = "IUT"  or obj.value = "PUR" or obj.value = "REP" or obj.value = "SAL" or obj.value = "SER" then
		sOrgID = document.formname.hUnit.value

		OutValue = showModalDialog("SalesInvoicePartyPopup.asp?ORGID="&sOrgID&"&sWho="&obj.value,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")

		while UBound(arrTemp) = 0
			OutValue = showModalDialog("SalesInvoicePartyPopup.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
			arrTemp = split(OutValue,":")
		wend
		document.formname.hSupplier.value = arrTemp(1)
	elseif obj.value="SUB" then
		sOrgID = document.formname.hUnit.value
		'sIType = document.formname.selItmType.value
		OutValue = showModalDialog("SubContractPop.asp?sUnit="& sOrgID,"","dialogHeight:200px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		document.formname.hSubCon.value = OutValue
	end if
	'alert(OutValue)
end Function
Function DisplayTable(todaysDate)
	iCounter1 = 0
	set rootData = ItemData.DocumentElement
	set root = OutData2.DocumentElement
	set oNodTaxRoot = PurTypeData.DocumentElement
	set rootUoM = UoMData.DocumentElement
	IF document.formname.hCtr.value = "" then document.formname.hCtr.value = 0
	IF document.formname.hCtr.value <> 0 then 	ClearTable()

	sUnit = document.formname.hUnit.value
	'alert(root.xml)

	if root.hasChildNodes then
		sExp ="//ITEMDETAILS [ @DISPALYED = 'N']"
		Set CheckNode = Root.Selectnodes(sExp)
		if CheckNode.Length > 0 then
			iCounter = 0
			'iser = 0
			'alert(document.formname.hCtr.value)
			'IF document.formname.hCtr.value <> 0  then

			For iCounter = 0 to CheckNode.Length - 1
				'alert(CheckNode.Item(iCounter).xml)
				iEntNo = CheckNode.Item(iCounter).Attributes.getNamedItem("ENTRYNO").value
				iItem = CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMCODE").value
				iClass = CheckNode.Item(iCounter).Attributes.getNamedItem("CLASSCODE").value
				sOrgId = CheckNode.Item(iCounter).Attributes.getNamedItem("UNIT").value
				sItemDesc = replace(CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").value,"~~",chr(34))
				sCheck = CheckNode.Item(iCounter).Attributes.getNamedItem("DECIMAL").value
				sItemName = CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").value
				sAttributeList = CheckNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").value
			 	sUOM = CheckNode.Item(iCounter).Attributes.getNamedItem("UOM").value

			 	sStoreCode = 0
				sBinCode = 0
				'alert(sAttributeList)
				sArrList = split(sAttributeList,"#")
				if Ubound(sArrList)>0 then
				   sArrOptName = split(sArrList(1),":")
				 '  alert(sArrOptName(0))
				   if sArrOptName(0)="0" then
			            sOptName=""
				   else
	    			    sOptName = "["& sArrOptName(1)&"]"
    			   end if
				end if
				'alert(sOptName)
			 	Set objhttp = CreateObject("Microsoft.XMLHTTP")
			 	if trim(document.formname.hUsage.value)="PRD" then
					if trim(sOptName)<>"" then
						sItemDesc = sItemDesc & sOptName
					end if 'if trim(sOptName)<>"" then
				end if 'if trim(document.formname.hUsage.value)="PRD" then

				'To Get Stock Details

				sTemp = iItem &":"& iEntNo &":"& iClass &":"& document.formname.hUnit.value&":"&document.formname.hMinDate.value&":"&document.formname.hMaxDate.value&":"&sAttributeList
				'alert(sTemp)
				objhttp.Open "GET","XMLGetStockDetails.asp?Para="&sTemp,false
				objhttp.send

				'alert(objhttp.responsetext)
				'alert(objhttp.responseXML.XML)
				'alert document.all.tblLot.rows.length

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

				''''''''''''''''''''
				j = j + 1
				iser = iser + 1

		'		alert(document.all.tblLot.rows.length &":::"& iCtr)
				set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length - cint(iCtr))
				'Serial no
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtSerA"&CStr(iItem)&"A"&CStr(iClass)&"A"&CStr(iEntNo)&""" value="""&iser&""" size=""1"" class=""FormelemRead"" style=""text-align=center"" READONLY>" )
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"
				headerCell.rowspan="2"

				'Checkbox
				'set headerCell=oRow.insertCell()
				'set oText = document.createElement("<input type=""checkbox"" name=""chkDeleteA"&CStr(iItem)&"A"&CStr(iClass)&"A"&CStr(iEntNo)&""" value="""&iser&""" class=""Formelem"" style=""text-align=right"">" )
				'headerCell.appendChild(oText)
				'headerCell.width = "10"
				'headerCell.className="ExcelInputCell"
				'Item desc

				set headerCell=oRow.insertCell()
				'if len(sItemDesc) > 20 then
				'	sItemDesc=mid(sItemDesc,1,20)&"..."
				'else
					'sItemDesc=sItemDesc
				'end if
				headerCell.innerHTML=sItemDesc
				headerCell.innerHTML="<a class='ExcelDisplayLink' href=#  name=""lnkA"&CStr(iItem)&"A"&CStr(iClass)&"A"&CStr(sOrgId)&""" onClick=""javascript:DisplayItem(this.name)"">" & sItemDesc & "</a>"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"
				headerCell.rowspan="2"
				'headerCell.colspan="2"

				'In Unit
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtStockZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" Value="""&iStock&""" size=""10"" class=""FormElemRead"" style=""text-align=right;cursor:hand;FONT-WEIGHT: bold"" alt=""In Unit Stock Details"" READONLY onkeypress=""DisplayStock(this)"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'Reserved
				set headerCell=oRow.insertCell()
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'Transit
				set headerCell=oRow.insertCell()
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'Issue 'refer(1---1)
				set headerCell=oRow.insertCell()

				headerCell.innerHtml = "<a><img name=""btnZ"&CStr(iClass)&"Z"&CStr(iItem)&"Z"&CStr(iEntNo)&"Z"&cstr(sOptName)&"Z"& CStr(sOrgId)&"Z"&cstr(sAttributeList)&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Pick Details"" onClick=""CheckLot(this,'','"&sAttributeList&"')""></a>"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""hidden"" name=""hStoName"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" Value=""0"">"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""hidden"" name=""hSchA"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" Value="""">"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""hidden"" name=""hStoA"&sStoreCode&"A"&sBinCode&""" Value=""0"">"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""text"" name=""txtQtyPPX"&CStr(iClass)&"X"&CStr(iItem)&"X"&CStr(iEntNo)&""" size=""10"" class=""FormElem"" onkeypress=""DoKeyPress('',7,3)"" style=""text-align:right"" onBlur=""GetSch(this)"" READONLY>"
				headerCell.width = "91"
				headerCell.className="ExcelInputCell"
				headerCell.align="left"

				'Transfer
				'set headerCell=oRow.insertCell()
				'headerCell.innerHtml = "<input type=""text"" name=""txtQtyTraX"&CStr(iClass)&"X"&CStr(iItem)&"X"&CStr(iEntNo)&""" value=""0"" size=""14"" onkeypress=""DoKeyPress('',7,3)"" style=""text-align:right"" class=""Formelem"" READONLY>"
				'headerCell.innerHtml = headerCell.innerHtml + "<a href=""javascript:void(0)""><img name=""btn:"&CStr(iClass)&":"&CStr(iItem)&":"&CStr(iEntNo)&":"&sOptName&":"&CStr(sOrgId)&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Stock Transfer"" onClick=""CheckST(this)""></a>"
				'headerCell.className="ExcelInputCell"
				'headerCell.width = ""
				'headerCell.align="left"
				'Purchase
				'set headerCell=oRow.insertCell()
				'headerCell.innerHtml = "<input type=""text"" name=""txtQtyPrX"&CStr(iClass)&"X"&CStr(iItem)&"X"&CStr(iEntNo)&""" value=""0"" size=""18"" onkeypress=""DoKeyPress('',7,3)"" style=""text-align:right"" class=""Formelem"" >"

				'headerCell.className="ExcelInputCell"
				'headerCell.align="left"

				'Total
				set headerCell=oRow.insertCell()
				'headerCell.innerHtml = "<input type=""text"" name=""txtQtyTotX"&CStr(iClass)&"X"&CStr(iItem)&"X"&CStr(iEntNo)&""" value=""0"" maxlength=""10"" size=""12"" onkeypress=""DoKeyPress('',7,3)"" style=""text-align:right"" class=""Formelem"" >"
				headerCell.innerHtml = ""
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"
				headerCell.width = "10"

				'''''''''''''''''''''
				'second row
				set oRow1 = document.all.tblLot.insertRow(document.all.tblLot.rows.length)

				'Other Unit
				set headerCell=oRow1.insertCell()
				headerCell.innerHtml = "0"
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'On Hold
				set headerCell=oRow1.insertCell()
				headerCell.innerHtml = "0"
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'Rejected
				set headerCell=oRow1.insertCell()
				headerCell.innerHtml = "0"
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'Date
				set headerCell=oRow1.insertCell()
				headerCell.innerHtml = ""
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'By Date
				'set headerCell=oRow1.insertCell()
				'set oText = document.createElement("<SELECT size=""1"" name=""selSTSchZ"&CStr(iClass)&"Z"&CStr(iItem)&"Z"&CStr(iEntNo)&"Z"&CStr(sOptName)&""" class=""FormElem"" onChange=""CheckSTPRSch(this,"&date()&" ,'ST')"">")

				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Immediate"
				'oText1.Value = "ID"
				'oText.Options.Add(oText1)

				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Within x Days"
				'oText1.Value = "WD"
				'oText.Options.Add(oText1)

				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Specific Date"
				'oText1.Value = "SD"
				'oText.Options.Add(oText1)

				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Scheduled"
				'oText1.Value = "S"
				'oText.Options.Add(oText1)
				'headerCell.className="ExcelFieldCell"
				'headerCell.align="left"
				'headerCell.width="30"
				'headerCell.appendChild(oText)


				'set headerCell=oRow1.insertCell()
			 	'set oText = document.createElement("<SELECT size=""1"" name=""selPRSchZ"&CStr(iClass)&"Z"&CStr(iItem)&"Z"&CStr(iEntNo)&"Z"&CStr(sOptName)&""" class=""FormElem"" onChange=""CheckSTPRSch(this,"&date()&",'ST')"">")
				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Immediate"
				'oText1.Value = "ID"
				'oText.Options.Add(oText1)

				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Within x Days"
				'oText1.Value = "WD"
				'oText.Options.Add(oText1)

				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Specific Date"
				'oText1.Value = "SD"
				'oText.Options.Add(oText1)

				'set oText1 = document.createElement("<Option>" )
				'oText1.Text = "Scheduled"
				'oText1.Value = "S"
				'oText.Options.Add(oText1)

				'headerCell.className="ExcelFieldCell"
				'headerCell.align="left"
				'headerCell.width="20"
				'headerCell.appendChild(oText)

				'UOM
				set headerCell=oRow1.insertCell()
				headerCell.innerHtml = sUOM
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				set headerCell = oRow.insertCell()
				set oText = document.createElement("<input type='button' name=""btnAddDet"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" class='ActionButtonX' value='Yes' onClick='GetAddDet("""+ iClass +""","""+ iItem+""","""+ sOrgid + ""","""+iEntNo+""")'>")
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headercell.align="center"
				headerCell.rowspan = "2"

				'document.formname.hCtr.value = document.all.tblLot.rows.length
				document.formname.hCtr.value = document.all.tblLot.rows.length + 1

			Next
		end if

	end if

End Function

Function ClearTable()
	Dim i
		'j = document.formname.hCtr.value  +1
		'alert(document.all.tblLot.rows.length &"-"& j)
		 K = document.all.tblLot.rows.length - 3
	for	i = 1 to  K
		document.all.tblLot.deleteRow(3)
	next

End function

Function popAC()
	document.formname.selAccHead.length = 2

	dim Root,HeaderNode
	sIssuedFor = document.formname.hUsage.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	sIType=trim(sIType)
	objhttp.Open "GET","XMLSelectAccountHead.asp?sIssuedFor="& sIssuedFor, false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData2.loadXML objhttp.responseXML.xml
		Set Root = OutData2.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				document.formname.selAccHead.length = document.formname.selAccHead.length+1
				document.formname.selAccHead.options(document.formname.selAccHead.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selAccHead.options(document.formname.selAccHead.length-1).Value = cstr(HeaderNode.Attributes.Item(0).nodeValue)
			next
		end if
	end if
End Function

Function CreateNew(obj)
	if obj.selectedIndex = "1" then
		idConsumption.style.display = "block"
	end if
End Function

Function hideDiv()
	idConsumption.style.display = "none"
	document.formname.selAccHead.selectedIndex = 0
End Function
Function CheckEntry()
	if trim(document.formname.txtCHead.value) = "" then
		alert("Enter Consumption Head")
		document.formname.txtCHead.select()
		exit Function
	elseif document.formname.selAcc.selectedIndex = "0" then
		alert("Select Account Head")
		document.formname.selAcc.focus()
		exit Function
	else
		sAHead = document.formname.selAcc.value
		sDesc = trim(document.formname.txtCHead.value)
		sIssFor = document.formname.hUsage.value

		Set Root = OutData2.documentElement
		sExp ="//AccountHead [ @ACCHEAD = "&sAHead&" and @CONSUM = '"&sDesc&"' and @ISSFOR = '"&sIssFor&"']"
		Set CheckNode = Root.Selectnodes(sExp)
		if CheckNode.Length > 0 then
			alert("Already Consumption Head exits")
			document.formname.txtCHead.select()
			exit function
		else
			Set newElem = OutData2.createElement("AccountHead")
			newElem.setAttribute "ACCHEAD", sAHead
			newElem.setAttribute "CONSUM", sDesc
			newElem.setAttribute "ISSFOR", sIssFor
			newElem.setAttribute "SRC", "N"

			Root.appendChild newElem

			set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.innerHTML= trim(document.formname.txtCHead.value)
			headerCell.className="ExcelDisplayCell"
			headerCell.align="left"

			set headerCell=oRow.insertCell()
			headerCell.innerHTML= document.formname.selAcc(document.formname.selAcc.selectedIndex).text
			headerCell.className="ExcelDisplayCell"
			headerCell.width="10"

			document.formname.selAcc.selectedIndex = "0"
			document.formname.txtCHead.value = ""

		end if
	end if
end Function

Function PopDone()
	Set Root = OutData2.documentElement

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","ConsumptionHeadInsert.asp", false
	objhttp.send OutData2.XMLDocument
'	alert(objhttp.responseText)
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(3).nodeValue = "N" then
				document.formname.selAccHead.length = document.formname.selAccHead.length+1
				document.formname.selAccHead.options(document.formname.selAccHead.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selAccHead.options(document.formname.selAccHead.length-1).Value = cstr(HeaderNode.Attributes.Item(0).nodeValue)
			end if
		next
	end if
	hideDiv
end Function
'=================================================================
Function init()
Dim Root,PartyNode,RefNode,subNode,sRefName,RefNo,IssName,sRefType
set Root= RefData.documentElement
'alert(Root.xml)

if Root.hasChildNodes() then
	For each subNode in Root.childNodes
		if strcomp(subNode.nodeName,"Ref")=0 then
			set RefNode = subNode
		elseif strcomp(subNode.nodeName,"Party")=0 then
			set PartyNode = subNode
		end if
	next
end if 'if Root.hasChildNodes() then
'alert(PartyNode.xml)
'alert(Root.xml)
	document.formname.hUsage.value = Root.getAttribute("Usage")
	UsageName.innerHtml = Root.getAttribute("UsageName")
	IssName = RefNode.getAttribute("IssName")
	document.formname.hIssForType.value = RefNode.getAttribute("IssName")
    document.formname.hIssueToCode.value = RefNode.getAttribute("Issue")

	if trim(IssName)<>"" and trim(lcase(IssName))<>"select" and not IsNull(IssName) and trim(LCase(IssName))<>"party" then
		selIssueFor.innerHTML = IssName
	elseif trim(IssName)="Party" then
	    selIssueFor.innerHTML = PartyNode.getAttribute("Name")
	else
		selIssueFor.innerHTML = "NA"
	end if
	document.formname.hPartyCode.value = PartyNode.getAttribute("Code")
'	alert(document.formname.hRefCodes.value)
End Function
'===============================================================
Function EditUsageInfo()
	Dim OutValue,objhttp
	set OutValue =  showModalDialog("IssueUsageSelPop.asp?OrgID="& document.formname.hUnit.value,RefData,"dialogHeight:340px;dialogwidth=500px;center:yes;help:no;resizable:no;status:no")
	if OutValue.getAttribute("Done")="Y" then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "POST","XMLSave.asp?Name=UsageSelection&SessionFlag=true",false
		objhttp.send RefData.XMLDocument
		document.formname.action = "DirectIssueItemEntry.asp"
		document.formname.submit()
	end if
End Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="LoadData();SetDate('<%=FormatDate(Date())%>');init();">

<form method="POST" name="formname">
<!--OBJECT id=penDet type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" VIEWASTEXT>
<PARAM name="Command" value="HH Version"-->
</OBJECT>
<input type=hidden name="hMRSNo" value="<%=iMRSNo%>">
<input type=hidden name="hReqDate" value="<%=sMaxDate%>">
<input type=hidden name="hMinDate" Value="<%=sMinDate%>">
<input type=hidden name="hMaxDate" Value="<%=sMaxDate%>">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hItemType" value="">
<input type=hidden name="hUserId" Value="<%=iuserId%>">
<input type=hidden name="hISSFORTYPE" value="">
<input type="hidden" name="mrs" value="<%=iMRSNo%>">
<input type="hidden" name="sAct" value="mrsIssueItemEntry.asp">
<input type="hidden" name="hUsage" value="<%=sUsageCode%>">
<input type="hidden" name="hUnit" value="<%=sUnit%>">
<input type="hidden" name="hCtr" value="">
<input type="hidden" name="hRefCodes" value="">
<input type="hidden" name="hSupplier" value="">
<input type="hidden" name="hJobWorkNo" value="">
<input type="hidden" name="hSubCon" value="">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hRefType" value="">
<input type="hidden" name="hDocType" value="">
<input type=hidden name="hRefNo" value="" >
<input type=hidden name="hRefDate" value="" >
<input type=hidden name="hAutoConsumption" value="<%=sAutoConsumption%>">
<input type=hidden name="hIssueToCode" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Material Issue Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class="TabBodywithtopline">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                                    <td class="FieldCellSub">Reference Name</td>
													<td class="FieldCellSub">
														<select name="selRefName" class="FormElem">
														<%
														    RefTypePop 2,4
														%>
														</select>
													<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click Here to Edit Usage Information" width="11" height="11" onClick="GetDetails()"></a>
													</td>
												</td>
                                                    <td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Issue Date</td>
													<td class="FieldCellSub" valign="middle">
														<object id="ctlIssDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"     codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
															<param name="_ExtentX" value="2355">
															<param name="_ExtentY" value="529">
														</object>
													</td>
												</tr>
                                                   <tr>
                                                    <td class="FieldCellSub">Reference No - Date</td>
													<td class="FieldCellSub">

														<span class="DataOnly" align=center id="RefNoDate">NA</span>


													<!--<span class="DataOnly">N/A&nbsp;</span>-->
												</td>
													<td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Created By</td>
														<td class="FieldCellSub">
															<span class="dataonly"><%=sCreatedBy%></span>
														</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Issue For</td>
													<td class="FieldCellSub" valign="top">
														<span id="UsageName" class="DataOnly"></span>

													</td>
													<td class="FieldCellSub" width="2"></td>
													<td class="FieldCellSub" width="75">Acc. Head</td>
													<td class="FieldCellSub">
														<select size="1" name="selAccHead" class="FormElem" onChange="CreateNew(this)">
															<option value="select">Select</option>
															<option value="NEW">< NEW ></option>
														</select>

<DIV class=frmBody id=idConsumption style="Z-INDEX: 1; POSITION: absolute" style="width=350;display: none" >
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td colspan=2 align="center" class=PageTitle height="20"><p align="center">Consumption Head Configuration</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Usage of Item</td>
											<td class="FieldCellSub">
												<span class="DataOnly" id="idUsage">&nbsp;</span>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Consumption Head</td>
											<td class="FieldCellSub">
												<input type="text" name="txtCHead" size="20" maxlength=50 class="FormElem">
											</td>
										</tr>
										<tr>
											<td class="FieldCell" valign="top">Account Head</td>
											<td class="FieldCellSub">
												<select size="1" name="selAcc" class="FormElem">
													<option value="select">Select</option>
													<%	'Calling the Function which populates Account Head List
														populateAccountHead
													%>
												</select>
												<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center"></td>
								<td width="100%">
									<div class="frmBody" id="frm2" style="width: 100%; height:100;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center">Consumption</td>
												<td class="ExcelHeaderCell" align="center" width="100">Account Head</td>
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
												    <input type="button" value="Done" name="B1" class="ActionButton" onClick="PopDone()">
												    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="hideDiv()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
								    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	</table>
</div>


													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Issue To</td>
													<td class="FieldCellSub" valign="top">
														<span id="selIssueFor" class="DataOnly"></span>
													</td>
													<td class="FieldCellSub" width="2"></td>
													<td class="FieldCellSub">Cost Center</td>
													<td class="FieldCellSub" valign="top">
														<select size="1" name="selCC" class="FormElem">
															<option value="select">Select</option>
														<%	'Calling the Function which populates Cost Center List
															populateCostCenter
														%>
														</select>
													</td>
												</tr>

										    <td class="FieldCellSub">Received By</td>&nbsp;
											<td class="FieldCellSub">
												<input type="text" name="txtRecBy" size="35" class="FormElem" maxlength=35 style="text-align:left">
											</td>
											<td></td>
											<td class="FieldCellSub">Type</td>
											<td class="FieldCellSub" valign="top">
												<input type="checkbox" name="chkReqType" class="FormElem" value="Returnable">Returnable
											</td>
										</tr>
										<tr>
										    <td class="FieldCellsub">Remarks</td>

										    <td class="FieldCellSub" colspan="4">
										    <textarea name="Remarks" cols="90" class="Formelem" maxlength="100"></textarea>

										    <!--td class="FieldCell">Issue Date</td>
										    <td class="ExcelInutCell">
											  <%'Response.Write InsertDatePicker("ctlIssDate")%>
											</td-->
										<!--	<td class="FieldCellSub"></td>
											<td class="FieldCellSub">Issue Type</td>
											<td class="FieldCellSub">
												<input type="checkbox" name="selIssType" class="FormElem" value="Marked">Marked
											</td>						-->
										</tr>

                                    </table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel"></td>
								<td valign="top" class="FieldCell" width="100%"><center>
                                    <div align="left">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;</td>
															<td class='GroupTitle' width="50"><p align="center">Items</td></center>
															<td class='GroupTitleRight'><p align="left">&nbsp;</td>
														</tr>
													</table>
                                                </td>
											</tr>
											<tr>
												<td class=GroupTable><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="3"> </td>
														</tr>
														<tr>
															<td class=ClearPixel width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
															<td class=FieldCell>
																<DIV class=frmBody id=frm2 style="height:260;">
																	<table  id="tblLot" border="0" cellspacing="1" class="ExcelTable" width="100%">
																		<tr>
																			<td class="ExcelHeaderCell" align="center" width="10" rowspan="3">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="3">Item Description
<!--																			<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item (s)" width="11" height="11" onClick="GetItems('<%=FormatDate(date)%>')"></a>-->
																			</td>
																			<!--td class="ExcelHeaderCell" align="center" rowspan="3">Store</td-->
																			<td class="ExcelHeaderCell" align="center" colspan="3">Quantity Availability</td>
																			<td class="ExcelHeaderCell" align="center" colspan="2">Stock</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="3">Additional<br> Details</td>
																		</tr>
																	    <tr>

																			<td class="ExcelHeaderCell" align="center">In Unit</td>
																			<td class="ExcelHeaderCell" align="center">Reserved</td>
																			<td class="ExcelHeaderCell" align="center">Transit</td>
																			<td class="ExcelHeaderCell" align="center">Issue</td>
																			<!--td class="ExcelHeaderCell" align="center">Transfer</td>
																			<td class="ExcelHeaderCell" align="center">Purchase</td-->
																			<td class="ExcelHeaderCell" align="center">Total</td>
																	    </tr>
																		<tr>

																			<td class="ExcelHeaderCell" align="center">Other Unit</td>
																			<td class="ExcelHeaderCell" align="center">On Hold</td>
																			<td class="ExcelHeaderCell" align="center">Rejected</td>
																			<!--td class="ExcelHeaderCell" align="center">Date</td>
																			<td class="ExcelHeaderCell" align="center">By Date</td-->
																			<td class="ExcelHeaderCell" align="center">By Date</td>
																			<td class="ExcelHeaderCell" align="center">UoM</td>
																		</tr>
																		<% 'refer(1---1) Row 1 %>
																		<% 'refer(1---1) Row 2 %>

																		<%' End %>

																	</table>
																</div>
															</td>
															<td class=ClearPixel width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
														</tr>
														<tr>
															<td class=MiddlePack width="267" colspan="3"></td>
														</tr>
													</table>
                                                </td>
											</tr>
										</table>
                                    </div>
								</td>
								<td align="center" class="ClearPixel"></td>
							</tr>
                            <tr>
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <!--input type="button" value="Back" name="B2" class="ActionButton" onClick="Back()"-->
                                                    <input type="button" value="Issue" name="B15" class="ActionButton" onClick="CheckSubmit('<%=formatDate(date())%>')">
                                                    <!--input type="button" value="Issue" name="B15" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(Date())%>')"-->
                                                    <!--input type="reset" value="Reset" name="B16" class="ActionButton"-->
                                                    <input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('ISSUEMGMT.ASP?ACTN=L')">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
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
	' Function to populate Usage
	Function populateUsage()
		' Declaration of variables
		Dim dcrs,sUsageCode,sUsageDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR WHERE ISSUEDFORCODE <> 'INV' ORDER BY ISSUEDFORCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUsageCode = dcrs(0)
		set sUsageDesc = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUsageCode)&""">"&trim(sUsageDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Cost Center list
	Function populateCostCenter()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT COSTCENTERHEAD,CCACCOUNTDESCRIPTION FROM VWORGCOSTCENTER WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND USEABLE = 1 ORDER BY COSTCENTERHEAD"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>

<%
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTDESCRIPTION,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sUnit) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(2)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function

	Function Issue()
		MsgBox "ok"
	End Function
%>
