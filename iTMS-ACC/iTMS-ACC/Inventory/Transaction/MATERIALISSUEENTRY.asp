<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MATERIALISSUEENTRY.asp
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
	dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sQuery
	dim iuserId,sCreatedBy,rsUser
	dim sFinPeriod,arr,sMaxDate,sMinDate,sIssFor,sAutoConsumption,sType,sTypeName
	Dim sAppRefNo,sAppRefType,sCallFrom,sIssMode,sIssEntryNo

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
	sUnit = Session("organizationcode")
	sCallFrom = Request("hCallFrom")

	sQuery = "Select isNull(AutomaticConsumptionEntry,'N') from APP_M_ApplicationSetup"
	dcrs.open sQuery,con
	if not dcrs.eof then
	    sAutoConsumption = trim(dcrs(0))
	end if
	dcrs.close

    sOrgName = session("OrgName")
	sUsage = Request.Form("selUsage")
	sIssue = Request.Form("selReqType")
	sUsageCode = Request.Form("selIssType")
	sType = Request("TYPE")
	sAppRefNo = Request("AppRefNo")
	sAppRefType = Request("AppRefType")
	sIssEntryNo =  Request("ISSNO")

	if Trim(sIssEntryNo)<>"" then
	    sIssMode = "E"
	else
	    sIssMode = "N"
	end if

	if trim(sType)="" or IsNull(sType) then
	    sType="GEN"
	end if

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

	iuserid = Session("userid")
	Set rsUser = Server.CreateObject("ADODB.RecordSet")

	sCreatedBy = session("username")

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
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="StoreDetails"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="MRData"><Root></Root></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsIssueItemDetails.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">

Function Help()
    window.open "../HelpFiles/Issue.htm","","toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px;left=10;top=10"
End Function
</script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
'---------------------------------------
Function DeleteItem()
    Set Root = OutData2.documentElement

    if Root.hasChildNodes() then
        for each ndItem in Root.childNodes
            if ndItem.nodeName="ITEMDETAILS" then
                sItemCode  = ndItem.getAttribute("ITEMCODE")
                sClassCode = ndItem.getAttribute("CLASSCODE")
                sAttList = ndItem.getAttribute("ATTRIBUTELIST")

                set objChk = eval("document.formname.chkItemZ"&cstr(sItemCode)&"Z"&cstr(sClassCode)&"Z"&cstr(replace(replace(sAttList,"#",""),":","")))
                if objChk.checked=true then
                    Root.removeChild(ndItem)
                end if
            end if
        next
    end if

    set ndRoot = OutData.documentElement
    if ndRoot.hasChildNodes() then
       for each ndItemEnt in ndRoot.childNodes
           if ndItemEnt.nodeName="Item" then
               ndRoot.removeChild ndItemEnt
           end if
       next
   end if

   DisplayTable date
End Function
'********************************************
Function AddItem()


Dim nFlag
Dim s1

	nFlag=1
	iStock = "N"
	sorgID = document.formname.hUnit.value

    sTempValWindowSize = GetWindowSizeForPopup("1")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)

	set OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt=N",OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	sAct = UCase(trim(OutValue.getAttribute("Action")))
	sQuery = trim(OutValue.getAttribute("PassQuery"))

	If not OutValue.hasChildNodes Then 	exit function

    Set ndRoot = OutData.documentElement
    Set Root = OutData2.documentElement
   ' alert(ndRoot.xml)
   ' alert(Root.xml)

	if ndRoot.hasChildNodes() then
        for each ndChild in ndRoot.childNodes
            if ndChild.nodeName="Item" then
                sItemCode = ndChild.getAttribute("ItemCode")
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
		            newElem.setAttribute "ReqQty","0"
		            newElem.setAttribute "ONLYLOT",""
		            if not CheckAvailability(sItemCode) then
		                Root.appendChild newElem
		            end if 'if not CheckAvailability(sItemCode) then
		    end if  'if ndChild.nodeName="Item" then
		next 'for each ndChild in ndRoot.childNodes
    end if 'if ndRoot.hasChildNodes() then


   ' set ndRoot = OutData.documentElement
   ' if ndRoot.hasChildNodes() then
   '     for each ndItemEnt in ndRoot.childNodes
   '         if ndItemEnt.nodeName="Item" then
   '             ndRoot.removeChild ndItemEnt
   '         end if
   '     next
   ' end if

	     DisplayTable date
OutData1.loadxml OutData2.xml

End Function
'**********************************************
Function CheckAvailability(ItemCode)
    Set Root = OutData2.documentElement
    if Root.hasChildNodes() then
        for each ndItem in Root.childNodes
            if ndItem.nodeName="ITEMDETAILS" then
                if ndItem.getAttribute("ITEMCODE")=ItemCode then
                    CheckAvailability = true
                end if
            end if
        next
    end if
End Function
''''**************************
Function ScheduleDate(obj,iQty,sAttributeList)

dim sItem,sClass,sLoc,sBin,sType,sOrgID,sAttList,sArrList,sAttID,sTempAttribute
	arrTemp = split(obj.name,"Z")
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	sLoc = arrTemp(3)
	sBin = arrTemp(4)
	sOrgID = arrTemp(5)
	sTempAttribute = sAttributeList
	sAttributeList = Replace(sAttributeList,"#","$")

	sArrAttTemp = Split(sAttributeList,",")

	For iCnt = 0 to UBound(sArrAttTemp)
	    sArrAttSub1 = Split(sArrAttTemp(iCnt),"@")
	    sArrAttribute = split(sArrAttSub1(0),"$")

	    if Trim(sArrAttribute(1))<>"0" then
	        sAttList = sAttList &","& sArrAttribute(0)
	        sAttID = sAttID &","& sArrAttribute(1)
	    end if 'if Trim(sArrAttribute(1))<>"0" then
	Next


	if trim(sAttID)<>"" then
		sAttList = Mid(sAttID,2)
	    sAttID = Mid(sAttList,2)
	else
		sAttID = ""
		sAttList = ""
	end if ' if trim(arrTemp(6))<>"" then
	sPackFlag = document.formname.hPickPackFlag.value
	if document.formname.selIssType.checked = true then
        sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName&":"&sLoc&":"&sBin&":"&iQty &":"&sType&":"&sUsage&":"& sOrgID&":"& sAttList&":"&sAttID

        set OutDataValue = showModalDialog("mrsPickSchedulePoP.asp?sTemp="&sTempValues&"&AttributeList="&sAttributeList&"&PickPackFlag="&sPackFlag,OutData1,"dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
        Set Root = OutData1.documentElement
        if Root.getAttribute("DONE")="NO" then exit function
	end if

End Function
'---------------------------------------------
Function ReturnData(ItemCode,ClassCode,EntryNo)
    set obj = eval("document.formname.chkReturnable"&ItemCode&"Z"&ClassCode&"Z"&EntryNo)
    if obj.checked = true then
        eval("divItemData"&ItemCode&"Z"&ClassCode&"Z"&EntryNo).style.display="block"
    else
        eval("divItemData"&ItemCode&"Z"&ClassCode&"Z"&EntryNo).style.display="none"
    end if
End Function
'----------------------------------------------
Function RetItemData()
    if document.formname.radItem(0).checked = true then
        sReturnItem = document.formname.radItem(0).value
    else
        sReturnItem = document.formname.radItem(1).value
    end if

End Function
'--------------------------------------------
Function MarkData()
set Obj = document.formname.selIssType
    if Obj.checked = true then
        divPickPack.style.display="block"
    else
        divPickPack.style.display="none"
    end if
    PackData()
End Function
'---------------------------------------
Function PackData()
    if document.formname.radPick(0).checked = true then
        sPickFlag = document.formname.radPick(0).value
    else
        sPickFlag = document.formname.radPick(1).value
    end if
    document.formname.hPickPackFlag.value = sPickFlag

    set sObjIssType = eval("document.formname.selIssType")

    set ndRoot = OutData2.documentElement
    if ndRoot.hasChildNodes() then
        for each ndItem in ndRoot.childNodes
            if ndItem.nodeName="ITEMDETAILS" then
                iEntryNo = ndItem.getAttribute("ENTRYNO")
                iItemCode = ndItem.getAttribute("ITEMCODE")
                iClassCode = ndItem.getAttribute("CLASSCODE")
                sRcptNum = eval("document.formname.hRcptNumA"&cstr(iClassCode)&"A"&cstr(iItemCode)&"A"&cstr(iEntryNo)).value
                if trim(sRcptNum)="N" then
                    eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).readOnly=false
                    eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).className="FormElem"
                else
                    if sObjIssType.checked = true then
                        if sPickFlag="L" then
                            eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).readOnly=false
                            eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).className="FormElem"
                        else
                            eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).readOnly=true
                            eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).className="FormElemRead"
                        end if
                    else
                        eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).readOnly=true
                        eval("document.formname.txtQtyPPX"&cstr(iClassCode)&"X"&cstr(iItemCode)&"X"&cstr(iEntryNo)).className="FormElemRead"
                    end if
                end if
            end if
        next
    end if
End Function

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
    sUsage = document.formname.hIssueToCode.value
    sType =  document.formname.hType.value
    if trim(sType)<>"SUB" then
        if document.formname.selIssueTo.selectedIndex = 0 then
            alert("Select Issue To")
                document.formname.selIssueTo.focus()
            exit function
        end if
    end if 'if trim(sType)<>"SUB" then

   ' if trim(sType)="SUB" then
    '    for iCnt = 0 to document.formname.selIssueTo.length-1
    '        if lcase(trim(document.formname.selIssueTo(iCnt).value))="party" then
    '            document.formname.selIssueTo.selectedIndex = iCnt
    '            exit for
    '        end if
    '    next
    'end if


    sIssFor = document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value
    document.formname.hIssFrom.value = document.formname.cmbIssFrom(document.formname.cmbIssFrom.selectedIndex).value

    if lcase(trim(sIssFor))=lcase("Party") then
        if trim(document.formname.hIssueToCode.value)<>"" then
            sPartyCode = document.formname.hIssueToCode.value
        else
            sPartyCode = ""
        end if 'if trim(document.formname.hIssueToCode.value)<>"" then
    end if

sDispItem = 0
sCallFrom = ""
    RefTypeSelection sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem,sCallFrom
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
                    sPartyData = split(ndChild.getAttribute("Remarks"),"-")

                    if trim(sType)="SUB" then
                        txtParty.innerText = sPartyData(1)
                        document.formname.hIssueToType.value = "Party"
                        document.formname.hIssueToCode.value = sPartyData(0)
                        set objIssTo = eval("document.formname.selIssueTo")
                        For iCnt = 0 to objIssTo.length - 1
                            if ucase(objIssTo(iCnt).value) = UCase("Party") then
                               objIssTo.selectedIndex = iCnt
                               exit for
                            end if
                        Next
                    end if

                end if
            next

            sRefCodeDate = mid(sRefCodeDate,2)
            sRefNo = mid(sRefNo,2)
            sRefDate = mid(sRefDate,2)

            RefNoDate.innerHTML = sRefCodeDate
            document.formname.hRefNo.value = sRefNo
            document.formname.hRefDate.value = sRefDate

                if sRefNo <>"" then
                    set objhttp = CreateObject("Microsoft.XMLHTTP")
                    objhttp.open "GET","InvGetItemDetForRefType.asp?RefType="&sRefType &"&RefCodes="&trim(sRefNo)&"&orgID="& document.formname.hUnit.value,false
                    objhttp.send
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
			                            newElem.setAttribute "ReqQty",ndIChild.getAttribute("Qty")
			                            newElem.setAttribute "ONLYLOT",""
			                            newElem.setAttribute "RETURNABLE","N"
                                        newElem.setAttribute "RETURNITEM","S"
			                            Root.appendChild newElem
			                    end if 'if not bItemFlag then
			                elseif trim(ndIChild.nodeName)="SubContract" then
			                    Root.appendChild ndIChild
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
			            newElem.setAttribute "ReqQty","0"
			            newElem.setAttribute "ONLYLOT",""
			            newElem.setAttribute "RETURNABLE","N"
                        newElem.setAttribute "RETURNITEM","S"
			            Root.appendChild newElem
			    end if  'if ndChild.nodeName="Item" then
			next 'for each ndChild in ndRoot.childNodes
        end if 'if ndRoot.hasChildNodes() then
    end if

    Set Root = OutData2.documentElement
    if Root.hasChildNodes() then
        for each ndSubHead in Root.childNodes
            if trim(ndSubHead.nodeName)="SubContract" then
                sProName = ndSubHead.getAttribute("ProcessName")
                sInst = ndSubHead.getAttribute("Instruct")
                sLabChar = ndSubHead.getAttribute("LabourCharge")
                sHardWaste = ndSubHead.getAttribute("HardWaste")
                sInvWaste = ndSubHead.getAttribute("InvWaste")
                exit for
            end if
        next
    end if
    sTempValue = sProName
    if trim(sLabChar)<>"" then
        sTempValue = sTempValue &",LCrg.:"&sLabChar
    end if
    if trim(sHardWaste)<>"" then
        sTempValue = sTempValue &",HW%:"& sHardWaste
    end if
    if trim(sInvWaste)<>"" then
        sTempValue = sTempValue &",IW%:"& sInvWaste
    end if
    if Trim(sType)="SUB" then
        spnProcess.innerHTML = sTempValue
    end if' if Trim(sType)="SUB" then


    DisplayTable date
    OutData1.loadxml OutData2.xml
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=MRISSUEDETAILS&SessionFlag=true", false	'
	objhttp.send OutData1.XMLDocument
End Function
'***********************************************************
Function GetAddDet(iClass,iItem,sOrgid,iEntNo)
    sType = document.formname.hType.value
     sUsage = document.formname.hIssueToCode.value
	    set Q = eval("document.formname.txtQtyPPX"&iClass&"X"&iItem&"X"&iEntNo)
	    set objRetType = eval("document.formname.selReturnZ"&trim(iItem)&"Z"&trim(iClass)&"Z"&trim(iEntNo))
	        if trim(Q.value) = "" then
			    alert("Enter Quantity")
			    exit function
		    end if
		    if trim(sType)="JWK" then
		        if objRetType.selectedIndex<=0 then
		            alert("Please select the Returnable")
		            exit function
		        end if
		    end if
		    if cdbl(Q.value) = 0 then exit function

		    if objRetType.selectedIndex=0 then
		        sReturnable = "N"
		        sReturnItem = "S"
		    elseif objRetType.selectedIndex=1 then
		        sReturnable = "Y"
		        sReturnItem = "S"
		    elseif objRetType.selectedIndex=2 then
		        sReturnable = "Y"
		        sReturnItem = "D"
		    end if

            if trim(sType)="JWK" then
		            sTempValues = iClass&"|"&iItem&"|"&sOrgid&"|"&sUsage&"|"&trim(Q.value)&"|"&sItemMode&"|"&sReturnable&"|"&sReturnItem&"|"&iEntNo

	                set OutValue = showModalDialog("IssSubConProcessDetailsPop.asp?sTemp="&sTempValues,OutData1,"dialogHeight:370px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No")
                    'alert(OutValue.xml)
            else
		            sTempValues = iClass&"|"&iItem&"|"&sOrgid&"|"&sUsage&"|"&trim(Q.value)

	            set OutValue = showModalDialog("DirectIssueAddEntry.asp?sTemp="&sTempValues,OutData2,"dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	        end if 'if trim(sType)="SUB" then
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

			If not rootData.hasChildNodes Then 	exit function
			sIType = trim(rootData.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
			sTemp1 = sIType
		else
			if trim(document.formname.hIssueToCode.value)="PRD" then
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
				if root.hasChildNodes then

					sExp ="//ITEMDETAILS [@ENTRYNO = "&ndTemp.attributes.getNamedItem("EntryNo").Value&" and @ITEMCODE = "&ndTemp.attributes.getNamedItem("ItemCode").value&" and @CLASSCODE = "&ndTemp.attributes.getNamedItem("ClassCode").value&"]"
					Set CheckNode = Root.Selectnodes(sExp)
					if CheckNode.Length = 0 then
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
						newElem.setAttribute "ONLYLOT",""

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
					newElem.setAttribute "ONLYLOT",""
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
	Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

    sTempValWindowSize = GetWindowSizeForPopup("1")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)

Set objhttp = CreateObject("Microsoft.XMLHTTP")
	if 1 = 1 then
		sorgID = document.formname.hUnit.value
		sIType = document.formname.hItemType.value
	nFlag = 1
		set Root = OutSelectData.DocumentElement
		iStock = "Y"
		if trim(document.formname.hIssueToCode.value)<>"PRD" then
			set OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt=N",OutSelectData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			sQuery = trim(OutValue.getAttribute("PassQuery"))
			sIType = trim(OutValue.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
			If not OutValue.hasChildNodes Then 	exit function
			sIType = trim(OutValue.getAttribute("ItemType"))
			document.formname.hItemType.value = sIType
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
			sTemp1 = sIType
		end if ' if trim(document.formname.selUsage.value)<>"PRD" then

		If not OutValue.hasChildNodes Then 	exit function
		set rootData = OutSelectData.DocumentElement
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
				if root.hasChildNodes then

					sExp ="//ITEMDETAILS [@ENTRYNO = "&ndTemp.attributes.getNamedItem("EntryNo").Value&" and @ITEMCODE = "&ndTemp.attributes.getNamedItem("ItemCode").value&" and @CLASSCODE = "&ndTemp.attributes.getNamedItem("ClassCode").value&"]"
					Set CheckNode = Root.Selectnodes(sExp)
					if CheckNode.Length = 0 then
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
						newElem.setAttribute "ONLYLOT",""

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
					newElem.setAttribute "ONLYLOT",""
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
'***************************************************

Function DisplayTable(todaysDate)
    Dim sRefType
	iCounter1 = 0
	set rootData = ItemData.DocumentElement
	set root = OutData2.DocumentElement
	set oNodTaxRoot = PurTypeData.DocumentElement
	set rootUoM = UoMData.DocumentElement
	IF document.formname.hCtr.value = "" then document.formname.hCtr.value = 0
	IF document.formname.hCtr.value <> 0 then 	ClearTable()

    sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value

	sUnit = document.formname.hUnit.value
	'alert(root.xml)
	if root.hasChildNodes then
		sExp ="//ITEMDETAILS [ @DISPALYED = 'N']"
		Set CheckNode = Root.Selectnodes(sExp)
		if CheckNode.Length > 0 then
			iCounter = 0

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
			 	iRequested =CheckNode.Item(iCounter).Attributes.getNamedItem("ReqQty").value
			 	iIssQty = CheckNode.Item(iCounter).Attributes.getNamedItem("QTY").value
			 	sReturnable = CheckNode.Item(iCounter).Attributes.getNamedItem("RETURNABLE").value
			 	sReturnItem = CheckNode.Item(iCounter).Attributes.getNamedItem("RETURNITEM").value
				sOptName=""
			 	sStoreCode = 0
				sBinCode = 0

				set objhttp = CreateObject("Microsoft.XMLHTTP")
				objhttp.open "GET","../../Common/GetItemRcptNumbering.asp?ItemCode="&iItem,false
				objhttp.send
				if trim(objhttp.responseText)<>"" then
				    sRcptNumbering = trim(objhttp.responseText)
				end if
				'alert(sAttributeList)
				AttTemp = split(sAttributeList,",")
				    IF sAttributeList <> "" then
	                    For k = 0 to UBOUND(AttTemp)
	                        nTemp   = AttTemp(k)
		                    nTempVal  = split(nTemp,":")
		                    nVal  = nTempVal(0)
		                    nValTemp = split(nVal,"#")
		                    if UBound(nValtemp)=0 then
		                        set objhttp=CreateObject("Microsoft.XMLHTTP")
				                objhttp.open "GET","../../Include/GetAttrName.asp?AttID="&sAttributeList,false
				                objhttp.send
				                'alert(objhttp.responseText)
				                if Trim(objhttp.responseText)<>"" then
				                    sOptName = sOptName &"["& Trim(objhttp.responseText) &"]"
				                end if
				            else
				                IF trim(nValTemp(1)) = "0" then
		                            if Trim(sOptName)="" then
			                            sOptName = ""
			                        end if
		                        Else
		                            sOptName = sOptName &"["& nTempVal(1)&"]"
		                        End IF
				            end if 'if UBound(nValtemp)=0 then
	                    Next
	                End If


				'To Get Stock Details

				sTemp = iItem &":"& iEntNo &":"& iClass &":"& document.formname.hUnit.value&":"&document.formname.hMinDate.value&":"&document.formname.hMaxDate.value&":"&Replace(Replace(sAttributeList,":","@"),"#","$")

				objhttp.Open "GET","XMLGetStockDetails.asp?Para="&sTemp,false
				objhttp.send
                'alert(objhttp.responseText)
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

				sTempValues = "orgID="&document.formname.hUnit.value&"&RefCodes="&document.formname.hRefNo.value&"&RefType="&sRefType&"&ItemCode="&iItem&"&ClassCode="&iClass
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

			 	if trim(iToIssue)="" or IsNull(iToIssue) then iToIssue = "0"
				''''''''''''''''''''
				j = j + 1
				iser = iser + 1

				set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length - cint(iCtr))
				'Serial no
				set headerCell=oRow.insertCell()
				headerCell.innerText = iSer
				headerCell.className="ExcelSerial"
				headerCell.align="center"

				'Checkbox
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox"" name=""chkItemZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(replace(replace(sAttributeList,"#",""),":",""))&""" class=""FormElem"" style=""text-align=right"">" )
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"


				set headerCell=oRow.insertCell()
				sIssMode = document.formname.hIssMode.value
				if trim(sIssMode)="E" then
				    sItemDesc=sItemDesc & sOptName
				else
				    if Trim(sOptName)<>"" then
				        sItemDesc=sItemDesc& sOptName
				    else
				        sItemDesc = sItemDesc
				    end if
				end if
				headerCell.innerHTML="<a class='ExcelDisplayLink' href=#  name=""lnkA"&CStr(iItem)&"A"&CStr(iClass)&"A"&CStr(sOrgId)&""" onClick=""javascript:DisplayItem(this.name)"">" & sItemDesc & "</a>"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"

				'Returnbable '' added on Oct 17,2013 by ragav
				''begin
				set headerCell=oRow.insertCell()
				'headerCell.innerHtml = headerCell.innerHtml &"<table><tr><td class='FieldCell'><input type=""checkbox"" name=""chkReturnable"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" onClick=""ReturnData("&CStr(iItem)&","&CStr(iClass)&","&CStr(iEntNo)&")"">&nbspReturnable</td><td class='FieldCell'><div id=""divItemData"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" style=""display:none;""><input type=""radio"" name=""radItem"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" value=""S"" checked>Same<input type=""radio"" name=""radItem"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" value=""D"">Diff.</div></td>"
				headerCell.innerHtml="<Select id=""selReturnZ"&Cstr(iItem)&"Z"&Cstr(iClass)&"Z"&Cstr(iEntNo)&""" class=""FormElem""><option value=""N"">No</option><option value=""Y"">Yes</option><option value=""D"">Diff.</option></Select>"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"
				''end


				set objReturn = eval("document.formname.selReturnZ"&Cstr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo))
				For iCnt = 0 to objReturn.length-1
				    if trim(sReturnable)=trim(objReturn(iCnt).value) then
				        if Trim(sReturnItem)="S" then
				            objReturn.selectedIndex = iCnt
				        else
				            objReturn.selectedIndex = iCnt + 1
				        end if
				        exit for
				    end if
				Next

				'Requested
				set headerCell=oRow.insertCell()
				headercell.innerHtml = trim(iRequested)&"["&trim(iToIssue)&"]"
				set oText = document.createElement("<input type=""hidden"" name=""txtRequestedZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" Value="""&trim(iRequested)&""" size=""5"" class=""FormElemNone"" style=""text-align=right;cursor:hand;FONT-WEIGHT: bold"" alt=""In Unit Stock Details"" READONLY>")
				headerCell.appendChild(oText)
				set oText = document.createElement("<input type=""hidden"" name=""txtToIssueZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" Value="""&trim(iToIssue)&""" size=""5"" class=""FormElemNone"" style=""text-align=right;cursor:hand;FONT-WEIGHT: bold"" alt=""To Issue Details"" READONLY>")
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'In Unit
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtStockZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" Value="""&iStock&""" size=""5"" class=""FormElemNone"" style=""text-align=right;cursor:hand;FONT-WEIGHT: bold"" alt=""In Unit Stock Details"" READONLY>")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				sOptName = Replace(sOptName,Chr(39),"~~")
				sAttributeList =  Replace(sAttributeList,Chr(39),"~~")

				'Issue 'refer(1---1)
				set headerCell=oRow.insertCell()

				headerCell.innerHtml = "<a><img name=""btnZ"&CStr(iClass)&"Z"&CStr(iItem)&"Z"&CStr(iEntNo)&"Z"&cstr(sOptName)&"Z"& CStr(sOrgId)&"Z"&cstr(sAttributeList)&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Pick Details"" onClick=""CheckLot(this,'','"&replace(sAttributeList,":","@")&"')""></a>"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""hidden"" name=""hStoName"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" Value=""0"">"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""hidden"" name=""hSchA"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" Value="""">"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""hidden"" name=""hStoA"&sStoreCode&"A"&sBinCode&""" Value=""0"">"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""hidden"" name=""hRcptNumA"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" Value="""&sRcptNumbering&""">"
				headerCell.innerHtml = headerCell.innerHtml + "<input type=""text"" name=""txtQtyPPX"&CStr(iClass)&"X"&CStr(iItem)&"X"&CStr(iEntNo)&""" size=""10"" class=""FormElemRead"" disabled onkeypress=""DoKeyPress('',7,3)"" style=""text-align:right"" value="""& iIssQty &""" onBlur=""GetSch(this)"">"

				headerCell.width = "91"
				headerCell.className="ExcelInputCell"
				headerCell.align="left"

				'UOM
				set headerCell=oRow.insertCell()
				headerCell.innerHtml = sUOM
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				'Date
				set headerCell=oRow.insertCell()
				headerCell.innerHtml = "<a><img name=""btnScheduleZ"&CStr(iClass)&"Z"&CStr(iItem)&"Z"&CStr(iEntNo)&"Z"&cstr(sOptName)&"Z"& CStr(sOrgId)&"Z"&cstr(sAttributeList)&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" alt=""Pick Schedule"" onClick=""ScheduleDate(this,'','"&replace(sAttributeList,":","@")&"')""></a>"
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"


				sType = document.formname.hType.value
				if trim(sType)="SUB" then
				    set headerCell = oRow.insertCell()
				   ' headerCell.innerHtml = "<input type='radio' name=""radMatTypeZ"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" value='P'>Primary<br/>"
				   ' headerCell.innerHtml = headerCell.innerHtml & " <input type='radio' name=""radMatTypeZ"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" value='C'>Consumable<br/>"
				    'headerCell.innerHtml = headerCell.innerHtml & " <input type='radio' name=""radMatTypeZ"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" value='A'>Accessories<br/>"
				    headerCell.innerHtml = "N/A"
				    headerCell.className="ExcelDisplayCell"
				else
				    set headerCell = oRow.insertCell()
				    set oText = document.createElement("<input type='button' name=""btnAddDet"&cstr(iClass)&"A"&cstr(iItem)&"A"&cstr(iEntNo)&""" class='ActionButtonX' value='Yes' onClick='GetAddDet("""+ iClass +""","""+ iItem+""","""+ sOrgid + ""","""+iEntNo+""")'>")
				    headerCell.appendChild(oText)
				    headerCell.className="ExcelDisplayCell"
				    headercell.align="center"
				end if

				document.formname.hCtr.value = document.all.tblLot.rows.length + 1
			Next
		end if

	end if
    MarkData()
End Function

Function ClearTable()
	Dim i
		 K = document.all.tblLot.rows.length - 2
	for	i = 1 to  K
		document.all.tblLot.deleteRow(2)
	next

End function

Function popAC()
	document.formname.selAccHead.length = 2

	dim Root,HeaderNode
	sIssuedFor = document.formname.hIssueToCode.value
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
		sIssFor = document.formname.hIssueToCode.value

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
'===============================================================
Function EditUsageInfo()
	Dim OutValue,objhttp
	set OutValue =  showModalDialog("IssueUsageSelPop.asp?OrgID="& document.formname.hUnit.value,RefData,"dialogHeight:340px;dialogwidth=500px;center:yes;help:no;resizable:no;status:no")
	if OutValue.getAttribute("Done")="Y" then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "POST","XMLSave.asp?Name=UsageSelection&SessionFlag=true",false
		objhttp.send RefData.XMLDocument
		document.formname.action = "MaterialItemEntry.asp"
		document.formname.submit()
	end if
End Function
'*****************************************************
Function SelectProcess()
sOrgCode =  document.formname.hUnit.value
    set OutValue = showModalDialog("IssSubConProcessSelPop.asp?Unit="&sOrgCode,OutData1,"dialogHeight:370px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No")
    set ndRoot = OutData1.documentElement
    if ndRoot.hasChildNodes() then
        for each ndSubHead in ndRoot.childNodes
            if trim(ndSubHead.nodeName)="SubContract" then
                sProName = ndSubHead.getAttribute("ProcessName")
                sInst = ndSubHead.getAttribute("Instruct")
                sLabChar = ndSubHead.getAttribute("LabourCharge")
                sHardWaste = ndSubHead.getAttribute("HardWaste")
                sInvWaste = ndSubHead.getAttribute("InvWaste")
                exit for
            end if
        next
    end if
    sTempValue = sProName
    if trim(sLabChar)<>"" then
        sTempValue = sTempValue &",LCrg.:"&sLabChar
    end if
    if trim(sHardWaste)<>"" then
        sTempValue = sTempValue &",HW%:"& sHardWaste
    end if
    if trim(sInvWaste)<>"" then
        sTempValue = sTempValue &",IW%:"& sInvWaste
    end if
    spnProcess.innerHTML = sTempValue
End Function
'******************************************************
Function Init()
Dim sRefType,sOrgID,sItemType,sPartyCode,iStock,nFlag,bAddButton
sAppRefType = document.formname.hAppRefType.value
sAppRefNo = document.formname.hAppRefNo.value
sIssMode = document.formname.hIssMode.value
sIssEntryNo = document.formname.hIssEntryNo.value
if Trim(sIssMode)="E" then

    set objhttp = CreateObject("Microsoft.XMLHTTP")
        objhttp.open "GET","XMLGetIssueDetails.asp?IssEntNo="&sIssEntryNo,false
        objhttp.send
        if Trim(objhttp.responseXML.xml)<>"" then
            OutData2.loadXML(objhttp.responseXML.xml)
        else
            alert(objhttp.responseText)
        end if

        set root = OutData2.documentElement
        if root.hasChildNodes() then
            sAppRefType = root.getAttribute("AppRefType")
            sAppRefNo = root.getAttribute("AppRefNo")
        end if

        set objRef = eval("document.formname.selRefName")
        For iCnt = 0 to objRef.length-1
            if trim(objRef(iCnt).value)=sAppRefType then
                objRef.selectedIndex = iCnt
            end if
        Next

            sorgID = document.formname.hUnit.value
            sItemType = ""
            sPartyCode = ""
            iStock = "Y"
            nFlag=1
            bAddButton = "N"

            sUsage = document.formname.hIssueToCode.value
            sType =  document.formname.hType.value

            if Trim(sAppRefType)<>"N" and Trim(sAppRefType)<>"" and Trim(sAppRefType)<>"0" then

                set objhttp = CreateObject("Microsoft.XMLHTTP")
                    objhttp.open "GET","../../Common/GetInfoForRefType.asp?RefType="&sAppRefType &"&RefNo="&trim(sAppRefNo)&"&orgID="& sorgID,false
                    objhttp.send
                    if trim(objhttp.responseXML.xml)<>"" then
                        MRData.loadXML(objhttp.responseXML.xml)
                    else
                        alert(Objhttp.responseText)
                    end if

                    set ndMRRoot = MRData.documentElement
                    if ndMRRoot.hasChildNodes() then
                        for each ndChild in ndMRRoot.childNodes
                            if ndChild.nodeName="Ref" then
                                sRefCode = trim(ndChild.getAttribute("Code"))
                                sRefDate = sRefDate &","& ndChild.getAttribute("Date")
                                sRefNo = sRefNo &","& ndChild.getAttribute("No")
                                sRefDt = ndChild.getAttribute("Date")
                                sRefCodeDate =  sRefCodeDate &","& sRefCode &" - "& sRefDt
                                sPartyData = split(ndChild.getAttribute("Remarks"),"-")

                                if trim(sType)="SUB" then
                                    txtParty.innerText = sPartyData(1)
                                    document.formname.hIssueToType.value = "Party"
                                    document.formname.hIssueToCode.value = sPartyData(0)
                                    set objIssTo = eval("document.formname.selIssueTo")
                                    For iCnt = 0 to objIssTo.length - 1
                                        if ucase(objIssTo(iCnt).value) = UCase("Party") then
                                           objIssTo.selectedIndex = iCnt
                                           exit for
                                        end if
                                    Next
                                end if

                            end if
                        next

                        sRefCodeDate = mid(sRefCodeDate,2)
                        sRefNo = mid(sRefNo,2)
                        sRefDate = mid(sRefDate,2)

                        RefNoDate.innerHTML = sRefCodeDate
                        document.formname.hRefNo.value = sAppRefNo
                        document.formname.hRefDate.value = sAppRefDate
                    end if
            end if'if Trim(sAppRefType)<>"N" and Trim(sAppRefType)<>"" then

            sIssFor = document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value
            document.formname.hIssFrom.value = document.formname.cmbIssFrom(document.formname.cmbIssFrom.selectedIndex).value


            if trim(sType)<>"SUB" then
                sIssToType = root.getAttribute("ISSTOTYPE")
                sIssToCode = root.getAttribute("ISSTOCODE")
                sIssToSubCode = Root.getAttribute("ISSTOSUBCODE")
                document.formname.hIssueToType.value = sIssToType
                document.formname.hIssueToCode.value = sIssToCode
                document.formname.hIssueToSubCode.value = sIssToSubCode
                set objIssTo = eval("document.formname.selIssueTo")
                For iCnt = 0 to objIssTo.length - 1
                    if ucase(objIssTo(iCnt).value) = UCase(sIssToType)&":"&UCase(sIssToCode) then
                       objIssTo.selectedIndex = iCnt
                       exit for
                    end if
                Next
            end if

            if lcase(trim(sIssFor))=lcase("Party") then
                if trim(document.formname.hIssueToCode.value)<>"" then
                    sPartyCode = document.formname.hIssueToCode.value
                else
                    sPartyCode = ""
                end if 'if trim(document.formname.hIssueToCode.value)<>"" then
            end if

        sDispItem = 0
        sCallFrom = ""
        Set Root = OutData2.documentElement
        if Root.hasChildNodes() then
            for each ndSubHead in Root.childNodes
                if trim(ndSubHead.nodeName)="SubContract" then
                    sProName = ndSubHead.getAttribute("ProcessName")
                    sInst = ndSubHead.getAttribute("Instruct")
                    sLabChar = ndSubHead.getAttribute("LabourCharge")
                    sHardWaste = ndSubHead.getAttribute("HardWaste")
                    sInvWaste = ndSubHead.getAttribute("InvWaste")
                    exit for
                end if
            next
        end if
        sTempValue = sProName
        if trim(sLabChar)<>"" then
            sTempValue = sTempValue &",LCrg.:"&sLabChar
        end if
        if trim(sHardWaste)<>"" then
            sTempValue = sTempValue &",HW%:"& sHardWaste
        end if
        if trim(sInvWaste)<>"" then
            sTempValue = sTempValue &",IW%:"& sInvWaste
        end if
        if Trim(sType)="SUB" then
            spnProcess.innerHTML = sTempValue
        end if' if Trim(sType)="SUB" then


            DisplayTable date
            OutData1.loadxml OutData2.xml
	        set objhttp = CreateObject("Microsoft.XMLHTTP")
	        objhttp.Open "POST","XMLSave.asp?Name=MRISSUEDETAILS&SessionFlag=true", false	'
	        objhttp.send OutData1.XMLDocument
else
    if trim(sAppRefNo)<>"" then
        set objRef = eval("document.formname.selRefName")
        For iCnt = 0 to objRef.length-1
            if trim(objRef(iCnt).value)=sAppRefType then
                objRef.selectedIndex = iCnt
            end if
        Next

            sorgID = document.formname.hUnit.value
            sItemType = ""
            sPartyCode = ""
            iStock = "Y"
            nFlag=1
            bAddButton = "N"

            set Root = RefData.documentElement
            sUsage = document.formname.hIssueToCode.value
            sType =  document.formname.hType.value


                set objhttp = CreateObject("Microsoft.XMLHTTP")
                    objhttp.open "GET","../../Common/GetInfoForRefType.asp?RefType="&sAppRefType &"&RefNo="&trim(sAppRefNo)&"&orgID="& sorgID,false
                    objhttp.send
                    if trim(objhttp.responseXML.xml)<>"" then
                        MRData.loadXML(objhttp.responseXML.xml)
                    else
                        alert(Objhttp.responseText)
                    end if

                    set ndMRRoot = MRData.documentElement
                    if ndMRRoot.hasChildNodes() then
                        for each ndChild in ndMRRoot.childNodes
                            if ndChild.nodeName="Ref" then
                                sRefCode = trim(ndChild.getAttribute("Code"))
                                sRefDate = sRefDate &","& ndChild.getAttribute("Date")
                                sRefNo = sRefNo &","& ndChild.getAttribute("No")
                                sRefDt = ndChild.getAttribute("Date")
                                sRefCodeDate =  sRefCodeDate &","& sRefCode &" - "& sRefDt
                                sPartyData = split(ndChild.getAttribute("Remarks"),"-")

                                if trim(sType)="SUB" then
                                    txtParty.innerText = sPartyData(1)
                                    document.formname.hIssueToType.value = "Party"
                                    document.formname.hIssueToCode.value = sPartyData(0)
                                    set objIssTo = eval("document.formname.selIssueTo")
                                    For iCnt = 0 to objIssTo.length - 1
                                        if ucase(objIssTo(iCnt).value) = UCase("Party") then
                                           objIssTo.selectedIndex = iCnt
                                           exit for
                                        end if
                                    Next
                                end if

                            end if
                        next

                        sRefCodeDate = mid(sRefCodeDate,2)
                        sRefNo = mid(sRefNo,2)
                        sRefDate = mid(sRefDate,2)

                        RefNoDate.innerHTML = sRefCodeDate
                        document.formname.hRefNo.value = sAppRefNo
                        document.formname.hRefDate.value = sAppRefDate
                    end if

            sIssFor = document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value
            document.formname.hIssFrom.value = document.formname.cmbIssFrom(document.formname.cmbIssFrom.selectedIndex).value

            if lcase(trim(sIssFor))=lcase("Party") then
                if trim(document.formname.hIssueToCode.value)<>"" then
                    sPartyCode = document.formname.hIssueToCode.value
                else
                    sPartyCode = ""
                end if 'if trim(document.formname.hIssueToCode.value)<>"" then
            end if

        sDispItem = 0
        sCallFrom = ""
            Set ndRoot = OutData.documentElement
            Set Root = OutData2.documentElement
            if sAppRefNo <>"" then
                set objhttp = CreateObject("Microsoft.XMLHTTP")
                objhttp.open "GET","InvGetItemDetForRefType.asp?RefType="&sAppRefType &"&RefCodes="&trim(sAppRefNo)&"&orgID="&sorgid,false
                objhttp.send
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
                                    newElem.setAttribute "ReqQty",ndIChild.getAttribute("Qty")
                                    newElem.setAttribute "ONLYLOT",""
                                    newElem.setAttribute "RETURNABLE","N"
                                    newElem.setAttribute "RETURNITEM","S"
                                    Root.appendChild newElem
                            end if 'if not bItemFlag then
                        elseif trim(ndIChild.nodeName)="SubContract" then
			                    Root.appendChild ndIChild
                        end if  'if ndChild.nodeName="Item" then
                    next
                end if
            end if 'if sRefNo <>"" then


             Set Root = OutData2.documentElement
        if Root.hasChildNodes() then
            for each ndSubHead in Root.childNodes
                if trim(ndSubHead.nodeName)="SubContract" then
                    sProName = ndSubHead.getAttribute("ProcessName")
                    sInst = ndSubHead.getAttribute("Instruct")
                    sLabChar = ndSubHead.getAttribute("LabourCharge")
                    sHardWaste = ndSubHead.getAttribute("HardWaste")
                    sInvWaste = ndSubHead.getAttribute("InvWaste")
                    exit for
                end if
            next
        end if
        sTempValue = sProName
        if trim(sLabChar)<>"" then
            sTempValue = sTempValue &",LCrg.:"&sLabChar
        end if
        if trim(sHardWaste)<>"" then
            sTempValue = sTempValue &",HW%:"& sHardWaste
        end if
        if trim(sInvWaste)<>"" then
            sTempValue = sTempValue &",IW%:"& sInvWaste
        end if
        if Trim(sType)="SUB" then
            spnProcess.innerHTML = sTempValue
        end if' if Trim(sType)="SUB" then

            DisplayTable date
            OutData1.loadxml OutData2.xml
	        set objhttp = CreateObject("Microsoft.XMLHTTP")
	        objhttp.Open "POST","XMLSave.asp?Name=MRISSUEDETAILS&SessionFlag=true", false	'
	        objhttp.send OutData1.XMLDocument
    end if'if trim(sAppRefNo)<>"" then
end if 'if trim(sIssMode)="E" then
End Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin="0" MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();LoadData();SetDate('<%=FormatDate(Date())%>');">

<form method="POST" name="formname">
<!--OBJECT id=penDet type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" VIEWASTEXT>
<PARAM name="Command" value="HH Version">
</OBJECT-->
<input type=hidden name="hMRSNo" value="<%=iMRSNo%>">
<input type=hidden name="hReqDate" value="<%=sMaxDate%>">
<input type=hidden name="hMinDate" Value="<%=sMinDate%>">
<input type=hidden name="hMaxDate" Value="<%=sMaxDate%>">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hItemType" value="">
<input type=hidden name="hUserId" Value="<%=iuserId%>">

<input type="hidden" name="mrs" value="<%=iMRSNo%>">
<input type="hidden" name="sAct" value="mrsIssueItemEntry.asp">
<input type="hidden" name="hUnit" value="<%=sUnit%>">
<input type="hidden" name="hCtr" value="">
<input type="hidden" name="hRefCodes" value="">

<input type="hidden" name="hJobWorkNo" value="">
<input type="hidden" name="hSubCon" value="">
<input type="hidden" name="hRefType" value="">
<input type="hidden" name="hDocType" value="">
<input type=hidden name="hRefNo" value="" >
<input type=hidden name="hRefDate" value="" >
<input type=hidden name="hAutoConsumption" value="<%=sAutoConsumption%>">
<input type="hidden" name="hIssueToType" value="">
<input type="hidden" name="hIssueToCode" value="">
<input type="hidden" name="hIssueToSubCode" value="">
<input type="hidden" name="hPickPackFlag" value="">
<input type="hidden" name="hIssFrom" value="">

<input type="hidden" name="hType" value="<%=sType%>" />
<input type="hidden" name="hAppRefNo" value="<%=sAppRefNo%>" />
<input type="hidden" name="hAppRefType" value="<%=sAppRefType%>" />
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>" />
<input type="hidden" name="hIssEntryNo" value="<%=sIssEntryNo%>" />
<input type="hidden" name="hIssMode" value="<%=sIssMode%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%

    Select Case trim(sType)
        Case "GEN"
            sTypeName = "General"
        Case "SUB"
            sTypeName = "Subcontract"
        Case "SER"
            sTypeName = "Services"
        Case "JWK"
            sTypeName = "Job Work"
        Case "TRN"
            sTypeName = "Transfer"
        Case "POS"
            sTypeName = "POS Consumption"
    End Select

%>
	<tr><td height="1px"></td></tr>
	<tr>
		<td>
		    <table>
	            <tr>
	                <td class="PageTitle" >
	                    Material Issue (<%=sTypeName%>)
	                </td>
	                <td class="PageTitle" >
	                    <a style="text-decoration:none;font:color:black" href="#" onclick="Help()">Help</a>
	                </td>
	            </tr>
	        </table>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing=0 cellPadding=0 border=0 width="100%">
		            <tr>
					    <td height="20" valign="bottom">
						    <table border="0" cellpadding="0" cellspacing="0" >
							    <tr>
							        <td class="TabCell" valign="bottom" width="90">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="IssueMGMT.asp">
											    <td align="center">List
											    </td></a>
										    </tr>
									    </table>
								    </td>
							   	    <td class="TabCurrentCell" valign="bottom" align="center" width="50">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr>
												<td align="center">Basic
												</td>
											</tr>
										</table>
									</td>
								    <td class="TabCell" valign="bottom" width="145">
								        <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
									        <tr><a href="IssReturnEntry.asp">
										        <td align="center">Return
										        </td></a>
									        </tr>
								        </table>
							        </td>
								    <td class="TabCellEnd" valign="bottom" align="left">
									        &nbsp;
							        </td>
							    </tr>
						    </table>
					    </td>
            	    </tr>
			    <TR>
					<TD class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td class="FieldCellSub">Issue From</td>
													<td class="FieldCellSub" valign="top">
													    <select name="cmbIssFrom" class="FormElem" disabled="true">
													        <%
													               sQuery ="Select ApplnShortName,ApplicationName from MS_Applications"
														           dcrs.open sQuery,con
														           if not dcrs.eof then
														                do while not dcrs.eof
														                    if trim(dcrs(0))="IN" then
														                        Response.Write "<option value="& trim(dcrs(0)) &" selected>"& trim(dcrs(1))&"</option>"
														                    else
														                        Response.Write "<option value="& trim(dcrs(0)) &">"& trim(dcrs(1))&"</option>"
														                    end if
														                    dcrs.movenext
														                loop
														           end if
														           dcrs.close
													        %>
													    </select>
														<span id="UsageName" class="DataOnly">
														<%
														'   sQuery ="Select IssuedForDescription from Inv_M_IssuedFor where IssuedForCode = '"& sIssFor &"'"
														'   dcrs.open sQuery,con
														'   if not dcrs.eof then
														'        Response.Write trim(dcrs(0))
														'   end if
														'   dcrs.close
														%>
														</span>
													</td>

                                                    <td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Issue To</td>
													<td class="FieldCellSub" valign="top">
														<select name="selIssueTo" class="FormElem" onchange("issue") onChange="popIssueTo()">
														    <option value="0">Select</option>
													    <%
														    populateIssueTo sUnit,sType
													    %>
													    </select><br />
													    <span id="txtParty" class="DataOnly"></span>
													</td>

												</tr>
                                                <tr>
                                                   <td class="FieldCellSub">Reference Name</td>
													<td class="FieldCellSub">
														<select name="selRefName" class="FormElem" Onchange="GetDetails()">
														<%
														    RefTypePop 2,4
														%>
														</select>
													<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click Here to Edit Usage Information" width="11" height="11" onClick="GetDetails()"></a>
													</td>
                                                    <td class="FieldCellSub"></td>
                                                   <td class="FieldCellSub">Issue Date</td>
													<td class="FieldCellSub" valign="middle">
														<object id="ctlIssDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"     codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
															<param name="_ExtentX" value="2355">
															<param name="_ExtentY" value="529">
														</object>
													</td>
												</tr>

												<tr>
                                                    <td class="FieldCellSub">Reference No - Date</td>
													<td class="FieldCellSub">
														<span class="DataOnly" align=center id="RefNoDate">NA</span>
    												</td>
                                                    <td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Created By</td>
														<td class="FieldCellSub">
															<span class="dataonly"><%=sCreatedBy%></span>
														</td>
												</tr>
												<tr>
												    <td class="FieldCellSub">Cost Center</td>
													<td class="FieldCellSub" valign="top">
														<select size="1" name="selCC" class="FormElem">
															<option value="select">Select</option>
														<%	'Calling the Function which populates Cost Center List
															populateCostCenter
														%>
														</select>
													</td>
												   	<td class="FieldCellSub" width="2"></td>
													<td class="FieldCellSub" width="75">Acc. Head</td>
													<td class="FieldCellSub">
														<select size="1" name="selAccHead" class="FormElem" onChange="CreateNew(this)">
															<option value="select">Select</option>
															<option value="NEW">< NEW ></option>
															<%
															    populateConsumptionAccHead
															%>
														</select>
													</td>
												</tr>

												<tr>
												    <td class="FieldCellSub"></td>
													<td class="FieldCellSub" valign="top">
<DIV class=frmBody id=idConsumption style="Z-INDEX: 1; POSITION: absolute;width:350;display: none" >
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td valign="top">
			<table id="table1" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				<tr>
					<td class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
						    <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
							    <td align="center"></td>
								<td align="center" class=PageTitle height="20"><p align="center">Consumption Head Configuration</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Usage of Item</td>
											<td class="FieldCellSub">
												<span class="DataOnly" id="Span1">
												<%
														    if trim(sIssFor)="CON" then
														        Response.Write "General Consumption"
														    elseif trim(sIssFor)="DIS" then
														        Response.Write "For Despatch"
														    elseif trim(sIssFor)="IUT" then
														        Response.Write "Inter-Unit transfers"
														    elseif trim(sIssFor)="JWK" then
														        Response.Write "Job Work"
														    elseif trim(sIssFor)="MAT" then
														        Response.Write "Maintenance Consumption"
														    elseif trim(sIssFor)="OTH" then
														        Response.Write "Others"
														    elseif trim(sIssFor)="PRD" then
														        Response.Write "Production Consumption"
														    elseif trim(sIssFor)="PUR" then
														        Response.Write "Purchase Returns"
														    elseif trim(sIssFor)="REP" then
														        Response.Write "Replacement returns"
														    elseif trim(sIssFor)="SAL" then
														        Response.Write "Sales Returns"
														    elseif trim(sIssFor)="SER" then
														        Response.Write "Services"
														    elseif trim(sIssFor)="SUB" then
														        Response.Write "Subcontracting"
														    end if
														%>
												</span>
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
									<div class="frmBody" id="Div2" style="width: 100%; height:100;">
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
										    <td class="FieldCellSub">Received By</td>&nbsp;
											<td class="FieldCellSub">
												<input type="text" name="txtRecBy" size="35" class="FormElem" maxlength="35" style="text-align:left"/>
											</td>
											<td></td>
											<td class="FieldCellSub">Issue Type</td>
											<td class="FieldCell">
											    <table>
											        <tr>
											            <td class="FieldCellSub">
											                 <input type="checkbox" name="selIssType" class="FormElem" value="Marked" onclick="MarkData()">Marked
											            </td>
											            <td class="FieldCell">
											                <div id="divPickPack" style="display:none;">
												                <input type="radio" name="radPick" value="N" onclick="PackData()">Pick Pack Now
												                <input type="radio" name="radPick" value="L" onclick="PackData()" checked>Later
												            </div>
											            </td>
											        </tr>
											    </table>
											</td>
										</tr>
										<tr>
										    <td class="FieldCellSub">Remarks</td>
										    <td class="FieldCellSub">
										        <textarea name="Remarks" cols="50" rows="1" class="FormElem" ></textarea>
										    </td>
										    <td class="FieldCellSub"></td>
										    <td class="FieldCellSub" colspan="2">
										    <%if trim(sType)="SUB" then%>
										        <a href="#" class="ExcelDisplayLink" onclick="SelectProcess()">Process Detail</a>
										        <span id="spnProcess" class="DataOnly">&nbsp;</span>
										    <%end if  %>
										    </td>
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
										<table cellpadding="0" cellspacing="0" width="100%">
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
																<DIV class=frmBody id=frm2 style="height:215;">
																	<table  id="tblLot" border="0" cellspacing="1" class="ExcelTable" width="100%">
																		<tr>
																			<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2">
																			    <img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" align="center" width="15" height="15" alt="Delete Item(s)" onclick="DeleteItem()">
																			</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description
																			</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2" width="55">Ret.Type
																			</td>
																			<td class="ExcelHeaderCell" align="center" colspan="2">Quantity</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2" colspan="2">Issue</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2"></td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2">Additional<br> Details</td>
																		</tr>
																	    <tr>
																	        <td class="ExcelHeaderCell" align="center" width="100">Req.[To Issue]</td>
																			<td class="ExcelHeaderCell" align="center">In Unit</td>
																	    </tr>
																		<% 'refer(1---1) Row 1 %>
																		<% 'refer(1---1) Row 2 %>

																		<%' End %>

																	</table>
																	<input type="button" name="btnAddItem" class="AddButtonX" value="Add Item" onclick="AddItem()">
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

<%
Function populateConsumptionAccHead()
    Dim sQuery,rsTemp
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select AccountHead,ConsumptionDesc from INV_T_ConsumptionHeadRelation"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        do while not rsTemp.eof
            Response.Write "<option value='"& rsTemp(0) &"'>"& rsTemp(1) &"</option>"
            rsTemp.movenext
        loop
    end if
    rsTemp.close
End Function
%>
