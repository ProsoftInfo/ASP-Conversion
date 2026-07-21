<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MatConRcptSelPop.asp
	'Module Name				:	Inventory (Issues- Consumption)
	'Author Name				:	Ragavendran
	'Created On					:	May 24,2013
	'Modified By				:
	'Modified On				:
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
	dim dcrs,dcrs1,iItem,iClass,sOrgID,iMRSNo,arrTemp,iIssNo
	dim sItmName,iQty,sTemp,sAlt,iLot
	dim sRead,iCtr,dIssDate,sTitle,sVar,iLineNo
	dim arrUoM,sUoMDesc,sUoMCode,sType,iDINo,sSQL,iAccHead,sRecptNum
	Dim sAttributeList,sCallFrom,sIssueCode,iNoOfPacks
	Dim iIntRcptNo,iIntItemCode,sIntAttList,sRcptNum,iOutPutQty,iRcptQty
	Dim sItemCode,sClassCode,sCatCode,sItemName,sClassName,iRowCount
	Dim sFromDate,sToDate,sFinPeriod,sFinArr

	Const iPageSize=10	'How many records to show
    Dim iCurrentPage	'Current Page No.
    Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
    Dim iPageCtr		'Counter
	Dim lnPage

    iCurrentPage = Request.Form("hPageSelection")
    if iCurrentPage = "" or iCurrentPage = "0" then iCurrentPage = "1"
    iCurrentPage = CInt(iCurrentPage)

	iAccHead = 0
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	iCtr = 0
	Response.Write "<font color=red>"
	'Response.Write Request.QueryString("sTemp")

	arrTemp = split(trim(Request.QueryString("sTemp")),":")
	sType   = arrTemp(0)
	iIssNo  = arrTemp(1)
	iItem	= arrTemp(2)
	iClass	= arrTemp(3)
	sOrgID	= arrTemp(4)
	sTitle  = arrTemp(5)
	sVar    = arrTemp(6)
	sAttributeList = arrTemp(7)
	if UBound(arrTemp)>7 then
	    sCallFrom = arrTemp(8)
	    'Response.Write "CallFrom="& sCallFrom
	end if
	if trim(iItem)<>"" then
	    sSQL = "Select ReceiptNumbering from VwItem where ItemCode ="& iItem &" and ClassificationCode = "& iClass
	    'Response.Write sSQL
	    dcrs.open sSQL,con
	    if not dcrs.eof then
	        sRecptNum = trim(dcrs(0))
	    end if
	    dcrs.close
	end if 'if trim(iItem)<>"" then

	sItemCode = Request.QueryString("ItemCode")
	sClassCode = Request.QueryString("ClassCode")
	sCatCode = Request.QueryString("CatCode")
	sItemName = Request.QueryString("ItemName")
	iRowCount = Request.QueryString("hRowCount")
	iCurrentPage = cint(Request.QueryString("hSubmit"))
	sFromDate = Request.QueryString("FromDate")
	sToDate = Request.QueryString("ToDate")
	if Trim(sClassCode)<>"" then
	    sClassName = getClassName(sClassCode)
	end if
	if Trim(sFromDate)="" then
	    sFinPeriod = Session("FinPeriod")
	    sFinArr = Split(sFinPeriod,":")
	    sFromDate = "01/04/"&sFinArr(0)
	    sToDate = "31/03/"&sFinArr(1)
	end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - <%=sTitle%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemData">
<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemXML"><Root></Root></script>
<%
    if trim(iItem)<>"" then
        sItmName = ItemDisplay(iItem,iClass)
	    arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
	    'Response.Write  sOrgID&iClass
	    sUoMCode = arrUoM(0)
	    sUoMDesc = arrUoM(1)
	end if 'if trim(iItem)<>"" then

	sSQL = "Select IsNull(IssueEntryCode,IssueEntryNo),Convert(varchar,IssueDate,103) from Inv_T_MaterialIssueHeader where IssueEntryNo = "& iIssNo
	dcrs.open sSQL,con
	if not dcrs.eof then
	    sIssueCode= dcrs(0)
	    dIssDate = dcrs(1)
	end if
	dcrs.close
%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script language="javascript" src="../../scripts/PrintWindow.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim objTemp,Root,newElem,iQtyTot,RootO
dim iClass,iItem,iIssNo,iMRSNo,sOrgID,sType,iDINo
set objTemp = window.dialogarguments
set RootO = objTemp.documentElement
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
'*********************************
Function GenXML(iRowIndex)
    nCnt = document.formname.hCtr.value
    set ndRoot = ItemData.documentElement
        set objChk = eval("document.formname.ChkZ"&iRowIndex)
         sArrRcptData =split(objChk.value,":") ' RcptNo:RcptItemCode:RcptAttList,RcptNumbering
            RcptNo = sArrRcptData(0)
            RcptItemCode = sArrRcptData(1)
            RcptAttList = sArrRcptData(2)
            RcptNumber = sArrRcptData(3)
        if objChk.checked then
            set ndRcpt = ItemData.CreateElement("RcptItem")
                ndRcpt.setAttribute "No",RcptNo
                ndRcpt.setAttribute "Item",RcptItemCode
                ndRcpt.setAttribute "AttID",RcptAttList
                ndRcpt.setAttribute "RNumbering",RcptNumber
                ndRcpt.setAttribute "Qty","0"
                ndRcpt.setAttribute "ByProduct","N"
            ndRoot.appendChild ndRcpt
        else
            if ndRoot.hasChildNodes() then
                for each ndRcpt in ndRoot.childNodes
                    if ndRcpt.nodeName="RcptItem" then
                        if ndRcpt.getAttribute("No")=trim(RcptNo) and ndRcpt.getAttribute("Item")=trim(RcptItemCode) then
                            ndRoot.removeChild(ndRcpt)
                        end if
                    end if
                next
            end if
        end if
End Function
'**************************************
Function SelectClassifcation()

	OrgID = document.formname.hOrgID.value
	IType = 1
	ReturnData = showModalDialog("/include/ClassificationSelectPop.asp?sIType="&IType&"&sOrgID="&OrgID&"&sITypename="&ITypeName&"&SelMode=M","Classification","dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(ReturnData,"*****")
	if arrTemp(0) = "-1" then exit function


	for i = 0 to ubound(arrTemp)  - 1
		j = 0
		arrTempValue =  split(arrTemp(0),"|")
		for j = 0 to ubound(arrTempValue)
			arrTempClass =  split(arrTempValue(j),":")
			if UBound(ArrTempClass)>0 then
				sClass = sClass & "," & arrTempClass(ubound(arrTempClass))
				sCategory = sCategory & "," & arrTempClass(1)
			else
				'sClass = sClass & "," & mid(arrTempClass(0),4)
				sClass = ""
				sCategory = sCategory & "," & mid(arrTempClass(0),4)
			end if
		next

		arrTempName =  split(arrTemp(1),"|||")
		for z = 0 to ubound(arrTempName)
			arrTempClassName =  split(arrTempName(z),":")
			sClassName = sClassName & "," & arrTempClassName(ubound(arrTempClassName))
		next
	next

	sClass = mid(sClass,2)
	sClassName = mid(sClassName,2)
	sTemp = mid(sTemp,3)
	sCategory = Mid(sCategory,2)

	txtClass.innerText  = sClassName
	document.formname.hClassCode.value = sclass
	document.formname.hCatCode.value = sCategory
End Function
'**************************************
Function SelectItem()


	Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

    sTempValWindowSize = GetWindowSizeForPopup("1")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)

	sUnit = document.formname.hOrgID.value
	sClassCodes = document.formname.hClassCode.value
    iStock = "Y"
	set OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&Stock=" & iStock & "&hSelectMode=M&hDispButt=N&hClassCodes="&sClassCodes,ItemXML,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	if not OutValue.hasChildNodes() then
	    exit function
	end if

	if OutValue.hasChildNodes() then
		for each ndItem in OutValue.childNodes
		    if ndItem.nodeName="Item" then
			    sItemCode = sItemCode &","& ndItem.getAttribute("ItemCode")
		        sItemName = sItemName &","& ndItem.getAttribute ("ItemName")
		    end if 'if ndItem.nodeName="Item" then
		next
	end if
if Trim(sItemCode)<>"" then
	sItemCode = mid(sItemCode,2)
	sItemName = mid(sItemName,2)
end if 'if Trim(sItemCode)<>"" then

	txtItem.innerText  = sItemName
	document.formname.hItemCode.value = sItemCode
End Function
'*************
Function fnInit(obj)
	arrTemp = split(obj,":")
	sType   = arrTemp(0)
	iIssNo  = arrTemp(1)
	iItem	= arrTemp(2)
	iClass	= arrTemp(3)
	sOrgID	= arrTemp(4)

	document.formname.ctlFrom.setDate = document.formname.hFromdate.value
	document.formname.ctlTo.setDate = document.formname.hToDate.value
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement

	set ndRoot = ItemData.documentElement
	if document.formname.hCtr.value = 0 then exit function
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("Class").Value = iClass and HeaderNode.Attributes.getNamedItem("Item").Value  = iItem and HeaderNode.Attributes.getNamedItem("IssEntryNo").Value  = iIssNo then
			if HeaderNode.HaschildNodes() then
			    for each ndRcptItem in HeaderNode.childNodes
			        if ndRcptItem.nodeName="RcptItem" then
			            iRcptNo = ndRcptItem.getAttribute("No")
			            rItemCode = ndRcptItem.getAttribute("Item")
			            rAttList = ndRcptItem.getAttribute("AttID")
			            rNumber = ndRcptItem.getAttribute("RNumbering")
			            rQty = ndRcptItem.getAttribute("Qty")
			            rByProduct = ndRcptItem.getAttribute("ByProduct")
			            for iCnt = 1 to document.formname.hCtr.value
			                set objChk = eval("document.formname.ChkZ"&iCnt)
			                set objByPro = eval("document.formname.ChkBPZ"&iCnt)
			                set objOPQty = eval("document.formname.txtOPQtyZ"&trim(iRcptNo)&"Z"&trim(rItemCode)&"Z"&trim(rAttList)&"Z"&trim(rNumber))
			                sTempVal = trim(iRcptNo)&":"&trim(rItemCode)&":"&trim(rAttList)&":"&trim(rNumber)
			                if objChk.value = sTempVal then
			                    objChk.checked = true
			                end if
			                if trim(rByProduct)="Y" then
			                    objByPro.checked = true
			                end if
			                objOPQty.value = rQty
			            next
			        end if 'if ndRcptItem.nodeName="RcptItem" then
			    next
			end if

		end if
	Next

end Function
'************************************************
Function PackDisplay(IntRcptNo,ItemCode,AttList,Quantity,RcptNumbering,iRowIndex)

    set objChk = eval("document.formname.ChkZ"&iRowIndex)
    if objChk.checked then
        sTempValues = IntRcptNo&":"&ItemCode &":"&AttList&":"&Quantity&":"&RcptNumbering
        set OutDataValue = showModalDialog("MatConOPackSelPop.asp?sTemp="&sTempValues,ItemData,"dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
        set ndRoot = ItemData.documentElement
        if ndRoot.hasChildNodes() then
            for each ndRcptItem in ndRoot.childNodes
                if ndRcptItem.getAttribute("No")=IntRcptNo and ndRcptItem.getAttribute("Item")=ItemCode and ndRcptItem.getAttribute("AttID")=AttList and ndRcptItem.getAttribute("RNumbering")=RcptNumbering then
                    eval("document.formname.txtOPQtyZ"&IntRcptNo&"Z"&ItemCode&"Z"&AttList).value = ndRcptItem.getAttribute("Qty")
                end if
            next
        end if
    else
        alert("Please select the Receipt Number to select the Bag Details")
        eval("document.formname.ChkZ"&iRowIndex).focus
        exit function
    end if
End Function
'*************************************************
Function CheckLot(obj,ictr)
'alert(obj.name)
	sTempValues = obj.name
	set OutDataValue = showModalDialog("MatConSerialPop.asp?sTemp="&sTempValues,ItemData,"dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(obj.name,"`")
	sType  = arrTemp(1)
	sOrgID = arrTemp(2)
	iItem = arrTemp(3)
	iClass = arrTemp(4)
	sLot = arrTemp(5)
	if trim(sType) = "M" then
		iMRSNo = arrTemp(6)
	else
		iDINo = arrTemp(6)
	end if
	iQty = arrTemp(7)
	dIssDate = arrTemp(9)

	set Q = eval("document.formname.txtOPQtyZ"&ictr)
	Set ndRoot = ItemData.documentElement

	if ndRoot.Attributes.getNamedItem("Item").Value  = iItem and ndRoot.Attributes.getNamedItem("Class").Value = iClass and ndRoot.Attributes.getNamedItem("IssEntryNo").Value  = iIssNo then
		Q.value = ndRoot.Attributes.getNamedItem("SerQtyRet").Value
	end if

End Function
'***************************************
Function NextSelection(sPage)
    sRet="NEXT"

	sFromDate = document.formname.ctlFrom.getdate
	sToDate = document.formname.ctlTo.getdate
	sItemCode = document.formname.hItemCode.value
	sClassCode = document.formname.hClassCode.value
	sTemp = document.formname.hobjVal.value
	iCurr=document.formname.hCurrentPage.value
	sRowCount=document.formname.hRowCount.value
	sItemName = txtItem.innerHTML
	sClassName = txtClass.innerHTML

	objIss = document.formname.hObjVal.value
    arrTemp = split(objIss,":")
    sType   = arrTemp(0)
    iIssNo  = arrTemp(1)
    iItem	= arrTemp(2)
    iClass	= arrTemp(3)
    sOrgID	= arrTemp(4)


	sValue = sTemp&"|FromDate="&sFromDate&"|ToDate="&sToDate&"|ItemCode="&sItemCode&"|hCurrentPage="&iCurr&"|hWho="&sRet&"|hSubmit="&sPage&"|hRowCount="&sRowCount&"|ClassCode="&sClassCode&"|ItemName="&sItemName&"|ClassName="&sClassName
	set RootO = objTemp.documentElement
	if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
                for each ndPagination in ndIss.childNodes
                    if Trim(ndPagination.nodeName)="Pagination" then
                        ndIss.removeChild(ndPagination)
                    end if
                next
            end if
        next
    end if

    if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
                set ndPagination = objTemp.createElement("Pagination")
                ndPagination.setAttribute "Details",sValue
                ndIss.appendChild ndPagination
            end if
        next
    end if
    window.close()
End Function
'***********************************
Function CheckSubmit()
	dim ictr,objQ,objSTQ,objSerial

	ictr = document.formname.hCtr.value
	'alert(iCtr)

	if ictr = "0" then exit function

	for i=1 to ictr
		'alert i
		set objChk = eval("document.formname.ChkZ"&i)
		if objChk.checked then
		    sArrTemp = split(objChk.value,":")
		    IntRcptNo = sArrTemp(0)
		    ItemCode = sArrTemp(1)
		    AttList = sArrTemp(2)
		    RcptNum = sArrTemp(3)
		    set objQ = eval("document.formname.txtOPQtyZ"&IntRcptNo&"Z"&Itemcode&"Z"&AttList)
		    set objBPChk = eval("document.formname.ChkBPZ"&i)

		    set ndRoot = ItemData.documentElement
	        if ndRoot.hasChildNodes() then
	            for each ndRcptItem in ndRoot.childNodes
	                if ndRcptItem.getAttribute("No")=IntRcptNo and ndRcptItem.getAttribute("Item")=ItemCode and ndRcptItem.getAttribute("AttID")=AttList and RcptNum = ndRcptItem.getAttribute("RNumbering") then
	                    ndRcptItem.setAttribute "Qty",objQ.value
	                    if objBPChk.checked then
	                        ndRcptItem.setAttribute "ByProduct","Y"
	                    else
	                        ndRcptItem.setAttribute "ByProduct","N"
	                    end if
	                end if
	            next
	        end if
		end if
	next

	Set RootO = objTemp.documentElement
	objIss = document.formname.hObjVal.value
    arrTemp = split(objIss,":")
    sType   = arrTemp(0)
    iIssNo  = arrTemp(1)
    iItem	= arrTemp(2)
    iClass	= arrTemp(3)
    sOrgID	= arrTemp(4)

    if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
                set ndIssNode = ndIss
            end if
        next
    end if

    set ndRoot = ItemData.documentElement
    for i=1 to ictr
		'alert i
		set objChk = eval("document.formname.ChkZ"&i)
		if objChk.checked then
		    sArrTemp = split(objChk.value,":")
		    IntRcptNo = sArrTemp(0)
		    ItemCode = sArrTemp(1)
		    AttList = sArrTemp(2)
		    RcptNum = sArrTemp(3)
		    set objQ = eval("document.formname.txtOPQtyZ"&IntRcptNo&"Z"&Itemcode&"Z"&AttList)
		    set objBPChk = eval("document.formname.ChkBPZ"&i)
		    if ndRoot.hasChildNodes() then
                for each ndRcptItem in ndRoot.childNodes
                    if ndRcptItem.getAttribute("No")=IntRcptNo and ndRcptItem.getAttribute("Item")=ItemCode and ndRcptItem.getAttribute("AttID")=AttList and RcptNum = ndRcptItem.getAttribute("RNumbering") then
                        ndIssNode.appendChild ndRcptItem
                    end if
                next
            end if
        end if
    next

    sRet="DONE```YES"

	sFromDate = document.formname.ctlFrom.getdate
	sToDate = document.formname.ctlTo.getdate
	sItemCode = document.formname.hItemCode.value
	sClassCode = document.formname.hClassCode.value
	sTemp = document.formname.hobjVal.value
	iCurr=document.formname.hCurrentPage.value
	sRowCount=document.formname.hRowCount.value
	sItemName = txtItem.innerHTML
	sClassName = txtClass.innerHTML


	sValue = sTemp&"|FromDate="&sFromDate&"|ToDate="&sToDate&"|ItemCode="&sItemCode&"|hCurrentPage="&iCurr&"|hWho="&sRet&"|hSubmit="&sPage&"|hRowCount="&sRowCount&"|ClassCode="&sClassCode&"|ItemName="&sItemName&"|ClassName="&sClassName

	if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
                for each ndPagination in ndIss.childNodes
                    if Trim(ndPagination.nodeName)="Pagination" then
                        ndIss.removeChild(ndPagination)
                    end if
                next
            end if
        next
    end if

    if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
              set ndPagination = objTemp.createElement("Pagination")
                ndPagination.setAttribute "Details",sValue
                ndIss.appendChild ndPagination
            end if
        next
    end if
	window.close
end Function
'***********************
Function Func_Close()

Set RootO = objTemp.documentElement
	objIss = document.formname.hObjVal.value
    arrTemp = split(objIss,":")
    sType   = arrTemp(0)
    iIssNo  = arrTemp(1)
    iItem	= arrTemp(2)
    iClass	= arrTemp(3)
    sOrgID	= arrTemp(4)

    sRet="DONE```YES"

	sFromDate = document.formname.ctlFrom.getdate
	sToDate = document.formname.ctlTo.getdate
	sItemCode = document.formname.hItemCode.value
	sClassCode = document.formname.hClassCode.value
	sTemp = document.formname.hobjVal.value
	iCurr=document.formname.hCurrentPage.value
	sRowCount=document.formname.hRowCount.value
	sItemName = txtItem.innerHTML
	sClassName = txtClass.innerHTML


	sValue = sTemp&"|FromDate="&sFromDate&"|ToDate="&sToDate&"|ItemCode="&sItemCode&"|hCurrentPage="&iCurr&"|hWho="&sRet&"|hSubmit="&sPage&"|hRowCount="&sRowCount&"|ClassCode="&sClassCode

	if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
                for each ndPagination in ndIss.childNodes
                    if Trim(ndPagination.nodeName)="Pagination" then
                        ndIss.removeChild(ndPagination)
                    end if
                next
            end if
        next
    end if

    if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
                set ndPagination = objTemp.createElement("Pagination")
                ndPagination.setAttribute "Details",sValue
                ndIss.appendChild ndPagination
            end if
        next
    end if

	window.close
End Function
'**************************

Function window_onunload()
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

</script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0  onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">

<form method="POST" name="formname">
<input type=hidden name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hObjVal" value="<%=Request.QueryString("sTemp")%>" />
<input type="hidden" name="hOrgID" value="<%=sOrgID%>" />
<input type="hidden" name="hClassCode" value="<%=sClassCode%>" />
<input type="hidden" name="hCatCode" value="<%=sCatCode%>" />
<input type="hidden" name="hItemCode" value="<%=sItemCode%>" />
<input type="hidden" name="hFromDate" value="<%=sFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sToDate%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Receipt Details
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
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <%if trim(sCallFrom)="ITEM" then %>
                                        <tr>
                                            <td class="FieldCell">Item Description</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idItemName"><%=sItmName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <%end if 'if trim(sCallFrom)<>"" then %>
                                        <tr>
                                            <td class="FieldCell">Issue No - Date</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sIssueCode%>-<%=dIssDate%>&nbsp;</span>
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
                                <td>
                                    <table class="BodyTable" width="100%">
                                        <tr>
                                            <td class="FieldCellSub">Date From</td>
                                            <td class="FieldCellSub">
                                                <% Response.Write InsertDatePicker("ctlFrom")%>
                                            </td>
                                            <td class="FieldCellSub">To</td>
                                            <td class="FieldCellSub">
                                                <% Response.Write InsertDatePicker("ctlTo")%>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Classification</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="txtClass"><%=sClassName%>&nbsp;</span>&nbsp;&nbsp;<img style="cursor: hand" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Classification" onclick="SelectClassifcation()"></td>
                                             <td class="FieldCellSub">Item</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="txtItem"><%=sItemName%>&nbsp;</span>&nbsp;&nbsp;<img style="cursor: hand" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Classification" onclick="SelectItem()">&nbsp;&nbsp;<input type="button" name="btnGo" value=" GO " class="ActionButtonX" onclick="NextSelection('1')" /></td>
                                        </tr>
                                    </table>
                                </td>
                                <td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<DIV class=frmBody id=frm6 style="width: 100%; height:380;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="10"></td>
												<td class="ExcelHeaderCell" align="center" width="80">Receipt No-<br />Date</td>
												<td class="ExcelHeaderCell" align="center">Item</td>
												<td class="ExcelHeaderCell" align="center">Qty[Packs]</td>
												<td class="ExcelHeaderCell" align="center">Output</td>
												<td class="ExcelHeaderCell" align="center">By Product</td>
											</tr>
											<%
											    iCtr = 0
											    sSQL = "Select H.InternalReceiptNo,Convert(varchar,ReceivedOn,103),D.ItemCode,ItemDescription,SUM(QuantityReturn),IsNull(InvRecNo,0),IsNull(D.AttributeList,'') "
											    sSQL = sSQL & "from APP_T_InternalReceiptHeader H Join APP_T_InternalReceiptDetails D on "
											    sSQL = sSQL & "H.InternalReceiptNo = D.InternalReceiptNo Join Inv_M_ItemMaster I on D.ItemCode = I.ItemCode and (InvRecNo is not null or InvRecNo <>0) "
											    sSQL = sSQL & " and ReceivedOn between Convert(datetime,'"& sFromDate&"',103) and Convert(datetime,'"& sToDate &"',103)"

											    if Trim(sClassCode)<>"" then
											        sSQL = sSQL & " and D.ClassificationCode in ("& sClassCode  &")"
											    end if

											    if Trim(sItemCode)<>"" then
											        sSQL = sSQL & " and D.ItemCode in ("& sItemCode  &")"
											    end if
											    sSQL = sSQL & " Group By H.InternalReceiptNo,ReceivedOn,D.ItemCode,ItemDescription,InvRecNo,D.AttributeList"
											    'Response.write "<textarea>"& sSql &"</textarea>"
											    With dcrs
											        .activeConnection = con
											        .CursorLocation = 3
											        .CursorType = 3
											        .Source = sSQL
											        .open
											    End With
											    if not dcrs.eof then
											        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   														dcrs.PageSize = iPageSize
														If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
														dcrs.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
														iTotPage = dcrs.PageCount					'stores total no. of pages
													'''''''''''''''''''''''''''''''''''''''''''''''''''''''
														For iPageCtr = 1 to dcrs.PageSize

											                iNoOfPacks = 0
											                iIntRcptNo = dcrs(0)
											                iIntItemCode = dcrs(2)
											                sIntAttList = dcrs(6)
											                iRcptQty = dcrs(4)
											                sSQL = "Select Count(*) from INV_T_LocationLot where InventoryReceiptNo = "& dcrs(5) &" and ItemCode ="& dcrs(2)
											                sSql = sSql & " and SerialNumber not in (Select Serialno from INV_T_MaterialConsumptionOutput H join INV_T_MaterialConsumptionOutputDet D on H.ConsumptionNo = D.ConsumptionNo  and H.IssueEntryNo = D.IssueEntryNo and H.LineNumber = D.LineNumber)"
											                'Response.write "<textarea>"& sSql &"</textarea>"
											                dcrs1.open sSQL,con
											                if not dcrs1.eof then
											                    iNoOfPacks = dcrs1(0)
											                end if
											                dcrs1.close

											                sSql = "Select IsNull(Sum(OutputQuantity),0) from INV_T_MaterialConsumptionOutput where AppRefNo = "& iIntRcptNo &" and AppRefType = 39"
											                dcrs1.open sSql,con
											                if not dcrs1.eof then
											                    iOutPutQty = dcrs1(0)
											                end if
											                dcrs1.close

											               sRcptNum = GetItemRcptNum(iIntItemCode)

								                            if iRcptQty<>"" then
								                                if cdbl(iRcptQty)>cdbl(iOutPutQty) then
								                                    iRcptQty =  cdbl(iRcptQty)-cdbl(iOutputQty)
								                                else
								                                    iRcptQty = 0
								                                end if
								                            end if


											                if iRcptQty>0 then
											                iCtr = iCtr + 1
											                iRowCount = iRowCount + 1
											                %>
											                    <tr>
												                    <td class="ExcelSerial" align="center"><%=iCtr%></td>
												                    <td class="ExcelDisplayCell" align="center" width="10">
												                        <input type="checkbox" name="ChkZ<%=iCtr%>" value="<%=iIntRcptNo%>:<%=iIntItemCode%>:<%=sIntAttList%>:<%=sRcptNum%>" onclick="GenXML('<%=iCtr%>')"/>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center" width="80"><%=dcrs(0)%>-<%=dcrs(1)%></td>
												                    <td class="ExcelDisplayCell" align="left"><%=dcrs(3)%>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center"><%=iRcptQty%>
												                    <%if iNoOfPacks>0 then
												                        Response.write "["&iNoofPacks&"]"
												                      end if'if iNoOfPacks>0 then %>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center">
												                        <%if trim(sRcptNum)="N" then%>
												                        <input type="text" name="txtOPQtyZ<%=iIntRcptNo%>Z<%=iIntItemCode%>Z<%=sIntAttList%>" value="0" class="FormElem" style="text-align:right" size="12" >
												                        <%else%>
												                        <input type="text" name="txtOPQtyZ<%=iIntRcptNo%>Z<%=iIntItemCode%>Z<%=sIntAttList%>" value="0" class="FormElem" style="text-align:right" size="12" disabled="true">
												                        <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" id="imgZ<%=iIntRcptNo%>Z<%=iIntItemCode%>Z<%=sIntAttList%>" onClick="PackDisplay('<%=iIntRcptNo%>','<%=iIntItemCode%>','<%=sIntAttList%>','<%=dcrs(4)%>','<%=sRcptNum%>','<%=iCtr%>')" >
												                        <%end if 'if trim(sRcptNum)="N" then%>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center">
												                        <input type="checkbox" name="ChkBPZ<%=iCtr%>" value="0" class="FormElem" style="text-align:right">
												                    </td>
											                    </tr>
											                <%
											                end if 'if iRcptQty>0 then
											            dcrs.movenext
											            If dcrs.EOF Then Exit For
											        Next
											    end if
											    dcrs.close
											%>
										</table>
									</div>
								</td>
								<input type="hidden" name="hCtr" value="<%=iCtr%>">
								<input type="hidden" name="hRowCount" value="<%=iRowCount%>" />
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
									    <td colspan="2" align="right">
								                        <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
								                        <Input Type=Hidden name="hPageSelection" Value="1" >
														<%	If iTotPage >= 2 Then
																if iCurrentPage = 1 then
														%>
														<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
														<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
														<%		else	%>
														<input type="button" value=" |< " class="ActionButtonX" onclick="NextSelection('1')" id=button3 name=button3>
														<input type="button" value=" << " class="ActionButtonX" onclick="NextSelection('<%=iCurrentPage - 1%>')" id=button4 name=button4>
    													<%		end if	%>
    													<SELECT class="FormElem" onChange="NextSelection(this(this.selectedIndex).value)" id="selPage">
    													<%
															For lnPage = 1 To iTotPage
																If lnPage = iCurrentPage Then
														%>
															<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotPage%></OPTION>
														<%		else	%>
															<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
    													<%		end if
    														next
    													%>
    													</SELECT>
    													<%
    															if iCurrentPage = iTotPage then
    													%>
														<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
														<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

    													<%		else	%>
														<input type="button" value=" >> " class="ActionButtonX" onclick="NextSelection('<%=iCurrentPage + 1%>')" id=button7 name=button7>
														<input type="button" value=" >| " class="ActionButtonX" onclick="NextSelection('<%=iTotPage%>')" id=button8 name=button8>
    													<%		end if
															End If
														%>
												</td>
												<td></td>

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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Close" name="B2" class="ActionButton" onclick="Func_Close()">
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
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
<%
	' Function to populate Store
	Function DisplayUoM(sOrgID,iClass,iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.source
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
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sOrgID) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sOrgID) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				stypID = dcrs(0)
				stypName = dcrs(1)
				if cint(iAccHead) = cint(stypID) then
					Response.Write("<OPTION VALUE="""&trim(stypID)&""" SELECTED>"&trim(stypName)&"</OPTION>" &vbcrlf)
				else
					Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				end if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>
