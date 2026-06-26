<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURDetailsEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 01 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include File="../../include/CheckACCPrevFinYear.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo
dim sReferenceNo,sInvoiceNo,iBkAccHead,sPartyName,sSetInvDate,sBkAccDesc
Dim sCode,sValue
Dim sFinPeriod,sValTemp2,sFinFromDate,sFinTodate
sOrgId = Session("organizationcode")
sOrgName = Session("OrgShortName")
iBookNo=Request.Form("selBook")
sReferenceNo=Request.Form("txtReferenceNo")
sInvoiceNo=Request.Form("txtInvoiceNo")
sPartyName=Request.Form("txtPartyName")
sSetInvDate = Request.Form("hInvDate")

Set objRs = Server.CreateObject("ADODB.RecordSet")


sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFromDate = "01/04/"& sValTemp2(0)
sFinToDate = "31/03/"&sValTemp2(1)



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData">
<Details BasicValue="" Discount="" ActualValue="" VouDate=""/></XML>
<XML id="EntryData"><Entry No="0" PayTo="" Amount="" Qty="" UOM="" UOMValue="" Rate="" ActValue="" DisPer="" DisAmount="" ItemCode="" ClassCode="" /></XML>
<XML id="AccHeadData">
<account/>
</XML>
<xml id="UnitBookData"><Root></Root></xml>
<xml id="OutData"><Root></Root></xml>
<xml id="PartyData"><Root></Root></xml>
<xml id="ItemData"><Root></Root></xml>
<xml id="VoucherData"><Root></Root></xml>
<xml id="GLHeadData"><Root></Root></xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../../scripts/checkdate.js"></script>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag,DelRoot
dim iBookAccCode
iEntryNo=1
bVouFlag=false
bSavFlag=false
set VouRoot=DetData.documentElement
set EntryRoot=EntryData.documentElement
set DelRoot=EntryData.documentElement

'iBookAccCode=<%=iBkAccHead%>
FUNCTION CheckAccHead(nodRoot,sAccHead)
dim sExp
sExp="//AccHead[@No='"&sAccHead&"']"
set tempNode=nodRoot.selectNodes(sExp)

if tempNode.length > 0 then
	CheckAccHead=true
else
	CheckAccHead=false
end if
END FUNCTION

function popSalesHead(objAcc)
dim sOrgId,sTemp,sDesc
	sOrgId=document.formname.hOrgId.value
	if objAcc.selectedIndex >0 then
				if objAcc.value="G" then
					showGLHead sOrgId
				else
					sTemp=Split(objAcc.value,"?")
					sDesc=objAcc.options(objAcc.selectedIndex).text
					window.spAccHead.innerHTML=sDesc
					document.formname.txtDescription.value=sDesc
					Set newElem = EntryData.createElement("AccHead")
						newElem.setAttribute "No", trim(sTemp(0))
						newElem.setAttribute "CostCenter", trim(sTemp(1))
						newElem.setAttribute "Analytical", trim(sTemp(2))
						newElem.setAttribute "Name", sDesc
						newElem.setAttribute "Type", "G"
						newElem.setAttribute "Group", ""
	    			EntryRoot.appendChild newElem

					showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))
				End if 'End of select Account Head Type check GL or PARTY
	End if 'End of If any Account Head Selected Check
End function
'---------------------End Of Function selAccountHead-------------------
function showCCAnal(sOrgId,iAccCode,bCostCenter,bAnal)
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo

if cint(bCostCenter)=1 or cint(bAnal)=1 then
'If Selected GL Account Head has Cost Center
	Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+iAccCode,"","")
	if nodCCAnly.Attributes.Item(0).nodeValue=1 then
		'Set the Additional and CCANAL Display Layer Visible
		setADDDisplay 1
		For Each HeaderNode In nodCCAnly.childNodes
			if 	HeaderNode.nodeName="CostCenter" then
				EntryRoot.appendChild HeaderNode
				if HeaderNode.hasChildNodes then
					'If user has Selected Cost centers
					iSno=1
					setAnalDisplay "C",1
					ClearTable "tblCost",1,1
					for each  nodCC in HeaderNode.childNodes
						sCode=trim(nodCC.Attributes.Item(0).nodeValue)
						sDesc=nodCC.Attributes.Item(2).nodeValue
						dRatio=nodCC.Attributes.Item(3).nodeValue


						set oRow = document.all.tblCost.insertRow(iSno)
						InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,2,"txtCCRatio"&sCode,CStr(dRatio),"ExcelInputCell","","",4,3,0,0,""
						InsertCell oRow,2,"txtCCAmount"&sCode,"0","ExcelInputCell","","",12,10,0,0,""

						iSno=iSno+1
					next
				else
					'No Cost Center Selected
					setAnalDisplay "C",0
				end if 'End of Check for Selected Cost centers
			end if 'End of Check for Cost Center Node

			if 	HeaderNode.nodeName="Analytical" then

				EntryRoot.appendChild HeaderNode
				if HeaderNode.hasChildNodes then
					iSno=1
					setAnalDisplay "A",1
					ClearTable "tblAnal",1,1
					for each  nodANL in HeaderNode.childNodes
						sCode=trim(nodANL.Attributes.Item(0).nodeValue)
						sDesc=nodANL.Attributes.Item(2).nodeValue
						dRatio=nodANL.Attributes.Item(3).nodeValue
						sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value

						set oRow = document.all.tblAnal.insertRow(iSno)
						InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","","",0,0,0,0,""
						InsertCell oRow,2,"txtANALRatio"&sCode&"Z"&sGroupCode,dRatio,"ExcelInputCell","","",4,3,0,0,""
						InsertCell oRow,2,"txtANALAmount"&sCode&"Z"&sGroupCode,"0","ExcelInputCell","","",12,10,0,0,""
						iSno=iSno+1
					next
				else
					'No Analytical Selected
					setAnalDisplay "A",0
				end if 'End of Check for Selected Analytical
			end if 'End of Check for Analytical Node
		next	'End of Processing CCANAL Node
	else
		'User Has canceled CC,ANAL Selection
		'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
 		setADDDisplay 0
	end if	'End of CC,ANAL has Childs Check
else
	'Selected Head has no CC or ANAL
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if	'End of GL has Cost Center or not

set nodAccHead=nothing
set nodCCAnly=nothing
set nodCC=nothing

End function
'---------------------End Of Function showCCAnal--------------------------
function showGLHead(sOrgId)
dim iAccCode,bAnal,bCostCenter
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,arrTemp,sRetVal
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
iBookNo=document.formname.hBookcode.value

		sTempValWindowSize = GetWindowSizeForPopup("5")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)

		Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&BookId=01&BookNo="&iBookNo&"&AccHead="+cstr(iBookAcchead),GLHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,GLHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

'	set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'	    sQuery = OutValue.getAttribute("PassQuery")
'	    if OutValue.getAttribute("Action")="CLOSE" then exit function
'	while OutValue.getAttribute("Action")<>"Done"
'		set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'		    sQuery = OutValue.getAttribute("PassQuery")
'	        if OutValue.getAttribute("Action")="CLOSE" then exit function
'	wend
'alert(OutValue.xml)
if OutValue.hasChildNodes() then
    For each ndChild in OutValue.childNodes
        sRetVal = ndChild.getAttribute("RetField0")
        sRetVal = sRetVal&":"&ndChild.getAttribute("RetField1")
        sRetVal = sRetVal&":"&ndChild.getAttribute("RetField2")
        sRetVal = sRetVal&":"&ndChild.getAttribute("RetField3")
        sRetVal = sRetVal&":"&ndChild.getAttribute("RetField4")
        sRetVal = sRetVal&":"&ndChild.getAttribute("RetField5")
    next
end if
GetGlHeadXml(sRetVal)

Set nodAccHead = AccHeadData.documentElement

if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		iAccCode=HeaderNode.Attributes.Item(0).nodeValue
		bAnal=HeaderNode.Attributes.Item(1).nodeValue
		bCostCenter=HeaderNode.Attributes.Item(2).nodeValue
		window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue
		'document.formname.txtDescription.value=HeaderNode.Attributes.Item(3).nodeValue
		document.formname.txtDescription.value=""
		EntryRoot.appendChild HeaderNode
	next
	showCCAnal sOrgId,iAccCode,bCostCenter,bAnal
else
	'User canceled Account Head Selection
	document.formname.selAccountHead.selectedIndex=0
	document.formname.txtDescription.value=""
	window.spAccHead.innerHTML=""

	if EntryRoot.hasChildNodes then
		set oldChild = EntryRoot.removeChild(EntryRoot.childNodes.item(0))
	end if

	setADDDisplay 0
end if 'End of GL Head Processing
set nodAccHead=nothing
End function
'---------------------End Of Function showGLHead--------------------------
'**************************************
FUNCTION popPartyType()
	dim iHeadCount

	sOrgId=document.formname.hOrgID.value
	iBkNo=document.formname.selBook.value

	for iCounter=1 to document.formname.selPartyType.length
		document.formname.selPartyType.remove(1)
	next

'	set objhttp = CreateObject("MSXML2.XMLHTTP")
'	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo&"&sCallTy=P" , false
'	objhttp.send
'
'	if objhttp.responseXML.xml <> "" then
'		OutData.loadXML objhttp.responseXML.xml
'		Set Root = OutData.documentElement
'		iCounter=1
'		For Each HeaderNode In Root.childNodes
'			set oText1 = document.createElement("<Option>" )
'				oText1.Text = HeaderNode.text
'				oText1.Value = HeaderNode.Attributes.Item(0).nodeValue
'
'			document.formname.selPartyType.add oText1,iCounter
'			iCounter=CDbl(iCounter)+1
'		next
'	end if

    set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "GET","../../Common/PartySubType.asp?ParCode="&sParCode&"&OrgCode="&sOrgId,false
		objhttp.send
		'alert(objhttp.responseText)
		if trim(objhttp.responseXML.xml)<>"" then
			OutData.loadXML(objhttp.responseXML.xml)
			set sTempRoot = OutData.documentElement
			document.formname.selPartyType.length = 1
			for each subNode in sTempRoot.childNodes
				document.formname.selPartyType.length= document.formname.selPartyType.length + 1
				document.formname.selPartyType(document.formname.selPartyType.length-1).value = subNode.getAttribute("SubType")
				document.formname.selPartyType(document.formname.selPartyType.length-1).text = subNode.text
			next
		end if


END FUNCTION
'******************************
Function GetAccHead(obj)
    Dim iAccHead,sAccHdName
'    alert(obj.value)
    if Trim(obj.value)<>"S" then
        document.formname.hBookcode.value = obj.value
        set objhttp = createobject("Microsoft.xmlhttp")
        objhttp.open "GET","GetAccHeadName.asp?Mod=PUR&BookNo="&obj.value,false
        objhttp.send
    '    alert(objhttp.responseText)
        if trim(objhttp.responseText)<>"" then
            sTemp = split(objhttp.responseText,":")
            if uBound(sTemp)=0 then
                alert(stemp)
            else
                iAccHead = stemp(0)
                iAccHdName = stemp(1)
                document.formname.hSalAccCode.value = iAccHead
                document.formname.hSalAccName.value = iAccHdName
            end if
        end if
    end if 'if Trim(obj.value)<>"S" then
    if document.formname.hSalAccCode.value="0" then
	    for iCnt = 0 to cint(document.formname.selAccountHead.length)-1
	        if document.formname.selAccountHead(iCnt).value = "S" then
	            document.formname.selAccountHead.selectedIndex = iCnt
	            spAccHead.innerHTML =  document.formname.hSalAccName.value
	            document.formname.selAccountHead.disabled = true
	        end if
	    next
	else
	    for iCnt = 0 to cint(document.formname.selAccountHead.length)-1
	        if document.formname.selAccountHead(iCnt).value = "G" then
	            document.formname.selAccountHead.selectedIndex = iCnt
	            spAccHead.innerHTML =  document.formname.hSalAccName.value
	            document.formname.selAccountHead.disabled = true
	        end if
	    next
	end if

End Function
'*************************************
'************************************
Function AddEntry(bFlag)
dim iCode,dRatio,dAmount
dim HeaderNode,nodANL,sChExp

IF CStr(document.formname.hSalAccCode.Value) = "0" and document.formname.selAccountHead.selectedIndex = 0 Then
	MsgBox "Select Purchase Account Head "
	document.formname.focus()
	Exit Function
End IF

sChExp = "//AccHead"
Set AccNode = EntryRoot.selectNodes(sChExp)
IF AccNode.length = 0 and CDbl(document.formname.txtAmount.value) <> 0 Then
	GetGlHeadXmlForSalAcc()
	Set nodAccHead = AccHeadData.documentElement
	For Each HeaderNode in nodAccHead.childNodes
		EntryRoot.appendChild HeaderNode
	Next
End IF


if (iEntryNo=1 and bFlag="S") or bFlag="A" or bFlag="U"  then
	if not EntryRoot.hasChildNodes then
		Msgbox("Select a Account Head")
		document.formname.selAccountHead.focus
		exit Function
	end if

	IF document.formname.txtDescription.value = "" Then
		MsgBox "Select Item Description "
		Exit Function
	End IF
end if

if bFlag<>"U" then
	EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
Else
	EntryRoot.Attributes.Item(0).nodeValue=document.formname.hEditEntNo.value
end if


if EntryRoot.hasChildNodes then
	if not checkFileds then exit function
	'EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	EntryRoot.Attributes.Item(1).nodeValue = document.formname.txtDescription.value
	EntryRoot.Attributes.Item(2).nodeValue = FormatNumber(document.formname.txtAmount.value,2,,,0)
	EntryRoot.Attributes.Item(3).nodeValue = document.formname.txtQty.value
	EntryRoot.Attributes.Item(4).nodeValue = document.formname.selUOM.value
	EntryRoot.Attributes.Item(5).nodeValue = document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
	EntryRoot.Attributes.Item(6).nodeValue = FormatNumber(document.formname.txtRate.value,2,,,0)
	EntryRoot.Attributes.Item(7).nodeValue = FormatNumber(document.formname.txtValue.value,2,,,0)
	EntryRoot.Attributes.Item(8).nodeValue = FormatNumber(document.formname.txtDisPercentage.value,2,,,0)
	EntryRoot.Attributes.Item(9).nodeValue = FormatNumber(document.formname.txtDisAmount.value,2,,,0)
	EntryRoot.Attributes.Item(10).nodeValue = document.formname.hItemCode.value
	EntryRoot.Attributes.Item(11).nodeValue = document.formname.hClassCode.value

	VouRoot.Attributes.Item(3).nodeValue=document.formname.ctlDate.GetDate


	for each HeaderNode in EntryRoot.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				iCode=trim(nodANL.Attributes.Item(0).nodeValue)


				dRatio=eval("document.formname.txtCCRatio"&iCode).value
				dAmount=eval("document.formname.txtCCAmount"&iCode).value

				nodANL.Attributes.Item(3).nodeValue=dRatio
				nodANL.Attributes.Item(4).nodeValue=dAmount
			next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value
				dRatio=eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value
				dAmount=eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value

				nodANL.Attributes.Item(3).nodeValue=dRatio
				nodANL.Attributes.Item(4).nodeValue=dAmount
			next
		end if 'End of Check for Analytical Node
	next

'====== This is to Insert/append the the entry in same order as on the creation ==
	IF CStr(bFlag) = "U" Then
		Dim iCurrEntNo,insNode,sInsxp
		iCurrEntNo = EntryRoot.Attributes.Item(0).nodeValue
		sInsxp = "//Entry[@No="&iCurrEntNo+1&"]"
		Set insNode = VouRoot.selectNodes(sInsxp)
		IF insNode.length <> 0 Then
			VouRoot.insertBefore EntryRoot,insNode.Item(0)
		Else
			VouRoot.appendChild EntryRoot
		End IF
	Else
		VouRoot.appendChild EntryRoot
	End IF
'====================================================================================

end if
	if bFlag="S" then
	    IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			Exit Function
		End IF
	else
		DisplayVoucher
		iEntryNo=iEntryNo+1
		clearXML()
		setADDDisplay 0
		document.formname.txtDescription.value = ""
		document.formname.txtAmount.value = "0.00"
		document.formname.txtRate.value = "0.00"
		document.formname.txtValue.value  = "0.00"
		document.formname.txtDisAmount.value = "0.00"
		document.formname.txtDisPercentage.value = "0"
		document.formname.txtQty.value = "0"
		'document.formname.reset

		IF CStr(document.formname.hSalAccCode.Value) = "0" Then
			document.formname.selAccountHead.selectedIndex=0
			window.spAccHead.innerHTML=""
		Else
			document.formname.selAccountHead.selectedIndex=1
			window.spAccHead.innerHTML=document.formname.hSalAccName.Value
		End IF

		document.formname.btnAdd.disabled = False
		document.formname.btnDel.disabled = True
		document.formname.btnNext.disabled = False
		document.formname.btnUpdate.disabled = True
	end if
end Function
'---------------------End Of Function AddEntry--------------------------
Function SaveXML()
    if validate then
        VouCreate
        'alert(DetData.xml)
        set objhttp = CreateObject("Microsoft.XMLHTTP")
	    objhttp.Open "POST","XMLSave.asp?Mod=PUR&Name=Voucher Entry", false
	    objhttp.send VoucherData.XMLDocument

	    set objhttp = CreateObject("Microsoft.XMLHTTP")
	    objhttp.Open "POST","XMLUpdate.asp?Mod=PUR&Name=Voucher Entry", false
	    objhttp.send DetData.XMLDocument
	    if objhttp.responseText <> "" then
		    Msgbox(objhttp.responseText)
	    else
		    document.formname.submit()
	    end if
	end if 'if validate then
End Function
'---------------------End Of Function SaveXML--------------------------
Function DisplayVoucher()
dim iSno,sAmount,sRate,sQty,sValue,sDiscount,iRow
dim dTotal,sDescription,sUom

window.DisVoucher.style.height="200px"
window.DisVoucher.style.visibility="visible"
ClearTable "tblVoucher",1,1
dTotal=0
iRow = 1

For Each EntryNode in VouRoot.childNodes

	iSno=EntryNode.Attributes.Item(0).nodeValue
	sDescription=EntryNode.Attributes.Item(1).nodeValue
	sAmount=EntryNode.Attributes.Item(2).nodeValue
	sRate=EntryNode.Attributes.Item(6).nodeValue
	sQty=EntryNode.Attributes.Item(3).nodeValue &"&nbsp;"&EntryNode.Attributes.Item(5).nodeValue
	sUom = EntryNode.Attributes.Item(5).nodeValue
	sValue=EntryNode.Attributes.Item(7).nodeValue
	sDiscount=EntryNode.Attributes.Item(9).nodeValue
	dTotal=FormatNumber(CDbl(dTotal)+CDbl(sAmount),2,,,0)

	set oRow = document.all.tblVoucher.insertRow()
	InsertCell oRow,1,"",iRow,"ExcelSerial","Center","top",0,0,0,0,""
	InsertCell oRow,1,"","<a class=""ExcelDisplaylink"" href=""javascript:EditEntry('"&iSno&"')"">Edit</a>","ExcelDisplayCell","Center","top",0,0,0,0,""
	InsertCell oRow,1,"",sDescription,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sRate,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sQty,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sValue,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sDiscount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""

	iRow = iRow + 1

next'End of Voucher Node Loop
	set oRow = document.all.tblVoucher.insertRow()

	InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,7,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","left","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",dTotal,"ExcelDisplayCell","right","top",0,0,0,0,""
End Function
'---------------------End Of Function DisplayVoucher----------------------
Function  checkFileds()


	if Trim(document.formname.txtDescription.value) = "" and Cstr(document.formname.hSalAccCode.Value) = "0" Then
		MsgBox "Select Item Description "
		checkFileds=false
		exit Function
	elseif  trim(document.formname.txtQty.value)="" then
		Msgbox("Enter Quantity")
		document.formname.txtQty.select
		checkFileds=false
		exit Function
	elseif IsNumeric(document.formname.txtQty.value)=false then
		Msgbox("Enter Numeric values for Quantity")
		document.formname.txtQty.select
		checkFileds=false
		exit Function
	end if

	if  trim(document.formname.txtRate.value)="" then
		Msgbox("Enter Rate")
		document.formname.txtRate.select
		checkFileds=false
		exit Function
	elseif IsNumeric(document.formname.txtRate.value)=false then
		Msgbox("Enter Numeric values for Rate")
		document.formname.txtRate.select
		checkFileds=false
		exit Function
	end if
	if  trim(document.formname.txtDisAmount.value)="" then
		Msgbox("Enter Discount")
		document.formname.txtDisAmount.select
		checkFileds=false
		exit Function
	elseif IsNumeric(document.formname.txtDisAmount.value)=false then
		Msgbox("Enter Numeric values for Discount")
		document.formname.txtDisAmount.select
		checkFileds=false
		exit Function
	end if
	checkFileds=true
end Function
'---------------------End Of Function checkFileds--------------------------
Function calculateField(bFlag)
	if  trim(document.formname.txtQty.value)="" then
		Msgbox("Enter Quantity")
		document.formname.txtQty.select
		calculateField=false
		exit Function
	elseif IsNumeric(document.formname.txtQty.value)=false then
		Msgbox("Enter Numeric values for Quantity")
		document.formname.txtQty.select
		calculateField=false
		exit Function
	end if
	if  trim(document.formname.txtRate.value)="" then
		Msgbox("Enter Rate")
		document.formname.txtRate.select
		calculateField=false
		exit Function
	elseif IsNumeric(document.formname.txtRate.value)=false then
		Msgbox("Enter Numeric values for Rate")
		document.formname.txtRate.select
		calculateField=false
		exit Function
	end if

	IF Trim(document.formname.txtRate.value) = 0 Then
		Exit Function
	End IF


	select case bFlag
		case 1
				IF CDbl(document.formname.txtRate.value) <> 0 and  CDbl(document.formname.txtQty.value) <> 0 Then
					document.formname.txtValue.value= FormatNumber(CDbl(document.formname.txtRate.value)* CDbl(document.formname.txtQty.value),2,,,0)
				Else
					document.formname.txtValue.value = 0
				End IF
				if CDbl(document.formname.txtDisPercentage.value)>0 then
					document.formname.txtDisAmount.value= FormatNumber(CDbl(document.formname.txtValue.value)* (CDbl(document.formname.txtDisPercentage.value)/100),2,,,0)
				Else
					document.formname.txtDisAmount.value= 0
				end if

				document.formname.txtAmount.value= FormatNumber(CDbl(document.formname.txtValue.value)- CDbl(document.formname.txtDisAmount.value),2,,,0)
		case 2

				IF CDbl(document.formname.txtRate.value) <> 0 and  CDbl(document.formname.txtQty.value) <> 0 Then
					document.formname.txtValue.value= FormatNumber(CDbl(document.formname.txtRate.value)* CDbl(document.formname.txtQty.value),2,,,0)
				End IF
				if IsNumeric(document.formname.txtDisPercentage.value)=false then
					Msgbox("Enter Numeric values for Discount Percentage")
					document.formname.txtDisPercentage.select
					calculateField=false
					exit Function
				ELSEif CDbl(document.formname.txtDisPercentage.value) >100 then
					MsgBox "DisCount Percentage Should be less than 100"
					document.formname.txtDisPercentage.select
					calculateField=false
					exit function
				end if

				document.formname.txtDisAmount.value= FormatNumber(CDbl(document.formname.txtValue.value)* (CDbl(document.formname.txtDisPercentage.value)/100),2,,,0)
				document.formname.txtAmount.value= FormatNumber(CDbl(document.formname.txtValue.value)- CDbl(document.formname.txtDisAmount.value),2,,,0)
		case 3
				IF CDbl(document.formname.txtRate.value) <> 0 and  CDbl(document.formname.txtQty.value) <> 0 Then
					document.formname.txtValue.value= FormatNumber(CDbl(document.formname.txtRate.value)* CDbl(document.formname.txtQty.value),2,,,0)
				End IF

				if IsNumeric(document.formname.txtDisAmount.value)=false then
					Msgbox("Enter Numeric values for Discount Amount")
					document.formname.txtDisAmount.select
					calculateField=false
					exit Function
				elseif CDbl(document.formname.txtDisAmount.value) >CDbl(document.formname.txtValue.value) then
					MsgBox "DisCount Value Should be less than actual Value"
					document.formname.txtDisAmount.select
					calculateField=false
					exit function
				end if
				document.formname.txtDisPercentage.value=FormatNumber( ((CDbl(document.formname.txtDisAmount.value)/ CDbl(document.formname.txtValue.value))* 100),2,,,0)
				document.formname.txtAmount.value= FormatNumber(CDbl(document.formname.txtValue.value)- CDbl(document.formname.txtDisAmount.value),2,,,0)

	end select
	popAddAmount1()
	calculateField=true
end Function
'---------------------End Of Function calculateField----------------------------
Function clearXML()
	Set EntryRoot = EntryData.createElement("Entry")
		EntryRoot.setAttribute "No",iEntryNo
		EntryRoot.setAttribute "PayTo",""
		EntryRoot.setAttribute "Amount",""
		EntryRoot.setAttribute "Qty",""
		EntryRoot.setAttribute "UOM",""
		EntryRoot.setAttribute "UOMValue",""
		EntryRoot.setAttribute "Rate",""
		EntryRoot.setAttribute "ActValue",""
		EntryRoot.setAttribute "DisPer",""
		EntryRoot.setAttribute "DisAmount",""
		EntryRoot.setAttribute "ItemCode",""
		EntryRoot.setAttribute "ClassCode",""

end Function
'---------------------End Of Function clearXML----------------------------
Function ClearTable(objTable,startlen,Count)
	dim i

	for i=startlen to eval("document.all."&objTable).rows.length - Count
		eval("document.all."&objTable).deleteRow(startlen)
	next
end Function
Function setAnalDisplay(sDisplay,iFlag)
if sDisplay="A" then
	if iFlag=0 then
		window.DisAnal.style.height="1px"
		window.DisAnal.style.width ="1px"
		window.DisAnal.style.visibility="hidden"
	else
		window.DisAnal.style.height="100px"
		window.DisAnal.style.width ="280px"
		window.DisAnal.style.visibility="visible"
	end if
else
	if iFlag=0 then
		window.DisCost.style.height="1px"
		window.DisCost.style.width ="1px"
		window.DisCost.style.visibility="hidden"
	else
		window.DisCost.style.height="100px"
		window.DisCost.style.width ="280px"
		window.DisCost.style.visibility="visible"
	end if
end if
end Function
'---------------------End Of Function setAnalDisplay----------------------------
Function setADDDisplay(iFlag)
	if iFlag=0 then
		window.Disaddtional.style.height="1px"
		window.Disaddtional.style.visibility="hidden"
		window.DisCCANL.style.height="1px"
		window.DisCCANL.style.visibility="hidden"
	else
		window.Disaddtional.style.height="115px"
		window.Disaddtional.style.visibility="visible"
		window.DisCCANL.style.height="114px"
		window.DisCCANL.style.visibility="visible"
	end if

end Function
'---------------------End Of Function setAnalDisplay----------------------------
Function CancelAction(sPage)
	document.formname.action=sPage
	document.formname.submit
end Function
'---------------------End Of Function ActionCancel----------------------------
'===============================================================================================
FUNCTION popAddAmount1()

'dim dAmount,iChildCount,dRatio,dTotal,dRatioTotal,iCounter


if not checkFileds then
	document.formname.txtAmount.value=""
	exit function
end if

for each HeaderNode in EntryRoot.childNodes


	dim sGroupCode
	if HeaderNode.nodeName="CostCenter" then
		dAmount=CDbl(document.formname.txtAmount.value)
		dTotal=dAmount
		dRatioTotal=0
		iCounter=1
		iChildCount=HeaderNode.childNodes.length
		if Cint(iChildCount)> 0 then
			dRatio=Round(100 / iChildCount,2)
			dAmount= Round((dRatio*dAmount)/100,2)
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value

				if iCounter<iChildCount then
					eval("document.formname.txtCCRatio"&iCode).value=dRatio
					eval("document.formname.txtCCAmount"&iCode).value=dAmount
					nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
					nodANL.Attributes.getNamedItem("Amount").Value=dAmount
					dTotal=CDbl(dTotal)-dAmount
					dRatioTotal=CDbl(dRatioTotal)+dRatio
				else
					eval("document.formname.txtCCRatio"&iCode).value=100-dRatioTotal
					eval("document.formname.txtCCAmount"&iCode).value=dTotal
					nodANL.Attributes.getNamedItem("Ratio").Value=100-dRatioTotal
					nodANL.Attributes.getNamedItem("Amount").Value=dTotal
				end if
				iCounter=CInt(iCounter)+1

			next
		end if 'End of Check for Cost Center Child Count
	end if 'End of Check for Cost Center Node

	if HeaderNode.nodeName="Analytical" then

		dAmount=CDbl(document.formname.txtAmount.value)
		dTotal=dAmount
		dRatioTotal=0
		iCounter=1
		iChildCount=HeaderNode.childNodes.length
		if Cint(iChildCount)> 0 then
			dRatio=Round(100 / iChildCount,2)
			dAmount= Round((dRatio*dAmount)/100,2)
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value
				if iCounter<iChildCount then
				'Done by Manohar Since error occured on this on 03/05/04
					eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value=dRatio
					eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value=dAmount
					'eval("document.formname.txtANALRatio"&iCode).value=dRatio
					'eval("document.formname.txtANALAmount"&iCode).value=dAmount
					nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
					nodANL.Attributes.getNamedItem("Amount").Value=dAmount
					dTotal=CDbl(dTotal)-dAmount
					dRatioTotal=CDbl(dRatioTotal)+dRatio
				else
					'Done by Manohar Since error occured on this on 03/05/04
					eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value=100-dRatioTotal
					eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value=dTotal
					'eval("document.formname.txtANALRatio"&iCode).value=100-dRatioTotal
					'eval("document.formname.txtANALAmount"&iCode).value=dTotal

					nodANL.Attributes.getNamedItem("Ratio").Value=100-dRatioTotal
					nodANL.Attributes.getNamedItem("Amount").Value=dTotal
				end if
				iCounter=CInt(iCounter)+1

			next
		end if 'End of Check for Analytical Child Count
	end if 'End of Check for Analytical Node

	if HeaderNode.nodeName="PayRec" then
		dAmount=CDbl(document.formname.txtAmount.value)
		dTotal=dAmount
		iCounter=1
		iChildCount=HeaderNode.childNodes.length
		if Cint(iChildCount)> 0 then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dTransAmount=nodANL.Attributes.getNamedItem("TransAmount").Value
				dAmtAdjusted=nodANL.Attributes.getNamedItem("AmtAdjusted").Value
				dAmtToAccount=nodANL.Attributes.getNamedItem("AmtToAccount").Value

				dAmtAdjust=CDbl(dTransAmount)-(CDbl(dAmtAdjusted)+CDbl(dAmtToAccount))

				if  CDbl(dAmtAdjust)>CDbl(dTotal) then
					eval("document.formname.txtDocAmount"&iCode).value=FormatNumber(dTotal,2,,,0)
					nodANL.Attributes.getNamedItem("AmtToAdjust").Value=FormatNumber(dTotal,2,,,0)
					dTotal=0
				else
					eval("document.formname.txtDocAmount"&iCode).value=FormatNumber(dAmtAdjust,2,,,0)
					nodANL.Attributes.getNamedItem("AmtToAdjust").Value=FormatNumber(dAmtAdjust,2,,,0)
					dTotal=CDbl(dTotal)-dAmtAdjust
				end if

			next
		end if 'End of Check for PayRec Child Count
	end if 'End of Check for PayRec Node

next

END FUNCTION

Function SetDate()
	Dim sSetDate
	sSetDate = document.formname.hSetInvDate.value
	IF Trim(sSetDate) <> "" Then
		document.formname.ctlDate.SetDate = sSetDate
	End IF
	sFromDate = document.formname.hFromDate.value
	sToDate = document.formname.hToDate.value
	if DateDiff("d",sTodate,date)>0 then
	    document.formname.ctlDate.setmindate = sFromDate
	    document.formname.ctlDate.setMaxDate = sToDate
	    document.formname.ctldate.setDate = sToDate
	else
	    document.formname.ctlDate.setmindate = sFromDate
	    document.formname.ctlDate.setMaxDate = date
	    document.formname.ctldate.setDate = date
	end if
End Function

Function CheckVouStat()
	Dim sCurrDate
	sCurrDate = document.formname.hCurrDate.value
	IF DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		CheckVouStat = false
		Exit Function
	Else
		CheckVouStat = True
	End IF
End Function

'===============================================================================================

Function EditEntry(iVouNo)
	Dim sExp,TempNode,sUom,iCount,AccNode
	sExp = "//Entry[@No="&iVouNo&"]"
	Set TempNode = VouRoot.selectNodes(sExp)
	'MsgBox TempNode.Item(0).xml
	IF TempNode.length <> 0 Then
		document.formname.txtDescription.value = TempNode.Item(0).Attributes.getNamedItem("PayTo").Value
		document.formname.txtDescription.size = Len(TempNode.Item(0).Attributes.getNamedItem("PayTo").Value)+20
		document.formname.txtQty.value = TempNode.Item(0).Attributes.getNamedItem("Qty").Value
		document.formname.txtRate.value = TempNode.Item(0).Attributes.getNamedItem("Rate").Value
		document.formname.txtValue.value = TempNode.Item(0).Attributes.getNamedItem("ActValue").Value
		document.formname.txtDisPercentage.value = TempNode.Item(0).Attributes.getNamedItem("DisPer").Value
		document.formname.txtDisAmount.value = TempNode.Item(0).Attributes.getNamedItem("DisAmount").Value
		document.formname.txtAmount.value = TempNode.Item(0).Attributes.getNamedItem("Amount").Value
		document.formname.hItemCode.value = TempNode.Item(0).Attributes.getNamedItem("ItemCode").Value
		document.formname.hClassCode.value = TempNode.Item(0).Attributes.getNamedItem("ClassCode").Value

		sUom = TempNode.Item(0).Attributes.getNamedItem("UOM").Value


		For iCount = 0 To document.formname.selUOM.length - 1
			IF Trim(document.formname.selUOM(iCount).value) = Trim(sUom) Then
				document.formname.selUOM.selectedIndex = iCount
				Exit For
			End IF
		Next

		For Each AccNode in TempNode.Item(0).childNodes
			IF AccNode.nodeName = "AccHead" Then
				'SelectHead AccNode.Attributes.getNamedItem("No").value,"G",document.formname.selAccountHead,1
				document.formname.selAccountHead.selectedIndex = 1
				window.spAccHead.innerHTML=AccNode.Attributes.getNamedItem("Name").value
			End IF

			IF AccNode.nodeName = "CostCenter" Then
				setADDDisplay 1
				popCostCenter(AccNode)
			End IF

			IF AccNode.nodeName = "Analytical" Then
				setADDDisplay 1
				popAnalytical(AccNode)
			End IF
		Next

		Set EntryRoot = VouRoot.removeChild(TempNode.Item(0))
	End IF

	'MsgBox EntryRoot.xml
	document.formname.hEditEntNo.value = iVouNo
	document.formname.btnNext.disabled = True
	document.formname.btnAdd.disabled = True
	document.formname.btnUpdate.disabled = False
	document.formname.btnDel.disabled = False
End Function

Function DelEntry()

	clearXML()
	setADDDisplay 0
	document.formname.txtDescription.value = ""
	document.formname.txtQty.value = "0.00"
	document.formname.txtDisAmount.value = "0.00"
	document.formname.txtDisPercentage.value = "0.00"
	document.formname.txtRate.value = "0.00"
	document.formname.txtAmount.value = "0.00"
	document.formname.txtValue.value = "0.00"
	document.formname.hEditEntNo.value = "0"
	document.formname.selUOM.selectedIndex = 0
	document.formname.selAccountHead.selectedIndex = 0
	document.formname.btnAdd.disabled = False
	document.formname.btnNext.disabled = False
	document.formname.btnDel.disabled = True
	document.formname.btnUpdate.disabled = True
	DisplayVoucher()

End Function

Function GetItem()
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
sorgID = document.formname.hOrgId.value
sTempValWindowSize = GetWindowSizeForPopup("1")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)

		Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId,ItemData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,ItemData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if


  '  set	OutValue = showModalDialog("../../Common/ItemSelectCommon.asp?orgID=" & sorgID&"&CallFrom=P",ItemData,"dialogHeight:480px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
  '      sQuery = OutValue.getAttribute("PassQuery")
  '      if OutValue.getAttribute("Action")="CLOSE" then exit function

'	while OutValue.getAttribute("Action")<>"Done"
'		set OutValue = showModalDialog("../../Common/ItemSelectCommon.asp?"&sQuery,ItemData,"dialogHeight:480px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
'		    sQuery = OutValue.getAttribute("PassQuery")
'		    if OutValue.getAttribute("Action")="CLOSE" then exit function
'	wend
	if OutValue.hasChildNodes() then
	    For each ndChild in OutValue.childNodes
	        document.formname.txtDescription.value = ndChild.getAttribute("ItemName")
	        document.formname.hItemCode.value = ndChild.getAttribute("ItemCode")
	        document.formname.hClassCode.value = ndChild.getAttribute("ClassCode")
	    Next ' For each ndChild in OutValue.childNodes
	end if  ' if OutValue.hasChildNodes() then

end function
'**********************
Function DisplayBook()
dim iUnitNo,arrTemp
dim Root
iUnitNo = document.formname.hOrgId.value
	    document.formname.selBook.options.length = 1


		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=04&orgID=" & iUnitNo , false
		objhttp.send

		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement
			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if
		popPartyType

end Function
'**************************************************
function selAccHead(objAcc)
dim sOrgId,sPartyType
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp

sRetVal2 = "0:0:0"

sOrgId=document.formname.hOrgId.value
sPartyType=Replace(objAcc.value,"|","?") &"?"& Replace(objAcc.options(objAcc.selectedIndex).text,"&"," and ")

if objAcc.selectedIndex >0 then

	sTempValWindowSize = GetWindowSizeForPopup("2")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)

    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sOrgId&"&Party="&sPartyType,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

'	set OutValue = showModalDialog("../../Common/PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,PartyData,"dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'	    sQuery = OutValue.getAttribute("PassQuery")
'	    if OutValue.getAttribute("Action")="CLOSE" then exit function
'	while OutValue.getAttribute("Action")<>"Done"
'		set OutValue = showModalDialog("../../Common/PartySelection.asp?"&sQuery,PartyData,"dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'		    sQuery = OutValue.getAttribute("PassQuery")
'	        if OutValue.getAttribute("Action")="CLOSE" then exit function
'	wend
	if OutValue.hasChildNodes() then
	    for each ndChild in OutValue.childNodes
	        sRetValue = ndChild.getAttribute("RetField0")
	        sRetValue = sRetValue&":"& ndChild.getAttribute("RetField1")
	        sRetValue = sRetValue&":"& ndChild.getAttribute("RetField2")
	        sRetValue = sRetValue&":"& ndChild.getAttribute("RetField3")
	        sRetValue = sRetValue&":"& ndChild.getAttribute("RetField4")
	    next
    else
		objAcc.selectedIndex = 0
		document.formname.txtPartyName.value = ""
		document.formname.hPartyCode.value = "0"
		Exit Function
	End IF

	sTemp = Split(sRetValue,":")
	sParTy = sTemp(3)
	sParSubType = sTemp(4)
	sParCode = sTemp(1)
	sPartyName = sTemp(0)
	GetPartyHeadXml sParCode,sPartyName,sRetVal2
	Set nodAccHead = AccHeadData.documentElement

	if nodAccHead.hasChildNodes then
		'User Has Selected a Party
		For Each HeaderNode In nodAccHead.childNodes
			document.formname.hPartyCode.value=replace(sPartyType,"|","?")&"?"& HeaderNode.Attributes.Item(0).nodeValue
			document.formname.txtPartyName.value=HeaderNode.Attributes.Item(3).nodeValue
		next
	else
		objAcc.selectedIndex=0
	end if 'End of Party Head Processing
Else
	document.formname.txtPartyName.value = ""
	document.formname.hPartyCode.value = "0"
End if 'End of If any Account Head Selected Check
End function
'****************************
Function VouCreate
dim newElem,objHeader,sTemp

Dim sInvNo,sInvDate,sPartTy,sSendVal,sRetVal,objhttp
Set objhttp = CreateObject("MSXML2.XMLHTTP")

Set Root = VoucherData.documentElement
	if validate then
		set objHeader= VoucherData.createElement("Header")
		Root.appendChild objHeader

		sPartTy = document.formname.hPartyCode.value
		sInvNo = document.formname.txtInvoiceNo.value
		sInvDate = document.formname.ctlDate.GetDate()

		sSendVal = sPartTy&"?"&sInvNo&"?"&sInvDate&"?04"&"?"&document.formname.hOrgId.value
		'MsgBox sSendVal

		'Msgbox sSendVal

		objhttp.Open "GET","CheckInvCreate.asp?sValue="&sSendVal , false
		objhttp.send
		sRetVal = objHttp.responseText

		'Msgbox sRetVal

		IF CStr(sRetVal) <> "C" Then
			MsgBox "Purchase Voucher already Created for this Party,InvoiceNo and Invoice Date "
			Exit Function
		End IF

		Set newElem = VoucherData.createElement("Organization")
		newElem.setAttribute "OrgId",document.formname.hOrgId.value
		newElem.Text=document.formname.hOrgName.value
		objHeader.appendChild newElem

		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
				'document.formname.hBkAccHead.value=HeaderNode.Attributes.Item(2).nodeValue

				Set newElem = VoucherData.createElement("Book")
				newElem.setAttribute "BookId",document.formname.selBook.value
				newElem.setAttribute "BKAccHead",HeaderNode.Attributes.Item(2).nodeValue
				newElem.setAttribute "BKOtherUnits",HeaderNode.Attributes.Item(3).nodeValue
				newElem.Text=document.formname.selBook.options(document.formname.selBook.selectedIndex).Text
				objHeader.appendChild newElem
			end if
		next

		Set newElem = VoucherData.createElement("PurchaseType")
		newElem.setAttribute "PurTypeId",document.formname.selPurtype.value
		newElem.Text=document.formname.selPurtype.options(document.formname.selPurtype.selectedIndex).Text
		objHeader.appendChild newElem

		Set newElem = VoucherData.createElement("PurInvoice")
		newElem.setAttribute "PurInvNo",document.formname.txtInvoiceNo.value
		newElem.setAttribute "PurInvDate",document.formname.ctlDate.GetDate
		objHeader.appendChild newElem

		Set newElem = VoucherData.createElement("Party")

		sTemp=Split(trim(document.formname.hPartyCode.value),"?")

		newElem.setAttribute "ParType",sTemp(0)
		newElem.setAttribute "ParSubType",sTemp(1)
		newElem.setAttribute "ParSubTypeName",document.formname.selPartyType.options(document.formname.selPartyType.selectedIndex).Text
		newElem.setAttribute "ParCode",sTemp(3)

		newElem.Text=document.formname.txtPartyName.value
		objHeader.appendChild newElem
	else
		exit function
	end if

End function

function validate()
	Dim sCurrDate
	sCurrDate = document.formname.hCurrDate.Value

	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select Purchase Book")
		document.formname.selBook.focus
		validate= false
		exit function
	end if
	if document.formname.selPurtype.selectedIndex<1 then
		MsgBox ("Select Purchase type")
		document.formname.selPurtype.focus
		validate=false
		exit function
	end if

	if document.formname.selPartyType.selectedIndex<1 then
		MsgBox ("Select Party")
		document.formname.selPartyType.focus
		validate=false
		exit function
	end if
	if trim(document.formname.txtPartyName.value)="" then
		MsgBox ("Party Name should not be blank")
		document.formname.txtPartyName.select
		validate=false
		exit function
	end if
	if trim(document.formname.txtInvoiceNo.value)="" then
		MsgBox ("Invoice No should not be blank")
		document.formname.txtInvoiceNo.select
		validate=false
		exit function
	end if

	IF DateDiff("d",document.formname.ctlDate.GetDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		validate=false
		Exit Function
	End IF




	validate = True
End function
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate();DisplayBook();">

<form method="POST" name="formname" action="VouPURTaxEntry.asp" >
<input type="hidden" name="hVouCode" value="04">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hSetInvDate" value="<%=sSetInvDate%>">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hItemCode" value="0">
<input type="hidden" name="hClassCode" value="0">
<input type="hidden" name="hSalAccCode" value="<%=iBkAccHead%>">
<input type="hidden" name="hSalAccName" value="<%=sBkAccDesc%>">
<input type="hidden" name="hPartyCode" value="">
<input type="hidden" name="hFromDate" value="<%=sFinFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sFinToDate%>" />

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
			Purchase Voucher
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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
							<!--	<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>-->
								<td class="TabCurrentCell" valign="bottom" align="center" width="110px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105px">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Invoice Details</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="75px">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center"> Advance</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70px">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr><td align="center">Voucher</td></a>
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
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
                                <tr>

                                    <td align="center" width="5" class="ClearPixel" height="1px">
									    <img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								    </td>
								    <td>
								    <table border=0 class=TableOutLineOnly width=100%>
								        <tr>
											<td class="FieldCell" width="108">Purchase Book</td>
											<td class="FieldCell">
												<select size="1" name="selBook" class="FormElem" onchange="GetAccHead(this)">
													<option value="S">Select Book</option>
												</select>
											</td>
											<td class="FieldCell" width="108">Invoice Number</td>
											<td class="FieldCell">
											<input type="text" name="txtInvoiceNo" size="20" class="FormElem" >
                                        	</td>
										</tr>
										<tr>
											<td class="FieldCell" width="108">Purchase Type&nbsp;</td>
											<td class="FieldCell">
												<select size="1" name="selPurType" class="FormElem">
													<option value="0">Select Purchase Type</option>
													<%

														sQuery = "Select PurchaseType,PurchaseTypeName from APP_M_PurchaseTypes Where Active = 'Y' "
														With objRs
															.CursorLocation = 3
															.CursorType = 3
															.Source = sQuery
															.ActiveConnection = con
															.Open
														End with
														Set objRs.Activeconnection = nothing
														Set sCode = objRs(0)
														Set sValue = objRs(1)
														Do while not objRs.EOF
													%>
															<option value="<%Response.Write sCode%>"><%Response.Write sValue%></option>
													<%

															objRs.MoveNext
														Loop
														objRs.Close
													%>
												</select>
											</td>
											<td class="FieldCell" width="120"> Invoice Date</td>
											<td class="FieldCell">
												 <% ' Function Call to Insert Date Picker
															Response.Write InsertDatePicker("ctlDate")
												 %>
							            	</td>
										</tr>

										<tr>
											<td class="FieldCell" width="108">Party Type</td>
											<td class="FieldCell" colspan="3">
												<select size="1" name="selPartyType" class="FormElem" onChange="selAccHead(this)">
												<option value="A">Select Party Type</option>

												</select>
                                          	</td>
										</tr>
										<tr>
											<td class="FieldCell" width="108">Party Name</td>
											<td class="FieldCell" colspan="3"> <input type="text" name="txtPartyName" size="40" class="FormElem"></td>
										</tr>
							        </table>
							    </td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                            <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td height=10></td>
                                </tr>
                             <tr>
                            <td class="FieldCell" width="135">Purchase Account Head</td>
                            <td class="FieldCell" colspan="3">
                                <select size="1" name="selAccountHead" class="FormElem"  onChange="popSalesHead(this) ">
                            <% IF CStr(iBkAccHead) = "0" Then %>
								<option value="S" Selected>Purchase Account Head</option>
								<option value="G">GL Account Head</option>
							<%Else%>
								<option value="S">Purchase Account Head</option>
								<option value="G" Selected>GL Account Head</option>
							<%ENd IF %>
                            </select>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="125"></td>
                            <td class="FieldCell" colspan="3">
                            <% IF CStr(iBkAccHead) = "0" Then %>
								<span class="DataOnly" id="spAccHead"></span>
							<%Else%>
								<span class="DataOnly" id="spAccHead"><b><%Response.Write(sBkAccDesc)%></b> </span>
							<%End IF %>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Item Description</td>
                            <td class="FieldCell" colspan="3">
                            <input type="text" name="txtDescription" size="40" class="FormElem">
                            <a href="#" onClick="GetItem()">
									<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item Description" width="15" height="15">
								</a>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Quantity</td>
                            <td class="FieldCell" colspan="3" align="left">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td><input type="text" name="txtQty" size="15"  maxlength="14" style="text-align: Right" class="FormElem" value="0.00" onBlur="calculateField(1)"></td>
                                <td width="10">
                                </td>
                                <td>
                            <select size="1" name="selUOM" class="FormElem">
                         <%

								sQuery = "Select UoMCode,UoMShortDescription from Ms_UnitOfMeasurement"

							  	With objRs
							  		.CursorLocation = 3
							  		.CursorType = 3
							  		.Source = sQuery
							  		.ActiveConnection = con
							  		.Open
							  	End with
							  	Set objRs.Activeconnection = nothing
							  	Set sCode = objRs(0)
							  	Set sValue = objRs(1)

							  	Do while not objRs.EOF
									Response.Write "<option value="""&sCode&""">"&sValue&"</option>"
									objRs.MoveNext
								Loop
								objRs.Close
							%>
                            </select></td>
                              </tr>
                            </table>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Rate</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtRate" onBlur="calculateField(1)" size="15"  maxlength="13" style="text-align:right" class="FormElem" value="0.00"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Actual Value</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtValue" size="15"  maxlength="13" style="text-align:right" class="FormElem" value="0.00" readonly></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Discount</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="60" class="FieldCell"><input type="text" name="txtDisPercentage" onBlur="calculateField(2)" size="6"  maxlength="5" style="text-align:right" value="0" class="FormElem">%</td>
                                <td>
                            <input type="text" name="txtDisAmount" size="15" onBlur="calculateField(3)"  maxlength="13" style="text-align:right" value="0.00" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>

                                <tr>
                            <td class="FieldCell" width="115">Purchase Value</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtAmount" size="15" readonly maxlength="13" style="text-align:right" class="FormElem" onBlur="popAddAmount1()" value="0.00"></td>
                              </tr>


                            </table>
                                  </td>
                                </tr>

                                <tr>
                            <td class="FieldCell" width="115">Approval</td>
                            <td class="FieldCell" colspan="3">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td class="FieldCell">
                            <input type="radio" value="Y" checked name="optApprove" class="FormElem">
                             Yes&nbsp;&nbsp;
                            <input type="radio" value="N" name="optApprove" class="FormElem"> No
                            </td>
                              </tr>


                            </table>
                                  </td>
                                </tr>

                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="1">
                            &nbsp;
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td >
<DIV class=frmBody id="Disaddtional" style="height:1; visibility: hidden;">
<div id="DisCCANL" class=frmBody style="height:1; visibility: hidden;">
	<table cellpadding="0" cellspacing="0" >
		<tr>
			<td class=MiddlePack colspan="3"> </td>
		</tr>
		<tr>
			<td class=FieldCell>
				<DIV class=frmBody id="DisCost" style="width:280;height:100;">
					<table border="0" id="tblCost" cellspacing="1" class="ExcelTable">
						<tr>
							<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Cost Center Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
						 </tr>
					</table>
				</div><!--End of CostCenter Display Division -->
			</td>
			<td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">                   </td>
			<td class=FieldCell>
				<DIV class=frmBody id="DisAnal" style="width:280; height:100;">

					<table border="0" id="tblAnal" cellspacing="1" class="ExcelTable">
						<tr>
								<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Analytical Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
					    </tr>
					</table>
				</div>	<!--End of Analytical Display Division -->
			</td>
		</tr>
		<tr>
			<td class=MiddlePack  colspan="3"></td>
		</tr>
	</table>
</div> <!--End of CCANAL Display Division -->
</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Add Entry" name="btnAdd" class="ActionButton" onclick="AddEntry('A')" >
                                                                <input type="button" value="Update" onClick="AddEntry('U')" name="btnUpdate" class="ActionButton" disabled>
                                                                <input type="button" value="Delete" onClick="DelEntry()" name="btnDel" class="ActionButton" disabled>
                                                                <input type="button" value="Next" onClick="AddEntry('S')" name="btnNext" class="ActionButton" >

                                                               <input type="button" value="Cancel" name="btnCancel" onClick="Cancel('VouPURBookSelection.asp')" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="35">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" >&nbsp;
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" >
												<DIV class=frmBody id=DisVoucher style="width: 600; height:140;">
                                                <table border="0" id="tblVoucher" cellspacing="1" class="ExcelTable" width="584">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center">Account Head</td>
                                        <td class="ExcelHeaderCell" align="center">Rate</td>
                                        <td class="ExcelHeaderCell" align="center">Quantity</td>
                                        <td class="ExcelHeaderCell" align="center">Value</td>
                                        <td class="ExcelHeaderCell" align="center">Discount</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                            </tr>
                                                </table>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5" >
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
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