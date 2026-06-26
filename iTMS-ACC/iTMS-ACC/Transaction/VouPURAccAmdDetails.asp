<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURAccAmdDetails.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu R
	'Created On					:	Oct 28, 2004
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
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo,sVal,sValTemp
dim sReferenceNo,sInvoiceNo,iBkAccHead,sPartyName,sSetInvDate
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,objfs,iTransNo
Dim sExp,TempNode,sVouDate
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sBkAccDesc
dim sCode,sValue,bCostcenter,bAnalytical,sAMdTy,sFlag

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF
sAMdTy = Request("AmdType")

'Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

sOrgId=Request.Form("selUnitId")
iBookNo=Request.Form("selBook")
sReferenceNo=Request.Form("txtReferenceNo")
sInvoiceNo=Request.Form("txtInvoiceNo")
iBkAccHead=Request.Form("hBkAccHead")
sPartyName=Request.Form("txtPartyName")
sSetInvDate = Request.Form("hInvDate")
sVouDate = Request.Form("hVouDate")

iTransNo = Request("hTransNo")
sFlag = Request("sFlag")

'Response.Write sFlag
'IF objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")) then
'	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"))
'End IF
'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
'Response.Write iTransNo
'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")

oDOM.load server.MapPath("../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml")

Set Root = oDOM.documentElement
sExp = "//Header/Book"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iBkAccHead = TempNode.item(0).Attributes.getNamedItem("BKOtherUnits").Value
End IF

sExp = "//Party"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sPartyName = TempNode.Item(0).Text
End IF

sExp = "//Organization"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sOrgId = TempNode.Item(0).Attributes.getNamedItem("OrgId").Value
	sOrgName = TempNode.Item(0).Text
End IF

sExp = "//Book"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iBookNo = TempNode.Item(0).Attributes.getNamedItem("BookId").Value
End IF

sExp = "//PurInvoice"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sInvoiceNo = TempNode.Item(0).Attributes.getNamedItem("PurInvNo").Value
	sSetInvDate = TempNode.Item(0).Attributes.getNamedItem("PurInvDate").Value
End IF

'Response.Write sInvoiceNo

sExp = "//Details"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sVouDate = TempNode.Item(0).Attributes.getNamedItem("VouDate").Value
End IF

Set objRs = Server.CreateObject("ADODB.RecordSet")
sQuery = "Select M.AccountHead,M.AccountDescription From Acc_M_GLAccountHead M, "&_
		 "VwOrgBookNames V Where V.BookCode = '04' and V.BookNumber = "&iBookNo&" "&_
		 "and V.OUDefinitionID = '"&sOrgId&"' and V.BookAccountHead = M.AccountHead "
objRs.Open sQuery,con
IF Not objRs.EOF Then
	iBkAccHead = objRs(0)
	sBkAccDesc = objRs(1)
Else
	iBkAccHead = 0
	sBkAccDesc = ""
End IF
objRs.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData">
<Details BasicValue="" Discount="" ActualValue="" VouDate=""/></XML>
<XML id="EntryData"><Entry No="0" PayTo="" Amount="" Qty="" UOM="" UOMValue="" Rate="" ActValue="" DisPer="" DisAmount=""/></XML>
<XML id="TaxData"></XML>

<XML id="AccHeadData">
<account/>
</XML>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData" src="<%="../temp/transaction/Voucher AMD_PUR_"&Session.SessionID&".xml"%>"></XML>


<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../../scripts/checkdate.js"></script>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>

<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag,DelRoot
dim iBookAccCode
iEntryNo=1
bVouFlag=false
bSavFlag=false
set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement
set DelRoot=EntryData.documentElement

iBookAccCode=<%=iBkAccHead%>
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
iBookNo=document.formname.hBookcode.value

OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
arrTemp = split(OutValue,":")
while UBound(arrTemp) = 0
	OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")
wend
sRetVal = OutValue
if UBound(arrTemp) <= 1 then exit function
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
		document.formname.txtDescription.value=HeaderNode.Attributes.Item(3).nodeValue
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
Function AddEntry(bFlag)
dim iCode,dRatio,dAmount,sStr,TempNode
dim HeaderNode,nodANL,sChExp,sAcExp,AcNode

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
End if
'If trim(bFlag) = "S" then
'	'sParaXml = 	VoucherData.XML
'	document.formname.action = "AmdAccGenerate.asp"
'	document.formname.submit()
'End If

sStr = "//Details"
Set TempNode = VouRoot.selectNodes(sStr)

IF CStr(bFlag) <> "U" Then
	if (iEntryNo=1 and bFlag="S") or bFlag="A" or bFlag="U"  then
		sAcExp = "//AccHead[@No=0]"
		Set AcNode = EntryRoot.selectNodes(sAcExp)
		'MsgBox AcNode.Item(0).xml
		'if not EntryRoot.hasChildNodes then
		IF AcNode.length <> 0 Then
			Msgbox("Select a Account Head")
			document.formname.selAccountHead.focus
			exit Function
		end if
	end if
Else
	IF document.formname.selAccountHead.selectedIndex = 0 Then
		Msgbox("Select a Account Head")
		document.formname.selAccountHead.focus
		exit Function
	End IF
End IF

IF CStr(bFlag) = "S" and iEntryNo > 1 Then
	SaveXML()
	Exit Function
End IF


if bFlag<>"U" then
	EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
Else
	EntryRoot.Attributes.Item(0).nodeValue=document.formname.hEditEntNo.value
end if


if EntryRoot.hasChildNodes then
	if not checkFileds then exit function
	'EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	EntryRoot.Attributes.Item(1).nodeValue=document.formname.txtDescription.value
	EntryRoot.Attributes.Item(2).nodeValue=FormatNumber(document.formname.txtAmount.value,2,,,0)
	EntryRoot.Attributes.Item(3).nodeValue=document.formname.txtQty.value
	EntryRoot.Attributes.Item(4).nodeValue=document.formname.selUOM.value
	EntryRoot.Attributes.Item(5).nodeValue=document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
	EntryRoot.Attributes.Item(6).nodeValue=FormatNumber(document.formname.txtRate.value,2,,,0)
	EntryRoot.Attributes.Item(7).nodeValue=FormatNumber(document.formname.txtValue.value,2,,,0)
	EntryRoot.Attributes.Item(8).nodeValue=FormatNumber(document.formname.txtDisPercentage.value,2,,,0)
	EntryRoot.Attributes.Item(9).nodeValue=FormatNumber(document.formname.txtDisAmount.value,2,,,0)
	'MsgBox document.formname.hItemCode.value
	'MsgBox document.formname.hClassCode.value
	EntryRoot.Attributes.Item(10).nodeValue = document.formname.hItemCode.value
	EntryRoot.Attributes.Item(11).nodeValue = document.formname.hClassCode.value
	'EntryRoot.Attributes.Item(12).nodeValue = document.formname.selPurType.value

	'IF document.formname.chkVatElg.checked = True Then
	'	EntryRoot.Attributes.Item(13).nodeValue = "Y"
	'Else
	'	EntryRoot.Attributes.Item(13).nodeValue = "N"
	'End IF

	'VouRoot.Attributes.Item(3).nodeValue=document.formname.ctlDate.GetDate


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
		Set insNode = TempNode.item(0).selectNodes(sInsxp)

		IF insNode.length <> 0 Then
			TempNode.item(0).insertBefore EntryRoot,insNode.Item(0)
		Else
			TempNode.item(0).appendChild EntryRoot
		End IF
	Else
		TempNode.item(0).appendChild EntryRoot
	End IF

	'alert VouRoot.xml

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
		document.formname.reset


		IF CStr(document.formname.hSalAccCode.Value) = 0 Then
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

	Dim sExp,TempNode,sFirPurTy,iCnt,sDiffPurTy

	sExp = "//PurInvoice"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.Item(0).value = document.all.tInvNo.innerHtml
		TempNode.Item(0).Attributes.Item(1).value = document.formname.hSetInvDate.value

	End IF

	sExp = "//Details"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.getNamedItem("VouDate").value = document.formname.ctlDate.GetDate()
	End IF


	Set objhttp = CreateObject("Microsoft.XMLHTTP")

	'sExp = "//Entry"
	sDiffPurTy = "N"
	'Set TempNode = VouRoot.SelectNodes(sExp)
	'IF TempNode.length <> 0 Then
		'sFirPurTy = TempNode.Item(0).Attributes.getNamedItem("PurType").Value
		'For iCnt = 0 To TempNode.length - 1
			'MsgBox CStr(TempNode.Item(iCnt).Attributes.getNamedItem("PurType").Value)
		'	IF CStr(sFirPurTy) <> CStr(TempNode.Item(iCnt).Attributes.getNamedItem("PurType").Value) Then
		'		sDiffPurTy = "Y"
		'		Exit For
		'	Else
		'		sDiffPurTy = "N"
		'	ENd IF
		'Next
	'End IF


	'IF CStr(sDiffPurTy) = "N" Then
		document.formname.action="VouPURAmdTaxEntry.asp?sAmdType=A"
	'Else
	'	document.formname.action="VouPURAmdTaxEntryWithMulPurTy.asp?sAmdType=A"
	'End IF

	'Alert(voucherData.xml)
	objhttp.Open "POST","XMLSave.asp?Mod=PUR&Name=Voucher Amd", false
	objhttp.send VoucherData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		document.formname.submit()
		'MsgBox "OK"
	end if
End Function
'---------------------End Of Function SaveXML--------------------------
Function DisplayVoucher()
dim iSno,sAmount,sRate,sQty,sValue,sDiscount,iRow
dim dTotal,sDescription,sUom,sExp,TempNode,iCtr

window.DisVoucher.style.height="200px"
window.DisVoucher.style.visibility="visible"
ClearTable "tblVoucher",1,1
dTotal=0
iRow = 1

set VouRoot=VoucherData.documentElement

sExp = "//Entry"
Set TempNode = VouRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1

			TempNode.Item(iCtr).Attributes.Item(0).nodeValue = iCtr + 1
			iSno=TempNode.Item(iCtr).Attributes.Item(0).nodeValue
			sDescription=TempNode.Item(iCtr).Attributes.Item(1).nodeValue
			sAmount=TempNode.Item(iCtr).Attributes.Item(2).nodeValue
			sRate=TempNode.Item(iCtr).Attributes.Item(6).nodeValue
			sQty=TempNode.Item(iCtr).Attributes.Item(3).nodeValue &"&nbsp;"&TempNode.Item(iCtr).Attributes.Item(5).nodeValue
			sUom = TempNode.Item(iCtr).Attributes.Item(5).nodeValue
			sValue=TempNode.Item(iCtr).Attributes.Item(7).nodeValue
			sDiscount=TempNode.Item(iCtr).Attributes.Item(9).nodeValue

			sAmount = Cdbl(sValue) - Cdbl(sDiscount)
			dTotal=FormatNumber(CDbl(dTotal)+CDbl(sAmount),2,,,0)
			sAmount = FormatNumber(sAmount,2,,,0)


			set oRow = document.all.tblVoucher.insertRow()
			InsertCell oRow,1,"",iRow,"ExcelSerial","Center","top",0,0,0,0,""
			InsertCell oRow,1,"","<a href=""javascript:EditEntry('"&iSno&"')"" class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
			InsertCell oRow,1,"",sDescription,"ExcelDisplayCell","left","top",0,0,0,0,""
			InsertCell oRow,1,"",sRate,"ExcelDisplayCell","right","top",0,0,0,0,""
			InsertCell oRow,1,"",sQty,"ExcelDisplayCell","left","top",0,0,0,0,""
			InsertCell oRow,1,"",sValue,"ExcelDisplayCell","right","top",0,0,0,0,""
			InsertCell oRow,1,"",sDiscount,"ExcelDisplayCell","right","top",0,0,0,0,""
			InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
			iRow = iRow + 1
			iEntryNo = iRow
		next'End of Voucher Node Loop

End IF

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
	if  trim(document.formname.txtQty.value)="" then
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
		EntryRoot.setAttribute "ItemCode","0"
		EntryRoot.setAttribute "ClassCode","0"

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

dim dAmount,iChildCount,dRatio,dTotal,dRatioTotal,iCounter


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
	Dim sSetDate,sExp,TaxNode,TempNode,AdvNode
	Dim sFromYr,sToYr
	'alert VoucherData.xml
	sFromYr = document.formname.hFromYr.Value
	sToYr = document.formname.hToYr.Value
	sFromYr = "01/04/"&Trim(sFromYr)
	sToYr = "31/03/"&sToYr
	document.formname.ctlDate.setMinDate() = sFromYr
	document.formname.ctlDate.setMaxDate() = sToYr

	sSetDate = document.formname.hVouDate.value
	IF Trim(sSetDate) <> "" Then
		document.formname.ctlDate.SetDate = sSetDate
	End IF

	Set VouRoot = VoucherData.documentElement

	sExp = "//TaxDetails"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		Set TaxNode = TempNode.Item(0)
		'TempNode.removeall
		'VouRoot.RemoveChild TaxNode
	End IF

	sExp = "//AdvanceDetails"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		Set AdvNode = TempNode.Item(0)
		'TempNode.removeall
		'VouRoot.RemoveChild AdvNode
	End IF
	'----------Added Newly (Instead of TempNode.removeAll)
	IF VouRoot.haschildNodes then
		For each subNode in VouRoot.childnodes
			If SubNode.NodeName = "TaxDetails" then
				VouRoot.Removechild SubNode
			End IF
			If SubNode.NodeName ="AdvanceDetails" then
				VouRoot.Removechild SubNode
			End IF
		Next
	End If
	'-----------------------------------------------------
	DisplayVoucher()
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
	Dim sExp,TempNode,sUom,iCount,AccNode,iPurType,sVatElg
	sExp = "//PurchaseType"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		iPurType = TempNode.Item(0).Attributes.getNamedItem("PurTypeId").Value
	End IF

	sExp = "//Entry[@No="&iVouNo&"]"
	Set TempNode = VouRoot.selectNodes(sExp)
	'MsgBox TempNode.Item(0).xml
	IF TempNode.length <> 0 Then
		document.formname.txtDescription.value = TempNode.Item(0).Attributes.getNamedItem("PayTo").Value
		'document.formname.txtDescription.size = Len(TempNode.Item(0).Attributes.getNamedItem("PayTo").Value) + 3
		document.formname.txtQty.value = TempNode.Item(0).Attributes.getNamedItem("Qty").Value
		document.formname.txtRate.value = TempNode.Item(0).Attributes.getNamedItem("Rate").Value
		document.formname.txtValue.value = TempNode.Item(0).Attributes.getNamedItem("ActValue").Value
		document.formname.txtDisPercentage.value = TempNode.Item(0).Attributes.getNamedItem("DisPer").Value
		document.formname.txtDisAmount.value = TempNode.Item(0).Attributes.getNamedItem("DisAmount").Value
		document.formname.txtAmount.value = TempNode.Item(0).Attributes.getNamedItem("Amount").Value
		sUom = TempNode.Item(0).Attributes.getNamedItem("UOM").Value

		'iPurType = TempNode.Item(0).Attributes.getNamedItem("PurType").Value
		'sVatElg = TempNode.Item(0).Attributes.getNamedItem("VATElg").Value
		calculateField(1)

		For iCount = 0 To document.formname.selUOM.length - 1
			IF Trim(document.formname.selUOM(iCount).value) = Trim(sUom) Then
				document.formname.selUOM.selectedIndex = iCount
				Exit For
			End IF
		Next

		'For iCount = 0 To document.formname.selPurType.length - 1
		'	IF Trim(document.formname.selPurType(iCount).value) = Trim(iPurType) Then
		'		document.formname.selPurType.selectedIndex = iCount
		'		Exit For
		'	End IF
		'Next

		'IF CStr(sVatElg) = "Y" Then
		'	document.formname.chkVatElg.checked = True
		'Else
		'	document.formname.chkVatElg.checked = False
		'End IF

		For Each AccNode in TempNode.Item(0).childNodes
			IF AccNode.nodeName = "AccHead" Then
				'SelectHead AccNode.Attributes.getNamedItem("No").value,"G",document.formname.selAccountHead,1
				document.all.spAccHead.innerText = AccNode.Attributes.getNamedItem("Name").value
				document.formname.selAccountHead.selectedIndex = 1
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

		Set EntryRoot = TempNode.Item(0)
		'Set EntryRoot = VouRoot.removeChild(TempNode.Item(0))
		TempNode.removeall

		'Set EntryRoot = TempNode.Item(0)
		'Set EntryRoot = VouRoot.removeChild(TempNode.Item(0))
		'TempNode.removeall
		'VouRoot.RemoveChild EntryRoot

		'If VouRoot.haschildNodes then
			'Msgbox "1"
		'	For each SubNode in VouRoot.ChildNodes
		'		'MsgBox SubNode.NodeName
		'		If SubNode.NodeName = "EntryNode" then
		'			'MsgBox "inside "
		'			VouRoot.RemoveChild SubNode
		'		End If
		'	Next
		'End If
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
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">

<form method="POST" name="formname" action="VouPURAmdTaxEntry.asp" >
<input type="hidden" name="hVouCode" value="04">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hSetInvDate" value="<%=sSetInvDate%>">
<input type="hidden" name="hVouDate" value="<%=sVouDate%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hSalAccCode" value="<%=iBkAccHead%>">
<input type="hidden" name="hSalAccName" value="<%=sBkAccDesc%>">
<input type="hidden" name="hAmdType" value="<%=sAMdTy%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hItemCode" value="0">
<input type="hidden" name="hClassCode" value="0">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Voucher Amendment	</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Invoice Details</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="75">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center"> Advance</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
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
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                        <tr>
                            <td class="FieldCell" width="115">Party Name </td>
                            <td class="FieldCell" colspan="3">
                            <span class="DataOnly"><%=sPartyName%></span>
                            </td>


                                </tr>
                                <tr>
                                <%
									'sQuery = "Select OUDefinitionID,BookNumber,PayToRecdFrom,VoucherDate from Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
									'Response.Write sQuery
									 ' 	With objRs
									  '		.CursorLocation = 3
									  '		.CursorType = 3
									  '		.Source = sQuery
									  '		.ActiveConnection = con
									  '		.Open
									  '	End with
									  '	Set objRs.Activeconnection = nothing
										'	sOrgId=objRs("OUDefinitionID")
											'iBookNo=objRs("BookNumber")
											'IF InStr(1,objRs("PayToRecdFrom"),"-") <> 0 Then
											'	sInvoiceNo =Left(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-")-1)
											'	sSetInvDate=Mid(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-"))
											'Else
											'	sInvoiceNo = ""
											'	sSetInvDate = ""
											'End IF

											'if sSetInvDate="" then
											'	sSetInvDate=objRs("VoucherDate")
											'end if
										'objRs.Close
                                %>
                            <td class="FieldCell" width="115">Invoice Number </td>
                            <td class="FieldCell" width="125">
                            	<span class="DataOnly" id="tInvNo"><%=sInvoiceNo%></span>
                            </td>
                            <td width="90" class="FieldCell">
                                Voucher Date
                            </td>
                            <td class="FieldCell">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlDate")
%>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="135">Purchase Account Head</td>
                            <td class="FieldCell" colspan="3">
                            <% IF CStr(iBkAccHead) = "0" Then %>
								<select size="1" name="selAccountHead" class="FormElem" onChange="popSalesHead(this) ">
								<option value="S" Selected>Purchase Account Head</option>
								<option value="G">GL Account Head</option>
							<%Else%>
								<select size="1" name="selAccountHead" class="FormElem" onChange="popSalesHead(this)" disabled>
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
								<span class="DataOnly" id="spAccHead"><%Response.Write(sBkAccDesc)%> </span>
							<%End IF %>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Item Description</td>
                            <td class="FieldCell" colspan="3">
                            <input type="text" name="txtDescription" size="50" class="FormElem"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Quantity</td>
                            <td class="FieldCell" colspan="3" align="left">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td><input type="text" name="txtQty" size="15"  maxlength="14" style="text-align: Right" class="FormElem" value="0.00"></td>
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
                            <input type="text" name="txtValue" size="15"  maxlength="13" style="text-align:right" class="FormElem" value="0.00"></td>
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

                                <!--tr>
                            <td class="FieldCell" width="108">Purchase Type&nbsp;</td>
                            <td class="FieldCell" colspan="3">
                            <select size="1" name="selPurType" class="FormElem" >
									<option value="0">Select Purchase Type</option>
									<%
										'dim sCode,sValue
										sQuery = "Select PurchaseType,PurchaseTypeName from APP_M_PurchaseTypes Where Active = 'Y' Order By PurchaseTypeName "
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
                                </tr-->
                           <!--tr>
									<td class="FieldCell" width="115">VAT Elgibility</td>
									<td class="FieldCell" colspan="3" align="left">
										<input type="checkbox" value="Y" name="chkVatElg" class="FormElem" Checked> YES
									</td>
                            </tr-->

                                <tr>
                            <td class="FieldCell" width="115">Approval</td>
                            <td class="FieldCell" colspan="3">
                            <!--table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td class="FieldCell" -->
                            <input type="radio" value="Y" checked name="optApprove" class="FormElem">
                             Yes&nbsp;&nbsp;
                            <input type="radio" value="N" name="optApprove" class="FormElem"> No
                            </td>
                              <!--/tr>


                            </table>
                                  </td-->
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