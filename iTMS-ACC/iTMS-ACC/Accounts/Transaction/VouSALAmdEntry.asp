<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	VouSALAmdEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	Feburary 14 2003
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
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo,sSalType,saTemp
dim sReferenceNo,sInvoiceNo,iBkAccHead,sPartyName,sInvDate,iTransNo
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,objfs,sExp,TempNode
Dim sVouNo,sVouDate,sSetInvDate,iSalTy,iSalAccHead,sSalAccHdName
Dim sCode,sValue,sAgentName,sSelUOM,sSelPack,sFlag

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("hOrgName")

sSalType=Request.Form("hSalType")
'saTemp=Split(sSalType,"-- --")
'sSalType=saTemp(1)
iBookNo=Request.Form("selBook")
sReferenceNo=Request.Form("txtRefNo")
sInvoiceNo=Request.Form("txtInvoiceNo")
iBkAccHead=Request.Form("hBkAccHead")
sPartyName=Request.Form("txtPartyName")
sInvDate=Request.Form("hInvDate")
iSalTy = Request.Form("selSaleType")
iTransNo = Request("hTransNo")
sFlag = Request("sFlag")
'Response.Write iTransNo
'Response.Write sFlag
'Deletes the prev file
'if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")) then
'	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"))
'End IF

'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
oDOM.load server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")
Set Root = oDOM.documentElement
Set objRs = Server.CreateObject("ADODB.RecordSet")

sExp = "//Party"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sPartyName = TempNode.Item(0).Text
End IF

sExp = "//SalesType"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sSalType = TempNode.Item(0).Text
	iSalTy = TempNode.Item(0).Attributes.Item(0).nodeValue
End IF

sExp = "//Voucher"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sVouNo = TempNode.Item(0).Attributes.getNamedItem("CreatedVouNo").Value
End IF

sExp = "//Organization"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sOrgId = TempNode.Item(0).Attributes.getNamedItem("OrgId").Value
	sOrgName = TempNode.Item(0).Text
End IF

sExp = "//SaleInvoice"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sReferenceNo = TempNode.Item(0).Attributes.getNamedItem("RefNo").Value
	sInvoiceNo = TempNode.Item(0).Attributes.getNamedItem("InvNo").Value
	sInvDate = TempNode.Item(0).Attributes.getNamedItem("InvDate").Value
End IF

sExp = "//Book"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	iBookNo = TempNode.Item(0).Attributes.getNamedItem("BookId").Value
End IF

sExp = "//Details"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sVouDate = TempNode.Item(0).Attributes.getNamedItem("VouDate").Value
End IF

sExp = "//Agent"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sAgentName = TempNode.Item(0).Attributes.getNamedItem("Agentname").Value
End IF



sQuery = "Select AccountHead From App_R_OrgnTaxAccountHead Where TaxCode is Null and TaxCategoryCode  "&_
		 "is Null and InvoiceType = "&iSalTy&" and OUDefinitionID = '"&sOrgId&"'  "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	iSalAccHead = objRs(0)
Else
	iSalAccHead = 0
End IF
objRs.Close

sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where AccountHead = "&iSalAccHead&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sSalAccHdName = objRs(0)
Else
	sSalAccHdName = ""
End IF
objRs.Close



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData">
<Details BasicValue="" Discount="" ActualValue="" VouDate="<%=sInvDate%>"/></XML>
<XML id="EntryData"><Entry No="0" PayTo="" Amount="" Qty="" UOM="" UOMValue="" Rate="" ActValue="" DisPer="" DisAmount="" RndOff="" NoofPack="" PackType="" RatePer="" ItemCode="" ClassCode="" /></XML>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData" src="<%="../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"%>"></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../../scripts/checkdate.js"></script>

<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
dim iBookAccCode

iEntryNo=1
bVouFlag=false
bSavFlag=false


bEditFlag=true
set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement
iBookAccCode = iBkAccHead

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

function popSalesHeadPre(objAcc)
dim sOrgId,sTemp,sDesc,sStr,TempNode,iEntNo
	sOrgId=document.formname.hOrgId.value
	if objAcc.selectedIndex >0 then
				if objAcc.value="G" then
					'showGLHead sOrgId
					iEntNo = document.formname.hEntryNo.value
					sStr="//Entry[@No="&iEntNo&"]/AccHead"
					Set TempNode = VouRoot.selectNodes(sStr)
					IF TempNode.length <> 0 Then

						EntryRoot.appendChild TempNode.Item(0)
					End IF
				else
					sTemp=Split(objAcc.value,"?")
					sDesc=objAcc.options(objAcc.selectedIndex).text
					window.spAccHead.innerHTML=sDesc
					'document.formname.txtDescription.value=sDesc
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

						set oRow = document.all.tblAnal.insertRow(iSno)
						InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","","",0,0,0,0,""
						InsertCell oRow,2,"txtANALRatio"&sCode,dRatio,"ExcelInputCell","","",4,3,0,0,""
						InsertCell oRow,2,"txtANALAmount"&sCode,"0","ExcelInputCell","","",12,10,0,0,""
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

'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAccCode),"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")

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


		bVouFlag=true
		bAnal=HeaderNode.Attributes.Item(1).nodeValue
		bCostCenter=HeaderNode.Attributes.Item(2).nodeValue

		window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue
		document.formname.txtDescription.value=HeaderNode.Attributes.Item(3).nodeValue
		EntryRoot.appendChild HeaderNode
	next
	showCCAnal sOrgId,iAccCode,bCostCenter,bAnal
else
	'User canceled Account Head Selection

	document.formname.txtDescription.value=""
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if 'End of GL Head Processing
set nodAccHead=nothing
End function
'---------------------End Of Function showGLHead--------------------------
Function AddEntry(bFlag)
dim iCode,dRatio,dAmount,sStr,TempNode
dim HeaderNode,nodANL,sChExp,AccNode,nodAccHead,dAccAmt,dRndVal
Dim sBasValue,Entnode,sStr2,iCounter
Dim dNoofPack,dPackTy,dRatePer

sChExp = "//AccHead"
Set AccNode = EntryRoot.selectNodes(sChExp)
IF AccNode.length = 0 Then
	GetGlHeadXmlForSalAcc()
	Set nodAccHead = AccHeadData.documentElement
	For Each HeaderNode in nodAccHead.childNodes
		EntryRoot.appendChild HeaderNode
	Next
End IF

sStr = "//Details"
Set TempNode = VouRoot.selectNodes(sStr)

' New Validation for check blank data - included on 02/04/2004
if bFlag = "S" then


if iEntryNo >= 1 and document.formname.txtAmount.value = "0.00" then
	sStr2 = "//Entry"
	Set EntNode = VouRoot.selectNodes(sStr2)
	IF EntNode.length <> 0 Then
		For iCounter = 0 To EntNode.length - 1
			sBasValue = Cdbl(sBasValue + EntNode.Item(iCounter).Attributes.getNamedItem("Amount").value)
		Next
	End IF
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.getNamedItem("BasicValue").Value = sBasValue
		TempNode.Item(0).Attributes.getNamedItem("ActualValue").Value = sBasValue
	End IF
	SaveXML
	Exit Function
end if
end if
' End of Validation

	if not checkFileds then exit function

	if bFlag<>"U" then
		iEntryNo=iEntryNo+1
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	end if


	IF CStr(bFlag) = "U" Then
		EntryRoot.Attributes.Item(0).nodeValue=document.formname.hEntryNo.value
	Else
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	End IF


	dAccAmt = document.formname.txtAmount.value
	IF document.formname.optRound(0).checked = True Then
		dAccAmt = RndOff(dAccAmt)
		dRndVal = Round(CDbl(dAccAmt) - CDbl(document.formname.txtAmount.value),2)
	Else
		dRndVal = 0
	End IF

	'MsgBox dAccAmt
	'MsgBox dRndVal


	EntryRoot.Attributes.Item(1).nodeValue=document.formname.txtDescription.value
	EntryRoot.Attributes.Item(2).nodeValue=dAccAmt
	EntryRoot.Attributes.Item(3).nodeValue=document.formname.txtQty.value
	EntryRoot.Attributes.Item(4).nodeValue=document.formname.selUOM.value
	EntryRoot.Attributes.Item(5).nodeValue=document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
	EntryRoot.Attributes.Item(6).nodeValue=document.formname.txtRate.value
	EntryRoot.Attributes.Item(7).nodeValue=document.formname.txtValue.value
	EntryRoot.Attributes.Item(8).nodeValue=document.formname.txtDisPercentage.value
	EntryRoot.Attributes.Item(9).nodeValue=document.formname.txtDisAmount.value
	EntryRoot.Attributes.Item(10).nodeValue=dRndVal
	EntryRoot.Attributes.Item(11).nodeValue = document.formname.txtBagno.value
	EntryRoot.Attributes.Item(12).nodeValue = document.formname.selPack.value
	EntryRoot.Attributes.Item(13).nodeValue = document.formname.txtRatePer.value
	EntryRoot.Attributes.Item(14).nodeValue = document.formname.hItemCode.value
	EntryRoot.Attributes.Item(15).nodeValue = document.formname.hClassCode.value


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
				dRatio=eval("document.formname.txtANALRatio"&iCode).value
				dAmount=eval("document.formname.txtANALAmount"&iCode).value
				nodANL.Attributes.Item(3).nodeValue=dRatio
				nodANL.Attributes.Item(4).nodeValue=dAmount
			next
		end if 'End of Check for Analytical Node
	next


	'IF TempNode.length <> 0 Then
	'	TempNode.item(0).appendChild EntryRoot
	'End IF


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
'====================================================================================


	if bFlag="S" then

		sStr2 = "//Entry"
		Set EntNode = VouRoot.selectNodes(sStr2)
		IF EntNode.length <> 0 Then
			For iCounter = 0 To EntNode.length - 1
				sBasValue = Cdbl(sBasValue + EntNode.Item(iCounter).Attributes.getNamedItem("Amount").value)
			Next
		End IF
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("BasicValue").Value = sBasValue
			TempNode.Item(0).Attributes.getNamedItem("ActualValue").Value = sBasValue
		End IF
		SaveXML
	else
		DisplayVoucher
		'iEntryNo=iEntryNo+1
		clearXML()
		document.formname.hItemCode.value = 0
		document.formname.hClassCode.value = 0
		setADDDisplay 0
		document.formname.reset
		window.spAccHead.innerHTML=""

		document.all.spUOM.innerHTML = document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
		document.all.spPack.innerHTML = document.formname.selPack.options(document.formname.selPack.selectedIndex).text

	end if
	document.formname.btnAdd.disabled = false
	document.formname.btnCancel.disabled = false
	document.formname.btnDel.disabled = true
	document.formname.btnNext.disabled = false
	document.formname.btnUpdate.disabled = true

end Function
Function AddEntry1(bFlag)
dim iCode,dRatio,dAmount,sAppType
dim HeaderNode,nodANL,sChExp,AccNode,nodAccHead,dAccAmt,dRndVal
Dim dNoofPack,dPackTy,dRatePer

If Not validate() Then
	Exit Function
End IF

sChExp = "//AccHead"
Set AccNode = EntryRoot.selectNodes(sChExp)
IF AccNode.length = 0 Then
	GetGlHeadXmlForSalAcc()
	Set nodAccHead = AccHeadData.documentElement
	For Each HeaderNode in nodAccHead.childNodes
		EntryRoot.appendChild HeaderNode
	Next
End IF

' New Validation for check blank data - included on 02/04/2004
if bFlag = "S" then
	if iEntryNo >= 1 and document.formname.txtAmount.value = "0.00" then
		SaveXML
		Exit Function
	end if
end if
' End of Validation

if bFlag<>"U" then
	EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
Else
	EntryRoot.Attributes.Item(0).nodeValue=document.formname.hEditEntNo.value
end if

IF document.formname.optApproval(0).checked = True Then
	sAppType = "Y"
Else
	sAppType = "N"
End IF


if not checkFileds then exit function
	'EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	dAccAmt = document.formname.txtAmount.value
	IF document.formname.optRound(0).checked = True Then
		dAccAmt = RndOff(dAccAmt)
		dRndVal = Round(CDbl(dAccAmt) - CDbl(document.formname.txtAmount.value),2)
	End IF
	EntryRoot.Attributes.Item(1).nodeValue = document.formname.txtDescription.value
	EntryRoot.Attributes.Item(2).nodeValue = dAccAmt
	EntryRoot.Attributes.Item(3).nodeValue = document.formname.txtQty.value
	EntryRoot.Attributes.Item(4).nodeValue = document.formname.selUOM.value
	EntryRoot.Attributes.Item(5).nodeValue = document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
	EntryRoot.Attributes.Item(6).nodeValue = document.formname.txtRate.value
	EntryRoot.Attributes.Item(7).nodeValue = document.formname.txtValue.value
	EntryRoot.Attributes.Item(8).nodeValue = document.formname.txtDisPercentage.value
	EntryRoot.Attributes.Item(9).nodeValue = document.formname.txtDisAmount.value
	EntryRoot.Attributes.Item(10).nodeValue = dRndVal
	EntryRoot.Attributes.Item(11).nodeValue = document.formname.txtBagno.value
	EntryRoot.Attributes.Item(12).nodeValue = document.formname.selPack.value
	EntryRoot.Attributes.Item(13).nodeValue = document.formname.txtRatePer.value
	EntryRoot.Attributes.Item(14).nodeValue = document.formname.hItemCode.value
	EntryRoot.Attributes.Item(15).nodeValue = document.formname.hClassCode.value

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

	if bFlag="S" then
	    VouCreate
		SaveXML
	else
		DisplayVoucher
		iEntryNo=iEntryNo+1
		clearXML()
		document.formname.hItemCode.value = 0
		document.formname.hClassCode.value = 0
		setADDDisplay 0
		document.formname.txtDescription.value =""
		document.formname.txtQty.value = "0.00"
		document.formname.txtBagno.value=""
		document.formname.txtAmount.value = "0.00"
		document.formname.txtDisAmount.value="0.00"
		document.formname.txtDisPercentage.value ="0"
		document.formname.txtRate.value = "0.00"
		document.formname.txtRatePer.value = "1"
		document.formname.txtValue.value = "0.00"

		'document.formname.reset
		'This will retain the old value of the approver status.
		IF CStr(sAppType) = "Y" Then
			document.formname.optApproval(0).checked = True
		Else
			document.formname.optApproval(1).checked = True
		End IF

		document.all.spUOM.innerHTML = document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
		document.all.spPack.innerHTML = document.formname.selPack.options(document.formname.selPack.selectedIndex).text

		IF CStr(Trim(document.formname.hSalAccName.Value)) = "" Then
			window.spAccHead.innerHTML=""
			document.formname.selAccountHead.selectedIndex=0
		End IF

		document.formname.btnAdd.disabled = False
		document.formname.btnDel.disabled = True
		document.formname.btnNext.disabled = False
		document.formname.btnUpdate.disabled = True

	end if
end Function

Function DisplayVoucher()
dim iSno,sAmount,sRate,sQty,sValue,sDiscount
dim dTotal,sDescription,sStr,TempNode,iCtr,sQtyWithUom,sUom
Dim AccNode,sAccName,RootTest,dRndVal,dRatePer

Set VouRoot = VoucherData.documentElement
'MsgBox VouRoot.xml

window.DisVoucher.style.height="200px"
window.DisVoucher.style.visibility="visible"
ClearTable "tblVoucher",1,1
dTotal=0


sStr = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr)

IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		TempNode.Item(iCtr).Attributes.Item(0).nodeValue = iCtr+1
		iSno=TempNode.Item(iCtr).Attributes.Item(0).nodeValue
		sDescription=TempNode.Item(iCtr).Attributes.Item(1).nodeValue
		sAmount=FormatNumber(TempNode.Item(iCtr).Attributes.Item(2).nodeValue,2,,,0)

		sRate = TempNode.Item(iCtr).Attributes.Item(6).nodeValue
		sQty = TempNode.Item(iCtr).Attributes.Item(3).nodeValue
		sUom = TempNode.Item(iCtr).Attributes.Item(5).nodeValue
		dRatePer = TempNode.Item(iCtr).Attributes.Item(13).nodeValue
		sValue=FormatNumber(TempNode.Item(iCtr).Attributes.Item(7).nodeValue,2,,,0)
		sDiscount=FormatNumber(TempNode.Item(iCtr).Attributes.Item(9).nodeValue,2,,,0)
		IF Cstr(TempNode.Item(iCtr).Attributes.Item(10).nodeValue) = "" Then
			TempNode.Item(iCtr).Attributes.Item(10).nodeValue = 0
		End IF
		dRndVal = FormatNumber(TempNode.Item(iCtr).Attributes.Item(10).nodeValue,2,,,0)

		'MsgBox dRndVal

		sAmount = CDbl(sQty) * (CDbl(sRate) / CDbl(dRatePer))
		sAmount = CDbl(sAmount) - Cdbl(sDiscount)
		sAmount = CDbl(sAmount) + CDbl(dRndVal)
		'Msgbox sAmount

		'sAmount = CDbl(sAmount) + CDbl(dRndVal)
		sAmount = FormatNumber(sAmount,2,,,0)

		sQtyWithUom = sQty &" "& sUom
		'Msgbox sQtyWithUom

		dTotal=CDbl(dTotal)+CDbl(sAmount)
		dTotal=FormatNumber(dTotal,2,,,0)

		For Each AccNode in TempNode.Item(iCtr).childNodes
			IF CStr(AccNode.nodeName) = "AccHead" Then
				sAccName = AccNode.Attributes.getNamedItem("Name").Value
			End IF
		Next

		sDescription = sAccName &" - " & sDescription


		set oRow = document.all.tblVoucher.insertRow()
		InsertCell oRow,1,"",iCtr+1,"ExcelSerial","Center","top",0,0,0,0,""
		InsertCell oRow,1,"","<a href=""javascript:EditEntry('"&iSno&"')"" class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
		InsertCell oRow,1,"",sDescription,"ExcelDisplayCell","left","top",0,0,0,0,""
		InsertCell oRow,1,"",sQtyWithUom,"ExcelDisplayCell","left","top",0,0,0,0,""
		InsertCell oRow,1,"",sRate,"ExcelDisplayCell","right","top",0,0,0,0,""
		InsertCell oRow,1,"",sValue,"ExcelDisplayCell","right","top",0,0,0,0,""
		InsertCell oRow,1,"",sDiscount,"ExcelDisplayCell","right","top",0,0,0,0,""
		InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""

	next'End of Voucher Node Loop
End IF
	set oRow = document.all.tblVoucher.insertRow()

	InsertCell oRow,1,"","<b>Total</b>","ExcelSerial","right","top",0,0,7,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","left","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",dTotal,"ExcelDisplayCell","right","top",0,0,0,0,""
End Function
'---------------------End Of Function DisplayVoucher----------------------
Function  checkFileds()
	if ValidateAmount(document.formname.txtQty.value,"Quantity",1,9999999.999)=false then
		document.formname.txtQty.select
		checkFileds=false
		exit Function
	elseif 	ValidateAmount(document.formname.txtRate.value,"Rate",1,9999999999.99)=false then
		document.formname.txtRate.select
		checkFileds=false
		exit Function
	elseif ValidateAmount(document.formname.txtDisAmount.value,"Discount",0,9999999999.99)=false then
		document.formname.txtDisAmount.select
		checkFileds=false
		exit Function
	end if
	checkFileds=true
end Function
FUNCTION ValidateAmount(dAmount,sName,dForm,dTo)
	if  trim(dAmount)="" then
		Msgbox(sName+" Cannot be blank")
		ValidateAmount=false
		exit Function
	elseif IsNumeric(dAmount)=false then
		Msgbox("Enter Numeric values for "+sName)
		ValidateAmount=false
		exit Function
	elseif CDbl(dAmount)<CDbl(dForm) or CDbl(dAmount)>CDbl(dTo)  then
		Msgbox(sName+" should be >"+dForm+" and < "+dTo)
		ValidateAmount=false
		exit Function
	end if
	ValidateAmount=true
END FUNCTION
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

	if  trim(document.formname.txtRatePer.value)="" then
		Msgbox("Enter Rate Per")
		document.formname.txtRatePer.select
		calculateField=false
		exit Function
	elseif IsNumeric(document.formname.txtRatePer.value)=false then
		Msgbox("Enter Numeric values for Rate")
		document.formname.txtRatePer.select
		calculateField=false
		exit Function
	elseif Cdbl(document.formname.txtRatePer.value) <= 0 Then
		MsgBox "Rate Per Should be Greater than Zero "
		document.formname.txtRatePer.select
		calculateField=false
		exit Function
	end if


	select case bFlag
		case 1
				document.formname.txtValue.value= FormatNumber(CDbl(document.formname.txtRate.value)/ CDbl(document.formname.txtRatePer.value)* CDbl(document.formname.txtQty.value),2,,,0)
				if CDbl(document.formname.txtDisPercentage.value)>0 then
					document.formname.txtDisAmount.value= FormatNumber(CDbl(document.formname.txtValue.value)* (CDbl(document.formname.txtDisPercentage.value)/100),2,,,0)
				end if
				document.formname.txtAmount.value= FormatNumber(CDbl(document.formname.txtValue.value)- CDbl(document.formname.txtDisAmount.value),2,,,0)
		case 2
				document.formname.txtValue.value= FormatNumber(CDbl(document.formname.txtRate.value)/ CDbl(document.formname.txtRatePer.value)* CDbl(document.formname.txtQty.value),2,,,0)
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
				document.formname.txtValue.value= FormatNumber(CDbl(document.formname.txtRate.value)/ CDbl(document.formname.txtRatePer.value)* CDbl(document.formname.txtQty.value),2,,,0)
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
	calculateField=true
end Function
'---------------------End Of Function calculateField----------------------------
Function SaveXML()
	Dim sExp,TempNode
	sExp = "//SaleInvoice"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		TempNode.Item(0).Attributes.getNamedItem("InvDate").value = document.formname.hInvDate.value
		TempNode.Item(0).Attributes.getNamedItem("InvNo").Value = document.formname.hInvNo.Value
		TempNode.Item(0).Attributes.getNamedItem("RefNo").Value = document.formname.hRefNo.Value
	End IF





	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=Voucher AMD&Mod=SAL", false
	objhttp.send VoucherData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else

		document.formname.btnNext.disabled = True

		document.formname.submit()

		'alert(VoucherData.xml)
	end if
End Function

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
		EntryRoot.setAttribute "RndOff",""
		EntryRoot.setAttribute "NoofPack",""
		EntryRoot.setAttribute "PackType",""
		EntryRoot.setAttribute "RatePer",""
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

Function EditEntry(iVouEntryNo)
Dim sStr,TempNode,iCounter,sTemp,VouRoot,sStr2,AccNode,iAccCode,sAccType,dPackTy
Set VouRoot = VoucherData.documentElement
document.formname.hEntryNo.value = iVouEntryNo
bEditFlag = true
if bEditFlag then
	setADDDisplay 0
	bVouFlag=true
	sStr = "//Entry[@No="&iVouEntryNo&"]"
	sStr2 = "//Entry[@No="&iVouEntryNo&"]/AccHead"
	Set AccNode = VouRoot.selectNodes(sStr2)
	Set TempNode = VouRoot.selectNodes(sStr)
	IF TempNode.length <> 0 Then
		'MsgBox TempNode.Item(0).xml
		document.formname.txtAmount.value = TempNode.item(0).Attributes.getNamedItem("Amount").value
		document.formname.txtQty.value = TempNode.item(0).Attributes.getNamedItem("Qty").value
		document.formname.txtDisPercentage.value = TempNode.item(0).Attributes.getNamedItem("DisPer").value
		document.formname.txtDisAmount.value = TempNode.item(0).Attributes.getNamedItem("DisAmount").value
		document.formname.txtRate.value = TempNode.item(0).Attributes.getNamedItem("Rate").value
		document.formname.txtValue.value = TempNode.item(0).Attributes.getNamedItem("ActValue").value
		document.formname.txtDescription.value = TempNode.item(0).Attributes.getNamedItem("PayTo").value
		document.formname.txtRatePer.value = TempNode.item(0).Attributes.getNamedItem("RatePer").Value
		document.formname.txtBagno.value = TempNode.item(0).Attributes.getNamedItem("NoofPack").Value
		dPackTy = TempNode.item(0).Attributes.getNamedItem("PackType").Value

		IF CDbl(TempNode.item(0).Attributes.item(10).nodeValue) <> 0 Then
			document.formname.optRound(0).checked = True
		Else
			document.formname.optRound(1).checked = True
		End IF

		calculateField(1)
		For iCounter = 0 To document.formname.selUOM.length - 1
			IF Trim(Cstr(document.formname.selUOM(iCounter).value)) = Trim(Cstr(TempNode.item(0).Attributes.getNamedItem("UOMValue").value)) Then
				document.formname.selUOM.selectedIndex = iCounter
				Exit For
			End IF
		Next

		For iCount = 0 To document.formname.selPack.length - 1
			IF Trim(document.formname.selPack(iCount).value) = Trim(dPackTy) Then
				document.formname.selPack.selectedIndex = iCount
				Exit For
			End IF
		Next

		document.all.spUOM.innerHTML = document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
		document.all.spPack.innerHTML = document.formname.selPack.options(document.formname.selPack.selectedIndex).text
		document.formname.hItemCode.value = TempNode.item(0).Attributes.getNamedItem("ItemCode").Value
		document.formname.hClassCode.value = TempNode.item(0).Attributes.getNamedItem("ClassCode").Value


	End IF

	IF AccNode.length <> 0 Then
		iAccCode = Trim(AccNode.Item(0).Attributes.getNamedItem("No").value)
		sAccType = Trim(AccNode.Item(0).Attributes.getNamedItem("Type").value)

		For iCounter = 0 To document.formname.selAccountHead.length - 1
			sTemp = Split(document.formname.selAccountHead(iCounter).value,"?")
			IF sTemp(0) = iAccCode Then
				document.formname.selAccountHead.selectedIndex = iCounter
				Exit For
			End IF
		Next
		IF document.formname.selAccountHead.selectedIndex = 0 Then
			For iCounter = 0 To document.formname.selAccountHead.length - 1
				sTemp = Split(document.formname.selAccountHead(iCounter).value,"?")

				IF Cstr(sTemp(0)) = Cstr(sAccType) Then
					document.formname.selAccountHead.selectedIndex = iCounter
					Exit For
				End IF
			Next
		End IF
		popSalesHeadPre(document.formname.selAccountHead)
	End IF


	'MsgBox "Sucess "
	For Each HeaderNode in TempNode.Item(0).childNodes
		'if HeaderNode.nodeName="AccHead" then
			'For iCounter = 0 To document.formname.selAccountHead.length - 1
				'sTemp = Split(document.formname.selAccountHead(iCounter).value,"?")
				'IF sTemp(0) = HeaderNode.Attributes.getNamedItem("No").value Then
					'document.formname.selAccountHead.selectedIndex = iCounter
					'Exit For
				'End IF
			'Next
		'end if 'End of Check for Account head Node
		if 	HeaderNode.nodeName="CostCenter" then
			setADDDisplay 1
			popCostCenter(HeaderNode)
		end if 'End of Check for Cost Center Node

		if 	HeaderNode.nodeName="Analytical" then
			setADDDisplay 1
			popAnalytical(HeaderNode)
		end if 'End of Check for Analytical Node

		if 	HeaderNode.nodeName="PayRec" then
			popPayRec(HeaderNode)
		end if 'End of Check for Analytical Node
	next 'End of Entry Node Loop
	TempNode.removeAll
	'----------Added Newly (Instead of TempNode.removeAll)
	'alert(VouRoot.xml)

	'If VouRoot.haschildNodes then
	'	For each SubNode in VouRoot.ChildNodes
	'		If SubNode.NodeName = "Details" then
	'			For each SubDetNode in SubNode.childnodes
	'				If SubDetNode.NodeName = "Entry" then
	'					SubNode.RemoveChild SubDetNode
	'				End If
	'			Next
	'		End If
	'	Next
	'End If
	'----------------------------------------------------

	document.formname.btnadd.disabled=true
	document.formname.btnnext.disabled=true
	document.formname.btnupdate.disabled=false
	document.formname.btndel.disabled=false
	bEditFlag=false
	bSavFlag=true
end if

End Function

Function DelEntry
clearXML
	setADDDisplay 0

	DisplayVoucher

	document.formname.txtAmount.value = ""
	document.formname.txtQty.value = ""
	document.formname.txtDisPercentage.value = ""
	document.formname.txtDisAmount.value = ""
	document.formname.txtRate.value = ""
	document.formname.txtValue.value = ""
	document.formname.txtDescription.value = ""
	document.formname.hEditEntNo.value = "0"
	document.formname.reset

	document.formname.btnadd.disabled=false
	document.formname.btnnext.disabled=false
	document.formname.btnupdate.disabled=true
	document.formname.btndel.disabled=true
	bVouFlag=false
	bEditFlag=true
	bSavFlag=true
End Function

Function ChDisp(sObj)
	Dim sDispVal

	sDispVal = sObj.options(sObj.selectedIndex).Text
	IF UCase(sObj.name) = "SELUOM" Then
		document.all.spUOM.innerHTML = sDispVal
	Else
		document.all.spPack.innerHTML = sDispVal
	End IF

End Function

Function GetItem()

	sorgID = document.formname.hOrgId.value
	OutValue = showModalDialog("ItemSelectionPopup.asp?orgID=" & sorgID,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")

	while UBound(arrTemp) = 0
		OutValue = showModalDialog("ItemSelectionPopup.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
	if UBound(arrTemp) <= 1 then exit function
	document.formname.txtDescription.value = arrTemp(0)
	document.formname.hItemCode.value = arrtemp(1)
	document.formname.hClassCode.value = arrtemp(2)

end function
'---------------------End Of Function EditEntry----------------------------

</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="DisplayVoucher()">

<form method="POST" name="formname" action="VouSALAmdTaxEntry.asp">
<input type="hidden" name="hVouCode" value="04">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hInvDate" value="<%=sInvDate%>">
<input type="hidden" name="hRefNo" value="<%=sReferenceNo%>">
<input type="hidden" name="hInvNo" value="<%=sInvoiceNo%>">
<input type="hidden" name="hSalAccCode" value="<%=iSalAccHead%>">
<input type="hidden" name="hSalAccName" value="<%=sSalAccHdName%>">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hItemCode" value="0">
<input type="hidden" name="hClassCode" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher Amendment
		</td>
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
								  		<td align="center">Advance</td>
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
                            <!--tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
							<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
							<span style="cursor: hand" Title="Month wise Balance" >
							<p align="center"><font size="4" face="Webdings">?</font>
							</span>
							</td>
							<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
							<p align="center">
							<span style="cursor: hand" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
							</span>
							</p>
							</td>
							<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
							<p align="center">
							<span style="cursor: hand" Title="Voucher History">
							<font size="4" face="Webdings">?</font>
							</span>
							</p>
							</td>
							<td class="ToolBarCell">
							&nbsp;
							</td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr-->
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" colspan="2">
                              <table border="0" cellspacing="0" width="100%" class="TableOutlineOnly" cellpadding="0">
                                <tr>
                                  <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                <tr>
                                <%
									sQuery = "Select H.OUDefinitionID,H.BookNumber,H.PayToRecdFrom,convert(char,H.VoucherDate,103) VoucherDate,H.CreatedVoucherNo,D.OrgUnitShortDescription from Acc_T_CreatedVoucherHeader H inner join " _
									& "DCS_OrganizationUnitDefinitions D on H.OUDefinitionID=D.OUDefinitionID where H.CreatedTransNo="&iTransNo
									  	With objRs
									  		.CursorLocation = 3
									  		.CursorType = 3
									  		.Source = sQuery
									  		.ActiveConnection = con
									  		.Open
									  	End with
									  	Set objRs.Activeconnection = nothing
									  	if not objRs.EOF then
											sOrgName =objRs("OrgUnitShortDescription")
											iBookNo=objRs("BookNumber")
											sReferenceNo=objRs("CreatedVoucherNo")
											sInvoiceNo =objRs("PayToRecdFrom")
											'sInvoiceNo =Left(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-")-1)
											'sSetInvDate=Mid(objRs("PayToRecdFrom"),InStr(1,objRs("PayToRecdFrom"),"-"))
											sInvDate=objRs("VoucherDate")
										objRs.Close
										end if
                                %>
                                  <!--<td class="FieldCellSub" width="165">Unit Name</td>
                                  <td class="FieldCell"><span class="DataOnly"><%=sOrgName%>&nbsp;<span></td>-->

								  <td class="FieldCellSub" width="100">Voucher No-Date</td>
								  <td class="FieldCell" width="165">
								  <span class="DataOnly"><%=sVouNo%>-<%=sVouDate%></span></td>

                                </tr>
                                <tr>
                                  <td class="FieldCellSub" width="165">Party Name</td>
                                  <td class="FieldCell" colspan="3"> <span class="DataOnly"><%=sPartyName%></span></td>
                                 </tr>

                                 <tr>

                                  <td class="FieldCellSub" width="75">Sale Type</td>
                                  <td class="FieldCell" colspan="2"><span class="DataOnly"><%=sSalType%>&nbsp;<span></td>
                                </tr>
                                <tr>
                                  <td class="FieldCellSub" width="165">Reference / Invoice Number</td>
                                  <td class="FieldCell">                            	<span class="DataOnly"><%=sReferenceNo%>&nbsp;/&nbsp;<%=sInvoiceNo%>&nbsp;</span>
</td>
                                  <td class="FieldCellSub" width="75">Invoice Date</td>
                                  <td class="FieldCell" width="125"><span class="DataOnly"><%=sInvDate%></span>     </td>
                                </tr>
                                <%IF CStr(sAgentName) <> "" Then %>
									<tr>

									  <td class="FieldCellSub" width="75">Agent Name</td>
									  <td class="FieldCell" colspan="2"><span class="DataOnly"><%=sAgentName%>&nbsp;<span></td>
									</tr>
								<%End IF %>

                                <tr>
                                  <td class="MiddlePack" colspan="4"></td>
                                </tr>
                              </table>
                            </td>
                                </tr>
                                <tr>
                            <td class="MiddlePack" colspan="2"></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Sales Account Head</td>
                            <td class="FieldCell">
                            <select size="1" name="selAccountHead" class="FormElem" onChange="popSalesHead(this) ">
							<%IF CStr(iSalAccHead) = "0" Then %>
								<option value="S" Selected>Sales Account Head</option>
								<option value="G">GL Account Head</option>
							<%Else%>
								<option value="S">Sales Account Head</option>
								<option value="G"  Selected>GL Account Head</option>
							<%End IF %>
                            </select>
                            </td>
                            <Input type="hidden" name="hHeadCount" value="1">
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115"></td>
                            <td class="FieldCell">

                            <%IF CStr(iSalAccHead) <> "0" Then %>
								<span class="DataOnly" id="spAccHead"><%=sSalAccHdName%> </span>
							<%Else%>
								<span class="DataOnly" id="spAccHead"></span>
							<%End IF %>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Item Description</td>
                            <td class="FieldCell">
                            <input type="text" name="txtDescription" size="40" class="FormElem">
                            <a href="#" onClick="GetItem()">
									<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item Description" width="15" height="15">
								</a>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Quantity</td>
                            <td class="FieldCell">
							    <table border="0" cellpadding="0" cellspacing="0">
							     <tr>
							       <td width="65"></td>
							       <td><input type="text" name="txtQty" value="0.00" size="13"  maxlength="11" style="text-align: Right" class="FormElem"></td>
							       <td width="10">
							       </td>
							       <td>
							   <select size="1" name="selUOM" class="FormElem" onChange="ChDisp(this)">
							<%

									sQuery = "Select UoMCode,UoMShortDescription from Ms_UnitOfMeasurement"

								  	With objRs
								  		.CursorLocation = 3
								  		.CursorType = 3
								  		.Source = sQuery
								  		.ActiveConnection = con
								  		.Open
								  	End with

								  	IF Not objRs.EOF Then
								  		sSelUOM = objRs(0)
								  	End IF

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
							   <td class="FieldCell">&nbsp;&nbsp;In</td>
							   <td class="FieldCell"><input type="text" name="txtBagno" class="FormElem" size="6" style="text-align: Right"></td>
							   <td>
							   <select size="1" name="selPack" class="FormElem" onChange="ChDisp(this)">
							<%

									sQuery = "Select PackingCode,PackingShortName From APP_M_PackingType Order By PackingShortName "

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

								  	IF Not objRs.EOF Then
								  		sSelPack = objRs(1)
								  	End IF


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
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtRate" value="0.00" onBlur="calculateField(1)" size="15"  maxlength="13" style="text-align:right" class="FormElem"></td>
                            <td class="FieldCell">&nbsp;&nbsp;Per</td>
								<td class="FieldCell"><input type="text" name="txtRatePer" class="FormElem" size="6" style="text-align: Right" value="1" onBlur="calculateField(1)">
								&nbsp;&nbsp;<span id="spUOM" class="ExcelDisplayCell"><%=sSelUOM%> </span>&nbsp;&nbsp;
								In a <span id="spPack" class="ExcelDisplayCell"><%=sSelPack%> </span>
								</td>

                              </tr>

                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Actual Value</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtValue" readonly size="15" value="0.00" maxlength="13" style="text-align:right" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Discount</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="60" class="FieldCell"><input type="text" name="txtDisPercentage" onBlur="calculateField(2)" size="6"  maxlength="5" style="text-align:right" value="0" class="FormElem">%</td>
                                <td>
                            <input type="text" name="txtDisAmount" size="15" value="0.00" onBlur="calculateField(3)"  maxlength="13" style="text-align:right" value="0" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="115">Sales Value</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtAmount" size="15" readonly value="0.00" maxlength="13" style="text-align:right" class="FormElem"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                  <tr>
                            <td class="FieldCell" width="115">Approval</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td class="FieldCell">
                           <Input type="radio" name="optApproval" value="Y" class="FormElem" checked>Yes &nbsp;&nbsp;
                           <Input type="radio" name="optApproval" value="N" class="FormElem">No &nbsp;&nbsp;
                           </td>
                              </tr>
                            </table>
                                  </td>
                                  <tr>
                            <td class="FieldCell">Rounded Off</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td class="FieldCell">
                           <Input type="radio" name="optRound" value="Y" class="FormElem" >Yes &nbsp;&nbsp;
                           <Input type="radio" name="optRound" value="N" class="FormElem" checked>No &nbsp;&nbsp;
                           </td>
                              </tr>

                            </table>
                                  </td>
                                </tr>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="1">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                                                                 <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
                                                                <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >
                                                                <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >
                                                                <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouSalBookSelection.asp')" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="35">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                                        <td class="ExcelHeaderCell" align="center">Account Head - Item Description</td>
                                        <td class="ExcelHeaderCell" align="center" width="60">Quantity</td>
                                        <td class="ExcelHeaderCell" align="center">Rate</td>
                                        <td class="ExcelHeaderCell" align="center">Value</td>
                                        <td class="ExcelHeaderCell" align="center">Discount</td>
                                        <td class="ExcelHeaderCell" align="center">Amount</td>
                                            </tr>
                                                </table>
												</div>
								</td>
								<td align="center" class="ClearPixel" width="5" >
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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