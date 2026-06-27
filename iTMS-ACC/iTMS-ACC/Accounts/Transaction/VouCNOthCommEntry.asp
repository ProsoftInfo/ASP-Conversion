<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNOthCommEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	Feburary 14 2003
	'Modified By				:	Manohar Prabhu.R
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo
dim sBookName,sInvoiceNo,sTemp,arrPartyCode,sPartyCode,sPartyName
Dim sInvTemp,iCtr,sVouTemp,sVouchTy,sNarr,sAmount,sTempAmt,sUserid
Dim oDom,Root,MainNode,PartyNode,sVouchDetails

Set oDom = Server.CreateObject("Microsoft.XMLDOM")

sUserid = getUserID()

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("horgName")
iBookNo=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sInvoiceNo=Request.Form("selInvoiceNo")
sVouchTy = Request.Form("selVoucherType")
sVouchDetails = Request.Form("hVouDetails")
sInvTemp = Split(sInvoiceNo,",")

sVouTemp = Split(Request.Form("hVouDetails"),":")

sPartyName=Request.Form("txtPartyName")
arrPartyCode=split(Request.Form("hPartyCode"),"?")

Set objRs = Server.CreateObject("ADODB.RecordSet")

Set Root = oDom.createElement("Root")
oDom.appendChild Root
For iCtr = 0 To UBound(sInvTemp)
	Set MainNode = oDom.createElement("voucher")
	MainNode.setAttribute "UnitNo", sOrgId
	MainNode.setAttribute "UnitName", sOrgName
	MainNode.setAttribute "BookNo", iBookNo
	MainNode.setAttribute "BookName", sBookName
	MainNode.setAttribute "VouDate", ""
	MainNode.setAttribute "Approver", ""
	MainNode.setAttribute "SalTransNo", sInvTemp(iCtr)
	MainNode.setAttribute "SalVouNo", ""
	MainNode.setAttribute "SalVouDate", ""

	Set PartyNode = oDom.createElement("Party")
	PartyNode.setAttribute "ParType", trim(arrPartyCode(0))
	PartyNode.setAttribute "ParSubType", trim(arrPartyCode(1))
	PartyNode.setAttribute "ParCode", trim(arrPartyCode(3))
	PartyNode.text = sPartyName

	MainNode.appendChild PartyNode
	Root.appendChild MainNode
Next

oDOM.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-CNCommEntry.xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData" src="<%="../Temp/Transaction/"&Session.SessionID&"-CNCommEntry.xml"%>">
</XML>

<XML id="EntryData">
<Entry No="0" Payto="" Amount="" CRDR="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></XML>
</XML>
<XML id="AccHeadData">
<account/>
</XML>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>

<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
iEntryNo=1
bVouFlag=false
bSavFlag=false
set VouRoot=DetData.documentElement
set EntryRoot=EntryData.documentElement

function showNarration()
dim sOrgId,sBookCode,sNarration

sOrgId=document.formname.hOrgId.value
sBookCode="07?"&document.formname.hBookcode.value

sNarration = showModalDialog("NarrationSelection.asp?orgId="+sOrgId&"&BookCode="&sBookCode,"","")
if sNarration<>"" then document.formname.txtNarration.value=sNarration
End Function

function selGLHead(objAcc)
dim sOrgId,sTemp,sDesc
	sOrgId=document.formname.hOrgId.value
	if objAcc.selectedIndex >0 then
				if objAcc.value="G" then
					showGLHead sOrgId
				else
					sTemp=Split(objAcc.value,"?")
					sDesc=objAcc.options(objAcc.selectedIndex).text
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
'---------------------End Of Function selGLHead-------------------
function showGLHead(sOrgId)
dim iAccCode,bAnal,bCostCenter
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,arrTemp,sRetVal,sTemp2,sTdsElgi,sTempVal
iBookNo=document.formname.hBookcode.value

OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=07&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
arrTemp = split(OutValue,":")

while UBound(arrTemp) = 0
	OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")
wend

sRetVal = OutValue
sTempVal = OutValue

if UBound(arrTemp) <= 1 then exit function
sTemp2 = Split(sTempVal,":")
sTdsElgi = sTemp2(6)
document.formname.hTdsElgi.value = sTdsElgi
IF CStr(sTdsElgi) = "1" Then
	document.formname.txtTdsAmount.disabled = False
	document.formname.txtTdsper.disabled = False
Else
	document.formname.txtTdsAmount.disabled = True
	document.formname.txtTdsper.disabled = True
End IF
GetGlHeadXml(sRetVal)

Set nodAccHead = AccHeadData.documentElement


'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=07&BookNo="+iBookNo,"","")

if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		iAccCode=HeaderNode.Attributes.Item(0).nodeValue
		bAnal=HeaderNode.Attributes.Item(1).nodeValue
		bCostCenter=HeaderNode.Attributes.Item(2).nodeValue

		document.formname.txtPayTo.value=HeaderNode.Attributes.Item(3).nodeValue
		EntryRoot.appendChild HeaderNode
	next
	showCCAnal sOrgId,iAccCode,bCostCenter,bAnal
else
	'User canceled Account Head Selection

	document.formname.txtPayTo.value=""
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if 'End of GL Head Processing
set nodAccHead=nothing
End function
'---------------------End Of Function showGLHead--------------------------
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


Function AddEntrySC(sVal)
'MsgBox EntryRoot.xml

dim iCode,dRatio,dAmount
dim HeaderNode,nodANL,sStr,TempNode,iCtr

set VouRoot = DetData.documentElement

document.formname.txtAmount.readonly = false
document.formname.txtAmount.className = "FormElem"
document.formname.txtNarration.readOnly = false

sStr = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr)

IF CStr(sVal) <> "S" Then
	IF document.formname.selAccHead.selectedIndex = 0 Then
		MsgBox "Select Account Head "
		document.formname.selAccHead.focus()
		Exit Function
	End IF
ElseIF CStr(sVal) = "S" Then

	IF document.formname.selAccHead.selectedIndex = 0 Then
		SaveXML()
		Exit Function
	End IF
End IF

if not checkFileds then exit function
	IF CStr(sVal) <> "U" Then
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	Else
		EntryRoot.Attributes.Item(0).nodeValue = document.formname.hEditEntry.value
	End IF

	EntryRoot.Attributes.Item(1).nodeValue=document.formname.txtPayTo.value
	EntryRoot.Attributes.Item(2).nodeValue=document.formname.txtAmount.value
	IF document.formname.OptCRDR(0).checked = True Then
		EntryRoot.setAttribute "CRDR","C"
	Else
		EntryRoot.setAttribute "CRDR","D"
	End IF
	EntryRoot.Attributes.getNamedItem("TdsAmount").value=document.formname.txtTdsAmount.value
	EntryRoot.Attributes.getNamedItem("TdsElgi").value=document.formname.hTdsElgi.value
	EntryRoot.Attributes.getNamedItem("TdsPercentage").value=document.formname.txtTdsper.value

	sStr = "//voucher"
	Set TempNode = VouRoot.selectNodes(sStr)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			TempNode.Item(iCtr).Attributes.Item(4).nodeValue = document.formname.ctlDate.GetDate
		Next
	End IF

	Set newElem = EntryData.createElement("Narration")
	newElem.text= document.formname.txtNarration.value
	EntryRoot.appendChild newElem

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

'====== This is to Insert/append the the entry in same order as on the creation ==
	IF CStr(sVal) = "U" Then
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

	IF CStr(sVal) = "A" Then
		DisplayVoucher()
		iEntryNo=iEntryNo+1
		bVouFlag=false
		sTransFlag="A"
		clearXML()
		document.formname.selAccHead.selectedIndex = 0
		document.formname.txtPayTo.value = " "
		document.formname.txtNarration.value = " "
		document.formname.txtTdsAmount.value = "0.00"
		document.formname.txtTdsper.value = "0.00"
		document.formname.txtAmount.value = "0.00"
		setADDDisplay 0

		if iEntryNo>1 then
			document.formname.txtPayTo.readOnly=true
		end if

	Elseif CStr(sVal) = "U" Then
		DisplayVoucher()
		bVouFlag=false
		sTransFlag="A"
		clearXML()
		document.formname.selAccHead.selectedIndex = 0
		document.formname.txtPayTo.value = " "
		document.formname.txtNarration.value = " "
		document.formname.txtTdsAmount.value = "0.00"
		document.formname.txtTdsper.value = "0.00"
		document.formname.txtAmount.value = "0.00"
		document.formname.hEditEntry.value = "0"
		setADDDisplay 0
		document.formname.btnAdd.disabled = False
		document.formname.btnDel.disabled = True
		document.formname.btnNext.disabled = False
		document.formname.btnUpdate.disabled = True
	Else
		SaveXML
	End IF

end Function

Function DelEntry()
	clearXML
	setADDDisplay 0

	DisplayVoucher

	document.formname.txtPayTo.value=""
	document.formname.reset

	document.formname.btnadd.disabled=false
	document.formname.btnnext.disabled=false
	document.formname.btnupdate.disabled=true
	document.formname.btndel.disabled=true
	bVouFlag=false
	bEditFlag=true
	bSavFlag=true
End Function

FUNCTION DisplayVoucher()
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTdsAmt,sTdsPer
dim dTotal,sAccUnit,sTotalCRDR,sStr,TempNode,iRow
sAccUnit = document.formname.hOrgName.value

window.DisVoucher.style.height="200px"
window.DisVoucher.style.visibility="visible"
ClearTable "tblVoucher",1,1
dTotal=0
iRow = 1
dTotal = 0

For Each EntryNode in VouRoot.childNodes
	IF CStr(EntryNode.nodeName) = "Entry" Then

		EntryNode.Attributes.Item(0).nodeValue = iRow
		iSno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		sTdsAmt = EntryNode.Attributes.Item(4).nodeValue
		sTdsPer = EntryNode.Attributes.Item(6).nodeValue
		sAmount=FormatNumber(CDbl(sAmount),2,,,0)
		sTdsAmt=FormatNumber(CDbl(sTdsAmt),2,,,0)
		sTdsPer=FormatNumber(CDbl(sTdsPer),2,,,0)

		IF CStr(EntryNode.Attributes.Item(3).nodeValue) = "D" Then
			dTotal = dTotal + CDbl(sAmount)
		Else
			dTotal = dTotal - CDbl(sAmount)
		End IF

		'sAmount=sAmount&"&nbsp;"&EntryNode.Attributes.Item(1).nodeValue&"r"
		sAddtional=""
		For Each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					if HeaderNode.Attributes.Item(4).nodeValue="P" then
						sAccount=HeaderNode.Attributes.Item(3).nodeValue
					else
						sAccount= HeaderNode.Attributes.Item(3).nodeValue
					end if
			end if 'End of Check for Account head Node
			if 	HeaderNode.nodeName="Narration" then
					sNarration=HeaderNode.text
					sNarration = Mid(sNarration,1,Len(sNarration)-1)
			end if 'End of Check for Narration Node
			if 	HeaderNode.nodeName="CostCenter" then
					for each  nodANL in HeaderNode.childNodes
						sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue&"-"
						sAddtional=sAddtional&nodANL.Attributes.Item(3).nodeValue &"%&nbsp;"
						sAddtional=sAddtional&nodANL.Attributes.Item(4).nodeValue&"<br>"
					next
			end if 'End of Check for Cost Center Node
			if 	HeaderNode.nodeName="Analytical" and HeaderNode.hasChildnodes then
					sAddtional=sAddtional&"---------------------------  <br>"
					for each  nodANL in HeaderNode.childNodes
						sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue&"-"
						sAddtional=sAddtional&nodANL.Attributes.Item(3).nodeValue &"%&nbsp;"
						sAddtional=sAddtional&nodANL.Attributes.Item(4).nodeValue&"<br>"
					next
			end if 'End of Check for Analytical Node
			if 	HeaderNode.nodeName="PayRec" then
					for each  nodANL in HeaderNode.childNodes
						sAddtional=sAddtional&nodANL.Attributes.Item(1).nodeValue&":"
						sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue &"-&nbsp;"
						sAddtional=sAddtional&nodANL.Attributes.Item(5).nodeValue&"<br>"
					next
			end if 'End of Check for Analytical Node
		next 'End of Entry Node Loop
		set oRow = document.all.tblVoucher.insertRow()
		InsertCell oRow,1,"",iRow,"ExcelSerial","Center","top",0,0,0,0,""
		InsertCell oRow,1,"","<a href=""javascript:EditEntry('"&iSno&"')"">Edit</a>","ExcelDisplayCell","Center","top",0,0,0,0,""
		InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
		InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
		InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
		InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
		InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""
		InsertCell oRow,1,"",sTdsAmt,"ExcelDisplayCell","left","top",0,0,0,0,""
		InsertCell oRow,1,"",sTdsPer,"ExcelDisplayCell","left","top",0,0,0,0,""

		iRow = iRow + 1
	End IF


next'End of Voucher Node Loop
	if dTotal < 0 then
		sTotalCRDR="&nbsp;Cr"
		dTotal=CDbl(dTotal)*-1
	else
		sTotalCRDR="&nbsp;Dr"
	end if

	dTotal="Rs. &nbsp;"&FormatNumber(dTotal,2,,,0)

	set oRow = document.all.tblVoucher.insertRow(iSno+1)
	InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,5,0,""
	InsertCell oRow,1,"",CStr(dTotal)&sTotalCRDR ,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,3,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,3,0,""
END FUNCTION



Function  checkFileds()
	if document.formname.selAccHead.selectedIndex = 0 Then
		MsgBox "Select Account Head "
		document.formname.selAccHead.focus()
		checkFileds=false
		exit Function
	end if

	if  trim(document.formname.txtNarration.value)="" then
		Msgbox("Enter Narration")
		document.formname.txtNarration.select
		checkFileds=false
		exit Function
	end if

	if CDate(document.formname.ctlDate.GetDate) < CDate(document.formname.hInvDate.value) then
		Msgbox("Credit Note date should be >= Invoice date")
		document.formname.ctlDate.focus
		checkFileds=false
		exit Function
	end if
	checkFileds=true
end Function
'---------------------End Of Function checkFileds--------------------------
Function SaveXML()
	IF CheckApp() Then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLSave.asp?Mod=CN&Name=Voucher Entry", false
		objhttp.send DetData.XMLDocument
		if objhttp.responseText <> "" then
			Msgbox(objhttp.responseText)
		else
			document.formname.btnNext.disabled = True
			document.formname.submit()
		end if
	End IF
End Function

Function clearXML()
	Set EntryRoot = EntryData.createElement("Entry")
		EntryRoot.setAttribute "No",iEntryNo
		EntryRoot.setAttribute "PayTo",""
		EntryRoot.setAttribute "Amount",""
		EntryRoot.setAttribute "CRDR",""
		EntryRoot.setAttribute "TdsAmount",""
		EntryRoot.setAttribute "TdsElgi","0"
		EntryRoot.setAttribute "TdsPercentage",""

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

Function SelMisParty()
	Dim arrTemp,sRetValue,sParCode,sPartyName,sTemp

	OutValue = showModalDialog("MisPartySelection.asp","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	IF CStr(OutValue) = "AN" Then
		AddNewParty()
		Exit Function
	End IF
	arrTemp = split(OutValue,":")


	while UBound(arrTemp) = 0
		OutValue = showModalDialog("MisPartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend

	sRetValue = OutValue
	'MsgBox sRetValue
	if UBound(arrTemp) <= 1 then exit function

	sTemp = Split(sRetValue,":")
	document.formname.txtPayTo.value = sTemp(0)
	'sParTy = sTemp(4)
	'sParSubType = sTemp(3)
	'sParCode = sTemp(1)
	'sPartyName = sTemp(0)
End Function

Function AddNewParty()
	OutValue = showModalDialog("MisParCreate.asp?"&OutValue,"","dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'MsgBox OutValue
	document.formname.txtPayTo.value = OutValue
End Function

Function EditEntry(iVouEntryNo)
	Dim sCheckExp,CheckNode
	'if bEditFlag then
		setADDDisplay 0
		'MsgBox "OK "
		'setPayableDisplay 0
		bVouFlag=true
		'window.spEntryNo.innerHTML=iVouEntryNo
		document.formname.hEditEntry.value = iVouEntryNo

		sCheckExp = "//Entry[@TdsAmount]"
		Set CheckNode = VouRoot.selectNodes(sCheckExp)

		For Each EntryNode in VouRoot.childNodes
			if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
				document.formname.txtAmount.value=EntryNode.Attributes.Item(2).nodeValue
				if EntryNode.Attributes.Item(3).nodeValue ="C" then
						document.formname.OptCRDR(0).checked=true
				else
						document.formname.OptCRDR(1).checked=true
				end if

				'sAccUnit=EntryNode.Attributes.Item(5).nodeValue

				document.formname.txtPayTo.value = EntryNode.Attributes.Item(1).nodeValue
				IF CheckNode.length <> 0 Then
					document.formname.txtTdsAmount.value = EntryNode.Attributes.Item(4).nodeValue
					document.formname.txtTdsper.value = EntryNode.Attributes.Item(6).nodeValue

					IF CStr(EntryNode.Attributes.Item(5).nodeValue) = "1" Then
						document.formname.txtTdsAmount.disabled = False
						document.formname.txtTdsper.disabled = False
					Else
						document.formname.txtTdsAmount.disabled = True
						document.formname.txtTdsper.disabled = True
					End IF
					document.formname.hTDSElgi.value = EntryNode.Attributes.Item(5).nodeValue
				Else
					document.formname.txtTdsAmount.value = "0.00"
					document.formname.txtTdsper.value = "0.00"
					document.formname.hTDSElgi.value = "0"
				End IF

				For Each HeaderNode in EntryNode.childNodes
					if HeaderNode.nodeName="AccHead" then
						'SelectHead HeaderNode.Attributes.getNamedItem("No").value,"G",document.formname.selAccHead,1
						document.formname.selAccHead.selectedIndex = 1
					end if 'End of Check for Account head Node

					if 	HeaderNode.nodeName="Narration" then
						document.formname.txtNarration.value=HeaderNode.text
					end if 'End of Check for Narration Node

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

				set EntryRoot=VouRoot.removeChild(EntryNode)

			end if
		next'End of Voucher Node Loop

		document.formname.btnadd.disabled=true
		document.formname.btnNext.disabled=true
		document.formname.btnupdate.disabled=false
		document.formname.btndel.disabled=false
		bEditFlag=false
		bSavFlag=true
	'end if
End Function

</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="VouCNOthCommGen.asp">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hInvDate" value="0">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hVouchTy" value="<%=sVouchTy%>">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hEditEntry" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Commission
          Entry
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
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
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<tr><td align="center">Voucher</td>
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
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: hand" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">�</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Daywise Balance"><font size="3" face="Webdings">�</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Voucher History">
                    <font size="4" face="Webdings">�</font>
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
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="590">
                                    <tr>
										<td class="FieldCell" width="93">Unit</td>
										<td colspan="3"><span class="DataOnly"><%=sOrgName%>&nbsp;</span></td>

	                                </tr>
									<tr>
										<td class="FieldCell" width="93">Agent Name</td>
										<td width="230"><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
									<%IF CStr(sVouchTy) <> "SC" Then %>
										<td class="FieldCell" width="100">Invoice No-Date</td>
										<td><span class="DataOnly">1&nbsp;-&nbsp;1&nbsp;</span></td>
									<%End IF %>
	                                </tr>
									<!--tr>
										<td class="FieldCell" width="113">Commision Amount</td>
										<td width="230"><span class="DataOnly"><%=FormatNumber(sAmount,2,,,0)%></span></td>
										<td class="FieldCell" width="100"></td>
										<td></td>
	                                </tr-->
	                                <tr>
										<td class="FieldCell" width="113">Entry Type</td>
										<td width="230" class="FieldCellSub">
										<Input type="radio" name="OptCRDR" value="C" class="FormElem">Credit &nbsp;&nbsp;
										<Input type="radio" name="OptCRDR" value="D" class="FormElem" checked>Debit &nbsp;&nbsp;</td>
										<td class="FieldCell" width="100"></td>
										<td></td>
	                                </tr>

                                    </table>

								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="1">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" >
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="5" width="139"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Head</td>
                                                    <td class="FieldCell">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selGLHead(this)">
															<option value="A">Select Account Head</option>
															<option value="G">General Ledger</option>

                                                    </select>
													 </td>
                                                    <td class="FieldCell" colspan="2"><p align="center">Date
                                                    </td>
                                                    <td class="FieldCell"> <p align="center">
                                                    <% ' Function Call to Insert Date Picker
															Response.Write InsertDatePicker("ctlDate")
													%>

														</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td>
 </td>
                                                    <td colspan="2"><p align="center"><!--Number--></p>
                                                    </td>
                                                    <td class="FieldCellSub">  </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td class="FieldCell" colspan="4">
                                                    <input type="text" name="txtPayTo" size="40" class="Formelem">
                                                    &nbsp; <a href="javascript:SelMisParty()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a>
                                                    </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139" valign="top">Narration</td>
                                                    <td class="FieldCell" colspan="2" valign="top">

														<textarea rows="3" name="txtNarration" cols="50" class="FormElem"><%=Trim(sVouchDetails)%></textarea> </td>

                                                    <td class="FieldCell" colspan="2" valign="middle">
 </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell" colspan="4">

														<input type="text" name="txtAmount" value="" size="15" style="text-align:right" class="Formelem"> </td>

                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Deduction @</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    </td>
                                                        </tr>

                                                        <tr>
															<td class="FieldCellSub" width="139">Approval</td>
															<td class="FieldCell" colspan="4">
																<Input type="radio" name="optApprove" class="FormElem" value="Y" checked onClick="SetApp('Y')">Yes &nbsp;
																&nbsp;&nbsp;&nbsp;
																<Input type="radio" name="optApprove" class="FormElem" value="N" onClick="SetApp('N')">No
																&nbsp;&nbsp;&nbsp; Immediate Approver &nbsp;&nbsp; <select size="1" name="selUserId" class="FormElem">
																		<option value="I">Immediate Approver</option>
																		<%=populateEmployeeWithVal(sUserId)%>
																		    </select>
															</td>
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

																 <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntrySC('A')" class="ActionButton" >
																 <input type="Button" value="Update" name="btnUpdate" onClick="AddEntrySC('U')" class="ActionButton" disabled>
																 <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" class="ActionButton" disabled>



                                                                <input type="button" value="Next" onClick="AddEntrySC('S')" name="btnNext" class="ActionButton" >

                                                                <input type="reset" value="Cancel" name="B8" class="ActionButton" >

														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="35">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

							<tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
<DIV class=frmBody id="DisVoucher" style="width:585; visibility:hidden; height:1;">
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="700">
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
		<td class="ExcelHeaderCell" align="center" width="75">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="125">Amount</td>
		<td class="ExcelHeaderCell" align="center" >Additional Details</td>
		<td class="ExcelHeaderCell" align="center" width="80">Deduction Amount</td>
		<td class="ExcelHeaderCell" align="center" width="80">Deduction Percentage</td>

	</tr>
	</table>
</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
<tr>
								<td align="center" class="BottomPack" colspan="3">
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