<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNOtherEntry.asp
	'Module Name				:	ACCOUNTS (Transcation Debit Note Amendment For Other Voucher Type)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 31,2011
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo
dim sBookName,sInvoiceNo,sTemp,arrPartyCode,sPartyCode,sPartyName
Dim sInvTemp,iCtr,sVouTemp,sVouchTy,sNarr,sAmount,sTransno,ODom
Dim sStr,TempNode,VouRoot,sUserID
Dim sVouDate,sVouUnit,sVouNumber,sVouAmt,sSelVouTy,sSelInvNo
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sCallFrom,sVouCode,sVouName

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
'sOrgId=Request.Form("selUnitId")
sOrgId = Session("organizationcode")
iBookNo=Request.Form("selBook")
'Response.Write "iBookNo = "& iBookNo
sUserID = getUserID()

sPartyName=Request.Form("txtPartyName")
'sOrgName=Request.Form("horgName")
sOrgName = Session("OrgShortName")
sBookName=Request.Form("hBookName")
sPartyCode=Request.Form("hPartyCode")
sVouTemp = Request.Form("hVouDetails")
sSelVouTy = Request.Form("selVoucherType")
sSelInvNo = Request.Form("selInvoiceNo")

Set objRs = Server.CreateObject("ADODB.RecordSet")
sCallFrom = Request("CallFrom")
if Trim(sCallFrom)="GJ" then
    sVouCode = "08"
    sVouName = "GJ"
else
    sVouCode = "06"
    sVouName = "DN"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData">
<Root>

</Root>
</XML>
<xml id="GLHeadData"><Root /></xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>

<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=iBookNo%>" BookName="<%=sBookName%>" CRDR="D" VouDate="" PartyCode="<%=sPartyCode%>" PartyName="<%=Replace(sPartyName,"&","and")%>" Approver=""/></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName=""/>
</XML>

<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<xml id="GJVoucher"></xml>
<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
iEntryNo=1
bVouFlag=false
bSavFlag=false
bEditFlag = True

set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement
'************************
Function AddNew()
    if trim(document.formname.hAction.value)="Edit" then
        AddEntry("U")
    else
        AddEntry("A")
    end if
End Function
'*************************
function showNarration()
dim sOrgId,sBookCode,sNarration

sOrgId=document.formname.hOrgId.value
sBookCode="06?"&document.formname.hBookcode.value

sNarration = showModalDialog("NarrationSelection.asp?orgId="+sOrgId&"&BookCode="&sBookCode,"","")
if sNarration<>"" then document.formname.txtNarration.value=sNarration
End Function

FUNCTION selAccountHead(objAcc)

DIM sVouType,sOrgId,sTemp,iHeadCount,sDesc
iHeadCount=cint(document.formname.hHeadCount.value)
	if objAcc.selectedIndex >0 then

		if document.formname.hOtherUnitFlag.value=1 then
			if document.formname.selAccUnitId.selectedIndex <=0 then

				objAcc.selectedIndex=0
				document.formname.selAccUnitId.focus
			else
				sOrgId=document.formname.selAccUnitId.value

				if objAcc.selectedIndex <= iHeadCount then

					sTemp=Split(objAcc.value,"?")
					sDesc=objAcc.options(objAcc.selectedIndex).text
					bVouFlag=true
					Set newElem = EntryData.createElement("AccHead")
						newElem.setAttribute "No", trim(sTemp(0))
						newElem.setAttribute "CostCenter", trim(sTemp(1))
						newElem.setAttribute "Analytical", trim(sTemp(2))
						newElem.setAttribute "Name", sDesc
						newElem.setAttribute "Type", "G"
						newElem.setAttribute "TransFalg", trim(sTemp(3))
	   					EntryRoot.appendChild newElem

						sTransFlag=trim(sTemp(3))
						window.spAccHead.innerHTML=sDesc&"&nbsp;"


						document.formname.txtPayto.value = document.formname.hPayto.value


					showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))

				elseif objAcc.selectedIndex =iHeadCount+1 then
						showGLHead sOrgId
				else
					sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
					showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
				End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
				document.formname.txtNarration.focus
			end if	'END OF ACCOUNTING UNIT SELECTED OR NOT
		else
			sOrgId=document.formname.hOrgId.value
			if objAcc.selectedIndex <= iHeadCount then
					sTemp=Split(objAcc.value,"?")
					sDesc=objAcc.options(objAcc.selectedIndex).text
					bVouFlag=true
					Set newElem = EntryData.createElement("AccHead")
						newElem.setAttribute "No", trim(sTemp(0))
						newElem.setAttribute "CostCenter", trim(sTemp(1))
						newElem.setAttribute "Analytical", trim(sTemp(2))
						newElem.setAttribute "Name", sDesc
						newElem.setAttribute "Type", "G"
						newElem.setAttribute "TransFalg", trim(sTemp(3))
	   					EntryRoot.appendChild newElem

						sTransFlag=trim(sTemp(3))
						window.spAccHead.innerHTML=sDesc&"&nbsp;"
						document.formname.txtPayto.value = document.formname.hPayTo.value


					showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))

				elseif objAcc.selectedIndex =iHeadCount+1 then
						showGLHead sOrgId
				else
					sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
					showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
				End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
			document.formname.txtNarration.focus
		end if	'END OF BOOK HAS OTHER UNIT TRANSCATION OR NOT CHECK
	End if 'END OF IF ANY ACCOUNT HEAD SELECTED CHECK

END FUNCTION

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
dim sCode,sDesc,dRatio,iBookNo,arrTemp,sRetVal
iBookNo=document.formname.hBookcode.value

'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=07&BookNo="+iBookNo,"","")
set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
sQuery = OutValue.getAttribute("PassQuery")
if OutValue.getAttribute("Action")="CLOSE" then exit function

while OutValue.getAttribute("Action")<>"Done"
	set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	sQuery = OutValue.getAttribute("PassQuery")
	if OutValue.getAttribute("Action")="CLOSE" then exit function
wend

if OutValue.hasChildNodes() then
    for each ndChild in OutValue.childNodes
        sRetVal = ndChild.getAttribute("RetField0")
        sRetVal =  sRetVal  &":"& ndChild.getAttribute("RetField1")
        sRetVal =  sRetVal  &":"& ndChild.getAttribute("RetField2")
        sRetVal =  sRetVal  &":"& ndChild.getAttribute("RetField3")
        sRetVal =  sRetVal  &":"& ndChild.getAttribute("RetField4")
        sRetVal =  sRetVal  &":"& ndChild.getAttribute("RetField5")
        sRetVal =  sRetVal  &":"& ndChild.getAttribute("RetField6")
        sRetVal =  sRetVal  &":"& ndChild.getAttribute("RetField7")
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
						sGroupCode = nodANL.Attributes.Item(5).nodeValue

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

Function AddEntry(sVal)
dim iCode,dRatio,dAmount,TempNode
bVouFlag = true
bflag = sVal
'MsgBox bFlag
' New Validation for check blank data - included on 02/04/2004
if bFlag = "S" then
	if Trim(document.formname.txtAmount.value) = "0.00" then
		SaveXML
		Exit Function
	end if
end if
' End of Validation

if bVouFlag then
	if not checkFileds then exit function
	bSavFlag=true

	if bFlag<>"U" then
		iEntryNo=iEntryNo+1
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	end if

	VouRoot.Attributes.getNamedItem("VouDate").Value=document.formname.ctlDate.getdate
	VouRoot.Attributes.getNamedItem("VouDate").Value=document.formname.ctlDate.getdate
	VouRoot.Attributes.getNamedItem("UnitNo").Value=document.formname.hOrgId.value
	VouRoot.Attributes.getNamedItem("UnitName").Value=document.formname.hOrgName.value

	EntryRoot.Attributes.getNamedItem("CRDR").Value="C"
	EntryRoot.Attributes.getNamedItem("Payto").Value=document.formname.txtPayTo.value
	EntryRoot.Attributes.getNamedItem("Amount").Value=document.formname.txtAmount.value
	EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.hOrgId.value
	EntryRoot.Attributes.getNamedItem("AccName").Value=document.formname.hOrgName.value

	IF CStr(bFlag) <> "U" Then
		Set newElem = EntryData.createElement("Narration")
		newElem.text= document.formname.txtNarration.value
		EntryRoot.appendChild newElem
	Else
		for each HeaderNode in EntryRoot.childNodes
			IF CStr(HeaderNode.nodeName) = "Narration" Then
				HeaderNode.text = document.formname.txtNarration.value
			End IF
		Next
	End IF

	for each HeaderNode in EntryRoot.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dRatio=eval("document.formname.txtCCRatio"&iCode).value
				dAmount=eval("document.formname.txtCCAmount"&iCode).value
				nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
				nodANL.Attributes.getNamedItem("Amount").Value=dAmount
			next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value

				dRatio=eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value
				dAmount=eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value

				nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
				nodANL.Attributes.getNamedItem("Amount").Value=dAmount
			next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dAmount=eval("document.formname.txtDocAmount"&iCode).value
				nodANL.Attributes.getNamedItem("AmtToAdjust").Value=dAmount
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
		document.formname.hAction.value = "New"
	Else
	    window.spEntryNo.innerHTML = iEntryNo
		VouRoot.appendChild EntryRoot
	End IF
'====================================================================================

	if bFlag="S" then
		SaveXML
	else
		DisplayVoucher
		clearXML()
		setADDDisplay 0
		document.formname.txtPayTo.value=""
		document.formname.reset
	'	document.formname.btnadd.disabled=false
	'	document.formname.btnnext.disabled=false
	'	document.formname.btnupdate.disabled=true
	'	document.formname.btndel.disabled=true
		bEditFlag=true
		bVouFlag=false
	end if
else
	if bFlag="S" then
		SaveXML
	End if
end if
end Function

FUNCTION DisplayVoucher()

dim sNarration,sAccount,sAddtional,iSno,sAmount
dim dTotal,sAccUnit,iRow,idivFixed,idivHeight
set VouRoot=VoucherData.documentElement
idivFixed = 60
idivHeight = iEntryNo*cint(25)
window.DisVoucher.style.height=cint(idivFixed) +cint(idivHeight)&"px"
window.DisVoucher.style.visibility="visible"

ClearTable "tblVoucher",1,1
dTotal=0

iEntryNo=0
iRow = 1

sAccUnit =  VouRoot.Attributes.Item(1).nodeValue
'sDate=VouRoot.Attributes.Item(5).nodeValue
'document.formname.ctlDate.setDate=sDate
For Each EntryNode in VouRoot.childNodes
	IF EntryNode.nodeName = "Entry" Then
		iEntryNo=cint(iEntryNo)+1
		EntryNode.Attributes.Item(0).nodeValue=iEntryNo
		sAmount=FormatNumber(EntryNode.Attributes.Item(3).nodeValue,2,,,0) & "&nbsp;" & EntryNode.Attributes.Item(1).nodeValue
		dTotal=dTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)

		sAddtional=""
		For Each HeaderNode in EntryNode.childNodes
			if HeaderNode.nodeName="AccHead" then
					if HeaderNode.Attributes.Item(4).nodeValue="P" then
						sAccount=HeaderNode.Attributes.Item(3).nodeValue
					else
						sAccount=HeaderNode.Attributes.Item(0).nodeValue
						sAccount=sAccount& "-" & HeaderNode.Attributes.Item(3).nodeValue
					end if
			end if 'End of Check for Account head Node
			if 	HeaderNode.nodeName="Narration" then
					sNarration=HeaderNode.text
			end if 'End of Check for Narration Node
			if 	HeaderNode.nodeName="CostCenter" then
				for each  nodANL in HeaderNode.childNodes
					sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue&"-"
					sAddtional=sAddtional&nodANL.Attributes.Item(3).nodeValue &"%&nbsp;"
					sAddtional=sAddtional&nodANL.Attributes.Item(4).nodeValue&"<br>"
				next
			end if 'End of Check for Cost Center Node
			if 	HeaderNode.nodeName="Analytical" then
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
	InsertCell oRow,1,"","<img src='../../assets/images/iTMS%20icons/Deleteicon.gif' onClick=EditEntry('"&iEntryNo&"','D')>","ExcelDisplayCell","Center","top",0,0,0,0,""
	InsertCell oRow,1,"","<a class=""ExcelDisplaylink"" href=""javascript:EditEntry('"&iEntryNo&"','E')"">Edit</a>","ExcelDisplayCell","Center","top",0,0,0,0,""
	'InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""

	iRow = iRow + 1
End IF

next'End of Voucher Node Loop
	'set oRow = document.all.tblVoucher.insertRow(iEntryNo+1)
	'InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,6,0,""
	'InsertCell oRow,1,"",CStr(dTotal) ,"ExcelDisplayCell","right","top",0,0,0,0,""

	'window.spAccHead.innerHTML=""
	'window.spEntryNo.innerHTML=iEntryNo+1
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

	checkFileds=true
end Function
'---------------------End Of Function checkFileds--------------------------
Function SaveXML()
		'	set objhttp = CreateObject("Microsoft.XMLHTTP")
		'	objhttp.Open "POST","XMLSave.asp?Mod=DN&Name=Voucher Entry", false
		'	objhttp.send VoucherData.XMLDocument
		'	if objhttp.responseText <> "" then
		'		Msgbox(objhttp.responseText)
		'	else
		'		'MsgBox "OK "
		'		document.formname.btnNext.disabled = True
		'		document.formname.action = "VouCNGenerate.asp"
		'		document.formname.submit()
		'	end if
if trim(document.formname.hCallFrom.value)="DR" then
	IF CheckApp() Then
		'IF CheckNoSer() Then
			set objhttp = CreateObject("Microsoft.XMLHTTP")
			objhttp.Open "POST","XMLSave.asp?Mod=DN&Name=Voucher Entry", false
			objhttp.send VoucherData.XMLDocument
			if objhttp.responseText <> "" then
				Msgbox(objhttp.responseText)
			else
				document.formname.btnNext.disabled = True
				document.formname.action="VouCNGenerate.asp"
				document.formname.submit()
			end if
		'End IF
	End IF
else 
	set oNodRoot = VoucherData.documentElement
	set oDGjRoot = GJVoucher.createElement("voucher")
	GJVoucher.appendChild oDGjRoot 
			iCnt = 0
			for each oNodEntry in oNodRoot.childNodes
			sTempUnitNo = oNodRoot.getAttribute("UnitNo")
			sTempUnitName = oNodRoot.getAttribute("UnitName")
			sArrParVal = split(oNodRoot.getAttribute("PartyCode"),"?")
			sPartyDet = sTempUnitNo &"&ParSubType="&sArrParVal(1)&"&ParType="&sArrParVal(0)&"&PartyCode="&sArrParVal(3)
			oDGjRoot.setAttribute "UnitNo",oNodRoot.getAttribute("UnitNo")
			oDGjRoot.setAttribute "UnitName",oNodRoot.getAttribute("UnitName")
			oDGjRoot.setAttribute "BookNo",oNodRoot.getAttribute("BookNo")
			oDGjRoot.setAttribute "BookName",oNodRoot.getAttribute("BookName")
			oDGjRoot.setAttribute "CRDR",""
			oDGjRoot.setAttribute "VouDate",oNodRoot.getAttribute("VouDate")
			oDGjRoot.setAttribute "BookAcchead","0"
			oDGjRoot.setAttribute "Approver",oNodRoot.getAttribute("Approver")
				if oNodEntry.nodeName="Entry" then
					if setFlag = False then
						'First Entry
						set oDGjEntry= GJVoucher.createElement("Entry")
						iCnt = iCnt + 1
						oDGjEntry.setAttribute "No",iCnt
						oDGjEntry.setAttribute "CRDR","D"
						oDGjEntry.setAttribute "Payto","0"
						oDGjEntry.setAttribute "Amount",oNodEntry.getAttribute("Amount")
						oDGjEntry.setAttribute "AccUnit",sTempUnitNo 
						oDGjEntry.setAttribute "AccName",sTempUnitName 
						oDGjEntry.setAttribute "TdsAmount","0.00"
						oDGjEntry.setAttribute "TDSElgi","0"
						oDGjEntry.setAttribute "TdsPercentage","0"
						oDGjEntry.setAttribute "PayRecAmount","0"
						oDGjRoot.appendChild oDGjEntry 
						dTotalAmount =  oNodEntry.getAttribute("Amount")
							
								set objhttp = CreateObject("Microsoft.XMLHTTP")
								objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sPartyDet, false
								objhttp.send

								IF objhttp.responseText <> "" Then
									sRetVal2 = objhttp.responseText
									sArrValue = split(sRetVal2,":")
								End IF
								
								set oDGjAcc = GJVoucher.createElement("AccHead")
								oDGjAcc.setAttribute "No",oNodRoot.getAttribute("PartyCode") 
								oDGjAcc.setAttribute "Pay",sArrValue(0)
								oDGjAcc.setAttribute "Rec",sArrValue(1)
								oDGjAcc.setAttribute "Name",sPartyName 
								oDGjAcc.setAttribute "Type","P"
								oDGjAcc.setAttribute "Adv",sArrValue(2)
								oDGjEntry.appendChild oDGjAcc 
								
							set NodRecCount = GJVoucher.createElement("RecCount")
							NodRecCount.setAttribute "Val","1"
							oDGjEntry.appendChild NodRecCount 
							Set oDGjNarr = GJVoucher.CreateElement("Narration")
							oDGjNarr.Text = document.formname.txtNarration.value
							oDGjEntry.appendChild oDGjNarr
						
						setFlag = True
					end if ' 	if setFlag = False then
						
					'Second Entry
					set oDGjEntry= GJVoucher.createElement("Entry")
					iCnt = iCnt + 1
					oDGjEntry.setAttribute "No",iCnt
					oDGjEntry.setAttribute "CRDR","C"
					oDGjEntry.setAttribute "Payto",""
					oDGjEntry.setAttribute "Amount",dTotalAmount
					oDGjEntry.setAttribute "AccUnit",sTempUnitNo 
					oDGjEntry.setAttribute "AccName",sTempUnitName 
					oDGjEntry.setAttribute "TdsAmount","0.00"
					oDGjEntry.setAttribute "TDSElgi","0"
					oDGjEntry.setAttribute "TdsPercentage","0"
					oDGjEntry.setAttribute "PayRecAmount","0"
					oDGjRoot.appendChild oDGjEntry 
						for each oNodDeatils in oNodEntry.childNodes
							if oNodDeatils.nodeName="AccHead" then
								set oDGjAcc = GJVoucher.createElement("AccHead")
								oDGjAcc.setAttribute "No",oNodDeatils.getAttribute("No")
								oDGjAcc.setAttribute "CostCenter",oNodDeatils.getAttribute("CostCenter")
								oDGjAcc.setAttribute "Analytical",oNodDeatils.getAttribute("Analytical")
								oDGjAcc.setAttribute "Name",oNodDeatils.getAttribute("Name")
								oDGjAcc.setAttribute "Type","G"
								oDGjAcc.setAttribute "TransFlag","A"
								oDGjEntry.appendChild oDGjAcc 
							end if
						next
				end if
				exit for ''added by ragav on Sep 27 for avoid the Multiple Entry
			next
	
			dInvAmount = dTotalAmount
			dBasicTotal = dInvAmount
			dTotal	= dInvAmount
				
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLSave.asp?Mod=GJ&Name=Voucher Entry", false
		objhttp.send GJVoucher.XMLDocument
		if objhttp.responseText <> "" then
			Msgbox(objhttp.responseText)
		else
			document.formname.action="VouGenerate.asp"
			document.formname.submit()
		end if
end if'if trim(document.formname.hCallFrom.value)="DR" then
	
	
End Function

Function clearXML()
	Set EntryRoot = EntryData.createElement("Entry")
	EntryRoot.setAttribute "No",iEntryNo
	EntryRoot.setAttribute "CRDR",""
	EntryRoot.setAttribute "Payto",""
	EntryRoot.setAttribute "Amount",""
	EntryRoot.setAttribute "AccUnit",""
	EntryRoot.setAttribute "AccName",""
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

Function EditEntry(iVouEntryNo,sEditType)
Dim iCounter
if trim(sEditType)="E" then
    document.formname.hAction.value = "Edit"
    if bEditFlag then
	    document.formname.hEntryNo.value = iVouEntryNo
	    setADDDisplay 0
	    'setPayableDisplay 0
	    bVouFlag=true

	    sAccUnit = VouRoot.Attributes.Item(0).nodeValue
	    sVouDate = VouRoot.Attributes.Item(5).nodeValue

	    document.formname.ctlDate.setDate = svoudate

	    For Each EntryNode in VouRoot.childNodes
		    if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
			    document.formname.txtAmount.value=EntryNode.Attributes.Item(3).nodeValue
			    'sAccUnit=EntryNode.Attributes.Item(5).nodeValue
			    sAddtional=""
			    For Each HeaderNode in EntryNode.childNodes
				    if HeaderNode.nodeName="AccHead" then
					    For iCounter = 0 To document.formname.selAccHead.length - 1
						    IF CStr(HeaderNode.Attributes.getNamedItem("Type").value) = CStr(document.formname.selAccHead(iCounter).value) Then
							    document.formname.selAccHead.selectedIndex = iCounter
							    Exit For
						    End IF
					    Next

					    document.formname.txtPayTo.value=HeaderNode.Attributes.Item(3).nodeValue
					    'window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue
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
			    'MsgBox EntryRoot.xml
		    end if
	    next'End of Voucher Node Loop
	   ' document.formname.btnadd.disabled=true
	    'document.formname.btnnext.disabled=true
	    'document.formname.btnupdate.disabled=false
	    'document.formname.btndel.disabled=false
	    bEditFlag=false
	    bSavFlag=true
    end if
else ' if trim(sEditType)="E" then
    document.formname.hEntryNo.value = iVouEntryNo
	    setADDDisplay 0
	    bVouFlag=true
	    sAccUnit = VouRoot.Attributes.Item(0).nodeValue
	    sVouDate = VouRoot.Attributes.Item(5).nodeValue
	    document.formname.ctlDate.setDate = svoudate
	    For Each EntryNode in VouRoot.childNodes
		    if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
			    set EntryRoot=VouRoot.removeChild(EntryNode)
		    end if
	    next'End of Voucher Node Loop
	    bEditFlag=false
	    bSavFlag=true
	    DelEntry
end if 'if trim(sEditType)="E" then
End Function

Function DelEntry()
	clearXML
	setADDDisplay 0
	DisplayVoucher

	document.formname.txtPayTo.value=""

	document.formname.reset
   window.spEntryNo.innerHTML = iEntryNo
'	document.formname.btnadd.disabled=false
'	document.formname.btnnext.disabled=false
'	document.formname.btnupdate.disabled=true
'	document.formname.btndel.disabled=true
	bVouFlag=false
	bEditFlag=true
	bSavFlag=true
End Function

Function SetDate()
	Dim sFromYr,sToYr
	sFromYr = document.formname.hFromYr.Value
	sToYr = document.formname.hToYr.Value
	sFromYr = "01/04/"&Trim(sFromYr)
	sToYr = "31/03/"&sToYr
	document.formname.ctlDate.setMinDate() = sFromYr
	document.formname.ctlDate.setMaxDate() = sToYr
End Function



</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">

<form method="POST" name="formname">
<input type="hidden" name="hVouCode" value="<%=sVouCode%>">
<input type="hidden" name="hVouName" value="<%=sVouName%>">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hTransNo" value="0">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hSelVouTy" value="<%=sSelVouTy%>">
<input type="hidden" name="hInvNos" value="<%=sSelInvNo%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hAction" value="New" >
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Debit Note Other

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
							<!--tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr-->
                            <tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">

                            </table>
                            </td>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
                            <!--tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr-->
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border=0 cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
										<td class="FieldCell" width="75">Agent Name</td>
										<td ><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
										<td class="FieldCell" width="75">Voucher Date</td>
										<td class="FieldCell" width="75"> <p align="center">
                                        <% ' Function Call to Insert Date Picker
												Response.Write InsertDatePicker("ctlDate")
										%>

										</td>

	                                </tr>
									<tr>


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
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td class="MiddlePack" colspan="5"></td>
                                                        </tr>
                                                        <tr>
                                                            <td class="FieldCellSub" width="95">Accounting Head</td>
                                                            <td class="FieldCell">
                                                                    <select size="1" name="selAccHead" class="FormElem" onChange="selGLHead(this)">
															        <option value="A">Select Account Head</option>
															        <option value="G">General Ledger</option>
                                                            </select>
                                                            <input type="hidden" name="hHeadCount" value="1">
													         </td>
													         <td colspan=3 class=FieldCellSub>Entry No&nbsp;&nbsp;<span id="spEntryNo" class="DataOnly">1</span></td>
                                                        </tr>
                                                        <tr>
                                                            <td class="FieldCellSub" width="95"></td>
                                                            <td></td>
                                                            <td colspan="2"><p align="center"><!--Number--></p>
                                                            </td>
                                                            <td class="FieldCellSub">  </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="FieldCellSub" width="95"></td>
                                                            <td class="FieldCell" colspan="4">
                                                            <input type="text" name="txtPayTo" size="40" class="Formelem"> </td>
                                                        </tr>
                                                        <tr>
                                                                <td class="FieldCellSub" width="95" valign="top">Narration&nbsp;&nbsp;&nbsp;
														            <a href="javascript:showNarration()"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
                                                                </td>
                                                                <td class="FieldCell" colspan="2" valign="top">

														            <textarea rows="3" name="txtNarration" cols="50" class="FormElem" onKeyPress="ChkEnter()"></textarea> </td>

                                                                <td class="FieldCell" colspan="2" valign="middle">
                                                                </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="95">Amount</td>
                                                    <td class="FieldCell" colspan="4">

														<input type="text" name="txtAmount" value="0.00" size="15" style="text-align:right" class="Formelem" onblur="popAddAmount()"> </td>

                                                        </tr>

                                                        <tr>
														    <td class="FieldCellSub" width="95">Approval</td>
														    <td class="FieldCell" colspan="3">
														    <input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="SetApp('Y')">Yes&nbsp;
														    <input type="radio" value="N" name="optApprove" class="FormElem" onClick="SetApp('N')"> No
														    &nbsp; Immediate Approver &nbsp;
														    <select size="1" name="selUserId" class="FormElem">
															    <option value="I">Immediate Approver</option>
															    <%=populateEmployeeWithVal(sUserId)%>
														    </select>
														    &nbsp;<input type="Button" value="Add Entry" name="btnAdd" onClick="AddNew()" class="AddButton" >
														    </td>
													    </tr>
    					                                <tr>
								                            <td colspan=4 align=center>
                                                                    <DIV class=frmBody id="DisVoucher" style="width:95%; visibility:hidden; height:1;">
	                                                                    <table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="95%">
	                                                                    <tr>
		                                                                    <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="35">&nbsp;</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="35">&nbsp;</td>
		                                                                    <!--<td class="ExcelHeaderCell" align="center" width="75">AU</td>-->
		                                                                    <td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		                                                                    <td class="ExcelHeaderCell" align="center" width="100">Amount</td>
		                                                                    <td class="ExcelHeaderCell" align="center" >Additional Details</td>
	                                                                    </tr>
	                                                                    </table>
                                                                    </div>
								                            </td>
                                                        </tr>
                                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="1">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                            <!--tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr-->
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">

																 <!--<input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >-->
                                                                <!--<input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >-->
                                                                <input type="button" value="Save" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouCNBookSelection.asp')" class="ActionButton" >


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