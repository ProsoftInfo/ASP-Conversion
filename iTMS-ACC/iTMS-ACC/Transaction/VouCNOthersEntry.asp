
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNOthersEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
dim sOrgId,sOrgName,sBookName,objRs,sQuery,iBookNo
dim sPartyCode,sPartyName,sUserID,sSelVouTy,sVouNarr,sSelInvNo
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sCallFrom,sVouCode,sVouName

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

'sOrgId=Request.Form("selUnitId")
sOrgId = Session("organizationcode")
sOrgName = Session("OrgShortName")
iBookNo=Request.Form("selBook")
sPartyName=Request.Form("txtPartyName")
'sOrgName=Request.Form("horgName")
sBookName=Request.Form("hBookName")
sPartyCode=Request.Form("hPartyCode")
sSelVouTy = Request.Form("selVoucherType")
sSelInvNo = Request.Form("selInvoiceNo")

'Response.Write "<p> Vou Type = "& sSelVouTy 

IF CStr(sSelVouTy) <> "OT" Then
	sVouNarr = Request.Form("hVouDetails")
Else
	sVouNarr = ""
End IF

sUserID = getUserID()


Set objRs = Server.CreateObject("ADODB.RecordSet")
sCallFrom = Request("CallFrom")
if Trim(sCallFrom)="GJ" then
    sVouCode = "08"
    sVouName = "GJ"
else
    sVouCode = "07"
    sVouName = "CN"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=iBookNo%>" BookName="<%=sBookName%>" CRDR="C" VouDate="" PartyCode="<%=Replace(sPartyCode,"&"," and ")%>" Approver="" PartyName="<%=Replace(Trim(sPartyName),"&"," and ")%>" /></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></XML>
<XML id="AccHeadData">
<account/>
</XML>
<xml id="GLHeadData"><Root /></xml>
<xml id="GJVoucher"></xml>
<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
dim iBookAccCode

iEntryNo=1
bVouFlag=false
bSavFlag=false
set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement

iBookAccCode=0
'*****************
Function AddNew()
    if trim(document.formname.hAction.value)="Edit" then
        AddEntry("U")
    else
        AddEntry("A")
    end if
End Function
'***********************
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

FUNCTION PopAccHead(objAcc)
set VouRoot=VoucherData.documentElement

dim sOrgId,sTemp,sDesc
	sOrgId=document.formname.hOrgId.value
	if objAcc.selectedIndex >0 then
				if objAcc.value="G" then
					showGLHead sOrgId
				else
					sTemp=Split(objAcc.value,"?")
					document.formname.hTdsElgi.value = sTemp(4)
					IF CStr(sTemp(4)) = "0" Then
						document.formname.txtTdsAmount.disabled = True
						document.formname.txtTdsper.disabled = True
					Else
						document.formname.txtTdsAmount.disabled = False
						document.formname.txtTdsper.disabled = False
					End IF
					if CheckAccHead(VouRoot,trim(sTemp(0))) then
						MsgBox"Account Head already Exisit in Voucher"
						document.formname.selAccountHead.selectedIndex=0
						exit function
					end if
					sDesc=objAcc.options(objAcc.selectedIndex).text
					document.formname.txtNarration.focus()
					window.spAccHead.innerHTML=sDesc
					Set newElem = EntryData.createElement("AccHead")
						newElem.setAttribute "No", trim(sTemp(0))
						newElem.setAttribute "CostCenter", trim(sTemp(1))
						newElem.setAttribute "Analytical", trim(sTemp(2))
						newElem.setAttribute "Name", sDesc
						newElem.setAttribute "Type", "G"
						newElem.setAttribute "Group", ""
	    				EntryRoot.appendChild newElem

	    			bVouFlag=true
					showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))
				End if 'End of select Account Head Type check GL or PARTY
	End if 'End of If any Account Head Selected Check
END FUNCTION
'---------------------End Of Function PopAccHead-------------------
FUNCTION showGLHead(sOrgId)
dim iAccCode,bAnal,bCostCenter
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,arrTemp,sRetVal,sTemp2,sTdsElgi,sTempVal
iBookNo=document.formname.hBookcode.value
'alert(sOrgID)
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
        sTdsElgi = ndChild.getAttribute("RetField6")
    next
end if

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

if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		iAccCode=HeaderNode.Attributes.Item(0).nodeValue
		if CheckAccHead(VouRoot,iAccCode) then
			MsgBox"Account Head already Exisit in Voucher"
			document.formname.selAccountHead.selectedIndex=0
			exit function
		end if
		bVouFlag=true

		bAnal=HeaderNode.Attributes.Item(1).nodeValue
		bCostCenter=HeaderNode.Attributes.Item(2).nodeValue
		window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue
		EntryRoot.appendChild HeaderNode
	next
	showCCAnal sOrgId,iAccCode,bCostCenter,bAnal
else
	'User canceled Account Head Selection
	document.formname.selAccountHead.selectedIndex=0
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if 'End of GL Head Processing
set nodAccHead=nothing
END FUNCTION
'---------------------End Of Function showGLHead--------------------------

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

Function AddEntry(bFlag)
dim iCode,dRatio,dAmount

' New Validation for check blank data - included on 02/04/2004
if bFlag = "S" then
	if iEntryNo > 1 and document.formname.txtAmount.value = "0.00" then
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			Exit Function
		End IF

	end if
end if
' End of Validation

	if not bVouFlag then
		MsgBox "Select an Account Head"
		exit function
	end if
	if not checkFileds then exit function

	VouRoot.Attributes.getNamedItem("VouDate").value=document.formname.ctlDate.getdate
	IF CStr(bFlag) <> "U" Then
	    iEntryNo = iEntryNo + 1
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	End IF
	EntryRoot.Attributes.Item(1).nodeValue="D"

	EntryRoot.Attributes.Item(2).nodeValue=window.spAccHead.innerHTML
	EntryRoot.Attributes.Item(3).nodeValue=document.formname.txtAmount.value

	EntryRoot.Attributes.Item(4).nodeValue=document.formname.hOrgId.value
	EntryRoot.Attributes.Item(5).nodeValue=document.formname.hOrgName.value

	EntryRoot.Attributes.getNamedItem("TdsAmount").Value=document.formname.txtTdsAmount.value
	EntryRoot.Attributes.getNamedItem("TDSElgi").Value=document.formname.hTdsElgi.value
	EntryRoot.Attributes.getNamedItem("TdsPercentage").Value=document.formname.txtTdsper.value



	IF CStr(bFlag) <> "U" Then
		Set newElem = EntryData.createElement("Narration")
		newElem.text= document.formname.txtNarration.value
		EntryRoot.appendChild newElem
	End IF

	for each HeaderNode in EntryRoot.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
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
		if 	HeaderNode.nodeName="PayRec" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				dAmount=eval("document.formname.txtDocAmount"&iCode).value
				nodANL.Attributes.Item(5).nodeValue=dAmount
			next
		end if 'End of Check for Analytical Node

		IF HeaderNode.nodeName = "Narration" Then
			HeaderNode.text = document.formname.txtNarration.value
		End IF

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
		VouRoot.appendChild EntryRoot
	End IF
'====================================================================================
	'alert VouRoot.xml

	if bFlag="A" then
		DisplayVoucher
		'iEntryNo=iEntryNo+1
		clearXML()
		document.formname.txtTdsAmount.value = "0.00"
		document.formname.txtTdsper.value = "0.00"
		setADDDisplay 0

		document.formname.selAccountHead.selectedIndex=0
		window.spAccHead.innerHTML=""
		window.spEntryNo.innerHTML=iEntryNo
		document.formname.reset
	elseif bFlag = "U" Then
		DisplayVoucher
		clearXML
		document.formname.txtTdsAmount.value = "0.00"
		document.formname.txtTdsper.value = "0.00"
		document.formname.txtAmount.value = "0.00"
		document.formname.txtNarration.value = ""
		document.formname.selAccountHead.selectedIndex = 0
		setADDDisplay 0

		document.formname.selAccountHead.selectedIndex=0
		window.spAccHead.innerHTML=""
		'document.formname.btnUpdate.disabled = True
		'document.formname.btnDel.disabled = True
		'document.formname.btnAdd.disabled = False
		'document.formname.btnNext.disabled = False
	else
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			Exit Function
		End IF
	end if
	'MsgBox EntryRoot.xml
END FUNCTION
'---------------------END OF FUNCTION ADDENTRY----------------------------
FUNCTION DisplayVoucher()
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTdsAmt,sTdsPer
dim dTotal,sAccUnit,sTotalCRDR,iRow,iDivFixed,iDivHeight

iDivFixed = 60
iDivHeight = cint(iEntryNo) * cint(25)
window.DisVoucher.style.height= cint(iDivFixed)+CInt(iDivHeight) & "px"
window.DisVoucher.style.visibility="visible"
ClearTable "tblVoucher",1,1
dTotal=0
iRow = 1

For Each EntryNode in VouRoot.childNodes
	iSno=EntryNode.Attributes.Item(0).nodeValue
	sAmount=EntryNode.Attributes.Item(3).nodeValue
	sTdsAmt = EntryNode.Attributes.Item(6).nodeValue
	sTdsPer = EntryNode.Attributes.Item(8).nodeValue
	sAmount=FormatNumber(CDbl(sAmount),2,,,0)
	sTdsAmt=FormatNumber(CDbl(sTdsAmt),2,,,0)
	sTdsPer=FormatNumber(CDbl(sTdsPer),2,,,0)

	dTotal=dTotal+CDbl(sAmount)

	sAccUnit=EntryNode.Attributes.Item(5).nodeValue
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
	InsertCell oRow,1,"","<img src='../../assets/images/iTMS%20icons/Deleteicon.gif' onClick=EditEntry('"&iSno&"','D')>","ExcelDisplayCell","Center","top",0,0,0,0,""
	InsertCell oRow,1,"","<a class=""ExcelDisplayLink"" href=""javascript:EditEntry('"&iSno&"','E')"" Class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
	'InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sTdsAmt,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sTdsPer,"ExcelDisplayCell","left","top",0,0,0,0,""
	iRow = iRow + 1
next'End of Voucher Node Loop
	dTotal=FormatNumber(dTotal,2,,,0)

	set oRow = document.all.tblVoucher.insertRow()
	InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,5,0,""
	InsertCell oRow,1,"",CStr(dTotal),"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
END FUNCTION
'---------------------END OF FUNCTION DISPLAYVOUCHER----------------------

'---------------------End Of Function AddEntry--------------------------
Function SaveXML()
if trim(document.formname.hCallFrom.value)="CR" then
	IF CheckApp() Then
		'IF CheckNoSer() Then
			set objhttp = CreateObject("Microsoft.XMLHTTP")
			objhttp.Open "POST","XMLSave.asp?Mod=CN&Name=Voucher Entry", false
			objhttp.send VoucherData.XMLDocument
			if objhttp.responseText <> "" then
				Msgbox(objhttp.responseText)
			else
				document.formname.btnNext.disabled = True
				document.formname.btnAdd.disabled = True
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
						oDGjEntry.setAttribute "CRDR","C"
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
					oDGjEntry.setAttribute "CRDR","D"
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
end if'if trim(document.formname.hCallFrom.value)="CR" then
	
	
	
	
End Function
'---------------------End Of Function SaveXML--------------------------
Function  checkFileds()
	if trim(document.formname.txtNarration.value)="" then
		Msgbox("Enter Narration")
		document.formname.txtNarration.select
		checkFileds=false
		exit Function
	end if
	if ValidateAmount(document.formname.txtAmount.value)=false then
		document.formname.txtAmount.select
		checkFileds=false
		exit Function
	end if
	checkFileds=true
end Function
'---------------------End Of Function checkFileds--------------------------
Function CancelAction(sPage)
	document.formname.action=sPage
	document.formname.submit
end Function
'---------------------End Of Function ActionCancel----------------------------

Function CheckVoucStat()
	Dim sCurrDate
	sCurrDate = document.formname.hCurrDate.Value
	IF DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		CheckVouStat = false
		Exit Function
	Else
		CheckVouStat = True
	End IF
End Function

Function EditEntry(iVouNo,iEditType)
	Dim iCounter,sExp,CheckNode
sExp = "//Entry[@TdsAmount]"
Set CheckNode = VouRoot.selectNodes(sExp)
'document.formname.hEditEntNo.value = iVouNo
'if bEditFlag then
if iEditType ="E" then
    document.formname.hAction.value = "Edit"
	    document.formname.hEntryNo.value = iVouNo
	    window.spEntryNo.innerHTML = iVouNo
	    setADDDisplay 0
	    'setPayableDisplay 0
	    bVouFlag=true
	    sAccUnit = VouRoot.Attributes.Item(0).nodeValue
	    For Each EntryNode in VouRoot.childNodes
		    if EntryNode.Attributes.Item(0).nodeValue=iVouNo then
			    document.formname.txtAmount.value=EntryNode.Attributes.Item(3).nodeValue
			    IF CheckNode.length <> 0 Then
				    document.formname.txtTdsAmount.value = EntryNode.Attributes.Item(6).nodeValue
				    document.formname.txtTdsper.value = EntryNode.Attributes.Item(8).nodeValue

				    IF CStr(EntryNode.Attributes.Item(7).nodeValue) = "1" Then
					    document.formname.txtTdsAmount.disabled = False
					    document.formname.txtTdsper.disabled = False
				    Else
					    document.formname.txtTdsAmount.disabled = True
					    document.formname.txtTdsper.disabled = True
				    End IF
				    document.formname.hTDSElgi.value = EntryNode.Attributes.Item(7).nodeValue
			    Else
				    document.formname.txtTdsAmount.value = "0.00"
				    document.formname.txtTdsper.value = "0.00"
				    document.formname.hTDSElgi.value = "0"
			    End IF
			    'sAccUnit=EntryNode.Attributes.Item(5).nodeValue
			    sAddtional=""
			    For Each HeaderNode in EntryNode.childNodes
				    if HeaderNode.nodeName="AccHead" then
					    For iCounter = 0 To document.formname.selAccountHead.length - 1
						    IF CStr(HeaderNode.Attributes.getNamedItem("Type").value) = CStr(document.formname.selAccountHead(iCounter).value) Then
							    document.formname.selAccountHead.selectedIndex = iCounter
							    Exit For
						    End IF
					    Next

					    'document.formname.txtPayTo.value=HeaderNode.Attributes.Item(3).nodeValue
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
			    'MsgBox VouRoot.xml
		    end if
	    next'End of Voucher Node Loop
    '	document.formname.btnadd.disabled=true
    '	document.formname.btnnext.disabled=true
    '	document.formname.btnupdate.disabled=false
    '	document.formname.btndel.disabled=false
	    bEditFlag=false
	    bSavFlag=true
    'end if
Else ' if iEditType ="E" then
    document.formname.hEntryNo.value = iVouNo

	    setADDDisplay 0
	    bVouFlag=true
	    sAccUnit = VouRoot.Attributes.Item(0).nodeValue
	    For Each EntryNode in VouRoot.childNodes
		    if EntryNode.Attributes.Item(0).nodeValue=iVouNo then
			    set EntryRoot=VouRoot.removeChild(EntryNode)
			    'MsgBox VouRoot.xml
		    end if
	    next'End of Voucher Node Loop
        bEditFlag=false
	    bSavFlag=true
	    DelEntry
End If 'if iEditType ="E" then
End Function

Function DelEntry()
	clearXML()
	setADDDisplay 0
	'setPayableDisplay 0
	DisplayVoucher


	window.spEntryNo.innerHTML=iEntryNo

	'document.formname.selCRDR(0).disabled=false
	'document.formname.selCRDR(1).disabled=false

	document.formname.reset

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
	sFromDate = "01/04/"&Trim(sFromYr)
	sTodate = "31/03/"&sToYr
	if DateDiff("d",sTodate,date)>0 then
	    document.formname.ctlDate.setMinDate = sFromDate
	    document.formname.ctlDate.setMaxDate = sToDate
	    document.formname.ctlDate.setDate = sToDate
	else
	    document.formname.ctlDate.setMinDate = sFromDate
	    document.formname.ctlDate.setMaxDate = date
	    document.formname.ctlDate.setDate =date
	end if 
	
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
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hSelVouTy" value="<%=sSelVouTy%>">
<input type="hidden" name="hInvNos" value="<%=sSelInvNo%>">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hCallFrm" value="C">
<input type="hidden" name="hVouCRDR" value="">
<input type="hidden" name="hAction" value="New">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20">Other Credit Note 		</td>
    </tr>
	<tr>
		<td align="center" class=MiddlePack height="20"><p align="center"> 		</td>
    </tr>
    <tr>
								<td align="center">
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
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
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
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
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
                             <table cellpadding="0" cellspacing="0" width=100% border=0>
                                <tr>
                                    <td colspan=4 width=100%>
                                        <table border=0 cellpadding=0 cellspacing=0>
                                            <tr>
                                                <td class="FieldCellSub" width="15%">Party Name </td>
                                                <td class="FieldCell" width="55%" ><span class="DataOnly"><%=sPartyName%></span>
                                                </td>
                                                <td  class="FieldCell">
                                                    Voucher Date
                                                </td>
                                                <td align=right class="FieldCell">
                                                 <% ' Function Call to Insert Date Picker
								                    Response.Write InsertDatePicker("ctlDate")
							                    %>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan=4>
                                    <table border=0 cellspacing=0 cellpadding=0 class="TableOutlineOnly" width=100%>
                                            <tr>
                                        <td class="FieldCellSub" width="160">GL Account Head</td>
                                        <td class="FieldCell">
                                        <select size="1" name="selAccountHead" class="FormElem" onChange="PopAccHead(this) ">
							            <option value="S">Select Account Head</option>
							            <%
										            dim iHeadCount
										            iHeadCount=popFrequentHead(sOrgId,"07",iBookNo)

							            %>
							            <option value="G">GL Account Head</option>
                                        </select>
                                        </td>
                                        <td class="FieldCellSub" width="150" colspan=2 align=right>Entry No&nbsp;&nbsp;
                                        <span class="DataOnly" id="spEntryNo">1</span></td>
                                            </tr>
                                            <tr>
                                        <td class="FieldCellSub" width="125"></td>
                                        <td class="FieldCell" colspan="3"> <span class="DataOnly" id="spAccHead"></span></td>
                                            </tr>
                                            <tr>
                                                                <td width="139" valign="top">
                                                                 <table border="0" width="100%" cellspacing="1">
                                                                    <tr>
                                                                      <td width="50%" class="FieldCellSub">Narration</td>
                                                                      <td width="50%" class="FieldCellSub">
            <%

            sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
	            " OUDefinitionID='"&sOrgId&"'and BookCode='07' and BookNumber="&iBookNo

            with objRs
	            .CursorLocation = 3
	            .CursorType = 3
	            .Source = sQuery
	            .ActiveConnection = con
	            .Open
            end with
            set objRs.ActiveConnection = nothing

            if objRs(0)>0 then
            %>
                                                                        <p align="left">
                                                                <a href="javascript:showNarration('07')"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
            <%
            end if
            objRs.Close
            %>
                                                                       </td>
                                                                    </tr>
                                                                  </table>

                                                                </td>
                                                                <td class="FieldCell" colspan="3" valign="top"> <textarea rows="3" name="txtNarration" cols="50" class="FormElem"><%=Trim(sVouNarr)%></textarea> </td>

                                            </tr>
                                            <tr>
                                        <td class="FieldCellSub" width="115">Amount</td>
                                        <td class="FieldCell" colspan="3">
                                        <input type="text" name="txtAmount" value="0.00" size="15" maxlength="15" style="text-align:right" class="FormElem" onblur="popAddAmount()"> </td>
                                            </tr>
                                            <tr>
                                            <td class="FieldCellSub" width="133">Deduction @</td>
                                            <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                            % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                            </td>
                                                </tr>

                                            <tr>
                                                                <td class="FieldCellSub" width="139">Approval</td>
                                                                <td class="FieldCell" colspan="3">
                                                                <input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="SetApp('Y')">Yes&nbsp;&nbsp;

													            <input type="radio" value="N" name="optApprove" class="FormElem" onClick="SetApp('N')"> No
													            &nbsp;&nbsp; Immediate Approver &nbsp;
													            <select size="1" name="selUserId" class="FormElem">
														            <option value="I">Immediate Approver</option>
														            <%=populateEmployeeWithVal(sUserId)%>
													            </select>
													            &nbsp;<input type="button" value="Add Entry" name="btnAdd" class="AddButton" onclick="AddNew()" >
													            </td>
                                                            </tr>
                                                            <tr>
							                                    <td align=center width=100% colspan=4>
                                                                        <DIV class=frmBody id="DisVoucher" style="width:98%; visibility:hidden; height:1;">
                                                                            <table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="98%">
                                                                            <tr>
	                                                                            <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
	                                                                            <!--<td class="ExcelHeaderCell" align="center" width="75">AU</td>-->
	                                                                            <td class="ExcelHeaderCell" align="center">Account Code - Name</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="125">Narration</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="80">Amount</td>
	                                                                            <td class="ExcelHeaderCell" align="center">CC/AH Details</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="80">Amount</td>
	                                                                            <td class="ExcelHeaderCell" align="center" width="80">Amount</td>
                                                                            </tr>
                                                                            </table>
                                                                        </div>
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
								<td valign="top" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <!--<input type="button" value="Update" name="btnUpdate" class="ActionButton" onclick="AddEntry('U')" disabled>
                                                                <input type="button" value="Delete" name="btnDel" class="ActionButton" onclick="DelEntry()" disabled>-->
                                                                <input type="button" value="Save" onClick="AddEntry('S')" name="btnNext" class="ActionButton" >
                                                                <input type="button" value="Cancel" onClick="CancelAction('CREDITNOTETOCREATE.asp')" name="B8" class="ActionButton" >
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