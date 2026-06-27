<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AMDBankVoucher.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	June 10, 2005
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
dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,sQuery
dim iVouNo,objRs,objRs1,sVouDate,bActionFlag
dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sAccHead,sParType,sPartSubType
dim iEnNo,Entrynode,HeaderNode,dOpeningBal
dim sParCode,sNarration,sAccHeadname,sAccUnitName,bOtherUnits,iBookAccHead,dTransLimit
Dim sVouCkTy,sLastVouDt,sCallVouTy,sPopVouTy,sFinPeriod,sFinTemp,sMaxDate,sMinDate
Dim iPreBookVal,objDOM
Dim sFinFrm,sFinTo,sValTemp2,sFormVal,sSelArg
Dim sConStatus,sRetVal

'Response.Write Request.ServerVariables("SCRIPTNAME")

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sFinTemp = Split(sFinPeriod,":")
	sMaxDate = "31/03/"&sFinTemp(1)
	sMinDate = "01/04/"&sFinTemp(0)
End IF

dim sAccount,sAddtional,iSno
dim dTotal
dim sVoucDate,iBookCode,sPayTo,sUserId,sVal,sValTemp
sUserId = getUserID()
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2
sCallVouTy = Request("VouTy")
sSelArg = Request("voutype")
sFormVal = Request("hFormVal")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("horgName")
sOrgId   = session("organizationcode")
sOrgName = session("orgshortname")

sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("selVouTypepe")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")

iPreBookVal = sBookCode
sConStatus = GetContraStatus(sTransNo)
'Response.Write "sTransNo="&sTransNo
sVal=Request("Val")
sValTemp=Split(sVal,"~")

IF Cstr(sVal) <> "" Then
	sQuery = "Select H.OUDEFINITIONID,D.OrgUnitDescription From DCS_OrganizationUnitDefinitions D, "&_
		 "Acc_T_CreatedVoucherHeader H Where H.OUDEFINITIONID = D.OUDEFINITIONID "&_
		 "and H.CreatedTransNo = "&sValTemp(0)&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sOrgId = objRs(0)
		sOrgName = objRs(1)
	End IF
	objRs.Close

Else

	sQuery = "Select Top 1 OUDefinitionID,OrgUnitDescription From DCS_OrganizationUnitDefinitions "&_
			 "Where Len(OUDefinitionID) > 4 Order By OUDefinitionID "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sOrgId = objRs(0)
		sOrgName = objRs(1)
	End IF
	objRs.Close
End IF

IF CStr(iBookAccHead) = "" Then
	sQuery = "Select Top 1 BookNumber,BookName,isNull(BookAccountHead,0),OtherUnitTransaction From vwOrgBookNames Where  "&_
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '02' Order By BookName "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sBookCode = objRs(0)
		sBookName = objRs(1)
		iBookAccHead = objRs(2)
		bOtherUnits = objRs(3)
	Else
		sBookCode = "02"
		sBookName = ""
		iBookAccHead = 0
		bOtherUnits = 1
	End IF
	objRs.Close
End IF

'Response.Write sTransNo


oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

'Response.Write sTransNo

'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_BA_"&Session.SessionID&".xml")

IF CStr(sTransNo) <> "" Then
 	sRetVal = GetVouchXML(sTransNo)
	oDOM.Load server.MapPath(sRetVal)
	oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_BA_"&Session.SessionID&".xml")
End IF


IF CStr(sTransNo) = "" Then
	sTransNo = 0
End IF

IF CStr(sCallVouTy) = "R" Then
	sVouType = "D"
	sVouCkTy = "BAP"
	sPopVouTy = "D"
Else
	sVouType = "C"
	sVouCkTy = "BAR"
	sPopVouTy = "C"
End IF

'IF CStr(sVouType) = "C" Then
'	sVouCkTy = "BAP"
'Else
'	sVouCkTy = "BAR"
'End IF

sQuery = "Select Convert(Char,VoucherDate,103) From Acc_T_CreatedVoucherHeader Where CreatedTransNo =  "&_
		 "(Select Max(CreatedTransNo) From Acc_T_CreatedVoucherHeader Where BookCode = '02'  "&_
		 "and OUDefinitionID = '"&sOrgId&"' and TransactionType ='"&sVouCkTy&"' ) "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sLastVouDt = Trim(objRs(0))
End IF
objRs.Close

'sQuery = "Select CreatedBy From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sTransNo&" "
'objRs.Open sQuery,Con
'IF Not objRs.EOF Then
'	sUserId = Trim(objRs(0))
'End IF
'objRs.Close


'IF CStr(sUserId) = "" Then
sUserId = session("userid")
'End IF

sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp2(0))
sFinTo = Trim(sValTemp2(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"

Dim TempNode
Set Root = objDOM.CreateElement("Root")
objDOM.appendChild Root
Set Entrynode = objDOM.CreateElement("Details")
Root.appendChild Entrynode
'****************** This Voucher that is been Adjusted in Any Purchase/Sales Vouchers Will be
'****************** taken here for consideration *********************************************

sQuery = "Select H.CreatedVoucherNo,Convert(Varchar,H.VoucherDate,103) VoucherDate,H.VoucherAmount,C.AmountAdjusted, "&_
		 "A.CreatedAdvanceNo,C.CreatedTransNo,H.CrDrIndication From Acc_T_CreatedVoucherHeader H, "&_
		 "ACC_T_CreatedAdvanceAdj C,Acc_T_CreatedAdvances A Where A.CreatedTransNo = "&sTransNo&" "&_
		 "And A.CreatedAdvanceNo = C.CreatedAdvanceNo and C.CreatedTransNo = H.CreatedTransNo  "

'Response.Write sQuery
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountAdjusted")
	TempNode.setAttribute "AdjCrTrNo", objRs("CreatedTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CreatedAdvanceNo")
	TempNode.setAttribute "ValFrm", "CRADJADV"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close

sQuery = "Select CreatedVoucherNo,Convert(Varchar,VoucherDate,103) VoucherDate,VoucherAmount, "&_
		 "AmountPaid,AdjCrTransNo,CradvNo From CrPayAdvAdjDet Where CreatedTransNo = "&sTransNo
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountPaid")
	TempNode.setAttribute "AdjCrTrNo", objRs("AdjCrTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CradvNo")
	TempNode.setAttribute "ValFrm", "Pay"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close

sQuery = "Select CreatedVoucherNo,Convert(Varchar,VoucherDate,103) VoucherDate,VoucherAmount, "&_
		 "AmountReceived,AdjCrTransNo,CradvNo From CrRecAdvAdjDet Where CreatedTransNo = "&sTransNo
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Set objRs.ActiveConnection = Nothing

Do While Not objRs.EOF
	Set TempNode = objDOM.CreateElement("AdjDet")
	TempNode.setAttribute "AmdCrTrNo", sTransNo
	TempNode.setAttribute "AdjCrVouNo", objRs("CreatedVoucherNo")
	TempNode.setAttribute "AdjCrVouDate", objRs("VoucherDate")
	TempNode.setAttribute "AdjCrVouAmt", objRs("VoucherAmount")
	TempNode.setAttribute "AdjCrVouAdjAmt", objRs("AmountReceived")
	TempNode.setAttribute "AdjCrTrNo", objRs("AdjCrTransNo")
	TempNode.setAttribute "AmdCrAdvNo", objRs("CradvNo")
	TempNode.setAttribute "ValFrm", "Rec"
	Entrynode.appendChild TempNode
	objRs.MoveNext
loop
objRs.Close


objDOM.Save server.MapPath("../temp/transaction/Voucher ADJAMD_BA_"&Session.SessionID&".xml")



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS Bank Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<SCRIPT language="javascript" SRC="../scripts/VouSelection.js"></SCRIPT>
<script language="javascript" src="../scripts/VoucherEntryCore.js"></script>
<script language="javascript" src="../scripts/BankVoucher.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>

<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=sBookCode%>" BookName="<%=sBookName%>" CRDR="<%=sVouType%>" VouDate="" BookAcchead="0" Approver=""/></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="0" TDSElgi="0" TdsPercentage="0" PayRecAmount="0" /></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="TDSData"  ><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<XML ID="UnitBookData">
<Book/>
</XML>
<XML ID="TDSFlagData">
<Root/>
</XML>
<XML id="VoucherAmdData"></XML>
<XML id="IUTUnits"></XML>
<XML id="InsDet"></XML>
<xml id="GLHeadData"><Root /></xml>
<xml id="PartyHeadData"><Root /></xml>
<XML id="AdjDet" src="<%="../Temp/Transaction/Voucher ADJAMD_BA_"&Session.SessionID&".xml"%>"></XML>

<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot
dim bVouFlag,bSavFlag,bEditFlag,iBookAcchead,dTransLimit,sTransFlag,sAdjType
iEntryNo=0
bVouFlag=false
bSavFlag=false
bEditFlag=true

iBookAcchead=<%=iBookAccHead%>
dTransLimit=<%=dTransLimit%>
sTransFlag="A"

set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement

Function DisplayBook()
	Dim iUnitNo,arrTemp,sOrgName
	Dim Root,sVal	
	document.formname.selBook.options.length = 0
	'document.formname.selAccUnitId.selectedIndex = objUnit.selectedIndex
	
	
'if objUnit.selectedIndex <> "0" then
	'iUnitNo = objUnit(objUnit.selectedIndex).value
	'sOrgName = objUnit(objUnit.selectedIndex).Text
	
	iUnitNo = document.formname.hOrgId.value
	sOrgName = document.formname.hOrgName.value

	'DisplayIUTUnits iUnitNo,sOrgName

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=02&orgID=" & iUnitNo , false
	objhttp.send
	'alert objhttp.responseText

	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			sVal = Trim(HeaderNode.Attributes.Item(0).nodeValue)
			sVal = CStr(sVal)&"-"&Trim(HeaderNode.Attributes.Item(2).nodeValue)

			document.formname.selBook.length = document.formname.selBook.length+1
			document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
			'document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			document.formname.selBook.options(document.formname.selBook.length-1).Value = sVal

		next
	end if
'end if
	popAccHead
	IF document.formname.selBook.length > 1 Then
		document.formname.selBook.selectedIndex = 1
	Else
		document.formname.selBook.selectedIndex = 0
	End IF

end Function


Function SaveXML()
Dim sExp,TempNode
	if bSavFlag then
		'IF CheckVoucherDt() Then ' Checking for the VoucherDate between Finicial Years Only
			IF CheckApp() Then ' Checking For Selected/Entered Values
				IF CheckContraEnt() Then ' Checking Wheather No Series is Defined or Not
					'-----------Added by S.Maheswari on Fev 20th 2009 to update hNorToContra value(copied from Sangeeth ---------
						'alert document.formname.hCurrConStatus.value & "*****"& document.formname.ContraStatus.Value
					IF CStr(document.formname.hCurrConStatus.value) = "1" and CStr(document.formname.ContraStatus.Value) = "F" Then
						'MsgBox "Non Contra Voucher Cannot be Converted to Contra Voucher "
						'Exit Function
						document.formname.hNorToContra.value = "Y"
					End IF
					'---------------------------------------------------------------------------------------
					'IF document.formname.selBook.selectedIndex = 0 Then
					'	MsgBox "Select Book"
					'	Exit Function
					'End IF

					IF CheckFinDate Then
						UpdateXML()

						''''
						'added on Nov 19,2009
						nFinalAmt =CheckVouAmount()
						'blocked on July 24,2010
						'if cdbl(nFinalAmt) <> cdbl(document.formname.hInsrAmt.value) then
						if trim(nFinalAmt) <> trim(document.formname.hInsrAmt.value) then
							MsgBox "Voucher Amount and Instrument Amount should be same"
							Exit Function
						end if
						''''

						set objhttp = CreateObject("Microsoft.XMLHTTP")
						objhttp.Open "POST","XMLSave.asp?Name=Voucher Entry&Mod=BA", false
						document.formname.action = "AmdAccGenerate.asp"

						IF Not CheckAdjAmtDate() Then
							Exit Function
						End IF

						'alert VoucherData.XML
						'Exit function

						objhttp.send VoucherData.XMLDocument
						if objhttp.responseText <> "" then
							Msgbox(objhttp.responseText)
						else
							document.formname.btnNext.disabled = True
							IF CStr(document.formname.hAmendTy.value) = "Y" Then
								document.formname.action = "AmdAccGenerate.asp"
							End IF
							document.formname.btnNext.disabled = True
							document.formname.btnDel.disabled = True

							document.formname.submit()
						end if
					End IF
				End IF
			End IF
		'End IF
	end if
End Function

'Function AddEntry(var)
'	Dim val,sParaXml,iTransNo
'		Val = Var
'		If trim(Val) = "S" then
'			sParaXml = 	VoucherData.XML
'			'alert sParaXml
'			'iTransNo = document.formname.hTransNo.value
'			'alert(iTransNo)
'			document.formname.action = "AmdAccGenerate.asp"
'			document.formname.submit()
'
'		End If
'End Function

Function CheckPayRecAmt()
	Dim sExp,TempNode,iCtr,dAddAmt,dSubAmt,PayNode,sParTy,sTemp,sDispType
	Dim sType,dEntAMt,dTempAmt
	dAddAmt = 0
	dSubAmt = 0
	dEntAMt = 0
	dTempAmt = 0

	sType = "T"

	IF CStr(document.formname.hVouCRDR.Value) = "D" Then
		sDispType = "Receipt Amount Should be Greater Than Payment Amount "
	Else
		sDispType = "Payment Amount Should be Greater Than Receipt Amount "
	End IF

	For Each TempNode in EntryRoot.ChildNodes
		IF TempNode.nodeName = "AccHead" Then
			IF Cstr(TempNode.Attributes.Item(4).Value) = "P" Then
				sParTy = Left(TempNode.Attributes.Item(0).Value,2)
			Else
				CheckPayRecAmt = True
				Exit Function
			End IF
		End IF

		IF TempNode.nodeName = "PayRec" Then
			IF CStr(sParTy) = "CR" Then
				For Each PayNode in TempNode.childNodes
					Select Case CStr(PayNode.Attributes.getNamedItem("AdjType").Value)
						Case "PI"
							dAddAmt = CDbl(dAddAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "C"
							dAddAmt = CDbl(dAddAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "I"
							dSubAmt = CDbl(dSubAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "D"
							dSubAmt = CDbl(dSubAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "P"
							dSubAmt = CDbl(dSubAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
					End Select
				Next
			Else
				For Each PayNode in TempNode.childNodes
					Select Case CStr(PayNode.Attributes.getNamedItem("AdjType").Value)
						Case "I"
							dAddAmt = CDbl(dAddAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "D"
							dAddAmt = CDbl(dAddAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "PI"
							dSubAmt = CDbl(dSubAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "C"
							dSubAmt = CDbl(dSubAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
						Case "R"
							dSubAmt = CDbl(dSubAmt) + CDbl(PayNode.Attributes.getNamedItem("AmtToAdjust").Value)
					End Select
				Next
			End IF
		End IF
	Next

	'MsgBox CDbl(dAddAmt) - CDbl(dSubAmt)
	EntryRoot.Attributes.Item(9).nodeValue = CDbl(dAddAmt) - CDbl(dSubAmt)
	'MsgBox EntryRoot.Attributes.Item(9).nodeValue

	'MsgBox dAddAmt &" " & dSubAmt



	IF CDbl(dAddAmt) = 0 and CDbl(dSubAmt) = 0 Then
		'CheckPayRecAmt = True
		'Exit Function
		sType = "T"
	Elseif CDbl(dAddAmt) < CDbl(dSubAmt) Then
		'MsgBox sDispType
		'CheckPayRecAmt = False
		'Exit Function
		sType = "F"
	Else
		'CheckPayRecAmt = True
		'Exit Function
		sType = "T"
	End IF

	dAddAmt = Trim(dAddAmt)
	dSubAmt = Trim(dSubAmt)

	dAddAmt = Cdbl(dAddAmt)
	dSubAmt = Cdbl(dSubAmt)

	dAddAmt = FormatNumber(dAddAmt,2,,,0)
	dSubAmt = FormatNumber(dSubAmt,2,,,0)

	dTempAmt = Cdbl(dAddAmt - dSubAmt)
	dTempAmt = FormatNumber(dTempAmt,2,,,0)

	dEntAMt = Trim(document.formname.txtAmount.value)
	dEntAMt = CDbl(dEntAMt)
	dEntAMt = FormatNumber(dEntAMt,2,,,0)


	Dim sRetVal
	IF CStr(sType) = "T" Then
			dEntAMt = Trim(dEntAMt)
			dTempAmt = Trim(dTempAmt)

			dEntAMt = CDbl(dEntAMt)
			dTempAmt = CDbl(dTempAmt)

		'IF CStr(sParTy) = "CR" Then
			IF dEntAMt = dTempAmt Then
				CheckPayRecAmt = True
				Exit Function
			Elseif dEntAMt > dTempAmt Then
				sRetVal = MsgBox("Entered Value is Greater than Adjusted Value! Remaing amount will treat as Advance Continue!! ",4,"Bank Voucher")
				IF sRetVal = 6 Then
					CheckPayRecAmt = True
					Exit Function
				End IF
			Else
				MsgBox "Entry Amount Should be less than or equal to Adjusted Amount"
			End IF
		'Else

		'End IF
	Else
		IF dEntAMt >= dTempAmt Then
			CheckPayRecAmt = True
			Exit Function
		Else
			MsgBox sDispType
			CheckPayRecAmt = False
		End IF
	End IF

	'CheckPayRecAmt = True
End Function

Function CheckContraEnt()
	Dim sExp,ContCntNode,ContNode,iBookAccHead,iEntAccHead,iCtr,objhttp
	Dim iUnitNo,sAccTemp
	iUnitNo = document.formname.hOrgId.value
	sAccTemp = Split(document.formname.selBook.value,"-")
	iBookAcchead = sAccTemp(1)
	sExp = "//Entry"
	Set ContCntNode = VouRoot.selectNodes(sExp)
	IF ContCntNode.length > 1 Then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		sExp = "//AccHead[@Type=""G""]"
		Set ContNode = VouRoot.selectNodes(sExp)
		For iCtr = 0 To ContNode.length - 1
			iEntAccHead = ContNode.Item(iCtr).Attributes.getNamedItem("No").Value
			objhttp.Open "GET","XMLContraEntAccChk.asp?BkAccHd="&iBookAcchead&"&orgID=" & iUnitNo&"&AccHead="&iEntAccHead , false
			objhttp.send
			IF Cstr(Trim(objhttp.responseText)) <> "0" Then
				document.formname.hCurrConStatus.value = "1" 'Added by Maheswar on Feb 20th 2009 (copied from sangeeth live)
				MsgBox "Contra Entry is Created only One Entry is allowed "
				CheckContraEnt = False

				Exit Function
			End IF
		Next
	Else 'Added by Maheswar on  Feb 20th 2009 (copied from sangeeth live)
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		sExp = "//AccHead[@Type=""G""]"
		Set ContNode = VouRoot.selectNodes(sExp)
		For iCtr = 0 To ContNode.length - 1
			iEntAccHead = ContNode.Item(iCtr).Attributes.getNamedItem("No").Value
			objhttp.Open "GET","XMLContraEntAccChk.asp?BkAccHd="&iBookAcchead&"&orgID=" & iUnitNo&"&AccHead="&iEntAccHead , false
			objhttp.send
		'	MsgBox "2="&objhttp.responseText
			IF Cstr(Trim(objhttp.responseText)) <> "0" Then
				document.formname.hCurrConStatus.value = "1"
			Else
				document.formname.hCurrConStatus.value = "0"
			End IF

		Next

	End IF

	CheckContraEnt = True
End Function

'=================================================================================================
FUNCTION popAdjAmount()
Dim iPayNo,iSno,sGroupCode,dTempAmt,dPrnAmt
dTempAmt = 0
if not checkFileds then
	document.formname.txtAmount.value=""
	exit function
end if

alert EntryRoot.xml

for each HeaderNode in EntryRoot.childNodes
	if HeaderNode.nodeName="PayRec" then
		dAmount=CDbl(document.formname.txtAmount.value)
		dTotal=dAmount
		dTempAmt = dAmount
		iCounter=1
		iSno = 1
		iChildCount=HeaderNode.childNodes.length
		if Cint(iChildCount)> 0 then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dTransAmount=nodANL.Attributes.getNamedItem("TransAmount").Value
				dAmtAdjusted=nodANL.Attributes.getNamedItem("AmtAdjusted").Value
				dAmtToAccount=nodANL.Attributes.getNamedItem("AmtToAccount").Value
				sAdjTy = Trim(nodANL.Attributes.getNamedItem("AdjType").Value)
				iPayNo = Trim(nodANL.Attributes.getNamedItem("PayableNo").Value)
				dAmtAdjust = CDbl(dTransAmount)- CDbl(dAmtAdjusted) - CDbl(dAmtToAccount)

				'MsgBox sAdjType
				Select Case CStr(sAdjTy)
					Case "PI"
						IF CDbl(dTempAmt) >= CDbl(dAmtAdjust) Then
							dTempAmt = Cdbl(dTempAmt) - Cdbl(dAmtAdjust)
							dPrnAmt = dAmtAdjust
						Else
							dPrnAmt = dTempAmt
							dTempAmt = 0

						End IF
					Case "D"
						dPrnAmt = dAmtAdjust
				End Select



				eval("document.formname.txtDocAmount"&iCode&"Z"&iPayNo&"Z"&iSno).value=FormatNumber(dPrnAmt,2,,,0)
				iSno = iSno + 1
			next
		end if 'End of Check for PayRec Child Count
	end if 'End of Check for PayRec Node
next
END FUNCTION
'=================================================================================================

Function CheckAdjAmtDate()
	Dim dDrAmt,dCrAmt,sExp,TempNode,iCtr,dVouAmt
	Dim AdjRoot,dTotAdjAMt,sMsgRetVal,sVouDate,sBillDate
	dDrAmt = 0
	dCrAmt = 0

	sVouDate = document.formname.ctlDate.GetDate()
	sExp = "//Entry"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			IF Cstr(TempNode.Item(iCtr).Attributes.getNamedItem("CRDR").Value = "D") Then
				dDrAmt = CDbl(dDrAmt) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value)
			Else
				dCrAmt = CDbl(dCrAmt) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("Amount").Value)
			End IF
		Next
		dDrAmt = CDbl(dDrAmt)
		dCrAmt = CDbl(dCrAmt)


		dVouAmt = Abs(dDrAmt-dCrAmt)
	End IF
	dTotAdjAMt = 0
	Set AdjRoot = AdjDet.documentElement
	sExp = "//AdjDet"
	Set TempNode = AdjRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			dTotAdjAMt = CDbl(dTotAdjAMt) + CDbl(TempNode.Item(iCtr).Attributes.getNamedItem("AdjCrVouAdjAmt").Value)
		Next
	End IF

	IF CDbl(dTotAdjAMt) > CDbl(dVouAmt) Then
		sMsgRetVal = MsgBox ("Bills Adjusted for this Voucher is More than the Voucher Amount! Adjusted Amount is "&FormatNumber(dTotAdjAMt,2,,,0)&" Click Yes To View Bill Details ",4)
		IF CStr(sMsgRetVal) = "6" Then
			DispBillDetails()
		End IF
		CheckAdjAmtDate = False
		Exit Function
	End IF

	sExp = "//AdjDet"
	Set TempNode = AdjRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			sBillDate = TempNode.Item(iCtr).Attributes.getNamedItem("AdjCrVouDate").Value
			IF DateDiff("d",sVouDate,sBillDate) < 0 Then
				sMsgRetVal = MsgBox("Voucher Date Should be Greater than Adjsuted Bills Date. Click Yes to View Bills Details ",4,"Bank Voucher")
				IF CStr(sMsgRetVal) = "6" Then
					DispBillDetails()
				End IF
				CheckAdjAmtDate = False
				Exit Function
			End IF
		Next
	End IF

	CheckAdjAmtDate = True
End Function

Function DispBillDetails()
	Dim AdjRoot,sExp,TempNode,sDispVal,iCtr
	Set AdjRoot = AdjDet.documentElement
	sExp = "//AdjDet"
	Set TempNode = AdjRoot.selectNodes(sExp)
	For iCtr = 0 to TempNode.length - 1
		sDispVal = sDispVal&" Bill-No:"&TempNode.Item(iCtr).Attributes.getNamedItem("AdjCrVouNo").Value
		sDispVal = sDispVal&" Dt:"&TempNode.Item(iCtr).Attributes.getNamedItem("AdjCrVouDate").Value
		sDispVal = sDispVal&" Bill-Amt:"&TempNode.Item(iCtr).Attributes.getNamedItem("AdjCrVouAmt").Value
		sDispVal = sDispVal&" Adj-Amt:"&TempNode.Item(iCtr).Attributes.getNamedItem("AdjCrVouAdjAmt").Value
		sDispVal = sDispVal&vbCrLf
	Next

	MsgBox sDispVal,,"Bank Voucher "

End Function
'===========================Added by Maheshwari on Dec 31st 2007 for TDS Calculation===================
Function TDSAmount()
Dim Root,sOrgId,sGrpId,sEntryNo,sTotAmt,Nd1,Root1

sOrgId = document.formname.hOrgId.value
sGrpId = document.formname.SelTDSGrp.value
sEntryNo = spEntryNo.innerText
TdsAmt = document.formname.txtAmount.value

sTotAmt = 0
set objhttp = CreateObject("MSXML2.XMLHTTP")
document.formname.hTdsNew.value ="Y"
objhttp.Open "GET","TDSCalcCash.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&GrpId="&sGrpId, false
objhttp.send
	 'alert objhttp.responseText
set Root = TDSData.DocumentElement
If Root.haschildnodes then
	for each node1 in Root.childnodes
		If node1.nodename = "TDS" then
			set TDSNode = node1
			Root.Removechild TDSNode
		End If 'If node1.nodename = "TDS" then
	next 'for each node1 in Root.childnodes
End If 'If Root.haschildnodes then
'alert(OutValue.xml)
sTotAmt = 0
if objhttp.responseText <> "" then

	OutData.loadXML objhttp.responseXML.xml
	Set Root1 = OutData.DocumentElement
		'alert("OutData="&Root1.xml)
	If Root1.haschildnodes then
		For Each Nd1 in Root1.childnodes
			If Nd1.NodeName = "TDS" then
				sTemp = trim(Nd1.GetAttribute("PayRecAmount"))
				If trim(sTemp) = "" then
					sTemp = 0
				End If
				Root.appendchild Nd1
				sTotAmt = sTotAmt + sTemp
			End If 'If Nd1.NodeName = "TDS" then
		Next 'For Each Nd1 in Root.childnodes
	End If 'If Root1.haschildnodes then
End If 'if objhttp.responseXML.xml <> "" then
	sTotAmt = FormatNumber(sTotAmt,2,,,0)
	document.formname.txtTdsAmount.value = sTotAmt


	objhttp.Open "POST","XMLSaveForTDS.asp?Name=TDS_Bank", false
	objhttp.send TDSData.XMLDocument
'alert(Root.xml)

End Function

'=======================================================================================================
Function TDSCalc()
Dim sOrgId,OutValue,TdsAmt,node1,sTotAmt,TDSNode,objhttp
sOrgId = document.formname.hOrgId.value
sGrpId = document.formname.SelTDSGrp.value
sEntryNo = spEntryNo.innerText
TdsAmt = document.formname.txtAmount.value
sNewAmt = document.formname.hTdsAmt.value
'alert("TdsAmt="&TdsAmt)
'alert("NewTdsAmt="&sNewAmt)
'alert(sGrpId)
set objhttp = CreateObject("MSXML2.XMLHTTP")
set Root = TDSData.DocumentElement
'alert("Root="&Root.xml)
'alert("EntRoot="&EntryRoot.xml)
If document.formname.hAmendTy.value ="Y" then
	If EntryRoot.haschildnodes then
		For each ChildEnt in EntryRoot.childnodes
			If trim(ChildEnt.nodename) = "TDS" then
				Set Ndnode = ChildEnt
				EntryRoot.RemoveChild Ndnode
			End If 'If trim(ChildEnt.nodename) = "TDS" then
		Next 'For each ChildEnt in EntryRoot.childnodes
	End If 'If EntryRoot.hacchildnodes then
ElseIF document.formname.hAmendTy.value <> "Y" then
	If EntryRoot.haschildnodes then
		For each ChildEnt in EntryRoot.childnodes
			If trim(ChildEnt.nodename) = "TDS" then
				Set Ndnode = ChildEnt
				Root.AppendChild Ndnode

			End If 'If trim(ChildEnt.nodename) = "TDS" then
		Next 'For each ChildEnt in EntryRoot.childnodes
	End If 'If EntryRoot.hacchildnodes then
End If 'If document.formname.hAmendTy.value ="Y" then
'alert("New="& Root.xml)
'alert("sGrpId="&sGrpId)
If trim(sGrpId) = "0" then
	Alert("select TDS Group")
	Exit Function
End If
sCallFrom = document.formname.hCallFrm.value
sVouName  = document.formname.hVouName.value
sNewVal = document.formname.hTdsNew.value
sUpdate = document.formname.hUpdate.value

'alert(EntryRoot.xml)
'window.open "TDSGroupSelection.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&NewAmt="&sNewAmt&"&GrpId="&sGrpId,"TDSData","",""
Set OutValue = ShowModalDialog("TDSGroupSelectionCash.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&NewAmt="&sNewAmt&"&GrpId="&sGrpId&"&CallFrom="&sCallFrom&"&VouName="&sVouName&"&NewVal="&sNewVal&"&Update="&sUpdate,TDSData,"","dialogHeight:350px;dialogWidth:380px;center:Yes;status:no")

'window.open "TDSGroupSelection.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&GrpId="&sGrpId

set Root = TDSData.DocumentElement
If Root.haschildnodes then
	for each node1 in Root.childnodes
		If node1.nodename = "TDS" then
			set TDSNode = node1
			Root.Removechild TDSNode
		End If 'If node1.nodename = "TDS" then
	next 'for each node1 in Root.childnodes
End If 'If Root.haschildnodes then
'alert(OutValue.xml)
sTotAmt = 0
If OutValue.haschildnodes then
	for each node1 in OutValue.childnodes
		If node1.NodeName = "TDS" then
			set TDSNode = node1
			Root.appendchild TDSNode
			sTotAmt = CDbl(sTotAmt) + CDbl(TDSNode.getattribute("PayRecAmount"))
		End If 'If node1.NodeName = "TDS" then
	next 'for each node1 in OutValue.childnodes
End If 'If OutValue.haschildnodes then


sTotAmt = FormatNumber(sTotAmt,2,,,0)
'MsgBox sTotAmt
document.formname.txtTdsAmount.value = sTotAmt
	objhttp.Open "POST","XMLSaveForTDS.asp?Name=TDS_Bank", false
	objhttp.send TDSData.XMLDocument
'alert(Root.xml)
End Function
'=======================================================================================================
Function AmtFun()
Dim sGrpId,TdsAmt
sGrpId = document.formname.SelTDSGrp.value
TdsAmt = document.formname.txtAmount.value

'alert(sGrpId)
If trim(sGrpId) <> "0" then
	TDSAmount()
End If
End Function
'=======================================================================================================
'=======================================================================================================
Function PopCCAH()
    iCCCount=0
    iAHCount=0
 sOrgID =EntryRoot.getAttribute("AccUnit")
 if Trim(sOrgID) = "0" or Trim(sOrgID)="" or IsNull(sOrgID) then  sOrgID =document.formname.hOrgId.value 
 iVouEntryNo = window.spEntryNo.innerHTML
	        For Each HeaderNode in EntryRoot.childNodes
		        if HeaderNode.nodeName="AccHead" then
			        if HeaderNode.Attributes.getNamedItem("Type").value="P" then
				        SelectHead HeaderNode.Attributes.getNamedItem("No").value,"P",document.formname.selAccHead,CInt(document.formname.hHeadCount.value)
			        else
				        SelectHead HeaderNode.Attributes.getNamedItem("No").value,"G",document.formname.selAccHead,CInt(document.formname.hHeadCount.value)
			        end if
			        'document.formname.txtPayTo.value=HeaderNode.Attributes.Item(3).nodeValue
			        window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue
			        bCostCenter = HeaderNode.getAttribute("CostCenter")
			        bAnalytical = HeaderNode.getAttribute("Analytical")
			        nAccCode = HeaderNode.getAttribute("No")
		        end if 'End of Check for Account head Node
		        if 	HeaderNode.nodeName="Narration" then
			        document.formname.txtNarration.value=HeaderNode.text
		        end if 'End of Check for Narration Node

		        if 	HeaderNode.nodeName="CostCenter" then
			        setADDDisplay 1
			        popCostCenter(HeaderNode)
			        iCCCount = 1
		        end if 'End of Check for Cost Center Node

		        if 	HeaderNode.nodeName="Analytical" then
			        setADDDisplay 1
			        popAnalytical(HeaderNode)
			        iAHCount = 1
		        end if 'End of Check for Analytical Node
		        
		        if 	HeaderNode.nodeName="PayRec" then
			        popPayRec(HeaderNode)
		        end if 'End of Check for Analytical Node

	        next 'End of Entry Node Loop
	        
	    ''added by ragav on Jan 13 ,2012 - if Cash Voucher Creation case CC or AH is not selected means here can select
        ''begin 
	        if cint(bCostCenter)=1 or cint(bAnalytical)=1 then
                'If Selected GL Account Head has Cost Center
                   Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+nAccCode,EntryRoot,"")
                    if nodCCAnly.Attributes.Item(0).nodeValue=1 then
                        'Set the Additional and CCANAL Display Layer Visible
                        setADDDisplay 1
                        'alert(nodCCAnly.xml)
                        'alert(EntryRoot.xml)
                        if EntryRoot.hasChildNodes() then
                            for each ndHead in EntryRoot.childNodes
                                'alert(ndHead.nodeName)
                                if ndHead.nodeName="CostCenter" then
                                    EntryRoot.removeChild ndHead
                                end if
                                if ndHead.nodeName="Analytical" then
                                    EntryRoot.removeChild ndHead
                                end if
                            next
                        end if
                        For Each ndHeader In nodCCAnly.childNodes
                            if ndHeader.nodeName="CostCenter" then
	                            EntryRoot.appendChild ndHeader
	                            popCostCenter ndHeader
                            end if 'End of Check for Cost Center Node
                            if ndHeader.nodeName="Analytical" then
	                            EntryRoot.appendChild ndHeader
	                            popAnalytical(ndHeader)
                            end if 'End of Check for Analytical Node
                        next	'End of Processing CCANAL Node
                    else
                        'User Has canceled CC,ANAL Selection
                        'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
                        setADDDisplay 1
                    end if	'End of CC,ANAL has Childs Check
            else
                'Selected Head has no CC or ANAL
                'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
                setADDDisplay 1
            end if	'End of GL has Cost Center or not
        ''end 

End Function

</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="DisplayBook();SelUnBook()">
<form method="POST" name="formname" action="AmdAccGenerate.asp" >
<input type="hidden" name="hVouCode" value="02">
<input type="hidden" name="hVouType" value="<%=sCallVouTy%>">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hTdsAmt" value="">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=sBookCode%>">
<input type="hidden" name="hOtherUnitFlag" value="<%=bOtherUnits%>">
<input type="hidden" name="hActionFlag" value="<%=bActionFlag%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hPayTo" value="">
<input type="hidden" name="hTDSElgi" value="0">
<input type="hidden" name="hTotalAmt" value="0">
<input type="hidden" name="hPayRecCount" value="0">
<input type="hidden" name="hSelPayRecCount" value="0">
<input type="hidden" name="hTotType" value="N">

<%if Trim(sVal)<>"" then%>
	<input type="hidden" name="hTransNo" value="<%=sValTemp(0)%>">
	<input type="hidden" name="hAmendDet" value="<%=sValTemp(1)%>">
	<input type="hidden" name="hCallFrm" value="<%=sValTemp(2)%>">
<%else%>
	<input type="hidden" name="hCallFrm" value="C">
	<input type="hidden" name="hTransNo" value="0">
<%End if%>

<input type="hidden" name="hLastVouDt" value="<%=sLastVouDt%>">
<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">
<input type="hidden" name="hAmendTy" value="N">
<input type="hidden" name="hBookAccHead" value="<%=iBookAcchead%>">
<input type="hidden" name="hPopVouTy" value="<%=sPopVouTy%>">
<input type="hidden" name="hMinDate" value="<%=sMaxDate%>">
<input type="hidden" name="hMaxDate" value="<%=sMinDate%>">
<input type="hidden" name="hBookOtherUnit" value="1">
<input type="hidden" name="hPreBookSel" value="<%=iPreBookVal%>">
<input type="hidden" name="hInsDet" value="">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">

<input type="hidden" name="hFormVal" value="<%=sFormVal%>">
<input type="hidden" name="voutype" value="<%=sSelArg%>">
<input type="hidden" name="ContraStatus" value="<%=sConStatus%>">
<input type="hidden" name="hCurrConStatus" value="0">
<input type="hidden" name="hNorToContra" value="N">

<input type="hidden" name="hUpdate" value="N">
<input type="hidden" name="hTdsNew" value="N">

<input type="hidden" name="hInsrAmt" value="0">
<input type="hidden" name="hAction" value="New">


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<% IF CStr(sCallVouTy) = "R" Then
				Response.Write("Bank Receipt Voucher Amendment")
		   Else
				Response.Write("Bank Payment Voucher Amendment")
		   End IF
		%>
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
								<!--td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td-->
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
								  		<td align="center">Voucher</td>
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
								<td align="center" colspan="3" height="8">
                                  <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                   <tr>
                                      <td class="MiddlePack" colspan="6"></td>
                                    </tr>
                                    <tr>
                                       <td class="FieldCellSub" width="110">Book</td>
										<td class="FieldCell" colspan="3">
											<select size="1" name="selBook" class="FormElem" onChange="SetBookAccHead()">
												<!--option value="S">Select</option-->
											</select>
										</td>
										<td class="FieldCellSub" width="90">Voucher Type</td>
										<td class="FieldCellSub" colspan="2">
										<input type=hidden name=hAccUnit value="<%=sOrgId%>">
											<select size="1" name="selVouType" class="FormElem" onChange="CheckEntryType(this)" disabled>
												<option value="S">Select</option>
												<%IF CStr(sVouType) = "D" Then %>
													<option value="D" Selected>Receipt</option>
													<option value="C">Payment</option>
												<%Else%>
													<option value="D">Receipt</option>
													<option value="C" Selected>Payment</option>
												<%End IF %>
											</select>
										</td>

										</tr>
										
										<tr>

											<td class="FieldCellSub" width="110">Voucher No
											<td class="FieldCell" colspan="3">
												<%if Trim(sVal)<>"" then%>
													<input type="text" name="txtVouNo" size="20" class="FormElemRead" value="<%=sValTemp(1)%>" readonly>
												<%else%>
													<input type="text" name="txtVouNo" size="20" class="FormElemRead" readonly>
												<%end if%>
											</td>
											<!--td class="FieldCellSub">
												<a href="javascript:popVoucherNo('C','02','BA',document.formname.selUnitId.value,document.formname.hPopVouTy.Value,document.formname.selBook.value)">
												<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
												 &nbsp;&nbsp;
												 <a href="javascript:popVoucherNo('A','02','BA',document.formname.selUnitId.value,document.formname.hPopVouTy.Value,document.formname.selBook.value)">
												<img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers"></a>
											</td-->
                                      <td class="FieldCellSub" width="90">Date</td>
                                      <td class="FieldCell" width="110">
									      <%'Function Call to Insert Date Picker
											Response.Write InsertDatePicker("ctlDate")%>
 									  </td>
								</tr>
								
								<tr>
									<td class="FieldCellSub" width="110">Entry No</td>
									<td class="FieldCell" colspan="3"><span class="DataOnly" id="spEntryNo">1&nbsp;</span>
									</td>
								</tr>
								<tR>
									<td class="FieldCellSub" width="110">&nbsp;</td>
									<td class="FiledCell">
										<Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()">
									</td>
									<td class="FieldCell" colspan=5>
								    <span id="spInsDet" class="DataOnly">&nbsp;</span>
								    &nbsp; <a href="#" onclick="PopInsDet()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Enter Instrument Details"></a>
								</td>
								</tr>
								
								<!--<tr>-->
									<!--td class="FieldCellSub" >
										<Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()">
									</td-->
									<!--<td class="FieldCellSub" colspan="2">Entry No &nbsp; <span class="DataOnly" id="spEntryNo">1&nbsp;</span>-->
									<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()" -->
									<!--</td>
								</tr>-->
								<tr>

									<td class="FieldCellSub">Current Balance</td>
									<td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spCurrBal">
                                <%

										dOpeningBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
										'dOpeningBal = 0
										dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
										if dOpeningBal<0 then
											Response.Write FormatNumber(dOpeningBal*-1,2,,,0) &"&nbsp;Cr"
										else
											Response.Write FormatNumber(dOpeningBal,2,,,0) &"&nbsp;Dr"
										end if
                                %> </span>
                                &nbsp;</td>
                                <td class="FieldCellSub">Book Balance</td>
								<td class="FieldCellSub">
									<span class="DataOnly" id="spBookBal">
									<%
									dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
									'dOpeningBal = 0
									dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
									if dOpeningBal<0 then
									Response.Write FormatNumber(dOpeningBal*-1,2,,,0) &"&nbsp;Cr"
									else
									Response.Write FormatNumber(dOpeningBal,2,,,0) &"&nbsp;Dr"
									end if
									%> </span>&nbsp;
								</td>

								</tr>
								<tr id="BankDet" style="visibility: hidden;">
									<td class="FieldCellSub">Inst No</td>
									<td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spInsNo">
									<td class="FieldCellSub">Inst Date</td>
									<td class="FieldCellSub" colspan="2"><span class="DataOnly" id="spInsDT">

								</tr>

								<tr>
                            </table>
                            </td>
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
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
                                                            <table border="0" cellspacing="0" class="TableOutlineOnly" cellpadding="0" width="100%">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="2" ></td>
                                                        </tr>
                                                        <!--<tr>
                                                    <td class="FieldCellSub">Accounting Unit</td>
                                                    <td class="FieldCellSUB">

                                                     <select size="1" name="selAccUnitId" onChange="popAccHead()" class="FormElem">
														<option value="A" >Account Unit</option>
														<%'populateOrganizationListDBWithVal(sOrgID)%>
														<OPTION value="<%=sOrgId%>" Selected><%=sOrgName%></option>
													</select>
													<%
														'else
														'	Response.Write sOrgName
														'end if
													%>
													</td>-->
													<!--td class="FieldCellSub">Entry No
													&nbsp; <span class="DataOnly" id="spEntryNo">1&nbsp;</span>
													</td-->
                                                        <!--</tr>-->
                                                        <tr>
                                                        <td class="FieldCellSub" width="110">Entry Type
									<td class="FieldCell" colspan="2">
									<% IF CStr(sCallVouTy) = "R" Then %>
										<input type=radio name="selCRDR" value="C" checked>Receipts
										<input type=radio name="selCRDR" value="D" disabled>Payments&nbsp;
                                    <%Else%>
										<input type=radio name="selCRDR" value="C"disabled>Receipts
										<input type=radio name="selCRDR" value="D" checked>Payments&nbsp;
									<%End IF %>
									</td>
									</tr>
									
									<tr>
                                        <td class="FieldCellSub" >Accounting Head</td>
                                        <td class="FieldCellSub" colspan="2">
                                                <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
												<option value="A">Select Account Head</option>
												<%
													dim iHeadCount
												 	'iHeadCount=popFrequentHead(sOrgId,"02",sBookCode)
													iHeadCount=0
												%>
													<option value="G">General Ledger</option>
												<%populatePartyType(sOrgId)%>
                                        </select> &nbsp; <a href="javascript:selAccountHead(document.formname.selAccHead)"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Account Head"></a>
                                        </td>
                                        <input type="hidden" name="hHeadCount" value="<%=iHeadCount%>">
									</tr>
									
                                    <tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td class="FieldCellSub"><span class="DataOnly" id="spAccHead"></span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" >Pay to / Received from</td>
                                                    <td class="FieldCellSub" colspan="2"> <input type="text" name="txtPayTo" size="55" class="Formelem" maxlength="50">
                                                    &nbsp; <a href="javascript:SelMisParty()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a></td>
                                                        </tr>
                                                        <tr>
                                                    <td width="139" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td class="FieldCellSub">Narration</td>
                                                          <td class="FieldCellSub">
<%

'sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
'	" OUDefinitionID='"&sOrgId&"'and BookCode='02' and BookNumber="&sBookCode

'with objRs
'	.CursorLocation = 3
'	.CursorType = 3
'	.Source = sQuery
'	.ActiveConnection = con
'	.Open
'end with
'set objRs.ActiveConnection = nothing

'if objRs(0)>0 then
%>
                                                            <p align="left">
                                                    <a href="javascript:showNarration('02')"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
<%
'end if
'objRs.Close
%>
                                                           </td>
                                                        </tr>
                                                      </table>
                                                      &nbsp;</td>
                                                    <td class="FieldCellSub" valign="top"> <textarea rows="2" name="txtNarration" cols="50" class="FormElem" onKeyPress="ChkEnter()"></textarea> </td>
                                                        </tr>
                                                        <tr>
                                                  <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtAmount" size="15" value="0.00" style="text-align:right" maxlength="13" class="Formelem" onblur="popAddAmount();TDSAmount()"> </td>
                                                        </tr>
                                                        <tr>
                                                        <td colspan="2">
                                                        <div id="DisCCANL" class=frmBody style="height:1; visibility: hidden;">
                                                            <table cellpadding="0" cellspacing="0" >
                                                                <tr>
	                                                                <td class=MiddlePack colspan="3"> </td>
                                                                </tr>
                                                                <tr>
	                                                                <td class=FieldCell>
		                                                                <DIV class=frmBody id="DisCost" style="width:260;height:100;">
			                                                                <table border="0" id="tblCost" cellspacing="1" class="ExcelTable">
				                                                                <tr>
					                                                                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
						                                                                <td class="ExcelHeaderCell" align="center" width="150">Cost Center Head
						                                                                <img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Click Here to add Cost Center or Analytical Head" onclick="PopCCAH()">
						                                                                </td>
						                                                                <td class="ExcelHeaderCell" align="center">Ratio</td>
						                                                                <td class="ExcelHeaderCell" align="center">Amount</td>
				                                                                 </tr>
			                                                                </table>
		                                                                </div><!--End of CostCenter Display Division -->
	                                                                </td>
	                                                                <td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">                   </td>
	                                                                <td class=FieldCell>
		                                                                <DIV class=frmBody id="DisAnal" style="width:260; height:100;">

			                                                                <table border="0" id="tblAnal" cellspacing="1" class="ExcelTable">
				                                                                <tr>
						                                                                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
						                                                                <td class="ExcelHeaderCell" align="center" width="150">Analytical Head
						                                                                <img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Click Here to add Cost Center or Analytical Head" onclick="PopCCAH()">
						                                                                </td>
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
                                                        </td>
                                                        </tr>
                                                         <tr>
                                                    <td class="FieldCellSub" >Select TDS Group</td>
                                                    <td class="FieldCellSub" width="591">
                                                    <select size="1" name="SelTDSGrp" class="FormElem" onchange="TDSAmount()">
                                                    <Option Value="0" selected> Select </option>
		                                                  <% Dim sUseable,sGrpID,sTemp

																sQuery = "Select GroupID,GroupName from ACC_M_TDSGroup where OUDefinitionID = '"& sOrgId &"' and isNull(Useable,'Y') <> 'N' "
																	'Response.Write sQuery
																	With objRs1
																		.CursorLocation = 3
																		.CursorType = 3
																		.ActiveConnection = con
																		.Source = sQuery
																		.Open
																	End With
																	Do while Not objRs1.EOF
																		sGrpId = objRs1(0)
																	Response.Write objRs1(1)& "<BR>"%>
																	<option value="<%=objRs1(0)%>" <%'If trim(sGroupName) = trim(objRs1(0)) then Response.Write "selected" %>> <%=objRs1(1)%> </option>
																	<%objRs1.MoveNext
																Loop
																objrs1.Close
															%>
                                                    </select>
                                                    &nbsp; % On Amount &nbsp; <input type="text" name="txtTdsAmount" Value="" size="15" style="text-align:right" maxlength="13" class="Formelemread" readonly >
                                                    <a href="javascript:TDSCalc()"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="TDS Group Selection" width="10" height="11"></a>
                                                    <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
                                                    </td>
                                                        </tr>
                                                         <!--tr>
															<td class="FieldCellSub" width="133">Approval</td>
															<td class="FieldCell" width="591">
															<input type="radio" value="Y" checked name="optApprove" class="FormElem">
															Yes&nbsp;&nbsp;
															<input type="radio" value="N" name="optApprove" class="FormElem"> No </td>
														</tr-->
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
								<td class="FieldCellSub" width="639">Approval

								<input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="SetApp('Y')">
								Yes&nbsp;&nbsp;
								<input type="radio" value="N" name="optApprove" class="FormElem" onClick="SetApp('N')"> No
								&nbsp;&nbsp; Approver &nbsp; <select size="1" name="selUserId" class="FormElem">
											<option value="I">Immediate Approver</option>
											<%=populateEmployeeWithVal(sUserId)%>
											    </select></td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td >
<DIV class=frmBody id="Disaddtional" style="height:1; visibility: hidden;">
	<DIV class=frmBody id="DisPayable" style="width: 555; visibility: hidden; height:1;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="555">
			<tr>
				<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
				<td class="ExcelHeaderCell" align="center" colspan="2">Document</td>
				<td class="ExcelHeaderCell" align="center" width="275" colspan="5">Amount</td>
		    </tr>
		   <tr>
				<td class="ExcelHeaderCell" align="center">Detail</td>
				<td class="ExcelHeaderCell" align="center">Date</td>
				<td class="ExcelHeaderCell" align="center">Amount</td>
				<td class="ExcelHeaderCell" align="center">Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To Account</td>
				<td class="ExcelHeaderCell" align="center">To be Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To adjust</td>

		   </tr>
		</table>
	</div>
</div><!--End of Addtional Details Display  -->
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
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="Button" value="Update Entry" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButtonX" >
                                                                <input type="Button" value="Delete Entry" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButtonX" >
                                                                <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Delete Voucher" name="btnDelVou" onClick="DelVouch()" class="ActionButtonX" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('BANKVOUCHERS.ASP')" class="ActionButton">

														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" style="width:660;" >
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="25"></td>
		<td class="ExcelHeaderCell" align="center">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center">Additional Details</td>
		<td class="ExcelHeaderCell" align="center">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="70">Amount</td>
		<td class="ExcelHeaderCell" align="center" width="70">Deduction Amount</td>
		<td class="ExcelHeaderCell" align="center" width="70">Deduction Percentage</td>
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