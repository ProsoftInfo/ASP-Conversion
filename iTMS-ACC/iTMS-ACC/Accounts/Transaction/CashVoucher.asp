
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	CashVoucher.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	June 10, 2005
	'Modified By                :   Ragavendran R
	'Modified On                :   Jan 18,2011
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
<!--#include File="../../include/IncludeDatePicker.asp" -->
<!--#include File="../../include/CheckACCPrevFinYear.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,sBookName,sVouType,sTransNo,sQuery
dim iVouNo,objRs,objRs1,sVouDate,bActionFlag,sVal,sValTemp
dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sAccHead,sParType,sPartSubType
dim iEnNo,Entrynode,HeaderNode,dOpeningBal
dim sParCode,sNarration,sAccHeadname,sAccUnitName,bOtherUnits,iBookAccHead,dTransLimit
Dim sVouCkTy,sLastVouDt,sSelVouTy
Dim sFinPeriod,sFinFrm,sFinTo,sValTemp2,sFormVal,sSelArg
dim sAccount,sAddtional,iSno,sAction
dim dTotal
dim sVoucDate,iBookCode,sPayTo,sUserId,iPreBookVal
Dim sRetVal,sFinFromDate,sFinToDate
sUserId = session("userid")
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2,sLogUID

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("hVouType")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")
sSelVouTy = Request("VOUTY")

sSelArg = Request("voutype")
sFormVal = Request("hFormVal")

iPreBookVal = sBookCode

sOrgId = Session("organizationcode")
sOrgName = Session("OrgShortName")
'Response.Write "sOrgID = "& sOrgId
'Response.Write "sOrgName = "& sOrgName
'Response.Write sVouType
sAction = Session("ACTN")
sVal=Request("Val")
'Response.Write sVal
sValTemp=Split(sVal,"~")
IF CStr(sVouType) = "" Then
	IF CStr(sSelVouTy) = "P" Then
		sVouType = "C"
	Else
		sVouType = "D"
	End IF
End IF
'Response.Write sVouType
IF Cstr(sVal) <> "" Then
	sQuery = "Select H.OUDEFINITIONID,D.OrgUnitDescription,isNull(AccountHead,0),BookNumber From DCS_OrganizationUnitDefinitions D, "&_
		 "Acc_T_CreatedVoucherHeader H Where H.OUDEFINITIONID = D.OUDEFINITIONID "&_
		 "and H.CreatedTransNo = "&sValTemp(0)&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
'		sOrgId = objRs(0)
'		sOrgName = objRs(1)
        iBookAccHead = objrs(2)
        sbookCode = objrs(3)
	End IF
	objRs.Close

Else

	sQuery = "Select Top 1 OUDefinitionID,OrgUnitDescription From DCS_OrganizationUnitDefinitions "&_
			 "Where Len(OUDefinitionID) > 4 Order By OUDefinitionID "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
'		sOrgId = objRs(0)
'		sOrgName = objRs(1)
	End IF
	objRs.Close
End IF

IF CStr(iBookAccHead) = "" Then
	sQuery = "Select Top 1 BookNumber,BookName,isNull(BookAccountHead,0),OtherUnitTransaction From vwOrgBookNames Where  "&_
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '01' Order By BookName "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sBookCode = objRs(0)
		sBookName = objRs(1)
		iBookAccHead = objRs(2)
		bOtherUnits = objRs(3)
	Else
		iBookAccHead = 0
	End IF
	objRs.Close
else
sQuery = "Select Top 1 OtherUnitTransaction From vwOrgBookNames Where  "&_
			 "OUDefinitionID = '"&sOrgId&"' and BookCode = '01' Order By BookName "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		bOtherUnits = objRs(0)
	Else
		bOtherUnits = 1
	End IF
	objRs.Close
End IF

oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

''Blocked and Added by Ragav on Jan 13,2012
''Begin
'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
sRetVal = GetVouchXML(sTransNo)
IF Request.Form("hCallFrm") = "A" then
    oDOM.Load server.MapPath(sRetVal)
End IF
''End

'Response.Write Request.Form("hCallFrm")
'IF Request.Form("hCallFrm") <> "A" then
	'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")
'End IF


IF CStr(sVouType) = "C" Then
	sVouCkTy = "CAP"
Else
	sVouCkTy = "CAR"
End IF

sQuery = "Select Convert(Char,VoucherDate,103),CreatedBy From Acc_T_CreatedVoucherHeader Where CreatedTransNo =  "&_
		 "(Select Max(CreatedTransNo) From Acc_T_CreatedVoucherHeader Where BookCode = '01'  "&_
		 "and OUDefinitionID = '"&sOrgId&"' and TransactionType ='"&sVouCkTy&"' ) "

objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sLastVouDt = Trim(objRs(0))
	'sUserId = Trim(objRs(1))
End IF
objRs.Close

'IF CStr(sUserId) = "" Then
sUserId = session("userid")
'End IF

sLogUID = session("userid")

sFinPeriod = Session("FinPeriod")
sValTemp2 = Split(sFinPeriod,":")
sFinFrm = Trim(sValTemp2(0))
sFinTo = Trim(sValTemp2(1))
sFinFrm = sFinFrm&"04"
sFinTo = sFinTo&"03"
bOtherUnits = 1
sFinFromDate = "01/04/"& sValTemp2(0)
sFinToDate = "31/03/"&sValTemp2(1)



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS Cash Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<meta http-equiv="x-ua-compatible" content="IE=10">
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<SCRIPT language="javascript" SRC="../scripts/VouSelection.js"></SCRIPT>
<script language="javascript" src="../scripts/VoucherEntryCore.js"></script>
<script language="javascript" src="../scripts/CashVoucher.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=sBookCode%>" BookName="<%=sBookName%>" CRDR="<%=sVouType%>" VouDate="" BookAcchead="<%=iBookAccHead%>" Approver=""/></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="0" TDSElgi="0" TdsPercentage="0" PayRecAmount="0" /></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="TDSData"  ><Root/></xml>
<xml id="GLHeadData"><Root /></xml>
<xml id="PartyHeadData"><Root /></xml>
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

Function DisplayBook(objUnit)
dim iUnitNo,arrTemp
dim Root,sVal
document.formname.selBook.options.length = 1
'document.formname.selAccUnitId.selectedIndex = objUnit.selectedIndex
'document.formname.hOrgId.value = document.formname.selAccUnitId.Value
'document.formname.hOrgName.value = document.formname.selAccUnitId.Options(document.formname.selAccUnitId.selectedIndex).text

popAccHead
'if objUnit.selectedIndex <> "0" then
	'iUnitNo= objUnit(objUnit.selectedIndex).value
	iUnitNo = document.formname.hOrgId.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=01&orgID=" & iUnitNo , false
	objhttp.send
	'alert objhttp.responseXML.xml
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

IF document.formname.selBook.length > 1 Then
    For iCnt = 0 to cint(document.formname.selbook.length)-1
        if document.formname.selBook(iCnt).value =trim(document.formname.hBookCode.value)&"-"&trim(document.formname.hBookAccHead.value) then
            document.formname.selBook.selectedIndex = iCnt
        end if
    Next
	SetBookAccHead()
Else
	document.formname.selBook.selectedIndex = 0
End IF

end Function

Function SaveXML()

	if bSavFlag then
		'IF CheckVoucherDt() Then ' Checking for the VoucherDate between the Last Voucher Date and Current Date
			IF CheckApp() Then ' Checking For Selected/Entered Values
				IF CheckContraEnt() Then ' Checking Wheather No Series is Defined or Not
					IF document.formname.selBook.selectedIndex = 0 Then
						MsgBox "Select Book"
						Exit Function
					End IF
					' alert VoucherData.xml
					'exit function
					IF CheckFinDate Then
						UpdateXML()
						set objhttp = CreateObject("Microsoft.XMLHTTP")
					'	alert(document.formname.hAmendTy.value)
						IF Cstr(document.formname.hAmendTy.value) = "N" Then
							objhttp.Open "POST","XMLSave.asp?Name=Voucher Entry&Mod=CA", false
							document.formname.action = "VouGenerate.asp"
						Else
							objhttp.Open "POST","XMLSave.asp?Name=Voucher AMD&Mod=CA", false
							document.formname.action = "VouAmdGenerate.asp"
						End IF

						' alert VoucherData.XML
						objhttp.send VoucherData.XMLDocument

					    'exit function
						if objhttp.responseText <> "" then
							Msgbox(objhttp.responseText)
						else
							document.formname.btnNext.disabled = True
							IF CStr(document.formname.hAmendTy.value) = "Y" Then
								document.formname.action = "VouAmdGenerate.asp"
							End IF
							'document.formname.selVouType.disabled = False
							'MsgBox "ok "
							document.formname.btnNext.disabled = True
							document.formname.submit()
						end if
					End IF
				End IF
			End IF
		'End IF
	end if
End Function
'---------------------End Of Function SaveXML-----------------------------
Function PrnVouch()
	sTransNo = document.formname.hTransNo.value
	sOrgName = document.formname.hOrgName.Value
	sVouTy = document.formname.hVouType.Value

	sValue = sTransNo&":"&sOrgName
	IF CStr(sVouTy) = "D" Then
		sStatus= showModalDialog("PRNCashRecVouView2.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
	Else
		sStatus= showModalDialog("PRNCashPayVouView2.asp?Value="&sValue,"","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
	End IF
End Function

Function CheckPayRecAmt()
	Dim sExp,TempNode,iCtr,dAddAmt,dSubAmt,PayNode,sParTy,sTemp,sDispType
	Dim sType,sRetVal
	dAddAmt = 0
	dSubAmt = 0
	sType = "T"



	IF CStr(document.formname.hVouCRDR.Value) = "D" Then
		sDispType = "Receipt Amount Should be Greater Than Payment Amount "
	Else
		sDispType = "Payment Amount Should be Greater Than Receipt Amount "
	End IF

	'MsgBox EntryRoot.xml



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

			'MsgBox dAddAmt &" -- " & dSubAmt

		End IF

		'MsgBox EntryRoot.Attributes.Item(9).nodeValue
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

	IF CStr(sType) = "T" Then
		'IF CStr(sParTy) = "CR" Then
			IF CDbl(document.formname.txtAmount.value) = dAddAmt - dSubAmt Then
				CheckPayRecAmt = True
				Exit Function
			Elseif CDbl(document.formname.txtAmount.value) > dAddAmt - dSubAmt Then
				sRetVal = MsgBox("Entered Value is Greater than Adjusted Value! Remaing amount will treat as Advance Continue!! ",4,"Cash Voucher")
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
		IF CDbl(document.formname.txtAmount.value) >= CDbl(dAddAmt) - CDbl(dSubAmt) Then
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
	iUnitNo = document.formname.hOrgID.value
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
				MsgBox "Contra Entry is Created only One Entry is allowed "
				CheckContraEnt = False
				Exit Function
			End IF
		Next
	End IF

	CheckContraEnt = True
End Function
'===========================Added by Maheshwari on Dec 29th 2007 for TDS Calculation===================
Function TDSAmount()
Dim Root,sOrgId,sGrpId,sEntryNo,sTotAmt,Nd1,Root1

sOrgId = document.formname.hAccUnitId.value
sGrpId = document.formname.SelTDSGrp.value
sEntryNo = spEntryNo.innerText
TdsAmt = document.formname.txtAmount.value

sTotAmt = 0
set objhttp = CreateObject("MSXML2.XMLHTTP")
set Root = TDSData.DocumentElement
document.formname.hTdsNew.value ="Y"
'alert(EntryData.xml)
objhttp.Open "GET","TDSCalcCash.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&GrpId="&sGrpId, false
objhttp.send
 ' alert objhttp.responseText
set Root = TDSData.DocumentElement
If Root.haschildnodes then
	for each node1 in Root.childnodes
		If node1.nodename = "TDS" then
			set TDSNode = node1
			Root.Removechild TDSNode
		End If 'If node1.nodename = "TDS" then
	next 'for each node1 in Root.childnodes
End If 'If Root.haschildnodes then
'alert(TDSData.xml)
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


	objhttp.Open "POST","XMLSaveForTDS.asp?Name=TDS_Cash", false
	objhttp.send TDSData.XMLDocument
'alert(Root.xml)

End Function

'=======================================================================================================
Function TDSCalc()
Dim sOrgId,OutValue,TdsAmt,node1,sTotAmt,TDSNode,objhttp
sOrgId = document.formname.hAccUnitId.value
sGrpId = document.formname.SelTDSGrp.value
sEntryNo = spEntryNo.innerText
TdsAmt = document.formname.txtAmount.value
sNewAmt = document.formname.hTdsAmt.value
'alert(document.formname.hAmendTy.value)
'alert("TdsAmt="&TdsAmt)
'alert("NewTdsAmt="&sNewAmt)
'alert(sGrpId)
set objhttp = CreateObject("MSXML2.XMLHTTP")
set Root = TDSData.DocumentElement
'alert("Root="&Root.xml)
'alert("EntRoot="&EntryRoot.xml)
'IF document.formname.hTdsNew.value ="Y" then
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
End If
 'alert("New="& Root.xml)
'alert("sGrpId="&sGrpId)
'alert("Root="&Root.xml)
If trim(sGrpId) = "0" then
	Alert("select TDS Group")
	Exit Function
End If
'alert(document.formname.hCallFrm.value)

sCallFrom = document.formname.hCallFrm.value
sVouName  = document.formname.hVouName.value
sNewVal = document.formname.hTdsNew.value
sUpdate = document.formname.hUpdate.value

'window.open "TDSGroupSelectionCash.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&NewAmt="&sNewAmt&"&GrpId="&sGrpId&"&Para="&sPara,"TDSData","",""
Set OutValue = ShowModalDialog("TDSGroupSelectionCash.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&NewAmt="&sNewAmt&"&GrpId="&sGrpId&"&CallFrom="&sCallFrom&"&VouName="&sVouName&"&NewVal="&sNewVal&"&Update="&sUpdate,TDSData,"","dialogHeight:350px;dialogWidth:380px;center:Yes;status:no")
'window.open "TDSGroupSelectionCash.asp?EntNo="&sEntryNo&"&Amount="&TdsAmt&"&GrpId="&sGrpId

set Root = TDSData.DocumentElement
If Root.haschildnodes then
	for each node1 in Root.childnodes
		If node1.nodename = "TDS" then
			set TDSNode = node1
			Root.Removechild TDSNode
		End If 'If node1.nodename = "TDS" then
	next 'for each node1 in Root.childnodes
End If 'If Root.haschildnodes then
 'alert("1="&OutValue.xml)
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
'alert("TTT="&TDSData.xml)
document.formname.txtTdsAmount.value = sTotAmt
objhttp.Open "POST","XMLSaveForTDS.asp?Name=TDS_Cash", false
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
        if Trim(bCostCenter)="" or IsNull(bCostCenter) then bCostCenter = 0
        if Trim(bAnalytical)="" or IsNull(bAnalytical) then bAnalytical = 0
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
'*************************************
Function Init()
   sFromDate = document.formname.hFromDate.value
   sTodate = document.formname.hToDate.value
   if DateDiff("d",sTodate,date)>0 then
        document.formname.ctlDate.setMinDate=sFromDate
        document.formname.ctlDate.setMaxDate=sTodate
        document.formname.ctlDate.setDate=sTodate
   else
        document.formname.ctlDate.setMinDate=sFromDate
        document.formname.ctlDate.setMaxDate=date
        document.formname.ctlDate.setDate=date
   end if


   'document.formname.ctlDate.setMaxDate = date()
End Function
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();SelUnBook();DisplayBook('<%=sOrgID%>')">
<form method="POST" name="formname" action="VouGenerate.asp" >
<input type="hidden" name="hVouCode" value="01">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="CA">
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
<input type="hidden" name="hUpdate" value="N">
<input type="hidden" name="hTdsNew" value="N">
<input type="hidden" name="hAction" value="New">

<input type="hidden" name="hVouType" value="<%=sVouType %>">

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
<input type="hidden" name="hBookOtherUnit" value="1">
<input type="hidden" name="hPreBookSel" value="<%=iPreBookVal%>">
<input type="hidden" name="hFinFrm" value="<%=sFinFrm%>">
<input type="hidden" name="hFinTo" value="<%=sFinTo%>">
<input type="hidden" name="hFormVal" value="<%=sFormVal%>">
<input type="hidden" name="voutype" value="<%=sSelArg%>">
<input type="hidden" name="hFromDate" value="<%=sFinFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sFinToDate%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
		<% IF CStr(sVouType) = "C" Then
				Response.Write("Cash Payment Voucher")
		   Else
				Response.Write("Cash Receipt Voucher ")
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
											<td align="center">Entry Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70px">
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
					<TD class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <!--tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="left">
								<table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable">
									<tr>
										<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: hand" Title="Month wise Balance" >
                    <p align="center"><font face="Webdings" size="5">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Daywise Balance"><font face="Webdings" size="5">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Voucher History">
                    <font face="Webdings" size="5">?</font>
                    </span>
                    </p>
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
								</td>
							</tr>
							<tr>
							    <td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td align="center" height="8">
                                  <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                    <tr>
                                      <td class="FieldCellSub" width="110">Book</td>
                                      <td class="FieldCell" width="110">
										<select size="1" name="selBook" class="FormElem" onChange="SetBookAccHead()">
												<option value="S">Select</option>
											</select>
									   </td>
									   	<td class="FieldCellSub" width="90">Date</td>
										<td class="FieldCell" >
										    <% ' Function Call to Insert Date Picker
									Response.Write InsertDatePicker("ctlDate")%>
										</td>

										</tr>

										<tr>
											<td class="FieldCellSub" width="90">Current Balance
											<td class="FieldCell">
											<span class="DataOnly" id="spCurrBal">
                                                        <%
                                                         dOpeningBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
                                                         'dOpeningBal = 0
                                                         dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                         if dOpeningBal<0 then
															Response.Write dOpeningBal*-1 &"&nbsp;Cr"
														 else
															Response.Write dOpeningBal &"&nbsp;Dr"
														 end if
                                                        %> </span>
                                                        &nbsp;&nbsp;&nbsp;Book Balance
                                                        &nbsp;
															<span class="DataOnly" id="spBookBal">
															<%
															dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
															'dOpeningBal = 0
															dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
															if dOpeningBal<0 then
															Response.Write dOpeningBal*-1 &"&nbsp;Cr"
															else
															Response.Write dOpeningBal &"&nbsp;Dr"
															end if
															%> </span>&nbsp;
									        </td>
                                      <td class="FieldCellSub" width="80">Voucher No</td>

                                      <td class="FieldCell" width="110">
                                      <%if Trim(sVal)<>"" then%>
													<input type="text" name="txtVouNo" size="20" class="FormElem" value=<%=sValTemp(1)%> readonly>
												<%else%>
													<input type="text" name="txtVouNo" size="20" class="FormElem" readonly>
												<%end if%>

								</td>
                                </tr>
                                </tr>

                            </table>
                            </td>
                            <td align="center" width="5" class="ClearPixel">
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
                                                    <td class="MiddlePack" colspan="2" width=100%>
                                                            <tr>
                                                                <td class="FieldCellSub" width="100">Entry Type</td>
                                                                <td class="FieldCell">
                                                                <table border=0 width="100%">
                                                                        <tr>
                                                                        <td class="FieldCell">
													                     <input type=hidden name="hAccUnitId" value="<%=sOrgId%>">
													                    <%IF CStr(sVouType) = "D" Then %>
                                                                            <input type=radio name="selCRDR" value="C" checked>Receipts
                                                                            <input type=radio name="selCRDR" value="D" disabled>Payments&nbsp;
                                                                        <%Else%>
														                    <input type=radio name="selCRDR" value="C" disabled>Receipts
														                    <input type=radio name="selCRDR" value="D"  checked>Payments&nbsp;
                                                                        <%End IF %>
                                                                        </td>
                                                                        <td class="FieldCell"  align=right>
                                                                        &nbsp;&nbsp;&nbsp;&nbsp;Entry No&nbsp;&nbsp;&nbsp;
													                    <span class="DataOnly" id="spEntryNo"><b>1&nbsp;</b></span></td>
                                                                    </tr>
                                                            </table>
                                                   </tr>
                                                   <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Head</td>
                                                    <td class="FieldCell">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
															<option value="A">Select Account Head</option>
															<%
																dim iHeadCount
															 	'iHeadCount=popFrequentHead(sOrgId,"01",sBookCode)
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
                                                    <td><span class="DataOnly" id="spAccHead"></span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Pay to / Received from</td>
                                                    <td class="FieldCell"> <input type="text" name="txtPayTo" size="40" class="FormElem" maxlength="50">
                                                    &nbsp; <a href="javascript:SelMisParty()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a></td>
                                                        </tr>
                                                        <tr>
                                                    <td width="139" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td width="50%" class="FieldCellSub">Narration</td>
                                                          <td width="50%" class="FieldCellSub">
<%

IF Cstr(sBookCode) = "" Then
	sBookCode = 0
End IF
sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
	" OUDefinitionID='"&sOrgId&"'and BookCode='01' and BookNumber="&sBookCode
'Response.write "<textarea>"& sQuery &"</textarea>"


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
                                                    <a href="javascript:showNarration('01')"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Frequently Used Narrations"></a>
<%
end if
objRs.Close
%>
                                                           </td>
                                                        </tr>
                                                      </table>
                                                      &nbsp;</td>
                                                    <td class="FieldCell" valign="top"> <textarea rows="3" name="txtNarration" cols="50" class="FormElem" onKeyPress="ChkEnter()"></textarea> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell"> <input type="text" name="txtAmount" size="15" value="0.00" style="text-align:right" maxlength="13" class="FormElem" onblur="popAddAmount();TDSAmount()"> </td><!--popAddAmount()-->
                                                        </tr>
                                                        <tr>
                                                        <td colspan=2>
                                                        <div id="DisCCANL" class=frmBody style="height:1px; visibility: hidden;">
	                                                                <table cellpadding="0" cellspacing="0" >
		                                                                <tr>
			                                                                <td class=MiddlePack colspan="4"> </td>
		                                                                </tr>
		                                                                <tr>
		                                                                    <td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"></td>
			                                                                <td class=FieldCell>
				                                                                <DIV class=frmBody id="DisCost" style="width:260;height:100;">
					                                                                <table border="0" id="tblCost" cellspacing="1" class="ExcelTable">
						                                                                <tr>
							                                                                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								                                                                <td class="ExcelHeaderCell" align="center" width="150">
								                                                                Cost Center Head
								                                                                <img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Click Here to add Cost Center or Analytical Head" onclick="PopCCAH()">
								                                                                </td>
								                                                                <td class="ExcelHeaderCell" align="center">Ratio</td>
								                                                                <td class="ExcelHeaderCell" align="center">Amount</td>
						                                                                 </tr>
					                                                                </table>
				                                                                </div><!--End of CostCenter Display Division -->
			                                                                </td>
			                                                                <td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"></td>
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
			                                                                <td class=MiddlePack  colspan="4"></td>
		                                                                </tr>
	                                                                </table>
                                                                </div> <!--End of CCANAL Display Division -->
                                                            </td>
                                                          </tr>
                                                         <tr>
                                                    <td class="FieldCellSub" >Select TDS Group</td>
                                                    <td class="FieldCell" width="591">
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
                                                    &nbsp; % On Amount &nbsp;
                                                    <input type="text" name="txtTdsAmount" Value="" size="15" style="text-align:right" maxlength="13" class="FormElemRead" readonly >
                                                    <a href="javascript:TDSCalc()"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="TDS Group Selection" width="10" height="11"></a>
                                                     <!--input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled-->
                                                     &nbsp;&nbsp;<input type="Button" value="Add Entry" name="btnAdd" onClick="AddNew()" class="AddButton">
                                                    </td>
                                                        </tr>
<tr>
								<td colspan=2 width="100%" align=center>
                                    <DIV class=frmBody id="DisVoucher" style="width:98%; visibility:hidden; height:1px;">
	                                    <table border="0" cellspacing="1px" id="tblVoucher" class="ExcelTable" style="width:98%;" >
	                                    <tr>
		                                    <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		                                    <td class="ExcelHeaderCell" align="center" width="25"></td>
		                                    <td class="ExcelHeaderCell" align="center" width="25"></td>
		                                    <!--<td class="ExcelHeaderCell" align="center">AU</td>-->
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
							</tr>                                                         <!--tr>
															<td class="FieldCellSub" width="133">Approval</td>
															<td class="FieldCell" width="591">
															<input type="radio" value="Y" checked name="optApprove" class="FormElem">
															Yes&nbsp;&nbsp;
															<input type="radio" value="N" name="optApprove" class="FormElem"> No </td>
														</tr-->
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5px">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8px">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
                           <tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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
								<td align="center" width="5px" class="ClearPixel">
								</td>
								<td >
<DIV class=frmBody id="Disaddtional" style="height:1px; visibility: hidden;">
	<DIV class=frmBody id="DisPayable" style="width: 585px; visibility: hidden; height:1px;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="565px">
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
								<td align="center" class="ClearPixel" width="5px">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8px">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell" align="center">
													<!--<input type="Button" value="Update Entry" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButtonX" >-->
													<!--<input type="Button" value="Delete Entry" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButtonX" >-->
													<input type="button" value="Save" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
													<!--input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouCABookSelection.asp')" class="ActionButton" -->
													<!--input type="button" value="Delete Voucher" name="btnDelVou" onClick="DelVouch()" class="ActionButtonX" disabled-->
													<!--input type="button" value="Print" name="btnPrnVou" onClick="PrnVouch()" class="ActionButtonX" -->
													<input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('CASHVOUCHERS.ASP')" class="ActionButtonX">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel" width="5px">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
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