<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCAEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January 06,2003
	'Modified By				:	Manohar Prabhu.R
	'Modified On				:	Sep 28, 2004
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

dim sAccount,sAddtional,iSno
dim dTotal
dim sVoucDate,iBookCode,sPayTo,sUserId
sUserId = getUserID()
'XML DOM Variables
Dim oDOM,nodHeader,Root,newElem,newElem1,newElem2

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("horgName")
sBookCode=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sVouType=Request.Form("selVouType")
sTransNo=Request.Form("hTransno")
iVouNo=Request.Form("txtVouNo")
bOtherUnits=Request.Form("hBookOtherUnit")
iBookAccHead=Request.Form("hBookAccHead")
bActionFlag=Request.Form("hActionFlag")

oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS Cash Voucher</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>

<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="<%=sBookCode%>" BookName="<%=sBookName%>" CRDR="<%=sVouType%>" VouDate="" BookAcchead="<%=iBookAccHead%>" Approver=""/></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="0" TDSElgi="0" TdsPercentage="0" /></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<SCRIPT ID=oButtonScript FOR="ctlDate" EVENT="onblur()" LANGUAGE="vbscript">
DisplayBalamt()
</SCRIPT>

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

Function DisplayBalamt()
	Dim objHttp,sTemp,sRetVal,Temparr
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	sTemp = document.formname.hOrgId.Value
	sTemp = sTemp&":"&iBookAcchead
	sTemp = sTemp&":"&document.formname.ctlDate.GetDate

	objhttp.Open "GET","GetDayOpenByDate.asp?sValue="&sTemp , false
	objhttp.send
	sRetval = objHttp.responseText
	Temparr = Split(sRetVal,"*")
	IF UBound(Temparr) = 1 Then
		'MsgBox Temparr(0)
		'MsgBox Temparr(1)
		IF CDbl(Temparr(0)) >= 0 Then
			document.all.spBookBal.innerHtml = FormatNumber(abs(Temparr(0)),2,,,0) &" Dr "
		Else
			document.all.spBookBal.innerHtml = FormatNumber(abs(Temparr(0)),2,,,0) &" Cr "
		End IF

		IF CDbl(Temparr(1)) >= 0 Then
			document.all.spCurrBal.innerHtml = FormatNumber(abs(Temparr(1)),2,,,0) &" Dr "
		Else
			document.all.spCurrBal.innerHtml = FormatNumber(abs(Temparr(1)),2,,,0) &" Cr "
		End IF
	End IF
End Function

FUNCTION popAccHead()
	dim iHeadCount

	iUnitNo=document.formname.selAccUnitId.value
	iHeadCount=cint(document.formname.hHeadCount.value)
	iBkNo=document.formname.hBookcode.value
	document.formname.selAccHead.selectedIndex=0

	for iCounter=1 to iHeadCount
		document.formname.selAccHead.remove(1)
	next

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	'objhttp.Open "GET","XMLGetOrgFreqHeads.asp?BkCode=01&BkNo="&iBkNo&"&orgID=" & iUnitNo , false
	'objhttp.send

	'if objhttp.responseXML.xml <> "" then
	'	OutData.loadXML objhttp.responseXML.xml
	'	Set Root = OutData.documentElement
	'	iCounter=1
'
	'	For Each HeaderNode In Root.childNodes

	'		set oText1 = document.createElement("<Option>" )
	'			oText1.Text = HeaderNode.text
	'			oText1.Value = HeaderNode.Attributes.Item(0).nodeValue

	'		document.formname.selAccHead.add oText1,iCounter
	'		iCounter=CDbl(iCounter)+1
	'	next
	'		document.formname.hHeadCount.value=CDbl(iCounter)-1
	'		iHeadCount=CDbl(iCounter)+1
	'else
		document.formname.hHeadCount.value=0
		iHeadCount=2
	'end if

	'for iCounter=iHeadCount+1 to document.formname.selAccHead.length
	'	document.formname.selAccHead.remove(iHeadCount)
	'next

	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		iCounter=document.formname.selAccHead.length
		For Each HeaderNode In Root.childNodes
			set oText1 = document.createElement("<Option>" )
				oText1.Text = HeaderNode.text
				oText1.Value = HeaderNode.Attributes.Item(0).nodeValue

			document.formname.selAccHead.add oText1,iCounter
			iCounter=CDbl(iCounter)+1
		next
	end if

END FUNCTION

FUNCTION selAccountHead(objAcc)
	DisplayBalamt
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
					document.formname.hTdsElgi.value = sTemp(4)

					IF CStr(sTemp(4)) = "0" Then
						document.formname.txtTdsAmount.disabled = True
						document.formname.txtTdsper.disabled = True
					Else
						document.formname.txtTdsAmount.disabled = False
						document.formname.txtTdsper.disabled = False
					End IF

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
				IF iEntryNo > "0" Then
					document.formname.txtNarration.focus()
					document.formname.txtPayTo.readOnly = True
				Else
					document.formname.txtPayto.focus()
					document.formname.txtPayTo.readOnly = False
				End IF
			end if	'END OF ACCOUNTING UNIT SELECTED OR NOT
		else

			sOrgId=document.formname.hOrgId.value
			if objAcc.selectedIndex <= iHeadCount then
					sTemp=Split(objAcc.value,"?")
					document.formname.hTdsElgi.value = sTemp(4)
					IF CStr(sTemp(4)) = "0" Then
						document.formname.txtTdsAmount.disabled = True
						document.formname.txtTdsper.disabled = True
					Else
						document.formname.txtTdsAmount.disabled = False
						document.formname.txtTdsper.disabled = False
					End IF
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

			IF iEntryNo > 0 Then
				document.formname.txtNarration.focus()
				document.formname.txtPayTo.readOnly = True
			Else
				document.formname.txtPayTo.focus
				document.formname.txtPayTo.readOnly = False
			End IF
		end if	'END OF BOOK HAS OTHER UNIT TRANSCATION OR NOT CHECK
	End if 'END OF IF ANY ACCOUNT HEAD SELECTED CHECK

END FUNCTION
'---------------------END OF FUNCTION SELACCOUNTHEAD----------------------
FUNCTION showPartyHead(sOrgId,sPartyType,sVouType)

dim sPartyCode,bRecivable,bPayable
dim sDocNo,sInvNo,sInvDate,sAmtRec,sAmtRecd
dim nodAccHead,nodPayRec,nodCC,iSno
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim iPayRecCount,sExp,TempNode,iSelPayRec
Dim sAmtToAdjust

set objhttp = CreateObject("Microsoft.XMLHTTP")

OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
arrTemp = split(OutValue,":")

while UBound(arrTemp) = 0
	OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")
wend

if UBound(arrTemp) <= 1 then
	document.formname.selAccHead.selectedIndex = 0
	document.formname.selAccHead.focus()
	exit function
End IF

sRetValue = OutValue
sTemp = Split(sRetValue,":")
sParTy = sTemp(4)
sParSubType = sTemp(3)
sParCode = sTemp(1)
sPartyName = sTemp(0)

objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
objhttp.send



IF objhttp.responseText <> "" Then
	sRetVal2 = objhttp.responseText
	GetPartyHeadXml sParCode,sPartyName,sRetVal2
End IF
Set nodAccHead = AccHeadData.documentElement



'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		bVouFlag=true
		sPartyCode=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
		HeaderNode.Attributes.Item(0).nodeValue=sPartyCode
		bPayable=HeaderNode.Attributes.Item(1).nodeValue
		bRecivable=HeaderNode.Attributes.Item(2).nodeValue

		window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue&"&nbsp;"
		IF document.formname.txtPayTo.value = "" Then
			document.formname.txtPayto.value=HeaderNode.Attributes.Item(3).nodeValue
		End IF
		EntryRoot.appendChild HeaderNode
		sTransFlag="A"
	next
	if (cint(bRecivable)>=1 and sVouType="D") or (cint(bPayable)>=1 and sVouType="C") then
	sPartyCode = Replace(sPartyCode,"&","and")
	'MsgBox sPartyCode
	'If Selected Party Has Payable or Receiavable
		Set nodPayRec = showModalDialog("PayRecSelection.asp?orgId="+sOrgId+"&ParCode="+sPartyCode&"&Type="&sVouType,"","")

		sExp = "//RecCount"
		Set TempNode = nodPayRec.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			iPayRecCount = TempNode.Item(0).Attributes.Item(0).nodeValue
		End IF

		sExp = "//Doc"
		Set TempNode = nodPayRec.selectNodes(sExp)
		iSelPayRec =  TempNode.length

		document.formname.hSelPayRecCount.value = iSelPayRec
		document.formname.hPayRecCount.value = iPayRecCount

		if nodPayRec.Attributes.Item(0).nodeValue=1 then
			'Set the Additional Display Layer Visible
			For Each HeaderNode In nodPayRec.childNodes
					EntryRoot.appendChild HeaderNode
					if HeaderNode.hasChildNodes then
						'If user has Selected Documnets
						iSno=1
						setPayableDisplay 1
						ClearTable "tblPayable",2,1
						for each  nodCC in HeaderNode.childNodes
							sDocNo=nodCC.Attributes.getNamedItem("No").Value
							sInvNo=nodCC.Attributes.getNamedItem("InvNo").Value
							sInvDate=nodCC.Attributes.getNamedItem("InvDate").Value
							sTransAmount=nodCC.Attributes.getNamedItem("TransAmount").Value
							sAmtAdjusted=nodCC.Attributes.getNamedItem("AmtAdjusted").Value
							sAmtToAccount=nodCC.Attributes.getNamedItem("AmtToAccount").Value

							sTransAmount = CDbl(sTransAmount)
							sAmtAdjusted = CDbl(sAmtAdjusted)
							sAmtToAccount = CDbl(sAmtToAccount)

							sAmtToAdjust = Cdbl(sTransAmount - sAmtAdjusted - sAmtToAccount)

							set oRow = document.all.tblPayable.insertRow(iSno+1)
							InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
							InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sTransAmount,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sAmtAdjusted,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sAmtToAccount,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",FormatNumber(sAmtToAdjust,2,,,0),"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo),"0","ExcelInputCell","","",12,10,0,0,"style=""text-align:right"""
							iSno=iSno+1

						next
					end if 'End of Check Documnet Node
			next	'End of Processing PayRec Node
		else
			'User Has canceled Documnet Selection
			'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
     		setPayableDisplay 0
		end if	'End of Documnet has Childs Check
    else
		'Selected Head has no Documnets
		'Set the Additional Layer Display Layer Hidden
		setPayableDisplay 0
	end if	'End of Party Has Payable Or Recivables
else
	'User canceled Party Head Selection
	window.spAccHead.innerHTML=""
	document.formname.txtPayto.value=""
	'Set the Additional Layer Display Layer Hidden
	setPayableDisplay 0
end if 'End of Party Head Processing

set nodAccHead=nothing
set nodPayRec=nothing
set nodCC=nothing
END FUNCTION
'---------------------END OF FUNCTION showPartyHead--------------------------
function showGLHead(sOrgId)
dim iAccCode,bAnal,bCostCenter
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,sRetVal,arrTemp,sTemp2,sTdsElgi,sTempVal
iBookNo=document.formname.hBookcode.value

OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
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

'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")

if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		bVouFlag=true
		iAccCode=HeaderNode.Attributes.Item(0).nodeValue
		bAnal=HeaderNode.Attributes.Item(1).nodeValue
		bCostCenter=HeaderNode.Attributes.Item(2).nodeValue
		sTransFlag=HeaderNode.Attributes.Item(5).nodeValue

		window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue&"&nbsp;"

		'document.formname.txtPayto.value=HeaderNode.Attributes.Item(3).nodeValue
		EntryRoot.appendChild HeaderNode
	next
	if cint(bCostCenter)=1 or cint(bAnal)=1 then
	'If Selected GL Account Head has Cost Center
		Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+iAccCode,"","")
		if nodCCAnly.Attributes.Item(0).nodeValue=1 then
			'Set the Additional and CCANAL Display Layer Visible
			setADDDisplay 1
			For Each HeaderNode In nodCCAnly.childNodes
				if 	HeaderNode.nodeName="CostCenter" then
					EntryRoot.appendChild HeaderNode
					popCostCenter HeaderNode
				end if 'End of Check for Cost Center Node
				if 	HeaderNode.nodeName="Analytical" then
					EntryRoot.appendChild HeaderNode
					popAnalytical(HeaderNode)
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
else
	'User canceled Account Head Selection
	window.spAccHead.innerHTML=""
	document.formname.txtPayto.value=""
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if 'End of GL Head Processing

set nodAccHead=nothing
set nodCCAnly=nothing
set nodCC=nothing
End function
'---------------------End Of Function showGLHead--------------------------
Function AddEntry(bFlag)
dim iCode,dRatio,dAmount,sExp,TempNode,iCounter,iTdsPer,sCheckExp,CheckNode

' New Validation for check blank data - included on 02/04/2004
if bFlag = "S" then
	if Trim(document.formname.txtAmount.value) = "0.00" then
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			Exit Function
		End IF
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

	IF CStr(bFlag) <> "S" Then
		document.formname.hTotType.value = "A"
	End IF

	sCheckExp = "//Entry[@No="&iEntryNo&" and @TdsAmount]"
	Set CheckNode = VouRoot.selectNodes(sCheckExp)

	'This gets checks for all PayTo values in the entry node and updates the same.
	'if bFlag="U" then
		sExp = "//Entry"
		Set Tempnode = VouRoot.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			For iCounter = 0 To TempNode.length - 1
				Tempnode.Item(iCounter).Attributes.getNamedItem("Payto").value = document.formname.txtPayto.value
			Next
		End IF
		document.formname.hPayTo.value = document.formname.txtPayto.value
	'End IF

	VouRoot.Attributes.getNamedItem("VouDate").Value=document.formname.ctlDate.getdate

	if document.formname.selCRDR(0).checked then
		EntryRoot.Attributes.getNamedItem("CRDR").Value=document.formname.selCRDR(0).value
	else
		EntryRoot.Attributes.getNamedItem("CRDR").Value=document.formname.selCRDR(1).value
	end if


	if EntryRoot.Attributes.getNamedItem("CRDR").Value ="C" then
		dTotal=dTotal-CDbl(document.formname.txtAmount.value)
	else
		dTotal=dTotal+CDbl(document.formname.txtAmount.value)
	end if




	EntryRoot.Attributes.getNamedItem("Payto").Value=document.formname.txtPayTo.value
	EntryRoot.Attributes.getNamedItem("Amount").Value=document.formname.txtAmount.value
	IF CheckNode.length <> 0 Then
		EntryRoot.Attributes.getNamedItem("TdsAmount").Value=document.formname.txtTdsAmount.value
		EntryRoot.Attributes.getNamedItem("TdsPercentage").Value = document.formname.txtTdsper.value
		EntryRoot.Attributes.getNamedItem("TDSElgi").Value=document.formname.hTDSElgi.value
	Else
		EntryRoot.setAttribute "TdsAmount",document.formname.txtTdsAmount.value
		EntryRoot.setAttribute "TdsPercentage",document.formname.txtTdsper.value
		EntryRoot.setAttribute "TDSElgi",document.formname.hTDSElgi.value
	End IF

	if document.formname.hOtherUnitFlag.value=1 then
		EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.selAccUnitId.value
		EntryRoot.Attributes.getNamedItem("AccName").Value=document.formname.selAccUnitId.options(document.formname.selAccUnitId.selectedIndex).text
	else
		EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.hOrgId.value
		EntryRoot.Attributes.getNamedItem("AccName").Value=document.formname.hOrgName.value
	end if

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
	Else
		VouRoot.appendChild EntryRoot
	End IF
'====================================================================================

	if bFlag="S" then
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			document.formname.txtPayTo.readOnly = True
			Exit Function
		End IF
	else
		DisplayVoucher
		clearXML()
		setADDDisplay 0
		setPayableDisplay 0

		'document.formname.txtPayTo.value=""

		document.formname.selCRDR(0).disabled=false
		document.formname.selCRDR(1).disabled=false
		document.formname.selAccHead.selectedIndex = 0
		document.formname.txtNarration.value = ""
		document.formname.txtAmount.value = "0.00"
		document.formname.txtTdsAmount.value = "0.00"
		document.formname.txtTdsper.value = "0.00"
		document.formname.txtTdsAmount.disabled = True
		document.formname.txtTdsper.disabled = True

		'document.formname.reset

		sExp = "//Entry"
		Set Tempnode = VouRoot.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			document.formname.hPayTo.value = Tempnode.Item(0).Attributes.getNamedItem("Payto").value
		End IF

		document.formname.btnadd.disabled=false
		document.formname.btnnext.disabled=false
		document.formname.btnupdate.disabled=true
		document.formname.btndel.disabled=true
		bEditFlag=true
		bVouFlag=false
	end if
else
	if bFlag="S" then
		IF CheckVouStat() Then
			SaveXML
			Exit Function
		Else
			Exit Function
		End IF
	End if
end if

End Function

'---------------------End Of Function AddEntry----------------------------
Function EditEntry(iVouEntryNo)
Dim sCheckExp,CheckNode
if bEditFlag then
setADDDisplay 0
setPayableDisplay 0
bVouFlag=true
window.spEntryNo.innerHTML=iVouEntryNo
sCheckExp = "//Entry[@TdsAmount]"
Set CheckNode = VouRoot.selectNodes(sCheckExp)



	For Each EntryNode in VouRoot.childNodes
		if EntryNode.Attributes.Item(0).nodeValue=iVouEntryNo then
			document.formname.txtAmount.value=EntryNode.Attributes.Item(3).nodeValue
			if EntryNode.Attributes.Item(1).nodeValue ="C" then
					document.formname.selCRDR(0).checked=true
			else
					document.formname.selCRDR(1).checked=true
			end if

			sAccUnit=EntryNode.Attributes.Item(5).nodeValue

			document.formname.txtPayTo.value = EntryNode.Attributes.Item(2).nodeValue
			document.formname.txtPayTo.readOnly = False
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

			if document.formname.hOtherUnitFlag.value=1 then
				for i=1 to document.formname.selAccUnitId.length-1
					if document.formname.selAccUnitId.options(i).value =EntryNode.Attributes.Item(4).nodeValue then
						document.formname.selAccUnitId.selectedIndex=i
					end if
				next
				popAccHead
			end if
			sAddtional=""

			For Each HeaderNode in EntryNode.childNodes
				if HeaderNode.nodeName="AccHead" then
					if HeaderNode.Attributes.getNamedItem("Type").value="P" then
						SelectHead HeaderNode.Attributes.getNamedItem("No").value,"P",document.formname.selAccHead,CInt(document.formname.hHeadCount.value)
					else
						SelectHead HeaderNode.Attributes.getNamedItem("No").value,"G",document.formname.selAccHead,CInt(document.formname.hHeadCount.value)
					end if
					'document.formname.txtPayTo.value=HeaderNode.Attributes.Item(3).nodeValue
					window.spAccHead.innerHTML=HeaderNode.Attributes.Item(3).nodeValue
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
	document.formname.btnnext.disabled=true
	document.formname.btnupdate.disabled=false
	document.formname.btndel.disabled=false
	bEditFlag=false
	bSavFlag=true
	document.formname.txtPayTo.readOnly = True
end if
End Function
'---------------------End Of Function EditEntry----------------------------
Function DelEntry()
	clearXML
	setADDDisplay 0
	setPayableDisplay 0
	DisplayVoucher

	document.formname.txtPayTo.value=""
	window.spEntryNo.innerHTML=iEntryNo

	document.formname.selCRDR(0).disabled=false
	document.formname.selCRDR(1).disabled=false

	document.formname.reset

	document.formname.btnadd.disabled=false
	document.formname.btnnext.disabled=false
	document.formname.btnupdate.disabled=true
	document.formname.btndel.disabled=true
	bVouFlag=false
	bEditFlag=true
	bSavFlag=true
End Function
'---------------------End Of Function DelEntry----------------------------
Function DisplayVoucher()
dim sNarration,sAccount,sAddtional,iSno,sAmount,sCheckExp,CheckNode
dim dTotal,sAccUnit,iTdsAmount,iTdsTotAmount,iTdsPer

set VouRoot=VoucherData.documentElement

window.DisVoucher.style.height="200px"
window.DisVoucher.style.visibility="visible"

ClearTable "tblVoucher",1,1
dTotal=0

iEntryNo=0
icounter = 1

sDate=VouRoot.Attributes.Item(5).nodeValue
document.formname.ctlDate.setDate=sDate
For Each EntryNode in VouRoot.childNodes

	iEntryNo=cint(iEntryNo)+1

	sCheckExp = "//Entry[@No="&iEntryNo&" and @TdsAmount]"
	Set CheckNode = VouRoot.selectNodes(sCheckExp)

	EntryNode.Attributes.Item(0).nodeValue=iEntryNo
	sAmount=FormatNumber(EntryNode.Attributes.Item(3).nodeValue,2,,,0) & "&nbsp;" & EntryNode.Attributes.Item(1).nodeValue
	IF CStr(EntryNode.Attributes.Item(1).nodeValue) = "C" Then
		dTotal=dTotal-CDbl(EntryNode.Attributes.Item(3).nodeValue)
	Else
		dTotal=dTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
	End IF

	sAccUnit=EntryNode.Attributes.Item(5).nodeValue
	IF CheckNode.length <> 0 Then
		iTdsAmount = EntryNode.Attributes.Item(6).nodeValue
		iTdsPer = EntryNode.Attributes.Item(8).nodeValue
	Else
		iTdsAmount = 0
		iTdsPer = 0
	End IF
	iTdsTotAmount = iTdsTotAmount + CDbl(iTdsAmount)
	document.formname.hPayTo.value = EntryNode.Attributes.Item(2).nodeValue

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

	iTdsAmount = FormatNumber(iTdsAmount,2,,,0)
	iTdsPer = FormatNumber(iTdsPer,2,,,0)

	set oRow = document.all.tblVoucher.insertRow(iEntryNo)
	InsertCell oRow,1,"",icounter,"ExcelSerial","Center","top",0,0,0,0,""
	InsertCell oRow,1,"","<a href=""javascript:EditEntry('"&iEntryNo&"')"" class=""ExcelDisplayCell""><b>Edit</b></a>","ExcelDisplayCell","Center","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",iTdsAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",iTdsPer,"ExcelDisplayCell","right","top",0,0,0,0,""

	icounter = icounter + 1

next'End of Voucher Node Loop
	dTotal = FormatNumber(dTotal,2,,,0)
	set oRow = document.all.tblVoucher.insertRow(iEntryNo+1)
	InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,6,0,""
	InsertCell oRow,1,"","<input type=""text"" name=""txtTotalAmt"" value="&dTotal&" size=""13"" class=""Formelemread"" style=""text-align:right"" >","ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"iTotalAmt",FormatNumber(dTotal,2,,,0) ,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",FormatNumber(iTdsTotAmount,2,,,0) ,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","" ,"ExcelDisplayCell","right","top",0,0,0,0,""

	window.spAccHead.innerHTML=""
	window.spEntryNo.innerHTML=iEntryNo+1

End Function

Function SaveXML()
	if bSavFlag then
		IF CheckApp() Then
			set objhttp = CreateObject("Microsoft.XMLHTTP")
			objhttp.Open "POST","XMLSave.asp?Name=Voucher Entry&Mod=CA", false
			objhttp.send VoucherData.XMLDocument
			if objhttp.responseText <> "" then
				Msgbox(objhttp.responseText)
			else
				document.formname.btnNext.disabled = True
				document.formname.submit()
			end if
		End IF
	end if
End Function
'---------------------End Of Function SaveXML-----------------------------
FUNCTION  checkFileds()
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
	if CDbl(document.formname.txtAmount.value) > CDbl(dTransLimit) then
		select case sTransFlag
			case "W"
					MsgBox "Amount is greater than the amount limit",,"Warning"
			case "R"
					MsgBox "Amount should be less than "&dTransLimit
					checkFileds=false
					exit Function
		end select
	end if

	for each HeaderNode in EntryRoot.childNodes
		if HeaderNode.nodeName="PayRec" then
			dAmount=CDbl(document.formname.txtAmount.value)
			dTotalAmtAdjust=0
			iCounter=1
				for each  nodANL in HeaderNode.childNodes
					iCode=nodANL.Attributes.getNamedItem("No").Value
					dTransAmount=nodANL.Attributes.getNamedItem("TransAmount").Value
					dAmtAdjusted=nodANL.Attributes.getNamedItem("AmtAdjusted").Value
					dAmtToAccount=nodANL.Attributes.getNamedItem("AmtToAccount").Value
					sAdjType = nodANL.Attributes.getNamedItem("AdjType").Value

					IF CStr(sAdjType) = "I" Then
						dAmtAdjust=CDbl(dTransAmount)-(CDbl(dAmtAdjusted)+CDbl(dAmtToAccount))
					Else
						dAmtAdjust=CDbl(dTransAmount)-CDbl(dAmtAdjusted)
					End IF
					dTotal=eval("document.formname.txtDocAmount"&iCode).value
					if  CDbl(dTotal)>CDbl(dAmtAdjust) then
						MsgBox """To Adjust Amount"" should be less than ""Document Amount-(Adjusted +To Account)"""
						eval("document.formname.txtDocAmount"&iCode).focus
						checkFileds=false
						exit Function
					else
						dTotalAmtAdjust=CDbl(dTotalAmtAdjust)+CDbl(dTotal)
					end if
				next
				if  CDbl(dTotalAmtAdjust)>CDbl(dAmount) then
					MsgBox "Total of ""To Adjust Amount"" should be less than ""Voucher Amount"""
					checkFileds=false
					exit Function
				end if
		end if 'End of Check for PayRec Node
	next
	checkFileds=true
END FUNCTION
'---------------------END OF FUNCTION CHECKFILEDS-------------------------
Function CancelAction(sPage)
	document.formname.action=sPage
	document.formname.submit
end Function
'---------------------End Of Function ActionCancel----------------------------

Function AddNewParty()
	OutValue = showModalDialog("MisParCreate.asp?"&OutValue,"","dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'MsgBox OutValue
	document.formname.txtPayTo.value = OutValue
End Function

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

Function CheckVouStat()
	Dim iTotal,sVouType,sVouName,sAltVouName,sStatus,sCurrDate,sTempCurr,dCurrentBal
	Dim iSelPayRec,iTotPayRec,iRetVal

	iSelPayRec = document.formname.hSelPayRecCount.value
	iTotPayRec = document.formname.hPayRecCount.value
	sCurrDate = document.formname.hCurrDate.value
	sVouType = document.formname.hVouCRDR.value
	iTotal = CheckVouAmount()

	IF iEntryNo = 1  and DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		CheckVouStat = false
		Exit Function
	Else
		CheckVouStat = True
	End IF

	IF CDbl(iTotal) < 0 Then
		IF CStr(sVouType) = "C" Then
			MsgBox "Total Voucher Amount is more than the Payment Amount"
		Else
			MsgBox "Total Voucher Amount is more than the Receipt Amount"
		End IF
		CheckVouStat = False
		sStatus = "T"
		Exit Function
	End IF


	IF CDbl(iTotal) = 0 Then
		MsgBox "Total Voucher amount should be More than Zero "
		CheckVouStat = False
		Exit Function
	Else
		sStatus = "F"
	End IF


	iSelPayRec = CDbl(iSelPayRec)
	iTotPayRec = CDbl(iTotPayRec)

	IF iTotPayRec <> 0 and iSelPayRec = 0 Then
		iRetVal = MsgBox("Adjustment is Not made for the Party!!, Continue Without Adjustments? ",4,"Warning")
	End IF

	IF CStr(iRetVal) = "7" Then
		CheckVouStat = False
		iEntryNo = Cdbl(iEntryNo - 1)
		Exit Function
	End IF

	IF Not CheckAdjVal(iTotal) Then
		iRetVal = MsgBox("Payment Amount is made more than the bill value!!, Continue?  ",4,"Warning")
	End IF

	IF CStr(iRetVal) = "7" Then
		CheckVouStat = False
		iEntryNo = Cdbl(iEntryNo - 1)
		Exit Function
	End IF



	IF CStr(document.formname.hVouCRDR.value) = "C" Then

		sTempCurr = Split(document.all.spCurrBal.innerText," ")
		dCurrentBal = Trim(sTempCurr(0))

		dCurrentBal = CDbl(dCurrentBal)
		iTotal = Trim(iTotal)
		iTotal = CDbl(iTotal)
		IF iTotal > dCurrentBal Then
			MsgBox "Voucher Amount is Greater than the Current Balance "
			'document.formname.hTotalAmt.value = CDbl(document.formname.hTotalAmt.value) - CDbl(document.formname.txtAmount.value)
			Exit Function
		End IF

		'IF CDbl(iTotal) > 20000 Then
		'	MsgBox "Payment Voucher Amount should not be greater than 20,000 "
		'	CheckVouStat = false
		'	Exit Function
		'End IF

	End IF

	IF DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
		MsgBox "Voucher Date Should be Less than the System Date "
		CheckVouStat = false
		Exit Function
	Else
		CheckVouStat = True
	End IF

End Function

Function CheckAdjVal(iTotal)
	Dim sExp,TempNode,iAdjVal,iCtr
	sExp = "//PayRec/Doc"
	iAdjVal = 0

	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.Length - 1
			iAdjVal = iAdjVal + CDbl(TempNode.item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value)
		Next
	Else
		CheckAdjVal = True
		Exit Function
	End IF

	IF CDbl(iTotal) > CDbl(iAdjVal) Then
		CheckAdjVal = False
		Exit Function
	Else
		CheckAdjVal = True
		Exit Function
	End IF
End Function

Function CheckVouAmount()
	Dim sExp,TempNode,iCount,iRecpTotal,iPayTotal,iRetValue

	iRecpTotal = 0
	iPayTotal = 0

	'Taking values for the Receipt Amount
	sExp = "//Entry[@CRDR=""C""]"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCount = 0 To TempNode.length - 1
			iRecpTotal = iRecpTotal + Cdbl(TempNode.Item(iCount).Attributes.getNamedItem("Amount").Value)
		Next
	End IF

	'Taking values for the Payment Amount
	sExp = "//Entry[@CRDR=""D""]"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCount = 0 To TempNode.length - 1
			iPayTotal = iPayTotal + Cdbl(TempNode.Item(iCount).Attributes.getNamedItem("Amount").Value)
		Next
	End IF


	'if the Voucher type is Payment then Payment - Receipt else Receipt - payment

	IF CStr(document.formname.hVouCRDR.Value) = "D" Then
		iRetValue = CDbl(iRecpTotal) - CDbl(iPayTotal)
	Else
		iRetValue = CDbl(iPayTotal) - CDbl(iRecpTotal)
	End IF

	CheckVouAmount = iRetValue

End Function


</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname" action="VouGenerate.asp">
<input type="hidden" name="hVouCode" value="01">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="CA">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=sBookCode%>">
<input type="hidden" name="hOtherUnitFlag" value="<%=bOtherUnits%>">
<input type="hidden" name="hActionFlag" value="<%=bActionFlag%>">
<input type="hidden" name="hTransNo" value="0">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hPayTo" value="">
<input type="hidden" name="hTDSElgi" value="0">
<input type="hidden" name="hTotalAmt" value="0">
<input type="hidden" name="hPayRecCount" value="0">
<input type="hidden" name="hSelPayRecCount" value="0">
<input type="hidden" name="hTotType" value="N">

<input type="hidden" name="hCurrDate" value="<%=Day(Date)&"/"&MonthName(Month(Date),True)&"/"&Year(Date)%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
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
                    <p align="center"><font face="Webdings" size="5">�</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Daywise Balance"><font face="Webdings" size="5">�</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: hand" Title="Voucher History">
                    <font face="Webdings" size="5">�</font>
                    </span>
                    </p>
                    </td>
                        </tr>
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
								<td align="center" colspan="3" height="8">
                                                      <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                                       <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Voucher
                                                            Date</td>
                                                          <td class="FieldCellSub" width="125">
                                                          <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>

                                                          </td>
                                                          <td class="FieldCellSub" width="80">Entry
                                                      Number</td>
                                                          <td class="FieldCellSub"><span class="DataOnly" id="spEntryNo">1&nbsp;</span></td>
                                                          <td class="FieldCellSub" width="100">
                                                    Book Balance</td>
                                                          <td class="FieldCellSub">
   <span class="DataOnly">
                                                            <%
                                                             dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
                                                             dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                             if dOpeningBal<0 then
                                                             %>
																<span class="DataOnly" id="spBookBal"><%Response.Write dOpeningBal*-1 &"&nbsp;Cr"%></span>
                                                             <%

															 else
															%>
																<span class="DataOnly" id="spBookBal"><%Response.Write dOpeningBal &"&nbsp;Dr"%></span>
															<%

															 end if
                                                            %></span>
   &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Entry
                                                            Type</td>
                                                          <td class="FieldCellSub" colspan="3">
											<%if sVouType="C" then%>
                                                            <input type=radio name="selCRDR" value="C" disabled>Receipts
                                                            <input type=radio name="selCRDR" value="D" checked>Payments
                                                            <%else%>
                                                            <input type=radio name="selCRDR" value="C" checked>Receipts
                                                            <input type=radio name="selCRDR" value="D" disabled >Payments&nbsp;
											<%end if%>
                                                            </td>
                                                          <td class="FieldCellSub" width="100">Current
                                                      Balance&nbsp;</td>
                                                          <td class="FieldCellSub"><span class="DataOnly" id="iCurrentBal">
                                                            <%
                                                             dOpeningBal =GetDayOpeningCreated(sOrgId,iBookAccHead,FormatDate(date+1))
                                                             dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                             if dOpeningBal<0 then
                                                            %>
																<span class="DataOnly" id="spCurrBal"><%Response.Write dOpeningBal*-1 &"&nbsp;Cr"%></span>
                                                            <%

															 else
															%>
																<span class="DataOnly" id="spCurrBal"><%Response.Write dOpeningBal &"&nbsp;Dr"%></span>
															<%

															 end if
                                                            %></span>
                                                            &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                      </table>
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
                                                    <td class="MiddlePack" colspan="2" width="139"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Unit</td>
                                                    <td class="FieldCell">
													<%	if bOtherUnits=1 then%>
                                                     <select size="1" name="selAccUnitId" onChange="popAccHead()" class="FormElem">
													<option value="A">Account Unit</option>
													<option value="<%=sOrgId%>" selected><%=sOrgName%></option>
															<%=popIUTUnits(sOrgId)%>
													</select>
													<%
														else
															Response.Write sOrgName
														end if
													%>


 </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Head</td>
                                                    <td class="FieldCell">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
															<option value="A">Select Account Head</option>
															<%
																dim iHeadCount
															 	'iHeadCount=popFrequentHead(sOrgId,"01",sBookCode)

															%>
																<option value="G">General Ledger</option>
															<%populatePartyType(sOrgId)%>
                                                    </select>
                                                    </td>
                                                    <input type="hidden" name="hHeadCount" value="0">

														</tr>
                                                    	<tr>
                                                    <td class="FieldCellSub" width="139"></td>
                                                    <td><span class="DataOnly" id="spAccHead"></span> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Pay to / Received from</td>
                                                    <td class="FieldCell"> <input type="text" name="txtPayTo" size="40" class="Formelem">
                                                    &nbsp; <a href="javascript:SelMisParty()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a></td>
                                                        </tr>
                                                        <tr>
                                                    <td width="139" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td width="50%" class="FieldCellSub">Narration</td>
                                                          <td width="50%" class="FieldCellSub">
<%

sQuery ="select count(NarrationDesc) from VwOrgFrequentNarration where "&_
	" OUDefinitionID='"&sOrgId&"'and BookCode='01' and BookNumber="&sBookCode

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
                                                    <td class="FieldCell" valign="top"> <textarea rows="3" name="txtNarration" cols="50" class="FormElem"></textarea> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell"> <input type="text" name="txtAmount" size="15" value="0.00" style="text-align:right" maxlength="13" class="Formelem" onblur="popAddAmount()"> </td>
                                                        </tr>
                                                         <tr>
                                                    <td class="FieldCellSub" width="133">Deduction @</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    &nbsp; % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
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
								<td class="ExcelHeaderCell" align="center" width="150">Cost Center Head</td>
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
                                                                <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
                                                                <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >
                                                                <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >
                                                                <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="CancelAction('VouCABookSelection.asp')" class="ActionButton" >
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