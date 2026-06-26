<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouDNOtherAmend.asp
	'Module Name				:	ACCOUNTS (Transcation Debit Note Amendment For Other Voucher Type)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 04, 2004
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
Dim sStr,TempNode,VouRoot,sFlag
Dim sVouDate,sVouUnit,sVouNumber,sVouAmt,sForAcc
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sAmdTy,iSelBook,sRetVal

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("horgName")
iBookNo=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sInvoiceNo=Request.Form("selInvoiceNo")
sVouchTy = Request.Form("selVoucherType")
sForAcc = Request.Form("hForAcc")

'sInvTemp = Split(sInvoiceNo,",")
sTransno = Request.Form("hTransNo")
sFlag = Request.Form("sFlag")
sAmdTy = Request("AmdType")
IF CStr(sAmdTy) = "A" Then
	sFlag = "True"
End IF
'Response.Write "Flag="& sFlag

sTemp = Split(sTransno,"-")
sTransno = sTemp(0)
'Response.Write sTransno
'Response.Write sTransno

'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
sRetVal = GetVouchXML(sTransNo)
oDOM.Load server.MapPath(sRetVal)

oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml")

Set VouRoot = ODom.documentElement
sStr = "//voucher"
Set TempNode = VouRoot.selectNodes(sStr)
IF TempNode.length <> 0 Then
	sOrgId = TempNode.Item(0).Attributes.getNamedItem("UnitNo").value
	sVouDate = TempNode.Item(0).Attributes.getNamedItem("VouDate").value
	sVouUnit = TempNode.Item(0).Attributes.getNamedItem("UnitName").value
	sVouNumber = TempNode.Item(0).Attributes.getNamedItem("VoucherNo").value
	sPartyName = TempNode.Item(0).Attributes.getNamedItem("PartyName").value
	iSelBook = TempNode.Item(0).Attributes.getNamedItem("BookNo").value
End IF

sStr = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr)
IF TempNode.length <> 0 Then
	sVouAmt = TempNode.Item(0).Attributes.getNamedItem("Amount").value
End IF


'sPartyName=Request.Form("txtPartyName")
'arrPartyCode=split(Request.Form("hPartyCode"),"?")

Set objRs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData">
<Root>

</Root>
</XML>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>

<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData" src="<%="../temp/transaction/Voucher AMD_DN_"&Session.SessionID&".xml"%>"></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName=""/>
</XML>

<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
iEntryNo=1
bVouFlag=false
bSavFlag=false
bEditFlag = True

set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement

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
OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=07&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
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
dim iCode,dRatio,dAmount,TempNode,sChk
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
	
	
	sChk = "F"
	IF CStr(bFlag) <> "U" Then
		Set newElem = EntryData.createElement("Narration")
		newElem.text= document.formname.txtNarration.value
		EntryRoot.appendChild newElem
	Else
		for each HeaderNode in EntryRoot.childNodes
			IF CStr(HeaderNode.nodeName) = "Narration" Then
				HeaderNode.text = document.formname.txtNarration.value
				sChk = "T"
			End IF
		Next
	End IF
	
	IF Cstr(sChk) = "F" Then
		Set newElem = EntryData.createElement("Narration")
		newElem.text= document.formname.txtNarration.value
		EntryRoot.appendChild newElem
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
	
	VouRoot.appendChild EntryRoot
	
	if bFlag="S" then
		SaveXML
	else
		DisplayVoucher
		clearXML()
		setADDDisplay 0
		document.formname.txtPayTo.value=""
		document.formname.reset
		document.formname.btnadd.disabled=false
		document.formname.btnnext.disabled=false
		document.formname.btnupdate.disabled=true
		document.formname.btndel.disabled=true
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
	dim dTotal,sAccUnit
	set VouRoot=VoucherData.documentElement
	window.DisVoucher.style.height="200px"
	window.DisVoucher.style.visibility="visible"

	ClearTable "tblVoucher",1,1
	dTotal=0

	iEntryNo=0

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

			set oRow = document.all.tblVoucher.insertRow(iEntryNo)
			InsertCell oRow,1,"",icounter,"ExcelSerial","Center","top",0,0,0,0,""
			InsertCell oRow,1,"","<a class=""ExcelDisplaylink"" href=""javascript:EditEntry('"&iEntryNo&"')"">Edit</a>","ExcelDisplayCell","Center","top",0,0,0,0,""
			InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
			InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
			InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
			InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
			InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""
		End IF	
	Next
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
	Dim sExp,TempNode
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	
	IF CStr(document.formname.hFlag.Value) = "True" Then
		sExp = "//voucher"
		Set TempNode = VouRoot.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("BookNo").Value = document.formname.selBook.value
			TempNode.Item(0).Attributes.getNamedItem("BookName").Value = document.formname.selBook.options(document.formname.selBook.selectedIndex).Text
			TempNode.Item(0).Attributes.getNamedItem("VouDate").Value = document.formname.ctlDate.GetDate()
		End IF
	End IF
	objhttp.Open "POST","XMLSave.asp?Mod=DN&Name=Voucher AMD", false
	objhttp.send VoucherData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		'MsgBox "OK "
		document.formname.btnNext.disabled = True
		document.formname.btnAdd.disabled = True
		If trim(document.formname.hFlag.value) = "True" then
			document.formname.action = "AmdAccDbNtGenerate.asp"
		End If
		
		document.formname.submit()
	end if
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

Function EditEntry(iVouEntryNo)
Dim iCounter
if bEditFlag then
	document.formname.hEntryNo.value = iVouEntryNo
	setADDDisplay 0
	'setPayableDisplay 0
	bVouFlag=true
	'MsgBox VouRoot.xml
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
	document.formname.btnadd.disabled=true
	document.formname.btnnext.disabled=true
	document.formname.btnupdate.disabled=false
	document.formname.btndel.disabled=false
	bEditFlag=false
	bSavFlag=true
end if
End Function

Function SetDate()
	Dim sFromYr,sToYr
	sFromYr = document.formname.hFromYr.Value
	sToYr = document.formname.hToYr.Value
	sFromYr = "01/04/"&Trim(sFromYr)
	sToYr = "31/03/"&sToYr
	document.formname.ctlDate.setMinDate() = sFromYr
	document.formname.ctlDate.setMaxDate() = sToYr
	DisplayVoucher()
End Function
'action="VouDNOTAmdGenerate.asp"
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">

<form method="POST" name="formname" action="VouDNOTAmdGenerate.asp">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sVouUnit%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hInvDate" value="<%=sVouDate%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hVouchTy" value="<%=sVouchTy%>">
<input type="hidden" name="hEntryNo" value="">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hFrmEdit" value="E">

<input type="hidden" name="hTransNo" value="<%=sTransno%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Debit Note Other
          Amendment
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
                        <!--tr>
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
                        </tr-->
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
                                    <table cellpadding="0" cellspacing="0" width="590">
                                    <tr>
										<td class="FieldCell" width="93">Unit</td>
										<td colspan="3"><span class="DataOnly"><%=sVouUnit%>&nbsp;</span></td>

	                                </tr>
	                                
									<tr>
										<td class="FieldCell" width="93">Agent Name</td>
										<td width="230"><span class="DataOnly"><%=sPartyName%>&nbsp;</span></td>
										
	                                </tr>
	                                
	                                <tr>
										<td class="FieldCell" width="100">Invoice No-Date</td>
										<td><span class="DataOnly"><%=sVouNumber%>-&nbsp;<%=sVouDate%> </span></td>
										<td class="FieldCell" width="100">Voucher Amount</td>
										<td><span class="DataOnly"><%=FormatNumber(sVouAmt,2,,,0)%> </span></td>
	                                </tr>
									<%IF CStr(sAmdTy) = "A" Then %>
										<tr>
											<td class="FieldCell" width="100">Book</td>
											<td class="FieldCell" colspan="3">
												<select size="1" name="selBook" class="FormElem">
												<%
													sQuery = "Select BookNumber,BookName From VwOrgBookNames Where  "&_
															 "OUDefinitionID = '"&sOrgId&"' And BookCode = '06' Order By BookName "
													With objRs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sQuery
														.ActiveConnection = Con
														.Open
													End With
													Set objRs.ActiveConnection = Nothing
													Do While Not objRs.EOF
														IF CStr(iSelBook) = CStr(objRs(0)) Then 
												%>
															<Option Value="<%=objRs(0)%>" Selected><%=objRs(1)%></Option>
												<%		Else%>
															<Option Value="<%=objRs(0)%>"><%=objRs(1)%></Option>
												<%		
														End IF
														objRs.MoveNext
													Loop
													objRs.Close
																			
												%>	
												 </select>
											 </td>
										</tr>  
									<%End IF%>	   

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
                                                    <td class="MiddlePack" colspan="5" width="139"></td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Accounting Head</td>
                                                    <td class="FieldCell">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selGLHead(this)">
															<option value="A">Select Account Head</option>
															<option value="G">General Ledger</option>

                                                    </select>
                                                    <input type="hidden" name="hHeadCount" value="1">
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
                                                    <input type="text" name="txtPayTo" size="40" class="Formelem"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139" valign="top">Narration</td>
                                                    <td class="FieldCell" colspan="2" valign="top">
                                                    
														<textarea rows="3" name="txtNarration" cols="50" class="FormElem" onKeyPress="ChkEnter()"></textarea> </td>
													
                                                    <td class="FieldCell" colspan="2" valign="middle">
 </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell" colspan="4"> 
                                                    
														<input type="text" name="txtAmount" value="0.00" size="15" style="text-align:right" class="Formelem"> </td>
													
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Approval</td>
                                                    <td class="FieldCell">
                                                    <Input type="radio" name="optApprove" class="Formelem" value="Y" checked> Yes
													&nbsp;&nbsp;&nbsp;<Input type="radio" name="optApprove" class="Formelem" Value="N">
                                                    No </td>
                                                    
														
													
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
															
																 <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
															
																 <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" disabled=true class="ActionButton" >
                                                                <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" disabled=true class="ActionButton" >
                                                                <input type="button" value="Next" name="btnNext" onClick="AddEntry('S')" class="ActionButton" >
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
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
<DIV class=frmBody id="DisVoucher" style="width:585; visibility:hidden; height:1;">
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="700">
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="35">&nbsp;</td>
		<td class="ExcelHeaderCell" align="center" width="75">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="125">Amount</td>
		<td class="ExcelHeaderCell" align="center" >Additional Details</td>
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