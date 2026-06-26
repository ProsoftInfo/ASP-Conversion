<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNCommissionAmd.asp
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo
dim sBookName,sInvoiceNo,sTemp,arrPartyCode,sPartyCode,sPartyName
Dim sInvTemp,iCtr,sVouTemp,sVouchTy,sNarr,sAmount,sTempAmt,sUserid
Dim Root,sExp,TempNode,sRetVal

Dim oDom,iTransNo
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")


Dim sFinPeriod,sFromYr,sToYr,sTempYr

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF

sUserid = getUserID()

sOrgId=Request.Form("selUnitId")
sOrgName=Request.Form("horgName")
iBookNo=Request.Form("selBook")
sBookName=Request.Form("hBookName")
sInvoiceNo=Request.Form("selInvoiceNo")
sVouchTy = Request.Form("selVoucherType")
sInvTemp = Split(sInvoiceNo,",")
iTransNo = Request.Form("hTransNo")

sVouchTy = "SC"

sVouTemp = Split(Request.Form("hVouDetails"),":")



IF CStr(sVouchTy) <> "SC" Then
	sTemp=Split(Request.Form("hVouDetails"),"-")
Else
	sNarr = "For Invoice "
	sAmount = 0
	sAmount = CDbl(sAmount)
	For iCtr = 0 To UBound(sVouTemp)
		sTemp = Split(sVouTemp(iCtr),"--")
		sNarr = sNarr&" "&sTemp(0)&" "&sTemp(1)&", "
		sAmount = CDbl(sAmount + sTemp(2))
	Next
End IF


sPartyName=Request.Form("txtPartyName")
arrPartyCode=split(Request.Form("hPartyCode"),"?")

Set objRs = Server.CreateObject("ADODB.RecordSet")

'oDOM.load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
Response.Write iTransNo

sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)
oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CNComm_"&Session.SessionID&".xml")

Set Root = oDom.documentElement
sExp = "//Party"
Set TempNode = Root.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sPartyName = TempNode.Item(0).Text
End IF

sExp = "//voucher"
Set TempNode = Root.selectNodes(sExp)

IF TempNode.length <> 0 Then
	sNarr = "For Invoice "
	sAmount = 0
	sOrgId = TempNode.Item(iCtr).attributes.getNamedItem("UnitNo").Value
	sOrgName  = TempNode.Item(iCtr).attributes.getNamedItem("UnitName").Value

	For iCtr = 0 To TempNode.length - 1
		sNarr = sNarr&" "&TempNode.Item(iCtr).attributes.getNamedItem("SalVouNo").Value
		sNarr = sNarr&" "&TempNode.Item(iCtr).attributes.getNamedItem("SalVouDate").Value
		sNarr = sNarr&", "
		IF CStr(TempNode.Item(iCtr).attributes.getNamedItem("CommisionValue").Value) <> "" Then
			sAmount = CDbl(sAmount + TempNode.Item(iCtr).attributes.getNamedItem("CommisionValue").Value)
		End IF
	Next
End IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData" src="<%="../temp/transaction/Voucher AMD_CNComm_"&Session.SessionID&".xml"%>"></XML>
<XML id="EntryData">
<Entry No="0" Payto="" Amount="" CRDR="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></XML>
</XML>
<XML id="AccHeadData">
<account/>
</XML>
<XML ID="UnitBookData"><Book/></XML>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>

<SCRIPT LANGUAGE=javascript>
function addclick(rightCombo , leftCombo  , removeBtn)
{
	var thelist = document.forms[0].elements[leftCombo]; //"selectItem"
	var tmplist = document.forms[0].elements[rightCombo]; //"selectedItem"

  	tmplength = tmplist.options.length;

	//tmplist.size=5;
        if(tmplength == -25)
		alert("Only 25 Items can be selected");
	else
	{
    		for( i = thelist.options.length -1; i >=0 ; i--)
    		{
        		if(thelist.options[i].selected)
        		{
            			newname = thelist.options[i].text;
            			newname1 = thelist.options[i].value;
            			addit = true;
            			for(j = tmplength-1; j >=0 ; j--)
            			{
				 	//if(tmplist.options[j].text == newname)
				 	if(tmplist.options[j].value == newname1)
				 	{
						alert("Selected item already  in the list");
                    				addit = false;
                    				break;
					}
	    			}
	    			if(addit)
	    			{
					tmplist.options.length = ++tmplength;
					tmplist.options[tmplength-1].text = newname;
					tmplist.options[tmplength-1].value = newname1;
					document.forms[0].elements[removeBtn].disabled = false;
					//document.forms[0].next.disabled = false;
					thelist.options[i] = null;
	    			}
			}
		}// end of for loop
	}// end of else

    //tmplist.size = 5;
    //thelist.size = 5;
    UpdateNarrAmt();
}

function removeclick(rightCombo , leftCombo  , removeBtn)
{
	var thelist = document.forms[0].elements[leftCombo]; //"select Item"
	var tmplist = document.forms[0].elements[rightCombo]; //"selected Item"

	tmplength = tmplist.options.length;

    	for(i = tmplist.options.length -1; i >=0 ; i--)
    	{
        	if(tmplist.options[i].selected)
        	{
            		thelist.options[thelist.options.length] = new Option(tmplist.options[i].text,tmplist.options[i].value);
            		tmplist.options[i] = null;
	 	}
    	}

    	if(tmplist.options.length == 0)
    	{
	  document.forms[0].elements[removeBtn].disabled = true;
	  //document.forms[0].next.disabled = true;
    	}

    	//tmplist.size = 5;
    	//thelist.size = 5;
    	UpdateNarrAmt();
}

</Script>

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

Function AddEntry(sVal)
'IF Voucher Type is Sal Comm Then calling a Seprate Functions
IF document.formname.hVouchTy.value = "SC" Then
	AddEntrySC(sVal)
	Exit Function
End IF

dim iCode,dRatio,dAmount
dim HeaderNode,nodANL
	if not checkFileds then exit function
	EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	EntryRoot.Attributes.Item(1).nodeValue=document.formname.txtPayTo.value
	EntryRoot.Attributes.Item(2).nodeValue=document.formname.txtAmount.value

	VouRoot.Attributes.Item(4).nodeValue=document.formname.ctlDate.GetDate

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
	VouRoot.appendChild EntryRoot
	SaveXML





end Function

Function AddEntrySC(sVal)
'MsgBox EntryRoot.xml

dim iCode,dRatio,dAmount
dim HeaderNode,nodANL,sStr,TempNode,iCtr

IF CStr(document.formname.hEditEntry.value) = "1" Then
	UpdateXml()
	document.formname.add.disabled = True
	document.formname.remove.disabled = True
	'Exit Function
End IF

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

	IF CStr(sVal) <> "U" Then
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
	VouRoot.appendChild EntryRoot

	IF CStr(sVal) = "A" Then
		DisplayVoucher("F")
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
		DisplayVoucher("F")
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

	DisplayVoucher("F")

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

FUNCTION DisplayVoucher(sCallTy)
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTdsAmt,sTdsPer
dim dTotal,sAccUnit,sTotalCRDR,sStr,TempNode,iRow
Dim sParCode,sParTy,sParSubTy,sTempNam,iOldInvNo

sTempNam ="Fill"

sAccUnit = document.formname.hOrgName.value

window.DisVoucher.style.height="200px"
window.DisVoucher.style.visibility="visible"
ClearTable "tblVoucher",1,1
dTotal=0
iRow = 1
Set VouRoot = DetData.documentElement

sExp = "//Party"
Set TempNode = VouRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sParCode = TempNode.Item(0).Attributes.getNamedItem("ParCode").value
	sParTy = TempNode.Item(0).Attributes.getNamedItem("ParType").value
	sParSubTy = TempNode.Item(0).Attributes.getNamedItem("ParSubType").value
	document.formname.hPartyCode.value = sParTy&"?"&sParSubTy&"?"&sTempNam&"?"&sParCode
End IF

IF CStr(sCallTy) = "B" Then
	PopulateInvoices()
	For Each EntryNode in VouRoot.childNodes
		IF CStr(EntryNode.nodeName) = "voucher" Then
			iOldInvNo = iOldInvNo &","&EntryNode.Attributes.Item(6).nodeValue
		End IF
	Next
	iOldInvNo = Trim(iOldInvNo)
	iOldInvNo = Mid(iOldInvNo,2)
	'MsgBox iOldInvNo
	document.formname.hOldInvCode.value = iOldInvNo

End IF


For Each EntryNode in VouRoot.childNodes
	IF CStr(EntryNode.nodeName) = "voucher" Then
		sAccUnit = EntryNode.Attributes.Item(1).nodeValue
	End IF

	IF CStr(EntryNode.nodeName) = "Entry" Then
		iSno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(2).nodeValue
		sTdsAmt = EntryNode.Attributes.Item(4).nodeValue
		sTdsPer = EntryNode.Attributes.Item(6).nodeValue
		sAmount=FormatNumber(CDbl(sAmount),2,,,0)
		sTdsAmt=FormatNumber(CDbl(sTdsAmt),2,,,0)
		sTdsPer=FormatNumber(CDbl(sTdsPer),2,,,0)


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


next
iEntryNo = iSno + 1



'End of Voucher Node Loop
	'if dTotal < 0 then
	'	sTotalCRDR="&nbsp;Cr"
	'	dTotal=CDbl(dTotal)*-1
	'else
	'	sTotalCRDR="&nbsp;Dr"
	'end if

	'dTotal="Rs. &nbsp;"&FormatNumber(dTotal,2,,,0)

	'set oRow = document.all.tblVoucher.insertRow(iSno+1)
	'InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,4,0,""
	'InsertCell oRow,1,"",CStr(dTotal)&sTotalCRDR ,"ExcelDisplayCell","right","top",0,0,0,0,""
	'InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
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

	'if CDate(document.formname.ctlDate.GetDate) < CDate(document.formname.hInvDate.value) then
	'	Msgbox("Credit Note date should be >= Invoice date")
	'	document.formname.ctlDate.focus
	'	checkFileds=false
	'	exit Function
	'end if
	checkFileds=true
end Function
'---------------------End Of Function checkFileds--------------------------
Function SaveXML()
	'IF CheckApp() Then
		'MsgBox document.formname.hOldInvCode.value
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLSave.asp?Mod=CNComm&Name=Voucher AMD", false
		objhttp.send DetData.XMLDocument
		if objhttp.responseText <> "" then
			Msgbox(objhttp.responseText)
		else
			document.formname.btnNext.disabled = True
			document.formname.btnAdd.disabled = True
			document.formname.submit()
		end if
	'End IF
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
		IF CStr(iVouEntryNo) = "1" Then
			document.formname.add.disabled = False
			document.formname.remove.disabled = False
		Else
			document.formname.add.disabled = True
			document.formname.remove.disabled = True
		End IF

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

				'alert VouRoot.xml

				set EntryRoot=VouRoot.removeChild(EntryNode)

				'alert VouRoot.xml

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

Function PopulateInvoices()
	Dim sParExp,sParNode,iCode,sInvNo
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	Dim sOrgid
	sOrgid = Trim(document.formname.hOrgId.value)
	objhttp.Open "GET","XMLCommisionDetails.asp?OrgId="&sOrgId&"&AgentCode=" & document.formname.hPartyCode.value , false
	objhttp.send
	'alert objhttp.responseText
	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			document.formname.selFrombox.length = document.formname.selFrombox.length+1
			document.formname.selFrombox.options(document.formname.selFrombox.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
			document.formname.selFrombox.options(document.formname.selFrombox.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
		next
	end if	'End of Agent Has Any Commision Check

	sParExp = "//voucher"
	Set sParNode = VouRoot.selectNodes(sParExp)
	IF sParNode.length <> 0 Then
		For iCode = 0 To sParNode.length - 1
			sInvNo = sInvNo&","&sParNode.Item(iCode).Attributes.getNamedItem("SalTransNo").value
		Next
		sInvNo = Mid(sInvNo,2)

		objhttp.Open "GET","XMLCommisionDetails.asp?OrgId="&sOrgId&"&AgentCode=" & document.formname.hPartyCode.value&"&sSelInv=" & sInvNo , false
		objhttp.send
		'alert objhttp.responseText
		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement
			For Each HeaderNode In Root.childNodes
				document.formname.selTobox.length = document.formname.selTobox.length+1
				document.formname.selTobox.options(document.formname.selTobox.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
				document.formname.selTobox.options(document.formname.selTobox.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			next

		end if	'End of Agent Has Any Commision Check
	End IF

	'document.formname.selFrombox.disabled = True
	'document.formname.selTobox.disabled = True
	document.formname.add.disabled = True
	document.formname.remove.disabled = True

End Function

Function UpdateXml()
	Dim sExp,TempNode,CheckNode,iCtr,sSelInv,iLen,iInvNo,OldRoot,newElem
	Dim sTempVal,sTemparr,PartyNode,OldPartyNode

	sSelInv = document.formname.selTobox.value
	iLen = document.formname.selTobox.length
	sExp = "//voucher"
	Set TempNode = VouRoot.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			Set OldRoot = VouRoot.removeChild(TempNode.item(iCtr))
		Next
	End IF

	sExp = "//Party"
	Set PartyNode = OldRoot.selectNodes(sExp)
	'Set OldPartyNode = VouRoot.removeChild(PartyNode.Item(0))
	'MsgBox PartyNode.length

	sExp = "//Entry"
	Set TempNode = VouRoot.selectNodes(sExp)
	'MsgBox TempNode.length

	For iCtr = 0 To iLen - 1
		sSelInv = sSelInv&","&document.formname.selTobox.options(iCtr).value
		sTempVal = document.formname.selTobox.options(iCtr).text
		sTemparr = Split(sTempVal,"--")

		iInvNo = document.formname.selTobox.options(iCtr).value

		Set newElem = DetData.createElement("voucher")
		newElem.setAttribute "UnitNo", OldRoot.Attributes.Item(0).nodeValue
		newElem.setAttribute "UnitName", OldRoot.Attributes.Item(1).nodeValue
		newElem.setAttribute "BookNo", OldRoot.Attributes.Item(2).nodeValue
		newElem.setAttribute "BookName", OldRoot.Attributes.Item(3).nodeValue
		newElem.setAttribute "VouDate", OldRoot.Attributes.Item(4).nodeValue
		newElem.setAttribute "Approver", OldRoot.Attributes.Item(5).nodeValue
		newElem.setAttribute "SalTransNo", iInvNo
		newElem.setAttribute "SalVouNo", sTemparr(0)
		newElem.setAttribute "SalVouDate", sTemparr(1)
		newElem.setAttribute "CrTransNo", OldRoot.Attributes.Item(9).nodeValue
		newElem.setAttribute "CrVoucherNo", OldRoot.Attributes.Item(10).nodeValue
		newElem.setAttribute "TransNo", OldRoot.Attributes.Item(11).nodeValue
		newElem.setAttribute "VoucherNo", OldRoot.Attributes.Item(12).nodeValue
		newElem.appendChild PartyNode.Item(0)

		VouRoot.insertBefore newElem,TempNode.item(0)
	Next
	'alert VouRoot.xml

	'VouRoot.insertBefore PartyNode.Item(0),TempNode.item(0)
	'sSelInv = Mid(sSelInv,2)
	'MsgBox sSelInv

	'alert VouRoot.xml

End Function

Function UpdateNarrAmt()
	Dim iCtr,sTemp,sTemparr,sNarr,iAmount

	document.formname.txtNarration.value = ""
	document.formname.txtamount.value = "0.00"
	sNarr = "For Invoice "
	iAmount = 0

	For iCtr = 0 To document.formname.selTobox.length - 1
		sTemp = document.formname.selTobox.options(iCtr).text
		sTemparr = Split(sTemp,"--")
		sNarr = sNarr &sTemparr(0)&" "& sTemparr(1) &", "
		iAmount = iAmount + CDbl(Trim(sTemparr(2)))
	Next

	sNarr = Trim(sNarr)
	sNarr = Left(sNarr,Len(sNarr)-1)
	document.formname.txtNarration.value = sNarr
	document.formname.txtAmount.value = FormatNumber(iAmount,2,,,0)

End Function

Function SetDate()
	Dim sFromYr,sToYr
	sFromYr = document.formname.hFromYr.Value
	sToYr = document.formname.hToYr.Value
	sFromYr = "01/04/"&Trim(sFromYr)
	sToYr = "31/03/"&sToYr
	document.formname.ctlDate.setMinDate() = sFromYr
	document.formname.ctlDate.setMaxDate() = sToYr
	DisplayVoucher("B")
End Function

</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="SetDate()">

<form method="POST" name="formname" action="VouCNCommAmdGenrate.asp">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hBookName" value="<%=sBookName%>">
<input type="hidden" name="hVouchTy" value="<%=sVouchTy%>">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hEditEntry" value="0">
<input type="hidden" name="hPartyCode" value="0">
<input type="hidden" name="hOldInvCode" value="0">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">

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
										<td><span class="DataOnly"></span></td>
									<%End IF %>
	                                </tr>

									<!--tr>
										<td class="FieldCell" width="113">Commision Amount</td>
										<td width="230"><span class="DataOnly"></span></td>
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

	                                <tr>
	                                <td class="FieldCellSub" colspan="6">
	                                <table border="0" cellspacing="1" width="100%" class="TableOutlineOnly">
                                         <tr>
                                     <td colspan="2" class="TableHeader" width="50%"><p align="left">Select Invoice &nbsp;&nbsp;&nbsp;&nbsp; <!--input ID="FormsEditField10" NAME="txtSearch" VALUE SIZE="10" ONKEYUP="selectTheItem(this,'selFrombox')" class="formelem" --></td>
                                         </tr>
                                         <tr>
                                     <td width="50%" class="TableHeader" align="center">Select Item</td>
                                     <td width="50%" class="TableHeader" align="center">Selected Items</td>
                                         </tr>
                                         <tr>
                                     <td width="50%" class="TableInput"><p align="center"><select size="5" name="selFrombox" multiple class="FormElem">
                                    </select></td>
                                     <td width="50%" class="TableInput"><p align="center"><select size="5" name="selTobox" multiple class="FormElem">
                                     </select></td>
                                         </tr>
                                         <tr>
                                     <td class="TableFooter" width="50%"><p align="center">
                                     <input type="button" value="Add &gt;&gt;" NAME="add" ONCLICK="addclick('selTobox','selFrombox','remove')" class="AddButton" ></td>
                                     <td class="TableFooter" width="50%"><p align="center">
                                     <input type="button" value="&lt;&lt; Remove" NAME="remove" ONCLICK="removeclick('selTobox','selFrombox','remove')" class="AddButton" ></td>
                                         </tr>
                                             </table>

                                       </td>
                                      <tr>

	                                <!--tr>
										<td class="FieldCell" width="113" valign="top">Sales Invoices</td>
										<td colspan="3" class="FieldCellSub">
										<Select name="selInvoiceNo" class="FormElem" Multiple>
										</select>
	                                </tr-->

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

													<textarea readonly rows="3" name="txtNarration" cols="50" class="FormElem"></textarea> </td>

                                                    <td class="FieldCell" colspan="2" valign="middle">
 </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="139">Amount</td>
                                                    <td class="FieldCell" colspan="4">
                                                    <%IF CStr(sVouchTy) <> "SC" Then %>
														<input type="text" name="txtAmount" value="<%=FormatNumber(sAmount,2,,,0)%>" size="15" style="text-align:right" readonly class="FormelemRead"> </td>
													<%Else%>
														<input type="text" name="txtAmount" value="<%=FormatNumber(sAmount,2,,,0)%>" size="15" style="text-align:right" readonly class="FormelemRead"> </td>
													<%End IF %>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Deduction @</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
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
															<% IF CStr(sVouchTy) = "SC" Then %>
																 <input type="Button" value="Add Entry" name="btnAdd" onClick="AddEntry('A')" class="ActionButton" >
																 <input type="Button" value="Update" name="btnUpdate" onClick="AddEntry('U')" class="ActionButton" disabled>
																 <input type="Button" value="Delete" name="btnDel" onClick="DelEntry()" class="ActionButton" disabled>

															<%End IF %>

                                                                <input type="button" value="Next" onClick="AddEntry('S')" name="btnNext" class="ActionButton" >

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