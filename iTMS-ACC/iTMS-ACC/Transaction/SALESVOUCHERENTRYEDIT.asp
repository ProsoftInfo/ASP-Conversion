
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SALESVOUCHERENTRYEDIT.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	May 25, 2011
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
<!--#include file="../../include/Salpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include File="../../include/CheckACCPrevFinYear.asp"-->
<%
dim sOrgId,sOrgName,sBookCode,objRs,sQuery,iBookNo,sSalType,saTemp
dim sReferenceNo,sInvoiceNo,iBkAccHead,sPartyName,sInvDate,sSetInvDate
Dim sAccUnit,sAccUnitName,iSalTy,iSalAccHead,sSalAccHdName,sAgName,sPartyCode
Dim sAccBookRel,sName,nCreatedTransNo,nPartyType,nSaleType,sPayToRecFrom,nInvNo

sOrgId=Session("organizationcode")
sOrgName=Session("OrgShortName")
sAccUnit = sOrgId
sAccUnitName= sOrgName

Set objRs = Server.CreateObject("ADODB.RecordSet")

nCreatedTransNo = Request.QueryString("nTransNo")
Response.Write "<p><font color=red>nCreatedTransNo="&nCreatedTransNo

sQuery = " SELECT A.CREATEDTRANSNO,A.TRANSACTIONTYPE,A.PARTYTYPE,"&_
		 " A.ACCOUNTHEAD,A.CREATEDVOUCHSTATUS,A.BANKINSTRUMENTTYPE,V.PARTYNAME,A.BOOKCODE,A.PayToRecdFrom,A.PartySubType,A.PARTYCODE "&_
		 " FROM ACC_T_CREATEDVOUCHERHEADER AS A INNER JOIN APP_M_PARTYMASTER AS V "&_
		 " ON A.PARTYCODE=V.PARTYCODE WHERE  A.OUDEFINITIONID='"& sOrgId &"' AND "&_
		 " isNull(A.OtherApplnTransNo,0) = 0 AND A.CREATEDTRANSNO = '"& nCreatedTransNo &"' "
objRs.Open sQuery,con
Response.Write sQuery
If Not objRs.EOF Then
	nPartyType = objrs(2)
	sPartyName = objRs(6)
	iBookNo = objrs(7)
	nSaleType = objRs(5)
	sPayToRecFrom = objrs(8)
	sPartyCode = nPartyType & "?"& objrs(9)& "?" & "" & "?"& objRs(10)
End IF
objRs.Close 
If sPayToRecFrom <> "" Then 
	If InStr(1,sPayToRecFrom,"-") Then
		nInvNo = split(sPayToRecFrom,"-")(0)
		sSetInvDate = split(sPayToRecFrom,"-")(1)
	Else
		nInvNo = sPayToRecFrom
		sSetInvDate = date()
	End IF
End IF 

Dim objfs,oDOM,sRetVal
Set objfs = CreateObject("Scripting.FileSystemObject")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sRetVal = GetVouchXML(nCreatedTransNo)
oDOM.Load server.MapPath(sRetVal)

if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")) then
	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"))
End IF

oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<XML id="DetData">
<Details BasicValue="" Discount="" ActualValue="" VouDate=""/></XML>
<XML id="EntryData"><Entry No="0" PayTo="" Amount="" Qty="" UOM="" UOMValue="" Rate="" ActValue="" DisPer="" DisAmount="" ItemCode="" ClassCode="" TransBasicamt="" TransRate="" TransDisAmt="" TransInvAmt="" RatePer="" NoofPack="" PackType="" RndOff="" /></XML>
<XML id="AccHeadData">
<account/>
</XML>
<xml id="PartyData"><Root /></xml>
<xml id="OutData"><Root /></xml>
<xml id="UnitBookData"><Book /></xml>
<xml id="SaleTypeData"><Book /></xml>
<!--<xml id="VoucherData"><Voucher/></xml>-->
<XML ID="VoucherData" src="<%="../temp/transaction/Voucher AMD_SAL_"&Session.SessionID&".xml"%>"></XML>
<xml id="TEMPXML"><Root></Root></xml>
<xml id="GLHeadData"><Root></Root></xml>
<xml id="ItemData"><Root></Root></xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../../scripts/checkdate.js"></script>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
dim iBookAccCode

iEntryNo=1
bVouFlag=false
bSavFlag=false
set VouRoot=DetData.documentElement
set EntryRoot=EntryData.documentElement
'iBookAccCode=<%=iBkAccHead%>
'***********************************

Function GetAccHead(obj)
    Dim iAccHead,sAccHdName
'    alert(obj.value)
    set objhttp = createobject("Microsoft.xmlhttp")
    objhttp.open "GET","GetAccHeadName.asp?Mod=SAL&InvType="&obj.value,false
    objhttp.send
'    alert(objhttp.responseText)
    if trim(objhttp.responseText)<>"" then
        sTemp = split(objhttp.responseText,":")
        if uBound(sTemp)=0 then
            alert(stemp)
        else
            iAccHead = stemp(0)
            iAccHdName = stemp(1)
'            alert(iAccHead)
            if trim(iAccHead)="" then

            end if
            document.formname.hSalAccCode.value = iAccHead
            document.formname.hSalAccName.value = iAccHdName
        end if
    end if



    if trim(document.formname.hSalAccCode.value)="0" then
	    document.formname.selAccountHead.disabled = true
        for iCnt=0 to cint(document.formname.selAccountHead.length)-1
            if document.formname.selAccountHead(iCnt).value="S" then
                document.formname.selAccountHead.selectedIndex = iCnt
            end if
        next
    else
        document.formname.selAccountHead.disabled = true
        for iCnt=0 to cint(document.formname.selAccountHead.length)-1
            if document.formname.selAccountHead(iCnt).value="G" then
                document.formname.selAccountHead.selectedIndex = iCnt
            end if
        next
        spAccHead.innerHTML =  document.formname.hSalAccName.value
    end if 'if trim(document.formname.hSalAccCode.value)="0" then

End Function
'*************************************
Function SelPartyName()
    dim sOrgId,sPartyType
    Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp

    sRetVal2 = "0:0"
    sOrgId=document.formname.hOrgId.value
    'sPartyType=objAcc.value &"?"& objAcc.options(objAcc.selectedIndex).text

    'if objAcc.selectedIndex >0 then

	'    set OutValue = showModalDialog("../../Common/PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,PartyData,"dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	 '   sQuery = OutValue.getAttribute("PassQuery")
	'    if OutValue.getAttribute("Action")="CLOSE" then exit function
'
'	    while OutValue.getAttribute("Action")<>"Done"
'		    set OutValue = showModalDialog("../../Common/PartySelection.asp?"&sQuery,PartyData,"dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'		    sQuery = OutValue.getAttribute("PassQuery")
'	        if OutValue.getAttribute("Action")="CLOSE" then exit function
'	    wend

	    'alert(OutValue.xml)
	    
	    sTempValWindowSize = GetWindowSizeForPopup("2")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
	    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sOrgId&"&Party="&sPartyType,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

        if OutValue.hasChildNodes() then
            For each ndChild in OutValue.childNodes
	            sParTy = ndChild.getAttribute("RetField3")
	            sParSubType = ndChild.getAttribute("RetField4")
	            sParCode = ndChild.getAttribute("RetField1")
	            sPartyName = ndChild.getAttribute("RetField0")
	            document.formname.hParCode.value = sParCode
	            document.formname.txtPartyName.value = sPartyName
	            sRetVal2 = ndChild.getAttribute("RetField3")
	            sRetVal2 =sRetVal2 &":"& ndChild.getAttribute("RetField4")
	            sRetVal2 =sRetVal2 &":"& ndChild.getAttribute("RetField1")
	            sRetVal2 =sRetVal2 &":"& ndChild.getAttribute("RetField0")
	        Next
	    end if 'if OutValue.hasChildNodes() then

	    set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "GET","../../Common/PartySubType.asp?ParCode="&sParCode&"&OrgCode="&sOrgId,false
		objhttp.send
		'alert(objhttp.responseText)
		if trim(objhttp.responseXML.xml)<>"" then
			TEMPXML.loadXML(objhttp.responseXML.xml)
			set sTempRoot = TEMPXML.documentElement
			document.formname.selParType.length = 1
			for each subNode in sTempRoot.childNodes
				document.formname.selParType.length= document.formname.selParType.length + 1
				document.formname.selParType(document.formname.selParType.length-1).value = subNode.getAttribute("SubType")
				document.formname.selParType(document.formname.selParType.length-1).text = subNode.text
			next
		end if

    'End if 'End of If any Account Head Selected Check
End function
'*************************************
Function PartyType(obj)
    if obj.value<>"A" then
        sParCode = document.formname.hParCode.value
        sParName = document.formname.txtPartyName.value
        Set nodAccHead = AccHeadData.documentElement
'        alert(nodAccHead.xml)
        if nodAccHead.hasChildNodes then
		    For Each HeaderNode In nodAccHead.childNodes
		        nodAccHead.removeChild(HeaderNode)
		    next
	    end if 'End of Party Head Processing

        sRetVal2 = replace(obj.value,"|",":")&":"&sParCode&":"&sParName&":0"

	    GetPartyHeadXml sParCode,sParName,sRetVal2
	   ' alert(nodAccHead.xml)
	    if nodAccHead.hasChildNodes then
		    'User Has Selected a GL Account Head
		    For Each HeaderNode In nodAccHead.childNodes
			    document.formname.hPartyCode.value=replace(obj.value,"|","?")&"?"& obj(obj.selectedIndex).text &"?"& HeaderNode.Attributes.Item(0).nodeValue
			    document.formname.txtPartyName.value=HeaderNode.Attributes.Item(3).nodeValue
		    next
	    else
		    obj.selectedIndex=0
	    end if 'End of Party Head Processing
	else
	    alert("Select Party Sub Type")
	    document.formname.selParType.focus()
	    exit function
    end if
End Function
'***********************************************
Function PopulateSalTy
	Dim iUnitNo,iBookNo
	iUnitNo = document.formname.hOrgId.value
	iBookNo = document.formname.selBook(document.formname.selBook.selectedIndex).value
	document.formname.hBookcode.value = iBookNo

	if trim(iBookNo)<>"S" then
	    Set objhttp = CreateObject("MSXML2.XMLHTTP")
	    objhttp.Open "GET","XMLGetBookSalPurType.asp?BkCode=05&orgID="& iUnitNo&"&BookNo="&iBookNo , false
	    objhttp.send

	    'Msgbox objhttp.responseText
	    if objhttp.responseXML.xml <> "" then
		    SaleTypeData.loadXML objhttp.responseXML.xml
		    Set Root = SaleTypeData.documentElement
		    document.formname.selSaleType.length = 1
		    For Each HeaderNode In Root.childNodes
			    document.formname.selSaleType.length = document.formname.selSaleType.length+1
			    document.formname.selSaleType.options(document.formname.selSaleType.length-1).text = HeaderNode.Attributes.Item(2).nodeValue
			    document.formname.selSaleType.options(document.formname.selSaleType.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
		    next
	    end if
	else
	    alert("Select Sales Book")
	    exit function
	end if ' if trim(iBookNo)<>"S" then
End Function
'**************************************************
Function DisplayBook()
dim iUnitNo,arrTemp
dim Root
	document.formname.selBook.options.length = 1
		iUnitNo= document.formname.hOrgId.value

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=05&orgID=" & iUnitNo , false
		objhttp.send

		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement
			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if
		popPartyType
		
		'nBookCode = document.formname.hBookcode.value
		'For i = 0 To document.formname.selBook.length -1
		'	If document.formname.selBook.options(i).value <> "S" Then
		'		If len(document.formname.selBook.options(i).value) = 1 Then
		'			nSelBookCode = "0"&document.formname.selBook.options(i).value
		'		End If
		'		If nBookCode = nSelBookCode Then
		'			document.formname.selBook.selectedIndex = i
		'		End IF
		'	End IF
		'Next
end Function

'************************************

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
					showGLHead
				else
					sTemp=Split(objAcc.value,"?")
					sDesc=objAcc.options(objAcc.selectedIndex).text
					window.spAccHead.innerHTML=sDesc
					document.formname.txtDescription.value=sDesc	'u
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
function showGLHead()
dim iAccCode,bAnal,bCostCenter
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,arrTemp,sRetVal
iBookNo=document.formname.hBookcode.value
sOrgId= document.formname.hOrgId.value
'	set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'	    sQuery = OutValue.getAttribute("PassQuery")
'	    if OutValue.getAttribute("Action")="CLOSE" then exit function
'	while OutValue.getAttribute("Action")<>"Done"
'		set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'		    sQuery = OutValue.getAttribute("PassQuery")
'	        if OutValue.getAttribute("Action")="CLOSE" then exit function
'	wend

sTempValWindowSize = GetWindowSizeForPopup("5")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
		Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&BookId=01&BookNo="&iBookNo,GLHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,GLHeadData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

'alert(OutValue.xml)
if OutValue.hasChildNodes() then
    For each ndChild in OutValue.childNodes
        sRetVal = ndChild.getAttribute("RetField0")
        sRetVal = sRetVal &":"& ndChild.getAttribute("RetField1")
        sRetVal = sRetVal &":"& ndChild.getAttribute("RetField2")
        sRetVal = sRetVal &":"& ndChild.getAttribute("RetField3")
        sRetVal = sRetVal &":"& ndChild.getAttribute("RetField4")
        sRetVal = sRetVal &":"& ndChild.getAttribute("RetField5")
        sRetVal = sRetVal &":"& ndChild.getAttribute("RetField6")
        sRetVal = sRetVal &":"& ndChild.getAttribute("RetField7")
    Next
end if

'if UBound(arrTemp) <= 1 then exit function
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
		'document.formname.txtDescription.value=HeaderNode.Attributes.Item(3).nodeValue
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
'alert(bFlag)
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
'' End of Validation

	if not checkFileds then exit function

	if bFlag<>"U" then
		iEntryNo=iEntryNo+1
		EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	end if


	IF CStr(bFlag) = "U" Then
		EntryRoot.Attributes.Item(0).nodeValue=document.formname.hEntryNo.value
		document.formname.hEntryNo.value="0"
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
	'MsgBox "A"
	'MsgBox VouRoot.xml

	EntryRoot.setAttribute "PayTo",document.formname.txtDescription.value
	EntryRoot.setAttribute "Amount",dAccAmt
	EntryRoot.setAttribute "Qty",document.formname.txtQty.value
	EntryRoot.setAttribute "UOM",document.formname.selUOM.value
	EntryRoot.setAttribute "UOMValue",document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
	EntryRoot.setAttribute "Rate",document.formname.txtRate.value
	EntryRoot.setAttribute "ActValue",document.formname.txtValue.value
	EntryRoot.setAttribute "DisPer",document.formname.txtDisPercentage.value
	EntryRoot.setAttribute "DisAmount",document.formname.txtDisAmount.value
	EntryRoot.setAttribute "ItemCode",document.formname.hItemCode.value
	EntryRoot.setAttribute "ClassCode", document.formname.hClassCode.value
	EntryRoot.setAttribute "TransBasicamt",document.formname.txtValue.value
	EntryRoot.setAttribute "TransRate",document.formname.txtRate.value
	EntryRoot.setAttribute "TransDisAmt",document.formname.txtDisAmount.value
	EntryRoot.setAttribute "TransInvAmt","0"
	EntryRoot.setAttribute "RatePer", document.formname.txtRatePer.value
	EntryRoot.setAttribute "NoofPack", document.formname.txtBagno.value
	EntryRoot.setAttribute "PackType",0
	EntryRoot.setAttribute "RndOff",dRndVal
	
	
	


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
	'MsgBox "B"
	'MsgBox VouRoot.xml

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
		
		document.formname.txtDescription.value=""
		document.formname.txtQty.value = "0.00"
		document.formname.txtDisAmount.value = "0.00"
		document.formname.txtDisPercentage.value = "0.00"
		document.formname.txtRate.value = "0.00"
		document.formname.txtAmount.value = "0.00"
		document.formname.txtValue.value = "0.00"
		
		'document.formname.reset
		window.spAccHead.innerHTML=""

		document.all.spUOM.innerHTML = document.formname.selUOM.options(document.formname.selUOM.selectedIndex).text
		'document.all.spPack.innerHTML = document.formname.selPack.options(document.formname.selPack.selectedIndex).text

	end if
	document.formname.btnAdd.disabled = false
	document.formname.btnCancel.disabled = false
	document.formname.btnDel.disabled = true
	document.formname.btnNext.disabled = false
	document.formname.btnUpdate.disabled = true

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
	'alert(TempNode.Item(iCtr).xml)
		TempNode.Item(iCtr).Attributes.Item(0).nodeValue = iCtr+1
		iSno=TempNode.Item(iCtr).Attributes.Item(0).nodeValue
		sDescription=TempNode.Item(iCtr).Attributes.Item(1).nodeValue
		sAmount=FormatNumber(TempNode.Item(iCtr).Attributes.Item(2).nodeValue,2,,,0)

		sRate = TempNode.Item(iCtr).Attributes.Item(6).nodeValue
		sQty = TempNode.Item(iCtr).Attributes.Item(3).nodeValue
		sUom = TempNode.Item(iCtr).Attributes.Item(5).nodeValue
		dRatePer = TempNode.Item(iCtr).Attributes.Item(16).nodeValue
		sValue=FormatNumber(TempNode.Item(iCtr).Attributes.Item(7).nodeValue,2,,,0)
		sDiscount=FormatNumber(TempNode.Item(iCtr).Attributes.Item(9).nodeValue,2,,,0)
		IF Cstr(TempNode.Item(iCtr).Attributes.Item(10).nodeValue) = "" Then
			TempNode.Item(iCtr).Attributes.Item(10).nodeValue = 0
		End IF
		dRndVal = FormatNumber(TempNode.Item(iCtr).Attributes.Item(19).nodeValue,2,,,0)

	'	MsgBox dRndVal
	'	msgbox " sQty = "& sQty &" Rate = "& sRate &" RatePer = "& dRatePer

		sAmount = CDbl(sQty) * (CDbl(sRate) / CDbl(dRatePer))
		sAmount = CDbl(sAmount) - Cdbl(sDiscount)
		sAmount = CDbl(sAmount) + CDbl(dRndVal)
	'	Msgbox sAmount

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
	elseif 	ValidateAmount(document.formname.txtRate.value,"Rate",0,9999999999.99)=false then
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

	IF Trim(document.formname.txtRate.value) = 0 Then
		Exit Function
	End IF

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
	popAddAmount1()
	calculateField=true
end Function
'---------------------End Of Function calculateField----------------------------
Function SaveXML()
	
	Set Root = VoucherData.documentElement
	
	If Not validate Then
		Exit Function
	End IF
	
	Set UnitRoot = UnitBookData.documentElement
		For Each HeaderNode In UnitRoot.childNodes
			if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
				document.formname.hBkAccHead.value=HeaderNode.Attributes.Item(2).nodeValue

				sExp = "//Book"
				Set TempNode = Root.selectNodes(sExp)
				IF TempNode.length <> 0 Then
					TempNode.Item(0).Attributes.getNamedItem("BookId").Value = document.formname.selBook.value
					TempNode.Item(0).Attributes.getNamedItem("BKAccHead").Value = HeaderNode.Attributes.Item(2).nodeValue
					TempNode.Item(0).Attributes.getNamedItem("BKOtherUnits").Value = HeaderNode.Attributes.Item(3).nodeValue
					TempNode.Item(0).Text = document.formname.selBook.options(document.formname.selBook.selectedIndex).Text

				end if
			End IF
		next

		document.formname.hInvDate.value = document.formname.ctlDate.GetDate
		sExp = "//SalesType"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("SalType").Value = document.formname.selSaletype.value
			TempNode.Item(0).Text = document.formname.selSaletype.options(document.formname.selSaletype.selectedIndex).Text
		End IF

		sExp = "//SaleInvoice"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("InvNo").Value = document.formname.txtInvoiceNo.value
			TempNode.Item(0).Attributes.getNamedItem("InvDate").Value = document.formname.ctlDate.GetDate
			TempNode.Item(0).Attributes.getNamedItem("RefNo").Value = document.formname.txtRefNo.value
		End IF

		sExp = "//Details"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("VouDate").Value = document.formname.ctlDate.GetDate
		End IF

		if document.formname.optAgentExist(0).checked then
			sAgChk = "Y"
		else
			sAgChk = "N"
		end if

		sExp = "//Party"
		Set TempNode = Root.selectNodes(sExp)
		sTemp=Split(trim(document.formname.hPartyCode.value),"?")
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("ParType").Value = sTemp(0)
			TempNode.Item(0).Attributes.getNamedItem("ParSubType").Value = sTemp(1)
			TempNode.Item(0).Attributes.getNamedItem("ParSubTypeName").Value = document.formname.selParType.options(document.formname.selParType.selectedIndex).Text
			TempNode.Item(0).Attributes.getNamedItem("ParCode").Value = sTemp(3)
			TempNode.Item(0).Attributes.getNamedItem("Agent").Value = sAgChk
			TempNode.Item(0).Text = document.formname.txtPartyName.value
		End IF
		
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Mod=SAL&Name=Voucher AMD", false
	objhttp.send VoucherData.XMLDocument
	'alert(VoucherData.xml)
	'Exit Function
	
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		'alert(OldVouData.xml)
		document.formname.submit()
	end if
End Function

Function SaveXML1()
    'alert(DetData.xml)
	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Mod=SAL&Name=Voucher Entry", false
	objhttp.send VoucherData.XMLDocument

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLUpdate.asp?Mod=SAL&Name=Voucher Entry", false
	objhttp.send DetData.XMLDocument

	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		IF document.formname.txtDescription.value = "" and iEntryNo < 1  Then
			Msgbox "Select Item "
			document.formname.txtDescription.focus()
			Exit Function
		End IF
		document.formname.btnNext.disabled = True
		document.formname.btnAdd.disabled = True
		document.formname.submit()
	end if
End Function
'*******************************
'*******************************

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
		EntryRoot.setAttribute "ItemCode",""
		EntryRoot.setAttribute "ClassCode",""
		EntryRoot.setAttribute "TransBasicamt",""
		EntryRoot.setAttribute "TransRate",""
		EntryRoot.setAttribute "TransDisAmt",""
		EntryRoot.setAttribute "TransInvAmt",""
		EntryRoot.setAttribute "RatePer",""
		EntryRoot.setAttribute "NoofPack",""
		EntryRoot.setAttribute "PackType",""
		EntryRoot.setAttribute "RndOff",""
		
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
'=============================================================================================

FUNCTION popAddAmount1()

'dim dAmount,iChildCount,dRatio,dTotal,dRatioTotal,iCounter


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
'=============================================================================================

Function EditEntry(iVouEntryNo)
Dim sStr,TempNode,iCounter,sTemp,VouRoot,sStr2,AccNode,iAccCode,sAccType,dPackTy
Set VouRoot = VoucherData.documentElement
if document.formname.hEntryNo.value <>"0" then
    Msgbox "Update the selected entry and edit this entry"
    exit function
end if
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
		'document.all.spPack.innerHTML = document.formname.selPack.options(document.formname.selPack.selectedIndex).text
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


Function GetItem()

	sorgID = document.formname.hOrgId.value
'	set OutValue = showModalDialog("../../Common/ItemSelectCommon.asp?orgID=" & sorgID,ItemData,"dialogHeight:480px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
'
'	    sQuery = OutValue.getAttribute("PassQuery")
'	    if OutValue.getAttribute("Action")="CLOSE" then exit function
'
'	while OutValue.getAttribute("Action")<>"Done"
'		set OutValue = showModalDialog("../../Common/ItemSelectCommon.asp?"&sQuery,ItemData,"dialogHeight:480px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
'		sQuery = OutValue.getAttribute("PassQuery")
'	    if OutValue.getAttribute("Action")="CLOSE" then exit function
'	wend

	sTempValWindowSize = GetWindowSizeForPopup("1")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
		Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId,ItemData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,ItemData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

	if OutValue.hasChildNodes() then
	    For each ndChild in OutValue.childNodes
	        document.formname.txtDescription.value = ndChild.getAttribute("ItemName")
	        document.formname.hItemCode.value = ndChild.getAttribute("ItemCode")
	        document.formname.hClassCode.value = ndChild.getAttribute("ClassCode")
	    Next
	end if

end function

Function ChDisp(sObj)
	Dim sDispVal

	sDispVal = sObj.options(sObj.selectedIndex).Text
	IF UCase(sObj.name) = "SELUOM" Then
		document.all.spUOM.innerHTML = sDispVal
	Else
		document.all.spPack.innerHTML = sDispVal
	End IF

End Function
'**********************************
FUNCTION popPartyType()
	dim iHeadCount

	iUnitNo=document.formname.hOrgId.value
	iBkNo=document.formname.selBook(document.formname.selBook.selectedIndex).value

	for iCounter=1 to document.formname.selParType.length
		document.formname.selParType.remove(1)
	next

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo&"&sCallTy=P" , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		iCounter=1
		For Each HeaderNode In Root.childNodes
			set oText1 = document.createElement("<Option>" )
				oText1.Text = HeaderNode.text
				oText1.Value = HeaderNode.Attributes.Item(0).nodeValue

			document.formname.selParType.add oText1,iCounter
			iCounter=CDbl(iCounter)+1
		next
	end if
	
	nPartyType =document.formname.hPartyType.value	
	
	'For i = 0 to document.formname.selParType.length - 1
	'	If document.formname.selParType.options(i).value <> "A" Then
	'		If nPartyType = Split(document.formname.selParType.options(i).value,"?")(0) Then
	'			document.formname.selParType.selectedIndex = i
	'		End If
	'	End IF
	'Next
END FUNCTION

'**********************************

Function showAgent(bFlag)
dim Returnvalue,sExp,TempNode
Set Root = VoucherData.documentElement
if Root.hasChildNodes() then
    For Each HeaderNode In Root.childNodes
	    if HeaderNode.nodeName="AgentDetails" then
		    set temp=Root.removeChild(HeaderNode)
	    end if
    next
end if 'if Root.hasChildNodes() then

if bFlag<>"N" then
		Set Returnvalue = showModalDialog ("AgentCommisionEntry.asp?OrgID="&document.formname.hOrgId.value&"&AgentType="&bFlag ,OutData,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		if Returnvalue.hasChildNodes then
			Root.appendChild Returnvalue
			sExp = "//Agent"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.Length <> 0 Then
				document.all.spAgentName.innerHTML = TempNode.Item(0).Attributes.Item(1).nodeValue
				document.formname.hCommName.Value = TempNode.Item(0).Attributes.Item(1).nodeValue
			Else
				document.all.spAgentName.innerHTML = ""
				document.formname.hCommName.Value = ""
			End IF
		else
			document.formname.optAgentExist(2).checked=true
		end if
end if
End function
'************
function validate()
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select SalesBook")
		document.formname.selBook.focus
		validate= false
		exit function
	end if
	if document.formname.selSaletype.selectedIndex<1 then
		MsgBox ("Select Sales type")
		document.formname.selSaletype.focus
		validate=false
		exit function
	end if

	if document.formname.selParType.selectedIndex<1 then
		MsgBox ("Select Party")
		document.formname.selParType.focus
		validate=false
		exit function
	end if
	if trim(document.formname.txtPartyName.value)="" then
		MsgBox ("Party Name should not be blank")
		document.formname.txtPartyName.select
		validate=false
		exit function
	end if
	IF document.formname.txtInvoiceNo.value = "" Then
		MsgBox "Enter Invoice Number "
		document.formname.txtInvoiceNo.focus()
		validate=false
		Exit Function
	End IF

	validate=true
End function
'*********************
Function VouCreate
	Dim newElem,objHeader,sTemp,sCurrDate,TempNode,UnitRoot,sAgChk
	Dim sPartyTy,sInvNo,sInvDate,sSendVal,sRetVal,objhttp,sExp
	sCurrDate = document.formname.hCurrDate.value

	Set objhttp = CreateObject("MSXML2.XMLHTTP")
	Set Root = VoucherData.documentElement

	if validate then

		IF document.formname.txtInvoiceNo.value = "" Then
			MsgBox "Enter Invoice Number "
			document.formname.txtInvoiceNo.focus()
			Exit Function
		End IF

		IF DateDiff("d",document.formname.ctlDate.getDate(),sCurrDate) < 0 Then
			MsgBox "Voucher Date Should be Less than the System Date "
			Exit Function
		End IF

		sPartTy = document.formname.hPartyCode.value
		sInvNo = document.formname.txtInvoiceNo.value
		sInvDate = document.formname.ctlDate.GetDate()

		sSendVal = sPartTy&"?"&sInvNo&"?"&sInvDate&"?05"&"?"&document.formname.hOrgID.value
		'MsgBox sSendVal

		document.formname.hInvDate.value=document.formname.ctlDate.GetDate
		document.formname.hOrgName.value=document.formname.hOrgName.value	'document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).Text
		'document.formname.hSalType.value=document.formname.selSaleType.options(document.formname.selSaleType.selectedIndex).Text

		sExp = "//Organization"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("OrgId").Value = document.formname.hOrgID.value
			TempNode.Item(0).Text = document.formname.hOrgName.value 'document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).Text
		End IF

		Set UnitRoot = UnitBookData.documentElement
		For Each HeaderNode In UnitRoot.childNodes
			if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
				document.formname.hBkAccHead.value=HeaderNode.Attributes.Item(2).nodeValue

				sExp = "//Book"
				Set TempNode = Root.selectNodes(sExp)
				IF TempNode.length <> 0 Then
					TempNode.Item(0).Attributes.getNamedItem("BookId").Value = document.formname.selBook.value
					TempNode.Item(0).Attributes.getNamedItem("BKAccHead").Value = HeaderNode.Attributes.Item(2).nodeValue
					TempNode.Item(0).Attributes.getNamedItem("BKOtherUnits").Value = HeaderNode.Attributes.Item(3).nodeValue
					TempNode.Item(0).Text = document.formname.selBook.options(document.formname.selBook.selectedIndex).Text

				end if
			End IF
		next

		document.formname.hInvDate.value = document.formname.ctlDate.GetDate
		sExp = "//SalesType"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("SalType").Value = document.formname.selSaletype.value
			TempNode.Item(0).Text = document.formname.selSaletype.options(document.formname.selSaletype.selectedIndex).Text
		End IF

		sExp = "//SaleInvoice"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("InvNo").Value = document.formname.txtInvoiceNo.value
			TempNode.Item(0).Attributes.getNamedItem("InvDate").Value = document.formname.ctlDate.GetDate
			TempNode.Item(0).Attributes.getNamedItem("RefNo").Value = document.formname.txtRefNo.value
		End IF

		sExp = "//Details"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("VouDate").Value = document.formname.ctlDate.GetDate
		End IF

		if document.formname.optAgentExist(0).checked then
			sAgChk = "Y"
		else
			sAgChk = "N"
		end if

		sExp = "//Party"
		Set TempNode = Root.selectNodes(sExp)
		sTemp=Split(trim(document.formname.hPartyCode.value),"?")
		IF TempNode.length <> 0 Then
			TempNode.Item(0).Attributes.getNamedItem("ParType").Value = sTemp(0)
			TempNode.Item(0).Attributes.getNamedItem("ParSubType").Value = sTemp(1)
			TempNode.Item(0).Attributes.getNamedItem("ParSubTypeName").Value = document.formname.selParType.options(document.formname.selParType.selectedIndex).Text
			TempNode.Item(0).Attributes.getNamedItem("ParCode").Value = sTemp(3)
			TempNode.Item(0).Attributes.getNamedItem("Agent").Value = sAgChk
			TempNode.Item(0).Text = document.formname.txtPartyName.value
		End IF

	else
		exit function
	end if
	'SaveXML()
End function

Function VouCreate1
dim newElem,objHeader,sTemp,sCurrDate
Dim sPartyTy,sInvNo,sInvDate,sSendVal,sRetVal,objhttp
sCurrDate = document.formname.hCurrDate.value

Set objhttp = CreateObject("MSXML2.XMLHTTP")
Set Root = VoucherData.documentElement

	if validate then

		IF document.formname.txtInvoiceNo.value = "" Then
			MsgBox "Enter Invoice Number "
			document.formname.txtInvoiceNo.focus()
			Exit Function
		End IF

		sPartTy = document.formname.hPartyCode.value
		sInvNo = document.formname.txtInvoiceNo.value
		sInvDate = document.formname.ctlDate.GetDate()

		sSendVal = sPartTy&"?"&sInvNo&"?"&sInvDate&"?05"&"?"&document.formname.hOrgId.value
		'MsgBox sSendVal

		objhttp.Open "GET","CheckInvCreate.asp?sValue="&sSendVal , false
		objhttp.send
'		alert(objhttp.responseText)
		sRetVal = objHttp.responseText
		'MsgBox sRetVal
		IF CStr(sRetVal) <> "C" Then
			MsgBox "Sales Voucher already Created for this Party,InvoiceNo and Invoice Date "
			Exit Function
		End IF


		set objHeader= VoucherData.createElement("Header")
		Root.appendChild objHeader

		Set newElem = VoucherData.createElement("Organization")
		newElem.setAttribute "OrgId",document.formname.hOrgId.value
		newElem.Text=document.formname.hOrgName.value
		objHeader.appendChild newElem

		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
				'document.formname.hBkAccHead.value=HeaderNode.Attributes.Item(2).nodeValue

				Set newElem = VoucherData.createElement("Book")
				newElem.setAttribute "BookId",document.formname.selBook(document.formname.selBook.selectedIndex).value
				newElem.setAttribute "BKAccHead",HeaderNode.Attributes.Item(2).nodeValue
				newElem.setAttribute "BKOtherUnits",HeaderNode.Attributes.Item(3).nodeValue
				newElem.Text=document.formname.selBook.options(document.formname.selBook.selectedIndex).Text
				objHeader.appendChild newElem
			end if
		next


		Set newElem = VoucherData.createElement("SalesType")
		newElem.setAttribute "SalType",document.formname.selSaletype(document.formname.selSaletype.selectedIndex).value
		newElem.Text=document.formname.selSaletype.options(document.formname.selSaletype.selectedIndex).Text
		objHeader.appendChild newElem

		Set newElem = VoucherData.createElement("SaleInvoice")
		newElem.setAttribute "InvNo",document.formname.txtInvoiceNo.value
		newElem.setAttribute "InvDate",document.formname.ctlDate.GetDate
		newElem.setAttribute "RefNo",document.formname.txtRefNo.value

		objHeader.appendChild newElem

		Set newElem = VoucherData.createElement("Party")

		sTemp=Split(trim(document.formname.hPartyCode.value),"?")

		newElem.setAttribute "ParType",sTemp(0)
		newElem.setAttribute "ParSubType",sTemp(1)
		newElem.setAttribute "ParSubTypeName",sTemp(2)
		newElem.setAttribute "ParCode",sTemp(3)

		if document.formname.optAgentExist(0).checked then
			newElem.setAttribute "Agent","Y"
		else
			newElem.setAttribute "Agent","N"
		end if
		newElem.Text=document.formname.txtPartyName.value
		objHeader.appendChild newElem
	else
		exit function
	end if

End function

Function DispOldVal()
	Dim sTempNode,sExp,sTemp,iCtr,Root,sTemp2,sTemp3,sAgChk,sAgType,sAgName
	Set Root = VoucherData.documentElement
	
	sExp = "//Book"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("BookId").Value
		sTemp2 = sTempNode.Item(0).Attributes.getNamedItem("BKAccHead").Value
		document.formname.hBkAccHead.value = sTemp2

		For iCtr = 0 To document.formname.selBook.length - 1
			IF document.formname.selBook.options(iCtr).value = sTemp Then
				document.formname.selBook.selectedIndex = iCtr
				Exit For
			End IF
		Next
		PopulateSalTy
	End IF

	sExp = "//Party"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		sTemp = sTempNode.Item(0).Attributes.getNamedItem("ParType").Value
		sTemp2 = sTempNode.Item(0).Attributes.getNamedItem("ParSubType").Value
		sTemp3 = sTempNode.Item(0).Attributes.getNamedItem("ParCode").Value
		sAgChk = sTempNode.Item(0).Attributes.getNamedItem("Agent").Value
		document.formname.txtPartyName.value = sTempNode.Item(0).Text


		sTemp = sTemp&"?"&sTemp2
		For iCtr = 0 To document.formname.selParType.length - 1
			IF document.formname.selParType.options(iCtr).value = sTemp Then
				document.formname.selParType.selectedIndex = iCtr
				Exit For
			End IF
		Next
		sTemp = sTemp&"?"&sTempNode.Item(0).Text
		sTemp = sTemp&"?"&sTemp3

		document.formname.hPartyCode.value = sTemp
	End IF

	IF CStr(sAgChk) = "Y" Then
		sExp = "//AgentDetails/Agent"
		Set sTempNode = Root.selectNodes(sExp)
		IF sTempNode.length <> 0 Then
			sAgType = sTempNode.Item(0).Attributes.getNamedItem("PartyType").Value
			sAgName = sTempNode.Item(0).Attributes.getNamedItem("Agentname").Value
		End IF
	Else
		sAgType = ""
		sAgName = ""
	End IF

	document.all.spAgentName.innerHTML = sAgName
	document.formname.hCommName.Value = sAgName

	IF CStr(sAgType) = "CR" Then
		document.formname.optAgentExist(0).checked = True
	Elseif CStr(sAgChk) = "DR" Then
		document.formname.optAgentExist(1).checked = True
	Else
		document.formname.optAgentExist(2).checked = True
	End IF

	sExp = "//SaleInvoice"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		document.formname.txtInvoiceNo.value = sTempNode.Item(0).Attributes.getNamedItem("InvNo").Value
		document.formname.ctlDate.setDate = sTempNode.Item(0).Attributes.getNamedItem("InvDate").Value
		document.formname.txtRefNo.value = sTempNode.Item(0).Attributes.getNamedItem("RefNo").Value

	End IF

	sExp = "//SalesType"
	Set sTempNode = Root.selectNodes(sExp)
	IF sTempNode.length <> 0 Then
		For iCtr = 0 To document.formname.selSaleType.length - 1
			IF CStr(document.formname.selSaleType.options(iCtr).Value) = CStr(sTempNode.Item(0).Attributes.getNamedItem("SalType").Value) Then
				document.formname.selSaleType.selectedIndex = iCtr
				Exit For
			End IF
		Next
	End IF


	'sExp = "//PurCategory"
	'Set sTempNode = Root.selectNodes(sExp)
	'IF sTempNode.length <> 0 Then
	'	sTemp = sTempNode.Item(0).Attributes.getNamedItem("Code").Value
	'	For iCtr = 0 To document.formname.selPurCat.length - 1
	'		IF document.formname.selPurCat.options(iCtr).value = sTemp Then
	'			document.formname.selPurCat.selectedIndex = iCtr
	'			Exit For
	'		End IF
	'	Next
	'End IF


End Function


</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="DisplayBook();DispOldVal();DisplayVoucher()">

<form method="POST" name="formname" action="VouSALAmdTaxEntry.asp">
<input type="hidden" name="hVouCode" value="04">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hEditEntNo" value="0">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=iBookNo%>">
<input type="hidden" name="hAccUnit" value="<%=sAccUnit%>">
<input type="hidden" name="hAccUnitName" value="<%=sAccUnitName%>">
<input type="hidden" name="hSalAccCode" value="<%=iSalAccHead%>">
<input type="hidden" name="hSalAccName" value="<%=sSalAccHdName%>">
<input type="hidden" name="hItemCode" value="0">
<input type="hidden" name="hClassCode" value="0">
<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
<input type="hidden" name="hParCode" value="">
<input type="hidden" name="hCommName" value="">
<input type="hidden" name="hCurrDate" value="<%=date()%>">
<input type="hidden" name="hBkAccHead" value="">

<Input type="hidden" name="hPartyType" value="<%=nPartyType%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hTransNo" value="<%=nCreatedTransNo%>">
<input type="hidden" name="hInvDate" value="<%=sSetInvDate%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Sales Voucher Entry
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
								<!--<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>-->
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
								<td class="TabCell" valign="bottom" align="center" width="100">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Commission</td>
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
								  	<tr><td align="center">Voucher</td></tr>
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
			        <td height="20" valign="bottom">
			            <table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>

								<td valign="top" width="100%" align=left>
								    <table cellpadding="0" cellspacing="0" width="100%" class="TableOutLineOnly">
                                    <tr>
                                            <td height="8" colspan=2></td>
                                    </tr>
                                    <tr>
                                            <td class="FieldCellSub" width="108">Sales Book</td>
                                            <td class="FieldCell">
												<select size="1" name="selBook" class="FormElem" onChange="PopulateSalTy()">
													<option value="S">Select Book</option>
												</select>
                                            </td>
                                            <td class="FieldCellSub" width="108">Reference Number</td>
                                            <td class="FieldCell"><input type="text" name="txtRefNo" size="20" maxlength="30" class="FormElem">
                                            </td>
                                    </tr>
                                        <tr>
                                            <td class="FieldCellSub" width="108">Sale Type&nbsp;</td>
                                            <td class="FieldCell">
                                            	<select size="1" name="selSaleType" class="FormElem" onchange="GetAccHead(this)" >
									                <option value="0">Select Sale Type</option>
									                <%

										                IF CStr(sAccBookRel) <> "T" Then 'Book and Account Head is Not Done
											                sQuery = "Select InvoiceType,InvTypeShortName,InvoiceTypeName from Sal_M_InvoiceTypes where TobeAccounted=1 and Useable = 1 Order By InvoiceTypeName "
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
									  		                Set sName = objRs(2)
									  		                Do while not objRs.EOF
									  							If Trim(sCode) = trim(nSaleType) Then
									  							%>
																	<option value="<%Response.Write sCode%>" selected><%=trim(sName)%></option>
																<%
																Else
																%>
																	<option value="<%Response.Write sCode%>"><%=trim(sName)%></option>
																<%
																End IF
												                objRs.MoveNext
											                Loop
											                objRs.Close
										                End IF
								                    %>
								                </select>
                                            </td>
                                            <td class="FieldCellSub" width="108">Invoice Number</td>
											<td class="FieldCell">
												<input type="text" name="txtInvoiceNo" size="20" class="FormElem" onblur="VouCreate()">
                                            </td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub" width="108">Party Name</td>
                                            <td class="FieldCell" valign=middle>
                                                <input type="text" name="txtPartyName" size="40" class="FormElemRead" value="<%=sPartyName%>">
                                                &nbsp;&nbsp;<a href="#" onclick="SelPartyName()">
                                                <img src="../../assets/images/iTMS%20icons/Entry.gif" alt='Select Party Name' height="10" width="10"></a>
                                            </td>
                                            <td class="FieldCellSub" width="120"> Invoice Date</td>
											<td class="FieldCell">
												  <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
												  %>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub" width="108">Party Type</td>
                                            <td class="FieldCell" colspan="3">
                            	                <select size="1" name="selParType" class="FormElem" onchange="PartyType(this)">
								                <option value="A">Select Party Type</option>
								                </select>
                                              </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub" width="108">Agent </td>
                                            <td class="FieldCell" colspan="3">
                                                <table border="0" cellpadding="0" cellspacing="0">
                                                  <tr>
                                                    <td class="FieldCell">
                                                        <input type="radio" value="C" name="optAgentExist" onClick="showAgent('1')" class="Formelem">
                                                    </td>
                                                    <td class="FieldCell">Commission Agent</td>
                                                    <td class="FieldCell">
                                                        <input type="radio" value="D" name="optAgentExist" onClick="showAgent('2')" class="Formelem">
                                                    </td>
                                                    <td class="FieldCell">Depo Agent</td>
                                                    <td class="FieldCell"> <input type="radio" value="No" onClick="showAgent('N')" checked name="optAgentExist" class="Formelem">
                                                    </td>
                                                    <td class="FieldCell">
                                                  		No Agent</td>
                                                  	<td>
                                                  		<span ID="spAgentName" class="DataOnly"></span>
                                                  	</td>
                                                  </tr>
                                                </table>
                                            </td>
                                        </tr>
										<!--tr>
											<td class="FieldCell" width="108">&nbsp;</td>
											<td class="FieldCell"><span ID="spAgentName" class="DataOnly"></span></td>
											<td class="FieldCell" width="120"></td>
											<td class="FieldCell"></td>
										</tr-->
                                    </table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
        			        </td>
			                </tr>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td valign="top" width="100%" >
                                <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                <td class="MiddlePack" colspan="2"></td>
                                </tr>
                                <tr>
                            <td class="FieldCellSub">Sales Account Head</td>
                            <td class="FieldCell">
                            <select size="1" name="selAccountHead" class="FormElem" onfocus="VouCreate()" onChange="popSalesHead(this)" >
                            <%IF CStr(iSalAccHead) = "0" Then %>
								<option value="S" Selected>Sales Account Head</option>
								<option value="G">GL Account Head</option>
							<%Else%>
								<option value="S">Sales Account Head</option>
								<option value="G"  Selected>GL Account Head</option>
							<%End IF %>
                            </select>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell"></td>
                            <td class="FieldCell">
                            <%IF CStr(iSalAccHead) <> "0" Then %>
								<span class="DataOnly" id="spAccHead"><%=sSalAccHdName%> </span>
							<%Else%>
								<span class="DataOnly" id="spAccHead"></span>
							<%End IF %>
							</td>
                                </tr>
                                <tr>
                            <td class="FieldCellSub">Item Description</td>
                            <td class="FieldCell">
								<input type="text" name="txtDescription" size="40" class="FormElem" onfocus="VouCreate()">
                                <a href="#" onClick="GetItem()">
									<img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item Description" width="15" height="15">
								</a>
							</td>
							</tr>
                            <tr>
                            <td class="FieldCellSub">Quantity</td>
                            <td class="FieldCell">
							    <table border="0" cellpadding="0" cellspacing="0">
							     <tr>
							       <td width="65"></td>
							       <td><input type="text" name="txtQty" value="0.00" size="13"  maxlength="11" style="text-align: Right" class="FormElem" onBlur="calculateField(1)"></td>
							       <td width="10">
							       </td>
							       <td>
							   <select size="1" name="selUOM" class="FormElem" onChange="ChDisp(this)">
							<%
									Dim sCode,sValue,sSelUOM,sSelPack
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

								  	IF Not objRs.EOF Then
								  		sSelUOM = objRs(1)
								  	End IF

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
                            <td class="FieldCellSub">Rate</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
									<input type="text" name="txtRate" value="0.00" onBlur="calculateField(1)" size="15"  maxlength="13" style="text-align:right" class="FormElem">
								</td>
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
                            <td class="FieldCellSub">Actual Value</td>
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
                            <td class="FieldCellSub">Discount</td>
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
                            <td class="FieldCellSub">Sales Value</td>
                            <td class="FieldCell">
                            <table border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="65"></td>
                                <td>
                            <input type="text" name="txtAmount" size="15" readonly value="0.00" maxlength="13" style="text-align:right" class="FormElem" onBlur="popAddAmount1()"></td>
                              </tr>
                            </table>
                                  </td>
                                </tr>
                                  <tr>
                            <td class="FieldCellSub">Approval</td>
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
                                </tr>
                                <tr>
                            <td class="FieldCellSub">Rounded Off</td>
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
								<td valign="top" class="FieldCell" width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Add Entry" name="btnAdd" class="ActionButton" onclick="AddEntry('A')" >
                                                                <input type="button" value="Update" onClick="AddEntry('U')" name="btnUpdate" class="ActionButton" disabled>
                                                                <input type="button" value="Delete" onClick="DelEntry()" name="btnDel" class="ActionButton" disabled>
                                                                <input type="button" value="Next" onClick="AddEntry('S')" name="btnNext" class="ActionButton" >
                                                                <input type="reset" value="Cancel" name="btnCancel" class="ActionButton" onClick="Cancel('VouSALBookSelection.asp')" >
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
									<td valign="top" class="FieldCell" >
												<DIV class=frmBody id=DisVoucher style="width: 100%; height:140;">
                                                <table border="0" id="tblVoucher" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="10">&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center">Account Head</td>
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
<%
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
	    'VouCreate
		SaveXML
		Exit Function
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
		'document.all.spPack.innerHTML = document.formname.selPack.options(document.formname.selPack.selectedIndex).text

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

%>