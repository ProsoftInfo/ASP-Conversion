<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MsiVouBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 20, 2004
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
<%
dim sOrgId,sBookId,iBookNo,sOrgName,sBookName
dim sFlag,sFromVal,sToVal,sRefType,sRefName,sVouStatus,sAppCode,sAction
dim objRs,objRs1
Dim iTotalPages,iPrevPage,iNextPage
Const iPageSize=20
Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt,iCnt

sOrgId=Session("organizationcode")

sBookId=Request("selVoucher")
iBookNo=Request("selBook")
sOrgName=Request("horgName")
sBookName=Request("hBookName")
sFlag=Request("optCriteria")

sAction = Request.QueryString("ACTN")
if sAction = "P" then
    sAppCode = 2
else
    sAppCode = 3
end if

Dim sFinPeriod,sFromYr,sToYr,sTempYr

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
End IF


select case sFlag
	case "VouNo"
			sFromVal=Request("txtNoFrom")
			sToVal=Request("txtNoFrom")
	case "VouDate"
			sFromVal=Request("hFDate")
			sToVal=Request("hTDate")
	case "Amount"
			sFromVal=Request("txtGAmount")
			sToVal=Request("txtLAmount")
	case "AccHead"
			sFromVal=Request("SelAccHead")
			sToVal=Request("hAccHead")
	Case "Exist"
		sOrgName=Session("OrgName")
		sBookName=Session("BookName")

		sFromVal=Session("FromValue")
		sToVal=Session("ToValue")
		sFlag=Session("Flag")
end select

Session("FromValue")=sFromVal
Session("ToValue")=sToVal
Session("Flag")=sFlag

Session("OrgName")=sOrgName
Session("BookName")=sBookName

%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML ID="OutData"><PartyType/></XML>
<XML id="AccHeadData">
<account/>
</XML>
<XML id="PartyData"><Root></Root></XML>
<XML id="TempXMLData"><Root></Root></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT language="vbscript">
dim sFlag
'*************************************
Function ChkSubmit()
    nCnt = document.formname.hCnt.value
    For iCnt = 1 to nCnt
        set sobj = eval("document.formname.ChkMiscZ"&iCnt)
        if sObj.checked = true then
            sValue = sobj.value
            iSelCount = iSelCount + 1
        end if
    Next
    if iSelCount >1 or iSelCount < 1 then
        alert("Select any one to create")
        exit function
    end if
    sArrValue = Split(sValue,"Z")
    if sArrValue(4)<>"010101" then
        alert("Already Entry Created for the Invoice")
        exit function
    end if
    if sArrValue(0) = "C" then
        document.formname.action = "MsiVouEntry.asp?TransNo="&sArrValue(1)&"&OrgName="&sArrValue(2)&"&OrgID="&sArrValue(3)
    else
       document.formname.action =  "MsiVouEntryForBank.asp?TransNo="&sArrValue(1)&"&OrgName="&sArrValue(2)&"&OrgID="&sArrValue(3)
    end if
    document.formname.submit
End Function
'********************************
Function AssignPage(nPage)
'alert(nPage)
	document.formname.hPage.value = nPage
	document.formname.action = "MsiVouBookSelection.asp"
	document.formname.submit()
end function
'------------------------------------------------------------------------------------------
Function popPartType()
set objhttp = CreateObject("MSXML2.XMLHTTP")

	iUnitNo=document.formname.hUnitId.value
	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentElement
			iCounter=document.formname.SelAccHead.length
			For Each HeaderNode In Root.childNodes
				set oText1 = document.createElement("<Option>" )
					oText1.Text = HeaderNode.text
					oText1.Value = HeaderNode.Attributes.getNamedItem("ParType").Value
				document.formname.selAccHead.add oText1,iCounter
				iCounter=CDbl(iCounter)+1
			next
	end if

end function
Function DisplayBook()
dim iUnitNo,arrTemp,BkCode
dim Root
	document.formname.selBook.options.length = 1
		iUnitNo= document.formname.hUnitId.value
		BkCode= document.formname.selVoucher.value

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode="&BkCode&"&orgID=" & iUnitNo , false
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

end Function
function validate()

	If sFlag="VouNo" Then
		if sFlag="VouDate" Then
			sFromDate=document.formname.CtlVouFromDate.GetDate
			sToDate=document.formname.CtlVouToDate.GetDate
			If dateDiff("d",sFromDate,sToDate)<0 Then
				Msgbox "To Date Should be Greater than From Date"
				Exit Function
			end if
'------------- Coding For the case Amount is Selected ------------------
		Elseif sFlag="Amount" Then
			If	document.formname.txtGAmount.value="" Then
				Msgbox "Enter From Amount"
				document.formname.txtGAmount.select
				Exit Function
			Elseif document.formname.txtLAmount.value="" Then
				Msgbox "Enter To Amount"
				document.formname.txtLAmount.select
				Exit Function
			Elseif not(IsNumeric(document.formname.txtGAmount.value)) Then
				Msgbox "Enter Numbers Only"
				document.formname.txtGAmount.select
				Exit Function
			Elseif not(IsNumeric(document.formname.txtLAmount.value)) Then
				Msgbox "Enter Numbers Only"
				document.formname.txtLAmount.select
				Exit Function
			Else
				dGAmount=cdbl(document.formname.txtGAmount.value)
				dLAmount=cdbl(document.formname.txtLAmount.value)
				If cdbl(dGAmount)>cdbl(dLAmount) Then
					Msgbox "To Amount Should be Greater Than From Amount "
					document.formname.txtLAmount.value =""
					document.formname.txtLAmount.select
					Exit Function
				End if
			end if
		end if
	End IF
'	alert(document.formname.optCriteria(0).checked)
IF document.formname.optCriteria(0).checked then
	 document.formname.hoptCriteria.value = document.formname.optCriteria(0).value
ElseIF document.formname.optCriteria(1).checked then
 	document.formname.hoptCriteria.value = document.formname.optCriteria(1).value
End If

	document.formname.horgName.value=document.formname.hOrgName.value
	document.formname.hFDate.value=document.formname.CtlVouFromDate.GetDate
	document.formname.hTDate.value=document.formname.CtlVouToDate.GetDate
	document.formname.action = "MiscPayments.asp"
	document.formname.submit()
End function

Function OptSelection()
	if document.formname.optCriteria(0).checked then
		sFlag=document.formname.optCriteria(0).value
		document.formname.txtGAmount.value = ""
		document.formname.txtLAmount.value = ""
		document.formname.txtGAmount.readOnly = True
		document.formname.txtLAmount.readOnly = True
	Elseif document.formname.optCriteria(1).checked then
		sFlag=document.formname.optCriteria(1).value
		document.formname.txtGAmount.value = ""
		document.formname.txtLAmount.value = ""
		document.formname.txtGAmount.readOnly = false
		document.formname.txtLAmount.readOnly = false

	End if
	document.formname.hoptCriteria.value = sFlag
End Function

Function SelNew()
	document.formname.SelAccHead.selectedIndex=0
	document.formname.txtGAmount.value  =""
	document.formname.txtLAmount.value =""
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	window.spAccHead.innerHTML =""
End Function
'---------- Selection of Account Head From Pop up Screen --------------

Function SelectAccHead()
dim iGlHead,sOrgId,sAccHead,arrTemp,sRetVal
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sRetValue,sTemp
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

set objhttp = CreateObject("Microsoft.XMLHTTP")

if document.formname.selVoucher.value="0" Then
		Msgbox "Select Voucher Type"
		document.formname.selVoucher.focus
		document.formname.SelAccHead.selectedIndex=0
		Exit Function
Elseif document.formname.selBook.value="S" Then
		Msgbox "Select Book"
		document.formname.selBook.focus
		document.formname.SelAccHead.selectedIndex=0
		Exit Function
Else
	sOrgId=document.formname.hUnitId.value
	sBookid=document.formname.selVoucher.value
	sBookNo=document.formname.selBook.value
	if 	document.formname.SelAccHead.value="G" then

	    sTempValWindowSize = GetWindowSizeForPopup("5")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)

		Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

	'	'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgid="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	'	OutValue = showModalDialog("GLHeadSelection.asp?orgid="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'	arrTemp = split(OutValue,":")
'
'		while UBound(arrTemp) = 0
'			OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'			arrTemp = split(OutValue,":")
'		wend
'
'		sRetVal = OutValue
'
'		if UBound(arrTemp) <= 1 then exit function

        if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    sRetVal = ndEntry.getAttribute("RetField0")&":"&ndEntry.getAttribute("RetField1")&":"&ndEntry.getAttribute("RetField2")&":"&ndEntry.getAttribute("RetField3")&":"&ndEntry.getAttribute("RetField4")&":"&ndEntry.getAttribute("RetField5")&":"&ndEntry.getAttribute("RetField6")
                end if
            next
        end if

		GetGlHeadXml(sRetVal)

		Set nodAccHead = AccHeadData.documentElement


		if nodAccHead.hasChildNodes then
			For Each HeaderNode In nodAccHead.childNodes
				document.formname.hAccHead.value=HeaderNode.Attributes.getNamedItem("No").Value
				window.spAccHead.innerHTML=HeaderNode.Attributes.getNamedItem("Name").Value&"&nbsp;"
			next
		else
			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
			window.spAccHead.innerHTML=""
		End if
	else

	    sPartyType=document.formname.SelAccHead.value& "?" & document.formname.SelAccHead.options(document.formname.SelAccHead.selectedIndex).text

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


	'	'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	'	OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'	arrTemp = split(OutValue,":")
'
'		while UBound(arrTemp) = 0
'			OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'			arrTemp = split(OutValue,":")
'		wend
'
'		sRetValue = OutValue
'		sTemp = Split(sRetValue,":")
'		sParTy = sTemp(4)
'		sParSubType = sTemp(3)
'		sParCode = sTemp(1)
'		sPartyName = sTemp(0)

        if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    sParTy = ndEntry.getAttribute("RetField3")
		            sParSubType = ndEntry.getAttribute("RetField4")
		            sParCode = ndEntry.getAttribute("RetField1")
		            sPartyName = ndEntry.getAttribute("RetField0")
		        exit for
                end if
            next
        end if


		objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
		objhttp.send

		IF objhttp.responseText <> "" Then
			sRetVal2 = objhttp.responseText
			GetPartyHeadXml sParCode,sPartyName,sRetVal2
		End IF
		Set nodAccHead = AccHeadData.documentElement

		if nodAccHead.hasChildNodes then
			For Each HeaderNode In nodAccHead.childNodes
				document.formname.hAccHead.value=sPartyType&"?"& HeaderNode.Attributes.getNamedItem("No").Value
				window.spAccHead.innerHTML=HeaderNode.Attributes.getNamedItem("Name").Value&"&nbsp;"
			next
		else
			document.formname.SelAccHead.selectedIndex=0
			document.formname.hAccHead.value="0"
			window.spAccHead.innerHTML=""
		End if
	end if
End if

End Function
'----------------------------------------------------------------------------------------
Function SetDate()
 	document.formname.ctlVouFromDate.setDate = document.formname.hFDate.value
	document.formname.ctlVouToDate.setDate = document.formname.hTDate.value
	'alert(document.formname.hoptCriteria.value)
	IF document.formname.hoptCriteria.value <> "" then
		IF document.formname.hoptCriteria.value = "VouDate" then
			document.formname.optCriteria(0).checked  = true
		ElseIF document.formname.hoptCriteria.value = "Amount" then
			document.formname.optCriteria(1).checked  = true
		End IF
	End IF
End Function
'----------------------------------------------------------------------------------------
Function MinDate()
	sFromDate = document.formname.ctlVouFromDate.GetDate
	sToDate = document.formname.ctlVouToDate.GetDate
	sMinDate = document.formname.hFDate.value
	sMaxDate = document.formname.hTDate.value
'alert(sMinDate & sMaxDate)
	If dateDiff("d",sFromDate,document.formname.hFDate.value) > 0 or  dateDiff("d",sFromDate,document.formname.hTDate.value) < 0 then
		alert("Date Should be within the Financial Year  "& sMinDate&" to " & sMaxDate )
		document.formname.ctlVouFromDate.SetDate =	document.formname.hFDate.value
		exit Function
	end if
	If dateDiff("d",sToDate,document.formname.hFDate.value) > 0 or datediff("d",sToDate,document.formname.hTDate.value) < 0  then
		alert("Date Should be within the Financial Year  "& sMinDate &" to " & sMaxDate )
		document.formname.ctlVouToDate.SetDate = document.formname.hTDate.value
		exit Function
	end if
End Function
'-------------------------------------------------------------------------------------------


</script>
<%
dim sFinTemp,sMaxDate,sMinDate,Da,Mo,Yr
iPageNo=trim(Request("hPage"))
	if iPageNo="" then iPageNo=1

sFinPeriod = Session("FinPeriod")
'Response.Write sFinPeriod
IF CStr(sFinPeriod) <> "" Then
	sFinTemp = Split(sFinPeriod,":")
	sMaxDate = "31/03/"&sFinTemp(1)
	sMinDate = "01/04/"&sFinTemp(0)
End IF

IF year(sMinDate) = Year(Date()) then
	Da = Day(Date())
	IF len(Da) = 1 then Da = 0&Da
	Mo = Month(Date())
	IF len(Mo) = 1 then Mo = 0&Mo
	 sMaxDate = Da&"/"&Mo&"/"&Year(Date())
End IF
'	Response.Write
'Response.Write sMinDate & " *** "& sMaxDate

%>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload ="SetDate()">

<form method="POST" name="formname" action="">
<input type=hidden name="hUnitID" value="<%=sOrgID%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<Input type="hidden" name="hFDate" value="<%=sMinDate%>">
<Input type="hidden" name="hTDate" value="<%=sMaxDate%>">
<Input type="hidden" name="hoptCriteria" value="">
<input type=hidden name="hPage" value="">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Supplementary Pay/Invoices
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>
<tr>
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%">
<div id="idUnprocessed" style="display: none">
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%" border=0>
<tr>
<td class="MiddlePack" colspan="6">
</td>
</tr>
 	<tr>
 	    <td align="left" class="FieldCellSub"></td>
 	  <td class="FieldCellSub">
      <input onclick="OptSelection()" type="radio" value="VouDate" name="optCriteria">
        Voucher Date</td>
        <td align="right" class="FieldCellSub" >From</td>
      <td align="left" class="FieldCellSub">
		 <object id="ctlVouFromDate"  onblur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"  codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
			<param name="_ExtentX" value="2355">
			<param name="_ExtentY" value="529">
		</object>
	  </td>
	  <td align="left" class="FieldCellSub" >To</td>
      <td align="left" class="FieldCellSub">
		 <object id="ctlVouToDate"  onblur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
			<param name="_ExtentX" value="2355">
			<param name="_ExtentY" value="529">
		</object>
		</td>

    </tr>
<tr>
  <td align="left" class="FieldCellSub"></td>
  <td class="FieldCellSub"><input type="radio" onclick="OptSelection()"  value="Amount" name="optCriteria">
    Amount</td>
  <td align="left" class="FieldCellSub"></td>
  <td align="left" class="FieldCellSub"><input class="FormElem" size="11" Readonly name="txtGAmount"></td>
  <td align="left" class="FieldCellSub"></td>
  <td align="left" class="FieldCellSub"><input class="FormElem" size="11" Readonly name="txtLAmount"></td>
  <td align="left" class="FieldCellSub"></td>
</tr>

<tr>
<td class="FieldCell" colspan="4" align="center">
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
</td>

</table>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
<table border="0" cellspacing="1px" class="ExcelTable" width="100%" >

<tr>
<td class="ExcelHeaderCell" width="10px">S.No.</td>
<td class="ExcelHeaderCell" width="10px"></td>
<td class="ExcelHeaderCell">Reference Type</td>
<td class="ExcelHeaderCell">Reference No-Date</td>
<td class="ExcelHeaderCell">Voucher No - Date</td>
<td class="ExcelHeaderCell">Amount</td>
<td class="ExcelHeaderCell">Created By</td>
</tr>
<!--td class="ExcelHeaderCell" align="center">Type</td-->
<SCRIPT LANGUAGE=vbscript RUNAT=Server>

</SCRIPT>
<%
dim sQuery,iSno,iTransNo,sVouNo,sVouDate,dAmount,sType,iCreatedBy,sCreatedByName,iApproveLevel
Dim nSlNo,iRecCtr,iTotalRecords, iStartRec,iEndRec,nPageCtr,rsStatus,rsTemp
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Response.Write "<font color=#000000>"
sQuery = "Select MiscTransNo,isNull(CreatedMiscPymtNo,''),Convert(Char,VoucherDate,103), "&_
		 "VoucherAmount,CreatedBy,isNull(BankInstrumentType,'C'),AppRefType,CreatedVouchStatus From Acc_T_MiscPymtRequestHeader Where  "&_
		 "OUDefinitionID = '"&sOrgId&"' and ApplicationCode = "& sAppCode

		' Response.Write sQuery

select case sFlag
	case "VouDate"
		sQuery=sQuery&" and VoucherDate >= convert(datetime,'"&sFromVal&"',103) and VoucherDate <= convert(datetime,'"&sToVal&"',103) "
	case "Amount"
		sQuery=sQuery&" and VoucherAmount>="&sFromVal&" and VoucherAmount<="&sToVal&" "

end select
	sQuery=sQuery&" order by 1"

'	Response.write sQUery
nslno=0
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.PageSize=iPageSize
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

iSno=1
iCnt=0
iRecCtr = 1
If not objRs.EOF then

	iTotalPages = objRs.PageCount
	iTotalRecords = objRs.RecordCount
	objRs.AbsolutePage = iPageNo
Else
	iTotalPages = 0
	iTotalRecords = 0

	iStartRec = 0
	iEndRec = 0
End If
	'Response.Write"<p>rfq="&iRFQNo

if trim(iPageNo) = 1 then
	iPrevPage = 0
else
	iPrevPage = iPageNo - 1
end if

if iTotalPages >= iPageNo + 1 then
	iNextPage = iPageNo + 1
else
	iNextPage = 0
end if


'Response.Write Objrs.PageSize


	Do While Not objRs.EOF and iSno <= Objrs.PageSize

	iCnt = iCnt + 1

		set iTransNo = objRs(0)
		set sVouNo = objRs(1)
		set sVouDate = objRs(2)
		set dAmount = objRs(3)
		set iCreatedBy= objRs(4)
		set sRefType = objRs(6)
		set sVouStatus = objRs(7)

		if Trim(sRefType)<>"" then
			sQuery = "Select ReferenceName from VW_ReferenceTypes where ReferenceEntryNo = "& sRefType
			'Response.Write sQuery
			objRs1.Open sQuery,con
			if not objRs1.EOF then
			    sRefName = trim(objRs1(0))
			end if
			objRs1.Close
		end if'if Trim(sRefType)<>"" then


		dAmount = FormatNumber(dAmount,2,,,0)

		sQuery="select EmployeeName from  Ms_EmployeeMaster where EmployeeNumber="&iCreatedBy
		with objRs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		sCreatedByName=objRs1(0)
		objRs1.Close


%>
    <tr>
    <td class="ExcelSerial"><%=iSno%></td>
	<td class="ExcelDisplayCell">
	    <input type="checkbox" name="ChkMiscZ<%=iSNO%>" value="<%=objRs(5)%>Z<%=iTransNo%>Z<%=sOrgName%>Z<%=sOrgID%>Z<%=sVouStatus%>" />
	<%' IF CStr(objRs(5)) = "C" Then %>
		<!--<A href="MsiVouEntry.asp?TransNo=<%=iTransNo%>&OrgName=<%=sOrgName%>&OrgID=<%=sOrgId%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>-->
	<%'Else%>
		<!--<A href="MsiVouEntryForBank.asp?TransNo=<%=iTransNo%>&OrgName=<%=sOrgName%>&OrgID=<%=sOrgId%>"><img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="View Details"></a></td>-->
	<%'End IF %>
	</td>
	<td class="ExcelDisplayCell"><%=sRefName%></td>
	<td class="ExcelDisplayCell"><%=sVouNo%></td>
	<td class="ExcelDisplayCell"><p align="center"><%=iTransNo%>-<%=sVouDate%></td>
	<td class="ExcelDisplayCell"><%=dAmount%></td>
	<td class="ExcelDisplayCell"><%=sCreatedByName%></td>


                                            </tr>
<%
		objRs.MoveNext

		iSno=CInt(iSno)+1

	LOOP



objRs.Close

%>
</table>
<!--/div-->
</td>
<td align="center" class="ClearPixel" width="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top" align="right">
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hCnt" value=<%=iCnt  %>>
<input type=hidden name="hPageSelection" value="0">

<input type="button" value=" |< " class="ActionButtonX" id=ButFirst name=ButFirst onClick="AssignPage('1')">

<%if trim(iPrevPage) = "0" then  %>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev >
<%else%>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev onClick="AssignPage('<%=iPrevPage%>')">
<%end if %>


<SELECT class="FormElem" onChange="AssignPage(this.value)"  id="mCmbPage" name="mCmbPage">

<%for nPageCtr= 1 to iTotalPages %>
	<option value="<%=nPageCtr%>" <%if trim(iPageNo) = trim(nPageCtr) then Response.Write "Selected" %> >Page <%=nPageCtr%> of <%=iTotalPages %></option>
<%next%>

</SELECT>
<%if trim(iNextPage) = "0" then  %>
	<input type="button" value=" >> " class="ActionButtonX" id=ButNext name=ButNext >
<%else%>
	<input type="button" value=" >> " class="ActionButtonX" onclick="AssignPage('<%=iNextPage%>')" id=ButNext name=ButNext >
<%end if%>

<input type="button" value=" >| " class="ActionButtonX" id=ButLast name=ButLast OnClick="AssignPage('<%=iTotalPages %>')">

</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>
<tr>
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td class="ActionCell">
    <input type="button" value="Create Voucher" name="btnCreate" class="ActionButtonX" tabindex="3" onclick="ChkSubmit()">
    <input type="button" value="Edit" name="B9" class="ActionButton" tabindex="4" >
    <input type="button" value="Approve" name="B10" class="ActionButton" tabindex="5" >
</td>
</tr>
</table>
</td>
<td align="center" class="ClearPixel" width="5">
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

</table>
</td>
</tr>

</table>
</td>
</tr>

</table>
</form>
</body>
</html>
