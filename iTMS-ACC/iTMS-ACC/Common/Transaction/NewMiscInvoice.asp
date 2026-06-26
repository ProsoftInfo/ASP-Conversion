<%@ Language=VBScript %>
<%	Option Explicit%>
<%

	'Program Name				:  NewMiscInvoice.asp
	'Module Name				:  Purchase
	'Author Name				:  Ragavendran R
	'Created On					:  April 07,2011
	'Modified By			    :
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
%>
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/Purpopulate.asp" -->
<!-- #include File="../../include/PurItemCommon.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!--#include file="../../include/CommonFunctions.asp"-->
<%
Dim rsTemp
Dim sOrgID,sOrgName,sCreatedBy,sQuery,sAppCode,sParType
Dim iCreatedBy
Dim sFinPeriod,sArrPeriod,sFromDate,sToDate

set rsTemp = Server.CreateObject("ADODB.Recordset")

sOrgID = Session("organizationcode")
sOrgName = Session("OrgShortName")
iCreatedBy = Session("userid")
sCreatedBy = Session("username")
sAppCode = Request.QueryString("APPCODE")
if sAppCode = 2 then
    sParType = "CR"
else
    sParType = "DR"
end if

sFinPeriod = Session("FinPeriod")
sArrPeriod = Split(sFinPeriod,":")
sFromDate = "01/04/"& sArrPeriod(0)
sToDate = "31/03/"& sArrPeriod(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>iTMS - Misc. Invoice Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<xml id="OutData"><Root></Root></xml>
<xml id="MiscData"><Root></Root></xml>
<xml id="PartyData"><Root></Root></xml>
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="" BookName="" CRDR="" VouDate="" BookAcchead="" Approver="" PartyCode="" PartyType="" PartySubType=""  ReferenceNo="" hPayTo="" hPayFor="" hRefNo=""  PayFor="" PayForName="" PaymentThru="" AppRefNo="" AppRefDate="" AppRefType="" /></XML>
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="<%=sOrgId%>" AccName="<%=sOrgName%>" /></XML>
<Script Language="javascript" SRC="../../scripts/itms-modern-compat.js"></Script>
<Script Language="javascript" SRC="../../scripts/RefTypePop.js"></Script>
<Script Language="javascript" SRC="../../scripts/MiscInvoiceCompat.js"></Script>
<script language="javascript" src="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=vbscript>
''**********************************************
Function ChangePaymentMode()
    set objPay = eval("document.formname.radPayThru")
    if objPay(0).checked then
        tdChequeNo.style.display="none"
        tdChequeDate.style.display="none"
    elseif objPay(1).checked then
        tdChequeNo.style.display="block"
        tdChequeDate.style.display="block"
    end if
End Function
'************************************************
Function PaymentForChange()
 sValue = document.formname.selPayFor(document.formname.selPayFor.selectedIndex).value
	if sValue = "O" then
		document.formname.txtPayFor.value = "Payment for "
        document.formname.txtPayFor.disabled = false
	elseif sValue="F" then
        document.formname.txtPayFor.value = "Freight Payment for | "& RefNoDate.innerText
        document.formname.txtPayFor.disabled = true
    else
        alert("Select Payment For ")
        document.formname.selPayFor.focus
        exit function
	end if
End Function
'**************************************************
Function LoadPartySubType()
    sParCode = document.formname.hSupplierCode.value
    sOrgCode = document.formname.hOrgID.value
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","../PartySubType.asp?ParCode="&sParCode&"&OrgCode="&sOrgCode,false
    objhttp.send
    if trim(objhttp.responseXML.xml)<>"" then
        PartyData.loadXML(objhttp.responseXML.xml)
    else
        alert(objhttp.responseText)
    end if
    set ndRoot = PartyData.documentElement
    if ndRoot.hasChildNodes() then
       document.formname.selPartySubType.length = 0
        for each ndChild in ndRoot.childNodes
            document.formname.selPartySubType.length = document.formname.selPartySubType.length + 1
            document.formname.selPartySubType(document.formname.selPartySubType.length - 1).value =ndChild.getAttribute("SubType")
            document.formname.selPartySubType(document.formname.selPartySubType.length - 1).text =ndChild.text
        next
    end if
End Function
''****************************************************
Function RefType_Click()
Dim ndInvRoot,ndType
Dim ndInvHeader,ndHead,ndInvItemDet,ndItem
Dim ItemNode,MaterialNode,EntryNode
Dim sRefType,sPartyCode,sOrgID,sarrValue
Dim RcptNo,ForUnit,Flag,iItemEntryNo

set ndInvRoot = MiscData.documentElement

sOrgID = document.formname.hOrgID.value
if trim(sOrgID)="" or IsNull(sOrgID) then
		sOrgID = document.formname.hOrgID.value
end if

nFlag=1
iStock = "N"
bAddButton = "Y"
sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
sPartyCode = document.formname.hSupplierCode.value
if trim(sRefType)="N" then
    if trim(document.formname.hSupplierCode.value)="" then
        alert("Select the Party")
        exit function
    end if
end if
sDispItem = 0
RefTypeSelection sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem,"PUR"
	if trim(sRefType)<>"N" then
		set Root = OutData.documentElement
		'alert(Root.xml)
		if Root.hasChildNodes() then
			For each RefNode in Root.childNodes
			    if RefNode.nodeName="Reference" then

				    RcptNo	= RefNode.getAttribute("ReferenceNo")
				    ForUnit = sOrgID
				    Flag	= ""

				    RefNoDate.innerText = RefNode.getAttribute("ReferenceCode") &" - "& RefNode.getAttribute("ReferenceDate")

				    document.formname.hRefTypeCode.value = sRefType
				    document.formname.hRefno.value = RcptNo
				    document.formname.hOrgID.value = sOrgID
				    document.formname.hRefDate.value = RefNode.getAttribute("ReferenceDate")


				    if trim(RefNode.getAttribute("Remarks"))<>"" then
					    PartyCode = split(RefNode.getAttribute("Remarks"),"-")(0)
					    idSupplier.innerText = split(RefNode.getAttribute("Remarks"),"-")(1)
				    end if 'if trim(RefNode.getAttribute("Remarks"))<>"" then

				    document.formname.hSupplierCode.value = PartyCode
				    document.formname.hSupplierName.value  = idSupplier.innerText

				    set ndVouRoot = VoucherData.documentElement
				    ndVouRoot.setAttribute "ReferenceNo",RefNoDate.innerText
				    ndVouRoot.setAttribute "hRefNo",RcptNo
				    ndVouRoot.setAttribute "AppRefNo",RcptNo
				    ndVouRoot.setAttribute "AppRefDate", document.formname.hRefDate.value
				    ndVouRoot.setAttribute "AppRefType",sRefType
				    ndVouRoot.setAttribute "PartyCode",PartyCode
				    LoadPartySubType
				end if'if RefNode.nodeName="Reference" then
			next
		end if
	end if 'if trim(sRefType)<>"N" then
End Function
'*********************************************************
Function AddNewParty()
	OutValue = showModalDialog("MisParCreate.asp?"&OutValue,"","dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'MsgBox OutValue
	idPayTo.innerHTML = OutValue
End Function

Function SelMisParty()
	Dim arrTemp,sRetValue,sParCode,sPartyName,sTemp

	OutValue = showModalDialog("../../Common/MisPartySelection.asp","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	IF CStr(OutValue) = "AN" Then
		AddNewParty()
		'Exit Function
	End IF
	arrTemp = split(OutValue,":")

	while UBound(arrTemp) = 0
		OutValue = showModalDialog("../../Common/MisPartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
	sRetValue = OutValue
	'MsgBox sRetValue
	if UBound(arrTemp) <= 1 then exit function

	sTemp = Split(sRetValue,":")
	idPayTo.innerHTML = sTemp(0)
	document.formname.hMisPartyCode.value = sTemp(1)

	IF sTemp(0) <> "" Then
		document.formname.txtPayTo.readOnly = True
	End IF

	'sParTy = sTemp(4)
	'sParSubType = sTemp(3)
	'sParCode = sTemp(1)
	'sPartyName = sTemp(0)
End Function
'*************************************************
Function Supplierselect()
	Dim opt,nFlag,sAct,sQuery,Root
	nFlag = 1
	sUnit = document.formname.hOrgID.value
	sParType = document.formname.hParType.value
	set OutValue=showModalDialog("../SupplierSelect.asp?Unit="+sUnit+"&hSelectMode=S&Flag="+cstr(nFlag)&"&ParType="&sParType,OutData,"status:no")

	'msgbox OutValue.xml

	sAct = UCase(trim(OutValue.getAttribute("Action")))
	'alert(sAct)
	sQuery = trim(OutValue.getAttribute("PassQuery"))
	if ucase(trim(sAct)) <> "CLOSE" then
		do while sAct <> "DONE"
			set OutValue=showModalDialog("../SupplierSelect.asp?" & sQuery,OutData,"status:no")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			sQuery = trim(OutValue.getAttribute("PassQuery"))

			if ucase(Trim(sAct)) = "CLOSE" then exit do
		loop
	end if 'if ucase(trim(sAct)) <> "CLOSE" then


	If not OutValue.hasChildNodes Then 	exit function

	set Root = OutData.DocumentElement

	For each Node2 in OutValue.childNodes
		if ucase(Node2.nodename) = ucase("Supplier") then
			ssCode = trim(ssCode) & trim(Node2.getAttribute("SCode")) & ","
			ssupcode = trim(ssupcode) & trim(Node2.getAttribute("SuppCode")) & ","
			ssuppname= trim(ssuppname) & trim(Node2.getAttribute("SuppName")) & ","
		end if 'if Strcomp(Node2.nodename,"Supplier")= 0 then
	Next
	if right(sscode,1) = "," then  sscode = mid(sscode,1,len(sscode) - 1 )
	if right(ssuppname,1) = "," then  ssuppname = mid(ssuppname,1,len(ssuppname) - 1 )
	if right(ssupcode,1) = "," then  ssupcode = mid(ssupcode,1,len(ssupcode) - 1 )

	idSupplier.innerHTML = ssuppname
	document.formname.hSupplierCode.value=trim(ssupcode)
	document.formname.hSupplierName.value=trim(ssuppname)
	LoadPartySubType
End function
'**************************************************
Function CheckSubmit()
	Dim sCheckPayToSelection

    if document.formname.hSupplierCode.value="" then
        alert("Select Party")
        exit function
    end if
    If IDPayTo.innerHTML  = "" and document.formname.txtPayTo.value = ""  Then
		alert("Select Pay to Received from Party or Enter the party name in Textbox")
		Exit Function
    End IF
    if document.formname.selPayFor.selectedIndex = 0 then
        alert("Select Payment For")
        document.formname.selPayFor.focus
        exit function
    end if

    sPayFor = document.formname.selPayFor(document.formname.selPayFor.selectedIndex).value

    if document.formname.radPayThru(0).checked =true then
        sPaymentThru  = document.formname.radPayThru(0).value
    else
        sPaymentThru  = document.formname.radPayThru(1).value
    end if

    If document.formname.AdjAgainBill.checked Then
		AdjustAgainstInvoice = "Y"
	Else
		AdjustAgainstInvoice = "N"
	End IF

    'if document.formname.radTransType(0).checked = true then
    '    sCRDR = document.formname.radTransType(0).value
    'else
    '    sCRDR = document.formname.radTransType(1).value
    'end if 'if document.formname.radTransType(0).checked = true then

    sCRDR = "C"

    sParType = document.formname.selPartySubType(document.formname.selPartySubType.selectedIndex).value

    if trim(sParType)<>"0" or trim(sParType)<>"" then
        sArrPartyType = split(sParType,"|")
        sPartyType = sArrPartyType(0) ' CR,DR
        sPartySubType = sArrPartyType(1) '  Sub Levels
    end if

    If IDPayTo.innerHTML <> "" Then
		sCheckPayToSelection = "YES"
	Else
		sCheckPayToSelection = "NO"
	End IF

	set objPay = eval("document.formname.radPayThru")
    if objPay(1).checked then
        sChequeNo = document.formname.txtChequeNo.value
        sChequeDate = document.formname.ctlChequeDate.getdate()
    else
        sChequeNo = ""
        sChequeDate = ""
    end if

    set ndVouRoot = VoucherData.documentElement
        ndVouRoot.setAttribute "VouDate",document.formname.ctlInvoiceDate.getDate()
        ndVouRoot.setAttribute "PartyType",sPartyType
        ndVouRoot.setAttribute "PartySubType",sPartySubType
        ndVouRoot.setAttribute "PayFor",sPayFor
        ndVouRoot.setAttribute "hPayFor",sPayFor
        ndVouRoot.setAttribute "PaymentThru",sPaymentThru
        ndVouRoot.setAttribute "Code",document.formname.hSupplierCode.value
        ndVouRoot.setAttribute "CheNo",sChequeNo
        ndVouRoot.setAttribute "CheDate",sChequeDate
   ' alert(ndVouRoot.xml)

    set ndEntRoot = EntryData.documentElement
        ndEntRoot.setAttribute "No","1"
        ndEntRoot.setAttribute "CRDR",sCRDR
        ndEntRoot.setAttribute "Payto",document.formname.hSupplierName.value
        ndEntRoot.setAttribute "Amount",document.formname.txtAmount.value
        ndEntRoot.setAttribute "PayToSelCheck",sCheckPayToSelection 'Check whether MiscPartysel or Enter Name
        If IDPayTo.innerHTML <> "" Then
			ndEntRoot.setAttribute "MiscPartyName",IDPayTo.innerHTML
        Else
			ndEntRoot.setAttribute "MiscPartyName",document.formname.txtPayTo.value
		End IF
        ndEntRoot.setAttribute "MiscPartyCode",document.formname.hMisPartyCode.value
        ndEntRoot.setAttribute "CheckVal",AdjustAgainstInvoice	'ChekBoxVal

    set ndNarr = EntryData.createElement("Narration")
    ndNarr.text = document.formname.txtPayFor.value
    ndEntRoot.appendChild ndNarr

    ndVouRoot.appendChild ndEntRoot
    'alert(ndVouRoot.xml)

    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "POST","XMLSave.asp?Name=MISCPayment&SessionFlag=true",false
    objhttp.send ndVouRoot.xml

    document.formname.action = "MiscInvoiceInsert.asp"
    document.formname.submit
End Function
'*******************************************************
Function setdate()
sFromDate = document.formname.hFromDate.value
sToDate = document.formname.hToDate.value
if DateDiff("d",sTodate,date)>0 then
    document.formname.ctlInvoiceDate.setMinDate = sFromDate
    document.formname.ctlInvoiceDate.setMaxDate = sTodate
    document.formname.ctlInvoicedate.setdate = sTodate
else
    document.formname.ctlInvoiceDate.setMinDate = sFromDate
    document.formname.ctlInvoiceDate.setMaxDate = date
    document.formname.ctlInvoicedate.setdate = date
end if

End Function
'*********************************
</SCRIPT>
</HEAD>


<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<BODY leftMargin=0 topMargin=0 onload="setdate()">
<FORM NAME="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=Hidden name="hSupplierCode" value="">
<input type=Hidden name="hSupplierName" value="">
<input type=hidden name="hParType" value="<%=sParType%>">

<input type=hidden name="hRefTypeCode" value="">
<input type=hidden name="hRefno" value="">
<input type=hidden name="hRefDate" value="">
<input type=hidden name="hAppCode" value="<%=sAppCode%>">
<input type=hidden name="hMisPartyCode" value="0">
<input type=hidden name="hFromDate" value="<%=sFromDate%>">
<input type=hidden name="hToDate" value="<%=sToDate%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
          Miscellaneous Invoice (Create)
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%" align="left">
                                       <table BORDER="0" CELLSPACING="0" CELLPADDING="0" width="100%">
                                        <tr>
                                            <td class="FieldCellSub">Reference Type</td>
                                            <td class="FieldCell">
                                                <select name=SelRefName class="FormElem" >
                                                <%
                                                    if sAppCode = 2 then
                                                        RefTypePop 13,2
                                                    else
                                                        RefTypePop 5,3
                                                    end if
                                                %>
                                                </select>
                                                &nbsp;<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item (s)" width="11" height="11" onClick="RefType_Click()"></a>
                                            </td>
                                            <td class="FieldCellSub">Invoice Date</td>
                                            <td class="FieldCell">
                                                <%  Response.write InsertDatePicker("ctlInvoiceDate") %>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Reference No - Date</td>
                                            <td class="FieldCell">
                                                <span id="RefNoDate" class="DataOnly">&nbsp;N/A&nbsp;</span>
                                            </td>
                                            <td class="FieldCellSub">Created By</td>
                                            <td class="FieldCell">
                                                <span class="DataOnly"><%=sCreatedBy%></span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Party</td>
                                            <td class="FieldCell">
                                            	<span class="dataonly" id="idSupplier"></span> &nbsp;
	                                            &nbsp;<img id="Img1" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="Supplierselect()" align="middle" alt="Supplier Selection" width="10" height="11">
                                            </td>
                                            <td class="FieldCellSub">Party SubType</td>
                                            <td class="FieldCell">
                                                <select name="selPartySubType" class="FormElem">

                                                </select>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Pay To</td>
                                            <td class="FieldCell" colspan="3">
                                            	<span class="dataonly" id="IDPayTo"></span> &nbsp;
	                                            &nbsp;<img id="Img1" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="SelMisParty()" align="middle" alt="Miscellaneous Party" width="10" height="11">
	                                            <input type=text name="txtPayTo" value="" class=FormElem style="text-align:right">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Pay For</td>
                                            <td class="FieldCell" colspan=3 valign="Top">
                                                <select name="selPayFor" class="FormElem" onchange="PaymentForChange()">
                                                    <option value="S">Select</option>
                                                    <option value="F">Freight</option>
                                                    <option value="O">Others</option>
                                                </select>
                                                &nbsp;&nbsp;<textarea name="txtPayFor" class="FormElem" rows="2" cols="65"></textarea>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Amount</td>
                                            <td class="FieldCell" colspan="3">
                                                <input type=text name=txtAmount value="0" class="FormElem" size="10" style="text-align:right">
                                                &nbsp;&nbsp;<Input type="checkbox" name="AdjAgainBill">To be Adjusted Against bill[Select If amount to be borne by party]
                                            </td>
                                            <!--<td class="FieldCellSub">Transaction Type</td>
                                            <td class="FieldCell">
                                                <input type=radio name=radTransType value="C" checked>Credit&nbsp;&nbsp;
                                                <input type=radio name=radTransType value="D">Debit
                                            </td>-->
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Payment Thru</td>
                                            <td class="FieldCell">
                                                <input type=radio name=radPayThru value="C" onclick="ChangePaymentMode()" checked>Cash&nbsp;&nbsp;
                                                <input type=radio name=radPayThru value="D" onclick="ChangePaymentMode()">Cheque
                                            </td>
                                            <td class="FieldCell" align="left" id="tdChequeNo" style="display:none">
                                                No: <input type="text" name="txtChequeNo" class="FormElem" /> Date:
                                            </td>
                                            <td class="FieldCell" align="left" id="tdChequeDate" style="display:none">
                                               <%  Response.write InsertDatePicker("ctlChequeDate") %>
                                            </td>
                                        </tr>
                                     </table>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                             <tr>
                                <td align="center">
								</td>
								<td align="center"class="ActionCell">
								    <input type=button name=btnClose value="Save" onclick="CheckSubmit()" class="ActionButton">
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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
</html>
